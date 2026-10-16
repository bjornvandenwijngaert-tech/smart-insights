"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

process.env.TZ = "Europe/London";
const root = path.resolve(__dirname, "..");
const variants = [
  { label: "standalone legacy", file: "legacy-trip-report/src/app.js", legacy: true },
  { label: "parent legacy", file: "src/app.js", legacy: true },
  { label: "parent activity", file: "src/app.js", legacy: false },
  { label: "standalone activity", file: "activity-report/src/app.js", legacy: false, privateScope: true }
];

function fixture() {
  const devices = [
    { id: "bVan", name: "Vehicle One - TEST01", driver: { id: "bCurrent" } },
    { id: "bOther", name: "Vehicle Two - TEST02" }
  ];
  const users = [
    { id: "bDriverA", firstName: "Alex", lastName: "Example", name: "alex@example.test" },
    { id: "bDriverB", firstName: "Jamie", lastName: "Sample", name: "jamie@example.test", isDriver: false, activeTo: "2026-10-13T00:00:00Z" }
  ];
  function trip(id, device, driver, day, hour, distance, stopped) {
    const start = `2026-10-${day}T${String(hour).padStart(2, "0")}:00:00Z`;
    const stop = `2026-10-${day}T${String(hour + 1).padStart(2, "0")}:00:00Z`;
    return { id, device: { id: device }, driver: { id: driver }, start, stop, distance,
      drivingDuration: "01:00:00", idlingDuration: "00:05:00", stopDuration: stopped,
      stopPoint: { x: Number(id.slice(1)), y: 50 } };
  }
  const trips = [
    trip("t1", "bVan", "bDriverA", 12, 7, 10, "01:00:00"),
    trip("t2", "bVan", "bDriverB", 12, 9, 20, "01:00:00"),
    trip("t3", "bVan", "bDriverA", 12, 11, 30, "19:00:00"),
    trip("t4", "bVan", "bDriverB", 13, 7, 40, "01:00:00"),
    trip("t5", "bVan", "UnknownDriverId", 13, 9, 5, null),
    trip("t6", "bOther", "bDriverA", 12, 14, 5, "00:30:00")
  ];
  trips[4].driverName = "Current Driver Must Not Be Used";
  return { devices, users, trips };
}

function loadApp(variant, data = fixture(), denyDrivers = false) {
  const elements = {}, calls = [], downloads = [], pdfs = [], workbooks = [];
  function element(id) {
    if (!elements[id]) elements[id] = {
      value: "", innerHTML: "", checked: false, children: [],
      classList: { add() {}, remove() {}, toggle() {} },
      addEventListener() {}, appendChild(child) { this.children.push(child); }
    };
    return elements[id];
  }
  element("filter-from").value = "2026-10-07";
  element("filter-to").value = "2026-10-13";
  element("report-type").value = variant.legacy ? "legacy-trip-history" : "activity-report";
  element("col-total-engine").checked = true;
  element("legacy-length").value = "7";
  element("range-length").value = "7";

  class FakeBlob {
    constructor(parts, options) { this.parts = parts; this.type = options.type; }
  }
  class FakePdf {
    constructor(options) {
      this.tables = []; this.texts = []; this.pages = 1;
      this.internal = { pageSize: { getWidth: () => options.orientation === "landscape" ? 297 : 210, getHeight: () => options.orientation === "landscape" ? 210 : 297 }, getNumberOfPages: () => this.pages };
      pdfs.push(this);
    }
    setFillColor() {} rect() {} setTextColor() {} setFontSize() {} setFont() {} roundedRect() {}
    text(value) { this.texts.push(value); }
    splitTextToSize(value) { return [value]; }
    addPage() { this.pages++; } setPage() {}
    autoTable(options) { this.tables.push(options); this.lastAutoTable = { finalY: 80 }; }
    save(filename) { this.filename = filename; }
  }
  class FakeWorkbook {
    constructor() { this.worksheets = []; this.xlsx = { writeBuffer: () => Promise.resolve(new Uint8Array([1])) }; workbooks.push(this); }
    addWorksheet(name) {
      assert(!this.worksheets.some(ws => ws.name === name), "Duplicate worksheet name");
      const ws = { name, rows: [], addRow(values) { const row = { values: Array.from(values) }; this.rows.push(row); return row; }, addImage() {} };
      this.worksheets.push(ws); return ws;
    }
    addImage() { return 1; }
  }
  const context = vm.createContext({
    console, Date, Map, Blob: FakeBlob, Uint8Array,
    setTimeout, clearTimeout, geotab: { addin: {} },
    localStorage: { getItem() { return null; }, setItem() {} },
    URL: { createObjectURL(blob) { downloads.push(blob); return "blob:test"; }, revokeObjectURL() {} },
    alert(message) { throw new Error(message); },
    document: {
      getElementById: element,
      createElement(tag) { return tag === "canvas" ? { getContext() { return {}; }, toDataURL() { return "data:image/png;base64,test"; } } : { click() {} }; }
    },
    jspdf: { jsPDF: FakePdf }, ExcelJS: { Workbook: FakeWorkbook },
    Chart: function () { this.destroy = function () {}; }
  });
  context.window = context;
  context.location = { hostname: "my.geotab.com", pathname: "/demo/" };
  let source = fs.readFileSync(path.join(root, variant.file), "utf8");
  if (variant.privateScope) {
    assert(/\}\)\(\);\s*$/.test(source));
    source = source.replace(/\}\)\(\);\s*$/, "globalThis.reportTestHooks = {state, runActivityReport, filterReportByVehicle, exportActivityCsv, exportActivityPdf, exportActivityExcel, reportDriverId, isUnassignedDriver, reportUserName, reportDriverInfo, resolveReportDrivers, groupReportDriverDays, buildByDevice, buildDayBlocks, fmtDateInput, applyRangeLength, localDayKey};\n})();");
  }
  vm.runInContext(source, context, { filename: variant.file });
  const hooks = variant.privateScope ? context.reportTestHooks : context;
  hooks.state.api = { call(method, params, success, failure) {
    calls.push({ method, params: JSON.parse(JSON.stringify(params)) });
    if (method === "Get" && params.typeName === "Device") success(data.devices);
    else if (method === "Get" && params.typeName === "Trip") success(data.trips);
    else if (method === "Get" && params.typeName === "User") {
      if (denyDrivers) failure(new Error("User permission denied"));
      else success(data.users.filter(user => params.search.userIds.includes(user.id)));
    } else if (method === "GetAddresses") success(params.coordinates.map(pt => ({ formattedAddress: "Stop " + pt.x })));
    else throw new Error("Unexpected API operation: " + method + " " + JSON.stringify(params));
  } };
  hooks.state.server = "my.geotab.com";
  hooks.state.dbName = "demo";
  hooks.state.unitSystem = "Metric";
  return { hooks, elements, calls, downloads, pdfs, workbooks, data };
}

function runReport(app, variant) {
  let completions = 0;
  app.hooks[variant.legacy ? "runLegacyTripHistoryReport" : "runActivityReport"]("2026-10-07", "2026-10-13", () => completions++);
  assert.equal(completions, 1);
  return variant.legacy ? app.hooks.state.legacyByDevice : app.hooks.state.activityByDevice;
}

function parseCsv(blob) {
  const text = blob.parts.join("");
  const rows = [], row = [];
  let value = "", quoted = false;
  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    if (char === '"' && quoted && text[i + 1] === '"') { value += '"'; i++; }
    else if (char === '"') quoted = !quoted;
    else if (!quoted && (char === "," || char === "\n")) {
      row.push(value); value = "";
      if (char === "\n") { rows.push(row.splice(0)); }
    } else value += char;
  }
  row.push(value); rows.push(row);
  return rows;
}

for (const variant of variants) {
  test(variant.label + ": historical names and asset/driver/day groups", () => {
    const app = loadApp(variant);
    const byDevice = runReport(app, variant);
    const groups = byDevice.bVan.driverGroups;
    assert.equal(groups.length, 3);
    assert.deepEqual(Array.from(groups.find(g => g.id === "bDriverA").days[0].rows, r => r.trip.id), ["t1", "t3"]);
    assert.deepEqual(Array.from(groups.find(g => g.id === "bDriverB").days, d => d.date), ["2026-10-12", "2026-10-13"]);
    const resumed = groups.find(g => g.id === "bDriverA").days[0].rows[1];
    assert.equal(resumed.prevTrip.id, "t2", "Driver handover must preserve the vehicle's actual predecessor");
    const html = app.elements["report-output"].innerHTML;
    assert(html.includes("Driver: Alex Example"));
    assert(html.includes("Driver: Jamie Sample"), "Archived/former driver must still resolve by id");
    assert(html.includes("Unassigned / unknown driver"));
    assert(!html.includes("Current Driver Must Not Be Used"));
    assert.equal((html.match(/<table /g) || []).length, 5);
    const userCalls = app.calls.filter(c => c.params.typeName === "User");
    assert.equal(userCalls.length, 1);
    assert.deepEqual(userCalls[0].params.search, { userIds: ["bDriverA", "bDriverB"] });
  });

  test(variant.label + ": CSV attributes each trip to one driver", () => {
    const app = loadApp(variant); runReport(app, variant);
    app.hooks[variant.legacy ? "exportReportCsv" : "exportActivityCsv"]();
    const rows = parseCsv(app.downloads.at(-1));
    assert.equal(rows[0][1], "Driver");
    assert.equal(rows.length, 7);
    assert.equal(rows.filter(r => r[1] === "Alex Example").length, 3);
    assert.equal(rows.filter(r => r[1] === "Jamie Sample").length, 2);
    assert.equal(rows.filter(r => r[1] === "Unassigned / unknown driver").length, 1);
    if (!variant.legacy) {
      const resumed = rows.find(r => r[7] === "Stop 3");
      assert.equal(resumed[4], "Stop 2", "Activity CSV must not skip the other driver's preceding stop");
    }
  });

  test(variant.label + ": PDF repeats vehicle, driver and day context", () => {
    const app = loadApp(variant); runReport(app, variant);
    app.hooks[variant.legacy ? "exportLegacyTripHistoryPdf" : "exportActivityPdf"]();
    const tables = app.pdfs[0].tables;
    assert.equal(tables.length, 5);
    assert(tables.every(t => t.head[0][0].content.includes(" | Driver: ")));
    assert(tables.some(t => t.head[0][0].content.includes("Jamie Sample | 2026-10-13")));
    assert(tables.some(t => t.head[0][0].content.includes("Alex Example | 2026-10-12")));
  });

  test(variant.label + ": vehicle filter updates rows, exports and summary", () => {
    const app = loadApp(variant); runReport(app, variant);
    const originalSummary = app.elements["report-summary"].innerHTML;
    app.hooks.filterReportByVehicle("bOther");
    assert.notEqual(app.elements["report-summary"].innerHTML, originalSummary);
    assert(app.elements["report-summary"].innerHTML.includes("Number of Stops</div><div class='summary-value'>1<"));
    assert(!app.elements["report-output"].innerHTML.includes("Vehicle One - TEST01"));
    app.hooks[variant.legacy ? "exportReportCsv" : "exportActivityCsv"]();
    assert.equal(parseCsv(app.downloads.at(-1)).length, 2);
    app.hooks.filterReportByVehicle("");
    assert.equal(app.elements["report-summary"].innerHTML, originalSummary);
    if (variant.legacy) assert(originalSummary.includes("Number of Stops</div><div class='summary-value'>4<"));
  });

  test(variant.label + ": unavailable names remain distinct from unassigned trips", () => {
    const app = loadApp(variant, fixture(), true); runReport(app, variant);
    const html = app.elements["report-output"].innerHTML;
    assert(html.includes("Driver name unavailable (bDriverA)"));
    assert(html.includes("Driver name unavailable (bDriverB)"));
    assert(html.includes("Unassigned / unknown driver"));
    assert.equal(app.hooks.state.reportData.length, 6);
  });

  test(variant.label + ": identity keys survive duplicate display names and cache resets", () => {
    const data = fixture(); data.users[1].firstName = "Alex"; data.users[1].lastName = "Example";
    const app = loadApp(variant, data);
    let byDevice = runReport(app, variant);
    assert.equal(byDevice.bVan.driverGroups.length, 3);
    assert(byDevice.bVan.driverGroups.some(g => g.name.includes("alex@example.test")));
    assert(byDevice.bVan.driverGroups.some(g => g.name.includes("jamie@example.test")));
    data.users[0].firstName = "Updated";
    byDevice = runReport(app, variant);
    assert(byDevice.bVan.driverGroups.some(g => g.name === "Updated Example"), "Each run must refresh names rather than reuse stale database/user data");
  });

  test(variant.label + ": local date inputs and day grouping agree in BST", () => {
    const data = fixture(); data.trips = [data.trips[0]];
    data.trips[0].start = "2026-06-30T23:30:00Z"; data.trips[0].stop = "2026-07-01T00:30:00Z";
    const app = loadApp(variant, data);
    assert.equal(app.hooks.fmtDateInput(new Date("2026-07-01T00:00:00")), "2026-07-01");
    for (const [days, expectedFrom] of [["7", "2026-06-25"], ["30", "2026-06-02"]]) {
      app.elements["filter-to"].value = "2026-07-01";
      app.elements[variant.privateScope ? "range-length" : "legacy-length"].value = days;
      app.hooks[variant.privateScope ? "applyRangeLength" : "applyLegacyLength"]();
      assert.equal(app.elements["filter-from"].value, expectedFrom);
      assert.equal(app.elements["filter-to"].value, "2026-07-01");
    }
    const byDevice = runReport(app, variant);
    assert.equal(byDevice.bVan.driverGroups[0].days[0].date, "2026-07-01");
  });

  if (!variant.legacy) test(variant.label + ": Excel has driver sections and a driver on every trip row", async () => {
    const app = loadApp(variant); runReport(app, variant);
    app.hooks.exportActivityExcel();
    await Promise.resolve();
    const workbook = app.workbooks[0];
    assert.equal(workbook.worksheets.length, 3);
    const sheet = workbook.worksheets.find(ws => ws.name === "Vehicle One - TEST01");
    assert.equal(sheet.rows.filter(r => String(r.values[0]).startsWith("Driver: ")).length, 3);
    assert.equal(sheet.rows.filter(r => r.values[0] === "Driver" && r.values[1] === "Start Time").length, 4);
    const trips = sheet.rows.filter(r => r.values.length > 7 && !String(r.values[1]).includes("Total") && r.values[1] !== "Start Time" && r.values[2] !== "(Ignition On)");
    assert.equal(trips.length, 5);
    assert.equal(trips.filter(r => r.values[0] === "Alex Example").length, 2);
    assert.equal(trips.find(r => r.values[4] === "Stop 3").values[2], "Stop 2");
  });

  if (!variant.legacy) test(variant.label + ": identically named assets remain separate in Excel", async () => {
    const data = fixture(); data.devices[1].name = data.devices[0].name;
    const app = loadApp(variant, data);
    const byDevice = runReport(app, variant);
    assert.equal(Object.keys(byDevice).length, 2);
    app.hooks.exportActivityExcel();
    await Promise.resolve();
    assert.deepEqual(app.workbooks[0].worksheets.map(ws => ws.name), ["Summary", "Vehicle One - TEST01", "Vehicle One - TEST01 (2)"]);
  });
}

test("driver lookup is batched by unique ids and helper copies stay synchronized", () => {
  const app = loadApp(variants[0], { devices: [], users: [], trips: [] });
  const trips = Array.from({ length: 201 }, (_, i) => ({ driver: { id: "bDriver" + i } }));
  let finished = false;
  app.hooks.resolveReportDrivers(trips.concat(trips), () => { finished = true; });
  assert(finished);
  assert.deepEqual(app.calls.map(c => c.params.search.userIds.length), [100, 100, 1]);
  const copies = [variants[0], variants[1], variants[3]].map(v => loadApp(v).hooks);
  for (const name of ["reportDriverId", "isUnassignedDriver", "reportUserName", "reportDriverInfo", "resolveReportDrivers", "groupReportDriverDays"]) {
    assert.equal(copies[0][name].toString(), copies[1][name].toString(), name);
    assert.equal(copies[0][name].toString(), copies[2][name].toString(), name);
  }
});

test("release manifests, badges and local asset cache versions agree", () => {
  for (const folder of ["", "legacy-trip-report", "activity-report"]) {
    const directory = path.join(root, folder);
    const manifest = JSON.parse(fs.readFileSync(path.join(directory, "config.json"), "utf8"));
    const html = fs.readFileSync(path.join(directory, "index.html"), "utf8");
    assert(html.includes('class="version-badge">v' + manifest.version + "</span>"), folder);
    for (const asset of ["styles.css", "logo-data.js", "app.js"]) assert(html.includes(asset + "?v=" + manifest.version), folder + "/" + asset);
  }
});
