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
  const devices = [{ id: "bVan", name: "Vehicle One - TEST01", driver: { id: "bCurrent" } }, { id: "bOther", name: "Vehicle Two - TEST02" }];
  const users = [
    { id: "bDriverA", firstName: "Alex", lastName: "Example", name: "alex@example.test" },
    { id: "bDriverB", firstName: "Jamie", lastName: "Sample", name: "jamie@example.test", isDriver: false, activeTo: "2026-10-13T00:00:00Z" }
  ];
  function trip(id, device, driver, day, hour, distance, stopped) {
    return { id, device: { id: device }, driver: { id: driver },
      start: `2026-10-${day}T${String(hour).padStart(2, "0")}:00:00Z`, stop: `2026-10-${day}T${String(hour + 1).padStart(2, "0")}:00:00Z`,
      distance, drivingDuration: "01:00:00", idlingDuration: "00:05:00", stopDuration: stopped, stopPoint: { x: Number(id.slice(1)), y: 50 } };
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
  const elements = {}, calls = [], downloads = [], filenames = [], pdfs = [], workbooks = [];
  function element(id) {
    if (!elements[id]) elements[id] = { value: "", innerHTML: "", checked: false, children: [],
      classList: { values: new Set(), add(value) { this.values.add(value); }, remove(value) { this.values.delete(value); },
        toggle(value) { this.values.has(value) ? this.values.delete(value) : this.values.add(value); }, contains(value) { return this.values.has(value); } },
      addEventListener() {}, appendChild(child) { this.children.push(child); } };
    return elements[id];
  }
  element("filter-from").value = "2026-10-07"; element("filter-to").value = "2026-10-13";
  element("report-type").value = variant.legacy ? "legacy-trip-history" : "activity-report";
  element("col-total-engine").checked = true;
  element("legacy-length").value = "7"; element("range-length").value = "7";
  class FakeBlob { constructor(parts, options) { this.parts = parts; this.type = options.type; } }
  class FakePdf {
    constructor(options) {
      this.tables = []; this.texts = []; this.pages = 1;
      this.internal = { pageSize: { getWidth: () => options.orientation === "landscape" ? 297 : 210, getHeight: () => options.orientation === "landscape" ? 210 : 297 }, getNumberOfPages: () => this.pages };
      pdfs.push(this);
    }
    setFillColor() {} rect() {} setTextColor() {} setFontSize() {} setFont() {} roundedRect() {}
    text(value) { this.texts.push(value); } splitTextToSize(value) { return [value]; }
    addPage() { this.pages++; } setPage() {}
    autoTable(options) { this.tables.push(options); this.lastAutoTable = { finalY: 80 }; }
    save(filename) { this.filename = filename; }
  }
  class FakeWorkbook {
    constructor() { this.worksheets = []; this.xlsx = { writeBuffer: () => Promise.resolve(new Uint8Array([1])) }; workbooks.push(this); }
    addWorksheet(name) {
      assert(!this.worksheets.some(ws => ws.name === name), "Duplicate worksheet name");
      const ws = { name, rows: [], properties: {}, get rowCount() { return this.rows.length; }, addRow(values) {
        const row = { number: this.rows.length + 1, values: Array.from(values), cells: {}, getCell(column) {
          if (!this.cells[column]) {
            let n = column, letter = "";
            while (n) { letter = String.fromCharCode(65 + (n - 1) % 26) + letter; n = Math.floor((n - 1) / 26); }
            const cell = { address: letter + this.number };
            Object.defineProperty(cell, "value", { get: () => this.values[column - 1], set: value => { this.values[column - 1] = value; } });
            this.cells[column] = cell;
          }
          return this.cells[column];
        } };
        this.rows.push(row); return row;
      }, addImage() {} };
      this.worksheets.push(ws); return ws;
    }
    addImage() { return 1; }
  }
  const context = vm.createContext({
    console, Date, Map, Blob: FakeBlob, Uint8Array, setTimeout, clearTimeout, geotab: { addin: {} },
    localStorage: { getItem() { return null; }, setItem() {} },
    URL: { createObjectURL(blob) { downloads.push(blob); return "blob:test"; }, revokeObjectURL() {} },
    alert(message) { throw new Error(message); },
    document: { getElementById: element, createElement(tag) {
      return tag === "canvas" ? { getContext() { return {}; }, toDataURL() { return "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII="; } }
        : { click() { if (this.download) filenames.push(this.download); } };
    } }, jspdf: { jsPDF: FakePdf }, ExcelJS: { Workbook: FakeWorkbook }, Chart: function () { this.destroy = function () {}; }
  });
  context.window = context; context.location = { hostname: "my.geotab.com", pathname: "/demo/" };
  let source = fs.readFileSync(path.join(root, variant.file), "utf8");
  if (variant.privateScope) source = source.replace(/\}\)\(\);\s*$/, "globalThis.reportTestHooks = {state, runActivityReport, runReport, getActivityCols, restoreActivityColumns, saveSettings, filterReportByVehicle, exportActivityCsv, exportActivityPdf, exportActivityExcel, reportDriverId, isUnassignedDriver, reportUserName, reportDriverInfo, resolveReportDrivers, groupReportDriverDays, buildActivityByDevice, fmtDateInput, applyRangeLength, localDayKey, activityVehicleIds, activityReportRange, activityRangeLabel, activityTotals, fmtActivityTime, activityLastTrip, activityDayEndpoints, activityDayElapsed, activitySummaryColumns, activitySummaryHeaders, activitySummaryGridStyle, activitySummaryNumbers, activityNumberText, activityDayValues, activityVehicleNumbers, renderActivityReport, activityWallDate, addActivityGroupedWorksheet};\n})();");
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
  hooks.state.server = "my.geotab.com"; hooks.state.dbName = "demo"; hooks.state.unitSystem = "Metric";
  return { hooks, elements, calls, downloads, filenames, pdfs, workbooks, data, context };
}

function runReport(app, variant, from = "2026-10-07", to = "2026-10-13") {
  let completions = 0;
  app.hooks[variant.legacy ? "runLegacyTripHistoryReport" : "runActivityReport"](from, to, () => completions++);
  assert.equal(completions, 1);
  return variant.legacy ? app.hooks.state.legacyByDevice : app.hooks.state.activityByDevice;
}

function parseCsv(blob) {
  const text = blob.parts.join(""), rows = [], row = [];
  let value = "", quoted = false;
  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    if (char === '"' && quoted && text[i + 1] === '"') { value += '"'; i++; }
    else if (char === '"') quoted = !quoted;
    else if (!quoted && (char === "," || char === "\n")) { row.push(value); value = ""; if (char === "\n") rows.push(row.splice(0)); }
    else value += char;
  }
  row.push(value); rows.push(row); return rows;
}

for (const variant of variants.filter(v => v.legacy)) {
  test(variant.label + ": historical names and asset/driver/day groups", () => {
    const app = loadApp(variant), groups = runReport(app, variant).bVan.driverGroups;
    assert.equal(groups.length, 3);
    assert.deepEqual(Array.from(groups.find(g => g.id === "bDriverA").days[0].rows, r => r.trip.id), ["t1", "t3"]);
    assert.deepEqual(Array.from(groups.find(g => g.id === "bDriverB").days, d => d.date), ["2026-10-12", "2026-10-13"]);
    assert.equal(groups.find(g => g.id === "bDriverA").days[0].rows[1].prevTrip.id, "t2");
    const html = app.elements["report-output"].innerHTML;
    for (const label of ["Driver: Alex Example", "Driver: Jamie Sample", "Unassigned / unknown driver"]) assert(html.includes(label));
    assert(!html.includes("Current Driver Must Not Be Used"));
    assert.equal((html.match(/<table /g) || []).length, 5);
    assert.deepEqual(app.calls.filter(c => c.params.typeName === "User").map(c => c.params.search), [{ userIds: ["bDriverA", "bDriverB"] }]);
  });
  test(variant.label + ": CSV attributes every trip to one historical driver", () => {
    const app = loadApp(variant); runReport(app, variant); app.hooks.exportReportCsv();
    const rows = parseCsv(app.downloads.at(-1));
    assert.equal(rows[0][1], "Driver"); assert.equal(rows.length, 7);
    assert.equal(rows.filter(r => r[1] === "Alex Example").length, 3);
    assert.equal(rows.filter(r => r[1] === "Jamie Sample").length, 2);
    assert.equal(rows.filter(r => r[1] === "Unassigned / unknown driver").length, 1);
  });
  test(variant.label + ": PDF repeats vehicle, driver and day context", () => {
    const app = loadApp(variant); runReport(app, variant); app.hooks.exportLegacyTripHistoryPdf();
    const tables = app.pdfs[0].tables;
    assert.equal(tables.length, 5); assert(tables.every(t => t.head[0][0].content.includes(" | Driver: ")));
    assert(tables.some(t => t.head[0][0].content.includes("Jamie Sample | 2026-10-13")));
    assert(tables.some(t => t.head[0][0].content.includes("Alex Example | 2026-10-12")));
  });
  test(variant.label + ": vehicle filter updates rows, exports and cards", () => {
    const app = loadApp(variant); runReport(app, variant);
    const summary = app.elements["report-summary"].innerHTML, output = app.elements["report-output"].innerHTML;
    app.hooks.filterReportByVehicle("bOther");
    assert.notEqual(app.elements["report-summary"].innerHTML, summary);
    assert(app.elements["report-summary"].innerHTML.includes("Number of Stops</div><div class='summary-value'>1<"));
    assert(!app.elements["report-output"].innerHTML.includes("Vehicle One - TEST01"));
    app.hooks.exportReportCsv(); assert.equal(parseCsv(app.downloads.at(-1)).length, 2);
    app.hooks.filterReportByVehicle("");
    assert.equal(app.elements["report-output"].innerHTML, output); assert.equal(app.elements["report-summary"].innerHTML, summary);
    assert(summary.includes("Number of Stops</div><div class='summary-value'>4<"));
  });
  test(variant.label + ": unavailable names stay distinct from unassigned trips", () => {
    const app = loadApp(variant, fixture(), true); runReport(app, variant);
    const html = app.elements["report-output"].innerHTML;
    for (const label of ["Driver name unavailable (bDriverA)", "Driver name unavailable (bDriverB)", "Unassigned / unknown driver"]) assert(html.includes(label));
    assert.equal(app.hooks.state.reportData.length, 6);
  });
  test(variant.label + ": identity keys survive duplicate driver names and cache resets", () => {
    const data = fixture(); data.users[1].firstName = "Alex"; data.users[1].lastName = "Example";
    const app = loadApp(variant, data); let groups = runReport(app, variant).bVan.driverGroups;
    assert.equal(groups.length, 3); assert(groups.some(g => g.name.includes("alex@example.test"))); assert(groups.some(g => g.name.includes("jamie@example.test")));
    data.users[0].firstName = "Updated"; groups = runReport(app, variant).bVan.driverGroups;
    assert(groups.some(g => g.name === "Updated Example"));
  });
  test(variant.label + ": local inputs and grouping agree in BST", () => {
    const data = fixture(); data.trips = [data.trips[0]]; data.trips[0].start = "2026-06-30T23:30:00Z"; data.trips[0].stop = "2026-07-01T00:30:00Z";
    const app = loadApp(variant, data);
    assert.equal(app.hooks.fmtDateInput(new Date("2026-07-01T00:00:00")), "2026-07-01");
    for (const [days, from] of [["7", "2026-06-25"], ["30", "2026-06-02"]]) {
      app.elements["filter-to"].value = "2026-07-01"; app.elements["legacy-length"].value = days; app.hooks.applyLegacyLength();
      assert.equal(app.elements["filter-from"].value, from); assert.equal(app.elements["filter-to"].value, "2026-07-01");
    }
    assert.equal(runReport(app, variant).bVan.driverGroups[0].days[0].date, "2026-07-01");
  });
}

for (const variant of variants.filter(v => !v.legacy)) {
  test(variant.label + ": one vehicle/day combines all historical and unassigned trips", () => {
    const app = loadApp(variant), vehicle = runReport(app, variant).bVan;
    assert.equal(vehicle.driverGroups, undefined); assert.equal(vehicle.days.length, 2);
    assert.deepEqual(Array.from(vehicle.days[0].rows, r => r.trip.id), ["t1", "t2", "t3"]);
    assert.deepEqual(Array.from(vehicle.days[1].rows, r => r.trip.id), ["t4", "t5"]);
    assert.equal(vehicle.days[0].rows[2].prevTrip.id, "t2"); assert.equal(vehicle.days[1].rows[0].prevTrip.id, "t3");
    assert.equal(app.hooks.activityTotals(vehicle.days[0].rows.map(r => r.trip)).distance, 60);
    assert.equal(app.hooks.activityTotals(vehicle.days[1].rows.map(r => r.trip)).distance, 45);
    assert.equal(app.calls.filter(call => call.params.typeName === "User").length, 0);
  });
  test(variant.label + ": daily summary rows are sorted with a single vehicle footer", () => {
    const data = fixture(); data.trips.reverse(); data.devices.reverse();
    const app = loadApp(variant, data); runReport(app, variant);
    const html = app.elements["report-output"].innerHTML;
    assert.equal((html.match(/class='activity-vehicle-block'/g) || []).length, 2);
    assert.equal((html.match(/class='activity-day-row activity-summary-grid'/g) || []).length, 3);
    assert.equal((html.match(/class='activity-vehicle-total activity-summary-grid'/g) || []).length, 2);
    for (const label of ["Driver:", "Unassigned / unknown", "activity-driver-fold", "activity-week-fold", "summary-card"]) assert(!html.includes(label));
    assert(!/<(?:details|summary|table|tbody|tr)\b/.test(html), "No trip-level or expandable content, including hidden markup");
    assert(!html.includes("Stop 2"), "Intermediate trip destinations are not emitted");
    assert(html.indexOf("Vehicle One") < html.indexOf("Vehicle Two"));
    assert(html.indexOf("data-day='2026-10-12'") < html.indexOf("data-day='2026-10-13'"));
  });
  test(variant.label + ": CSV has daily rows and vehicle totals without driver allocation", () => {
    const app = loadApp(variant); runReport(app, variant); app.hooks.exportActivityCsv();
    const rows = parseCsv(app.downloads.at(-1));
    assert.equal(rows.length, 6); assert(!rows[0].includes("Driver"));
    assert.deepEqual(rows[1].slice(1, 6), ["12/10/2026", "08:00", "(location unknown)", "13:00", "Stop 3"]);
    assert.deepEqual(rows[1].slice(6), ["5.00", "21.00", "0.25", "3.00", "3:15", "60.00"]);
    assert.deepEqual(rows[2].slice(6), ["3.00", "1.00", "0.17", "2.00", "2:10", "45.00"]);
    assert.equal(rows[3][0], "Vehicle One - TEST01 Total"); assert.equal(rows[3][11], "105.00");
  });
  test(variant.label + ": PDF repeats vehicle daily headers and one footer", () => {
    const app = loadApp(variant); runReport(app, variant); app.hooks.exportActivityPdf();
    const tables = app.pdfs[0].tables;
    assert.equal(tables.length, 2); assert.equal(tables[0].body.length, 3);
    assert.equal(tables[0].head[0][0].content, "Vehicle One - TEST01");
    assert.equal(tables[0].body[0][0], "12/10/2026"); assert.equal(tables[0].body[0][10], "60.00");
    assert.equal(tables[0].body.at(-1)[0].content, "Vehicle One - TEST01 Total");
    assert(!JSON.stringify(tables).includes("Driver:"));
  });
  test(variant.label + ": Excel exports daily rows only with no hidden trip itemisation", async () => {
    const app = loadApp(variant); runReport(app, variant); await app.hooks.exportActivityExcel();
    const wb = app.workbooks[0], ws = wb.worksheets[1];
    assert.equal(wb.views[0].activeTab, 1); assert.equal(ws.rowCount, 6);
    assert.equal(ws.properties.outlineLevelRow, 1); assert(ws.properties.outlineProperties.summaryBelow);
    assert(ws.views[0].showOutlineSymbols); assert.equal(ws.views[0].ySplit, 1);
    assert.deepEqual(ws.rows.slice(1).filter(r => !r.hidden).map(r => r.outlineLevel), [1, 1, 0, 1, 0]);
    assert(ws.rows.slice(1).every(r => r.hidden === false && !r.collapsed && r.outlineLevel <= 1));
    const daily = ws.rows.filter(r => r.outlineLevel === 1);
    assert.equal(daily.length, 3);
    assert.equal(new Set(daily.map(r => r.values[0] + ":" + r.values[1].toISOString())).size, 3);
    assert(!ws.rows[0].values.includes("Driver")); assert(ws.columns.slice(-2).every(c => c.hidden));
    assert.deepEqual(ws.rows[0].values.slice(-2), ["Day Start (UTC)", "Day End (UTC)"]);
    assert.equal(ws.rows[1].values[3], "(location unknown)"); assert.equal(ws.rows[1].values[5], "Stop 3");
    assert.equal(ws.rows[2].values[3], "Stop 3"); assert.equal(ws.rows[2].values[5], "Stop 5");
    assert.equal(ws.rows[1].values[1].toISOString(), "2026-10-12T00:00:00.000Z");
    assert.equal(ws.rows[1].getCell(2).numFmt, "dd/mm/yyyy"); assert.equal(ws.rows[1].getCell(3).numFmt, "hh:mm");
    assert(!JSON.stringify(wb.worksheets).includes("Stop 2"), "No intermediate trip records in any worksheet");
  });
  test(variant.label + ": Excel totals count trip metrics and elapsed daily spans exactly once", async () => {
    const app = loadApp(variant); runReport(app, variant); await app.hooks.exportActivityExcel();
    const ws = app.workbooks[0].worksheets[1];
    assert.deepEqual(ws.rows[1].values.slice(6, 12).map(v => typeof v === "number" ? v : v.result), [5, 21, 0.25, 3, 195 / 1440, 60]);
    assert.equal(ws.rows[3].values[6].result, 8); assert.equal(ws.rows[3].values[11].result, 105);
    assert.equal(ws.rows[3].values[6].formula, "SUM(G2:G3)");
    assert.equal(ws.rows[1].values[6].formula, "24*(N2-M2)");
    assert.equal(ws.rows[1].values[10].formula, "(I2+J2)/24");
    ws.rows.forEach((row, index) => {
      if (row.outlineLevel !== 0) return;
      for (let c = 7; c <= 12; c++) {
        const value = row.values[c - 1]; if (!value || !value.formula) continue;
        const match = /^SUM\(([A-Z]+)(\d+):\1(\d+)\)$/.exec(value.formula); assert(match);
        const first = Number(match[2]), last = Number(match[3]); assert(last < index + 1);
        const expected = ws.rows.slice(first - 1, last).reduce((sum, r) => {
          const value = r.values[c - 1]; return sum + (typeof value === "number" ? value : value.result);
        }, 0);
        assert(Math.abs(value.result - expected) < 1e-10, value.formula);
      }
    });
  });
  test(variant.label + ": units, work split and >24-hour engine totals are numeric", async () => {
    const data = fixture(); Object.assign(data.trips[0], { drivingDuration: "1.02:00:00", workDrivingDuration: "01:30:00", workIdlingDuration: "00:15:00", afterHoursDrivingDuration: "00:30:00", afterHoursIdlingDuration: "00:05:00" });
    data.trips[1].drivingDuration = null;
    const app = loadApp(variant, data); runReport(app, variant); app.hooks.state.unitSystem = "Imperial";
    app.elements["col-work-split"].checked = true; app.hooks.filterReportByVehicle("bVan"); await app.hooks.exportActivityExcel();
    const ws = app.workbooks[0].worksheets[1], total = ws.rows.at(-1);
    assert.equal(ws.rows[0].values[11], "Distance (mi)"); assert.equal(ws.rows[0].values.length, 16);
    assert(Math.abs(total.values[11].result - 105 * 0.621371) < 1e-9); assert.equal(total.values[9].result, 30);
    assert(Math.abs(total.values[10].result - 1825 / 1440) < 1e-10); assert.equal(total.getCell(11).numFmt, "[h]:mm");
    assert.equal(total.values[12].result, 1.75); assert.equal(total.values[13].result, 35 / 60);
    app.elements["col-work-split"].checked = false; await app.hooks.exportActivityExcel();
    assert.equal(app.workbooks[1].worksheets[1].rows[0].values.length, 14);
  });
  test(variant.label + ": duplicate vehicle names remain separate by device identity", async () => {
    const data = fixture(); data.devices[1].name = data.devices[0].name;
    const app = loadApp(variant, data); const groups = runReport(app, variant); assert.equal(Object.keys(groups).length, 2);
    await app.hooks.exportActivityExcel();
    assert.deepEqual(app.workbooks[0].worksheets[1].rows.filter(r => r.outlineLevel === 0).map(r => r.values[11].result), [5, 105]);
  });
  test(variant.label + ": daily elapsed time includes gaps and ends at latest arrival", () => {
    const data = fixture(); data.trips[1].stop = "2026-10-12T14:00:00Z";
    const app = loadApp(variant, data), vehicle = runReport(app, variant).bVan, day = vehicle.days[0];
    assert.equal(app.hooks.activityDayElapsed(day), 420); assert.equal(app.hooks.activityDayEndpoints(day, app.hooks.state.activityAddrMap).end, "Stop 2");
    assert.equal(app.hooks.activityVehicleNumbers(vehicle, [])[0], 10);
    assert.equal(app.hooks.activityTotals(vehicle.trips).driving, 300, "Driving is not the elapsed span");
  });
  test(variant.label + ": overnight and DST use real elapsed hours with complete start-day trips", async () => {
    const data = fixture(); data.trips = [data.trips[0]];
    data.trips[0].start = "2026-10-24T22:00:00Z"; data.trips[0].stop = "2026-10-25T02:00:00Z";
    const app = loadApp(variant, data), day = runReport(app, variant, "2026-10-24", "2026-10-25").bVan.days[0];
    assert.equal(day.date, "2026-10-24"); assert.equal(app.hooks.activityDayElapsed(day), 240);
    const end = app.hooks.activityDayEndpoints(day, app.hooks.state.activityAddrMap);
    assert.equal(end.startTime, "23:00"); assert.equal(end.endTime, "02:00");
    await app.hooks.exportActivityExcel(); const ws = app.workbooks[0].worksheets[1];
    assert.equal(ws.rowCount, 3);
    assert.equal(ws.rows[1].values[6].result, 4); assert.equal(ws.rows[2].values[6].result, 4);
    assert.equal(ws.rows[1].values[12].toISOString(), "2026-10-24T22:00:00.000Z");
    assert.equal(ws.rows[1].values[13].toISOString(), "2026-10-25T02:00:00.000Z");
  });
  test(variant.label + ": filters update daily rows, footer and all exports", () => {
    const app = loadApp(variant); runReport(app, variant); const output = app.elements["report-output"].innerHTML;
    app.hooks.filterReportByVehicle("bOther"); const html = app.elements["report-output"].innerHTML;
    assert(!html.includes("Vehicle One")); assert(html.includes("Vehicle Two")); assert(!html.includes("105.00"));
    assert.equal(app.elements["report-summary"].innerHTML, ""); assert(app.elements["report-summary"].classList.contains("hidden"));
    app.hooks.exportActivityCsv(); assert.equal(parseCsv(app.downloads.at(-1)).length, 3);
    app.hooks.exportActivityPdf(); assert.equal(app.pdfs[0].tables.length, 1);
    app.hooks.filterReportByVehicle(""); assert.equal(app.elements["report-output"].innerHTML, output);
  });
  test(variant.label + ": all export dates stay with completed report results", async () => {
    const app = loadApp(variant); runReport(app, variant); const output = app.elements["report-output"].innerHTML;
    app.elements["filter-from"].value = "2027-02-01"; app.elements["filter-to"].value = "2027-02-28";
    app.hooks.filterReportByVehicle(""); assert.equal(app.elements["report-output"].innerHTML, output);
    app.hooks.exportActivityCsv(); assert.equal(app.filenames.at(-1), "activity_report_2026-10-07_to_2026-10-13.csv");
    app.hooks.exportActivityPdf(); assert.equal(app.pdfs[0].filename, "activity_report_2026-10-07_to_2026-10-13.pdf");
    assert(app.pdfs[0].texts.includes("Report Totals for: 2026-10-07 – 2026-10-13"));
    await app.hooks.exportActivityExcel(); assert.equal(app.filenames.at(-1), "activity_report_2026-10-07_to_2026-10-13.xlsx");
  });
  test(variant.label + ": rerun publishes trips, range and addresses together", () => {
    const app = loadApp(variant); runReport(app, variant); const call = app.hooks.state.api.call;
    let resolveAddresses, completed = false;
    app.hooks.state.api.call = (method, params, success, failure) => method === "GetAddresses" ? (resolveAddresses = () => call(method, params, success, failure)) : call(method, params, success, failure);
    app.hooks.runActivityReport("2026-10-10", "2026-10-14", () => { completed = true; });
    assert(resolveAddresses && !completed); assert.equal(app.hooks.state.activityRange, null); assert.equal(app.hooks.state.activityByDevice, null);
    assert.equal(app.hooks.state.activityAddrMap, null); assert.equal(app.hooks.state.reportData.length, 0); assert(app.elements["export-activity-btn"].classList.contains("hidden"));
    const before = app.downloads.length; app.hooks.exportActivityCsv(); assert.equal(app.downloads.length, before);
    resolveAddresses(); assert(completed); assert.deepEqual({ ...app.hooks.activityReportRange() }, { from: "2026-10-10", to: "2026-10-14" });
    assert.equal(app.hooks.state.reportData.length, 6); assert.equal(app.hooks.state.activityByDevice.bVan.days.length, 2);
    assert.equal(Object.keys(app.hooks.state.activityAddrMap).length, 6); assert(!app.elements["export-activity-btn"].classList.contains("hidden"));
  });
  test(variant.label + ": local day and date presets agree in BST", () => {
    const data = fixture(); data.trips = [data.trips[0]]; data.trips[0].start = "2026-06-30T23:30:00Z"; data.trips[0].stop = "2026-07-01T00:30:00Z";
    const app = loadApp(variant, data);
    assert.equal(app.hooks.fmtDateInput(new Date("2026-07-01T00:00:00")), "2026-07-01");
    assert.equal(runReport(app, variant, "2026-07-01", "2026-07-01").bVan.days[0].date, "2026-07-01");
    assert.equal(app.hooks.fmtActivityTime(data.trips[0].start), "00:30");
    for (const [days, from] of [["7", "2026-06-25"], ["30", "2026-06-02"]]) {
      app.elements["filter-to"].value = "2026-07-01"; app.elements[variant.privateScope ? "range-length" : "legacy-length"].value = days;
      app.hooks[variant.privateScope ? "applyRangeLength" : "applyLegacyLength"](); assert.equal(app.elements["filter-from"].value, from);
    }
  });
  test(variant.label + ": zero and unavailable stops do not invent driver groups or lose trips", () => {
    const data = fixture(); data.trips.forEach(t => { t.driver = null; t.stopDuration = null; t.idlingDuration = "00:00:00"; t.distance = 0; });
    const app = loadApp(variant, data, true), vehicle = runReport(app, variant).bVan;
    assert.equal(vehicle.trips.length, 5); assert.equal(vehicle.days.length, 2); assert.equal(app.calls.filter(c => c.params.typeName === "User").length, 0);
    assert.deepEqual({ ...app.hooks.activityTotals(vehicle.trips) }, { distance: 0, driving: 300, idle: 0, stopped: 0, engine: 300 });
    assert(!app.elements["report-output"].innerHTML.includes("Unassigned / unknown driver"));
  });
  test(variant.label + ": Demo 01 July 31 actions produce exactly one daily row in every output", async () => {
    const data = fixture(), template = data.trips[0];
    data.devices = [{ id: "bVan", name: "Demo - 01" }];
    data.trips = [{ ...template, id: "j30", start: "2026-07-30T21:37:00Z", stop: "2026-07-31T00:02:00Z", distance: 134.5 }];
    for (let i = 0; i < 12; i++) {
      const hour = String(i + 1).padStart(2, "0");
      data.trips.push({ ...template, id: "j31-" + i, driver: { id: i % 2 ? "UnknownDriverId" : "bDriverB" },
        start: "2026-07-31T" + hour + ":08:00Z", stop: "2026-07-31T" + hour + ":38:00Z", distance: 1 });
    }
    const app = loadApp(variant, data), vehicle = runReport(app, variant, "2026-07-30", "2026-07-31").bVan;
    assert.equal(vehicle.trips.length, 13); assert.equal(vehicle.days.length, 2);
    const html = app.elements["report-output"].innerHTML;
    assert.equal((html.match(/data-day='2026-07-31'/g) || []).length, 1); assert(!html.includes("<details"));
    await app.hooks.exportActivityExcel();
    const sheet = app.workbooks[0].worksheets[1];
    assert.equal(sheet.rowCount, 4, "Header + two daily rows + vehicle total, no action rows");
    assert(sheet.rows.slice(1).every(row => !row.hidden));
    assert.equal(sheet.rows[2].values[1].toISOString(), "2026-07-31T00:00:00.000Z");
    assert.equal(sheet.rows[2].values[11], 12);
    assert.equal(sheet.rows[3].values[11].result, 146.5);
    app.hooks.exportActivityCsv(); assert.equal(parseCsv(app.downloads.at(-1)).length, 4);
    app.hooks.exportActivityPdf(); assert.equal(app.pdfs[0].tables[0].body.length, 3);
  });
  test(variant.label + ": all 16 engine reporting selections align daily UI and every export", async () => {
    const ids = ["total-engine", "drive-only", "idle-only", "work-split"];
    const fixed = ["Name", "Start Date", "Start Time", "Start Location", "End Time", "End Location", "Total Hours (h)", "Total Stop Duration (h)", "Total Idle Duration (h)", "Total Driving Duration (h)"];
    const values = { "Total Hours (h)": 5, "Total Stop Duration (h)": 21, "Total Idle Duration (h)": 0.25, "Total Driving Duration (h)": 3,
      "Engine On (Total)": 195 / 1440, "Driving Time": 180 / 1440, "Idle Time": 15 / 1440, "Distance (km)": 60, "Work Engine (h)": 35 / 60, "After-hours Engine (h)": 0.5 };
    const totals = { "Total Hours (h)": 8, "Total Stop Duration (h)": 22, "Total Idle Duration (h)": 25 / 60, "Total Driving Duration (h)": 5,
      "Engine On (Total)": 325 / 1440, "Driving Time": 300 / 1440, "Idle Time": 25 / 1440, "Distance (km)": 105, "Work Engine (h)": 35 / 60, "After-hours Engine (h)": 0.5 };
    const text = { "Total Hours (h)": "5.00", "Total Stop Duration (h)": "21.00", "Total Idle Duration (h)": "0.25", "Total Driving Duration (h)": "3.00",
      "Engine On (Total)": "3:15", "Driving Time": "3:00", "Idle Time": "0:15", "Distance (km)": "60.00", "Work Engine (h)": "0.58", "After-hours Engine (h)": "0.50" };
    const duration = new Set(["Engine On (Total)", "Driving Time", "Idle Time"]);
    for (let mask = 0; mask < 16; mask++) {
      const data = fixture();
      Object.assign(data.trips[0], { workDrivingDuration: "00:30:00", workIdlingDuration: "00:05:00", afterHoursDrivingDuration: "00:30:00", afterHoursIdlingDuration: "00:00:00" });
      const app = loadApp(variant, data), selected = ids.filter((id, i) => mask & (1 << i));
      app.hooks.restoreActivityColumns(selected);
      assert.deepEqual(Array.from(app.hooks.getActivityCols()), selected, "selected ids mask " + mask);
      runReport(app, variant);
      const headers = fixed.concat(mask & 1 ? ["Engine On (Total)"] : [], mask & 2 ? ["Driving Time"] : [], mask & 4 ? ["Idle Time"] : [], ["Distance (km)"], mask & 8 ? ["Work Engine (h)", "After-hours Engine (h)"] : []);
      const html = app.elements["report-output"].innerHTML;
      const heading = html.match(/<div class='activity-summary-head[^>]*>(.*?)<\/div>/s)[1];
      assert.deepEqual(Array.from(heading.matchAll(/<span>(.*?)<\/span>/g), match => match[1]), headers.slice(1), "UI mask " + mask);
      assert(!/<(?:details|summary|table)\b/.test(html));
      app.hooks.exportActivityCsv(); const csv = parseCsv(app.downloads.at(-1));
      assert.deepEqual(csv[0], headers, "CSV headers mask " + mask);
      assert.equal(csv.length, 6); assert.deepEqual(csv[1].slice(6), headers.slice(6).map(header => text[header]));
      app.hooks.exportActivityPdf(); const pdf = app.pdfs[0].tables[0];
      assert.deepEqual(Array.from(pdf.head[1]), headers.slice(1), "PDF headers mask " + mask);
      assert.deepEqual(Array.from(pdf.body[0]), csv[1].slice(1)); assert.equal(pdf.body.length, 3);
      await app.hooks.exportActivityExcel(); const ws = app.workbooks[0].worksheets[1];
      assert.deepEqual(ws.rows[0].values.slice(0, headers.length), headers, "Excel headers mask " + mask);
      assert.equal(ws.rowCount, 6); assert(ws.rows.slice(1).every(row => !row.hidden && row.outlineLevel <= 1));
      assert.equal(ws.columns.length, headers.length + 2); assert(ws.columns.slice(-2).every(column => column.hidden));
      headers.slice(6).forEach((header, i) => {
        const column = i + 7, value = ws.rows[1].values[column - 1], numeric = value && value.formula ? value.result : value;
        assert(Math.abs(numeric - values[header]) < 1e-10, header + " daily mask " + mask);
        assert(Math.abs(ws.rows[3].values[column - 1].result - totals[header]) < 1e-10, header + " vehicle mask " + mask);
        assert.equal(ws.rows[1].getCell(column).numFmt, duration.has(header) ? "[h]:mm" : "0.00");
      });
      if (!mask) assert.equal(ws.rows[1].getCell(11).value, 60, "Distance in K must not be overwritten by an unselected engine formula");
    }
  });
  test(variant.label + ": original menu choices and saved selection arrays are preserved", () => {
    const ids = ["total-engine", "drive-only", "idle-only", "work-split"];
    const app = loadApp(variant);
    let saved;
    app.context.localStorage.setItem = (key, value) => { saved = JSON.parse(value); };
    for (let mask = 0; mask < 16; mask++) {
      const expected = ids.filter((id, i) => mask & (1 << i));
      app.hooks.restoreActivityColumns(expected);
      assert.deepEqual(Array.from(app.hooks.getActivityCols()), expected);
      app.hooks[variant.privateScope ? "saveSettings" : "saveActivitySettings"]("2026-10-07", "2026-10-13");
      assert.deepEqual(saved.cols, expected);
    }
    for (const missing of [undefined, null, {}, "invalid"]) {
      app.hooks.restoreActivityColumns(missing);
      assert.deepEqual(Array.from(app.hooks.getActivityCols()), ["total-engine"]);
    }
    app.hooks.restoreActivityColumns([]); assert.deepEqual(Array.from(app.hooks.getActivityCols()), []);
    app.hooks.restoreActivityColumns(["work-split"]); assert.deepEqual(Array.from(app.hooks.getActivityCols()), ["work-split"]);
    const html = fs.readFileSync(path.join(root, variant.privateScope ? "activity-report/index.html" : "index.html"), "utf8");
    for (const id of ids) assert.equal((html.match(new RegExp('id="col-' + id + '"', 'g')) || []).length, 1);
    assert(html.includes('id="col-total-engine" checked> Engine On (total)'));
    assert(html.includes('id="col-drive-only"> Driving Time')); assert(html.includes('id="col-idle-only"> Idle Time'));
  });
}

test("historical driver helpers stay synchronized for Legacy reports", () => {
  const app = loadApp(variants[0], { devices: [], users: [], trips: [] });
  const trips = Array.from({ length: 201 }, (_, i) => ({ driver: { id: "bDriver" + i } }));
  let done = false; app.hooks.resolveReportDrivers(trips.concat(trips), () => { done = true; });
  assert(done); assert.deepEqual(app.calls.map(c => c.params.search.userIds.length), [100, 100, 1]);
  const copies = [variants[0], variants[1], variants[3]].map(v => loadApp(v).hooks);
  for (const name of ["reportDriverId", "isUnassignedDriver", "reportUserName", "reportDriverInfo", "resolveReportDrivers", "groupReportDriverDays"]) {
    for (const copy of copies.slice(1)) assert.equal(copies[0][name].toString().replace(/\r\n/g, "\n"), copy[name].toString().replace(/\r\n/g, "\n"), name);
  }
});

test("Activity vehicle/day and worksheet helpers stay synchronized with parent", () => {
  const parent = loadApp(variants[2]).hooks, standalone = loadApp(variants[3]).hooks;
  for (const name of ["getActivityCols", "restoreActivityColumns", "buildActivityByDevice", "activityVehicleIds", "activityReportRange", "activityRangeLabel", "activityTotals", "fmtActivityTime", "activityLastTrip", "activityDayEndpoints", "activityDayElapsed", "activitySummaryColumns", "activitySummaryHeaders", "activitySummaryGridStyle", "activitySummaryNumbers", "activityNumberText", "activityDayValues", "activityVehicleNumbers", "renderActivityReport", "activityWallDate", "addActivityGroupedWorksheet"]) {
    assert.equal(parent[name].toString().replace(/\r\n/g, "\n"), standalone[name].toString().replace(/\r\n/g, "\n"), name);
  }
});

test("release manifests, badges and local asset cache versions agree", () => {
  for (const folder of ["", "legacy-trip-report", "activity-report"]) {
    const directory = path.join(root, folder), manifest = JSON.parse(fs.readFileSync(path.join(directory, "config.json"), "utf8")), html = fs.readFileSync(path.join(directory, "index.html"), "utf8");
    assert(html.includes('class="version-badge">v' + manifest.version + "</span>"), folder);
    for (const asset of ["styles.css", "logo-data.js", "app.js"]) assert(html.includes(asset + "?v=" + manifest.version), folder + "/" + asset);
  }
});
