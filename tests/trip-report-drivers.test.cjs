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
  const devices = [{ id: "bVan", name: "Vehicle One - TEST01", groups: [{ id: "gChild" }], driver: { id: "bCurrent" } }, { id: "bOther", name: "Vehicle Two - TEST02", groups: [{ id: "gOther" }] }];
  const users = [
    { id: "bDriverA", firstName: "Alex", lastName: "Example", name: "alex@example.test", isDriver: true, driverGroups: [{ id: "gChild" }], companyGroups: [{ id: "gOther" }] },
    { id: "bDriverB", firstName: "Jamie", lastName: "Sample", name: "jamie@example.test", driverGroups: [{ id: "gOther" }], companyGroups: [{ id: "gChild" }], isDriver: false, activeTo: "2026-10-13T00:00:00Z" }
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
  const groups = [{ id: "GroupCompanyId", name: "Company", children: [{ id: "gParent" }, { id: "gOther" }] },
    { id: "gParent", name: "North", children: [{ id: "gChild" }] }, { id: "gChild", name: "North depot", children: [] },
    { id: "gOther", name: "South", children: [] }, { id: "gSecurity", name: "Security only", children: [] }];
  return { devices, users, trips, groups };
}

function testClock() {
  let now = Date.parse("2026-10-29T12:00:00Z"), nextId = 0;
  const timers = new Map();
  class ClockDate extends Date {
    constructor(...args) { super(...(args.length ? args : [now])); }
    static now() { return now; }
  }
  function advance(ms) {
    const until = now + ms;
    for (let steps = 0; steps < 10000; steps++) {
      const next = Array.from(timers).sort((a, b) => a[1].at - b[1].at || a[0] - b[0])[0];
      if (!next || next[1].at > until) { now = until; return; }
      now = next[1].at; timers.delete(next[0]); next[1].callback();
    }
    throw new Error("Timer queue did not settle");
  }
  return { Date: ClockDate, advance, get now() { return now; }, get pending() { return timers.size; },
    setTimeout(callback, delay) { const id = ++nextId; timers.set(id, { callback, at: now + Math.max(0, Number(delay) || 0) }); return id; },
    clearTimeout(id) { timers.delete(id); },
    flush() {
      for (let steps = 0; timers.size && steps < 10000; steps++) advance(Math.max(0, Math.min(...Array.from(timers.values(), t => t.at)) - now));
      assert.equal(timers.size, 0, "Timer queue did not settle");
    }
  };
}

function loadApp(variant, data = fixture(), denyDrivers = false) {
  const clock = testClock();
  const elements = {}, calls = [], downloads = [], filenames = [], pdfs = [], workbooks = [], listenerRegistrations = [], documentListeners = {};
  function element(id) {
    if (!elements[id]) elements[id] = { id, value: "", innerHTML: "", checked: false, children: [], style: {}, listeners: {},
      classList: { values: new Set(), add(value) { this.values.add(value); }, remove(value) { this.values.delete(value); },
        toggle(value, force) { const add = force === undefined ? !this.values.has(value) : force; add ? this.values.add(value) : this.values.delete(value); }, contains(value) { return this.values.has(value); } },
      attributes: {}, setAttribute(name, value) { this.attributes[name] = value; }, getAttribute(name) { return this.attributes[name]; },
      contains(target) { for (let current = target; current; current = current.parentNode) if (current === this) return true; return false; },
      focus() { context.document.activeElement = this; }, querySelector(selector) { return selector === "summary" ? element(id === "legacy-group-picker" ? "legacy-group-summary" : "legacy-entity-summary") : null; },
      addEventListener(name, handler) { listenerRegistrations.push({ id, name, handler }); this.listeners[name] = handler; }, appendChild(child) { child.parentNode = this; this.children.push(child); } };
    return elements[id];
  }
  element("filter-from").value = "2026-10-07"; element("filter-to").value = "2026-10-13";
  element("report-type").value = variant.legacy ? "legacy-trip-history" : "activity-report";
  element("col-total-engine").checked = true;
  element("legacy-length").value = "7"; element("range-length").value = "7";
  class FakeBlob { constructor(parts, options) { this.parts = parts; this.type = options.type; } }
  class FakePdf {
    constructor(options) {
      this.tables = []; this.texts = []; this.links = []; this.pages = 1;
      this.internal = { pageSize: { getWidth: () => options.orientation === "landscape" ? 297 : 210, getHeight: () => options.orientation === "landscape" ? 210 : 297 }, getNumberOfPages: () => this.pages };
      pdfs.push(this);
    }
    setFillColor() {} rect() {} setTextColor() {} setFontSize() {} setFont() {} roundedRect() {}
    text(value) { this.texts.push(value); } splitTextToSize(value) { return [value]; }
    addPage() { this.pages++; } setPage() {}
    link(x, y, width, height, options) { this.links.push({ x, y, width, height, url: options.url }); }
    autoTable(options) {
      this.tables.push(options); this.lastAutoTable = { finalY: 80 };
      if (options.didDrawCell) {
        for (const section of ["head", "body"]) {
          for (const row of options[section] || []) {
            let column = 0;
            for (const raw of row) {
              options.didDrawCell({ section, column: { index: column }, cell: { raw,
                text: [String(typeof raw === "object" ? raw.content : raw)], x: column * 20, y: 80, width: 20, height: 10 } });
              column += typeof raw === "object" && raw.colSpan || 1;
            }
          }
        }
      }
    }
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
    console, Date: clock.Date, Map, Blob: FakeBlob, Uint8Array, setTimeout: clock.setTimeout, clearTimeout: clock.clearTimeout, geotab: { addin: {} },
    localStorage: { getItem() { return null; }, setItem() {} },
    URL: { createObjectURL(blob) { downloads.push(blob); return "blob:test"; }, revokeObjectURL() {} },
    alert(message) { throw new Error(message); },
    document: { getElementById: element, addEventListener(name, handler) {
      listenerRegistrations.push({ id: "document", name, handler });
      (documentListeners[name] || (documentListeners[name] = [])).push(handler);
    }, createElement(tag) {
      return tag === "canvas" ? { getContext() { return {}; }, toDataURL() { return "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII="; } }
        : { click() { if (this.download) filenames.push(this.download); } };
    } }, jspdf: { jsPDF: FakePdf }, ExcelJS: { Workbook: FakeWorkbook }, Chart: function () { this.destroy = function () {}; }
  });
  context.window = context; context.location = { hostname: "my.geotab.com", pathname: "/demo/" };
  let source = fs.readFileSync(path.join(root, variant.file), "utf8");
  if (variant.privateScope) source = source.replace(/\}\)\(\);\s*$/, "globalThis.reportTestHooks = {state, runActivityReport, runReport, getActivityCols, restoreActivityColumns, saveSettings, filterReportByVehicle, exportActivityCsv, exportActivityPdf, exportActivityExcel, reportDriverId, isUnassignedDriver, reportUserName, reportDriverInfo, resolveReportDrivers, groupReportDriverDays, buildActivityByDevice, fmtDateInput, applyRangeLength, localDayKey, activityVehicleIds, activityReportRange, activityRangeLabel, activityTotals, fmtActivityTime, activityLastTrip, activityDayEndpoints, activityDayElapsed, activitySummaryColumns, activitySummaryHeaders, activitySummaryGridStyle, activitySummaryNumbers, activityNumberText, activityDayValues, activityVehicleNumbers, renderActivityReport, activityWallDate, addActivityGroupedWorksheet, setupReports, cancelActivityReport: typeof cancelActivityReport === 'function' ? cancelActivityReport : undefined, activityCoordinate: typeof activityCoordinate === 'function' ? activityCoordinate : undefined, activityLocationKey: typeof activityLocationKey === 'function' ? activityLocationKey : undefined, createActivityRun: typeof createActivityRun === 'function' ? createActivityRun : undefined, activityLocationCall: typeof activityLocationCall === 'function' ? activityLocationCall : undefined, resolveActivityLocations: typeof resolveActivityLocations === 'function' ? resolveActivityLocations : undefined};\n})();");
  if (!variant.privateScope) vm.runInContext(fs.readFileSync(path.join(root, "legacy-trip-report/src/filters.js"), "utf8"), context, { filename: "legacy-trip-report/src/filters.js" });
  vm.runInContext(source, context, { filename: variant.file });
  const hooks = variant.privateScope ? context.reportTestHooks : context;
  hooks.state.api = { call(method, params, success, failure) {
    calls.push({ method, params: JSON.parse(JSON.stringify(params)), at: clock.now });
    if (method === "Get" && params.typeName === "Device") success(data.devices);
    else if (method === "Get" && params.typeName === "Group") success(data.groups || []);
    else if (method === "Get" && params.typeName === "Trip") success(data.trips);
    else if (method === "Get" && params.typeName === "LogRecord") success(data.logRecords || []);
    else if (method === "Get" && params.typeName === "User") {
      if (denyDrivers && params.search.userIds) failure(new Error("User permission denied"));
      else success(data.users.filter(user => !params.search.userIds || params.search.userIds.includes(user.id)));
    } else if (method === "GetAddresses") success(params.coordinates.map(pt => ({ formattedAddress: "Stop " + pt.x })));
    else throw new Error("Unexpected API operation: " + method + " " + JSON.stringify(params));
  } };
  hooks.state.server = "my.geotab.com"; hooks.state.dbName = "demo"; hooks.state.unitSystem = "Metric";
  return { hooks, elements, calls, downloads, filenames, pdfs, workbooks, data, context, clock, listenerRegistrations, documentListeners };
}

function runReport(app, variant, from = "2026-10-07", to = "2026-10-13") {
  let completions = 0;
  app.hooks[variant.legacy ? "runLegacyTripHistoryReport" : "runActivityReport"](from, to, () => completions++);
  app.clock.flush();
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

function legacyHtmlRows(html) {
  return Array.from(html.matchAll(/<tr\b[^>]*>([\s\S]*?)<\/tr>/g), row =>
    Array.from(row[1].matchAll(/<t[dh]\b[^>]*>([\s\S]*?)<\/t[dh]>/g), cell => cell[1].replace(/<[^>]+>/g, "").trim()));
}

for (const variant of variants.filter(v => v.legacy)) {
  test(variant.label + ": legacy dates display DD-MM-YYYY without changing canonical dates or links", () => {
    const app = loadApp(variant), original = JSON.stringify(app.data.trips);
    const urls = app.data.trips.map(t => app.hooks.tripHistoryUrl(t.device.id, t)).sort();
    const groups = runReport(app, variant);
    const rows = legacyHtmlRows(app.elements["report-output"].innerHTML);
    assert(rows.some(r => r[0].startsWith("12-10-2026 — Starting from:")));
    assert(rows.some(r => r[0] === "13-10-2026 Total"));
    assert.equal(groups.bVan.driverGroups[0].days[0].date, "2026-10-12", "Grouping keys remain sortable ISO dates");
    app.hooks.exportReportCsv();
    assert.deepEqual([...new Set(parseCsv(app.downloads.at(-1)).slice(1).map(r => r[2]))].sort(), ["12-10-2026", "13-10-2026"]);
    app.hooks.exportLegacyTripHistoryPdf(); const pdf = app.pdfs[0];
    assert(pdf.texts.some(text => typeof text === "string" && text.startsWith("Created: 29-10-2026 ")));
    assert(pdf.texts.includes("Report Totals for: 07-10-2026 \u2013 13-10-2026"));
    assert(pdf.tables.every(t => /\d{2}-\d{2}-\d{4}$/.test(t.head[0][0].content)));
    assert(pdf.tables.some(t => t.body[0][0].content.startsWith("12-10-2026  Starting from:")));
    assert.equal(pdf.filename, "legacy_trip_history_07-10-2026_to_13-10-2026.pdf");
    assert.deepEqual(pdf.links.map(link => link.url).sort(), urls);
    assert.equal(app.elements["filter-from"].value, "2026-10-07"); assert.equal(app.elements["filter-to"].value, "2026-10-13");
    const search = app.calls.find(c => c.params.typeName === "Trip").params.search;
    assert.equal(search.fromDate, "2026-10-06T23:00:00.000Z"); assert.equal(search.toDate, "2026-10-13T22:59:59.000Z");
    assert.equal(JSON.stringify(app.data.trips), original);
  });
  test(variant.label + ": legacy date padding and year-boundary ordering use canonical keys", () => {
    const data = fixture(); data.trips = [
      { ...data.trips[0], start: "2026-01-01T08:00:00Z", stop: "2026-01-01T09:00:00Z" },
      { ...data.trips[0], id: "tYearEnd", start: "2025-12-31T08:00:00Z", stop: "2025-12-31T09:00:00Z" }
    ];
    const app = loadApp(variant, data);
    assert.equal(app.hooks.fmtLegacyDate("2026-01-03"), "03-01-2026");
    assert.equal(app.hooks.fmtLegacyDate(""), ""); assert.equal(app.hooks.fmtLegacyDate(null), "");
    const groups = runReport(app, variant, "2025-12-31", "2026-01-01");
    assert.deepEqual(Array.from(groups.bVan.driverGroups[0].days, day => day.date), ["2025-12-31", "2026-01-01"]);
    app.hooks.exportReportCsv(); assert.deepEqual(parseCsv(app.downloads.at(-1)).slice(1).map(row => row[2]), ["31-12-2025", "01-01-2026"]);
    const date = "2026-01-03T12:00:00Z";
    assert.equal(app.hooks.fmtDateReadable(date), new Date(date).toLocaleDateString(undefined, { day: "2-digit", month: "short", year: "numeric" }), "Other reports retain their existing date formatter");
  });
  for (const metric of [true, false]) {
    test(variant.label + ": legacy unit conversion and separate columns " + (metric ? "metric" : "imperial"), () => {
      const data = fixture(); data.trips = [data.trips[0]]; data.trips[0].distance = 50;
      const app = loadApp(variant, data); app.hooks.state.unitSystem = metric ? "Metric" : "Imperial";
      runReport(app, variant);
      const unit = metric ? "km" : "mi", distance = metric ? "50.00" : "31.07";
      assert(app.elements["report-output"].innerHTML.includes(distance + " " + unit), "50 canonical km must be displayed as " + distance + " " + unit);
      const rows = legacyHtmlRows(app.elements["report-output"].innerHTML);
      assert.deepEqual(rows[0], ["Start Time", "Distance", "Duration", "Stop Location", "Arrival Time", "Idle Duration", "Stop Duration", "Trip History"]);
      assert.equal(rows[2][1], distance + " " + unit); assert.equal(rows[2][2], "1h 00m 00s");
      assert.equal(rows[3][1], distance + " " + unit); assert.equal(rows[3][2], "1h 00m 00s");
      assert.equal(app.elements["report-summary"].innerHTML, "", "Legacy summary tiles were intentionally removed");
      assert.equal(app.hooks.legacyTotals(app.hooks.state.reportData).distance, 50);
      app.hooks.exportReportCsv(); const csv = parseCsv(app.downloads.at(-1));
      assert.equal(csv.length, 2); assert.equal(csv[0][4], "Distance (" + unit + ")"); assert.equal(csv[1][4], distance);
      app.hooks.exportLegacyTripHistoryPdf(); const pdf = app.pdfs[0], table = pdf.tables[0];
      assert.deepEqual(Array.from(table.head[1]), ["Start Time", "Distance", "Duration", "Stop Location", "Arrival Time", "Idle Duration", "Stop Duration", "Trip History"]);
      assert.equal(table.head[0][0].colSpan, 8); assert.equal(table.body[0][0].colSpan, 8);
      assert.equal(table.body[1][1], distance + " " + unit); assert.equal(table.body[1][2], "1h 00m 00s");
      assert(pdf.texts.includes("Total Distance Travelled (" + unit + ")"));
      assert.equal(data.trips[0].distance, 50, "Never convert the canonical Trip.distance in place");
    });
  }
  test(variant.label + ": legacy totals cover day, driver period and vehicle period once", () => {
    const app = loadApp(variant); runReport(app, variant);
    const rows = legacyHtmlRows(app.elements["report-output"].innerHTML);
    const driver = rows.filter(r => r[0] === "Driver Total");
    const vehicle = rows.filter(r => r[0] === "Vehicle Total");
    assert.deepEqual(driver.map(r => r.slice(1, 3)), [["40.00 km", "2h 00m 00s"], ["60.00 km", "2h 00m 00s"], ["5.00 km", "1h 00m 00s"], ["5.00 km", "1h 00m 00s"]]);
    assert.deepEqual(vehicle.map(r => r.slice(1, 3)), [["105.00 km", "5h 00m 00s"], ["5.00 km", "1h 00m 00s"]]);
    assert.equal(vehicle[0][6], "3h 00m 00s", "Overnight and unknown stop durations remain excluded");
    assert.equal(vehicle[0][7], "3 stops");
    assert.equal(rows.filter(r => /^\d{2}-\d{2}-2026 Total$/.test(r[0])).length, 5);
    app.hooks.exportLegacyTripHistoryPdf();
    const pdfRows = app.pdfs[0].tables.flatMap(t => Array.from(t.body, row => Array.from(row, c => typeof c === "object" ? c.content : c)));
    for (const label of ["12-10-2026 Total", "13-10-2026 Total", "Driver Total", "Vehicle Total"]) {
      const totals = pdfRows.filter(r => r[0] === label);
      assert(totals.length && totals.every(r => r.length === 8 && /^\d+ stops$/.test(r[7])), "Stop totals share the final Trip History column");
      assert.deepEqual(totals, rows.filter(r => r[0] === label));
    }
    for (const table of app.pdfs[0].tables) {
      for (const row of table.body) assert.equal(row.reduce((sum, c) => sum + (typeof c === "object" && c.colSpan || 1), 0), 8, "PDF spans cover all columns");
      assert.equal(table.columnStyles[7].cellWidth, 22); assert.equal(table.columnStyles[8], undefined);
    }
    app.hooks.filterReportByVehicle("bOther");
    assert.deepEqual(legacyHtmlRows(app.elements["report-output"].innerHTML).filter(r => r[0] === "Vehicle Total").map(r => r[1]), ["5.00 km"]);
    app.hooks.filterReportByVehicle(""); assert.equal(app.elements["report-output"].innerHTML.includes("105.00 km"), true);
  });
  test(variant.label + ": legacy totals retain seconds, driving fallback and unrounded km", () => {
    const data = fixture(); data.trips = data.trips.slice(0, 2);
    data.trips[0].distance = 0.006; data.trips[1].distance = 0.006;
    data.trips[0].drivingDuration = "00:00:31"; data.trips[1].drivingDuration = null;
    data.trips[1].stop = "2026-10-12T09:00:31Z";
    const app = loadApp(variant, data); app.hooks.state.unitSystem = "Imperial"; runReport(app, variant);
    const row = legacyHtmlRows(app.elements["report-output"].innerHTML).find(r => r[0] === "Vehicle Total");
    assert(row, "A vehicle-period total must be present");
    assert.equal(row[1], "0.01 mi", "Convert and round the sum, not rounded trip cells"); assert.equal(row[2], "01m 02s");
  });
  test(variant.label + ": PDF trip-history column has real per-trip URI annotations", () => {
    const app = loadApp(variant); runReport(app, variant); app.hooks.exportLegacyTripHistoryPdf();
    const pdf = app.pdfs[0];
    const expected = app.data.trips.map(t => app.hooks.tripHistoryUrl(t.device.id, t)).sort();
    assert.deepEqual(pdf.links.map(link => link.url).sort(), expected);
    assert.equal(pdf.internal.pageSize.getWidth(), 210, "Portrait A4 is required");
    assert.equal(pdf.internal.pageSize.getHeight(), 297);
    for (const table of pdf.tables) {
      assert.equal(table.styles.overflow, "linebreak");
      assert.equal(table.columnStyles[3].cellWidth, 33, "Address width is capped");
      assert.equal(Object.values(table.columnStyles).reduce((sum, column) => sum + column.cellWidth, 0), 182);
    }
    assert(pdf.links.every(link => link.width > 0 && link.height > 0 && !link.url.includes("&#39;")), "Clickable PDF links use complete raw URLs");
    const options = pdf.tables[0], raw = options.body.find(row => row[7] && row[7].tripHistoryUrl)[7];
    assert.equal(raw.content, "Trip History");
    for (const [section, column, text] of [["head", 7, ["Trip History"]], ["body", 6, ["Trip History"]], ["body", 7, []], ["body", 7, [""]], ["body", 7, [" "]]]) {
      options.didDrawCell({ section, column: { index: column }, cell: { raw, text, x: 10, y: 10, width: 22, height: 10 } });
    }
    assert.equal(pdf.links.length, expected.length, "Headers, other columns and blank page-split continuations do not create links");
    for (const table of pdf.tables) {
      for (const row of table.body.filter(row => row[7] && /^\d+ stops$/.test(row[7].content || ""))) {
        table.didDrawCell({ section: "body", column: { index: 7 }, cell: { raw: row[7], text: [row[7].content], x: 10, y: 10, width: 22, height: 10 } });
      }
    }
    assert.equal(pdf.links.length, expected.length, "Stop totals are plain counts, never clickable trip links");
    app.hooks.filterReportByVehicle("bOther"); app.hooks.exportLegacyTripHistoryPdf();
    assert.equal(app.pdfs[1].links.length, 1); assert.equal(app.pdfs[1].links[0].url, app.hooks.tripHistoryUrl("bOther", app.data.trips[5]));
  });
  test(variant.label + ": PDF never creates unusable trip-history links without a host", () => {
    const app = loadApp(variant); runReport(app, variant); app.hooks.state.server = "";
    app.hooks.exportLegacyTripHistoryPdf(); assert.equal(app.pdfs[0].links.length, 0);
    const trips = app.pdfs[0].tables.flatMap(table => Array.from(table.body)).filter(row => row[7] && typeof row[7] === "object" && "tripHistoryUrl" in row[7]);
    assert.equal(trips.length, app.data.trips.length); assert(trips.every(row => row[7].content === "\u2014"));
  });
  for (const metric of [true, false]) {
    for (const wrapped of [true, false]) {
      test(variant.label + ": viewer profile wins over database and driver units " + metric + "/" + wrapped, () => {
        const app = loadApp(variant); app.hooks.state.unitSystem = metric ? "Imperial" : "Metric";
        const requests = [];
        app.hooks.state.api = { getSession(cb) { cb(wrapped ? { credentials: { userName: "viewer@example.test" } } : { userName: "viewer@example.test" }); },
          call(method, params, success) { requests.push(params); success([{ name: "viewer@example.test", isMetric: metric }]); } };
        let done = 0; app.hooks.loadUserUnits(() => done++); app.clock.flush();
        assert.equal(done, 1); assert.equal(app.hooks.distUnit(), metric ? "km" : "mi");
        assert.equal(app.hooks.fmtDist(50, 2), metric ? "50.00 km" : "31.07 mi");
        assert.deepEqual(JSON.parse(JSON.stringify(requests)), [{ typeName: "User", search: { name: "viewer@example.test" } }]);
      });
    }
  }
}

test("standalone legacy: initialization waits for viewer's regional preferences", () => {
  const app = loadApp(variants[0]); app.hooks.state.unitSystem = "Imperial";
  app.elements.main = app.context.document.getElementById("main"); app.elements.main.classList.add("hidden");
  let completeUser, initialized = 0;
  const api = { getSession(cb) { cb({ userName: "viewer@example.test" }); }, call(method, params, success) {
    if (params.search && params.search.name) completeUser = success; else success([]);
  } };
  app.context.geotab.addin.legacyTripHistory().initialize(api, { database: "demo" }, () => initialized++);
  assert.equal(initialized, 0); assert(app.elements.main.classList.contains("hidden"));
  completeUser([{ name: "viewer@example.test", isMetric: true }]);
  assert.equal(initialized, 1); assert(!app.elements.main.classList.contains("hidden"));
  assert.equal(app.hooks.distUnit(), "km"); app.clock.flush();
});

test("Legacy aggregation, layout and distance helpers stay synchronized", () => {
  const standalone = loadApp(variants[0]).hooks, parent = loadApp(variants[1]).hooks;
  for (const name of ["fmtLegacyDate", "legacyTotals", "legacyTotalCells", "legacyHtmlTotal", "renderLegacyTripHistoryOutput", "exportLegacyTripHistoryPdf", "isMetric", "distVal", "distUnit", "fmtDist"]) {
    assert.equal(standalone[name].toString().replace(/\r\n/g, "\n"), parent[name].toString().replace(/\r\n/g, "\n"), name);
  }
});

// Synthetic control-flow fixture in the documented LogRecord shape. This does
// not stand in for the required authenticated producer/consumer smoke check.
function gpsAt(trip, role = "start", longitude = -0.1275, latitude = 51.5072) {
  return { device: { id: trip.device.id }, dateTime: role === "start" ? trip.start : trip.stop, longitude, latitude, speed: 0 };
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
    assert.deepEqual(app.calls.filter(c => c.params.typeName === "User" && c.params.search.userIds).map(c => c.params.search), [{ userIds: ["bDriverA", "bDriverB"] }]);
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
    assert(tables.some(t => t.head[0][0].content.includes("Jamie Sample | 13-10-2026")));
    assert(tables.some(t => t.head[0][0].content.includes("Alex Example | 12-10-2026")));
  });
  test(variant.label + ": vehicle filter updates rows, exports and totals without summary tiles", () => {
    const app = loadApp(variant); runReport(app, variant);
    const summary = app.elements["report-summary"].innerHTML, output = app.elements["report-output"].innerHTML;
    app.hooks.filterReportByVehicle("bOther");
    assert.equal(app.elements["report-summary"].innerHTML, "");
    assert.equal(app.hooks.legacyTotals(app.hooks.state.reportData).stops, 1);
    assert(!app.elements["report-output"].innerHTML.includes("Vehicle One - TEST01"));
    app.hooks.exportReportCsv(); assert.equal(parseCsv(app.downloads.at(-1)).length, 2);
    app.hooks.filterReportByVehicle("");
    assert.equal(app.elements["report-output"].innerHTML, output); assert.equal(app.elements["report-summary"].innerHTML, summary);
    assert.equal(app.hooks.legacyTotals(app.hooks.state.reportData).stops, 4);
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
    assert(legacyHtmlRows(app.elements["report-output"].innerHTML).some(row => row[0] === "01-07-2026 Total"));
    app.hooks.exportReportCsv(); assert.equal(parseCsv(app.downloads.at(-1))[1][2], "01-07-2026");
  });
}

for (const variant of variants.filter(v => !v.legacy)) {
  test(variant.label + ": resolves first departure through GPS and native geocoder in every output", async () => {
    const data = fixture(); data.logRecords = [gpsAt(data.trips[0])];
    const original = JSON.stringify(data.trips), app = loadApp(variant, data);
    runReport(app, variant);
    const location = "Stop -0.1275";
    app.hooks.exportActivityCsv();
    assert.equal(parseCsv(app.downloads.at(-1))[1][3], location, "First departure has a GPS-backed address without a preceding fetched trip");
    assert(app.elements["report-output"].innerHTML.includes(location));
    app.hooks.exportActivityPdf(); assert.equal(app.pdfs[0].tables[0].body[0][2], location);
    await app.hooks.exportActivityExcel(); assert.equal(app.workbooks[0].worksheets[0].rows[1].values[3], location);
    const gps = app.calls.filter(c => c.params.typeName === "LogRecord");
    assert.equal(gps.length, 2, "Only the two vehicles' missing first departures need GPS");
    assert.equal(gps[0].params.search.deviceSearch.id, "bVan");
    assert.equal(gps[0].params.search.fromDate, "2026-10-12T07:00:00.000Z");
    assert.equal(Date.parse(gps[0].params.search.toDate) - Date.parse(gps[0].params.search.fromDate), 1);
    assert.equal(gps[0].params.resultsLimit, 3); assert(gps[1].at - gps[0].at >= 100);
    const coordinates = app.calls.find(c => c.method === "GetAddresses").params.coordinates;
    assert(coordinates.some(p => p.x === -0.1275 && p.y === 51.5072), "Longitude is x; latitude is y");
    assert.deepEqual(coordinates.map(p => [p.x, p.y]).sort(), [[-0.1275, 51.5072], [3, 50], [5, 50], [6, 50]].sort());
    assert.equal(JSON.stringify(data.trips), original, "Enrichment never changes source trips or metrics");
    const calls = app.calls.length, before = app.elements["report-output"].innerHTML;
    app.hooks.filterReportByVehicle("bOther"); app.hooks.filterReportByVehicle("");
    assert.equal(app.elements["report-output"].innerHTML, before); assert.equal(app.calls.length, calls);
  });
  test(variant.label + ": missing final stop coordinates are recovered at the actual latest arrival", () => {
    const data = fixture(); data.trips[4].stopPoint = null; data.logRecords = [gpsAt(data.trips[4], "end", -0.2, 51.6)];
    const app = loadApp(variant, data), groups = runReport(app, variant);
    const endpoints = app.hooks.activityDayEndpoints(groups.bVan.days[1], app.hooks.state.activityAddrMap);
    assert.equal(endpoints.end, "Stop -0.2"); assert.equal(endpoints.endDateTime, data.trips[4].stop);
    assert(app.calls.some(c => c.params.typeName === "LogRecord" && c.params.search.fromDate === "2026-10-13T10:00:00.000Z"));
  });
  test(variant.label + ": GPS identity includes the vehicle and endpoint instant", () => {
    const data = fixture(); data.trips = [data.trips[0], { ...data.trips[5], start: data.trips[0].start }];
    data.logRecords = [gpsAt(data.trips[1], "start", -1, 52), gpsAt(data.trips[0])];
    const app = loadApp(variant, data), groups = runReport(app, variant);
    assert.equal(app.calls.filter(c => c.params.typeName === "LogRecord").length, 2);
    assert.equal(app.hooks.activityDayEndpoints(groups.bVan.days[0], app.hooks.state.activityAddrMap).start, "Stop -0.1275");
    assert.equal(app.hooks.activityDayEndpoints(groups.bOther.days[0], app.hooks.state.activityAddrMap).start, "Stop -1");
  });
  test(variant.label + ": coincident missing start and end share one GPS and address lookup", () => {
    const data = fixture(); data.trips = [{ ...data.trips[0], stop: data.trips[0].start, stopPoint: null }]; data.logRecords = [gpsAt(data.trips[0])];
    const app = loadApp(variant, data), groups = runReport(app, variant);
    assert.equal(app.calls.filter(c => c.params.typeName === "LogRecord").length, 1);
    assert.equal(app.calls.find(c => c.method === "GetAddresses").params.coordinates.length, 1);
    const endpoints = app.hooks.activityDayEndpoints(groups.bVan.days[0], app.hooks.state.activityAddrMap);
    assert.equal(endpoints.start, "Stop -0.1275"); assert.equal(endpoints.end, endpoints.start);
    assert.equal(Object.keys(app.hooks.state.activityLocations.endpoints).length, 2);
  });
  test(variant.label + ": unrelated, older and invalid GPS records cannot fabricate an address", () => {
    const data = fixture(); data.trips = [data.trips[0]];
    data.logRecords = [
      { ...gpsAt(data.trips[0]), device: { id: "bDifferent" } },
      { ...gpsAt(data.trips[0]), dateTime: "2026-10-12T06:59:59Z" },
      gpsAt(data.trips[0], "start", null, 51), gpsAt(data.trips[0], "start", -0.1, 91)
    ];
    const app = loadApp(variant, data), groups = runReport(app, variant);
    assert.equal(app.hooks.activityDayEndpoints(groups.bVan.days[0], app.hooks.state.activityAddrMap).start, "(location unknown)");
    assert.deepEqual(app.calls.find(c => c.method === "GetAddresses").params.coordinates, [{ x: 1, y: 50 }]);
  });
  test(variant.label + ": malformed stop points are excluded while valid zero coordinates are retained", () => {
    const data = fixture(); data.trips = [data.trips[0], data.trips[5]];
    data.trips[0].stopPoint = { x: "invalid", y: 50 }; data.trips[1].stopPoint = { x: 0, y: 0 };
    const app = loadApp(variant, data), groups = runReport(app, variant);
    assert.equal(app.hooks.activityDayEndpoints(groups.bVan.days[0], app.hooks.state.activityAddrMap).end, "(location unknown)");
    assert.equal(app.hooks.activityDayEndpoints(groups.bOther.days[0], app.hooks.state.activityAddrMap).end, "Stop 0");
    assert.deepEqual(app.calls.find(c => c.method === "GetAddresses").params.coordinates, [{ x: 0, y: 0 }]);
  });
  for (const stage of ["gps", "address"]) {
    for (const behavior of ["timeout", "error then late success", "duplicate success", "synchronous throw"]) {
      test(variant.label + ": " + stage + " " + behavior + " settles once and preserves other locations", () => {
        const data = fixture(); data.trips = [data.trips[0]]; data.logRecords = [gpsAt(data.trips[0])];
        const app = loadApp(variant, data), original = app.hooks.state.api.call;
        let late, completions = 0;
        app.hooks.state.api.call = (method, params, success, failure) => {
          if (stage === "gps" ? params.typeName !== "LogRecord" : method !== "GetAddresses") return original(method, params, success, failure);
          const value = stage === "gps" ? data.logRecords : params.coordinates.map(() => ({ formattedAddress: "Resolved address" }));
          late = () => success(value);
          if (behavior === "timeout") return;
          if (behavior === "synchronous throw") throw new Error("Synthetic request failure");
          if (behavior === "error then late success") { failure(new Error("Synthetic request failure")); success(value); }
          else { success(value); success(value); }
        };
        app.hooks.runActivityReport("2026-10-12", "2026-10-12", () => completions++);
        app.clock.flush(); assert.equal(completions, 1); assert.equal(app.clock.pending, 0);
        const before = JSON.stringify(app.hooks.state.activityLocations), html = app.elements["report-output"].innerHTML, calls = app.calls.length;
        if (late) late(); app.clock.flush();
        assert.equal(completions, 1); assert.equal(JSON.stringify(app.hooks.state.activityLocations), before);
        assert.equal(app.elements["report-output"].innerHTML, html); assert.equal(app.calls.length, calls);
        if (behavior === "duplicate success") assert(html.includes(stage === "gps" ? "Stop -0.1275" : "Resolved address"));
        else if (stage === "gps") { assert(html.includes("(location lookup failed)")); assert(html.includes("Stop 1")); }
        else { assert(html.includes("(address lookup failed) 51.50720, -0.12750")); assert(html.includes("(address lookup failed) 50.00000, 1.00000")); }
      });
    }
    test(variant.label + ": superseding a pending " + stage + " request preserves the newer run's controls and snapshot", () => {
      const data = fixture(); data.trips = [data.trips[0]];
      const app = loadApp(variant, data), original = app.hooks.state.api.call, pending = [];
      app.hooks.state.api.call = (method, params, success, failure) => {
        if (stage === "gps" ? params.typeName === "LogRecord" : method === "GetAddresses") pending.push({ params, success });
        else original(method, params, success, failure);
      };
      app.hooks.runReport(); app.clock.advance(1000); assert.equal(pending.length, 1);
      app.elements["filter-to"].value = "2026-10-14";
      app.hooks.runReport(); app.clock.advance(1000); assert.equal(pending.length, 2);
      const value = stage === "gps" ? [gpsAt(data.trips[0])] : [{ formattedAddress: "New address" }];
      pending[0].success(value);
      assert.equal(app.elements["run-report"].disabled, true); assert.equal(app.hooks.state.activityLocations, null);
      pending[1].success(value); app.clock.flush();
      assert.equal(app.elements["run-report"].disabled, false); assert.equal(app.hooks.state.activityRange.to, "2026-10-14");
      const html = app.elements["report-output"].innerHTML, snapshot = app.hooks.state.activityLocations;
      pending[0].success(value); app.clock.flush();
      assert.equal(app.elements["report-output"].innerHTML, html); assert.equal(app.hooks.state.activityLocations, snapshot);
    });
  }
  test(variant.label + ": leaving the add-in cancels pending requests and queued dispatch", () => {
    const app = loadApp(variant), original = app.hooks.state.api.call;
    let late;
    app.hooks.state.api.call = (method, params, success, failure) => params.typeName === "LogRecord" ? (late = success) : original(method, params, success, failure);
    app.hooks.runReport(); assert(late); assert(app.elements["run-report"].disabled);
    app.context.geotab.addin[variant.privateScope ? "activityReport" : "smartInsights"]().blur();
    const calls = app.calls.length;
    late([]); app.clock.flush();
    assert.equal(app.calls.length, calls); assert.equal(app.hooks.state.activityLocations, null);
    assert(!app.elements["run-report"].disabled); assert(app.elements["export-activity-btn"].classList.contains("hidden"));
  });
  test(variant.label + ": geocoder batches and paces only displayed endpoint coordinates", () => {
    const data = fixture(), first = data.trips[0];
    data.trips = Array.from({ length: 801 }, (_, i) => {
      const start = new Date(Date.UTC(2024, 0, i + 1, 7)), stop = new Date(start.getTime() + 3600000);
      return { ...first, id: "batch" + i, start: start.toISOString(), stop: stop.toISOString(), stopPoint: { x: -170 + i / 10, y: 50 } };
    });
    const app = loadApp(variant, data); runReport(app, variant, "2024-01-01", "2026-03-11");
    const calls = app.calls.filter(c => c.method === "GetAddresses");
    assert.deepEqual(calls.map(c => c.params.coordinates.length), [400, 400, 1]);
    for (let i = 1; i < calls.length; i++) assert(calls[i].at - calls[i - 1].at >= 150);
  });
  if (!variant.privateScope) test("parent activity: changing report type away and back cannot resurrect pending Activity output", () => {
    const app = loadApp(variant), original = app.hooks.state.api.call;
    app.hooks.setupReports(); let late;
    app.hooks.state.api.call = (method, params, success, failure) => params.typeName === "LogRecord" ? (late = success) : original(method, params, success, failure);
    app.hooks.runReport(); assert(late);
    const select = app.elements["report-type"];
    select.value = "trips"; select.listeners.change.call(select);
    select.value = "activity-report"; select.listeners.change.call(select);
    const html = app.elements["report-output"].innerHTML;
    late([]); app.clock.flush();
    assert.equal(app.elements["report-output"].innerHTML, html); assert.equal(app.hooks.state.activityLocations, null);
    assert(app.elements["export-activity-btn"].classList.contains("hidden")); assert(!app.elements["run-report"].disabled);
  });
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
    const app = loadApp(variant); runReport(app, variant); delete app.context.Chart; await app.hooks.exportActivityExcel();
    const wb = app.workbooks[0], ws = wb.worksheets[0];
    assert.deepEqual(wb.worksheets.map(sheet => sheet.name), ["Activity Report"]);
    assert.equal(wb.views[0].activeTab, 0); assert.equal(ws.rowCount, 6);
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
    const ws = app.workbooks[0].worksheets[0];
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
    const ws = app.workbooks[0].worksheets[0], total = ws.rows.at(-1);
    assert.equal(ws.rows[0].values[11], "Distance (mi)"); assert.equal(ws.rows[0].values.length, 16);
    assert(Math.abs(total.values[11].result - 105 * 0.621371) < 1e-9); assert.equal(total.values[9].result, 30);
    assert(Math.abs(total.values[10].result - 1825 / 1440) < 1e-10); assert.equal(total.getCell(11).numFmt, "[h]:mm");
    assert.equal(total.values[12].result, 1.75); assert.equal(total.values[13].result, 35 / 60);
    app.elements["col-work-split"].checked = false; await app.hooks.exportActivityExcel();
    assert.equal(app.workbooks[1].worksheets[0].rows[0].values.length, 14);
  });
  test(variant.label + ": duplicate vehicle names remain separate by device identity", async () => {
    const data = fixture(); data.devices[1].name = data.devices[0].name;
    const app = loadApp(variant, data); const groups = runReport(app, variant); assert.equal(Object.keys(groups).length, 2);
    await app.hooks.exportActivityExcel();
    assert.deepEqual(app.workbooks[0].worksheets[0].rows.filter(r => r.outlineLevel === 0).map(r => r.values[11].result), [5, 105]);
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
    await app.hooks.exportActivityExcel(); const ws = app.workbooks[0].worksheets[0];
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
    app.clock.advance(1000);
    assert(resolveAddresses && !completed); assert.equal(app.hooks.state.activityRange, null); assert.equal(app.hooks.state.activityByDevice, null);
    assert.equal(app.hooks.state.activityAddrMap, null); assert.equal(app.hooks.state.reportData.length, 0); assert(app.elements["export-activity-btn"].classList.contains("hidden"));
    const before = app.downloads.length; app.hooks.exportActivityCsv(); assert.equal(app.downloads.length, before);
    resolveAddresses(); assert(completed); assert.deepEqual({ ...app.hooks.activityReportRange() }, { from: "2026-10-10", to: "2026-10-14" });
    assert.equal(app.hooks.state.reportData.length, 6); assert.equal(app.hooks.state.activityByDevice.bVan.days.length, 2);
    assert.deepEqual(Object.keys(app.hooks.state.activityAddrMap).sort(), ["50.00000,3.00000", "50.00000,5.00000", "50.00000,6.00000"]);
    assert(!app.elements["export-activity-btn"].classList.contains("hidden"));
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
    const sheet = app.workbooks[0].worksheets[0];
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
      delete app.context.Chart;
      await app.hooks.exportActivityExcel(); const ws = app.workbooks[0].worksheets[0];
      assert.deepEqual(app.workbooks[0].worksheets.map(sheet => sheet.name), ["Activity Report"], "Single worksheet, mask " + mask);
      assert.deepEqual(ws.rows[0].values.slice(0, headers.length), headers, "Excel headers mask " + mask);
      assert.equal(ws.rowCount, 6); assert(ws.rows.slice(1).every(row => !row.hidden && row.outlineLevel <= 1));
      assert(ws.rows[0].font.bold, "Keep the Excel header bold");
      for (const row of ws.rows.slice(1)) {
        assert.equal(row.font.bold, row.outlineLevel === 0, "Only vehicle-total data rows are bold, mask " + mask);
        for (let c = 1; c <= headers.length; c++) {
          const font = row.getCell(c).font || row.font;
          assert.equal(font.bold, row.outlineLevel === 0, "Effective cell weight, mask " + mask + ", row " + row.number + ", column " + c);
        }
      }
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
  for (const name of ["getActivityCols", "restoreActivityColumns", "activityCoordinate", "activityLocationKey", "cancelActivityReport", "createActivityRun", "activityLocationCall", "resolveActivityLocations", "buildActivityByDevice", "activityVehicleIds", "activityReportRange", "activityRangeLabel", "activityTotals", "fmtActivityTime", "activityLastTrip", "activityDayEndpoints", "activityDayElapsed", "activitySummaryColumns", "activitySummaryHeaders", "activitySummaryGridStyle", "activitySummaryNumbers", "activityNumberText", "activityDayValues", "activityVehicleNumbers", "renderActivityReport", "activityWallDate", "addActivityGroupedWorksheet"]) {
    assert.equal(parent[name].toString().replace(/\r\n/g, "\n"), standalone[name].toString().replace(/\r\n/g, "\n"), name);
  }
});

// Synthetic catalog/control-flow fixtures; authenticated MyGeotab/native matching
// remains a separate smoke check, not something these mocks establish.
for (const variant of variants.filter(v => v.legacy)) {
  function setup(data) { const app = loadApp(variant, data); app.hooks.setupLegacyFilters(); return app; }
  function ids(app) { return Array.from(app.hooks.state.reportData, trip => trip.id).sort(); }
  function change(app, id, value) { const el = app.elements[id]; el.value = value; el.listeners.change.call(el); }
  test(variant.label + ": groups include descendants, union memberships and constrain choices before Run", () => {
    const app = setup(), f = app.hooks.legacyFilters();
    assert.equal(f.mode, "assets"); assert.equal(f.ready, true);
    assert(!app.elements["legacy-group-options"].innerHTML.includes("Security only"));
    f.groups = ["gParent"]; app.hooks.renderLegacyChoices();
    assert.deepEqual(Array.from(app.hooks.legacyEntities(), entity => entity.id), ["bVan"]);
    runReport(app, variant); assert.deepEqual(ids(app), ["t1", "t2", "t3", "t4", "t5"]);
    f.groups.push("gOther"); app.hooks.applyLegacyFilters(); assert.equal(ids(app).length, 6);
    f.groups = ["gMissing"]; app.hooks.applyLegacyFilters(); assert.deepEqual(ids(app), []);
    assert(app.elements["export-pdf-legacy"].classList.contains("hidden"));
  });
  test(variant.label + ": driver home groups, not company/data-access groups, scope historical trips", () => {
    const app = setup(), f = app.hooks.legacyFilters();
    change(app, "legacy-mode", "drivers"); f.groups = ["gParent"];
    runReport(app, variant);
    assert.deepEqual(ids(app), ["t1", "t3", "t6"], "Driver scope follows home group even in another asset group");
    f.groups = ["gOther"]; app.hooks.applyLegacyFilters(); assert.deepEqual(ids(app), ["t2", "t4"], "Archived driver is retained");
    f.groups = []; app.hooks.applyLegacyFilters(); assert.equal(ids(app).length, 6);
    assert(app.hooks.legacyEntities().some(entity => entity.id === "unassigned"));
  });
  test(variant.label + ": select all minus one survives reruns and asset/driver mode changes", () => {
    const app = setup();
    app.hooks.setLegacySelected("bVan", false); runReport(app, variant);
    assert.deepEqual(ids(app), ["t6"]);
    runReport(app, variant); assert.deepEqual(ids(app), ["t6"]);
    change(app, "legacy-mode", "drivers"); app.hooks.setLegacySelected("id:bDriverA", false); app.hooks.applyLegacyFilters();
    assert.deepEqual(ids(app), ["t2", "t4", "t5"]);
    change(app, "legacy-mode", "assets"); assert.deepEqual(ids(app), ["t6"]);
    change(app, "legacy-mode", "drivers"); runReport(app, variant); assert.deepEqual(ids(app), ["t2", "t4", "t5"]);
  });
  test(variant.label + ": select-all checkbox supports search subsets, partial selection and selecting none", () => {
    const app = setup(); runReport(app, variant);
    app.elements["legacy-entity-search"].value = "Vehicle One";
    const all = app.elements["legacy-select-all"];
    all.checked = false; all.listeners.change.call(all);
    assert.deepEqual(ids(app), ["t6"]);
    app.elements["legacy-entity-search"].value = ""; app.hooks.renderLegacyChoices();
    assert.equal(all.indeterminate, true); assert.equal(all.checked, false);
    all.checked = true; all.listeners.change.call(all); assert.equal(ids(app).length, 6);
    all.checked = false; all.listeners.change.call(all); assert.deepEqual(ids(app), []);
    runReport(app, variant); assert.deepEqual(ids(app), [], "Empty selection must never mean all");
    assert(app.elements["export-csv"].classList.contains("hidden"));
    app.hooks.exportReportCsv(); assert.equal(app.downloads.length, 0);
  });
  test(variant.label + ": driver filtering preserves handover predecessors and full-timeline overnight decisions in exports", () => {
    const app = setup(), f = app.hooks.legacyFilters();
    f.mode = "drivers"; f.drivers = { all: false, ids: { "id:bDriverA": true } };
    runReport(app, variant);
    const row = app.hooks.state.legacyByDevice.bVan.driverGroups[0].days[0].rows[1];
    assert.equal(row.prevTrip.id, "t2"); assert.equal(app.hooks.legacyStopMins(row.trip), null);
    app.hooks.exportReportCsv();
    const csv = parseCsv(app.downloads[0]); assert.equal(csv.length, 4); assert(csv.slice(1).every(r => r[1] === "Alex Example"));
    assert.equal(app.hooks.state.legacyStopInfo.get(row.trip).excluded, true, "CSV must not rebuild a filtered timeline");
    app.hooks.exportLegacyTripHistoryPdf(); assert.equal(app.pdfs[0].links.length, 3);
    f.drivers = { all: false, ids: { "id:bDriverB": true } }; app.hooks.applyLegacyFilters();
    assert(app.elements["report-output"].innerHTML.includes("Starting from: Stop 1"));
    assert.equal(app.hooks.state.legacyByDevice.bVan.driverGroups[0].days[0].rows[0].prevTrip.id, "t1");
  });
  test(variant.label + ": duplicate labels do not merge selected assets or drivers", () => {
    const data = fixture(); data.devices[1].name = data.devices[0].name; data.users[1].firstName = "Alex"; data.users[1].lastName = "Example";
    const app = setup(data), f = app.hooks.legacyFilters();
    assert(app.hooks.legacyEntities().every(entity => entity.name.includes(entity.id)));
    f.mode = "drivers"; f.drivers = { all: false, ids: { "id:bDriverB": true } };
    runReport(app, variant); assert.deepEqual(ids(app), ["t2", "t4"]);
  });
  test(variant.label + ": completed export dates remain tied to completed data", () => {
    const app = setup(); runReport(app, variant);
    app.elements["filter-from"].value = "2026-11-01"; app.elements["filter-to"].value = "2026-11-02";
    app.hooks.exportLegacyTripHistoryPdf();
    assert.equal(app.pdfs[0].filename, "legacy_trip_history_07-10-2026_to_13-10-2026.pdf");
  });
  test(variant.label + ": empty/error reruns retain choices but never expose previous results", () => {
    const app = setup(), f = app.hooks.legacyFilters();
    app.hooks.setLegacySelected("bOther", false); runReport(app, variant);
    app.data.trips = []; runReport(app, variant);
    assert.equal(app.hooks.legacyFilters(), f); assert.deepEqual(ids(app), []);
    assert(app.elements["export-pdf-legacy"].classList.contains("hidden"));
    const original = app.hooks.state.api.call;
    app.hooks.state.api.call = (method, params, success, failure) => params.typeName === "Trip" ? failure(new Error("Trip denied")) : original(method, params, success, failure);
    runReport(app, variant);
    assert.equal(app.hooks.state.legacyByDevice, null); assert.equal(app.hooks.state.legacyReportRange, null);
    assert.equal(app.hooks.legacySelected("bOther"), false);
  });
  test(variant.label + ": late run callbacks cannot mutate a newer snapshot or dispatch more lookups", () => {
    const app = setup(), original = app.hooks.state.api.call;
    let late, first = true;
    app.hooks.state.api.call = (method, params, success, failure) => {
      if (params.typeName === "User" && params.search.userIds && first) { first = false; late = success; }
      else original(method, params, success, failure);
    };
    let oldDone = 0; app.hooks.runLegacyTripHistoryReport("2026-10-07", "2026-10-13", () => oldDone++);
    app.hooks.setLegacySelected("bVan", false); runReport(app, variant);
    const before = app.calls.length, snapshot = app.hooks.state.legacyByDevice;
    late([{ id: "bDriverA", firstName: "Stale", lastName: "User" }]); app.clock.flush();
    assert.equal(app.calls.length, before); assert.equal(oldDone, 0); assert.equal(app.hooks.state.legacyByDevice, snapshot);
    assert.equal(app.hooks.state.reportDriverMap.bDriverA.name, "Alex Example"); assert.deepEqual(ids(app), ["t6"]);
  });
  test(variant.label + ": catalog errors/timeouts fail visibly and Retry preserves choices", () => {
    const app = loadApp(variant), original = app.hooks.state.api.call;
    app.hooks.setLegacySelected("bOther", false);
    app.hooks.state.api.call = (method, params, success, failure) => params.typeName === "Group" ? undefined : original(method, params, success, failure);
    app.hooks.setupLegacyFilters(); app.clock.flush();
    assert.equal(app.hooks.legacyFilters().ready, false);
    assert(!app.elements["legacy-filter-retry"].classList.contains("hidden"));
    app.hooks.state.api.call = original; runReport(app, variant); assert.deepEqual(ids(app), ["t1", "t2", "t3", "t4", "t5"]);
  });
  test(variant.label + ": initialization boundary clears selections/catalogs and rejects old-identity callbacks", () => {
    const app = loadApp(variant), original = app.hooks.state.api.call;
    let late;
    app.hooks.state.api.call = (method, params, success, failure) => params.typeName === "Group" ? late = success : original(method, params, success, failure);
    app.hooks.setupLegacyFilters(); app.hooks.setLegacySelected("bVan", false);
    app.hooks.resetLegacyPage(); app.hooks.state.api.call = original;
    app.data.devices[0].name = "New identity asset";
    runReport(app, variant); const snapshot = app.hooks.state.legacyByDevice;
    late([{ id: "gOther", name: "Old database group" }]); app.clock.flush();
    assert.equal(app.hooks.state.legacyByDevice, snapshot); assert.equal(app.hooks.legacySelected("bVan"), true);
    assert.equal(snapshot.bVan.name, "New identity asset"); assert(!app.elements["legacy-group-options"].innerHTML.includes("Old database group"));
  });
  test(variant.label + ": initialization removes previous identity content before new data arrives", () => {
    const app = setup(); runReport(app, variant);
    app.elements["db-name"] = app.context.document.getElementById("db-name"); app.elements["db-name"].textContent = "old-database";
    app.hooks.resetLegacyPage();
    assert(!app.elements["report-output"].innerHTML.includes("TEST01"), "Old asset/driver/address preview must disappear immediately");
    assert.equal(app.elements["report-summary"].innerHTML, "");
    assert(app.elements["report-summary"].classList.contains("hidden"));
    assert.equal(app.elements["db-name"].textContent, "");
    assert.deepEqual(Object.keys(app.hooks.state.deviceMap), []);
    assert.equal(app.elements["legacy-group-options"].innerHTML, "");
    assert.equal(app.elements["legacy-entity-options"].innerHTML, "");
  });
  test(variant.label + ": obsolete initialization cannot change a newer viewer's units or continue setup", () => {
    const app = loadApp(variant);
    let late, setupCount = 0, completedA = 0, completedB = 0;
    if (variant.label === "parent legacy") {
      for (const name of ["setupNav", "applyUnitLabels", "restoreActivitySettings", "setupEditMode", "setupWidgetPicker", "setupModal", "setupParamEditor", "setupMap", "setupIncident", "loadSuggestions"]) app.hooks[name] = () => {};
      app.hooks.setupReports = () => setupCount++;
      app.hooks.restoreDashboard = done => done();
    }
    function api(viewer, pending) { return { getSession(done) { done({ userName: viewer }); }, call(method, params, success) {
      if (params.search && params.search.name) { if (pending) late = success; else success([{ name: viewer, isMetric: false }]); }
      else success([]);
    } }; }
    const addin = app.context.geotab.addin[variant.label === "parent legacy" ? "smartInsights" : "legacyTripHistory"]();
    addin.initialize(api("old@example.test", true), { database: "old" }, () => completedA++);
    addin.initialize(api("new@example.test", false), { database: "new" }, () => completedB++);
    assert.equal(app.hooks.distUnit(), "mi");
    late([{ name: "old@example.test", isMetric: true }]); app.clock.flush();
    assert.equal(app.hooks.distUnit(), "mi", "Old metric user cannot overwrite the current Imperial user");
    assert.equal(completedB, 1); assert.equal(completedA, variant.label === "parent legacy" ? 1 : 0);
    if (variant.label === "parent legacy") assert.equal(setupCount, 1, "Only current initialization runs asynchronous setup");
  });
  test(variant.label + ": every Run refreshes asset and group memberships without resetting exclusions", () => {
    const app = setup(), f = app.hooks.legacyFilters(); runReport(app, variant);
    app.hooks.setLegacySelected("bOther", false);
    app.data.devices = app.data.devices.concat({ id: "bNew", name: "Newly added asset", groups: [{ id: "gChild" }] });
    app.data.trips = app.data.trips.concat({ ...app.data.trips[0], id: "t7", device: { id: "bNew" } });
    runReport(app, variant);
    assert.deepEqual(ids(app), ["t1", "t2", "t3", "t4", "t5", "t7"], "All assets includes newly created assets while keeping the excluded asset out");
    assert.equal(app.hooks.legacyFilters(), f); assert.equal(app.hooks.legacySelected("bOther"), false);
    f.groups = ["gParent"];
    app.data.devices = app.data.devices.map(device => device.id === "bVan" ? { ...device, groups: [{ id: "gOther" }] } : device);
    runReport(app, variant); assert.deepEqual(ids(app), ["t7"], "Group membership is refreshed too");
    app.data.groups = app.data.groups.map(group => group.id === "gParent" ? { ...group, children: [] } : group);
    runReport(app, variant); assert.deepEqual(ids(app), [], "Removed child relationships cannot remain cached");
  });
}

test("parent: reinitializing a selected Legacy report reloads pre-run selectors without a report-type change", () => {
  const app = loadApp(variants[1]), original = app.hooks.state.api.call;
  app.hooks.setupReports(); runReport(app, variants[1]);
  app.hooks.setLegacySelected("bOther", false);
  for (const name of ["setupNav", "applyUnitLabels", "restoreActivitySettings", "setupEditMode", "setupWidgetPicker", "setupModal", "setupParamEditor", "setupMap", "setupIncident", "loadSuggestions"]) app.hooks[name] = () => {};
  app.hooks.restoreDashboard = done => done();
  app.data.devices = app.data.devices.map(device => ({ ...device, name: "New database " + device.name }));
  const api = { getSession(done) { done({ userName: "new@example.test" }); }, call(method, params, success, failure) {
    if (params.search && params.search.name) success([{ name: "new@example.test", isMetric: false }]);
    else original(method, params, success, failure);
  } };
  app.context.geotab.addin.smartInsights().initialize(api, { database: "new" }, () => {});
  app.clock.flush();
  assert(app.hooks.state.legacyFilters && app.hooks.state.legacyFilters.ready, "Pre-run catalogs must load during initialization, not only after a change event");
  assert(app.elements["legacy-entity-options"].innerHTML.includes("New database Vehicle One"));
  assert(app.elements["legacy-group-options"].innerHTML.includes("North depot"));
  assert.equal(app.hooks.legacySelected("bOther"), true, "A new identity does not inherit old exclusions");
});

test("parent: pending non-Legacy report cannot overwrite Legacy after a report-type switch", () => {
  const app = loadApp(variants[1]), original = app.hooks.state.api.call;
  app.hooks.setupReports(); app.elements["report-type"].value = "trips";
  let late;
  app.hooks.state.api.call = (method, params, success, failure) => params.typeName === "Trip" ? late = success : original(method, params, success, failure);
  app.hooks.runReport();
  assert.equal(app.elements["run-report"].disabled, true);
  app.elements["report-type"].value = "legacy-trip-history";
  app.elements["report-type"].listeners.change.call(app.elements["report-type"]);
  assert.equal(app.elements["run-report"].disabled, false, "Cancellation must make Run clickable for the newly selected report");
  app.hooks.state.api.call = original; app.hooks.runReport(); app.clock.flush();
  assert.equal(app.elements["run-report"].disabled, false);
  const snapshot = app.hooks.state.legacyByDevice, html = app.elements["report-output"].innerHTML;
  late([]); assert.equal(app.hooks.state.legacyByDevice, snapshot); assert.equal(app.elements["report-output"].innerHTML, html);
});

for (const variant of variants.filter(v => v.legacy)) {
  function groupedFixture() {
    const data = fixture();
    data.groups[0].children.push({ id: "GroupAssetTypeId" }, { id: "GroupAssetPurposeId" }, { id: "bCustomType" });
    data.groups.push({ id: "GroupAssetTypeId", name: "Tipo de activo", children: [{ id: "GroupVehicleId" }, { id: "bSystemChild" }] },
      { id: "GroupVehicleId", name: "Vehicle", children: [] }, { id: "GroupAssetPurposeId", name: "Asset Purpose", children: [] },
      { id: "bSystemChild", name: "Customer extension of asset type", children: [] }, { id: "bCustomType", name: "Asset Purpose", children: [] });
    return data;
  }
  function setup(data = groupedFixture()) { const app = loadApp(variant, data); app.hooks.setupLegacyFilters(); return app; }
  test(variant.label + ": customer groups lead, built-in branches are collapsed and names do not classify groups", () => {
    const app = setup();
    const sections = app.hooks.legacyGroupSections();
    assert(sections.custom.some(group => group.id === "bCustomType"), "A customer-created namesake remains a customer group");
    for (const id of ["GroupCompanyId", "GroupAssetTypeId", "GroupAssetPurposeId", "GroupVehicleId", "bSystemChild"]) {
      assert(sections.system.some(group => group.id === id), id + " remains available in the system section");
      assert(!sections.custom.some(group => group.id === id));
    }
    const html = app.elements["legacy-group-options"].innerHTML;
    const top = html.slice(0, html.indexOf("id='legacy-system-toggle'"));
    assert(top.includes("data-group='gParent'")); assert(!top.includes("data-group='GroupAssetTypeId'"));
    assert(html.includes("aria-expanded='false'")); assert(html.includes("id='legacy-system-options' class='hidden'"));
    app.elements["legacy-group-search"].value = "Tipo"; app.hooks.renderLegacyGroups();
    assert(app.elements["legacy-group-options"].innerHTML.includes("No matching customer groups."));
    assert.equal(app.elements["legacy-system-toggle"].textContent, "Show system groups (1 match)");
    assert.equal(app.elements["legacy-system-toggle"].attributes["aria-expanded"], "false");
  });
  test(variant.label + ": system leaf classification also works with a restricted partial group tree", () => {
    const data = groupedFixture();
    data.groups = data.groups.filter(group => group.id === "GroupVehicleId" || group.id === "gChild");
    data.devices[0].groups.push({ id: "GroupVehicleId" });
    const app = setup(data), sections = app.hooks.legacyGroupSections();
    assert.deepEqual(Array.from(sections.system, group => group.id), ["GroupVehicleId"]);
    assert.deepEqual(Array.from(sections.custom, group => group.id), ["gChild"]);
  });
  test(variant.label + ": expanding/collapsing system choices preserves filters and Company is not All groups", () => {
    const app = setup(), f = app.hooks.legacyFilters();
    const toggle = app.context.document.getElementById("legacy-system-toggle");
    const options = app.elements["legacy-group-options"];
    options.listeners.click({ target: toggle });
    assert.equal(toggle.attributes["aria-expanded"], "true");
    f.groups = ["GroupCompanyId"]; runReport(app, variant);
    assert.equal(app.hooks.state.reportData.length, 6, "Company's descendant asset groups are unchanged");
    f.mode = "drivers"; app.hooks.applyLegacyFilters();
    assert.equal(app.hooks.state.reportData.length, 5, "Company excludes unassigned drivers");
    options.listeners.click({ target: toggle });
    assert.equal(toggle.attributes["aria-expanded"], "false");
    assert.equal(f.groups[0], "GroupCompanyId");
    f.groups = []; app.hooks.applyLegacyFilters(); assert.equal(app.hooks.state.reportData.length, 6, "All groups remains unrestricted");
  });
  test(variant.label + ": outside clicks close pickers without suppressing inside label/checkbox clicks", () => {
    const app = setup(), group = app.elements["legacy-group-picker"], entity = app.elements["legacy-entity-picker"];
    const groupOptions = app.elements["legacy-group-options"]; groupOptions.parentNode = group;
    const label = app.context.document.getElementById("test-group-label"); label.parentNode = groupOptions;
    const checkbox = app.context.document.getElementById("test-group-checkbox"); checkbox.parentNode = label;
    checkbox.attributes["data-group"] = "gParent"; checkbox.checked = true;
    const click = target => (app.documentListeners.click || []).forEach(handler => handler({ target }));
    group.open = true; entity.open = true; click(label);
    assert.equal(group.open, true); assert.equal(entity.open, false);
    groupOptions.listeners.change({ target: checkbox }); click(checkbox);
    assert.equal(group.open, true); assert.equal(app.hooks.legacyFilters().groups[0], "gParent");
    const summary = app.elements["legacy-entity-summary"]; summary.parentNode = entity;
    entity.open = true; click(summary); assert.equal(group.open, false); assert.equal(entity.open, true);
    click(app.context.document.getElementById("run-report")); assert.equal(entity.open, false);
    assert.equal(app.hooks.legacyFilters().groups[0], "gParent");
    entity.open = true; entity.listeners.keydown.call(entity, { key: "Escape" });
    assert.equal(entity.open, false); assert.equal(app.context.document.activeElement, summary);
    app.hooks.setupLegacyFilters(); app.hooks.resetLegacyPage(); app.hooks.setupLegacyFilters();
    assert.equal(app.listenerRegistrations.filter(registration => registration.id === "document" && registration.name === "click").length, 1);
  });
  test(variant.label + ": closed-picker catalog errors remain visible and announced without creating a toolbar row", () => {
    const app = loadApp(variant, groupedFixture()), original = app.hooks.state.api.call;
    app.hooks.state.api.call = (method, params, success, failure) => params.typeName === "Group" ? failure(new Error("Groups denied")) : original(method, params, success, failure);
    app.hooks.setupLegacyFilters();
    assert.match(app.context.document.getElementById("legacy-group-summary").textContent || "", /unavailable/i);
    assert.match(app.elements["legacy-filter-status"].textContent, /unavailable/i);
    const html = app.elements["legacy-filters"].innerHTML;
    assert(html.indexOf("id='legacy-filter-status'") > html.lastIndexOf("</details>"), "Live status is outside the collapsed panels");
    assert(html.includes("id='legacy-filter-status' class='legacy-sr-only'"));
    app.hooks.state.api.call = original; app.elements["legacy-filter-retry"].listeners.click();
    assert.equal(app.elements["legacy-group-summary"].textContent, "All groups");
  });
  test(variant.label + ": Legacy summary tiles are removed while table and PDF totals remain", () => {
    const app = setup(); runReport(app, variant);
    assert.equal(app.elements["report-summary"].innerHTML, ""); assert(app.elements["report-summary"].classList.contains("hidden"));
    assert(app.elements["report-output"].innerHTML.includes("Vehicle Total"));
    assert.equal(app.hooks.legacyTotals(app.hooks.state.reportData).distance, 110);
    app.hooks.exportLegacyTripHistoryPdf(); assert(app.pdfs[0].texts.includes("Total Distance Travelled (km)"));
  });
}

test("parent: other report summary tiles are unaffected by Legacy-only removal", () => {
  const app = loadApp(variants[1]); app.elements["report-type"].value = "trips";
  app.hooks.runTripsReport("2026-10-07", "2026-10-13", () => {});
  assert(app.elements["report-summary"].innerHTML.includes("summary-card"));
  assert(!app.elements["report-summary"].classList.contains("hidden"));
  app.elements["report-type"].value = "legacy-trip-history";
  runReport(app, variants[1]); assert.equal(app.elements["report-summary"].innerHTML, "");
});

test("release manifests, badges and local asset cache versions agree", () => {
  for (const folder of ["", "legacy-trip-report", "activity-report"]) {
    const directory = path.join(root, folder), manifest = JSON.parse(fs.readFileSync(path.join(directory, "config.json"), "utf8")), html = fs.readFileSync(path.join(directory, "index.html"), "utf8");
    assert(html.includes('class="version-badge">v' + manifest.version + "</span>"), folder);
    for (const asset of ["styles.css", "logo-data.js", "app.js"]) assert(html.includes(asset + "?v=" + manifest.version), folder + "/" + asset);
  }
});
