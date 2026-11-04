"use strict";

// Real renderer smoke (no browser, no added npm dependencies).
// node tests/legacy-pdf-layout.cjs <directory containing pinned CDN scripts>
// Scripts: jspdf@2.5.1/dist/jspdf.umd.min.js and
// jspdf-autotable@3.8.2/dist/jspdf.plugin.autotable.min.js.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const root = process.env.LEGACY_SOURCE_ROOT || path.resolve(__dirname, "..");
const libraries = process.argv[2];
assert(libraries, "Pass the directory containing the pinned jsPDF/AutoTable scripts");

for (const file of ["legacy-trip-report/src/app.js", "src/app.js"]) {
  const elements = {
    "filter-from": { value: "2026-10-12" }, "filter-to": { value: "2026-10-12" }
  };
  const context = vm.createContext({ console, Date, Map, Uint8Array, TextEncoder, TextDecoder, setTimeout, clearTimeout,
    atob: value => Buffer.from(value, "base64").toString("binary"), btoa: value => Buffer.from(value, "binary").toString("base64"),
    navigator: { userAgent: "node-render-smoke" }, document: { getElementById: id => elements[id] },
    geotab: { addin: {} }, alert: message => { throw new Error(message); } });
  context.window = context; context.self = context;
  for (const library of ["jspdf.umd.min.js", "jspdf.plugin.autotable.min.js"]) {
    vm.runInContext(fs.readFileSync(path.join(libraries, library), "utf8"), context, { filename: library });
  }
  vm.runInContext(fs.readFileSync(path.join(root, "legacy-trip-report/src/filters.js"), "utf8"), context, { filename: "legacy-trip-report/src/filters.js" });
  vm.runInContext(fs.readFileSync(path.join(root, file), "utf8"), context, { filename: file });
  const RealPdf = context.jspdf.jsPDF;
  let pdf, bytes, filename;
  const cells = [], links = [], headings = [];
  context.jspdf.jsPDF = function (options) {
    pdf = new RealPdf(options);
    const autoTable = pdf.autoTable, link = pdf.link, text = pdf.text;
    pdf.text = function (value, x, y, options) {
      headings.push({ lines: Array.isArray(value) ? Array.from(value) : [value], x, y, width: pdf.getTextWidth(Array.isArray(value) ? value.join("") : value) });
      return text.call(this, value, x, y, options);
    };
    pdf.autoTable = function (tableOptions) {
      const originalHook = tableOptions.didDrawCell;
      autoTable.call(this, { ...tableOptions, didDrawCell(data) {
        if (originalHook) originalHook(data);
        cells.push({ section: data.section, column: data.column.index, width: data.cell.width,
          x: data.cell.x, height: data.cell.height, text: Array.from(data.cell.text), raw: data.cell.raw });
      } });
    };
    pdf.link = function (x, y, w, h, options) { links.push(options.url); return link.call(this, x, y, w, h, options); };
    pdf.save = name => { filename = name; bytes = Buffer.from(pdf.output("arraybuffer")); };
    return pdf;
  };
  context.state.unitSystem = "Metric";
  context.state.server = "my.geotab.com"; context.state.dbName = "synthetic-demo";
  context.state.reportDriverMap = { bDriver: { name: "Example Driver", login: "driver@example.test" } };
  const addresses = ["Short stop", "123 Long Industrial Estate Road, Warehouse Seven, North Business Park, Example City, Example County, AB12 3CD, United Kingdom".repeat(4), "Unbroken".repeat(400)];
  const trips = Array.from({ length: 40 }, (_, index) => {
    const start = new Date(Date.parse("2026-10-12T07:00:00Z") + index * 8 * 60000);
    return { id: "t" + index, device: { id: "bVan" }, driver: { id: "bDriver" }, start: start.toISOString(),
      stop: new Date(start.getTime() + 240000).toISOString(), distance: 12.345,
      drivingDuration: "00:04:00", idlingDuration: "00:01:00", stopDuration: "00:04:00", stopPoint: { x: index, y: 50 } };
  });
  const assetName = "Example Asset with a long customer-entered description and registration - North Depot Vehicle 123";
  context.state.legacyByDevice = context.buildLegacyByDevice(trips, { bVan: assetName });
  context.state.legacyAddrMap = {};
  trips.forEach((trip, index) => { context.state.legacyAddrMap[context.coordKey(trip.stopPoint)] = addresses[index] || addresses[0]; });
  context.state.legacyReportRange = { from: "2026-10-12", to: "2026-10-12" };
  context.exportLegacyTripHistoryPdf();
  assert(pdf.internal.pageSize.getWidth() < pdf.internal.pageSize.getHeight(), "Export must be portrait");
  assert(Math.abs(pdf.internal.pageSize.getWidth() - 210) < 0.01, "A4 portrait width");
  const assetHeading = headings.find(heading => heading.lines.join(" ") === assetName);
  assert(assetHeading && assetHeading.lines.length > 1, "Long asset headings wrap on portrait pages");
  pdf.setFont("helvetica", "bold"); pdf.setFontSize(12);
  assert(assetHeading.lines.every(line => pdf.getTextWidth(line) <= 182.01), "Asset heading lines fit the portrait content width");
  assert(pdf.internal.getNumberOfPages() > 1, "Exercise actual AutoTable pagination");
  const addressCells = cells.filter(cell => cell.section === "body" && cell.column === 3 && typeof cell.raw === "string" && addresses.includes(cell.raw));
  const short = addressCells.find(cell => cell.raw === addresses[0]);
  const long = addressCells.find(cell => cell.raw === addresses[1]);
  const unbroken = addressCells.filter(cell => cell.raw === addresses[2]);
  assert(long.text.length > 1 && long.height > short.height, "Long addresses wrap and grow row height");
  assert(unbroken.length > 1, "Oversize address continues across portrait pages");
  assert.equal(unbroken.flatMap(cell => cell.text).join(""), addresses[2], "Never truncate even unbroken address text");
  assert(addressCells.every(cell => Math.abs(cell.width - 33) < 0.01), "Address column never widens");
  assert(cells.every(cell => cell.x >= 13.99 && cell.x + cell.width <= 196.02), "All cells fit portrait page margins");
  assert.equal(links.length, trips.length, "Each trip retains one complete PDF URI annotation");
  assert(links.every(url => url.startsWith("https://my.geotab.com/synthetic-demo/#tripsHistory,")));
  assert(cells.some(cell => cell.raw && cell.raw.content === "Vehicle Total"), "Vehicle-period footer survives pagination");
  assert.equal(filename, "legacy_trip_history_12-10-2026_to_12-10-2026.pdf");
  assert(bytes.length > 10000 && bytes.subarray(0, 4).toString() === "%PDF");
  console.log(file + ": PASS — portrait, wrapped asset headings/address rows, oversize pagination, 40 links, totals (" + pdf.internal.getNumberOfPages() + " pages)");
}
