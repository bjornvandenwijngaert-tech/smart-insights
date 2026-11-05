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
 for (const zones of [false, true]) {
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
  const logoFile = file.startsWith("legacy-trip-report/") ? "legacy-trip-report/src/logo-data.js" : "src/logo-data.js";
  vm.runInContext(fs.readFileSync(path.join(root, logoFile), "utf8"), context, { filename: logoFile });
  vm.runInContext(fs.readFileSync(path.join(root, "legacy-trip-report/src/zones.js"), "utf8"), context, { filename: "legacy-trip-report/src/zones.js" });
  vm.runInContext(fs.readFileSync(path.join(root, "legacy-trip-report/src/zone-preference.js"), "utf8"), context, { filename: "legacy-trip-report/src/zone-preference.js" });
  vm.runInContext(fs.readFileSync(path.join(root, "legacy-trip-report/src/filters.js"), "utf8"), context, { filename: "legacy-trip-report/src/filters.js" });
  vm.runInContext(fs.readFileSync(path.join(root, file), "utf8"), context, { filename: file });
  const RealPdf = context.jspdf.jsPDF;
  let pdf, bytes, filename;
  const cells = [], links = [], headings = [], rectangles = [], kpiBoxes = [], logos = [], tableLayouts = [];
  context.jspdf.jsPDF = function (options) {
    pdf = new RealPdf(options);
    const autoTable = pdf.autoTable, link = pdf.link, text = pdf.text, rect = pdf.rect, roundedRect = pdf.roundedRect, addImage = pdf.addImage;
    pdf.rect = function (x, y, w, h, style) { rectangles.push({ x, y, w, h, style }); return rect.call(this, x, y, w, h, style); };
    pdf.roundedRect = function (x, y, w, h, rx, ry, style) {
      kpiBoxes.push({ x, y, w, h }); return roundedRect.call(this, x, y, w, h, rx, ry, style);
    };
    pdf.addImage = function (image, format, x, y, w, h) {
      const result = addImage.call(this, image, format, x, y, w, h);
      logos.push({ format, x, y, w, h }); return result;
    };
    pdf.text = function (value, x, y, options) {
      headings.push({ lines: Array.isArray(value) ? Array.from(value) : [value], x, y,
        page: pdf.internal.getCurrentPageInfo().pageNumber,
        width: pdf.getTextWidth(Array.isArray(value) ? value.join("") : value) });
      return text.call(this, value, x, y, options);
    };
    pdf.autoTable = function (tableOptions) {
      tableLayouts.push({ margin: tableOptions.margin, tableWidth: tableOptions.tableWidth, head: tableOptions.head });
      const originalHook = tableOptions.didDrawCell;
      autoTable.call(this, { ...tableOptions, didDrawCell(data) {
        if (originalHook) originalHook(data);
        cells.push({ section: data.section, column: data.column.index, width: data.cell.width,
          x: data.cell.x, y: data.cell.y, height: data.cell.height, page: data.pageNumber,
          text: Array.from(data.cell.text), raw: data.cell.raw });
      } });
    };
    pdf.link = function (x, y, w, h, options) { links.push(options.url); return link.call(this, x, y, w, h, options); };
    pdf.save = name => { filename = name; bytes = Buffer.from(pdf.output("arraybuffer")); };
    return pdf;
  };
  context.state.unitSystem = "Metric";
  context.legacyFilters().zones = zones;
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
  const zoneName = "Customer depot and distribution centre with a deliberately long zone name ".repeat(3).trim();
  context.state.legacyZoneSnapshot = { status: "ready", labels: {} };
  trips.forEach(trip => { context.state.legacyZoneSnapshot.labels[context.legacyZoneKey(trip.stopPoint)] = zoneName; });
  context.exportLegacyTripHistoryPdf();
  assert(pdf.internal.pageSize.getWidth() < pdf.internal.pageSize.getHeight(), "Export must be portrait");
  assert(Math.abs(pdf.internal.pageSize.getWidth() - 210) < 0.01, "A4 portrait width");
  assert.deepEqual(rectangles.slice(0, 2), [
    { x: 0, y: 0, w: pdf.internal.pageSize.getWidth(), h: 24, style: "F" },
    { x: 14, y: 32, w: pdf.internal.pageSize.getWidth() - 28, h: 8, style: "F" }
  ], "Top navy header and black totals bar keep their original positions");
  assert.equal(logos.length, 1, "Real embedded Geotab logo must be drawn");
  assert.equal(logos[0].format, "PNG");
  assert.equal(logos[0].x, pdf.internal.pageSize.getWidth() - 14 - logos[0].w);
  assert.equal(logos[0].y, (24 - logos[0].h) / 2);
  assert.equal(kpiBoxes.length, 6, "The existing six KPI boxes stay in the PDF");
  assert.equal(kpiBoxes[0].x, 14); assert.equal(kpiBoxes[0].y, 46);
  assert.equal(kpiBoxes[3].x, 14); assert.equal(kpiBoxes[3].y, 66);
  for (const [prefix, x, y] of [["Legacy Trip History Report", 14, 11], ["Created:", 14, 18], ["Report Totals for:", 17, 37.5]]) {
    assert(headings.some(heading => heading.page === 1 && heading.lines[0].startsWith(prefix) && heading.x === x && heading.y === y), prefix + " remains in the original header position");
  }
  for (let page = 1; page <= pdf.internal.getNumberOfPages(); page++) {
    assert(headings.some(heading => heading.page === page && heading.lines[0] === "Confidential" && heading.x === 14 && heading.y === pdf.internal.pageSize.getHeight() - 8), "Left footer unchanged on page " + page);
    assert(headings.some(heading => heading.page === page && /^Page \d+ of \d+$/.test(heading.lines[0]) && heading.x === pdf.internal.pageSize.getWidth() - 14 && heading.y === pdf.internal.pageSize.getHeight() - 8), "Right footer unchanged on page " + page);
  }
  assert(tableLayouts.length > 0);
  assert(tableLayouts.every(layout => layout.margin.left === 5 && layout.margin.right === 5 && layout.margin.top === undefined && layout.margin.bottom === undefined && Math.abs(layout.tableWidth - 200) < 0.01), "Only the table side margins change to 5 mm");
  if (zones) assert(tableLayouts.every(layout => layout.head[1][4] === "Zone"), "PDFs use the singular column heading");
  else assert(tableLayouts.every(layout => !layout.head[1].includes("Zone")), "Zones-off PDF has no Zone column");
  const assetHeading = headings.find(heading => heading.lines.join(" ") === assetName);
  assert(assetHeading && assetHeading.lines.length > 1, "Long asset headings wrap on portrait pages");
  pdf.setFont("helvetica", "bold"); pdf.setFontSize(12);
  assert(assetHeading.lines.every(line => pdf.getTextWidth(line) <= 182.01), "Asset headings keep their original 14 mm inset");
  assert(pdf.internal.getNumberOfPages() > 1, "Exercise actual AutoTable pagination");
  const addressCells = cells.filter(cell => cell.section === "body" && cell.column === 3 && typeof cell.raw === "string" && addresses.includes(cell.raw));
  const short = addressCells.find(cell => cell.raw === addresses[0]);
  const long = addressCells.find(cell => cell.raw === addresses[1]);
  const unbroken = addressCells.filter(cell => cell.raw === addresses[2]);
  assert(long.text.length > 1 && long.height > short.height, "Long addresses wrap and grow row height");
  assert(unbroken.length > 1, "Oversize address continues across portrait pages");
  assert.equal(unbroken.flatMap(cell => cell.text).join(""), addresses[2], "Never truncate even unbroken address text");
  assert(addressCells.every(cell => Math.abs(cell.width - (zones ? 38 : 41)) < 0.01), "Wider address column remains capped");
  const zoneCells = cells.filter(cell => cell.section === "body" && cell.column === 4 && cell.raw === zoneName);
  if (zones) {
    assert(zoneCells.length >= trips.length && zoneCells.every(cell => Math.abs(cell.width - 28) < 0.01), "Zones fit their allocated portrait column");
    assert(zoneCells.some(cell => cell.text.length > 1), "Long zone names wrap");
    assert(cells.some(cell => cell.section === "head" && cell.raw === "Zone"));
  } else assert.equal(zoneCells.length, 0, "No empty zone column when off");
  assert(cells.every(cell => cell.x >= 4.99 && cell.x + cell.width <= 205.02), "All table cells fit the 5 mm side margins");
  assert(cells.some(cell => Math.abs(cell.x - 5) < 0.01) && cells.some(cell => Math.abs(cell.x + cell.width - 205) < 0.02), "The table actually uses the newly freed space on both sides");
  assert(cells.every(cell => cell.y >= 0 && cell.y + cell.height <= pdf.internal.pageSize.getHeight()), "Wrapped cells stay within page height");
  assert.equal(links.length, trips.length, "Each trip retains one complete PDF URI annotation");
  assert(links.every(url => url.startsWith("https://my.geotab.com/synthetic-demo/#tripsHistory,")));
  assert(cells.some(cell => cell.raw && cell.raw.content === "Vehicle Total"), "Vehicle-period footer survives pagination");
  assert.equal(filename, "legacy_trip_history_12-10-2026_to_12-10-2026.pdf");
  assert(bytes.length > 10000 && bytes.subarray(0, 4).toString() === "%PDF");
  assert(bytes.includes(Buffer.from("/Subtype /Image")), "Logo must actually be embedded in the PDF");
  console.log(file + " zones=" + zones + ": PASS — portrait, wrapped asset/zone/address rows, oversize pagination, 40 links, totals (" + pdf.internal.getNumberOfPages() + " pages)");
 }
}
