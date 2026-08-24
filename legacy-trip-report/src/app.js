"use strict";

// ─── Legacy Trip History — standalone MyGeotab add-in ─────────────────────────
// A single-purpose extraction of the Legacy Trip History Report from the Smart
// Insights add-in. Recreates a classic per-vehicle/per-day trip log (start time,
// distance, driving duration, reverse-geocoded stop address, arrival time, idle
// duration, stop duration, daily totals) with CSV + branded PDF export.
//
// Addresses are resolved via MyGeotab's own GetAddresses method (Trip.stopPoint →
// coordinates → GetAddresses → ReverseGeocodeAddress): coordinates never leave
// Geotab's infrastructure, and no third-party geocoder is used.
//
// KEEP IN SYNC: this file is a lift of the legacy-report functions in the parent
// Smart Insights src/app.js. If you fix a bug here, mirror it there (and vice
// versa) until the two are formally split into separate repos.

// ─── State ─────────────────────────────────────────────────────────────────
var state = {
  api:            null,
  reportRaw:      [],
  reportData:     [],
  deviceMap:      {},
  dbName:         "",
  server:         "",    // MyGeotab host for Trip History links, e.g. "my.geotab.com"
  hostVia:        "",    // where server/dbName came from, for the diagnostics line
  legacyByDevice: null,  // { deviceId: { name, trips, drivers } }
  legacyAddrMap:  null,  // "lat,lng" -> resolved address string
  legacyStopInfo: null   // Map(trip -> { mins, excluded }) — see buildLegacyStopInfo
};

// ─── MyGeotab addin lifecycle ──────────────────────────────────────────────
if (typeof geotab === "undefined") { var geotab = { addin: {} }; }

geotab.addin.legacyTripHistory = function () {
  return {
    initialize: function (api, freshState, callback) {
      try {
        state.api = api;
        if (freshState && freshState.database) {
          state.dbName = freshState.database;
          state.hostVia = "session";
          document.getElementById("db-name").textContent = freshState.database;
        }
        resolveHost();
        setupReports();
        document.getElementById("loading").classList.add("hidden");
        document.getElementById("main").classList.remove("hidden");
      } catch (err) {
        showError("Init error: " + err.message);
      }
      if (callback) callback();
    },
    focus: function () {},
    blur:  function () {}
  };
};

// ─── Setup / controls ──────────────────────────────────────────────────────
function setupReports() {
  var today = new Date();
  document.getElementById("filter-to").value = fmtDateInput(today);
  applyLegacyLength(); // fills filter-from from the default length preset
  document.getElementById("run-report").addEventListener("click", runReport);
  document.getElementById("export-csv").addEventListener("click", exportReportCsv);
  document.getElementById("export-pdf-legacy").addEventListener("click", exportLegacyTripHistoryPdf);
  document.getElementById("legacy-length").addEventListener("change", applyLegacyLength);
}

// Recompute filter-from from filter-to based on the selected preset length.
// "custom" leaves both date inputs alone so the user can pick any range manually.
function applyLegacyLength() {
  var days = document.getElementById("legacy-length").value;
  if (days === "custom") return;
  var toEl = document.getElementById("filter-to");
  var to   = toEl.value ? new Date(toEl.value + "T00:00:00") : new Date();
  var from = new Date(to);
  from.setDate(to.getDate() - (parseInt(days, 10) - 1));
  document.getElementById("filter-from").value = fmtDateInput(from);
  document.getElementById("filter-to").value   = fmtDateInput(to);
}

function runReport() {
  var from   = document.getElementById("filter-from").value;
  var to     = document.getElementById("filter-to").value;
  var btn    = document.getElementById("run-report");
  var output = document.getElementById("report-output");
  if (!from || !to) { alert("Please select a date range."); return; }

  btn.disabled = true; btn.textContent = "Loading...";
  output.innerHTML = "<p class='placeholder'>Fetching data...</p>";
  document.getElementById("report-summary").classList.add("hidden");
  document.getElementById("export-csv").classList.add("hidden");
  document.getElementById("export-pdf-legacy").classList.add("hidden");
  document.getElementById("report-vehicle").classList.add("hidden");

  function done() { btn.disabled = false; btn.textContent = "Run"; }
  runLegacyTripHistoryReport(from, to, done);
}

// ─── Duration parsing / formatting ─────────────────────────────────────────
// Parse a MyGeotab TimeSpan value to minutes. MyGeotab serializes TimeSpan as a
// .NET-style string: "[-][d.]hh:mm:ss[.fffffff]" (e.g. "00:44:00", "1.02:28:55",
// "17:08:00"). Also tolerates ISO-8601 ("PT4M22S") and plain numeric seconds as
// fallbacks. Returns null if unparseable so callers can fall back to timestamp math.
function parseDurationToMins(val) {
  if (val == null) return null;
  if (typeof val === "number") return val / 60; // assume seconds
  if (typeof val !== "string") return null;
  var s = val.trim();

  // .NET TimeSpan: [-][d.]hh:mm:ss[.fffffff]  (days optional, fraction optional)
  var ts = s.match(/^(-)?(?:(\d+)\.)?(\d{1,3}):([0-5]?\d):([0-5]?\d)(?:\.(\d+))?$/);
  if (ts) {
    var sign = ts[1] ? -1 : 1;
    var days = parseFloat(ts[2] || 0);
    var hh   = parseFloat(ts[3] || 0);
    var mm   = parseFloat(ts[4] || 0);
    var ss   = parseFloat(ts[5] || 0);
    var frac = ts[6] ? parseFloat("0." + ts[6]) : 0;
    return sign * (days * 1440 + hh * 60 + mm + (ss + frac) / 60);
  }

  // ISO-8601 duration fallback: PnDTnHnMnS
  var iso = s.match(/^P(?:(\d+)D)?T?(?:(\d+)H)?(?:(\d+)M)?(?:([\d.]+)S)?$/);
  if (iso && (iso[1] || iso[2] || iso[3] || iso[4])) {
    return parseFloat(iso[1] || 0) * 1440 + parseFloat(iso[2] || 0) * 60 + parseFloat(iso[3] || 0) + parseFloat(iso[4] || 0) / 60;
  }
  return null;
}
// Whole-minute duration for summary cards: "4h 22m" / "44m"
function fmtDurWhole(mins) {
  if (mins == null || isNaN(mins)) return "\u2014";
  var total = Math.round(mins);
  return total < 60 ? total + "m" : Math.floor(total / 60) + "h " + (total % 60) + "m";
}
// Second-precision duration for table rows: "02m 02s" / "1h 42m 28s"
function fmtDurPrecise(mins) {
  if (mins == null || isNaN(mins)) return "\u2014";
  var totalSecs = Math.max(0, Math.round(mins * 60));
  var h = Math.floor(totalSecs / 3600);
  var m = Math.floor((totalSecs % 3600) / 60);
  var s = totalSecs % 60;
  var mm = (m < 10 ? "0" : "") + m;
  var ss = (s < 10 ? "0" : "") + s;
  return h > 0 ? h + "h " + mm + "m " + ss + "s" : mm + "m " + ss + "s";
}
// Trip.Distance is in KILOMETRES per the MyGeotab SDK (confirmed against live
// data: a multi-hour drive showed ~0.01 mi when distance was wrongly treated as
// metres). 1 km = 0.621371 miles.
function milesFromDistance(km) { return (km || 0) * 0.621371; }

// ─── Overnight parking ─────────────────────────────────────────────────────
// Trip.StopDuration runs from the engine stopping to the NEXT trip's start, so a
// vehicle parked at 18:00 and driven again at 07:00 arrives as a single 13-hour
// stop on the day's last trip. Nothing on the Trip record separates that from a
// real stop, so we infer it. A stop is treated as an overnight park, and left
// out of every figure, when all three hold:
//   1. it is that vehicle's last trip of the local day,
//   2. it runs longer than OVERNIGHT_MIN_HOURS,
//   3. it ends on a later local day than it began.
// A long mid-shift stop that ends the same day is genuine and still counts.
var OVERNIGHT_MIN_HOURS = 3;

// Local calendar day. Everything the report displays is local time, so the day
// bucket has to be local too — toISOString() would file a 00:30 BST trip under
// the previous day, and an evening trip west of Greenwich under the next one.
function localDayKey(iso) {
  if (!iso) return "";
  var d = new Date(iso);
  if (isNaN(d.getTime())) return "";
  var m = d.getMonth() + 1, day = d.getDate();
  return d.getFullYear() + "-" + (m < 10 ? "0" : "") + m + "-" + (day < 10 ? "0" : "") + day;
}

// One vehicle's trips, already sorted by start -> Map(trip -> { mins, excluded }).
// mins is null for an excluded stop, which is exactly what the renderers already
// treat as "unknown" and draw as an em-dash.
function buildLegacyStopInfo(sortedTrips) {
  var lastOfDay = {};
  sortedTrips.forEach(function (t, i) { lastOfDay[localDayKey(t.start)] = i; });

  var info = new Map();
  sortedTrips.forEach(function (t, i) {
    var mins = parseDurationToMins(t.stopDuration);
    var excluded;
    if (mins == null) {
      // Usually the last trip in the range: Geotab has no next trip to measure
      // to. We cannot count what we cannot measure.
      excluded = true;
    } else if (i === lastOfDay[localDayKey(t.start)] && mins > OVERNIGHT_MIN_HOURS * 60) {
      var endsAt = new Date(new Date(t.stop).getTime() + mins * 60000);
      excluded = localDayKey(endsAt.toISOString()) > localDayKey(t.stop);
    } else {
      excluded = false;
    }
    info.set(t, { mins: excluded ? null : mins, excluded: excluded });
  });
  return info;
}

// Stop minutes for the legacy report, or null when the stop is an overnight park
// or unmeasurable. Falls back to the raw field if the trip predates the map.
function legacyStopMins(t) {
  var entry = state.legacyStopInfo && state.legacyStopInfo.get(t);
  return entry ? entry.mins : parseDurationToMins(t.stopDuration);
}

// ─── Data assembly ─────────────────────────────────────────────────────────
function buildLegacyByDevice(trips, deviceMap) {
  var byDevice = {};
  trips.forEach(function (t) {
    var vid = t.device && t.device.id; if (!vid) return;
    if (!byDevice[vid]) byDevice[vid] = { name: deviceMap[vid] || vid, trips: [], drivers: {} };
    byDevice[vid].trips.push(t);
    if (t.driverName) byDevice[vid].drivers[t.driverName] = 1;
  });
  // Overnight detection needs each vehicle's trips in order, so it rides along
  // with the sort rather than being recomputed in each of the three renderers.
  var stopInfo = new Map();
  Object.keys(byDevice).forEach(function (vid) {
    byDevice[vid].trips.sort(function (a, b) { return new Date(a.start) - new Date(b.start); });
    buildLegacyStopInfo(byDevice[vid].trips).forEach(function (v, k) { stopInfo.set(k, v); });
  });
  state.legacyStopInfo = stopInfo;
  return byDevice;
}

// Round a coordinate to ~1m precision so repeat stops (depots, regular customer
// sites) share one GetAddresses lookup instead of a fresh one each time.
function coordKey(pt) { return pt.y.toFixed(5) + "," + pt.x.toFixed(5); }

// Collect every unique StopPoint, resolve them in as few GetAddresses calls as
// possible (450/min limit → chunk generously), return "lat,lng" -> address map.
function resolveStopAddresses(byDevice, cb) {
  var uniq = {};
  Object.keys(byDevice).forEach(function (vid) {
    byDevice[vid].trips.forEach(function (t) {
      var sp = t.stopPoint;
      if (sp && sp.y != null && sp.x != null) uniq[coordKey(sp)] = { x: sp.x, y: sp.y };
    });
  });
  var keys = Object.keys(uniq);
  if (!keys.length) { cb({}); return; }

  var CHUNK = 400;
  var addrMap = {};
  var idx = 0;

  function runChunk() {
    if (idx >= keys.length) { cb(addrMap); return; }
    var sliceKeys = keys.slice(idx, idx + CHUNK);
    var coords = sliceKeys.map(function (k) { return { x: uniq[k].x, y: uniq[k].y }; });
    apiCall("GetAddresses", { coordinates: coords }, function (results) {
      (results || []).forEach(function (addr, i) { addrMap[sliceKeys[i]] = formatReverseGeocodeAddress(addr); });
      idx += CHUNK;
      runChunk();
    }, function () {
      idx += CHUNK; // leave failed keys unresolved (renderer falls back to coords)
      runChunk();
    });
  }
  runChunk();
}

function formatReverseGeocodeAddress(addr) {
  if (!addr) return null;
  if (addr.formattedAddress) return addr.formattedAddress;
  var parts = [];
  if (addr.streetNumber) parts.push(addr.streetNumber);
  if (addr.street || addr.streetName) parts.push(addr.street || addr.streetName);
  if (addr.city) parts.push(addr.city);
  if (addr.postalCode) parts.push(addr.postalCode);
  return parts.length ? parts.join(" ") : null;
}

function addressForPoint(pt, addrMap) {
  if (!pt || pt.y == null || pt.x == null) return "(location unknown)";
  var resolved = addrMap[coordKey(pt)];
  return resolved || (pt.y.toFixed(5) + ", " + pt.x.toFixed(5));
}

// Split one vehicle's chronological trip list into day blocks, each row carrying
// the previous trip for the synthetic "Starting from" / "(Ignition On)" lead-in.
function buildLegacyDayBlocks(vTrips, addrMap) {
  var days = {};
  var order = [];
  vTrips.forEach(function (t, i) {
    var day = localDayKey(t.start);
    if (!days[day]) { days[day] = { date: day, rows: [] }; order.push(day); }
    days[day].rows.push({ trip: t, prevTrip: i > 0 ? vTrips[i - 1] : null });
  });
  return order.map(function (day) { return days[day]; });
}

// ─── Trip History deep link ────────────────────────────────────────────────
// Ported from the Activity Report add-in (~/repos/vehicle-activity-log). Keep the
// two in sync: the URL shape below is not documented anywhere reliable.
//
// The format in the SDK guide ("Using MyGeotab URLs") is stale — its examples date
// from 2015 and Geotab simplified the URLs in 2022. entityType and selectedEntities
// are no longer read by the Trips History page, so links built from the docs open a
// blank page. This shape was copied out of the address bar after selecting one
// vehicle and a custom date range by hand in a live database:
//
//   #tripsHistory,
//   dateRange:(endDate:'…',label:Custom,startDate:'…'),
//   devices:!(bC),
//   expandedCardIds:!('bC_UnknownDriverId_Tue+Aug+18'),
//   isReplayPlayerHidden:!f,
//   routes:(bC:!((start:'…',stop:'…')))
//
// What each part does:
//   devices               THE vehicle filter. Not entityType/selectedEntities.
//   routes                device id -> trip segments to draw. Does not select the
//                         vehicle on its own; devices does that. One segment is
//                         deliberate: land on this row's trip, not redraw the day.
//   isReplayPlayerHidden  rison !f is false, so the replay player opens.
//   dateRange             scopes the trip list. label:Custom needs explicit dates.
//   expandedCardIds       opens the matching trip card in the side list.
//                         "<deviceId>_<driverId>_<Ddd+Mmm+D>", spaces as "+".
//   mapBounds             viewport only; omitted so the map fits the route.
//
// Timezone note: the Activity Report builds these from the MyGeotab profile
// timezone. This report has no profile-timezone plumbing — every time it prints
// (fmtTime, localDayKey) is browser-local — so the link uses browser-local too and
// stays consistent with the row it sits on.
var DOW_SHORT = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
var MON_SHORT = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

function cardDatePart(iso) {
  var d = new Date(iso);
  if (isNaN(d.getTime())) return "";
  return DOW_SHORT[d.getDay()] + "+" + MON_SHORT[d.getMonth()] + "+" + d.getDate();
}

// Where the MyGeotab host comes from, best source first. Add-in pages are injected
// into the MyGeotab page rather than iframed, so window.location is normally
// already the real MyGeotab URL: https://my.geotab.com/<database>/#legacyTripHistory
function resolveHost() {
  if (!state.server) {
    var host = window.location.hostname || "";
    // Ignore the case where the page really is served from its own host.
    if (host && !/(^|\.)github\.io$/i.test(host) && host !== "localhost" && host !== "127.0.0.1") {
      state.server  = host;
      state.hostVia = "window.location";
      var seg = window.location.pathname.split("/").filter(Boolean);
      if (!state.dbName && seg.length) state.dbName = decodeURIComponent(seg[0]);
    }
  }
  return !!(state.server && state.dbName);
}

// Returns "" when no host resolved, which the renderer turns into a disabled cell
// rather than a link that quietly lands on an empty Trips History page.
function tripHistoryUrl(deviceId, trip) {
  if (!state.server || !state.dbName) return "";
  var d = new Date(trip.start);
  if (isNaN(d.getTime())) return "";

  // Midnight to 23:59:59 of that local day, as UTC instants, which is what
  // MyGeotab itself puts in dateRange.
  var dayStart = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 0, 0, 0);
  var dayEnd   = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 23, 59, 59);

  // Trip.driver is a Driver entity, or the "UnknownDriverId" sentinel when nobody
  // was assigned. The card id needs whichever.
  var driverId = (trip.driver && trip.driver.id) ? trip.driver.id : "UnknownDriverId";
  var cardId   = deviceId + "_" + driverId + "_" + cardDatePart(trip.start);

  // Keys emitted alphabetically, matching how MyGeotab serialises them.
  var parts = [
    "dateRange:(endDate:'" + dayEnd.toISOString() + "',label:Custom,startDate:'" + dayStart.toISOString() + "')",
    "devices:!(" + deviceId + ")"
  ];
  if (cardId) parts.push("expandedCardIds:!('" + cardId + "')");
  parts.push("isReplayPlayerHidden:!f");
  parts.push("routes:(" + deviceId + ":!((start:'" + new Date(trip.start).toISOString() +
             "',stop:'" + new Date(trip.stop).toISOString() + "')))");

  return "https://" + state.server + "/" + state.dbName + "/#tripsHistory," + parts.join(",");
}

// A real href, not "#": the browser status bar then shows where the link goes, and
// middle-click / copy-link-address both work, which is what makes this diagnosable
// without a debugger.
function tripHistoryCell(deviceId, trip) {
  var url = tripHistoryUrl(deviceId, trip);
  if (!url) {
    return "<td><span class='trip-history-link is-disabled' title='No MyGeotab host resolved, so no Trip History link could be built.'>Trip History</span></td>";
  }
  return "<td><a href='" + esc(url) + "' target='_blank' rel='noopener' class='trip-history-link'" +
         " title='" + esc(url) + "'>Trip History</a></td>";
}

// ─── Run ───────────────────────────────────────────────────────────────────
function runLegacyTripHistoryReport(from, to, done) {
  var fromDate = new Date(from + "T00:00:00").toISOString();
  var toDate   = new Date(to   + "T23:59:59").toISOString();
  apiCall("Get", { typeName: "Device", search: {} }, function (devices) {
    var deviceMap = {};
    devices.forEach(function (d) { deviceMap[d.id] = d.name; });
    state.deviceMap = deviceMap;
    apiCall("Get", { typeName: "Trip", search: { fromDate: fromDate, toDate: toDate } }, function (trips) {
      if (!trips.length) {
        document.getElementById("report-output").innerHTML = "<p class='placeholder'>No trips found for this period.</p>";
        done(); return;
      }
      var byDevice = buildLegacyByDevice(trips, deviceMap);
      state.reportRaw  = trips;
      state.reportData = trips;
      state.legacyByDevice = byDevice;
      populateReportVehicleFilter(trips, deviceMap);

      document.getElementById("report-output").innerHTML = "<p class='placeholder'>Resolving stop addresses\u2026</p>";
      resolveStopAddresses(byDevice, function (addrMap) {
        state.legacyAddrMap = addrMap;

        var totalDistM = trips.reduce(function (s, t) { return s + (t.distance || 0); }, 0);
        var totalDrive = trips.reduce(function (s, t) { var dm = parseDurationToMins(t.drivingDuration); return s + (dm != null ? dm : durationMins(t.start, t.stop)); }, 0);
        var totalIdle  = trips.reduce(function (s, t) { return s + (parseDurationToMins(t.idlingDuration) || 0); }, 0);
        // Overnight parks carry no stop figure, so they are not stops here either.
        // Counting them in the denominator while dropping them from the total would
        // make the average disagree with the two cards it sits between.
        var counted    = trips.filter(function (t) { return legacyStopMins(t) != null; });
        var totalStop  = counted.reduce(function (s, t) { return s + legacyStopMins(t); }, 0);
        var numStops   = counted.length;
        var avgStop    = numStops ? totalStop / numStops : 0;

        showReportSummary([
          summaryCard("Total Distance",         milesFromDistance(totalDistM).toFixed(0), "mi"),
          summaryCard("Total Stop Duration",    fmtDurWhole(totalStop), ""),
          summaryCard("Total Idle Time",        fmtDurWhole(totalIdle), ""),
          summaryCard("Total Travel Duration",  fmtDurWhole(totalDrive), ""),
          summaryCard("Average Stop Duration",  fmtDurWhole(avgStop), ""),
          summaryCard("Number of Stops",        numStops, "")
        ]);

        renderLegacyTripHistoryOutput(byDevice, addrMap);
        document.getElementById("export-csv").classList.remove("hidden");
        document.getElementById("export-pdf-legacy").classList.remove("hidden");
        done();
      });
    }, function (err) { reportError(err); done(); });
  }, function (err) { reportError(err); done(); });
}

// ─── Render (in-app preview) ───────────────────────────────────────────────
function renderLegacyTripHistoryOutput(byDevice, addrMap) {
  var output = document.getElementById("report-output");
  var vids = Object.keys(byDevice);
  if (!vids.length) { output.innerHTML = "<p class='placeholder'>No trips found.</p>"; return; }

  output.innerHTML = vids.map(function (vid) {
    var v = byDevice[vid];
    var driverNames = Object.keys(v.drivers).join(", ") || "No driver assigned";
    var dayBlocks = buildLegacyDayBlocks(v.trips, addrMap);

    var dayHtml = dayBlocks.map(function (block) {
      var rowsHtml = block.rows.map(function (r) {
        var t = r.trip;
        var driveMins = parseDurationToMins(t.drivingDuration); if (driveMins == null) driveMins = durationMins(t.start, t.stop);
        var idleMins  = parseDurationToMins(t.idlingDuration) || 0;
        var stopMins  = legacyStopMins(t);
        var stopStr   = stopMins != null ? fmtDurPrecise(stopMins) : "\u2014";
        return "<tr>" +
          "<td>" + fmtTime(t.start) + "</td>" +
          "<td>" + milesFromDistance(t.distance).toFixed(2) + " mi<br><span style='color:var(--text-muted);font-size:11px'>" + fmtDurPrecise(driveMins) + "</span></td>" +
          "<td>" + esc(addressForPoint(t.stopPoint, addrMap)) + "</td>" +
          "<td>" + fmtTime(t.stop) + "</td>" +
          "<td>" + fmtDurPrecise(idleMins) + "</td>" +
          "<td>" + stopStr + "</td>" +
          tripHistoryCell(vid, t) +
        "</tr>";
      }).join("");

      var dayDistM = block.rows.reduce(function (s, r) { return s + (r.trip.distance || 0); }, 0);
      var dayDrive = block.rows.reduce(function (s, r) { var dm = parseDurationToMins(r.trip.drivingDuration); return s + (dm != null ? dm : durationMins(r.trip.start, r.trip.stop)); }, 0);
      var dayIdle  = block.rows.reduce(function (s, r) { return s + (parseDurationToMins(r.trip.idlingDuration) || 0); }, 0);
      var dayCount = block.rows.filter(function (r) { return legacyStopMins(r.trip) != null; });
      var dayStop  = dayCount.reduce(function (s, r) { return s + legacyStopMins(r.trip); }, 0);

      var firstTrip = block.rows[0].trip;
      var prevTrip  = block.rows[0].prevTrip;
      var startingFromHtml = "<tr><td colspan='7' style='font-weight:600;background:var(--surface-2)'>" + block.date + " \u2014 Starting from: " + esc(addressForPoint(prevTrip ? prevTrip.stopPoint : null, addrMap)) + "</td></tr>";
      // "(Ignition On)" marks the day's first drive. We deliberately do NOT show a
      // duration here: the gap to the previous trip is overnight/parked time, not
      // idle, and Trip data has no pre-drive idle figure to show instead.
      var ignitionHtml = prevTrip
        // No Trip History cell here: this synthetic row marks the same trip whose
        // own row, directly below, already carries the link.
        ? "<tr style='color:var(--accent)'><td>" + fmtTime(firstTrip.start) + "</td><td colspan='3'>(Ignition On)</td><td>\u2014</td><td></td><td></td></tr>"
        : "";

      return "<div class='dd-table-wrap' style='margin-bottom:12px'><table class='dd-table'>" +
        "<thead><tr><th>Start Time</th><th>Distance / Duration</th><th>Stop Location</th><th>Arrival Time</th><th>Idle Duration</th><th>Stop Duration</th><th>Trip History</th></tr></thead>" +
        "<tbody>" + startingFromHtml + ignitionHtml + rowsHtml +
        "<tr style='font-weight:700;background:var(--surface-2)'><td>" + block.date + " Total</td><td>" + milesFromDistance(dayDistM).toFixed(2) + " mi in " + fmtDurWhole(dayDrive) + "</td><td></td><td>" + dayCount.length + " stops</td><td>" + fmtDurWhole(dayIdle) + "</td><td>" + fmtDurWhole(dayStop) + "</td><td></td></tr>" +
        "</tbody></table></div>";
    }).join("");

    return "<div style='margin-bottom:24px'><h3 style='margin-bottom:4px'>" + esc(v.name) + "</h3>" +
      "<div style='color:var(--text-muted);font-size:12px;margin-bottom:8px'>Driver(s): " + esc(driverNames) + "</div>" +
      dayHtml + "</div>";
  }).join("");
}

// ─── Vehicle filter ────────────────────────────────────────────────────────
function populateReportVehicleFilter(data, deviceMap) {
  var dMap = deviceMap || state.deviceMap;
  var ids = {};
  data.forEach(function (item) {
    var vid = item.device && item.device.id;
    if (vid) ids[vid] = dMap[vid] || vid;
  });
  var sel = document.getElementById("report-vehicle");
  sel.innerHTML = "<option value=''>All Vehicles</option>";
  Object.keys(ids).sort(function (a, b) { return ids[a].localeCompare(ids[b]); }).forEach(function (id) {
    var opt = document.createElement("option");
    opt.value = id; opt.textContent = ids[id];
    sel.appendChild(opt);
  });
  sel.classList.remove("hidden");
  if (!sel._filterBound) {
    sel._filterBound = true;
    sel.addEventListener("change", function () { filterReportByVehicle(this.value); });
  }
}

function filterReportByVehicle(vid) {
  var raw = state.reportRaw;
  state.reportData = !vid ? raw : raw.filter(function (t) { return t.device && t.device.id === vid; });
  var filteredByDevice = buildLegacyByDevice(state.reportData, state.deviceMap);
  state.legacyByDevice = filteredByDevice;
  renderLegacyTripHistoryOutput(filteredByDevice, state.legacyAddrMap || {});
}

// ─── CSV export ────────────────────────────────────────────────────────────
function exportReportCsv() {
  var data = state.reportData;
  if (!data || !data.length) return;
  var addrMap = state.legacyAddrMap || {};
  var rows = [["Vehicle", "Driver(s)", "Date", "Start Time", "Distance (mi)", "Driving Duration", "Stop Location", "Arrival Time", "Idle Duration", "Stop Duration"]];
  var byDevice = buildLegacyByDevice(data, state.deviceMap);
  Object.keys(byDevice).forEach(function (vid) {
    var v = byDevice[vid];
    var driverNames = Object.keys(v.drivers).join(", ") || "No driver assigned";
    v.trips.forEach(function (t) {
      var driveMins = parseDurationToMins(t.drivingDuration); if (driveMins == null) driveMins = durationMins(t.start, t.stop);
      var idleMins  = parseDurationToMins(t.idlingDuration) || 0;
      var stopMins  = legacyStopMins(t);
      rows.push([
        v.name, driverNames, localDayKey(t.start), fmtTime(t.start),
        milesFromDistance(t.distance).toFixed(2), fmtDurPrecise(driveMins),
        addressForPoint(t.stopPoint, addrMap), fmtTime(t.stop),
        fmtDurPrecise(idleMins), stopMins != null ? fmtDurPrecise(stopMins) : ""
      ]);
    });
  });
  if (rows.length > 1) downloadCsvBlob(rows, "legacy_trip_history_report.csv");
}

// ─── PDF branding helper ───────────────────────────────────────────────────
// Draws the Geotab logo in the top-right of a PDF header bar. The logo PNG has a
// navy (#27325D / rgb 39,50,93) background, so header bars that host it are drawn
// in the same navy for a seamless fit. Logo data + native dimensions come from
// src/logo-data.js (loaded before app.js). No-ops safely if that file is missing.
var GEOTAB_NAVY = [39, 50, 93];
function drawPdfHeaderLogo(doc, pageW, margin, barH) {
  if (typeof GEOTAB_LOGO_DATAURL === "undefined") return;
  var logoH = Math.min(9, barH - 6);
  var logoW = logoH * (GEOTAB_LOGO_W / GEOTAB_LOGO_H);
  var x = pageW - margin - logoW;
  var y = (barH - logoH) / 2;
  try { doc.addImage(GEOTAB_LOGO_DATAURL, "PNG", x, y, logoW, logoH); } catch (e) {}
}

// ─── PDF export ────────────────────────────────────────────────────────────
function exportLegacyTripHistoryPdf() {
  var byDevice = state.legacyByDevice;
  var addrMap  = state.legacyAddrMap || {};
  if (!byDevice || !Object.keys(byDevice).length) { alert("Run the report first."); return; }
  var from = document.getElementById("filter-from").value;
  var to   = document.getElementById("filter-to").value;

  var doc    = new jspdf.jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
  var pageW  = doc.internal.pageSize.getWidth();
  var pageH  = doc.internal.pageSize.getHeight();
  var margin = 14;

  doc.setFillColor(GEOTAB_NAVY[0], GEOTAB_NAVY[1], GEOTAB_NAVY[2]);
  doc.rect(0, 0, pageW, 24, "F");
  drawPdfHeaderLogo(doc, pageW, margin, 24);
  doc.setTextColor(255, 255, 255); doc.setFontSize(16); doc.setFont("helvetica", "bold");
  doc.text("Legacy Trip History Report", margin, 11);
  doc.setFontSize(9); doc.setFont("helvetica", "normal");
  var now = new Date();
  doc.text("Created: " + fmtDateReadable(now.toISOString()) + " " + fmtTime(now.toISOString()), margin, 18);

  var yPos = 32;
  doc.setFillColor(20, 20, 20);
  doc.rect(margin, yPos, pageW - margin * 2, 8, "F");
  doc.setTextColor(255, 255, 255); doc.setFontSize(10); doc.setFont("helvetica", "bold");
  doc.text("Report Totals for: " + from + " \u2013 " + to, margin + 3, yPos + 5.5);
  yPos += 14;

  var allTrips = [];
  Object.keys(byDevice).forEach(function (vid) { allTrips = allTrips.concat(byDevice[vid].trips); });
  var totalDistM = allTrips.reduce(function (s, t) { return s + (t.distance || 0); }, 0);
  var totalDrive = allTrips.reduce(function (s, t) { var dm = parseDurationToMins(t.drivingDuration); return s + (dm != null ? dm : durationMins(t.start, t.stop)); }, 0);
  var totalIdle  = allTrips.reduce(function (s, t) { return s + (parseDurationToMins(t.idlingDuration) || 0); }, 0);
  // Same counted set as the on-screen KPIs — overnight parks are not stops.
  var counted    = allTrips.filter(function (t) { return legacyStopMins(t) != null; });
  var totalStop  = counted.reduce(function (s, t) { return s + legacyStopMins(t); }, 0);
  var numStops   = counted.length;
  var avgStop    = numStops ? totalStop / numStops : 0;

  function drawKpiRow(items, y) {
    var boxW = (pageW - margin * 2 - 8) / 3;
    items.forEach(function (b, i) {
      var bx = margin + i * (boxW + 4);
      doc.setFillColor(232, 244, 253); doc.roundedRect(bx, y, boxW, 16, 2, 2, "F");
      doc.setTextColor(0, 90, 158); doc.setFontSize(13); doc.setFont("helvetica", "bold");
      doc.text(b.value, bx + 4, y + 8);
      doc.setFontSize(7); doc.setFont("helvetica", "normal"); doc.setTextColor(100, 100, 100);
      doc.text(b.label, bx + 4, y + 13, { maxWidth: boxW - 8 });
    });
  }
  drawKpiRow([
    { label: "Total Stop Duration",   value: fmtDurWhole(totalStop) },
    { label: "Total Idle Time",       value: fmtDurWhole(totalIdle) },
    { label: "Total Travel Duration", value: fmtDurWhole(totalDrive) }
  ], yPos);
  yPos += 20;
  drawKpiRow([
    { label: "Total Distance Travelled (miles)", value: milesFromDistance(totalDistM).toFixed(0) },
    { label: "Average Stop Duration",            value: fmtDurWhole(avgStop) },
    { label: "Number of Stops",                  value: String(numStops) }
  ], yPos);
  yPos += 24;

  var headers = [["Start Time", "Distance / Duration", "Stop Location", "Arrival Time", "Idle Duration", "Stop Duration", "Stops"]];

  Object.keys(byDevice).forEach(function (vid) {
    var v = byDevice[vid];
    if (yPos > pageH - 40) { doc.addPage(); yPos = 16; }
    doc.setTextColor(20, 20, 20); doc.setFontSize(12); doc.setFont("helvetica", "bold");
    doc.text(v.name, margin, yPos);
    doc.setFontSize(9); doc.setFont("helvetica", "normal"); doc.setTextColor(100, 100, 100);
    var driverNames = Object.keys(v.drivers).join(", ") || "No driver assigned";
    doc.text("Driver(s): " + driverNames, margin, yPos + 5);
    yPos += 10;

    var dayBlocks = buildLegacyDayBlocks(v.trips, addrMap);
    dayBlocks.forEach(function (block) {
      var body = [];
      var firstTrip = block.rows[0].trip;
      var prevTrip  = block.rows[0].prevTrip;

      body.push([{ content: block.date + "  Starting from: " + addressForPoint(prevTrip ? prevTrip.stopPoint : null, addrMap), colSpan: 7, styles: { fontStyle: "bold", fillColor: [245, 245, 245] } }]);

      if (prevTrip) {
        // No duration on this row — the gap to the previous trip is overnight/
        // parked time, not idle (see HTML renderer note).
        body.push([
          fmtTime(firstTrip.start),
          { content: "(Ignition On)", colSpan: 3, styles: { textColor: [217, 119, 6] } },
          "\u2014",
          "", ""
        ]);
      }

      block.rows.forEach(function (r) {
        var t = r.trip;
        var driveMins = parseDurationToMins(t.drivingDuration); if (driveMins == null) driveMins = durationMins(t.start, t.stop);
        var idleMins  = parseDurationToMins(t.idlingDuration) || 0;
        var stopMins  = legacyStopMins(t);
        body.push([
          fmtTime(t.start),
          milesFromDistance(t.distance).toFixed(2) + " mi\n" + fmtDurPrecise(driveMins),
          addressForPoint(t.stopPoint, addrMap),
          fmtTime(t.stop),
          fmtDurPrecise(idleMins),
          stopMins != null ? fmtDurPrecise(stopMins) : "\u2014",
          ""
        ]);
      });

      var dayDistM = block.rows.reduce(function (s, r) { return s + (r.trip.distance || 0); }, 0);
      var dayDrive = block.rows.reduce(function (s, r) { var dm = parseDurationToMins(r.trip.drivingDuration); return s + (dm != null ? dm : durationMins(r.trip.start, r.trip.stop)); }, 0);
      var dayIdle  = block.rows.reduce(function (s, r) { return s + (parseDurationToMins(r.trip.idlingDuration) || 0); }, 0);
      var dayCount = block.rows.filter(function (r) { return legacyStopMins(r.trip) != null; });
      var dayStop  = dayCount.reduce(function (s, r) { return s + legacyStopMins(r.trip); }, 0);
      var totalStyles = { fontStyle: "bold", fillColor: [245, 245, 245] };
      body.push([
        { content: block.date + " Total", styles: totalStyles },
        { content: milesFromDistance(dayDistM).toFixed(2) + " mi in " + fmtDurWhole(dayDrive), styles: totalStyles },
        { content: "", styles: totalStyles },
        { content: "", styles: totalStyles },
        { content: fmtDurWhole(dayIdle), styles: totalStyles },
        { content: fmtDurWhole(dayStop), styles: totalStyles },
        { content: dayCount.length + " stops", styles: totalStyles }
      ]);

      doc.autoTable({
        startY: yPos, head: headers, body: body,
        margin: { left: margin, right: margin },
        styles: { fontSize: 8, cellPadding: 2.2, textColor: [45, 55, 72] },
        headStyles: { fillColor: [0, 120, 212], textColor: 255, fontStyle: "bold", fontSize: 8 },
        columnStyles: { 6: { cellWidth: 16 } }
      });
      yPos = doc.lastAutoTable.finalY + 6;
      if (yPos > pageH - 30) { doc.addPage(); yPos = 16; }
    });
    yPos += 4;
  });

  var pageCount = doc.internal.getNumberOfPages();
  for (var p = 1; p <= pageCount; p++) {
    doc.setPage(p); doc.setFontSize(8); doc.setTextColor(160, 160, 160);
    doc.text("Confidential", margin, pageH - 8);
    doc.text("Page " + p + " of " + pageCount, pageW - margin, pageH - 8, { align: "right" });
  }
  doc.save("legacy_trip_history_" + from + "_to_" + to + ".pdf");
}

// ─── Report helpers ────────────────────────────────────────────────────────
function showReportSummary(cards) {
  var el = document.getElementById("report-summary");
  el.innerHTML = cards.join("");
  el.classList.remove("hidden");
}
function summaryCard(label, value, unit) {
  return "<div class='summary-card'><div class='summary-label'>" + label + "</div><div class='summary-value'>" + value + "<span class='summary-unit'>" + unit + "</span></div></div>";
}
function reportError(err) {
  document.getElementById("report-output").innerHTML = "<p class='placeholder' style='color:var(--critical)'>Error: " + esc((err && err.message) || String(err)) + "</p>";
}

// ─── Generic helpers ───────────────────────────────────────────────────────
function apiCall(method, params, onSuccess, onError) {
  state.api.call(method, params, onSuccess, onError || function (err) { console.error(method, err); });
}
function downloadCsvBlob(rows, filename) {
  var csv  = rows.map(function (r) { return r.map(function (c) { return '"' + String(c).replace(/"/g, '""') + '"'; }).join(","); }).join("\n");
  var blob = new Blob([csv], { type: "text/csv" });
  var url  = URL.createObjectURL(blob);
  var a    = document.createElement("a"); a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}
function fmtDateInput(d)      { return d.toISOString().slice(0, 10); }
function fmtDateShort(iso)    { return iso ? new Date(iso).toISOString().slice(0, 10) : ""; }
function fmtDateReadable(iso) { if (!iso) return ""; var d = new Date(iso); return d.toLocaleDateString(undefined, { day: "2-digit", month: "short", year: "numeric" }); }
function fmtTime(iso)         { return iso ? new Date(iso).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }) : "--"; }
function durationMins(a, b)   { return (!a || !b) ? 0 : (new Date(b) - new Date(a)) / 60000; }
// Single quotes MUST be escaped here. Every attribute in this file is written with
// single-quoted delimiters, and the rison Trip History URLs quote their date values
// with apostrophes, so leaving ' alone silently truncates every href at the first
// date — the link then lands on a bare map instead of the trip.
function esc(str) {
  return String(str == null ? "" : str)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
function showError(msg)       { document.getElementById("loading").innerHTML = "<p style='color:var(--critical);padding:40px'>" + msg + "</p>"; }
