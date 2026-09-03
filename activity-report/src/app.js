"use strict";

// ─── Activity Report — standalone MyGeotab add-in ─────────────────────────────
// A single-purpose extraction of the Activity Report from the Smart Insights
// add-in: a per-vehicle, per-day trip log with start location, distance, driving
// duration, stop location, arrival time, idle and stop duration, selectable
// engine-hour columns, daily totals, and CSV / Excel / branded PDF export.
//
// Units follow the LOGGED-IN USER, not the database. MyGeotab returns everything
// metric; User.isMetric decides how it is displayed. See the Units section.
//
// Addresses are resolved via MyGeotab's own GetAddresses method (Trip.stopPoint →
// coordinates → GetAddresses → ReverseGeocodeAddress): coordinates never leave
// Geotab's infrastructure, and no third-party geocoder is used.
//
// KEEP IN SYNC: this file is a lift of the Activity Report functions in the
// parent Smart Insights src/app.js. If you fix a bug here, mirror it there (and
// vice versa) until the two are formally split into separate repos. Two helpers
// are renamed for readability in a single-report file — the parent calls them
// buildLegacyByDevice / buildLegacyDayBlocks, here they are buildByDevice /
// buildDayBlocks (they are shared with the Legacy Trip History report upstream).
//
// Everything is wrapped in an IIFE. MyGeotab injects add-in pages into its own
// document rather than iframing them, so add-ins installed side by side share one
// global scope — a bare `var state` here would collide with the parent add-in's.
(function () {

// ─── State ─────────────────────────────────────────────────────────────────
var state = {
  api:              null,
  reportRaw:        [],
  reportData:       [],
  deviceMap:        {},
  dbName:           "",
  activityByDevice: null,        // { deviceId: { name, trips, drivers } }
  activityAddrMap:  null,        // "lat,lng" -> resolved address string
  unitSystem:       "Imperial",  // from the logged-in User.isMetric ("Metric" or "Imperial")
  legacyLengthEl:   null
};

// Settings are stored under MyGeotab's own origin, which every add-in on that
// database shares, so the key is namespaced away from the parent add-in's.
var SETTINGS_KEY = "activity_report_standalone_settings";

// ─── MyGeotab addin lifecycle ──────────────────────────────────────────────
// The registry must be reached through window, not a bare `geotab`: inside this
// IIFE a `var geotab` fallback would hoist and shadow the real global, so the
// add-in would register itself on a throwaway object and never load.
if (!window.geotab) window.geotab = {};
if (!window.geotab.addin) window.geotab.addin = {};

window.geotab.addin.activityReport = function () {
  return {
    initialize: function (api, freshState, callback) {
      try {
        state.api = api;
        if (freshState && freshState.database) {
          state.dbName = freshState.database;
          document.getElementById("db-name").textContent = freshState.database;
        }
        // Units are a property of the person looking, not of the database, so
        // they are resolved before anything renders.
        loadUserUnits(function () {
          setupReports();
          restoreSettings();
          document.getElementById("loading").classList.add("hidden");
          document.getElementById("main").classList.remove("hidden");
        });
      } catch (err) {
        showError("Init error: " + err.message);
      }
      if (callback) callback();
    },
    focus: function () {},
    blur:  function () {}
  };
};

// ─── Units ───────────────────────────────────────────────────────────────────
// MyGeotab hands everything back metric regardless of who is asking: Trip.distance
// in km, speeds in km/h, fuelUsed in litres. How those get *displayed* is a
// property of the person looking, not of the database — User.isMetric is a
// per-user regional setting. Read once at init; SystemSettings.measurementSystem
// is only a fallback for when the User read fails, and it is database-wide so it
// can disagree with the user's own profile.
//
// Every distance in this file goes through these helpers. Do not reintroduce a
// bare * 0.621371 or a hardcoded "mi" at a call site.
var MI_PER_KM = 0.621371;

function loadUserUnits(cb) {
  var done = false;
  function finish() { if (!done) { done = true; cb(); } }

  function applyUser(user) {
    if (user && typeof user.isMetric === "boolean") {
      state.unitSystem = user.isMetric ? "Metric" : "Imperial";
      finish();
      return true;
    }
    return false;
  }

  function fallbackToSystemSettings() {
    apiCall("Get", { typeName: "SystemSettings" }, function (settings) {
      if (settings && settings[0] && settings[0].measurementSystem) {
        state.unitSystem = settings[0].measurementSystem;
      }
      finish();
    }, finish);
  }

  function fetchUserByName(userName) {
    if (!userName) { fallbackToSystemSettings(); return; }
    apiCall("Get", { typeName: "User", search: { name: userName } }, function (users) {
      if (!users || !users.length || !applyUser(users[0])) fallbackToSystemSettings();
    }, fallbackToSystemSettings);
  }

  // getSession's callback shape varies across MyGeotab versions, and on some it
  // never fires at all, so the read is guarded by a timeout rather than trusted.
  try {
    if (state.api && typeof state.api.getSession === "function") {
      var settled = false;
      var guard = setTimeout(function () {
        if (!settled) { settled = true; fallbackToSystemSettings(); }
      }, 8000);
      state.api.getSession(function (a) {
        if (settled) return;
        settled = true;
        clearTimeout(guard);
        var creds = (a && a.credentials) || a || {};
        fetchUserByName(creds.userName || creds.username);
      });
      return;
    }
  } catch (e) { /* fall through */ }
  fallbackToSystemSettings();
}

function isMetric()      { return state.unitSystem === "Metric"; }
function distUnit()      { return isMetric() ? "km" : "mi"; }
function distVal(km)     { return isMetric() ? (km || 0) : (km || 0) * MI_PER_KM; }
function fmtDist(km, dp) { return distVal(km).toFixed(dp == null ? 1 : dp) + " " + distUnit(); }

// ─── Setup / controls ──────────────────────────────────────────────────────
function setupReports() {
  var today = new Date();
  var week  = new Date(today);
  week.setDate(today.getDate() - 6);
  document.getElementById("filter-from").value = fmtDateInput(week);
  document.getElementById("filter-to").value   = fmtDateInput(today);

  document.getElementById("run-report").addEventListener("click", runReport);
  document.getElementById("range-length").addEventListener("change", applyRangeLength);
  document.getElementById("filter-from").addEventListener("change", function () {
    syncToDateFromFromDate();
  });
  document.getElementById("filter-to").addEventListener("change", function () {
    if (rangeLengthDays() != null) syncToDateFromFromDate();
  });
  document.getElementById("export-activity-btn").addEventListener("click", function () {
    var fmt = document.getElementById("export-format").value;
    if      (fmt === "csv")   exportActivityCsv();
    else if (fmt === "excel") exportActivityExcel();
    else                      exportActivityPdf();
  });
  applyRangeLength();
}

function rangeLengthDays() {
  var raw = document.getElementById("range-length").value;
  if (raw === "custom") return null;
  var n = parseInt(raw, 10);
  return isNaN(n) || n < 1 ? null : n;
}

// Preset lengths are driven by the From date: pick a start day, and To follows.
// Custom leaves both date inputs editable.
function syncToDateFromFromDate() {
  var days = rangeLengthDays();
  if (days == null) return;
  var from = document.getElementById("filter-from").value;
  if (!from) return;
  document.getElementById("filter-to").value = addDaysDateInput(from, days - 1);
}

function applyRangeLength() {
  var days = rangeLengthDays();
  var toEl = document.getElementById("filter-to");
  if (days == null) {
    toEl.disabled = false;
    toEl.title = "";
    return;
  }

  toEl.disabled = true;
  toEl.title = "Set by the selected range";

  var fromEl = document.getElementById("filter-from");
  if (!fromEl.value) fromEl.value = fmtDateInput(new Date());

  syncToDateFromFromDate();
}

function parseDateInput(v) {
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(v || "");
  if (!m) return null;
  return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
}

function addDaysDateInput(v, days) {
  var d = parseDateInput(v);
  if (!d) return v || "";
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().slice(0, 10);
}

function runReport() {
  applyRangeLength();

  var from   = document.getElementById("filter-from").value;
  var to     = document.getElementById("filter-to").value;
  var btn    = document.getElementById("run-report");
  var output = document.getElementById("report-output");

  if (!from || !to) { alert("Please select a date range."); return; }
  if (from > to)    { alert("From date is after To date."); return; }

  btn.disabled = true; btn.textContent = "Loading...";
  output.innerHTML = "<p class='placeholder'>Fetching data...</p>";
  document.getElementById("report-summary").classList.add("hidden");
  document.getElementById("report-vehicle").classList.add("hidden");
  document.getElementById("export-format").classList.add("hidden");
  document.getElementById("export-activity-btn").classList.add("hidden");

  runActivityReport(from, to, function () { btn.disabled = false; btn.textContent = "Run"; });
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
  if (mins == null || isNaN(mins)) return "—";
  var total = Math.round(mins);
  return total < 60 ? total + "m" : Math.floor(total / 60) + "h " + (total % 60) + "m";
}
// Second-precision duration for table rows: "02m 02s" / "1h 42m 28s"
function fmtDurPrecise(mins) {
  if (mins == null || isNaN(mins)) return "—";
  var totalSecs = Math.max(0, Math.round(mins * 60));
  var h = Math.floor(totalSecs / 3600);
  var m = Math.floor((totalSecs % 3600) / 60);
  var s = totalSecs % 60;
  var mm = (m < 10 ? "0" : "") + m;
  var ss = (s < 10 ? "0" : "") + s;
  return h > 0 ? h + "h " + mm + "m " + ss + "s" : mm + "m " + ss + "s";
}

// ─── Data assembly ─────────────────────────────────────────────────────────
// Local calendar day. Everything this report displays is local time, so the day
// bucket has to be local too — toISOString() would file a 00:30 BST trip under
// the previous day, and an evening trip west of Greenwich under the next one.
function localDayKey(iso) {
  if (!iso) return "";
  var d = new Date(iso);
  if (isNaN(d.getTime())) return "";
  var m = d.getMonth() + 1, day = d.getDate();
  return d.getFullYear() + "-" + (m < 10 ? "0" : "") + m + "-" + (day < 10 ? "0" : "") + day;
}

// Group fetched trips by device, sorted chronologically. Driver names come
// straight off Trip.driverName.
//
// NOTE: unlike the Legacy Trip History report, the Activity Report deliberately
// does not suppress overnight parks — it reports raw Trip.stopDuration, so the
// last stop of a day carries the full time until the next morning's first trip.
function buildByDevice(trips, deviceMap) {
  var byDevice = {};
  trips.forEach(function (t) {
    var vid = t.device && t.device.id; if (!vid) return;
    if (!byDevice[vid]) byDevice[vid] = { name: deviceMap[vid] || vid, trips: [], drivers: {} };
    byDevice[vid].trips.push(t);
    if (t.driverName) byDevice[vid].drivers[t.driverName] = 1;
  });
  Object.keys(byDevice).forEach(function (vid) {
    byDevice[vid].trips.sort(function (a, b) { return new Date(a.start) - new Date(b.start); });
  });
  return byDevice;
}

// Round a coordinate to ~1m precision so repeat stops (depots, regular
// customer sites) share one GetAddresses lookup instead of a fresh one each time.
function coordKey(pt) { return pt.y.toFixed(5) + "," + pt.x.toFixed(5); }

// Collect every unique StopPoint across the whole dataset, resolve them all in
// as few GetAddresses calls as possible (450/min limit, so chunk generously),
// and return a lookup map: "lat,lng" -> formatted address string.
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

  var CHUNK = 400; // stay comfortably under the 450/min GetAddresses limit
  var addrMap = {};
  var idx = 0;

  function runChunk() {
    if (idx >= keys.length) { cb(addrMap); return; }
    var sliceKeys = keys.slice(idx, idx + CHUNK);
    var coords = sliceKeys.map(function (k) { return { x: uniq[k].x, y: uniq[k].y }; });
    apiCall("GetAddresses", { coordinates: coords }, function (results) {
      (results || []).forEach(function (addr, i) {
        addrMap[sliceKeys[i]] = formatReverseGeocodeAddress(addr);
      });
      idx += CHUNK;
      runChunk();
    }, function () {
      // On failure, leave these keys unresolved (renderer falls back to coordinates)
      idx += CHUNK;
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

// Where a trip began. Trip has no startPoint — a vehicle starts where it last
// stopped, so the previous trip's stopPoint is the start location. The first trip
// of the fetched range has no predecessor and is genuinely unknown.
function startAddressForRow(row, addrMap) {
  return addressForPoint(row.prevTrip ? row.prevTrip.stopPoint : null, addrMap);
}

function addressForPoint(pt, addrMap) {
  if (!pt || pt.y == null || pt.x == null) return "(location unknown)";
  var resolved = addrMap[coordKey(pt)];
  return resolved || (pt.y.toFixed(5) + ", " + pt.x.toFixed(5));
}

// Build the day-block rows for one vehicle's trip list, including the synthetic
// "Starting from" / "(Ignition On)" lead-in row derived from the previous trip
// (which may be from an earlier day, even outside the selected range if it
// happened to already be in the fetched dataset).
function buildDayBlocks(vTrips) {
  var days = {}; // 'yyyy-mm-dd' -> { date, rows:[{trip, prevTrip}] }
  var order = [];
  vTrips.forEach(function (t, i) {
    var day = localDayKey(t.start);
    if (!days[day]) { days[day] = { date: day, rows: [] }; order.push(day); }
    days[day].rows.push({ trip: t, prevTrip: i > 0 ? vTrips[i - 1] : null });
  });
  return order.map(function (day) { return days[day]; });
}

// ─── Column selection ──────────────────────────────────────────────────────
function fmtActivityDistance(km) { return fmtDist(km, 2); }

function fmtActivityDistanceKpi(km) {
  return { value: distVal(km).toFixed(0), unit: distUnit() };
}

function getActivityCols() {
  return ["total-engine", "drive-only", "idle-only", "work-split"].filter(function (id) {
    var el = document.getElementById("col-" + id);
    return el && el.checked;
  });
}

function activityEngineColHeaders(cols) {
  var headers = [];
  cols.forEach(function (id) {
    if (id === "total-engine") headers.push("Engine On (Total)");
    else if (id === "drive-only") headers.push("Driving Time");
    else if (id === "idle-only") headers.push("Idle Time");
    else if (id === "work-split") { headers.push("Work Hours Engine"); headers.push("After-hours Engine"); }
  });
  return headers;
}

function activityEngineColValues(trip, cols) {
  var vals = [];
  cols.forEach(function (id) {
    var driveMins = parseDurationToMins(trip.drivingDuration); if (driveMins == null) driveMins = durationMins(trip.start, trip.stop);
    var idleMins  = parseDurationToMins(trip.idlingDuration) || 0;
    if (id === "total-engine") {
      vals.push(fmtDurPrecise(driveMins + idleMins));
    } else if (id === "drive-only") {
      vals.push(fmtDurPrecise(driveMins));
    } else if (id === "idle-only") {
      vals.push(fmtDurPrecise(idleMins));
    } else if (id === "work-split") {
      var workDrive = parseDurationToMins(trip.workDrivingDuration) || 0;
      var workIdle  = parseDurationToMins(trip.workIdlingDuration)  || 0;
      var ahDrive   = parseDurationToMins(trip.afterHoursDrivingDuration) || 0;
      var ahIdle    = parseDurationToMins(trip.afterHoursIdlingDuration)  || 0;
      vals.push(fmtDurPrecise(workDrive + workIdle));
      vals.push(fmtDurPrecise(ahDrive + ahIdle));
    }
  });
  return vals;
}

function activityEngineDayTotals(rows, cols) {
  var totals = [];
  cols.forEach(function (id) {
    var driveMins = rows.reduce(function (s, r) { var dm = parseDurationToMins(r.trip.drivingDuration); return s + (dm != null ? dm : durationMins(r.trip.start, r.trip.stop)); }, 0);
    var idleMins  = rows.reduce(function (s, r) { return s + (parseDurationToMins(r.trip.idlingDuration) || 0); }, 0);
    if (id === "total-engine") {
      totals.push(fmtDurWhole(driveMins + idleMins));
    } else if (id === "drive-only") {
      totals.push(fmtDurWhole(driveMins));
    } else if (id === "idle-only") {
      totals.push(fmtDurWhole(idleMins));
    } else if (id === "work-split") {
      var workDrive = rows.reduce(function (s, r) { return s + (parseDurationToMins(r.trip.workDrivingDuration) || 0); }, 0);
      var workIdle  = rows.reduce(function (s, r) { return s + (parseDurationToMins(r.trip.workIdlingDuration) || 0); }, 0);
      var ahDrive   = rows.reduce(function (s, r) { return s + (parseDurationToMins(r.trip.afterHoursDrivingDuration) || 0); }, 0);
      var ahIdle    = rows.reduce(function (s, r) { return s + (parseDurationToMins(r.trip.afterHoursIdlingDuration) || 0); }, 0);
      totals.push(fmtDurWhole(workDrive + workIdle));
      totals.push(fmtDurWhole(ahDrive + ahIdle));
    }
  });
  return totals;
}

// ─── Run ───────────────────────────────────────────────────────────────────
function runActivityReport(from, to, done) {
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
      var byDevice = buildByDevice(trips, deviceMap);
      state.reportRaw        = trips;
      state.reportData       = trips;
      state.activityByDevice = byDevice;
      populateReportVehicleFilter(trips, deviceMap);

      document.getElementById("report-output").innerHTML = "<p class='placeholder'>Resolving stop addresses…</p>";
      resolveStopAddresses(byDevice, function (addrMap) {
        state.activityAddrMap = addrMap;

        var totalDistKm = trips.reduce(function (s, t) { return s + (t.distance || 0); }, 0);
        var totalDrive  = trips.reduce(function (s, t) { var dm = parseDurationToMins(t.drivingDuration); return s + (dm != null ? dm : durationMins(t.start, t.stop)); }, 0);
        var totalIdle   = trips.reduce(function (s, t) { return s + (parseDurationToMins(t.idlingDuration) || 0); }, 0);
        var totalStop   = trips.reduce(function (s, t) { return s + (parseDurationToMins(t.stopDuration) || 0); }, 0);
        var numStops    = trips.length;
        var avgStop     = numStops ? totalStop / numStops : 0;
        var distKpi     = fmtActivityDistanceKpi(totalDistKm);

        showReportSummary([
          summaryCard("Total Distance",         distKpi.value, distKpi.unit),
          summaryCard("Total Stop Duration",    fmtDurWhole(totalStop), ""),
          summaryCard("Total Idle Time",        fmtDurWhole(totalIdle), ""),
          summaryCard("Total Travel Duration",  fmtDurWhole(totalDrive), ""),
          summaryCard("Average Stop Duration",  fmtDurWhole(avgStop), ""),
          summaryCard("Number of Stops",        numStops, "")
        ]);

        renderActivityReport(byDevice, addrMap);
        document.getElementById("export-format").classList.remove("hidden");
        document.getElementById("export-activity-btn").classList.remove("hidden");
        saveSettings(from, to);
        done();
      });
    }, function (err) { reportError(err); done(); });
  }, function (err) { reportError(err); done(); });
}

function dayBlockStats(block) {
  var dayDistKm = block.rows.reduce(function (s, r) { return s + (r.trip.distance || 0); }, 0);
  var dayDrive  = block.rows.reduce(function (s, r) {
    var dm = parseDurationToMins(r.trip.drivingDuration);
    return s + (dm != null ? dm : durationMins(r.trip.start, r.trip.stop));
  }, 0);
  var dayIdle   = block.rows.reduce(function (s, r) { return s + (parseDurationToMins(r.trip.idlingDuration) || 0); }, 0);
  var dayStop   = block.rows.reduce(function (s, r) { return s + (parseDurationToMins(r.trip.stopDuration) || 0); }, 0);
  return {
    distKm: dayDistKm,
    driveMins: dayDrive,
    idleMins: dayIdle,
    engineMins: dayDrive + dayIdle,
    stopMins: dayStop,
    stops: block.rows.length
  };
}

function vehicleStats(v) {
  var distKm = v.trips.reduce(function (s, t) { return s + (t.distance || 0); }, 0);
  var drive  = v.trips.reduce(function (s, t) {
    var dm = parseDurationToMins(t.drivingDuration);
    return s + (dm != null ? dm : durationMins(t.start, t.stop));
  }, 0);
  var idle   = v.trips.reduce(function (s, t) { return s + (parseDurationToMins(t.idlingDuration) || 0); }, 0);
  var stop   = v.trips.reduce(function (s, t) { return s + (parseDurationToMins(t.stopDuration) || 0); }, 0);
  return {
    distKm: distKm,
    driveMins: drive,
    idleMins: idle,
    engineMins: drive + idle,
    stopMins: stop,
    stops: v.trips.length
  };
}

// ─── Render (in-app preview) ───────────────────────────────────────────────
function renderActivityReport(byDevice, addrMap) {
  var cols   = getActivityCols();
  var output = document.getElementById("report-output");
  var vids   = Object.keys(byDevice);
  if (!vids.length) { output.innerHTML = "<p class='placeholder'>No trips found.</p>"; return; }

  var extraHeaders = activityEngineColHeaders(cols);
  var baseHeaders  = ["Start Time", "Start Location", "Distance / Duration", "Stop Location", "Arrival Time", "Idle Duration", "Stop Duration"];
  var allHeaders   = baseHeaders.concat(extraHeaders);
  var colSpan      = allHeaders.length;

  output.innerHTML = vids.map(function (vid) {
    var v = byDevice[vid];
    var driverNames = Object.keys(v.drivers).join(", ") || "No driver assigned";
    var dayBlocks   = buildDayBlocks(v.trips);
    var vStats      = vehicleStats(v);

    var dayHtml = dayBlocks.map(function (block) {
      var stats = dayBlockStats(block);
      var rowsHtml = block.rows.map(function (r) {
        var t          = r.trip;
        var driveMins  = parseDurationToMins(t.drivingDuration);
        if (driveMins == null) driveMins = durationMins(t.start, t.stop);
        var idleMins   = parseDurationToMins(t.idlingDuration) || 0;
        var stopMins   = parseDurationToMins(t.stopDuration);
        var engineVals = activityEngineColValues(t, cols);
        return "<tr>"
          + "<td>" + fmtTime(t.start) + "</td>"
          + "<td>" + esc(startAddressForRow(r, addrMap)) + "</td>"
          + "<td>" + fmtActivityDistance(t.distance) + "<br><span style='color:var(--text-muted);font-size:11px'>" + fmtDurPrecise(driveMins) + "</span></td>"
          + "<td>" + esc(addressForPoint(t.stopPoint, addrMap)) + "</td>"
          + "<td>" + fmtTime(t.stop) + "</td>"
          + "<td>" + fmtDurPrecise(idleMins) + "</td>"
          + "<td>" + (stopMins != null ? fmtDurPrecise(stopMins) : "—") + "</td>"
          + engineVals.map(function (v2) { return "<td>" + v2 + "</td>"; }).join("")
          + "</tr>";
      }).join("");

      var engineTotals = activityEngineDayTotals(block.rows, cols);
      var firstTrip = block.rows[0].trip;
      var prevTrip  = block.rows[0].prevTrip;

      var startingFromHtml = "<tr><td colspan='" + colSpan + "' style='font-weight:600;background:var(--surface-2)'>"
        + block.date + " — Starting from: " + esc(addressForPoint(prevTrip ? prevTrip.stopPoint : null, addrMap))
        + "</td></tr>";
      var ignitionHtml = prevTrip
        ? "<tr style='color:var(--accent)'><td>" + fmtTime(firstTrip.start)
          + "</td><td colspan='4'>(Ignition On)</td><td>—</td><td></td>"
          + extraHeaders.map(function () { return "<td></td>"; }).join("") + "</tr>"
        : "";

      var totalRowCells = [
        "<td>" + block.date + " Total</td>",
        "<td></td>",
        "<td>" + fmtActivityDistance(stats.distKm) + " in " + fmtDurWhole(stats.driveMins) + "</td>",
        "<td></td>",
        "<td>" + stats.stops + " stops</td>",
        "<td>" + fmtDurWhole(stats.idleMins) + "</td>",
        "<td>" + fmtDurWhole(stats.stopMins) + "</td>"
      ].concat(engineTotals.map(function (v3) { return "<td>" + v3 + "</td>"; })).join("");

      var daySummary = block.date
        + " · " + stats.stops + " stops"
        + " · " + fmtActivityDistance(stats.distKm)
        + " · engine on " + fmtDurWhole(stats.engineMins)
        + " · drive " + fmtDurWhole(stats.driveMins)
        + " · idle " + fmtDurWhole(stats.idleMins);

      return "<details class='ar-day-fold'>"
        + "<summary class='ar-day-summary'>"
        + "<span class='ar-day-main'>" + esc(block.date) + "</span>"
        + "<span class='ar-day-meta'>" + esc(daySummary) + "</span>"
        + "</summary>"
        + "<div class='ar-day-body'>"
        + "<div class='dd-table-wrap'><table class='dd-table'>"
        + "<thead><tr>" + allHeaders.map(function (h) { return "<th>" + h + "</th>"; }).join("") + "</tr></thead>"
        + "<tbody>" + startingFromHtml + ignitionHtml + rowsHtml
        + "<tr style='font-weight:700;background:var(--surface-2)'>" + totalRowCells + "</tr>"
        + "</tbody></table></div>"
        + "</div></details>";
    }).join("");

    var vehicleSummary = dayBlocks.length + " day" + (dayBlocks.length === 1 ? "" : "s")
      + " · " + vStats.stops + " stops"
      + " · " + fmtActivityDistance(vStats.distKm)
      + " · engine on " + fmtDurWhole(vStats.engineMins)
      + " · drive " + fmtDurWhole(vStats.driveMins)
      + " · idle " + fmtDurWhole(vStats.idleMins);

    return "<details class='ar-vehicle-fold'>"
      + "<summary class='ar-vehicle-summary'>"
      + "<span class='ar-vehicle-main'>" + esc(v.name) + "</span>"
      + "<span class='ar-vehicle-meta'>" + esc(vehicleSummary) + "</span>"
      + "</summary>"
      + "<div class='ar-vehicle-body'>"
      + "<div class='ar-vehicle-driver'>Driver(s): " + esc(driverNames) + "</div>"
      + dayHtml
      + "</div></details>";
  }).join("");
}

// ─── Settings persistence ──────────────────────────────────────────────────
function saveSettings(from, to) {
  try {
    localStorage.setItem(SETTINGS_KEY, JSON.stringify({
      from: from,
      to:   to,
      cols: getActivityCols(),
      len:  document.getElementById("range-length").value
    }));
  } catch (e) {}
}

function restoreSettings() {
  try {
    var raw = localStorage.getItem(SETTINGS_KEY);
    if (raw) {
      var s = JSON.parse(raw);
      if (s.len)  { var lenEl = document.getElementById("range-length"); if (lenEl) lenEl.value = s.len; }
      if (s.from) { var fromEl = document.getElementById("filter-from"); if (fromEl) fromEl.value = s.from; }
      if (s.to)   { var toEl   = document.getElementById("filter-to");   if (toEl)   toEl.value   = s.to;   }
      // Re-check saved column selections (default is only total-engine checked)
      ["total-engine", "drive-only", "idle-only", "work-split"].forEach(function (id) {
        var el = document.getElementById("col-" + id);
        if (el) el.checked = (s.cols || ["total-engine"]).indexOf(id) !== -1;
      });
    }

    applyRangeLength();
  } catch (e) {}
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
  state.reportData = !vid ? raw : raw.filter(function (item) {
    return item.device && item.device.id === vid;
  });
  var filteredByDevice = buildByDevice(state.reportData, state.deviceMap);
  state.activityByDevice = filteredByDevice;
  renderActivityReport(filteredByDevice, state.activityAddrMap || {});
}

// ─── CSV export ────────────────────────────────────────────────────────────
function exportActivityCsv() {
  var data    = state.reportData;
  var addrMap = state.activityAddrMap || {};
  var cols    = getActivityCols();
  if (!data || !data.length) return;
  var from = document.getElementById("filter-from").value;
  var to   = document.getElementById("filter-to").value;
  var baseHeaders = ["Vehicle", "Driver(s)", "Date", "Start Time", "Start Location", "Distance (" + distUnit() + ")", "Driving Duration", "Stop Location", "Arrival Time", "Idle Duration", "Stop Duration"];
  var rows = [baseHeaders.concat(activityEngineColHeaders(cols))];
  var byDevice = buildByDevice(data, state.deviceMap);
  Object.keys(byDevice).forEach(function (vid) {
    var v = byDevice[vid];
    var driverNames = Object.keys(v.drivers).join(", ") || "No driver assigned";
    v.trips.forEach(function (t, i) {
      var driveMins = parseDurationToMins(t.drivingDuration); if (driveMins == null) driveMins = durationMins(t.start, t.stop);
      var idleMins  = parseDurationToMins(t.idlingDuration) || 0;
      var stopMins  = parseDurationToMins(t.stopDuration);
      // trips are sorted per vehicle, so the previous one holds this trip's origin
      var startAddr = startAddressForRow({ prevTrip: i > 0 ? v.trips[i - 1] : null }, addrMap);
      var baseRow = [
        v.name, driverNames, fmtDateShort(t.start), fmtTime(t.start), startAddr,
        distVal(t.distance).toFixed(2), fmtDurPrecise(driveMins),
        addressForPoint(t.stopPoint, addrMap), fmtTime(t.stop),
        fmtDurPrecise(idleMins), stopMins != null ? fmtDurPrecise(stopMins) : ""
      ];
      rows.push(baseRow.concat(activityEngineColValues(t, cols)));
    });
  });
  if (rows.length > 1) downloadCsvBlob(rows, "activity_report_" + from + "_to_" + to + ".csv");
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
function exportActivityPdf() {
  var byDevice = state.activityByDevice;
  var addrMap  = state.activityAddrMap || {};
  var cols     = getActivityCols();
  if (!byDevice || !Object.keys(byDevice).length) { alert("Run the report first."); return; }
  var from = document.getElementById("filter-from").value;
  var to   = document.getElementById("filter-to").value;

  var doc    = new jspdf.jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
  var pageW  = doc.internal.pageSize.getWidth();
  var pageH  = doc.internal.pageSize.getHeight();
  var margin = 14;

  doc.setFillColor(GEOTAB_NAVY[0], GEOTAB_NAVY[1], GEOTAB_NAVY[2]);
  doc.rect(0, 0, pageW, 24, "F");
  drawPdfHeaderLogo(doc, pageW, margin, 24);
  doc.setTextColor(255, 255, 255); doc.setFontSize(16); doc.setFont("helvetica", "bold");
  doc.text("Activity Report", margin, 11);
  doc.setFontSize(9); doc.setFont("helvetica", "normal");
  var now = new Date();
  doc.text("Created: " + fmtDateReadable(now.toISOString()) + " " + fmtTime(now.toISOString()), margin, 18);

  var yPos = 32;
  doc.setFillColor(20, 20, 20);
  doc.rect(margin, yPos, pageW - margin * 2, 8, "F");
  doc.setTextColor(255, 255, 255); doc.setFontSize(10); doc.setFont("helvetica", "bold");
  doc.text("Report Totals for: " + from + " – " + to, margin + 3, yPos + 5.5);
  yPos += 14;

  var allTrips = [];
  Object.keys(byDevice).forEach(function (vid) { allTrips = allTrips.concat(byDevice[vid].trips); });
  var totalDistKm = allTrips.reduce(function (s, t) { return s + (t.distance || 0); }, 0);
  var totalDrive  = allTrips.reduce(function (s, t) { var dm = parseDurationToMins(t.drivingDuration); return s + (dm != null ? dm : durationMins(t.start, t.stop)); }, 0);
  var totalIdle   = allTrips.reduce(function (s, t) { return s + (parseDurationToMins(t.idlingDuration) || 0); }, 0);
  var totalStop   = allTrips.reduce(function (s, t) { return s + (parseDurationToMins(t.stopDuration) || 0); }, 0);
  var numStops    = allTrips.length;
  var avgStop     = numStops ? totalStop / numStops : 0;
  var distKpi     = fmtActivityDistanceKpi(totalDistKm);

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
    { label: "Total Distance (" + distKpi.unit + ")", value: distKpi.value },
    { label: "Average Stop Duration", value: fmtDurWhole(avgStop) },
    { label: "Number of Stops",       value: String(numStops) }
  ], yPos);
  yPos += 24;

  var extraHeaders = activityEngineColHeaders(cols);
  var headers = [["Start Time", "Start Location", "Distance / Duration", "Stop Location", "Arrival Time", "Idle Duration", "Stop Duration"].concat(extraHeaders)];
  // Two free-text address columns in one landscape row will squeeze everything else
  // unless they are pinned; the rest of the columns hold short fixed-width values.
  var pdfColumnStyles = { 0: { cellWidth: 16 }, 1: { cellWidth: 52 }, 2: { cellWidth: 26 }, 3: { cellWidth: 52 }, 4: { cellWidth: 16 } };

  Object.keys(byDevice).forEach(function (vid) {
    var v = byDevice[vid];
    if (yPos > pageH - 40) { doc.addPage(); yPos = 16; }
    doc.setTextColor(20, 20, 20); doc.setFontSize(12); doc.setFont("helvetica", "bold");
    doc.text(v.name, margin, yPos);
    doc.setFontSize(9); doc.setFont("helvetica", "normal"); doc.setTextColor(100, 100, 100);
    doc.text("Driver(s): " + (Object.keys(v.drivers).join(", ") || "No driver assigned"), margin, yPos + 5);
    yPos += 10;

    var dayBlocks = buildDayBlocks(v.trips);
    dayBlocks.forEach(function (block) {
      var body = [];
      var firstTrip = block.rows[0].trip;
      var prevTrip  = block.rows[0].prevTrip;
      var numCols   = 7 + extraHeaders.length;
      body.push([{ content: block.date + "  Starting from: " + addressForPoint(prevTrip ? prevTrip.stopPoint : null, addrMap), colSpan: numCols, styles: { fontStyle: "bold", fillColor: [245, 245, 245] } }]);

      if (prevTrip) {
        var ignRow = [fmtTime(firstTrip.start), { content: "(Ignition On)", colSpan: 4, styles: { textColor: [217, 119, 6] } }, "—", ""];
        for (var ei = 0; ei < extraHeaders.length; ei++) ignRow.push("");
        body.push(ignRow);
      }

      block.rows.forEach(function (r) {
        var t         = r.trip;
        var driveMins = parseDurationToMins(t.drivingDuration); if (driveMins == null) driveMins = durationMins(t.start, t.stop);
        var idleMins  = parseDurationToMins(t.idlingDuration) || 0;
        var stopMins  = parseDurationToMins(t.stopDuration);
        var row = [
          fmtTime(t.start),
          startAddressForRow(r, addrMap),
          fmtActivityDistance(t.distance) + "\n" + fmtDurPrecise(driveMins),
          addressForPoint(t.stopPoint, addrMap),
          fmtTime(t.stop),
          fmtDurPrecise(idleMins),
          stopMins != null ? fmtDurPrecise(stopMins) : "—"
        ];
        body.push(row.concat(activityEngineColValues(t, cols)));
      });

      var dayDistKm = block.rows.reduce(function (s, r) { return s + (r.trip.distance || 0); }, 0);
      var dayDrive  = block.rows.reduce(function (s, r) { var dm = parseDurationToMins(r.trip.drivingDuration); return s + (dm != null ? dm : durationMins(r.trip.start, r.trip.stop)); }, 0);
      var dayIdle   = block.rows.reduce(function (s, r) { return s + (parseDurationToMins(r.trip.idlingDuration) || 0); }, 0);
      var dayStop   = block.rows.reduce(function (s, r) { return s + (parseDurationToMins(r.trip.stopDuration) || 0); }, 0);
      var totalStyles = { fontStyle: "bold", fillColor: [245, 245, 245] };
      var totalRow = [
        { content: block.date + " Total", styles: totalStyles },
        { content: "", styles: totalStyles },
        { content: fmtActivityDistance(dayDistKm) + " in " + fmtDurWhole(dayDrive), styles: totalStyles },
        { content: "", styles: totalStyles },
        { content: "", styles: totalStyles },
        { content: fmtDurWhole(dayIdle), styles: totalStyles },
        { content: fmtDurWhole(dayStop), styles: totalStyles }
      ];
      activityEngineDayTotals(block.rows, cols).forEach(function (v) {
        totalRow.push({ content: v, styles: totalStyles });
      });
      body.push(totalRow);

      doc.autoTable({
        startY: yPos, head: headers, body: body,
        margin: { left: margin, right: margin },
        columnStyles: pdfColumnStyles,
        styles: { fontSize: 7.5, cellPadding: 2, textColor: [45, 55, 72], overflow: "linebreak" },
        headStyles: { fillColor: [0, 120, 212], textColor: 255, fontStyle: "bold", fontSize: 7.5 }
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
  doc.save("activity_report_" + from + "_to_" + to + ".pdf");
}

// ─── Excel export ──────────────────────────────────────────────────────────
function exportActivityExcel() {
  var byDevice = state.activityByDevice;
  var addrMap  = state.activityAddrMap || {};
  var cols     = getActivityCols();
  if (!byDevice || !Object.keys(byDevice).length) { alert("Run the report first."); return; }
  if (typeof ExcelJS === "undefined") { alert("Excel library not loaded. Please check your internet connection."); return; }
  var from  = document.getElementById("filter-from").value;
  var to    = document.getElementById("filter-to").value;
  var dUnit = distUnit();
  var vids  = Object.keys(byDevice);

  var allTrips = [];
  vids.forEach(function (vid) { allTrips = allTrips.concat(byDevice[vid].trips); });

  var totalDistKm = allTrips.reduce(function (s, t) { return s + (t.distance || 0); }, 0);
  var totalDrive  = allTrips.reduce(function (s, t) {
    var dm = parseDurationToMins(t.drivingDuration);
    return s + (dm != null ? dm : durationMins(t.start, t.stop));
  }, 0);
  var totalIdle   = allTrips.reduce(function (s, t) { return s + (parseDurationToMins(t.idlingDuration) || 0); }, 0);
  var totalStop   = allTrips.reduce(function (s, t) { return s + (parseDurationToMins(t.stopDuration) || 0); }, 0);
  var distKpi     = fmtActivityDistanceKpi(totalDistKm);

  var wb = new ExcelJS.Workbook();
  wb.creator = "Activity Report";
  wb.created = new Date();

  var extraHeaders  = activityEngineColHeaders(cols);
  var baseHeaders   = ["Start Time", "Start Location", "Distance / Duration", "Stop Location", "Arrival Time", "Idle Duration", "Stop Duration"];
  var allColHeaders = baseHeaders.concat(extraHeaders);

  var ws = wb.addWorksheet("Activity Report");
  ws.columns = [
    { width: 12 }, { width: 44 }, { width: 28 }, { width: 44 },
    { width: 14 }, { width: 16 }, { width: 16 }
  ].concat(extraHeaders.map(function () { return { width: 22 }; }));

  var titleRow = ws.addRow(["Activity Report — " + from + " to " + to]);
  titleRow.font = { bold: true, size: 13 };
  if (state.dbName) {
    var dbRow = ws.addRow(["Database: " + state.dbName]);
    dbRow.font = { color: { argb: "FF6B7280" }, size: 10 };
  }
  ws.addRow([]);

  var metricHdr = ws.addRow(["Metric", "Value"]);
  metricHdr.font = { bold: true };
  metricHdr.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE8F4FD" } };

  [
    ["Total Distance (" + dUnit + ")", distKpi.value + " " + distKpi.unit],
    ["Total Travel Duration", fmtDurWhole(totalDrive)],
    ["Total Idle Time", fmtDurWhole(totalIdle)],
    ["Total Stop Duration", fmtDurWhole(totalStop)],
    ["Number of Stops", allTrips.length],
    ["Average Stop Duration", fmtDurWhole(allTrips.length ? totalStop / allTrips.length : 0)],
    ["Vehicles in Report", vids.length]
  ].forEach(function (m) { ws.addRow(m); });
  ws.addRow([]);

  vids.forEach(function (vid) {
    var v = byDevice[vid];
    var driverNames = Object.keys(v.drivers).join(", ") || "No driver assigned";
    var vStats = vehicleStats(v);
    var dayBlocks = buildDayBlocks(v.trips);

    var vehicleSummary = dayBlocks.length + " day" + (dayBlocks.length === 1 ? "" : "s")
      + " · " + vStats.stops + " stops"
      + " · " + fmtActivityDistance(vStats.distKm)
      + " · engine on " + fmtDurWhole(vStats.engineMins)
      + " · drive " + fmtDurWhole(vStats.driveMins)
      + " · idle " + fmtDurWhole(vStats.idleMins);

    var vehicleRow = ws.addRow([v.name, vehicleSummary]);
    vehicleRow.font = { bold: true, color: { argb: "FF27325D" } };
    vehicleRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF5F8FF" } };

    var driverRow = ws.addRow(["Driver(s): " + driverNames]);
    driverRow.font = { color: { argb: "FF6B7280" }, size: 10 };

    dayBlocks.forEach(function (block) {
      var stats = dayBlockStats(block);
      var prevTrip = block.rows[0].prevTrip;

      var daySummary = block.date
        + " · " + stats.stops + " stops"
        + " · " + fmtActivityDistance(stats.distKm)
        + " · engine on " + fmtDurWhole(stats.engineMins)
        + " · drive " + fmtDurWhole(stats.driveMins)
        + " · idle " + fmtDurWhole(stats.idleMins);

      var dayHdr = ws.addRow([block.date, daySummary]);
      dayHdr.font = { bold: true };
      dayHdr.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF1F3F7" } };

      var sfRow = ws.addRow(["Starting from: " + addressForPoint(prevTrip ? prevTrip.stopPoint : null, addrMap)]);
      sfRow.font = { italic: true, color: { argb: "FF4B5563" } };

      var hRow = ws.addRow(allColHeaders);
      hRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
      hRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0078D4" } };

      if (prevTrip) {
        var ignRow = [fmtTime(block.rows[0].trip.start), "(Ignition On)", "", "", "", "—", ""];
        for (var ei = 0; ei < extraHeaders.length; ei++) ignRow.push("");
        var iRow = ws.addRow(ignRow);
        iRow.font = { color: { argb: "FFD97706" } };
      }

      block.rows.forEach(function (r) {
        var t         = r.trip;
        var driveMins = parseDurationToMins(t.drivingDuration);
        if (driveMins == null) driveMins = durationMins(t.start, t.stop);
        var idleMins  = parseDurationToMins(t.idlingDuration) || 0;
        var stopMins  = parseDurationToMins(t.stopDuration);
        var row = [
          fmtTime(t.start),
          startAddressForRow(r, addrMap),
          fmtActivityDistance(t.distance) + " | " + fmtDurPrecise(driveMins),
          addressForPoint(t.stopPoint, addrMap),
          fmtTime(t.stop),
          fmtDurPrecise(idleMins),
          stopMins != null ? fmtDurPrecise(stopMins) : ""
        ];
        ws.addRow(row.concat(activityEngineColValues(t, cols)));
      });

      var totalRow = ws.addRow([
        block.date + " Total",
        "",
        fmtActivityDistance(stats.distKm) + " in " + fmtDurWhole(stats.driveMins),
        "",
        stats.stops + " stops",
        fmtDurWhole(stats.idleMins),
        fmtDurWhole(stats.stopMins)
      ].concat(activityEngineDayTotals(block.rows, cols)));
      totalRow.font = { bold: true };
      totalRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF5F5F5" } };

      ws.addRow([]);
    });

    ws.addRow([]);
  });

  // ── Write and trigger download ─────────────────────────────────────────────
  wb.xlsx.writeBuffer().then(function (buffer) {
    var blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    var url  = URL.createObjectURL(blob);
    var a    = document.createElement("a");
    a.href = url; a.download = "activity_report_" + from + "_to_" + to + ".xlsx"; a.click();
    URL.revokeObjectURL(url);
  });
}

// ─── Generic helpers ───────────────────────────────────────────────────────
function apiCall(method, params, onSuccess, onError) {
  state.api.call(method, params, onSuccess, onError || function (e) { console.error(method, e); });
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
function esc(str)             { return String(str).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"); }
function showError(msg)       { document.getElementById("loading").innerHTML = "<p style='color:var(--critical);padding:40px'>" + msg + "</p>"; }

})();
