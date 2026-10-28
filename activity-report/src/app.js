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
// vice versa) until the two are formally split into separate repos. Activity's
// vehicle/day builder is separate from the parent's driver-based Legacy report.
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
  reportDriverMap:  {}, // Users resolved from the historical Trip.driver references
  dbName:           "",
  activityByDevice: null,        // { deviceId: { name, trips, days } }
  activityAddrMap:  null,        // "lat,lng" -> resolved address string
  unitSystem:       "Imperial",  // from the logged-in User.isMetric ("Metric" or "Imperial")
  legacyLengthEl:   null
};

// Settings are stored under MyGeotab's own origin, which every add-in on that
// database shares, so the key is namespaced away from the parent add-in's.
var SETTINGS_KEY = "activity_report_standalone_settings";

// ─── Historical trip drivers (shared Legacy helpers; unused by Activity) ────
// Trip.driver identifies the driver of THAT trip. Never substitute the device's
// current driver: a vehicle can change hands repeatedly within the report range.
function reportDriverId(trip) {
  var driver = trip && trip.driver;
  return typeof driver === "string" ? driver : (driver && driver.id) || "";
}
function isUnassignedDriver(id) {
  return /^(?:UnknownDriver(?:Id)?|NoDriver(?:Id)?)$/i.test(id || "");
}
function reportUserName(user) {
  if (!user) return "";
  return [user.firstName || "", user.lastName || ""].join(" ").trim() || (user.name || "").trim();
}
function reportDriverInfo(trip) {
  var id = reportDriverId(trip);
  if (isUnassignedDriver(id)) return { key: "unassigned", id: "", name: "Unassigned / unknown driver", login: "" };
  var user = id && state.reportDriverMap[id];
  var name = (user && user.name) || reportUserName(trip && trip.driver) || ((trip && trip.driverName) || "").trim();
  if (id) return { key: "id:" + id, id: id, name: name || "Driver name unavailable (" + id + ")", login: (user && user.login) || "" };
  return name ? { key: "name:" + name, id: "", name: name, login: "" }
    : { key: "unassigned", id: "", name: "Unassigned / unknown driver", login: "" };
}
function resolveReportDrivers(trips, done) {
  state.reportDriverMap = {};
  var unique = {};
  trips.forEach(function (trip) {
    var id = reportDriverId(trip);
    if (id && !isUnassignedDriver(id)) unique[id] = true;
  });
  var ids = Object.keys(unique), offset = 0;
  function nextBatch() {
    if (offset >= ids.length) { done(); return; }
    var batch = ids.slice(offset, offset + 100);
    offset += batch.length;
    // Explicit ids include historical/archived users. A current-driver-only or
    // active-date search would lose drivers who worked earlier in the period.
    apiCall("Get", { typeName: "User", search: { userIds: batch } }, function (users) {
      (users || []).forEach(function (user) {
        if (batch.indexOf(user.id) === -1) return;
        var name = reportUserName(user);
        if (name) state.reportDriverMap[user.id] = { name: name, login: user.name || "" };
      });
      nextBatch();
    }, nextBatch); // keep trips visible; unresolved ids get an explicit label
  }
  nextBatch();
}
function groupReportDriverDays(dayBlocks) {
  var groups = {}, order = [];
  dayBlocks.forEach(function (block) {
    block.rows.forEach(function (row) {
      var info = reportDriverInfo(row.trip);
      if (!groups[info.key]) {
        groups[info.key] = { key: info.key, id: info.id, name: info.name, login: info.login, days: [], byDay: {} };
        order.push(info.key);
      }
      var group = groups[info.key];
      if (!group.byDay[block.date]) {
        group.byDay[block.date] = { date: block.date, rows: [] };
        group.days.push(group.byDay[block.date]);
      }
      // Preserve the predecessor from the WHOLE vehicle timeline, including a
      // different driver's preceding trip. It defines the real start location.
      group.byDay[block.date].rows.push(row);
    });
  });
  var names = {};
  order.forEach(function (key) { var name = groups[key].name; names[name] = (names[name] || 0) + 1; });
  var result = order.map(function (key) {
    var group = groups[key];
    if (names[group.name] > 1) group.name += " (" + (group.login || group.id) + ")";
    return group;
  });
  result.sort(function (a, b) { return a.name.localeCompare(b.name) || a.key.localeCompare(b.key); });
  return result;
}

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
  week.setDate(today.getDate() - 7);
  document.getElementById("filter-from").value = fmtDateInput(week);
  document.getElementById("filter-to").value   = fmtDateInput(today);

  document.getElementById("run-report").addEventListener("click", runReport);
  document.getElementById("range-length").addEventListener("change", applyRangeLength);
  document.getElementById("export-activity-btn").addEventListener("click", function () {
    var fmt = document.getElementById("export-format").value;
    if      (fmt === "csv")   exportActivityCsv();
    else if (fmt === "excel") exportActivityExcel();
    else                      exportActivityPdf();
  });
  applyRangeLength();
}

// Recompute filter-from from filter-to based on the selected preset length.
// "custom" leaves both date inputs alone so the user can pick any range manually.
function applyRangeLength() {
  var days = document.getElementById("range-length").value;
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

// Group the complete vehicle timeline by stable Device.id and local start day.
//
// NOTE: unlike the Legacy Trip History report, the Activity Report deliberately
// does not suppress overnight parks — it reports raw Trip.stopDuration, so the
// last stop of a day carries the full time until the next morning's first trip.
function buildActivityByDevice(trips, deviceMap) {
  var byDevice = {};
  trips.forEach(function (t) {
    var vid = t.device && t.device.id; if (!vid) return;
    if (!byDevice[vid]) byDevice[vid] = { name: deviceMap[vid] || vid, trips: [], days: [] };
    byDevice[vid].trips.push(t);
  });
  Object.keys(byDevice).forEach(function (vid) {
    var vehicle = byDevice[vid], days = {};
    vehicle.trips.sort(function (a, b) { return new Date(a.start) - new Date(b.start); });
    vehicle.trips.forEach(function (trip, i) {
      var date = localDayKey(trip.start);
      if (!days[date]) { days[date] = { date: date, rows: [] }; vehicle.days.push(days[date]); }
      days[date].rows.push({ trip: trip, prevTrip: i ? vehicle.trips[i - 1] : null });
    });
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

// ─── Column selection ──────────────────────────────────────────────────────
function fmtActivityDistance(km) { return fmtDist(km, 2); }

function fmtActivityDistanceKpi(km) {
  return { value: distVal(km).toFixed(0), unit: distUnit() };
}

function getActivityCols() {
  var el = document.getElementById("col-work-split");
  return el && el.checked ? ["work-split"] : [];
}

// ─── Run ───────────────────────────────────────────────────────────────────
function runActivityReport(from, to, done) {
  state.activityByDevice = null;
  state.activityAddrMap = null;
  state.activityRange = null;
  state.reportRaw = [];
  state.reportData = [];
  document.getElementById("export-format").classList.add("hidden");
  document.getElementById("export-activity-btn").classList.add("hidden");
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
      var byDevice = buildActivityByDevice(trips, deviceMap);
      document.getElementById("report-output").innerHTML = "<p class='placeholder'>Resolving stop addresses…</p>";
      resolveStopAddresses(byDevice, function (addrMap) {
        state.reportRaw = trips;
        state.reportData = trips;
        state.activityByDevice = byDevice;
        state.activityAddrMap = addrMap;
        state.activityRange = { from: from, to: to };
        populateReportVehicleFilter(trips, deviceMap);
        renderActivityReport(byDevice, addrMap);
        document.getElementById("export-format").classList.remove("hidden");
        document.getElementById("export-activity-btn").classList.remove("hidden");
        saveSettings(from, to);
        done();
      });
    }, function (err) { reportError(err); done(); });
  }, function (err) { reportError(err); done(); });
}

// ─── Render (in-app preview) ───────────────────────────────────────────────
function activityVehicleIds(byDevice) {
  return Object.keys(byDevice).sort(function (a, b) {
    return byDevice[a].name.localeCompare(byDevice[b].name) || a.localeCompare(b);
  });
}

function activityReportRange() {
  return state.activityRange || { from: "", to: "" };
}

function activityRangeLabel(from, to) {
  function dateLabel(day) { return fmtDateReadable(day + "T12:00:00"); }
  return from === to ? dateLabel(from) : dateLabel(from) + " – " + dateLabel(to);
}

function activityTotals(trips) {
  var totals = { distance: 0, driving: 0, idle: 0, stopped: 0, engine: 0 };
  trips.forEach(function (t) {
    totals.distance += t.distance || 0;
    var mins = parseDurationToMins(t.drivingDuration);
    totals.driving += mins != null ? mins : durationMins(t.start, t.stop);
    totals.idle += parseDurationToMins(t.idlingDuration) || 0;
    totals.stopped += parseDurationToMins(t.stopDuration) || 0;
  });
  totals.engine = totals.driving + totals.idle;
  return totals;
}

function fmtActivityTime(iso) {
  if (!iso) return "--";
  var date = new Date(iso);
  return String(date.getHours()).padStart(2, "0") + ":" + String(date.getMinutes()).padStart(2, "0");
}

function activityLastTrip(block) {
  return block.rows.reduce(function (last, row) {
    return new Date(row.trip.stop) > new Date(last.stop) ? row.trip : last;
  }, block.rows[0].trip);
}

function activityDayEndpoints(block, addrMap) {
  var first = block.rows[0], last = activityLastTrip(block);
  return {
    startDateTime: first.trip.start, endDateTime: last.stop,
    startTime: fmtActivityTime(first.trip.start), start: startAddressForRow(first, addrMap),
    endTime: fmtActivityTime(last.stop), end: addressForPoint(last.stopPoint, addrMap)
  };
}

function activityDayElapsed(block) {
  return Math.max(0, durationMins(block.rows[0].trip.start, activityLastTrip(block).stop));
}

function activitySummaryHeaders(cols) {
  return ["Name", "Start Date", "Start Time", "Start Location", "End Time", "End Location", "Total Hours (h)",
    "Total Stop Duration (h)", "Total Idle Duration (h)", "Total Driving Duration (h)", "Engine On (Total)", "Distance (" + distUnit() + ")"]
    .concat(cols.indexOf("work-split") < 0 ? [] : ["Work Engine (h)", "After-hours Engine (h)"]);
}

function activitySummaryNumbers(trips, elapsed, cols) {
  var totals = activityTotals(trips);
  var values = [elapsed / 60, totals.stopped / 60, totals.idle / 60, totals.driving / 60, totals.engine / 1440, distVal(totals.distance)];
  if (cols.indexOf("work-split") >= 0) {
    ["work", "afterHours"].forEach(function (prefix) {
      values.push(trips.reduce(function (sum, trip) {
        return sum + (parseDurationToMins(trip[prefix + "DrivingDuration"]) || 0) + (parseDurationToMins(trip[prefix + "IdlingDuration"]) || 0);
      }, 0) / 60);
    });
  }
  return values;
}

function activityNumberText(values) {
  return values.map(function (value, i) {
    if (i !== 4) return value.toFixed(2);
    var minutes = Math.round(value * 1440);
    return Math.floor(minutes / 60) + ":" + String(minutes % 60).padStart(2, "0");
  });
}

function activityDayValues(block, addrMap, cols) {
  var endpoints = activityDayEndpoints(block, addrMap);
  var trips = block.rows.map(function (row) { return row.trip; });
  return [block.date.split("-").reverse().join("/"), endpoints.startTime, endpoints.start, endpoints.endTime, endpoints.end]
    .concat(activityNumberText(activitySummaryNumbers(trips, activityDayElapsed(block), cols)));
}

function activityVehicleNumbers(vehicle, cols) {
  var elapsed = vehicle.days.reduce(function (sum, day) { return sum + activityDayElapsed(day); }, 0);
  return activitySummaryNumbers(vehicle.trips, elapsed, cols);
}

function renderActivityReport(byDevice, addrMap) {
  var summary = document.getElementById("report-summary");
  if (summary) { summary.innerHTML = ""; summary.classList.add("hidden"); }
  var cols   = getActivityCols();
  var output = document.getElementById("report-output");
  var vids   = activityVehicleIds(byDevice);
  if (!vids.length) { output.innerHTML = "<p class='placeholder'>No trips found.</p>"; return; }
  var range = activityReportRange();
  var period = "Selected period: " + activityRangeLabel(range.from, range.to);
  var summaryHeaders = activitySummaryHeaders(cols).slice(1);

  output.innerHTML = vids.map(function (vid) {
    var v = byDevice[vid];
    var dayHtml = v.days.map(function (block) {
      return "<div class='activity-day-row activity-summary-grid' data-day='" + block.date + "'>" + activityDayValues(block, addrMap, cols).map(function (value, i) {
          return "<span class='" + (i >= 5 ? "activity-number" : "activity-text") + "'><span class='activity-sr-only'>" + esc(summaryHeaders[i]) + ": </span>" + esc(value) + "</span>";
        }).join("") + "</div>";
    }).join("");
    var footer = "<div class='activity-vehicle-total activity-summary-grid'><span class='activity-total-label'>" + esc(v.name) + " Total</span>"
      + activityNumberText(activityVehicleNumbers(v, cols)).map(function (value) { return "<strong class='activity-number'>" + value + "</strong>"; }).join("") + "</div>";
    return "<section class='activity-vehicle-block'><h3>" + esc(v.name) + "</h3><p class='activity-period'>" + esc(period) + "</p>"
      + "<div class='activity-summary-table" + (cols.indexOf("work-split") >= 0 ? " activity-with-work" : "") + "'>"
      + "<div class='activity-summary-head activity-summary-grid' aria-hidden='true'>" + summaryHeaders.map(function (header) { return "<span>" + esc(header) + "</span>"; }).join("") + "</div>"
      + dayHtml + footer + "</div></section>";
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
    if (!raw) return;
    var s = JSON.parse(raw);
    if (s.len)  { var lenEl = document.getElementById("range-length"); if (lenEl) lenEl.value = s.len; }
    if (s.from) { var fromEl = document.getElementById("filter-from"); if (fromEl) fromEl.value = s.from; }
    if (s.to)   { var toEl   = document.getElementById("filter-to");   if (toEl)   toEl.value   = s.to;   }
    var el = document.getElementById("col-work-split");
    if (el) el.checked = (s.cols || []).indexOf("work-split") !== -1;
  } catch (e) {}
}

// ─── Report helpers ────────────────────────────────────────────────────────
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
  var filteredByDevice = buildActivityByDevice(state.reportData, state.deviceMap);
  state.activityByDevice = filteredByDevice;
  renderActivityReport(filteredByDevice, state.activityAddrMap || {});
}

// ─── CSV export ────────────────────────────────────────────────────────────
function exportActivityCsv() {
  var data    = state.reportData;
  var addrMap = state.activityAddrMap || {};
  var cols    = getActivityCols();
  if (!data || !data.length) return;
  var range = activityReportRange(), from = range.from, to = range.to;
  var rows = [activitySummaryHeaders(cols)];
  var byDevice = state.activityByDevice;
  activityVehicleIds(byDevice).forEach(function (vid) {
    var v = byDevice[vid];
    v.days.forEach(function (day) { rows.push([v.name].concat(activityDayValues(day, addrMap, cols))); });
    rows.push([v.name + " Total", "", "", "", "", ""].concat(activityNumberText(activityVehicleNumbers(v, cols))));
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
  var range = activityReportRange(), from = range.from, to = range.to;

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

  var headers = [activitySummaryHeaders(cols).slice(1)];
  var pdfColumnStyles = { 0: { cellWidth: 18 }, 1: { cellWidth: 14 }, 2: { cellWidth: 38 }, 3: { cellWidth: 14 }, 4: { cellWidth: 38 } };

  activityVehicleIds(byDevice).forEach(function (vid) {
    var v = byDevice[vid];
    if (yPos > pageH - 40) { doc.addPage(); yPos = 16; }
    doc.setTextColor(20, 20, 20); doc.setFontSize(12); doc.setFont("helvetica", "bold");
    doc.text(v.name, margin, yPos);
    yPos += 6;
      var body = v.days.map(function (day) { return activityDayValues(day, addrMap, cols); });
      var totalStyles = { fontStyle: "bold", fillColor: [232, 244, 253] };
      body.push([{ content: v.name + " Total", colSpan: 5, styles: totalStyles }].concat(
        activityNumberText(activityVehicleNumbers(v, cols)).map(function (value) { return { content: value, styles: totalStyles }; })
      ));

      doc.autoTable({
        startY: yPos,
        head: [[{ content: v.name, colSpan: headers[0].length }]].concat(headers),
        body: body,
        margin: { left: margin, right: margin },
        columnStyles: pdfColumnStyles,
        styles: { fontSize: 7.5, cellPadding: 2, textColor: [45, 55, 72], overflow: "linebreak" },
        headStyles: { fillColor: [0, 120, 212], textColor: 255, fontStyle: "bold", fontSize: 7.5 }
      });
      yPos = doc.lastAutoTable.finalY + 6;
      if (yPos > pageH - 30) { doc.addPage(); yPos = 16; }
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
function activityWallDate(iso, dateOnly) {
  var date = new Date(iso);
  return new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate(),
    dateOnly ? 0 : date.getHours(), dateOnly ? 0 : date.getMinutes(), dateOnly ? 0 : date.getSeconds(), dateOnly ? 0 : date.getMilliseconds()));
}

function addActivityGroupedWorksheet(wb, byDevice, addrMap, cols) {
  var headers = activitySummaryHeaders(cols);
  var utcStart = headers.length + 1, utcEnd = utcStart + 1;
  var ws = wb.addWorksheet("Activity Report");
  wb.views = [{ activeTab: wb.worksheets.indexOf(ws) }];
  ws.properties.outlineLevelRow = 1;
  ws.properties.outlineProperties = { summaryBelow: true, summaryRight: false };
  ws.views = [{ state: "frozen", ySplit: 1, showOutlineSymbols: true }];
  ws.columns = [36, 14, 12, 48, 12, 48, 17, 22, 22, 22, 20, 18].concat(headers.slice(12).map(function () { return 24; }))
    .map(function (width) { return { width: width }; }).concat([{ hidden: true }, { hidden: true }]);
  var header = ws.addRow(headers.concat(["Day Start (UTC)", "Day End (UTC)"]));
  header.font = { name: "Arial", bold: true, color: { argb: "FFFFFFFF" } };
  header.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0078D4" } };
  header.alignment = { vertical: "middle", wrapText: true };
  header.height = 30;
  var startLetter = header.getCell(utcStart).address.replace(/[0-9]/g, "");
  var endLetter = header.getCell(utcEnd).address.replace(/[0-9]/g, "");

  function addRow(values, level, total) {
    var row = ws.addRow(values);
    row.outlineLevel = level;
    row.hidden = false;
    // Daily rows have no child detail. Keep vehicle groups expanded initially.
    Object.defineProperty(row, "collapsed", { value: false, configurable: true });
    row.font = { name: "Arial", size: 11, bold: total };
    if (total) row.fill = { type: "pattern", pattern: "solid", fgColor: { argb: level === 0 ? "FFE8F4FD" : "FFF5F5F5" } };
    row.getCell(2).numFmt = "dd/mm/yyyy";
    row.getCell(3).numFmt = "hh:mm";
    row.getCell(5).numFmt = "hh:mm";
    for (var c = 7; c <= headers.length; c++) {
      row.getCell(c).numFmt = c === 11 ? "[h]:mm" : "0.00";
      if (level === 1) row.getCell(c).font = { name: "Arial", size: 11, bold: true };
    }
    return row;
  }
  function vehicleTotal(label, first, totals) {
    var last = ws.rowCount;
    var row = addRow([label], 0, true);
    totals.forEach(function (value, i) {
      var cell = row.getCell(i + 7);
      var letter = cell.address.replace(/[0-9]/g, "");
      cell.value = { formula: "SUM(" + letter + first + ":" + letter + last + ")", result: value };
    });
  }
  activityVehicleIds(byDevice).forEach(function (vid) {
    var vehicle = byDevice[vid], vehicleFirst = ws.rowCount + 1;
    vehicle.days.forEach(function (day) {
      var trips = day.rows.map(function (row) { return row.trip; });
      var values = activitySummaryNumbers(trips, activityDayElapsed(day), cols);
      var endpoints = activityDayEndpoints(day, addrMap);
      var row = addRow([vehicle.name, activityWallDate(endpoints.startDateTime, true), activityWallDate(endpoints.startDateTime), endpoints.start,
        activityWallDate(endpoints.endDateTime), endpoints.end].concat(values, [new Date(endpoints.startDateTime), new Date(endpoints.endDateTime)]), 1, false);
      row.getCell(7).value = { formula: "24*(" + endLetter + row.number + "-" + startLetter + row.number + ")", result: values[0] };
      row.getCell(11).value = { formula: "(I" + row.number + "+J" + row.number + ")/24", result: values[4] };
    });
    vehicleTotal(vehicle.name + " Total", vehicleFirst, activityVehicleNumbers(vehicle, cols));
  });
  return ws;
}

function exportActivityExcel() {
  var byDevice = state.activityByDevice;
  var addrMap  = state.activityAddrMap || {};
  var cols     = getActivityCols();
  if (!byDevice || !Object.keys(byDevice).length) { alert("Run the report first."); return; }
  if (typeof ExcelJS === "undefined") { alert("Excel library not loaded. Please check your internet connection."); return; }
  var range = activityReportRange(), from = range.from, to = range.to;
  var dUnit = distUnit();
  var vids  = activityVehicleIds(byDevice);

  // ── Build per-vehicle KPI data ────────────────────────────────────────────
  var allTrips    = [];
  var vehicleKpis = []; // [{name, distKm, engineMins}]
  vids.forEach(function (vid) {
    var v = byDevice[vid];
    allTrips = allTrips.concat(v.trips);
    var distKm    = v.trips.reduce(function (s, t) { return s + (t.distance || 0); }, 0);
    var driveMins = v.trips.reduce(function (s, t) { var dm = parseDurationToMins(t.drivingDuration); return s + (dm != null ? dm : durationMins(t.start, t.stop)); }, 0);
    var idleMins  = v.trips.reduce(function (s, t) { return s + (parseDurationToMins(t.idlingDuration) || 0); }, 0);
    vehicleKpis.push({ name: v.name, distKm: distKm, engineMins: driveMins + idleMins });
  });
  var totalDistKm = allTrips.reduce(function (s, t) { return s + (t.distance || 0); }, 0);
  var totalDrive  = allTrips.reduce(function (s, t) { var dm = parseDurationToMins(t.drivingDuration); return s + (dm != null ? dm : durationMins(t.start, t.stop)); }, 0);
  var totalIdle   = allTrips.reduce(function (s, t) { return s + (parseDurationToMins(t.idlingDuration) || 0); }, 0);
  var totalStop   = allTrips.reduce(function (s, t) { return s + (parseDurationToMins(t.stopDuration) || 0); }, 0);
  var distKpi     = fmtActivityDistanceKpi(totalDistKm);

  // ── Render grouped bar chart to an offscreen canvas ───────────────────────
  var chartCanvas = document.createElement("canvas");
  chartCanvas.width  = 900;
  chartCanvas.height = 420;

  var distValues   = vehicleKpis.map(function (k) { return parseFloat(distVal(k.distKm).toFixed(2)); });
  var engineValues = vehicleKpis.map(function (k) { return parseFloat((k.engineMins / 60).toFixed(2)); }); // hours

  var chartInst = new Chart(chartCanvas.getContext("2d"), {
    type: "bar",
    data: {
      labels: vehicleKpis.map(function (k) { return k.name; }),
      datasets: [
        { label: "Distance (" + dUnit + ")", data: distValues,   backgroundColor: "#0078D4", borderRadius: 4 },
        { label: "Engine On (hrs)",          data: engineValues, backgroundColor: "#059669", borderRadius: 4 }
      ]
    },
    options: {
      animation: false,
      responsive: false,
      plugins: {
        legend: { display: true, position: "top", labels: { font: { size: 13 }, color: "#374151", padding: 16 } },
        title: {
          display: true,
          text: "Vehicle Summary — " + from + " to " + to,
          font: { size: 15, weight: "bold" },
          color: "#111827",
          padding: { bottom: 16 }
        }
      },
      scales: {
        x: { ticks: { color: "#374151", font: { size: 11 } }, grid: { color: "rgba(0,0,0,.06)" } },
        y: { beginAtZero: true, ticks: { color: "#374151", font: { size: 11 } }, grid: { color: "rgba(0,0,0,.06)" } }
      }
    }
  });
  var chartPng = chartCanvas.toDataURL("image/png");
  chartInst.destroy();
  // Strip the data:image/png;base64, prefix — ExcelJS wants raw base64
  var chartBase64 = chartPng.replace(/^data:image\/png;base64,/, "");

  // ── Build workbook with ExcelJS ───────────────────────────────────────────
  var wb = new ExcelJS.Workbook();
  wb.creator = "Activity Report";
  wb.created = new Date();

  // ── Summary sheet — KPIs left (cols A-B), chart right (cols D-K) ──────────
  var summarySheet = wb.addWorksheet("Summary");
  summarySheet.columns = [
    { width: 36 }, { width: 24 }, { width: 4 },  // A, B, C (spacer)
    { width: 18 }, { width: 18 }, { width: 18 }, { width: 18 },
    { width: 18 }, { width: 18 }, { width: 18 }   // D-J (chart area)
  ];

  var titleRow = summarySheet.addRow(["Activity Report — " + from + " to " + to]);
  titleRow.font = { bold: true, size: 13 };
  summarySheet.addRow([]);

  var hdrRow = summarySheet.addRow(["Metric", "Value"]);
  hdrRow.font = { bold: true };
  hdrRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE8F4FD" } };

  [
    ["Total Distance (" + dUnit + ")", distKpi.value + " " + distKpi.unit],
    ["Total Travel Duration", fmtDurWhole(totalDrive)],
    ["Total Idle Time", fmtDurWhole(totalIdle)],
    ["Total Stop Duration", fmtDurWhole(totalStop)],
    ["Number of Stops", allTrips.length],
    ["Average Stop Duration", fmtDurWhole(allTrips.length ? totalStop / allTrips.length : 0)],
    ["Vehicles in Report", vids.length]
  ].forEach(function (r) { summarySheet.addRow(r); });

  // Chart image — placed to the right of the KPI table (col D = index 3, row 0)
  var imgId = wb.addImage({ base64: chartBase64, extension: "png" });
  summarySheet.addImage(imgId, { tl: { col: 3, row: 0 }, br: { col: 10, row: 20 } });

  addActivityGroupedWorksheet(wb, byDevice, addrMap, cols);

  // ── Write and trigger download ─────────────────────────────────────────────
  return wb.xlsx.writeBuffer().then(function (buffer) {
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
function fmtDateInput(d)      { return localDayKey(d); }
function fmtDateShort(iso)    { return iso ? new Date(iso).toISOString().slice(0, 10) : ""; }
function fmtDateReadable(iso) { if (!iso) return ""; var d = new Date(iso); return d.toLocaleDateString(undefined, { day: "2-digit", month: "short", year: "numeric" }); }
function fmtTime(iso)         { return iso ? new Date(iso).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }) : "--"; }
function durationMins(a, b)   { return (!a || !b) ? 0 : (new Date(b) - new Date(a)) / 60000; }
function esc(str)             { return String(str).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"); }
function showError(msg)       { document.getElementById("loading").innerHTML = "<p style='color:var(--critical);padding:40px'>" + msg + "</p>"; }

})();
