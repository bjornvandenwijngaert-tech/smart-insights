"use strict";

// Optional zone membership uses Geotab's own containment calculation. Coordinate
// keys are exact: the address cache's ~1m rounding is unsafe at polygon borders.
function legacyZoneKey(point) {
  return legacyCoordinate(point) ? JSON.stringify([point.x, point.y]) : "";
}
function legacyZonesEnabled() { return !!legacyFilters().zones; }
function legacyHeaders() {
  var headers = ["Start Time", "Distance", "Duration", "Stop Location", "Arrival Time", "Idle Duration", "Stop Duration", "Trip History"];
  if (legacyZonesEnabled()) headers.splice(4, 0, "Zone");
  return headers;
}
function legacyPdfColumns() {
  var widths = legacyZonesEnabled() ? [16, 16, 22, 38, 28, 16, 22, 22, 20] : [22, 20, 25, 41, 22, 24, 24, 22];
  var styles = {};
  widths.forEach(function (width, index) { styles[index] = { cellWidth: width }; });
  return styles; // both layouts occupy 200mm with 5mm table side margins on A4
}
function legacyZoneLabel(trip) {
  var key = legacyZoneKey(trip.stopPoint), snapshot = state.legacyZoneSnapshot;
  if (!key) return "(location unknown)";
  if (!snapshot || snapshot.status === "loading") return "Loading...";
  if (snapshot.status === "error") return snapshot.message || "Zone lookup failed";
  return snapshot.labels[key] || "";
}
function legacyZoneExportReady() {
  if (legacyZonesEnabled() && (!state.legacyZoneSnapshot || state.legacyZoneSnapshot.status === "loading")) {
    alert("Wait for zones to finish loading, or turn Zones off."); return false;
  }
  return true;
}
function updateLegacyZoneControls() {
  var enabled = legacyZonesEnabled(), snapshot = state.legacyZoneSnapshot;
  var pending = enabled && snapshot && snapshot.status === "loading";
  ["export-csv", "export-pdf-legacy"].forEach(function (id) {
    var button = document.getElementById(id);
    if (button) button.disabled = !!(legacyIsActive() && pending);
  });
  var toggle = document.getElementById("legacy-zones");
  if (toggle) toggle.checked = enabled;
  var status = document.getElementById("legacy-zone-status");
  if (status) status.textContent = !enabled || !snapshot ? "" : pending ? "Loading zones..." : snapshot.status === "error" ? snapshot.message + ". Retry zones." : "";
  var retry = document.getElementById("legacy-zone-retry");
  if (retry) retry.classList.toggle("hidden", !(enabled && snapshot && snapshot.status === "error"));
  legacyPreferenceStatus();
}
function cancelLegacyZones() {
  var lookup = state.legacyZoneLookup;
  state.legacyZoneLookup = null;
  if (lookup) lookup.timers.forEach(clearTimeout);
  var interrupted = state.legacyZoneSnapshot && state.legacyZoneSnapshot.status === "loading";
  if (interrupted) state.legacyZoneSnapshot = { status: "error", labels: {}, message: "Zone lookup interrupted" };
  // This button is shared with other parent reports. Cancellation always releases it.
  ["export-csv", "export-pdf-legacy"].forEach(function (id) {
    var button = document.getElementById(id); if (button) button.disabled = false;
  });
  return !!interrupted;
}
function setupLegacyZones() {
  var toggle = document.getElementById("legacy-zones");
  if (!toggle || toggle._zonesBound) { updateLegacyZoneControls(); return; }
  toggle._zonesBound = true;
  toggle.addEventListener("change", function () {
    var pref = legacyPreferenceOwner();
    if (state.legacyIdentityStale || state.legacyFocusChecking || (pref && (pref.status === "loading" || pref.status === "saving"))) {
      updateLegacyZoneControls(); return;
    }
    legacyFilters().zones = this.checked;
    if (!this.checked) cancelLegacyZones();
    applyLegacyFilters(); updateLegacyZoneControls();
    legacyPreferenceSave();
  });
  document.getElementById("legacy-zone-retry").addEventListener("click", function () {
    if (legacyZonesEnabled() && state.legacyFullByDevice && !state.legacyLoading) loadLegacyZones();
  });
  updateLegacyZoneControls();
}
function loadLegacyZones() {
  cancelLegacyZones();
  var source = state.legacyFullByDevice, epoch = state.reportEpoch, api = state.api, database = state.dbName;
  if (!source || !legacyZonesEnabled() || state.legacyLoading || !legacyIsActive()) return;
  var lookup = { timers: [], nextRequestAt: 0 };
  state.legacyZoneLookup = lookup;
  state.legacyZoneSnapshot = { status: "loading", labels: {} };
  function current() {
    return state.legacyZoneLookup === lookup && state.reportEpoch === epoch && state.api === api && state.dbName === database
      && state.legacyFullByDevice === source && legacyZonesEnabled() && legacyIsActive();
  }
  function finish(labels, error) {
    if (!current()) return;
    lookup.timers.forEach(clearTimeout); state.legacyZoneLookup = null;
    state.legacyZoneSnapshot = { status: error ? "error" : "ready", labels: labels, message: error ? "Zone lookup failed" : "" };
    applyLegacyFilters();
  }
  function fail() { finish({}, true); }
  function call(method, params, success) {
    if (!current()) return;
    var delay = lookup.nextRequestAt - Date.now();
    if (delay > 0) { lookup.timers.push(setTimeout(function () { call(method, params, success); }, delay)); return; }
    lookup.nextRequestAt = Date.now() + 150;
    var settled = false, timer = setTimeout(function () { settle(null, true); }, 20000);
    lookup.timers.push(timer);
    function settle(value, error) {
      if (settled) return;
      settled = true; clearTimeout(timer);
      if (!current()) return;
      if (error) fail(); else success(value);
    }
    try { api.call(method, params, function (value) { settle(value, false); }, function () { settle(null, true); }); }
    catch (error) { if (settled) throw error; settle(null, true); }
  }
  var points = Object.create(null), labels = Object.create(null);
  Object.keys(source).forEach(function (id) {
    source[id].trips.forEach(function (trip) {
      var key = legacyZoneKey(trip.stopPoint);
      if (key) points[key] = { x: trip.stopPoint.x, y: trip.stopPoint.y };
    });
  });
  var keys = Object.keys(points), offset = 0;
  applyLegacyFilters(); // render loading cells and disable exports for this report
  if (!keys.length) { finish(labels, false); return; }
  call("Get", { typeName: "Zone", search: { fromDate: new Date().toISOString() }, resultsLimit: 10000 }, function (zones) {
    if (!Array.isArray(zones) || zones.length >= 10000) { fail(); return; }
    var names = Object.create(null);
    zones.forEach(function (zone) { if (zone && zone.id) names[zone.id] = typeof zone.name === "string" ? zone.name.trim() : ""; });
    if (!zones.length) { finish(labels, false); return; }
    function next() {
      if (!current()) return;
      if (offset >= keys.length) { finish(labels, false); return; }
      var batch = keys.slice(offset, offset + 400); offset += batch.length;
      call("GetAddresses", { coordinates: batch.map(function (key) { return points[key]; }) }, function (addresses) {
        if (!Array.isArray(addresses) || addresses.length !== batch.length) { fail(); return; }
        for (var i = 0; i < batch.length; i++) {
          var address = addresses[i], zonesHere = address && address.zones;
          if (!address || typeof address !== "object" || Array.isArray(address) || (zonesHere != null && !Array.isArray(zonesHere))) { fail(); return; }
          var seen = Object.create(null), matches = [];
          for (var j = 0; j < (zonesHere || []).length; j++) {
            var zone = zonesHere[j], id = typeof zone === "string" ? zone : zone && zone.id;
            if (typeof id !== "string" || !id) { fail(); return; }
            if (seen[id]) continue;
            seen[id] = true;
            matches.push({ id: id, name: names[id] || "Zone name unavailable" });
          }
          matches.sort(function (a, b) { return a.name.localeCompare(b.name) || a.id.localeCompare(b.id); });
          labels[batch[i]] = matches.map(function (zone) { return zone.name; }).join("; ");
        }
        next();
      });
    }
    next();
  });
}
