"use strict";

// Shared by standalone Legacy and the parent app. Deliberately page-memory only:
// a refresh/close resets choices, while Run and report-type changes do not.
function legacyFilters() {
  if (!state.legacyFilters) state.legacyFilters = {
    mode: "assets", groups: [], assets: { all: true, ids: {} }, drivers: { all: true, ids: {} },
    devices: [], users: [], groupMap: {}, ready: false, loading: false, waiters: []
  };
  return state.legacyFilters;
}
function cancelLegacyReport() {
  // Run is shared by every report type, including guarded non-Legacy requests.
  document.getElementById("run-report").disabled = false;
  document.getElementById("run-report").textContent = "Run";
  state.reportEpoch = (state.reportEpoch || 0) + 1;
  state.legacyLoading = false;
}
function clearLegacyResults() {
  state.legacyFullByDevice = null;
  state.legacyByDevice = null;
  state.legacyAddrMap = null;
  state.legacyStopInfo = null;
  state.legacyReportRange = null;
  ["export-csv", "export-pdf-legacy"].forEach(function (id) { document.getElementById(id).classList.add("hidden"); });
}
function resetLegacyPage() {
  cancelLegacyReport();
  state.initializationEpoch = (state.initializationEpoch || 0) + 1;
  state.legacyFilters = null; // initialize is also the database/principal boundary
  clearLegacyResults();
  state.reportRaw = []; state.reportData = []; state.reportDriverMap = {}; state.deviceMap = {};
  state.server = ""; state.dbName = ""; state.hostVia = "";
  state.unitSystem = "Imperial"; state.fuelEconomyUnit = "MPGImperial";
  document.getElementById("db-name").textContent = "";
  document.getElementById("report-output").innerHTML = "<p class='report-placeholder'>Select a date range and click Run.</p>";
  document.getElementById("report-summary").innerHTML = "";
  document.getElementById("report-summary").classList.add("hidden");
  document.getElementById("main").classList.add("hidden");
  document.getElementById("loading").classList.remove("hidden");
  ["legacy-group-options", "legacy-entity-options"].forEach(function (id) {
    var element = document.getElementById(id); if (element) element.innerHTML = "";
  });
  ["legacy-group-search", "legacy-entity-search"].forEach(function (id) {
    var element = document.getElementById(id); if (element) element.value = "";
  });
  ["legacy-group-picker", "legacy-entity-picker"].forEach(function (id) {
    var element = document.getElementById(id); if (element) element.open = false;
  });
  var mode = document.getElementById("legacy-mode"); if (mode) mode.value = "assets";
  var groupSummary = document.getElementById("legacy-group-summary"); if (groupSummary) groupSummary.textContent = "All groups";
  var entitySummary = document.getElementById("legacy-entity-summary"); if (entitySummary) entitySummary.textContent = "All assets";
  var all = document.getElementById("legacy-select-all");
  if (all) { all.checked = false; all.indeterminate = false; all.disabled = true; }
  ["export-format", "export-activity-btn", "report-vehicle"].forEach(function (id) {
    var element = document.getElementById(id); if (element) element.classList.add("hidden");
  });
}
// Used only by the parent's other report runners, not dashboard/API consumers.
function legacyGuardedReportCall(method, params, success, failure) {
  var epoch = state.reportEpoch;
  apiCall(method, params, function (value) {
    if (epoch === state.reportEpoch) success(value);
  }, function (err) { if (epoch === state.reportEpoch && failure) failure(err); });
}
function legacyIsActive() {
  var type = document.getElementById("report-type");
  return !type || type.value === "legacy-trip-history";
}
function setupLegacyFilters() {
  var wrap = document.getElementById("legacy-filters");
  if (!wrap) return;
  if (wrap._legacyBound) {
    document.getElementById("legacy-mode").value = legacyFilters().mode;
    loadLegacyCatalog(function () {}); return;
  }
  wrap._legacyBound = true;
  wrap.innerHTML = "<details id='legacy-group-picker' class='legacy-picker'><summary id='legacy-group-summary'>All groups</summary>" +
    "<div class='legacy-picker-panel'><label>Find groups<input id='legacy-group-search' type='search' class='select'></label>" +
    "<label class='legacy-check'><input id='legacy-all-groups' type='checkbox' checked> All groups</label>" +
    "<div id='legacy-group-options' class='legacy-options'></div></div></details>" +
    "<label class='legacy-mode-label'>Report by<select id='legacy-mode' class='select'><option value='assets'>Assets</option><option value='drivers'>Drivers</option></select></label>" +
    "<details id='legacy-entity-picker' class='legacy-picker'><summary id='legacy-entity-summary'>All assets</summary>" +
    "<div class='legacy-picker-panel'><label><span id='legacy-search-label'>Find assets</span><input id='legacy-entity-search' type='search' class='select'></label>" +
    "<label class='legacy-check'><input id='legacy-select-all' type='checkbox'> Select all filtered</label>" +
    "<div id='legacy-entity-options' class='legacy-options'></div><p id='legacy-selection-count' role='status'></p></div></details>" +
    "<span id='legacy-filter-status' role='status'></span><button id='legacy-filter-retry' type='button' class='btn btn-secondary hidden'>Retry filters</button>";
  document.getElementById("legacy-mode").value = legacyFilters().mode;
  document.getElementById("legacy-mode").addEventListener("change", function () {
    legacyFilters().mode = this.value;
    document.getElementById("legacy-entity-search").value = "";
    renderLegacyChoices(); applyLegacyFilters();
  });
  document.getElementById("legacy-all-groups").addEventListener("change", function () {
    legacyFilters().groups = []; renderLegacyGroups(); renderLegacyChoices(); applyLegacyFilters();
  });
  document.getElementById("legacy-group-options").addEventListener("change", function (event) {
    var id = event.target.getAttribute("data-group"), f = legacyFilters();
    if (!id) return;
    f.groups = f.groups.filter(function (value) { return value !== id; });
    if (event.target.checked) f.groups.push(id);
    document.getElementById("legacy-all-groups").checked = !f.groups.length;
    updateLegacyGroupSummary(); renderLegacyChoices(); applyLegacyFilters();
  });
  document.getElementById("legacy-entity-options").addEventListener("change", function (event) {
    var id = event.target.getAttribute("data-entity");
    if (!id) return;
    setLegacySelected(id, event.target.checked); updateLegacySelection(); applyLegacyFilters();
  });
  document.getElementById("legacy-select-all").addEventListener("change", function () {
    var checked = this.checked;
    legacyVisibleEntities().forEach(function (entity) { setLegacySelected(entity.id, checked); });
    renderLegacyChoices(); applyLegacyFilters();
  });
  document.getElementById("legacy-group-search").addEventListener("input", renderLegacyGroups);
  document.getElementById("legacy-entity-search").addEventListener("input", renderLegacyChoices);
  document.getElementById("legacy-filter-retry").addEventListener("click", function () { loadLegacyCatalog(function () {}); });
  ["legacy-group-picker", "legacy-entity-picker"].forEach(function (id) {
    document.getElementById(id).addEventListener("keydown", function (event) {
      if (event.key === "Escape") { this.open = false; this.querySelector("summary").focus(); }
    });
  });
  loadLegacyCatalog(function () {});
}
function loadLegacyCatalog(done, failure, refresh) {
  var f = legacyFilters();
  if (f.ready && !refresh) { done(); return; }
  f.waiters.push({ done: done, failure: failure });
  if (f.loading) return;
  f.loading = true; f.ready = false;
  document.getElementById("legacy-filter-status").textContent = "Loading filters...";
  document.getElementById("legacy-filter-retry").classList.add("hidden");
  var pending = 3, records = {}, settled = false;
  var timer = setTimeout(function () { fail(new Error("Filter lookup timed out.")); }, 20000);
  function current() { return state.legacyFilters === f && !settled; }
  function fail(err) {
    if (!current()) return;
    settled = true; clearTimeout(timer); f.loading = false;
    document.getElementById("legacy-filter-status").textContent = "Filters unavailable. Retry before running the report.";
    document.getElementById("legacy-filter-retry").classList.remove("hidden");
    var waiters = f.waiters.splice(0);
    waiters.forEach(function (waiter) { if (waiter.failure) waiter.failure(err); });
  }
  ["Device", "User", "Group"].forEach(function (type) {
    if (!current()) return;
    apiCall("Get", { typeName: type, search: {}, resultsLimit: 5000 }, function (items) {
      if (!current()) return;
      if (records[type]) return;
      if (!Array.isArray(items) || items.length >= 5000) { fail(new Error("Filter catalog is incomplete.")); return; }
      records[type] = items;
      if (--pending) return;
      settled = true; clearTimeout(timer); f.loading = false; f.ready = true;
      f.devices = records.Device;
      f.users = records.User;
      f.groupMap = {};
      records.Group.forEach(function (group) { f.groupMap[group.id] = group; });
      document.getElementById("legacy-filter-status").textContent = "Groups include subgroups; driver groups are home groups.";
      renderLegacyGroups(); renderLegacyChoices();
      f.waiters.splice(0).forEach(function (waiter) { waiter.done(); });
    }, fail);
  });
}
function legacyRefId(ref) { return typeof ref === "string" ? ref : ref && ref.id; }
function legacyDescendants(ids) {
  var map = legacyFilters().groupMap, result = {};
  function visit(id) {
    if (!id || result[id]) return;
    result[id] = true;
    ((map[id] && map[id].children) || []).forEach(function (child) { visit(legacyRefId(child)); });
  }
  ids.forEach(visit); return result;
}
function legacyCompanyGroups() {
  var f = legacyFilters(), seeds = ["GroupCompanyId"];
  f.devices.concat(f.users).forEach(function (entity) {
    (entity.groups || entity.driverGroups || []).forEach(function (ref) { seeds.push(legacyRefId(ref)); });
  });
  // Restricted users may receive a partial tree without Company. Retain visible
  // ancestors of actual memberships, never security/report/data-access groups.
  var allowed = legacyDescendants(seeds), changed = true;
  while (changed) {
    changed = false;
    Object.keys(f.groupMap).forEach(function (id) {
      if (!allowed[id] && (f.groupMap[id].children || []).some(function (child) { return allowed[legacyRefId(child)]; })) {
        allowed[id] = true; changed = true;
      }
    });
  }
  return Object.keys(f.groupMap).filter(function (id) { return allowed[id]; }).map(function (id) { return f.groupMap[id]; });
}
function legacyInGroups(entity, mode) {
  var f = legacyFilters();
  if (!f.groups.length) return true;
  var allowed = legacyDescendants(f.groups);
  return ((entity && entity[mode === "drivers" ? "driverGroups" : "groups"]) || []).some(function (ref) { return allowed[legacyRefId(ref)]; });
}
function legacyEntities() {
  var f = legacyFilters(), entities = [], seen = {}, historical = {};
  if (state.legacyFullByDevice) Object.keys(state.legacyFullByDevice).forEach(function (id) {
    state.legacyFullByDevice[id].trips.forEach(function (trip) { historical[reportDriverInfo(trip).key] = reportDriverInfo(trip); });
  });
  var source = f.mode === "assets" ? f.devices : f.users;
  source.forEach(function (entity) {
    var id = f.mode === "assets" ? entity.id : "id:" + entity.id;
    if (f.mode === "drivers" && !entity.isDriver && !(entity.driverGroups || []).length && !historical[id]) return;
    if (!legacyInGroups(entity, f.mode)) return;
    seen[id] = true;
    entities.push({ id: id, name: f.mode === "assets" ? entity.name || entity.id : reportUserName(entity) || "Driver name unavailable (" + entity.id + ")" });
  });
  if (f.mode === "drivers" && !f.groups.length) {
    historical.unassigned = { key: "unassigned", name: "Unassigned / unknown driver" };
    Object.keys(historical).forEach(function (key) { if (!seen[key]) entities.push({ id: key, name: historical[key].name }); });
  }
  var counts = {};
  entities.forEach(function (entity) { counts[entity.name] = (counts[entity.name] || 0) + 1; });
  entities.forEach(function (entity) { if (counts[entity.name] > 1) entity.name += " (" + entity.id.replace(/^id:/, "") + ")"; });
  return entities.sort(function (a, b) { return a.name.localeCompare(b.name) || a.id.localeCompare(b.id); });
}
function legacyVisibleEntities() {
  var query = document.getElementById("legacy-entity-search").value.toLocaleLowerCase();
  return legacyEntities().filter(function (entity) { return entity.name.toLocaleLowerCase().indexOf(query) !== -1; });
}
function legacySelected(id) {
  var f = legacyFilters(), selection = f[f.mode];
  return selection.all ? !selection.ids[id] : !!selection.ids[id];
}
function setLegacySelected(id, checked) {
  var f = legacyFilters(), selection = f[f.mode];
  if (checked !== selection.all) selection.ids[id] = true;
  else delete selection.ids[id];
}
function updateLegacyGroupSummary() {
  var f = legacyFilters();
  document.getElementById("legacy-group-summary").textContent = f.groups.length ? f.groups.length + " group(s)" : "All groups";
}
function renderLegacyGroups() {
  var f = legacyFilters(), query = document.getElementById("legacy-group-search").value.toLocaleLowerCase();
  document.getElementById("legacy-all-groups").checked = !f.groups.length;
  document.getElementById("legacy-group-options").innerHTML = legacyCompanyGroups().sort(function (a, b) { return a.name.localeCompare(b.name); }).filter(function (group) {
    return group.name.toLocaleLowerCase().indexOf(query) !== -1;
  }).map(function (group) {
    return "<label class='legacy-check'><input type='checkbox' data-group='" + esc(group.id) + "'" + (f.groups.indexOf(group.id) !== -1 ? " checked" : "") + "> " + esc(group.name) + "</label>";
  }).join("") || "<p>No matching groups.</p>";
  updateLegacyGroupSummary();
}
function renderLegacyChoices() {
  var f = legacyFilters();
  document.getElementById("legacy-search-label").textContent = "Find " + f.mode;
  document.getElementById("legacy-entity-search").setAttribute("aria-label", "Find " + f.mode);
  document.getElementById("legacy-entity-options").innerHTML = legacyVisibleEntities().map(function (entity) {
    return "<label class='legacy-check'><input type='checkbox' data-entity='" + esc(entity.id) + "'" + (legacySelected(entity.id) ? " checked" : "") + "> " + esc(entity.name) + "</label>";
  }).join("") || "<p>No matching " + f.mode + ".</p>";
  updateLegacySelection();
}
function updateLegacySelection() {
  var f = legacyFilters(), eligible = legacyEntities(), visible = legacyVisibleEntities();
  var count = eligible.filter(function (entity) { return legacySelected(entity.id); }).length;
  var selectedVisible = visible.filter(function (entity) { return legacySelected(entity.id); }).length;
  var all = document.getElementById("legacy-select-all");
  all.checked = visible.length > 0 && selectedVisible === visible.length;
  all.indeterminate = selectedVisible > 0 && selectedVisible < visible.length;
  all.disabled = !f.ready || !visible.length;
  document.getElementById("legacy-entity-summary").textContent = count === eligible.length && count ? "All " + f.mode + " (" + count + ")" : count + " " + f.mode + " selected";
  document.getElementById("legacy-selection-count").textContent = count + " of " + eligible.length + " " + f.mode + " selected";
}
function applyLegacyFilters() {
  if (!legacyIsActive() || state.legacyLoading || !state.legacyFullByDevice) return;
  var f = legacyFilters(), allowed = {}, byDevice = {}, trips = [];
  legacyEntities().forEach(function (entity) { if (legacySelected(entity.id)) allowed[entity.id] = true; });
  Object.keys(state.legacyFullByDevice).forEach(function (vid) {
    var full = state.legacyFullByDevice[vid];
    if (f.mode === "assets" && !allowed[vid]) return;
    var drivers = full.driverGroups.filter(function (group) { return f.mode === "assets" || allowed[group.key]; });
    var selectedTrips = full.trips.filter(function (trip) { return f.mode === "assets" || allowed[reportDriverInfo(trip).key]; });
    if (!selectedTrips.length) return;
    byDevice[vid] = { name: full.name, trips: selectedTrips, driverGroups: drivers };
    trips = trips.concat(selectedTrips);
  });
  state.reportData = trips; state.legacyByDevice = byDevice;
  updateLegacyReportSummary(trips);
  renderLegacyTripHistoryOutput(byDevice, state.legacyAddrMap || {});
  ["export-csv", "export-pdf-legacy"].forEach(function (id) { document.getElementById(id).classList.toggle("hidden", !trips.length); });
}
function runLegacyTripHistoryReport(from, to, done) {
  cancelLegacyReport(); clearLegacyResults();
  state.reportRaw = []; state.reportData = []; state.legacyLoading = true;
  document.getElementById("run-report").disabled = true;
  document.getElementById("run-report").textContent = "Loading...";
  var epoch = state.reportEpoch;
  var finished = false;
  function current() { return !finished && epoch === state.reportEpoch && legacyIsActive(); }
  function complete() {
    finished = true; state.legacyLoading = false;
    document.getElementById("run-report").disabled = false;
    document.getElementById("run-report").textContent = "Run";
    done();
  }
  function fail(err) { if (current()) { reportError(err); complete(); } }
  function call(method, params, success, failure) {
    if (!current()) return;
    var settled = false;
    var timer = setTimeout(function () { settle(new Error("Report lookup timed out.")); }, 20000);
    function settle(err, value) {
      if (settled) return;
      settled = true; clearTimeout(timer);
      if (!current()) return;
      if (err) (failure || fail)(err); else success(value);
    }
    try { apiCall(method, params, function (value) { settle(null, value); }, function (err) { settle(err); }); }
    catch (err) { settle(err); }
  }
  var fromDate = new Date(from + "T00:00:00").toISOString();
  var toDate = new Date(to + "T23:59:59").toISOString();
  loadLegacyCatalog(function () {
    if (!current()) return;
    var f = legacyFilters(), deviceMap = {};
    f.devices.forEach(function (device) { deviceMap[device.id] = device.name; });
    call("Get", { typeName: "Trip", search: { fromDate: fromDate, toDate: toDate } }, function (trips) {
      if (!current()) return;
      var unique = {}, users = {}, offset = 0;
      trips.forEach(function (trip) { var id = reportDriverId(trip); if (id && !isUnassignedDriver(id)) unique[id] = true; });
      var ids = Object.keys(unique);
      document.getElementById("report-output").innerHTML = "<p class='placeholder'>Resolving driver names...</p>";
      function nextDrivers() {
        if (!current()) return;
        if (offset < ids.length) {
          var batch = ids.slice(offset, offset + 100); offset += batch.length;
          call("Get", { typeName: "User", search: { userIds: batch } }, function (records) {
            if (!current()) return;
            (records || []).forEach(function (user) { if (batch.indexOf(user.id) !== -1) users[user.id] = user; });
            nextDrivers();
          }, nextDrivers);
          return;
        }
        // Assembly is local to this run. No global driver/stop map is mutated
        // until the guarded final callback publishes the completed snapshot.
        var points = {}, addrMap = {}, keys, position = 0;
        trips.forEach(function (trip) { if (trip.stopPoint && trip.stopPoint.x != null && trip.stopPoint.y != null) points[coordKey(trip.stopPoint)] = trip.stopPoint; });
        keys = Object.keys(points);
        document.getElementById("report-output").innerHTML = "<p class='placeholder'>Resolving stop addresses...</p>";
        function nextAddresses() {
          if (!current()) return;
          if (position < keys.length) {
            var chunk = keys.slice(position, position + 400); position += chunk.length;
            call("GetAddresses", { coordinates: chunk.map(function (key) { return points[key]; }) }, function (addresses) {
              if (!current()) return;
              (addresses || []).forEach(function (address, i) { if (chunk[i]) addrMap[chunk[i]] = formatReverseGeocodeAddress(address); });
              nextAddresses();
            }, nextAddresses);
            return;
          }
          state.reportDriverMap = {};
          Object.keys(users).forEach(function (id) {
            var user = users[id], name = reportUserName(user);
            if (name) state.reportDriverMap[id] = { name: name, login: user.name || "" };
            var existing = f.users.findIndex(function (candidate) { return candidate.id === id; });
            if (existing === -1) f.users.push(user); else f.users[existing] = user;
          });
          state.deviceMap = deviceMap; state.reportRaw = trips;
          state.legacyFullByDevice = buildLegacyByDevice(trips, deviceMap);
          state.legacyAddrMap = addrMap; state.legacyReportRange = { from: from, to: to };
          state.legacyLoading = false;
          renderLegacyChoices(); applyLegacyFilters(); complete();
        }
        nextAddresses();
      }
      nextDrivers();
    }, fail);
  }, fail, true); // refresh memberships/catalogs, never the user's selection IDs
}
