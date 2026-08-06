"use strict";

// ─── Preset widget type registry ──────────────────────────────────────────────
var PRESET_WIDGET_TYPES = {
  "carbon-monthly": {
    label:       "Monthly Carbon Report",
    description: "Monthly CO\u2082 emissions and fuel efficiency across your fleet",
    color:       "#059669",
    chartType:   "bar",
    statLabels:  ["Total CO\u2082", "Fleet MPG", "Best MPG"]
  },
  "speeding": {
    label:       "Speeding",
    description: "Trips exceeding a configurable speed threshold (default 80\u00a0mph)",
    color:       "#DC2626",
    chartType:   "bar",
    statLabels:  ["Total", "Per Day", "Worst Day"],
    hasParams:   true,
    defaultParams: { thresholdMph: 80 }
  },
  // NOTE: "maintenance-upcoming" and "maintenance-spend" removed from the picker.
  // MyGeotab has no classic SDK entity for work orders/maintenance; the next-gen
  // Work Order Management app (solutionId geotabWorkOrderManagement) stores its
  // data in a service not reachable via the classic api.call("Get", ...) path a
  // classic add-in uses. Handler functions are retained below (dead code) so these
  // can be restored quickly if/when Geotab ships a public Maintenance Center API.
  "fuel-economy-daily": {
    label:       "Daily Fuel Economy",
    description: "Daily fleet average MPG over the last 30 days",
    color:       "#059669",
    chartType:   "line",
    statLabels:  ["Fleet MPG", "Best Day", "Worst Day"]
  }
};

// ─── State ─────────────────────────────────────────────────────────────────
var state = {
  api:              null,
  reportData:       [],
  reportRaw:        [],
  widgets:          [],
  suggestions:      [],
  deviceMap:        {},
  driverMap:        {},
  dbName:           "",
  editMode:         false,
  drilldown:        { widgetType: null, widgetDef: null, data: [], filtered: [], meta: {} },
  nextX:            0,
  nextY:            0,
  dashboardDataId:  null,  // AddInData record id; null until first load or save
  legacyByDevice:   null,  // Legacy Trip History Report: { deviceId: { name, trips, drivers } }
  legacyAddrMap:    null,  // Legacy Trip History Report: "lat,lng" -> resolved address string
  unitSystem:       "Imperial",  // loaded from SystemSettings on init ("Metric" or "Imperial")
  activityByDevice: null,  // Activity Report: { deviceId: { name, trips, drivers } }
  activityAddrMap:  null   // Activity Report: "lat,lng" -> resolved address string
};
// ─── MyGeotab addin lifecycle ──────────────────────────────────────────────
if (typeof geotab === "undefined") { var geotab = { addin: {} }; }

geotab.addin.smartInsights = function () {
  return {
    initialize: function (api, freshState, callback) {
      try {
        state.api = api;
        if (freshState && freshState.database) {
          state.dbName = freshState.database;
          document.getElementById("db-name").textContent = freshState.database;
        }
        // Load measurement system from database settings so distance renders in the
        // correct unit without requiring the user to manually toggle.
        apiCall("Get", { typeName: "SystemSettings" }, function (settings) {
          if (settings && settings[0] && settings[0].measurementSystem) {
            state.unitSystem = settings[0].measurementSystem; // "Metric" or "Imperial"
          }
        }, function () {}); // silent fail — stays Imperial
        setupNav();
        setupReports();
        setupMyReports();
        restoreActivitySettings(); // pre-fill Activity Report toolbar from last session
        setupEditMode();
        setupWidgetPicker();
        setupModal();
        setupParamEditor();
        setupMap();
        setupIncident();
        // Keep the loading screen until AddInData has been fetched so the
        // dashboard doesn't flash empty before widgets appear.
        restoreDashboard(function () {
          document.getElementById("loading").classList.add("hidden");
          document.getElementById("main").classList.remove("hidden");
          loadSuggestions();
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

// ─── Navigation ─────────────────────────────────────────────────────────────
function setupNav() {
  document.querySelectorAll(".nav-btn").forEach(function (btn) {
    btn.addEventListener("click", function () {
      document.querySelectorAll(".nav-btn").forEach(function (b) { b.classList.remove("active"); });
      document.querySelectorAll(".tab").forEach(function (t) { t.classList.remove("active"); t.classList.add("hidden"); });
      btn.classList.add("active");
      var target = document.getElementById("tab-" + btn.dataset.tab);
      target.classList.remove("hidden");
      target.classList.add("active");
    });
  });
}

// ─── Edit mode ───────────────────────────────────────────────────────────────
function setupEditMode() {
  var editBtn = document.getElementById("edit-layout-btn");
  var doneBtn = document.getElementById("done-layout-btn");
  var grid    = document.getElementById("dashboard-grid");

  editBtn.addEventListener("click", function () {
    state.editMode = true;
    grid.classList.add("edit-mode");
    document.body.classList.add("edit-mode-active");
    editBtn.classList.add("hidden");
    doneBtn.classList.remove("hidden");
    document.getElementById("dashboard-hint").textContent = "Drag widgets to reposition. Pull the corner to resize.";
  });

  doneBtn.addEventListener("click", function () {
    state.editMode = false;
    grid.classList.remove("edit-mode");
    document.body.classList.remove("edit-mode-active");
    doneBtn.classList.add("hidden");
    editBtn.classList.remove("hidden");
    document.getElementById("dashboard-hint").textContent = "Click a widget to view details. Use Edit Layout to rearrange.";
    saveDashboard();
  });
}
// ─── Widget picker ────────────────────────────────────────────────────────────
function setupWidgetPicker() {
  document.getElementById("add-widget-btn").addEventListener("click", openWidgetPicker);
  document.getElementById("picker-close").addEventListener("click", closeWidgetPicker);
  document.getElementById("widget-picker").addEventListener("click", function (e) {
    if (e.target === this) closeWidgetPicker();
  });

  var grid = document.getElementById("picker-grid");
  var icons = {
    "carbon-monthly":       '<path d="M12 2a10 10 0 1 0 0 20 10 10 0 0 0 0-20z"/><path d="M8 14s1.5 2 4 2 4-2 4-2"/><path d="M9 9h.01"/><path d="M15 9h.01"/>',
    "speeding":          '<circle cx="12" cy="13" r="2"/><path d="M12 3a9 9 0 0 1 6.4 15.4"/><path d="M12 7v2"/><path d="M6.4 6.4 7.8 7.8"/><path d="M3 13h2"/>',
    "maintenance-upcoming": '<path d="M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z"/>',
    "maintenance-spend":    '<line x1="12" y1="1" x2="12" y2="23"/><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/>',
    "fuel-economy-daily":   '<path d="M3 22V10a2 2 0 0 1 2-2h8a2 2 0 0 1 2 2v12"/><path d="M15 14h1a2 2 0 0 1 2 2v1a2 2 0 0 0 2 2 2 2 0 0 0-2-2v-1a4 4 0 0 0-4-4"/><path d="M3 22h12"/><path d="M7 8V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v4"/>'
  };

  grid.innerHTML = Object.keys(PRESET_WIDGET_TYPES).map(function (key) {
    var t = PRESET_WIDGET_TYPES[key];
    return "<div class='picker-card' data-type='" + key + "'>" +
      "<div class='picker-icon' style='color:" + t.color + ";background:" + hexToRgba(t.color, 0.08) + "'>" +
        "<svg viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='1.8' stroke-linecap='round' stroke-linejoin='round'>" +
          (icons[key] || "") +
        "</svg>" +
      "</div>" +
      "<div class='picker-info'>" +
        "<div class='picker-name'>" + t.label + "</div>" +
        "<div class='picker-desc'>" + t.description + "</div>" +
      "</div>" +
      "<button class='btn btn-primary btn-sm'>Add</button>" +
    "</div>";
  }).join("");

  grid.querySelectorAll(".picker-card").forEach(function (card) {
    card.querySelector("button").addEventListener("click", function () {
      addPresetWidget(card.dataset.type);
    });
  });
}

function openWidgetPicker() {
  document.getElementById("widget-picker").classList.remove("hidden");
}

function closeWidgetPicker() {
  document.getElementById("widget-picker").classList.add("hidden");
}

function addPresetWidget(typeKey) {
  var id      = "widget-" + Date.now();
  var typeDef = PRESET_WIDGET_TYPES[typeKey];
  var widgetDef = { type: typeKey, params: typeDef.defaultParams ? JSON.parse(JSON.stringify(typeDef.defaultParams)) : {} };
  closeWidgetPicker();
  switchToTab("dashboard");
  addWidget(id, typeDef.label, widgetDef, state.nextX, state.nextY, DEFAULT_W, DEFAULT_H);
  state.nextX = (state.nextX + CASCADE) % 120;
  state.nextY = (state.nextY + CASCADE) % 120;
  saveDashboard();
}

function switchToTab(tabName) {
  document.querySelectorAll(".nav-btn").forEach(function (b) { b.classList.remove("active"); });
  document.querySelectorAll(".tab").forEach(function (t) { t.classList.remove("active"); t.classList.add("hidden"); });
  var btn = document.querySelector("[data-tab='" + tabName + "']");
  if (btn) btn.classList.add("active");
  var tab = document.getElementById("tab-" + tabName);
  if (tab) { tab.classList.remove("hidden"); tab.classList.add("active"); }
}
// ─── My Reports history ───────────────────────────────────────────────────────
var REPORT_LABELS = {
  "trips":               "Daily Trips",
  "carbon-monthly":      "Monthly Carbon",
  "speeding":            "Speeding >80mph",
  "fuel-economy-daily":  "Daily Fuel Economy",
  "legacy-trip-history": "Legacy Trip History",
  "activity-report":     "Activity Report"
};

function saveReportHistory(entry) {
  // entry: { type, label, from, to, preset, cols, speedThresh }
  try {
    var hist = getReportHistory();
    // Deduplicate — remove an identical prior entry to keep list tidy
    hist = hist.filter(function (h) {
      return !(h.type === entry.type && h.from === entry.from && h.to === entry.to && h.preset === entry.preset);
    });
    hist.unshift(entry);
    if (hist.length > 10) hist = hist.slice(0, 10);
    localStorage.setItem("si_report_history", JSON.stringify(hist));
    renderMyReports();
  } catch (e) {}
}

function getReportHistory() {
  try { return JSON.parse(localStorage.getItem("si_report_history") || "[]"); } catch (e) { return []; }
}

function clearReportHistory() {
  try { localStorage.removeItem("si_report_history"); } catch (e) {}
  renderMyReports();
}

function renderMyReports() {
  var hist = getReportHistory();
  var section = document.getElementById("my-reports-section");
  var list    = document.getElementById("my-reports-list");
  if (!section || !list) return;
  if (!hist.length) { section.classList.add("hidden"); return; }
  section.classList.remove("hidden");

  list.innerHTML = hist.map(function (h, i) {
    var isCustom = (h.preset === "custom" || !h.preset);
    var metaStr  = isCustom
      ? h.from + " – " + h.to
      : (h.preset === "1" ? "Daily" : h.preset === "7" ? "Weekly" : "Monthly") + " (last run: " + h.from + " – " + h.to + ")";
    var badge    = isCustom
      ? "<span class='my-report-badge custom'>Custom range</span>"
      : "<span class='my-report-badge'>" + (h.preset === "1" ? "Daily" : h.preset === "7" ? "Weekly" : "Monthly") + "</span>";
    return "<div class='my-report-item'>" +
      "<div class='my-report-icon'>" +
        "<svg viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='1.8' stroke-linecap='round' stroke-linejoin='round'>" +
          "<path d='M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z'/><polyline points='14 2 14 8 20 8'/>" +
          "<line x1='8' y1='13' x2='16' y2='13'/><line x1='8' y1='17' x2='16' y2='17'/>" +
        "</svg>" +
      "</div>" +
      "<div class='my-report-info'>" +
        "<div class='my-report-name'>" + esc(h.label) + "</div>" +
        "<div class='my-report-meta'>" + esc(metaStr) + "</div>" +
      "</div>" +
      badge +
      "<button class='btn btn-sm btn-primary' data-hist-idx='" + i + "'>Re-run</button>" +
    "</div>";
  }).join("");

  list.querySelectorAll("[data-hist-idx]").forEach(function (btn) {
    btn.addEventListener("click", function () {
      openRerunModal(hist[parseInt(this.dataset.histIdx, 10)]);
    });
  });
}

function openRerunModal(h) {
  var isCustom = (h.preset === "custom" || !h.preset);
  var modal    = document.getElementById("rerun-modal");
  var title    = document.getElementById("rerun-title");
  var subtitle = document.getElementById("rerun-subtitle");
  var body     = document.getElementById("rerun-body");
  var period   = document.getElementById("rerun-period");
  var editArea = document.getElementById("rerun-edit-area");
  var editBtn  = document.getElementById("rerun-edit");
  var confirmBtn = document.getElementById("rerun-confirm");

  title.textContent    = "Re-run: " + h.label;
  subtitle.textContent = isCustom ? "Custom date range" : "Preset: " + (h.preset === "1" ? "Daily" : h.preset === "7" ? "Weekly" : "Monthly");

  // Reset edit state
  editArea.style.display = "none";
  editBtn.textContent = "Edit Period";

  var resolvedFrom, resolvedTo;

  if (isCustom) {
    body.textContent = "You’re about to run this report again for the exact same period:";
    period.textContent = h.from + " – " + h.to;
    resolvedFrom = h.from;
    resolvedTo   = h.to;
    document.getElementById("rerun-from").value = h.from;
    document.getElementById("rerun-to").value   = h.to;
  } else {
    // Re-resolve relative preset from today
    var today   = new Date();
    var toDate  = new Date(today);
    var fromDate= new Date(today);
    fromDate.setDate(today.getDate() - (parseInt(h.preset, 10) - 1));
    resolvedFrom = fmtDateInput(fromDate);
    resolvedTo   = fmtDateInput(toDate);
    body.textContent = "This will run “" + h.label + "” for the past " + (h.preset === "1" ? "day" : h.preset === "7" ? "7 days" : "30 days") + " relative to today.";
    period.textContent = resolvedFrom + " – " + resolvedTo;
    document.getElementById("rerun-from").value = resolvedFrom;
    document.getElementById("rerun-to").value   = resolvedTo;
    // Preset reports don't offer period editing
    editBtn.style.display = "none";
  }

  if (isCustom) {
    editBtn.style.display = "";
    var editing = false;
    editBtn.onclick = function () {
      editing = !editing;
      editArea.style.display = editing ? "" : "none";
      editBtn.textContent = editing ? "Use original period" : "Edit Period";
    };
  }

  confirmBtn.onclick = function () {
    var fromVal = isCustom && document.getElementById("rerun-edit-area").style.display !== "none"
      ? document.getElementById("rerun-from").value
      : resolvedFrom;
    var toVal = isCustom && document.getElementById("rerun-edit-area").style.display !== "none"
      ? document.getElementById("rerun-to").value
      : resolvedTo;
    closeRerunModal();
    applyAndRunReport(h, fromVal, toVal);
  };

  modal.classList.remove("hidden");
}

function closeRerunModal() {
  document.getElementById("rerun-modal").classList.add("hidden");
}

function applyAndRunReport(h, from, to) {
  // Switch to Reports tab
  switchToTab("reports");
  // Set report type
  var typeSel = document.getElementById("report-type");
  typeSel.value = h.type;
  typeSel.dispatchEvent(new Event("change"));
  // Set dates
  document.getElementById("filter-from").value = from;
  document.getElementById("filter-to").value   = to;
  // Set preset length if applicable
  if (h.preset && h.preset !== "custom") {
    var lenSel = document.getElementById("legacy-length");
    if (lenSel) lenSel.value = h.preset;
  } else if (h.preset === "custom") {
    var lenSel2 = document.getElementById("legacy-length");
    if (lenSel2) lenSel2.value = "custom";
  }
  // Restore activity columns if saved
  if (h.type === "activity-report" && h.cols) {
    ["total-engine", "drive-only", "idle-only", "work-split"].forEach(function (id) {
      var el = document.getElementById("col-" + id);
      if (el) el.checked = h.cols.indexOf(id) !== -1;
    });
  }
  // Restore speed threshold if saved
  if (h.type === "speeding" && h.speedThresh) {
    var thEl = document.getElementById("filter-speed-thresh");
    if (thEl) thEl.value = h.speedThresh;
  }
  // Run
  runReport();
}

function setupMyReports() {
  document.getElementById("my-reports-clear").addEventListener("click", function () {
    clearReportHistory();
  });
  document.getElementById("rerun-close").addEventListener("click", closeRerunModal);
  document.getElementById("rerun-cancel").addEventListener("click", closeRerunModal);
  document.getElementById("rerun-modal").addEventListener("click", function (e) {
    if (e.target === this) closeRerunModal();
  });
  renderMyReports();
}

// ─── Reports ────────────────────────────────────────────────────────────────
function setupReports() {
  var today = new Date();
  var week  = new Date(today);
  week.setDate(today.getDate() - 7);
  document.getElementById("filter-from").value = fmtDateInput(week);
  document.getElementById("filter-to").value   = fmtDateInput(today);
  document.getElementById("run-report").addEventListener("click", runReport);
  document.getElementById("export-csv").addEventListener("click", exportReportCsv);
  document.getElementById("export-pdf-legacy").addEventListener("click", exportLegacyTripHistoryPdf);
  document.getElementById("report-type").addEventListener("change", function () {
    document.getElementById("report-vehicle").classList.add("hidden");
    document.getElementById("export-csv").classList.add("hidden");
    document.getElementById("export-pdf-legacy").classList.add("hidden");
    document.getElementById("export-format").classList.add("hidden");
    document.getElementById("export-activity-btn").classList.add("hidden");
    document.getElementById("activity-cols-wrap").classList.add("hidden");
    document.getElementById("report-summary").classList.add("hidden");
    document.getElementById("report-output").innerHTML = "<p class='report-placeholder'>Select a date range and click Run.</p>";
    // Show/hide date inputs based on type
    var type = this.value;
    var noDate = (type === "maintenance-upcoming");
    document.getElementById("filter-from").style.display = noDate ? "none" : "";
    document.getElementById("filter-to").style.display   = noDate ? "none" : "";
    var thWrap = document.getElementById("speed-thresh-wrap");
    if (thWrap) thWrap.classList.toggle("hidden", type !== "speeding");
    var lenWrap = document.getElementById("legacy-length");
    if (lenWrap) {
      lenWrap.classList.toggle("hidden", type !== "legacy-trip-history" && type !== "activity-report");
      if (type === "legacy-trip-history" || type === "activity-report") applyLegacyLength();
    }
    if (type === "activity-report") {
      document.getElementById("activity-cols-wrap").classList.remove("hidden");
    }
  });
  document.getElementById("legacy-length").addEventListener("change", applyLegacyLength);
  document.getElementById("export-activity-btn").addEventListener("click", function () {
    var fmt = document.getElementById("export-format").value;
    if (fmt === "csv")        exportActivityCsv();
    else if (fmt === "excel") exportActivityExcel();
    else                      exportActivityPdf();
  });
}

// Recompute filter-from from filter-to based on the selected preset length.
// "custom" leaves both date inputs alone so the user can pick any range manually.
function applyLegacyLength() {
  var lenSel = document.getElementById("legacy-length");
  var days   = lenSel.value;
  if (days === "custom") return;
  var toEl = document.getElementById("filter-to");
  var to   = toEl.value ? new Date(toEl.value + "T00:00:00") : new Date();
  var from = new Date(to);
  from.setDate(to.getDate() - (parseInt(days, 10) - 1));
  document.getElementById("filter-from").value = fmtDateInput(from);
  document.getElementById("filter-to").value   = fmtDateInput(to);
}

function runReport() {
  var type   = document.getElementById("report-type").value;
  var from   = document.getElementById("filter-from").value;
  var to     = document.getElementById("filter-to").value;
  var btn    = document.getElementById("run-report");
  var output = document.getElementById("report-output");

  if (type !== "maintenance-upcoming" && (!from || !to)) { alert("Please select a date range."); return; }

  btn.disabled = true; btn.textContent = "Loading...";
  output.innerHTML = "<p class='placeholder'>Fetching data...</p>";
  document.getElementById("report-summary").classList.add("hidden");
  document.getElementById("export-csv").classList.add("hidden");
  document.getElementById("report-vehicle").classList.add("hidden");

  // Capture preset/custom state for history
  var lenSel    = document.getElementById("legacy-length");
  var preset    = (type === "legacy-trip-history" || type === "activity-report") && lenSel
    ? lenSel.value
    : (type === "speeding" || type === "trips" || type === "carbon-monthly" || type === "fuel-economy-daily" ? "custom" : null);
  var histEntry = {
    type:       type,
    label:      REPORT_LABELS[type] || type,
    from:       from,
    to:         to,
    preset:     preset,
    cols:       type === "activity-report" ? getActivityCols() : null,
    speedThresh:type === "speeding" ? (document.getElementById("filter-speed-thresh") || {}).value || "80" : null
  };

  function done(success) {
    btn.disabled = false; btn.textContent = "Run";
    if (success !== false) saveReportHistory(histEntry);
  }

  if      (type === "trips")                 runTripsReport(from, to, done);
  else if (type === "carbon-monthly")        runCarbonReport(from, to, done);
  else if (type === "speeding")           runSpeedingReport(from, to, done);
  else if (type === "maintenance-upcoming")  runMaintenanceUpcomingReport(done);
  else if (type === "maintenance-spend")     runMaintenanceSpendReport(from, to, done);
  else if (type === "fuel-economy-daily")    runFuelEconomyReport(from, to, done);
  else if (type === "legacy-trip-history")   runLegacyTripHistoryReport(from, to, done);
  else if (type === "activity-report")       runActivityReport(from, to, done);
}

// ── Trips report (existing) ───────────────────────────────────────────────────
function runTripsReport(from, to, done) {
  var fromDate = new Date(from + "T00:00:00").toISOString();
  var toDate   = new Date(to   + "T23:59:59").toISOString();
  apiCall("Get", { typeName: "Device", search: {} }, function (devices) {
    var deviceMap = {};
    devices.forEach(function (d) { deviceMap[d.id] = d.name; });
    apiCall("Get", { typeName: "Trip", search: { fromDate: fromDate, toDate: toDate } }, function (trips) {
      state.reportRaw = trips;
      state.reportData = trips;
      populateReportVehicleFilter(trips, deviceMap);
      var totalKm   = trips.reduce(function (s, t) { return s + (t.distance || 0); }, 0);
      var totalMins = trips.reduce(function (s, t) { return s + durationMins(t.start, t.stop); }, 0);
      var vehicles  = {};
      trips.forEach(function (t) { if (t.device) vehicles[t.device.id] = 1; });
      showReportSummary([
        summaryCard("Total Distance", totalKm.toFixed(1), "km"),
        summaryCard("Drive Time",     fmtMins(totalMins), ""),
        summaryCard("Trips",          trips.length, ""),
        summaryCard("Active Vehicles",Object.keys(vehicles).length, "")
      ]);
      renderTripsOutput(trips, deviceMap);
      document.getElementById("export-csv").classList.remove("hidden");
      done();
    }, function (err) { reportError(err); done(false); });
  }, function (err) { reportError(err); done(false); });
}

function renderTripsOutput(trips, deviceMap) {
  var output = document.getElementById("report-output");
  if (trips.length === 0) { output.innerHTML = "<p class='placeholder'>No trips found.</p>"; return; }
  var sorted = trips.slice().sort(function (a, b) { return new Date(b.start) - new Date(a.start); });
  output.innerHTML = sorted.map(function (t) {
    var vName = deviceMap ? (deviceMap[t.device && t.device.id] || "Unknown") : (state.deviceMap[t.device && t.device.id] || "Unknown");
    var dist  = (t.distance || 0).toFixed(1);
    var mins  = durationMins(t.start, t.stop);
    return "<div class='trip-card'>" +
      "<div class='trip-top'>" +
        "<div><div class='trip-vehicle'>" + esc(vName) + "</div>" +
        "<div class='trip-driver'>" + esc(t.driverName || "No driver assigned") + "</div></div>" +
        "<div class='trip-time'>" + fmtDateShort(t.start) + "<br>" + fmtTime(t.start) + " &rarr; " + fmtTime(t.stop) + "</div>" +
      "</div>" +
      "<div class='trip-meta'>" +
        tripStat("Distance", dist + " km") + tripStat("Duration", fmtMins(mins)) +
        tripStat("Max Speed", mphStr(t.maximumSpeed)) + tripStat("Avg Speed", mphStr(t.averageSpeed)) +
      "</div>" +
    "</div>";
  }).join("");
}

// ── Carbon report ─────────────────────────────────────────────────────────────
function runCarbonReport(from, to, done) {
  var fromDate = new Date(from + "T00:00:00").toISOString();
  var toDate   = new Date(to   + "T23:59:59").toISOString();
  apiCall("Get", { typeName: "Device", search: {} }, function (devices) {
    var deviceMap = {};
    devices.forEach(function (d) { deviceMap[d.id] = d.name; });
    state.deviceMap = deviceMap;
    apiCall("Get", { typeName: "Trip", search: { fromDate: fromDate, toDate: toDate } }, function (trips) {
      var fuelTrips = trips.filter(function (t) { return (t.fuelUsed || 0) > 0; });
      if (fuelTrips.length === 0) {
        document.getElementById("report-output").innerHTML = "<p class='placeholder'>No fuel data found. Fuel tracking must be enabled on your devices.</p>";
        done(); return;
      }
      // Aggregate per vehicle
      var byVehicle = {};
      fuelTrips.forEach(function (t) {
        var vid = t.device && t.device.id;
        if (!vid) return;
        if (!byVehicle[vid]) byVehicle[vid] = { name: deviceMap[vid] || vid, fuel: 0, distM: 0, co2: 0, trips: 0 };
        byVehicle[vid].fuel  += (t.fuelUsed || 0);
        byVehicle[vid].distM += (t.distance || 0);
        byVehicle[vid].co2   += (t.fuelUsed || 0) * 2.4;
        byVehicle[vid].trips += 1;
      });
      state.reportRaw  = Object.values(byVehicle);
      state.reportData = state.reportRaw;
      populateReportVehicleFilter(fuelTrips, deviceMap);

      var totalCo2  = fuelTrips.reduce(function (s, t) { return s + (t.fuelUsed || 0) * 2.4; }, 0);
      var totalFuel = fuelTrips.reduce(function (s, t) { return s + (t.fuelUsed || 0); }, 0);
      var totalDistM= fuelTrips.reduce(function (s, t) { return s + (t.distance || 0); }, 0);
      var fleetMpg  = calcMpg(totalDistM, totalFuel);
      showReportSummary([
        summaryCard("Total CO\u2082",    Math.round(totalCo2), "kg"),
        summaryCard("Fleet MPG",         fleetMpg, ""),
        summaryCard("Total Fuel",         totalFuel.toFixed(0), "L"),
        summaryCard("Vehicles Reporting", Object.keys(byVehicle).length, "")
      ]);
      renderCarbonOutput(Object.values(byVehicle));
      document.getElementById("export-csv").classList.remove("hidden");
      done();
    }, function (err) { reportError(err); done(false); });
  }, function (err) { reportError(err); done(false); });
}

function renderCarbonOutput(rows) {
  var output = document.getElementById("report-output");
  if (!rows.length) { output.innerHTML = "<p class='placeholder'>No data.</p>"; return; }
  var sorted = rows.slice().sort(function (a, b) { return b.co2 - a.co2; });
  output.innerHTML = "<div class='dd-table-wrap'><table class='dd-table'>" +
    "<thead><tr><th>Vehicle</th><th>CO\u2082 (kg)</th><th>MPG</th><th>Distance</th><th>Trips</th></tr></thead><tbody>" +
    sorted.map(function (r) {
      return "<tr><td class='td-vehicle'>" + esc(r.name) + "</td>" +
        "<td>" + Math.round(r.co2) + " kg</td>" +
        "<td>" + calcMpg(r.distM, r.fuel) + "</td>" +
        "<td>" + ((r.distM * 0.621371).toFixed(1)) + " mi</td>" +
        "<td>" + r.trips + "</td></tr>";
    }).join("") + "</tbody></table></div>";
}
// ── Speeding report (configurable threshold) ──────────────────────────────────
function runSpeedingReport(from, to, done) {
  var thEl    = document.getElementById("filter-speed-thresh");
  var thMph   = thEl ? (parseFloat(thEl.value) || 80) : 80;
  var thKmh   = thMph / 0.621371;
  var fromDate = new Date(from + "T00:00:00").toISOString();
  var toDate   = new Date(to   + "T23:59:59").toISOString();
  apiCall("Get", { typeName: "Device", search: {} }, function (devices) {
    var deviceMap = {};
    devices.forEach(function (d) { deviceMap[d.id] = d.name; });
    state.deviceMap = deviceMap;
    apiCall("Get", { typeName: "Trip", search: { fromDate: fromDate, toDate: toDate } }, function (trips) {
      var speeding = trips.filter(function (t) { return (t.maximumSpeed || 0) > thKmh; });
      state.reportRaw  = speeding;
      state.reportData = speeding;
      populateReportVehicleFilter(speeding, deviceMap);
      var vehicles = {};
      speeding.forEach(function (t) { if (t.device) vehicles[t.device.id] = 1; });
      var maxSpeedTrip = speeding.reduce(function (m, t) { return (t.maximumSpeed || 0) > (m.maximumSpeed || 0) ? t : m; }, {});
      showReportSummary([
        summaryCard("Incidents >" + thMph + "mph", speeding.length, ""),
        summaryCard("Vehicles",          Object.keys(vehicles).length, ""),
        summaryCard("Highest Speed",     maxSpeedTrip.maximumSpeed ? toMph(maxSpeedTrip.maximumSpeed).toFixed(0) : "—", "mph"),
        summaryCard("Period",            from + " \u2013 " + to, "")
      ]);
      renderSpeedingOutput(speeding, deviceMap);
      document.getElementById("export-csv").classList.remove("hidden");
      done();
    }, function (err) { reportError(err); done(false); });
  }, function (err) { reportError(err); done(false); });
}

function renderSpeedingOutput(trips, deviceMap) {
  var output = document.getElementById("report-output");
  if (!trips.length) { output.innerHTML = "<p class='placeholder'>No trips exceeding 80mph found for this period.</p>"; return; }
  var sorted = trips.slice().sort(function (a, b) { return (b.maximumSpeed || 0) - (a.maximumSpeed || 0); });
  output.innerHTML = "<div class='dd-table-wrap'><table class='dd-table'>" +
    "<thead><tr><th>Vehicle</th><th>Driver</th><th>Date</th><th>Start Time</th><th>Max Speed</th><th>Distance</th></tr></thead><tbody>" +
    sorted.map(function (t) {
      var dMap = deviceMap || state.deviceMap;
      var vName = dMap[t.device && t.device.id] || "Unknown";
      return "<tr><td class='td-vehicle'>" + esc(vName) + "</td>" +
        "<td class='td-driver'>" + esc(t.driverName || "Unassigned") + "</td>" +
        "<td>" + fmtDateReadable(t.start) + "</td>" +
        "<td>" + fmtTime(t.start) + "</td>" +
        "<td><span class='speed-badge'>" + toMph(t.maximumSpeed || 0).toFixed(0) + " mph</span></td>" +
        "<td>" + ((t.distance || 0) * 0.621371).toFixed(1) + " mi</td></tr>";
    }).join("") + "</tbody></table></div>";
}

// ── Upcoming maintenance report ───────────────────────────────────────────────
function runMaintenanceUpcomingReport(done) {
  apiCall("Get", { typeName: "Device", search: {} }, function (devices) {
    var deviceMap = {};
    devices.forEach(function (d) { deviceMap[d.id] = d.name; });
    state.deviceMap = deviceMap;
    apiCall("Get", { typeName: "MaintenanceReminder", search: {} }, function (reminders) {
      var today = new Date(); today.setHours(0, 0, 0, 0);
      var processed = reminders.map(function (r) {
        var dueDate = r.nextServiceDate ? new Date(r.nextServiceDate) : null;
        var daysUntil = dueDate ? Math.round((dueDate - today) / 86400000) : null;
        return {
          vehicle:   deviceMap[r.device && r.device.id] || (r.device && r.device.id) || "Unknown",
          what:      r.comment || r.description || "Scheduled maintenance",
          dueDate:   dueDate,
          dueDateStr:dueDate ? fmtDateReadable(dueDate.toISOString()) : "—",
          odometer:  r.nextOdometerReading ? (r.nextOdometerReading / 1000).toFixed(0) + " km" : "—",
          daysUntil: daysUntil
        };
      }).filter(function (r) { return r.dueDate; });

      state.reportRaw  = processed;
      state.reportData = processed;

      var overdue  = processed.filter(function (r) { return r.daysUntil < 0; }).length;
      var due7     = processed.filter(function (r) { return r.daysUntil >= 0 && r.daysUntil <= 7; }).length;
      var due30    = processed.filter(function (r) { return r.daysUntil > 7 && r.daysUntil <= 30; }).length;
      showReportSummary([
        summaryCard("Overdue",      overdue, ""),
        summaryCard("Due in 7 Days",due7, ""),
        summaryCard("Due in 30 Days",due30, ""),
        summaryCard("Total Items",  processed.length, "")
      ]);
      renderMaintenanceUpcomingOutput(processed);
      document.getElementById("export-csv").classList.remove("hidden");
      done();
    }, function (err) {
      document.getElementById("report-output").innerHTML = "<p class='placeholder' style='color:var(--warning)'>Maintenance reminders not available. Please ensure maintenance reminders are configured in MyGeotab.</p>";
      done(false);
    });
  }, function (err) { reportError(err); done(false); });
}

function renderMaintenanceUpcomingOutput(items) {
  var output = document.getElementById("report-output");
  if (!items.length) { output.innerHTML = "<p class='placeholder'>No maintenance reminders found.</p>"; return; }
  var sorted = items.slice().sort(function (a, b) { return a.daysUntil - b.daysUntil; });
  output.innerHTML = "<div class='dd-table-wrap'><table class='dd-table'>" +
    "<thead><tr><th>Vehicle</th><th>What Is Due</th><th>Date Due</th><th>Odometer</th><th>Status</th></tr></thead><tbody>" +
    sorted.map(function (r) {
      var statusHtml;
      if (r.daysUntil < 0)
        statusHtml = "<span class='badge badge-critical'>Overdue " + Math.abs(r.daysUntil) + "d</span>";
      else if (r.daysUntil <= 7)
        statusHtml = "<span class='badge badge-warning'>Due in " + r.daysUntil + "d</span>";
      else
        statusHtml = "<span class='badge badge-low'>Due in " + r.daysUntil + "d</span>";
      return "<tr><td class='td-vehicle'>" + esc(r.vehicle) + "</td>" +
        "<td>" + esc(r.what) + "</td>" +
        "<td>" + r.dueDateStr + "</td>" +
        "<td>" + r.odometer + "</td>" +
        "<td>" + statusHtml + "</td></tr>";
    }).join("") + "</tbody></table></div>";
}

// ── Maintenance spend report ──────────────────────────────────────────────────
function runMaintenanceSpendReport(from, to, done) {
  var fromDate = new Date(from + "T00:00:00");
  var toDate   = new Date(to   + "T23:59:59");
  apiCall("Get", { typeName: "Device", search: {} }, function (devices) {
    var deviceMap = {};
    devices.forEach(function (d) { deviceMap[d.id] = d.name; });
    state.deviceMap = deviceMap;
    apiCall("Get", { typeName: "MaintenanceReminder", search: {} }, function (reminders) {
      // Show all reminders — both upcoming and overdue — as a spend/history view
      var processed = reminders.map(function (r) {
        var dueDate = r.nextServiceDate ? new Date(r.nextServiceDate) : null;
        return {
          vehicle:    deviceMap[r.device && r.device.id] || "Unknown",
          what:       r.comment || r.description || "Maintenance",
          date:       dueDate,
          dateStr:    dueDate ? fmtDateReadable(dueDate.toISOString()) : "—",
          cost:       r.cost != null ? r.cost : null,
          status:     dueDate && dueDate < new Date() ? "Overdue" : "Upcoming"
        };
      }).filter(function (r) {
        if (!r.date) return false;
        return r.date >= fromDate && r.date <= toDate;
      });

      state.reportRaw  = processed;
      state.reportData = processed;
      populateReportVehicleFilter(processed, null);

      var totalCost = processed.filter(function (r) { return r.cost != null; })
                               .reduce(function (s, r) { return s + r.cost; }, 0);
      var hasCost   = processed.some(function (r) { return r.cost != null; });
      showReportSummary([
        summaryCard("Total Events",  processed.length, ""),
        summaryCard("Overdue",       processed.filter(function (r) { return r.status === "Overdue"; }).length, ""),
        summaryCard(hasCost ? "Total Spend" : "Cost Data", hasCost ? "\u00a3" + totalCost.toFixed(2) : "Log in MYG", ""),
        summaryCard("Vehicles",      [... new Set(processed.map(function (r) { return r.vehicle; }))].length, "")
      ]);
      renderMaintenanceSpendOutput(processed);
      document.getElementById("export-csv").classList.remove("hidden");
      done();
    }, function (err) {
      document.getElementById("report-output").innerHTML = "<p class='placeholder' style='color:var(--warning)'>Maintenance reminders not available in this database.</p>";
      done(false);
    });
  }, function (err) { reportError(err); done(false); });
}

function renderMaintenanceSpendOutput(items) {
  var output = document.getElementById("report-output");
  if (!items.length) { output.innerHTML = "<p class='placeholder'>No maintenance records found for this period.</p>"; return; }
  var sorted = items.slice().sort(function (a, b) { return b.date - a.date; });
  output.innerHTML = "<div class='dd-table-wrap'><table class='dd-table'>" +
    "<thead><tr><th>Vehicle</th><th>Description</th><th>Date</th><th>Cost</th><th>Status</th></tr></thead><tbody>" +
    sorted.map(function (r) {
      var statusHtml = r.status === "Overdue"
        ? "<span class='badge badge-critical'>Overdue</span>"
        : "<span class='badge badge-low'>Upcoming</span>";
      return "<tr><td class='td-vehicle'>" + esc(r.vehicle) + "</td>" +
        "<td>" + esc(r.what) + "</td>" +
        "<td>" + r.dateStr + "</td>" +
        "<td>" + (r.cost != null ? "\u00a3" + r.cost.toFixed(2) : "<span style='color:var(--text-muted)'>Not logged</span>") + "</td>" +
        "<td>" + statusHtml + "</td></tr>";
    }).join("") + "</tbody></table></div>";
}

// ── Fuel economy report ───────────────────────────────────────────────────────
function runFuelEconomyReport(from, to, done) {
  var fromDate = new Date(from + "T00:00:00").toISOString();
  var toDate   = new Date(to   + "T23:59:59").toISOString();
  apiCall("Get", { typeName: "Device", search: {} }, function (devices) {
    var deviceMap = {};
    devices.forEach(function (d) { deviceMap[d.id] = d.name; });
    state.deviceMap = deviceMap;
    apiCall("Get", { typeName: "Trip", search: { fromDate: fromDate, toDate: toDate } }, function (trips) {
      var fuelTrips = trips.filter(function (t) { return (t.fuelUsed || 0) > 0; });
      if (!fuelTrips.length) {
        document.getElementById("report-output").innerHTML = "<p class='placeholder'>No fuel data available. Fuel tracking must be enabled on your devices.</p>";
        done(); return;
      }
      // Group by vehicle
      var byVehicle = {};
      fuelTrips.forEach(function (t) {
        var vid = t.device && t.device.id;
        if (!vid) return;
        if (!byVehicle[vid]) byVehicle[vid] = { name: deviceMap[vid] || vid, fuel: 0, distM: 0, trips: 0 };
        byVehicle[vid].fuel  += (t.fuelUsed || 0);
        byVehicle[vid].distM += (t.distance || 0);
        byVehicle[vid].trips += 1;
      });
      var rows = Object.values(byVehicle);
      state.reportRaw  = rows;
      state.reportData = rows;
      populateReportVehicleFilter(fuelTrips, deviceMap);
      var totalFuel  = fuelTrips.reduce(function (s, t) { return s + (t.fuelUsed || 0); }, 0);
      var totalDistM = fuelTrips.reduce(function (s, t) { return s + (t.distance || 0); }, 0);
      var fleetMpg   = calcMpg(totalDistM, totalFuel);
      var mpgValues  = rows.map(function (r) { return parseFloat(calcMpg(r.distM, r.fuel)) || 0; }).filter(function (v) { return v > 0; });
      var bestMpg    = mpgValues.length ? Math.max.apply(null, mpgValues).toFixed(1) : "—";
      showReportSummary([
        summaryCard("Fleet MPG",    fleetMpg, ""),
        summaryCard("Best Vehicle", bestMpg, "mpg"),
        summaryCard("Total Fuel",   totalFuel.toFixed(0), "L"),
        summaryCard("Vehicles",     rows.length, "")
      ]);
      renderFuelEconomyOutput(rows);
      document.getElementById("export-csv").classList.remove("hidden");
      done();
    }, function (err) { reportError(err); done(false); });
  }, function (err) { reportError(err); done(false); });
}

function renderFuelEconomyOutput(rows) {
  var output = document.getElementById("report-output");
  if (!rows.length) { output.innerHTML = "<p class='placeholder'>No data.</p>"; return; }
  var sorted = rows.slice().sort(function (a, b) { return parseFloat(calcMpg(b.distM, b.fuel)) - parseFloat(calcMpg(a.distM, a.fuel)); });
  output.innerHTML = "<div class='dd-table-wrap'><table class='dd-table'>" +
    "<thead><tr><th>Vehicle</th><th>MPG</th><th>Distance</th><th>Fuel Used</th><th>Trips</th></tr></thead><tbody>" +
    sorted.map(function (r) {
      return "<tr><td class='td-vehicle'>" + esc(r.name) + "</td>" +
        "<td><strong>" + calcMpg(r.distM, r.fuel) + "</strong></td>" +
        "<td>" + (r.distM * 0.621371).toFixed(1) + " mi</td>" +
        "<td>" + r.fuel.toFixed(1) + " L</td>" +
        "<td>" + r.trips + "</td></tr>";
    }).join("") + "</tbody></table></div>";
}
// ── Legacy Trip History Report ────────────────────────────────────────────────
// Replicates a classic "Daily Report" style export (per-vehicle, per-day trip
// log with reverse-geocoded stop addresses) using MyGeotab's own GetAddresses
// method — coordinates never leave Geotab's infrastructure, unlike a third-party
// geocoder. Trip.StopPoint gives the coordinate directly, no LogRecord
// correlation needed.
//
// KNOWN SIMPLIFICATION: the source format the customer provided distinguishes
// a "Working Total" row (work-hours only) from a "Total" row (whole day incl.
// after-hours idling), using Trip's Work*/AfterHours* fields. v1 collapses this
// into a single Total row per day using whole-trip figures. Revisit if the
// customer needs the work/after-hours split.

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
// metres). 1 km = 0.621371 miles. All reports in this file now treat distance as
// km: km->mi via *0.621371, and calcMpg() takes km. (Historically the code
// wrongly divided by 1609.344 / 1000, making distances & MPG ~1000x too small.)
function milesFromDistance(km) { return (km || 0) * 0.621371; }

// Group fetched trips by device, sorted chronologically. Driver names come
// straight off Trip.driverName, matching the convention already used elsewhere
// in this file (renderTripsOutput, renderSpeedingOutput).
function buildLegacyByDevice(trips, deviceMap) {
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

function addressForPoint(pt, addrMap) {
  if (!pt || pt.y == null || pt.x == null) return "(location unknown)";
  var resolved = addrMap[coordKey(pt)];
  return resolved || (pt.y.toFixed(5) + ", " + pt.x.toFixed(5));
}

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

        // Report-wide KPIs (whole period, all vehicles)
        var totalDistM  = trips.reduce(function (s, t) { return s + (t.distance || 0); }, 0);
        var totalDrive  = trips.reduce(function (s, t) { return s + (parseDurationToMins(t.drivingDuration) != null ? parseDurationToMins(t.drivingDuration) : durationMins(t.start, t.stop)); }, 0);
        var totalIdle   = trips.reduce(function (s, t) { return s + (parseDurationToMins(t.idlingDuration) || 0); }, 0);
        var totalStop   = trips.reduce(function (s, t) { return s + (parseDurationToMins(t.stopDuration) || 0); }, 0);
        var numStops    = trips.length;
        var avgStop     = numStops ? totalStop / numStops : 0;

        showReportSummary([
          summaryCard("Total Distance",  milesFromDistance(totalDistM).toFixed(0), "mi"),
          summaryCard("Total Stop Duration", fmtDurWhole(totalStop), ""),
          summaryCard("Total Idle Time", fmtDurWhole(totalIdle), ""),
          summaryCard("Total Travel Duration", fmtDurWhole(totalDrive), ""),
          summaryCard("Average Stop Duration", fmtDurWhole(avgStop), ""),
          summaryCard("Number of Stops", numStops, "")
        ]);

        renderLegacyTripHistoryOutput(byDevice, addrMap);
        document.getElementById("export-csv").classList.remove("hidden");
        document.getElementById("export-pdf-legacy").classList.remove("hidden");
        done();
      });
    }, function (err) { reportError(err); done(false); });
  }, function (err) { reportError(err); done(false); });
}

// Build the day-block rows for one vehicle's trip list, including the synthetic
// "Starting from" / "(Ignition On)" lead-in row derived from the previous trip
// (which may be from an earlier day, even outside the selected range if it
// happened to already be in the fetched dataset).
function buildLegacyDayBlocks(vTrips, addrMap) {
  var days = {}; // 'yyyy-mm-dd' -> { date, rows:[{trip, prevTrip}] }
  var order = [];
  vTrips.forEach(function (t, i) {
    var day = fmtDateShort(t.start);
    if (!days[day]) { days[day] = { date: day, rows: [] }; order.push(day); }
    days[day].rows.push({ trip: t, prevTrip: i > 0 ? vTrips[i - 1] : null });
  });
  return order.map(function (day) { return days[day]; });
}

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
        var stopMins  = parseDurationToMins(t.stopDuration);
        var stopStr   = stopMins != null ? fmtDurPrecise(stopMins) : "\u2014";
        return "<tr>" +
          "<td>" + fmtTime(t.start) + "</td>" +
          "<td>" + milesFromDistance(t.distance).toFixed(2) + " mi<br><span style='color:var(--text-muted);font-size:11px'>" + fmtDurPrecise(driveMins) + "</span></td>" +
          "<td>" + esc(addressForPoint(t.stopPoint, addrMap)) + "</td>" +
          "<td>" + fmtTime(t.stop) + "</td>" +
          "<td>" + fmtDurPrecise(idleMins) + "</td>" +
          "<td>" + stopStr + "</td>" +
        "</tr>";
      }).join("");

      var dayDistM  = block.rows.reduce(function (s, r) { return s + (r.trip.distance || 0); }, 0);
      var dayDrive  = block.rows.reduce(function (s, r) { var dm = parseDurationToMins(r.trip.drivingDuration); return s + (dm != null ? dm : durationMins(r.trip.start, r.trip.stop)); }, 0);
      var dayIdle   = block.rows.reduce(function (s, r) { return s + (parseDurationToMins(r.trip.idlingDuration) || 0); }, 0);
      var dayStop   = block.rows.reduce(function (s, r) { return s + (parseDurationToMins(r.trip.stopDuration) || 0); }, 0);

      var firstTrip = block.rows[0].trip;
      var prevTrip  = block.rows[0].prevTrip;
      var startingFromHtml = "<tr><td colspan='6' style='font-weight:600;background:var(--surface-2)'>" + block.date + " \u2014 Starting from: " + esc(addressForPoint(prevTrip ? prevTrip.stopPoint : null, addrMap)) + "</td></tr>";
      // "(Ignition On)" marks the day's first drive. We deliberately do NOT show a
      // duration here: the gap to the previous trip is overnight/parked time, not
      // idle, and Trip data has no pre-drive idle figure to show instead.
      var ignitionHtml = prevTrip
        ? "<tr style='color:var(--accent)'><td>" + fmtTime(firstTrip.start) + "</td><td colspan='3'>(Ignition On)</td><td>\u2014</td><td></td></tr>"
        : "";

      return "<div class='dd-table-wrap' style='margin-bottom:12px'><table class='dd-table'>" +
        "<thead><tr><th>Start Time</th><th>Distance / Duration</th><th>Stop Location</th><th>Arrival Time</th><th>Idle Duration</th><th>Stop Duration</th></tr></thead>" +
        "<tbody>" + startingFromHtml + ignitionHtml + rowsHtml +
        "<tr style='font-weight:700;background:var(--surface-2)'><td>" + block.date + " Total</td><td>" + milesFromDistance(dayDistM).toFixed(2) + " mi in " + fmtDurWhole(dayDrive) + "</td><td></td><td>" + block.rows.length + " stops</td><td>" + fmtDurWhole(dayIdle) + "</td><td>" + fmtDurWhole(dayStop) + "</td></tr>" +
        "</tbody></table></div>";
    }).join("");

    return "<div style='margin-bottom:24px'><h3 style='margin-bottom:4px'>" + esc(v.name) + "</h3>" +
      "<div style='color:var(--text-muted);font-size:12px;margin-bottom:8px'>Driver(s): " + esc(driverNames) + "</div>" +
      dayHtml + "</div>";
  }).join("");
}

// ─── Activity Report ──────────────────────────────────────────────────────────
// An enhanced trip history report with unit-aware distance, selectable engine-hour
// columns, Excel export, and localStorage settings persistence.

function fmtActivityDistance(km) {
  if (state.unitSystem === "Metric") return (km || 0).toFixed(2) + " km";
  return milesFromDistance(km).toFixed(2) + " mi";
}

function fmtActivityDistanceKpi(km) {
  if (state.unitSystem === "Metric") return { value: (km || 0).toFixed(0), unit: "km" };
  return { value: milesFromDistance(km).toFixed(0), unit: "mi" };
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
      var byDevice = buildLegacyByDevice(trips, deviceMap);
      state.reportRaw      = trips;
      state.reportData     = trips;
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
        saveActivitySettings(from, to);
        done();
      });
    }, function (err) { reportError(err); done(false); });
  }, function (err) { reportError(err); done(false); });
}

function renderActivityReport(byDevice, addrMap) {
  var cols   = getActivityCols();
  var output = document.getElementById("report-output");
  var vids   = Object.keys(byDevice);
  if (!vids.length) { output.innerHTML = "<p class='placeholder'>No trips found.</p>"; return; }
  var extraHeaders = activityEngineColHeaders(cols);
  var baseHeaders  = ["Start Time", "Distance / Duration", "Stop Location", "Arrival Time", "Idle Duration", "Stop Duration"];
  var allHeaders   = baseHeaders.concat(extraHeaders);
  var colSpan      = allHeaders.length;

  output.innerHTML = vids.map(function (vid) {
    var v = byDevice[vid];
    var driverNames = Object.keys(v.drivers).join(", ") || "No driver assigned";
    var dayBlocks   = buildLegacyDayBlocks(v.trips, addrMap);

    var dayHtml = dayBlocks.map(function (block) {
      var rowsHtml = block.rows.map(function (r) {
        var t         = r.trip;
        var driveMins = parseDurationToMins(t.drivingDuration); if (driveMins == null) driveMins = durationMins(t.start, t.stop);
        var idleMins  = parseDurationToMins(t.idlingDuration) || 0;
        var stopMins  = parseDurationToMins(t.stopDuration);
        var engineVals= activityEngineColValues(t, cols);
        return "<tr>" +
          "<td>" + fmtTime(t.start) + "</td>" +
          "<td>" + fmtActivityDistance(t.distance) + "<br><span style='color:var(--text-muted);font-size:11px'>" + fmtDurPrecise(driveMins) + "</span></td>" +
          "<td>" + esc(addressForPoint(t.stopPoint, addrMap)) + "</td>" +
          "<td>" + fmtTime(t.stop) + "</td>" +
          "<td>" + fmtDurPrecise(idleMins) + "</td>" +
          "<td>" + (stopMins != null ? fmtDurPrecise(stopMins) : "—") + "</td>" +
          engineVals.map(function (v) { return "<td>" + v + "</td>"; }).join("") +
        "</tr>";
      }).join("");

      var dayDistKm = block.rows.reduce(function (s, r) { return s + (r.trip.distance || 0); }, 0);
      var dayDrive  = block.rows.reduce(function (s, r) { var dm = parseDurationToMins(r.trip.drivingDuration); return s + (dm != null ? dm : durationMins(r.trip.start, r.trip.stop)); }, 0);
      var dayIdle   = block.rows.reduce(function (s, r) { return s + (parseDurationToMins(r.trip.idlingDuration) || 0); }, 0);
      var dayStop   = block.rows.reduce(function (s, r) { return s + (parseDurationToMins(r.trip.stopDuration) || 0); }, 0);
      var engineTotals = activityEngineDayTotals(block.rows, cols);

      var firstTrip = block.rows[0].trip;
      var prevTrip  = block.rows[0].prevTrip;
      var startingFromHtml = "<tr><td colspan='" + colSpan + "' style='font-weight:600;background:var(--surface-2)'>" + block.date + " — Starting from: " + esc(addressForPoint(prevTrip ? prevTrip.stopPoint : null, addrMap)) + "</td></tr>";
      var ignitionHtml = prevTrip
        ? "<tr style='color:var(--accent)'><td>" + fmtTime(firstTrip.start) + "</td><td colspan='3'>(Ignition On)</td><td>—</td><td></td>" + cols.map(function () { return "<td></td>"; }).join("") + "</tr>"
        : "";

      var totalRowCells = [
        "<td>" + block.date + " Total</td>",
        "<td>" + fmtActivityDistance(dayDistKm) + " in " + fmtDurWhole(dayDrive) + "</td>",
        "<td></td>",
        "<td>" + block.rows.length + " stops</td>",
        "<td>" + fmtDurWhole(dayIdle) + "</td>",
        "<td>" + fmtDurWhole(dayStop) + "</td>"
      ].concat(engineTotals.map(function (v) { return "<td>" + v + "</td>"; })).join("");

      return "<div class='dd-table-wrap' style='margin-bottom:12px'><table class='dd-table'>" +
        "<thead><tr>" + allHeaders.map(function (h) { return "<th>" + h + "</th>"; }).join("") + "</tr></thead>" +
        "<tbody>" + startingFromHtml + ignitionHtml + rowsHtml +
        "<tr style='font-weight:700;background:var(--surface-2)'>" + totalRowCells + "</tr>" +
        "</tbody></table></div>";
    }).join("");

    return "<div style='margin-bottom:24px'><h3 style='margin-bottom:4px'>" + esc(v.name) + "</h3>" +
      "<div style='color:var(--text-muted);font-size:12px;margin-bottom:8px'>Driver(s): " + esc(driverNames) + "</div>" +
      dayHtml + "</div>";
  }).join("");
}

// ─── Activity Report settings persistence ─────────────────────────────────────
function saveActivitySettings(from, to) {
  try {
    localStorage.setItem("si_activity_settings", JSON.stringify({
      from: from,
      to:   to,
      cols: getActivityCols(),
      len:  document.getElementById("legacy-length").value
    }));
  } catch (e) {}
}

function restoreActivitySettings() {
  try {
    var raw = localStorage.getItem("si_activity_settings");
    if (!raw) return;
    var s = JSON.parse(raw);
    // Switch report type to activity-report and fire the change event to show controls
    var sel = document.getElementById("report-type");
    if (sel) { sel.value = "activity-report"; sel.dispatchEvent(new Event("change")); }
    if (s.len)  { var lenEl = document.getElementById("legacy-length"); if (lenEl) { lenEl.value = s.len; applyLegacyLength(); } }
    if (s.from) { var fromEl = document.getElementById("filter-from"); if (fromEl) fromEl.value = s.from; }
    if (s.to)   { var toEl   = document.getElementById("filter-to");   if (toEl)   toEl.value   = s.to;   }
    // Re-check saved column selections (default is only total-engine checked)
    ["total-engine", "drive-only", "idle-only", "work-split"].forEach(function (id) {
      var el = document.getElementById("col-" + id);
      if (el) el.checked = (s.cols || ["total-engine"]).indexOf(id) !== -1;
    });
  } catch (e) {}
}

// ─── Report helpers ───────────────────────────────────────────────────────────
function showReportSummary(cards) {
  var el = document.getElementById("report-summary");
  el.innerHTML = cards.join("");
  el.classList.remove("hidden");
}
function summaryCard(label, value, unit) {
  return "<div class='summary-card'><div class='summary-label'>" + label + "</div><div class='summary-value'>" + value + "<span class='summary-unit'>" + unit + "</span></div></div>";
}
function tripStat(label, value) {
  return "<div class='trip-stat'><span class='trip-stat-label'>" + label + "</span><span class='trip-stat-value'>" + value + "</span></div>";
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
    else if (item.vehicle) ids[item.vehicle] = item.vehicle; // already resolved names
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
  var type = document.getElementById("report-type").value;
  var raw  = state.reportRaw;
  if (!vid) {
    state.reportData = raw;
  } else {
    state.reportData = raw.filter(function (item) {
      if (item.device) return item.device.id === vid;
      if (item.vehicle) return item.vehicle === vid || state.deviceMap[vid] === item.vehicle;
      return false;
    });
  }
  if      (type === "trips")                renderTripsOutput(state.reportData, null);
  else if (type === "carbon-monthly")       renderCarbonOutput(state.reportData);
  else if (type === "speeding")          renderSpeedingOutput(state.reportData, null);
  else if (type === "maintenance-upcoming") renderMaintenanceUpcomingOutput(state.reportData);
  else if (type === "maintenance-spend")    renderMaintenanceSpendOutput(state.reportData);
  else if (type === "fuel-economy-daily")   renderFuelEconomyOutput(state.reportData);
  else if (type === "legacy-trip-history") {
    var filteredByDevice = buildLegacyByDevice(state.reportData, state.deviceMap);
    state.legacyByDevice = filteredByDevice;
    renderLegacyTripHistoryOutput(filteredByDevice, state.legacyAddrMap || {});
  }
  else if (type === "activity-report") {
    var actFilteredByDevice = buildLegacyByDevice(state.reportData, state.deviceMap);
    state.activityByDevice = actFilteredByDevice;
    renderActivityReport(actFilteredByDevice, state.activityAddrMap || {});
  }
}

function exportReportCsv() {
  var type = document.getElementById("report-type").value;
  var data = state.reportData;
  if (!data || !data.length) return;
  var rows = [];
  if (type === "trips") {
    rows = [["Vehicle","Driver","Date","Start","Stop","Distance (km)","Duration","Max Speed (mph)","Avg Speed (mph)"]];
    data.forEach(function (t) {
      rows.push([
        state.deviceMap[t.device && t.device.id] || "",
        t.driverName || "",
        fmtDateShort(t.start), fmtTime(t.start), fmtTime(t.stop),
        (t.distance || 0).toFixed(1),
        fmtMins(durationMins(t.start, t.stop)),
        toMph(t.maximumSpeed || 0).toFixed(0),
        toMph(t.averageSpeed  || 0).toFixed(0)
      ]);
    });
  } else if (type === "carbon-monthly") {
    rows = [["Vehicle","CO2 (kg)","MPG","Distance (mi)","Trips"]];
    data.forEach(function (r) { rows.push([r.name, Math.round(r.co2), calcMpg(r.distM, r.fuel), (r.distM*0.621371).toFixed(1), r.trips]); });
  } else if (type === "speeding") {
    rows = [["Vehicle","Driver","Date","Start Time","Max Speed (mph)","Distance (mi)"]];
    data.forEach(function (t) {
      rows.push([
        state.deviceMap[t.device && t.device.id] || "",
        t.driverName || "",
        fmtDateReadable(t.start), fmtTime(t.start),
        toMph(t.maximumSpeed || 0).toFixed(0),
        ((t.distance || 0) * 0.621371).toFixed(1)
      ]);
    });
  } else if (type === "maintenance-upcoming") {
    rows = [["Vehicle","What Is Due","Date Due","Odometer","Days Until Due"]];
    data.forEach(function (r) { rows.push([r.vehicle, r.what, r.dueDateStr, r.odometer, r.daysUntil != null ? r.daysUntil : ""]); });
  } else if (type === "maintenance-spend") {
    rows = [["Vehicle","Description","Date","Cost","Status"]];
    data.forEach(function (r) { rows.push([r.vehicle, r.what, r.dateStr, r.cost != null ? r.cost.toFixed(2) : "", r.status]); });
  } else if (type === "fuel-economy-daily") {
    rows = [["Vehicle","MPG","Distance (mi)","Fuel (L)","Trips"]];
    data.forEach(function (r) { rows.push([r.name, calcMpg(r.distM, r.fuel), (r.distM*0.621371).toFixed(1), r.fuel.toFixed(1), r.trips]); });
  } else if (type === "legacy-trip-history") {
    var addrMap = state.legacyAddrMap || {};
    rows = [["Vehicle","Driver(s)","Date","Start Time","Distance (mi)","Driving Duration","Stop Location","Arrival Time","Idle Duration","Stop Duration"]];
    var byDevice = buildLegacyByDevice(data, state.deviceMap);
    Object.keys(byDevice).forEach(function (vid) {
      var v = byDevice[vid];
      var driverNames = Object.keys(v.drivers).join(", ") || "No driver assigned";
      v.trips.forEach(function (t) {
        var driveMins = parseDurationToMins(t.drivingDuration); if (driveMins == null) driveMins = durationMins(t.start, t.stop);
        var idleMins  = parseDurationToMins(t.idlingDuration) || 0;
        var stopMins  = parseDurationToMins(t.stopDuration);
        rows.push([
          v.name, driverNames, fmtDateShort(t.start), fmtTime(t.start),
          milesFromDistance(t.distance).toFixed(2), fmtDurPrecise(driveMins),
          addressForPoint(t.stopPoint, addrMap), fmtTime(t.stop),
          fmtDurPrecise(idleMins), stopMins != null ? fmtDurPrecise(stopMins) : ""
        ]);
      });
    });
  }
  if (rows.length > 1) downloadCsvBlob(rows, type + "_report.csv");
}
// ─── Activity Report exports ──────────────────────────────────────────────────
function exportActivityCsv() {
  var data    = state.reportData;
  var addrMap = state.activityAddrMap || {};
  var cols    = getActivityCols();
  if (!data || !data.length) return;
  var from = document.getElementById("filter-from").value;
  var to   = document.getElementById("filter-to").value;
  var distUnit = state.unitSystem === "Metric" ? "km" : "mi";
  var baseHeaders = ["Vehicle", "Driver(s)", "Date", "Start Time", "Distance (" + distUnit + ")", "Driving Duration", "Stop Location", "Arrival Time", "Idle Duration", "Stop Duration"];
  var rows = [baseHeaders.concat(activityEngineColHeaders(cols))];
  var byDevice = buildLegacyByDevice(data, state.deviceMap);
  Object.keys(byDevice).forEach(function (vid) {
    var v = byDevice[vid];
    var driverNames = Object.keys(v.drivers).join(", ") || "No driver assigned";
    v.trips.forEach(function (t) {
      var driveMins = parseDurationToMins(t.drivingDuration); if (driveMins == null) driveMins = durationMins(t.start, t.stop);
      var idleMins  = parseDurationToMins(t.idlingDuration) || 0;
      var stopMins  = parseDurationToMins(t.stopDuration);
      var baseRow = [
        v.name, driverNames, fmtDateShort(t.start), fmtTime(t.start),
        fmtActivityDistance(t.distance), fmtDurPrecise(driveMins),
        addressForPoint(t.stopPoint, addrMap), fmtTime(t.stop),
        fmtDurPrecise(idleMins), stopMins != null ? fmtDurPrecise(stopMins) : ""
      ];
      rows.push(baseRow.concat(activityEngineColValues(t, cols)));
    });
  });
  if (rows.length > 1) downloadCsvBlob(rows, "activity_report_" + from + "_to_" + to + ".csv");
}

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

  var allTrips   = [];
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
    { label: "Total Distance (" + (state.unitSystem === "Metric" ? "km" : "mi") + ")", value: distKpi.value },
    { label: "Average Stop Duration", value: fmtDurWhole(avgStop) },
    { label: "Number of Stops",       value: String(numStops) }
  ], yPos);
  yPos += 24;

  var extraHeaders = activityEngineColHeaders(cols);
  var headers = [["Start Time", "Distance / Duration", "Stop Location", "Arrival Time", "Idle Duration", "Stop Duration"].concat(extraHeaders)];

  Object.keys(byDevice).forEach(function (vid) {
    var v = byDevice[vid];
    if (yPos > pageH - 40) { doc.addPage(); yPos = 16; }
    doc.setTextColor(20, 20, 20); doc.setFontSize(12); doc.setFont("helvetica", "bold");
    doc.text(v.name, margin, yPos);
    doc.setFontSize(9); doc.setFont("helvetica", "normal"); doc.setTextColor(100, 100, 100);
    doc.text("Driver(s): " + (Object.keys(v.drivers).join(", ") || "No driver assigned"), margin, yPos + 5);
    yPos += 10;

    var dayBlocks = buildLegacyDayBlocks(v.trips, addrMap);
    dayBlocks.forEach(function (block) {
      var body = [];
      var firstTrip = block.rows[0].trip;
      var prevTrip  = block.rows[0].prevTrip;
      var numCols   = 6 + extraHeaders.length;
      body.push([{ content: block.date + "  Starting from: " + addressForPoint(prevTrip ? prevTrip.stopPoint : null, addrMap), colSpan: numCols, styles: { fontStyle: "bold", fillColor: [245, 245, 245] } }]);

      if (prevTrip) {
        var ignRow = [fmtTime(firstTrip.start), { content: "(Ignition On)", colSpan: 3, styles: { textColor: [217, 119, 6] } }, "—", ""];
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
        styles: { fontSize: 7.5, cellPadding: 2, textColor: [45, 55, 72] },
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

function exportActivityExcel() {
  var byDevice = state.activityByDevice;
  var addrMap  = state.activityAddrMap || {};
  var cols     = getActivityCols();
  if (!byDevice || !Object.keys(byDevice).length) { alert("Run the report first."); return; }
  if (typeof ExcelJS === "undefined") { alert("Excel library not loaded. Please check your internet connection."); return; }
  var from     = document.getElementById("filter-from").value;
  var to       = document.getElementById("filter-to").value;
  var distUnit = state.unitSystem === "Metric" ? "km" : "mi";
  var vids     = Object.keys(byDevice);

  // ── Build per-vehicle KPI data ────────────────────────────────────────────
  var allTrips    = [];
  var vehicleKpis = []; // [{name, distKm, engineMins}]
  vids.forEach(function (vid) {
    var v = byDevice[vid];
    allTrips = allTrips.concat(v.trips);
    var distKm     = v.trips.reduce(function (s, t) { return s + (t.distance || 0); }, 0);
    var driveMins  = v.trips.reduce(function (s, t) { var dm = parseDurationToMins(t.drivingDuration); return s + (dm != null ? dm : durationMins(t.start, t.stop)); }, 0);
    var idleMins   = v.trips.reduce(function (s, t) { return s + (parseDurationToMins(t.idlingDuration) || 0); }, 0);
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
  var distColor   = "#0078D4";
  var engineColor = "#059669";

  var distValues   = vehicleKpis.map(function (k) {
    return state.unitSystem === "Metric" ? parseFloat(k.distKm.toFixed(2)) : parseFloat(milesFromDistance(k.distKm).toFixed(2));
  });
  var engineValues = vehicleKpis.map(function (k) {
    return parseFloat((k.engineMins / 60).toFixed(2)); // hours
  });

  var chartInst = new Chart(chartCanvas.getContext("2d"), {
    type: "bar",
    data: {
      labels: vehicleKpis.map(function (k) { return k.name; }),
      datasets: [
        {
          label: "Distance (" + distUnit + ")",
          data: distValues,
          backgroundColor: distColor,
          borderRadius: 4
        },
        {
          label: "Engine On (hrs)",
          data: engineValues,
          backgroundColor: engineColor,
          borderRadius: 4
        }
      ]
    },
    options: {
      animation: false,
      responsive: false,
      plugins: {
        legend: {
          display: true,
          position: "top",
          labels: { font: { size: 13 }, color: "#374151", padding: 16 }
        },
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
  wb.creator = "Smart Insights";
  wb.created = new Date();

  // ── Summary sheet — KPIs left (cols A-B), chart right (cols D-K) ──────────
  var summarySheet = wb.addWorksheet("Summary");
  summarySheet.columns = [
    { width: 36 }, { width: 24 }, { width: 4 },  // A, B, C (spacer)
    { width: 18 }, { width: 18 }, { width: 18 }, { width: 18 },
    { width: 18 }, { width: 18 }, { width: 18 }   // D-J (chart area)
  ];

  // KPI table — written first so row indices are known
  var titleRow = summarySheet.addRow(["Activity Report — " + from + " to " + to]);
  titleRow.font = { bold: true, size: 13 };
  summarySheet.addRow([]);

  var hdrRow = summarySheet.addRow(["Metric", "Value"]);
  hdrRow.font = { bold: true };
  hdrRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE8F4FD" } };

  var kpiData = [
    ["Total Distance (" + distUnit + ")", distKpi.value + " " + distKpi.unit],
    ["Total Travel Duration", fmtDurWhole(totalDrive)],
    ["Total Idle Time", fmtDurWhole(totalIdle)],
    ["Total Stop Duration", fmtDurWhole(totalStop)],
    ["Number of Stops", allTrips.length],
    ["Average Stop Duration", fmtDurWhole(allTrips.length ? totalStop / allTrips.length : 0)],
    ["Vehicles in Report", vids.length]
  ];
  kpiData.forEach(function (r) { summarySheet.addRow(r); });

  // Chart image — placed to the right of the KPI table (col D = index 3, row 0)
  var imgId = wb.addImage({ base64: chartBase64, extension: "png" });
  summarySheet.addImage(imgId, { tl: { col: 3, row: 0 }, br: { col: 10, row: 20 } });

  // ── Per-vehicle sheets ────────────────────────────────────────────────────
  var extraHeaders  = activityEngineColHeaders(cols);
  var baseHeaders   = ["Start Time", "Distance (" + distUnit + ")", "Stop Location", "Arrival Time", "Idle Duration", "Stop Duration"];
  var allColHeaders = baseHeaders.concat(extraHeaders);
  var colWidths     = [{ width: 12 }, { width: 16 }, { width: 42 }, { width: 12 }, { width: 16 }, { width: 16 }]
                       .concat(extraHeaders.map(function () { return { width: 22 }; }));

  vids.forEach(function (vid) {
    var v           = byDevice[vid];
    var driverNames = Object.keys(v.drivers).join(", ") || "No driver assigned";
    var dayBlocks   = buildLegacyDayBlocks(v.trips, addrMap);
    var sheetName   = v.name.replace(/[\\\/\?\*\[\]:]/g, "_").slice(0, 31) || ("V" + vid.slice(-4));
    var ws          = wb.addWorksheet(sheetName);
    ws.columns      = colWidths;

    var r1 = ws.addRow(["Activity Report — " + v.name]);
    r1.font = { bold: true, size: 12 };
    var r2 = ws.addRow(["Period: " + from + " to " + to + "   |   Driver(s): " + driverNames]);
    r2.font = { color: { argb: "FF6B7280" }, size: 10 };
    ws.addRow([]);

    dayBlocks.forEach(function (block) {
      var prevTrip = block.rows[0].prevTrip;

      // "Starting from" header row
      var sfRow = ws.addRow([block.date + " — Starting from: " + addressForPoint(prevTrip ? prevTrip.stopPoint : null, addrMap)]);
      sfRow.font = { bold: true };
      sfRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF5F5F5" } };

      // Column header row
      var hRow = ws.addRow(allColHeaders);
      hRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
      hRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0078D4" } };

      // Ignition On row
      if (prevTrip) {
        var ignRow = [fmtTime(block.rows[0].trip.start), "(Ignition On)", "", "", "—", ""];
        for (var ei = 0; ei < extraHeaders.length; ei++) ignRow.push("");
        var iRow = ws.addRow(ignRow);
        iRow.font = { color: { argb: "FFD97706" } };
      }

      // Trip rows
      block.rows.forEach(function (r) {
        var t         = r.trip;
        var driveMins = parseDurationToMins(t.drivingDuration); if (driveMins == null) driveMins = durationMins(t.start, t.stop);
        var idleMins  = parseDurationToMins(t.idlingDuration) || 0;
        var stopMins  = parseDurationToMins(t.stopDuration);
        var row = [
          fmtTime(t.start),
          fmtActivityDistance(t.distance),
          addressForPoint(t.stopPoint, addrMap),
          fmtTime(t.stop),
          fmtDurPrecise(idleMins),
          stopMins != null ? fmtDurPrecise(stopMins) : ""
        ];
        ws.addRow(row.concat(activityEngineColValues(t, cols)));
      });

      // Daily total row
      var dayDistKm = block.rows.reduce(function (s, r) { return s + (r.trip.distance || 0); }, 0);
      var dayDrive  = block.rows.reduce(function (s, r) { var dm = parseDurationToMins(r.trip.drivingDuration); return s + (dm != null ? dm : durationMins(r.trip.start, r.trip.stop)); }, 0);
      var dayIdle   = block.rows.reduce(function (s, r) { return s + (parseDurationToMins(r.trip.idlingDuration) || 0); }, 0);
      var dayStop   = block.rows.reduce(function (s, r) { return s + (parseDurationToMins(r.trip.stopDuration) || 0); }, 0);
      var totalRowData = [
        block.date + " Total",
        fmtActivityDistance(dayDistKm) + " in " + fmtDurWhole(dayDrive),
        "",
        block.rows.length + " stops",
        fmtDurWhole(dayIdle),
        fmtDurWhole(dayStop)
      ];
      var tRow = ws.addRow(totalRowData.concat(activityEngineDayTotals(block.rows, cols)));
      tRow.font = { bold: true };
      tRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF5F5F5" } };

      ws.addRow([]);
    });
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

// ─── Suggestions ─────────────────────────────────────────────────────────────
function loadSuggestions() {
  var refreshBtn = document.getElementById("refresh-suggestions");
  if (refreshBtn && !refreshBtn._bound) { refreshBtn.addEventListener("click", loadSuggestions); refreshBtn._bound = true; }
  var list = document.getElementById("suggestions-list");
  list.innerHTML = "<p class='placeholder'>Scanning your fleet data...</p>";
  var nowDate = new Date(), curFrom = new Date(nowDate), prevTo = new Date(), prevFrom = new Date();
  curFrom.setDate(nowDate.getDate() - 30);
  prevTo = new Date(curFrom);
  prevFrom.setDate(curFrom.getDate() - 30);
  var pending = 4, res = { devices: [], current: [], previous: [], rules: [] };
  function onDone() { if (--pending === 0) processAndRenderSuggestions(res, list); }
  apiCall("Get", { typeName: "Device",  search: {} }, function (d) { res.devices = d; onDone(); }, function () { onDone(); });
  apiCall("Get", { typeName: "Rule",    search: {} }, function (r) { res.rules   = r; onDone(); }, function () { onDone(); });
  apiCall("Get", { typeName: "ExceptionEvent", search: { fromDate: curFrom.toISOString(),  toDate: nowDate.toISOString(),  includeInvalidated: false } }, function (e) { res.current  = e; onDone(); }, function () { onDone(); });
  apiCall("Get", { typeName: "ExceptionEvent", search: { fromDate: prevFrom.toISOString(), toDate: prevTo.toISOString(),   includeInvalidated: false } }, function (e) { res.previous = e; onDone(); }, function () { onDone(); });
}
function processAndRenderSuggestions(res, list) {
  var deviceMap = {}; res.devices.forEach(function (d) { deviceMap[d.id] = d.name || d.id; }); state.deviceMap = deviceMap;
  var ruleMap   = {}; res.rules.forEach(function (r) { ruleMap[r.id] = r; });
  var curCounts = {}, byRuleDevice = {};
  res.current.forEach(function (e) {
    var rId = e.rule && e.rule.id, dId = e.device && e.device.id; if (!rId) return;
    curCounts[rId] = (curCounts[rId] || 0) + 1;
    if (dId) { if (!byRuleDevice[rId]) byRuleDevice[rId] = {}; byRuleDevice[rId][dId] = (byRuleDevice[rId][dId] || 0) + 1; }
  });
  var prevCounts = {};
  res.previous.forEach(function (e) { var rId = e.rule && e.rule.id; if (rId) prevCounts[rId] = (prevCounts[rId] || 0) + 1; });
  var ranked = Object.keys(curCounts).map(function (id) {
    var curr = curCounts[id], prev = prevCounts[id] || 0;
    var trend = prev > 0 ? Math.round(((curr - prev) / prev) * 100) : null;
    var topVehicles = Object.keys(byRuleDevice[id] || {}).map(function (did) { return { name: deviceMap[did] || did, count: byRuleDevice[id][did] }; }).sort(function (a, b) { return b.count - a.count; }).slice(0, 3);
    return { rule: ruleMap[id], ruleId: id, count: curr, prev: prev, trend: trend, topVehicles: topVehicles };
  }).filter(function (s) { return s.rule; }).sort(function (a, b) { return b.count - a.count; }).slice(0, 8);
  state.suggestions = ranked;
  renderHealthScore(computeHealthScore(res.current.length, res.devices.length), res.current.length, res.devices.length);
  if (!ranked.length) { list.innerHTML = "<p class='placeholder'>No violations found in the last 30 days.</p>"; return; }
  list.innerHTML = ranked.map(function (s, i) { return suggestionCard(s, i); }).join("");
  list.querySelectorAll(".btn-add").forEach(function (btn) {
    btn.addEventListener("click", function (e) { e.stopPropagation(); addSuggestionToDashboard(state.suggestions[parseInt(btn.dataset.idx)]); });
  });
}
function computeHealthScore(v, d) { return d === 0 ? 100 : Math.max(0, Math.min(100, Math.round(100 - (v / d / 50) * 100))); }
function renderHealthScore(score, violations, devices) {
  var el = document.getElementById("health-score-container"); if (!el) return;
  var cls = score >= 70 ? "score-good" : score >= 40 ? "score-medium" : "score-poor";
  var lbl = score >= 70 ? "Good" : score >= 40 ? "Needs Attention" : "Poor";
  el.innerHTML = "<div class='health-card'><div class='health-score-wrap'><span class='health-score-num " + cls + "'>" + score + "</span><span class='health-score-denom'>/100</span></div><div class='health-divider'></div><div class='health-meta'><div class='health-label'>Fleet Health &mdash; " + lbl + "</div><div class='health-desc'>" + violations + " violation" + (violations !== 1 ? "s" : "") + " across " + devices + " vehicle" + (devices !== 1 ? "s" : "") + " in the last 30 days</div></div><div class='health-bar-wrap'><div class='health-bar-track'><div class='health-bar-fill " + cls + "' style='width:" + score + "%'></div></div></div></div>";
}
function suggestionCard(s, idx) {
  var sev = severity(s.count), sevLabel = sev === "critical" ? "Critical" : sev === "warning" ? "Warning" : "Low";
  return "<div class='suggestion-card sev-" + sev + "'><div class='sug-icon-wrap sev-" + sev + "'>" + ruleIconSvg(s.rule.name || "") + "</div><div class='sug-body'><div class='sug-header'><span class='sug-title'>" + esc(s.rule.name || "Unknown Rule") + "</span><span class='badge badge-" + sev + "'>" + sevLabel + "</span>" + trendBadge(s.trend) + "</div><div class='sug-desc'>" + s.count + " violation" + (s.count !== 1 ? "s" : "") + " in the last 30 days</div>" + topVehiclesHtml(s.topVehicles) + "</div><div class='sug-action'><button class='btn btn-add btn-sm' data-idx='" + idx + "'>+ Dashboard</button></div></div>";
}
function trendBadge(pct) {
  if (pct === null || pct === undefined) return "";
  if (Math.abs(pct) < 5) return "<span class='trend-badge trend-neutral'>&#8212; Stable</span>";
  if (pct > 0) return "<span class='trend-badge trend-up'>&#8593; " + pct + "% vs prev</span>";
  return "<span class='trend-badge trend-down'>&#8595; " + Math.abs(pct) + "% vs prev</span>";
}
function topVehiclesHtml(vehicles) {
  if (!vehicles || !vehicles.length) return "";
  var max = vehicles[0].count;
  return "<div class='sug-vehicles'>" + vehicles.map(function (v) {
    var pct = max > 0 ? Math.round((v.count / max) * 100) : 0;
    return "<div class='sug-vehicle-row'><span class='sug-vehicle-name'>" + esc(v.name) + "</span><div class='sug-vehicle-bar-wrap'><div class='sug-vehicle-bar' style='width:" + pct + "%'></div></div><span class='sug-vehicle-count'>" + v.count + "</span></div>";
  }).join("") + "</div>";
}
function severity(count) { return count >= 30 ? "critical" : count >= 10 ? "warning" : "low"; }
function ruleIconSvg(name) {
  var n = name.toLowerCase(), path;
  if (n.includes("speed"))     path = '<circle cx="12" cy="13" r="2"/><path d="M12 3a9 9 0 0 1 6.4 15.4"/><path d="M12 7v2"/><path d="M6.4 6.4 7.8 7.8"/><path d="M3 13h2"/>';
  else if (n.includes("idle")) path = '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3.5 3.5"/>';
  else if (n.includes("harsh") || n.includes("brake") || n.includes("accel")) path = '<path d="M10.3 5 3.6 16.5A2 2 0 0 0 5.3 19.5H18.7a2 2 0 0 0 1.7-3L13.7 5a2 2 0 0 0-3.4 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>';
  else if (n.includes("seat") || n.includes("belt")) path = '<circle cx="12" cy="7" r="3"/><path d="M9 20v-3a3 3 0 0 1 6 0v3"/><path d="M9 13 15 20"/>';
  else if (n.includes("fuel")) path = '<path d="M3 22V10a2 2 0 0 1 2-2h8a2 2 0 0 1 2 2v12"/><path d="M15 14h1a2 2 0 0 1 2 2v1a2 2 0 0 0 2 2 2 2 0 0 0-2-2v-1a4 4 0 0 0-4-4"/><path d="M3 22h12"/><path d="M7 8V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v4"/>';
  else if (n.includes("after") || n.includes("hour") || n.includes("curfew") || n.includes("night")) path = '<path d="M12 3a6 6 0 0 0 9 9 9 9 0 1 1-9-9z"/>';
  else path = '<rect x="3" y="12" width="4" height="8" rx="1"/><rect x="9" y="8" width="4" height="12" rx="1"/><rect x="15" y="4" width="4" height="16" rx="1"/>';
  return '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">' + path + '</svg>';
}

// ─── Shared chart / geo helpers ───────────────────────────────────────────────
// Distinct per-vehicle colour palette (cycles if fleet is larger than the list)
var VEHICLE_PALETTE = [
  "#0078D4", "#DC2626", "#059669", "#D97706", "#7C3AED",
  "#DB2777", "#0891B2", "#65A30D", "#EA580C", "#2563EB",
  "#9333EA", "#0D9488", "#CA8A04", "#E11D48", "#4F46E5"
];
function vehicleColor(i) { return VEHICLE_PALETTE[i % VEHICLE_PALETTE.length]; }

// Max GPS correlations per map render (user-configured cap)
var MAP_EVENT_CAP = 500;

// Decide bucket granularity from a date span: hour for <=24h, else day
function bucketGranularity(fromDate, toDate) {
  var hours = (toDate - fromDate) / 3600000;
  return hours <= 24.5 ? "hour" : "day";
}

// UTC bucket key for an ISO date string (matches existing fmtDateShort day bucketing)
function bucketKeyFor(iso, gran) {
  if (!iso) return null;
  var s = new Date(iso).toISOString();
  return gran === "hour" ? s.slice(0, 13) : s.slice(0, 10); // yyyy-mm-ddThh  |  yyyy-mm-dd
}

// Ordered [{key,label}] covering the range at the chosen granularity
function buildBuckets(fromDate, toDate, gran) {
  var buckets = [];
  var cursor = new Date(fromDate);
  if (gran === "hour") {
    cursor.setUTCMinutes(0, 0, 0);
    while (cursor <= toDate) {
      buckets.push({ key: cursor.toISOString().slice(0, 13), label: String(cursor.getUTCHours()).padStart(2, "0") + ":00" });
      cursor.setUTCHours(cursor.getUTCHours() + 1);
    }
  } else {
    cursor.setUTCHours(0, 0, 0, 0);
    while (cursor <= toDate) {
      buckets.push({ key: cursor.toISOString().slice(0, 10), label: cursor.toISOString().slice(5, 10) });
      cursor.setUTCDate(cursor.getUTCDate() + 1);
    }
  }
  return buckets;
}

// Build Chart.js stacked datasets (one per vehicle) from a
// { vid: { name, buckets: {key:number} } } map over an ordered bucket list.
function buildStackedDatasets(byVehicle, buckets) {
  return Object.keys(byVehicle).map(function (vid, i) {
    var clr = vehicleColor(i);
    return {
      label: byVehicle[vid].name,
      vehicleId: vid,
      data: buckets.map(function (b) { return byVehicle[vid].buckets[b.key] || 0; }),
      backgroundColor: clr,
      borderColor: clr,
      borderWidth: 1,
      borderRadius: 3,
      stack: "events"
    };
  });
}

// From a clicked bucket, derive a {from,to} window for drilldown
function bucketWindow(bucket, gran) {
  var start, end;
  if (gran === "hour") {
    start = new Date(bucket.key + ":00:00.000Z");
    end   = new Date(start); end.setUTCHours(end.getUTCHours() + 1);
  } else if (gran === "month") {
    start = new Date(bucket.key + "-01T00:00:00.000Z");
    end   = new Date(start); end.setUTCMonth(end.getUTCMonth() + 1);
  } else {
    start = new Date(bucket.key + "T00:00:00.000Z");
    end   = new Date(start); end.setUTCDate(end.getUTCDate() + 1);
  }
  return { from: start, to: end };
}

// ─── Dashboard ───────────────────────────────────────────────────────────────
var DEFAULT_W = 400, DEFAULT_H = 320, CASCADE = 24;

function addSuggestionToDashboard(suggestion) {
  var id = "widget-" + Date.now();
  switchToTab("dashboard");
  var widgetDef = { type: "exception", suggestion: suggestion };
  addWidget(id, suggestion.rule.name + " \u2014 30 days", widgetDef, state.nextX, state.nextY, DEFAULT_W, DEFAULT_H);
  state.nextX = (state.nextX + CASCADE) % 120;
  state.nextY = (state.nextY + CASCADE) % 120;
  saveDashboard();
}

function addWidget(id, title, widgetDef, x, y, w, h) {
  var grid  = document.getElementById("dashboard-grid");
  var empty = grid.querySelector(".grid-empty");
  if (empty) empty.remove();

  var type    = widgetDef.type || "exception";
  var typeDef = PRESET_WIDGET_TYPES[type];
  var dotClr, statLabels, chartType, chartClr;

  if (type === "exception") {
    var sev = severity(widgetDef.suggestion ? widgetDef.suggestion.count || 0 : 0);
    dotClr     = sev === "critical" ? "var(--critical)" : sev === "warning" ? "var(--warning)" : "var(--low)";
    statLabels = ["Total", "Daily Avg", "Worst Day"];
    chartType  = "bar";  // stacked per-vehicle bars (was single line)
    chartClr   = "#0078D4";
  } else {
    dotClr     = typeDef.color;
    statLabels = typeDef.statLabels;
    chartType  = typeDef.chartType;
    chartClr   = typeDef.color;
  }

  var widget = document.createElement("div");
  widget.className = "widget";
  widget.id = id;
  widget.style.left   = (x || 0) + "px";
  widget.style.top    = (y || 0) + "px";
  widget.style.width  = (w || DEFAULT_W) + "px";
  widget.style.height = (h || DEFAULT_H) + "px";

  var hasParams = !!(PRESET_WIDGET_TYPES[type] && PRESET_WIDGET_TYPES[type].hasParams);
  widget.innerHTML =
    "<div class='widget-header'>" +
      "<span class='widget-drag-hint'>&#8942;&#8942;</span>" +
      "<span class='widget-sev-dot' style='background:" + dotClr + "'></span>" +
      "<span class='widget-title' id='title-" + id + "'>" + esc(title) + "</span>" +
      "<button class='widget-edit-name' title='Edit name' data-id='" + id + "'>" +
        "<svg viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7'/><path d='M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z'/></svg>" +
      "</button>" +
      (hasParams
        ? "<button class='widget-edit-params' title='Edit parameters' data-id='" + id + "'>" +
            "<svg viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><circle cx='12' cy='12' r='3'/><path d='M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z'/></svg>" +
          "</button>"
        : "") +
      "<button class='widget-remove' title='Remove'>&#x2715;</button>" +
    "</div>" +
    "<div class='widget-body'><canvas id='canvas-" + id + "'></canvas></div>" +
    "<div class='widget-stats' id='stats-" + id + "'>" +
      statLabels.map(function (lbl, i) {
        return "<div class='widget-stat'><div class='widget-stat-label'>" + lbl + "</div><div class='widget-stat-value' id='stat-" + i + "-" + id + "'>—</div></div>";
      }).join("") +
    "</div>" +
    "<div class='widget-resize'></div>";

  grid.appendChild(widget);

  widget.addEventListener("click", function (e) {
    if (state.editMode) return;
    if (e.target.closest(".widget-header")) return;
    if (e.target.closest(".widget-resize"))  return;
    openDrilldown(widgetDef);
  });
  widget.querySelector(".widget-remove").addEventListener("click", function (e) {
    e.stopPropagation();
    var inst = state.widgets.find(function (w) { return w.id === id; });
    if (inst && inst.chart) inst.chart.destroy();
    state.widgets = state.widgets.filter(function (w) { return w.id !== id; });
    widget.remove();
    if (!grid.querySelector(".widget"))
      grid.innerHTML = "<div class='grid-empty'><p>No widgets yet.</p><p>Click <strong>+ Add Widget</strong> or head to <strong>Suggestions</strong>.</p></div>";
    saveDashboard();
  });

  // Pencil — inline name edit
  var pencilBtn = widget.querySelector(".widget-edit-name");
  if (pencilBtn) {
    pencilBtn.addEventListener("click", function (e) {
      e.stopPropagation();
      if (!state.editMode) return;
      enableNameEdit(id);
    });
  }
  // Settings — param editor
  var paramsBtn = widget.querySelector(".widget-edit-params");
  if (paramsBtn) {
    paramsBtn.addEventListener("click", function (e) {
      e.stopPropagation();
      if (!state.editMode) return;
      openParamEditor(id);
    });
  }

  var wState = { id: id, title: title, widgetDef: widgetDef, x: x || 0, y: y || 0, w: w || DEFAULT_W, h: h || DEFAULT_H, chart: null };
  state.widgets.push(wState);
  makeDraggable(widget, wState);
  makeResizable(widget, wState);

  // Build chart
  var bgColor = chartType === "bar" ? hexToRgba(chartClr, 0.72) : hexToRgba(chartClr, 0.08);
  var canvas  = document.getElementById("canvas-" + id);
  var chart   = new Chart(canvas, {
    type: chartType,
    data: { labels: [], datasets: [{ label: "Data", data: [], borderColor: chartClr, backgroundColor: bgColor, borderWidth: chartType === "bar" ? 1 : 2, tension: chartType === "line" ? 0.4 : 0, fill: chartType === "line", pointRadius: chartType === "line" ? 2 : 0, pointHoverRadius: 5, pointBackgroundColor: chartClr, borderRadius: chartType === "bar" ? 4 : 0 }] },
    options: {
      responsive: true, maintainAspectRatio: false, devicePixelRatio: window.devicePixelRatio || 2,
      interaction: { mode: "index", intersect: false },
      plugins: {
        legend: { display: false }, // toggled on per-vehicle in loaders that stack
        tooltip: { backgroundColor: "#111827", titleColor: "#9CA3AF", bodyColor: "#F9FAFB", padding: 10, cornerRadius: 6 }
      },
      scales: {
        x: { stacked: true, grid: { color: "rgba(0,0,0,.04)" }, ticks: { color: "#9CA3AF", font: { size: 10 }, maxRotation: 0, autoSkip: true, maxTicksLimit: 12 } },
        y: { stacked: true, beginAtZero: true, grid: { color: "rgba(0,0,0,.04)" }, ticks: { color: "#9CA3AF", font: { size: 10 }, precision: 0 } }
      },
      // Click a vehicle segment → drill into that vehicle for the clicked bucket.
      // Click empty chart area → event bubbles to the widget → full drilldown.
      onClick: function (evt, elements) {
        if (!elements || !elements.length) return;
        var meta = chart._siMeta;
        if (!meta) return;
        if (evt.native && evt.native.stopPropagation) evt.native.stopPropagation();
        var el      = elements[0];
        var ds      = chart.data.datasets[el.datasetIndex];
        var vid     = ds && ds.vehicleId;
        var bucket  = meta.buckets && meta.buckets[el.index];
        if (!vid || !bucket) return;
        var win = bucketWindow(bucket, meta.granularity);
        openDrilldown(meta.widgetDef, { vehicleId: vid, from: win.from, to: win.to });
      }
    }
  });
  chart._siMeta = null; // populated by stacked loaders with {widgetDef, buckets, granularity}
  wState.chart = chart;
  loadWidgetData(id, widgetDef, chart);
}
// ─── Overlap prevention ────────────────────────────────────────────────────────
var SNAP_GAP = 8; // px gap left between a snapped widget and its neighbour

function rectsOverlap(a, b) {
  return a.x < b.x + b.w &&
         a.x + a.w > b.x &&
         a.y < b.y + b.h &&
         a.y + a.h > b.y;
}

// Pure function — computes the nearest overlap-free {x, y} for the dragged widget.
// rx/ry are grid-local coordinates of the dragged widget's intended position.
// Other widgets' geometry is read directly from the DOM via getBoundingClientRect
// so the numbers reflect actual rendered pixels regardless of CSS box model or borders.
// Never mutates anything.
function computeSnapPosition(rx, ry, wState) {
  var grid     = document.getElementById("dashboard-grid");
  var gridRect = grid.getBoundingClientRect();
  var dw = wState.w, dh = wState.h;

  // Build grid-local rects for every other widget from the DOM — source of truth
  var others = [];
  state.widgets.forEach(function (w) {
    if (w.id === wState.id) return;
    var el = document.getElementById(w.id);
    if (!el) return;
    var r = el.getBoundingClientRect();
    others.push({ x: r.left - gridRect.left, y: r.top - gridRect.top, w: r.width, h: r.height });
  });

  // Find first overlapping widget
  var blocker = null;
  for (var i = 0; i < others.length; i++) {
    var o = others[i];
    if (rx < o.x + o.w && rx + dw > o.x && ry < o.y + o.h && ry + dh > o.y) {
      blocker = o; break;
    }
  }
  if (!blocker) return { x: rx, y: ry }; // no overlap — nothing to do

  // Distances from dragged centre to each edge of the blocker
  var cx      = rx + dw / 2;
  var cy      = ry + dh / 2;
  var dRight  = Math.abs(cx - (blocker.x + blocker.w));
  var dLeft   = Math.abs(cx - blocker.x);
  var dBottom = Math.abs(cy - (blocker.y + blocker.h));
  var dTop    = Math.abs(cy - blocker.y);
  var minD    = Math.min(dRight, dLeft, dBottom, dTop);

  var snapX = rx, snapY = ry;

  if (minD === dRight || minD === dLeft) {
    // ── Horizontal snap: keep y, resolve x ──────────────────────────────────
    var snapRight  = (minD === dRight);
    var hDir       = snapRight ? 1 : -1;
    var candidateX = snapRight
      ? blocker.x + blocker.w + SNAP_GAP
      : blocker.x - dw - SNAP_GAP;

    var hBlocking = others.filter(function (o) {
      return o.y < ry + dh && o.y + o.h > ry;
    }).sort(function (a, b) { return hDir === 1 ? a.x - b.x : b.x - a.x; });

    for (var hi = 0; hi < hBlocking.length; hi++) {
      var ho = hBlocking[hi];
      if (hDir === 1) {
        if (ho.x + ho.w <= candidateX) continue;      // fully left of candidate — skip
        if (ho.x >= candidateX + dw) break;           // gap found — fits here
        candidateX = ho.x + ho.w + SNAP_GAP;          // push further right
      } else {
        if (ho.x >= candidateX + dw) continue;        // fully right of candidate — skip
        if (ho.x + ho.w <= candidateX) break;         // gap found — fits here
        candidateX = ho.x - dw - SNAP_GAP;            // push further left
      }
    }
    snapX = Math.max(0, candidateX);

  } else {
    // ── Vertical snap: keep x, resolve y ────────────────────────────────────
    var snapDown   = (minD === dBottom);
    var vDir       = snapDown ? 1 : -1;
    var candidateY = snapDown
      ? blocker.y + blocker.h + SNAP_GAP
      : blocker.y - dh - SNAP_GAP;

    var vBlocking = others.filter(function (o) {
      return o.x < rx + dw && o.x + o.w > rx;
    }).sort(function (a, b) { return vDir === 1 ? a.y - b.y : b.y - a.y; });

    for (var vi = 0; vi < vBlocking.length; vi++) {
      var vo = vBlocking[vi];
      if (vDir === 1) {
        if (vo.y + vo.h <= candidateY) continue;      // fully above candidate — skip
        if (vo.y >= candidateY + dh) break;           // gap found — fits here
        candidateY = vo.y + vo.h + SNAP_GAP;          // push further down
      } else {
        if (vo.y >= candidateY + dh) continue;        // fully below candidate — skip
        if (vo.y + vo.h <= candidateY) break;         // gap found — fits here
        candidateY = vo.y - dh - SNAP_GAP;            // push further up
      }
    }
    snapY = Math.max(0, candidateY);
  }

  return { x: snapX, y: snapY };
}

// ─── Drag / Resize ────────────────────────────────────────────────────────────
function makeDraggable(widget, wState) {
  var header = widget.querySelector(".widget-header");
  header.addEventListener("mousedown", function (e) {
    if (!state.editMode) return;
    if (e.target.closest(".widget-remove")) return;
    e.preventDefault();
    var startX = e.clientX - wState.x, startY = e.clientY - wState.y;
    header.style.cursor = "grabbing"; document.body.style.userSelect = "none";
    function onMove(e) {
      var grid = document.getElementById("dashboard-grid");
      // Raw position tracks the mouse — wState holds this so overlap detection
      // always uses the true cursor-derived geometry, not a previously snapped value.
      wState.x = Math.max(0, Math.min(e.clientX - startX, grid.offsetWidth - wState.w));
      wState.y = Math.max(0, e.clientY - startY);
      // DOM shows the snapped display position live.
      var s = computeSnapPosition(wState.x, wState.y, wState);
      widget.style.left = s.x + "px"; widget.style.top = s.y + "px";
    }
    function onUp() {
      header.style.cursor = "grab"; document.body.style.userSelect = "";
      document.removeEventListener("mousemove", onMove); document.removeEventListener("mouseup", onUp);
      // Commit snapped position to wState so saveDashboard persists correct coordinates.
      var s = computeSnapPosition(wState.x, wState.y, wState);
      wState.x = s.x; wState.y = s.y;
      widget.style.left = wState.x + "px"; widget.style.top = wState.y + "px";
      saveDashboard();
    }
    document.addEventListener("mousemove", onMove); document.addEventListener("mouseup", onUp);
  });
}
function makeResizable(widget, wState) {
  var handle = widget.querySelector(".widget-resize");
  handle.addEventListener("mousedown", function (e) {
    if (!state.editMode) return;
    e.preventDefault(); e.stopPropagation();
    var startX = e.clientX, startY = e.clientY, startW = wState.w, startH = wState.h;
    document.body.style.userSelect = "none";
    function onMove(e) {
      wState.w = Math.max(280, startW + (e.clientX - startX)); wState.h = Math.max(240, startH + (e.clientY - startY));
      widget.style.width = wState.w + "px"; widget.style.height = wState.h + "px";
      if (wState.chart) wState.chart.resize();
    }
    function onUp() {
      document.body.style.userSelect = "";
      document.removeEventListener("mousemove", onMove); document.removeEventListener("mouseup", onUp);
      saveDashboard();
    }
    document.addEventListener("mousemove", onMove); document.addEventListener("mouseup", onUp);
  });
}

// ─── Widget data loaders ──────────────────────────────────────────────────────
function loadWidgetData(widgetId, widgetDef, chart) {
  var type = widgetDef.type || "exception";
  if      (type === "exception")             loadExceptionWidget(widgetId, widgetDef, chart);
  else if (type === "carbon-monthly")        loadCarbonWidget(widgetId, widgetDef, chart);
  else if (type === "speeding")              loadSpeedingWidget(widgetId, widgetDef, chart);
  else if (type === "maintenance-upcoming")  loadMaintenanceUpcomingWidget(widgetId, chart);
  else if (type === "maintenance-spend")     loadMaintenanceSpendWidget(widgetId, chart);
  else if (type === "fuel-economy-daily")    loadFuelEconomyWidget(widgetId, widgetDef, chart);
}

function setWidgetStat(widgetId, index, value) {
  var el = document.getElementById("stat-" + index + "-" + widgetId);
  if (el) el.textContent = value;
}

// Ensure state.deviceMap is populated before a loader needs vehicle names.
// Widgets can restore before loadSuggestions() finishes, so fetch on demand.
function ensureDeviceMap(cb) {
  if (state.deviceMap && Object.keys(state.deviceMap).length) { cb(); return; }
  apiCall("Get", { typeName: "Device", search: {} }, function (devices) {
    var m = {}; devices.forEach(function (d) { m[d.id] = d.name || d.id; });
    state.deviceMap = m; cb();
  }, function () { cb(); });
}

// Turn on a compact bottom legend for per-vehicle stacked/line charts.
function enableVehicleLegend(chart) {
  chart.options.plugins.legend = {
    display: true, position: "bottom",
    labels: { boxWidth: 10, boxHeight: 10, font: { size: 10 }, color: "#6B7280", padding: 8, usePointStyle: true }
  };
}

// Exception widget — per-day (or per-hour) event count, stacked per vehicle, click to drill
function loadExceptionWidget(widgetId, widgetDef, chart) {
  var suggestion = widgetDef.suggestion;
  var toDate = new Date(), fromDate = new Date();
  fromDate.setDate(toDate.getDate() - 30);
  var gran    = bucketGranularity(fromDate, toDate);
  var buckets = buildBuckets(fromDate, toDate, gran);

  ensureDeviceMap(function () {
    apiCall("Get", { typeName: "ExceptionEvent", search: { fromDate: fromDate.toISOString(), toDate: toDate.toISOString(), ruleSearch: { id: suggestion.rule.id }, includeInvalidated: false } },
      function (exceptions) {
        var byVehicle = {};
        exceptions.forEach(function (e) {
          var vid = e.device && e.device.id; if (!vid) return;
          var key = bucketKeyFor(e.activeFrom, gran); if (!key) return;
          if (!byVehicle[vid]) byVehicle[vid] = { name: state.deviceMap[vid] || vid, buckets: {} };
          byVehicle[vid].buckets[key] = (byVehicle[vid].buckets[key] || 0) + 1;
        });
        chart.data.labels   = buckets.map(function (b) { return b.label; });
        chart.data.datasets = buildStackedDatasets(byVehicle, buckets);
        enableVehicleLegend(chart);
        chart._siMeta = { widgetDef: widgetDef, buckets: buckets, granularity: gran };
        chart.update();

        var total = exceptions.length;
        var days  = Math.max(1, buckets.length);
        var perBucketTotals = buckets.map(function (b) {
          return Object.keys(byVehicle).reduce(function (s, vid) { return s + (byVehicle[vid].buckets[b.key] || 0); }, 0);
        });
        var peak = perBucketTotals.length ? Math.max.apply(null, perBucketTotals) : 0;
        setWidgetStat(widgetId, 0, total);
        setWidgetStat(widgetId, 1, (total / days).toFixed(1) + (gran === "hour" ? "/hr" : "/day"));
        setWidgetStat(widgetId, 2, peak);
      }, function () {});
  });
}

// Carbon widget — monthly CO2 (last 12 months), stacked per vehicle, click to drill.
// CO2 is additive so stacking is meaningful; granularity stays monthly by nature.
function loadCarbonWidget(widgetId, widgetDef, chart) {
  var toDate = new Date(), fromDate = new Date();
  fromDate.setFullYear(fromDate.getFullYear() - 1); fromDate.setDate(1);

  // Build 12 month buckets {key:'yyyy-mm', label:'Mon'}
  var buckets = [];
  for (var i = 11; i >= 0; i--) {
    var d = new Date(); d.setDate(1); d.setMonth(d.getMonth() - i);
    buckets.push({ key: d.toISOString().slice(0, 7), label: d.toLocaleString(undefined, { month: "short" }) });
  }

  ensureDeviceMap(function () {
    apiCall("Get", { typeName: "Trip", search: { fromDate: fromDate.toISOString(), toDate: toDate.toISOString() } }, function (trips) {
      var byVehicle = {};
      trips.forEach(function (t) {
        if (!(t.fuelUsed > 0)) return;
        var vid = t.device && t.device.id; if (!vid) return;
        var mo = t.start ? t.start.slice(0, 7) : null; if (!mo) return;
        if (!byVehicle[vid]) byVehicle[vid] = { name: state.deviceMap[vid] || vid, buckets: {} };
        byVehicle[vid].buckets[mo] = (byVehicle[vid].buckets[mo] || 0) + (t.fuelUsed || 0) * 2.4;
      });
      // round each vehicle's monthly CO2
      Object.keys(byVehicle).forEach(function (vid) {
        Object.keys(byVehicle[vid].buckets).forEach(function (k) { byVehicle[vid].buckets[k] = Math.round(byVehicle[vid].buckets[k]); });
      });
      chart.data.labels   = buckets.map(function (b) { return b.label; });
      chart.data.datasets = buildStackedDatasets(byVehicle, buckets);
      enableVehicleLegend(chart);
      chart._siMeta = { widgetDef: widgetDef, buckets: buckets, granularity: "month" };
      chart.update();

      var fuelTrips = trips.filter(function (t) { return (t.fuelUsed || 0) > 0; });
      var totalCo2  = fuelTrips.reduce(function (s, t) { return s + (t.fuelUsed || 0) * 2.4; }, 0);
      var totalFuel = fuelTrips.reduce(function (s, t) { return s + (t.fuelUsed || 0); }, 0);
      var totalDist = fuelTrips.reduce(function (s, t) { return s + (t.distance || 0); }, 0);
      setWidgetStat(widgetId, 0, Math.round(totalCo2) + " kg");
      setWidgetStat(widgetId, 1, calcMpg(totalDist, totalFuel) + " mpg");
      setWidgetStat(widgetId, 2, "—");
    }, function () { setWidgetStat(widgetId, 0, "Error"); });
  });
}

// Speeding widget — per-day (or per-hour) incident count, stacked per vehicle, click to drill
function loadSpeedingWidget(widgetId, widgetDef, chart) {
  var params    = widgetDef.params || {};
  var threshKmh = (params.thresholdMph || 80) / 0.621371;
  var toDate = new Date(), fromDate = new Date();
  fromDate.setDate(toDate.getDate() - 30);
  var gran    = bucketGranularity(fromDate, toDate);
  var buckets = buildBuckets(fromDate, toDate, gran);

  ensureDeviceMap(function () {
    apiCall("Get", { typeName: "Trip", search: { fromDate: fromDate.toISOString(), toDate: toDate.toISOString() } }, function (trips) {
      var speeding = trips.filter(function (t) { return (t.maximumSpeed || 0) > threshKmh; });
      var byVehicle = {};
      speeding.forEach(function (t) {
        var vid = t.device && t.device.id; if (!vid) return;
        var key = bucketKeyFor(t.start, gran); if (!key) return;
        if (!byVehicle[vid]) byVehicle[vid] = { name: state.deviceMap[vid] || vid, buckets: {} };
        byVehicle[vid].buckets[key] = (byVehicle[vid].buckets[key] || 0) + 1;
      });
      chart.data.labels   = buckets.map(function (b) { return b.label; });
      chart.data.datasets = buildStackedDatasets(byVehicle, buckets);
      enableVehicleLegend(chart);
      chart._siMeta = { widgetDef: widgetDef, buckets: buckets, granularity: gran };
      chart.update();

      var total = speeding.length;
      var days  = Math.max(1, buckets.length);
      var perBucketTotals = buckets.map(function (b) {
        return Object.keys(byVehicle).reduce(function (s, vid) { return s + (byVehicle[vid].buckets[b.key] || 0); }, 0);
      });
      var peak = perBucketTotals.length ? Math.max.apply(null, perBucketTotals) : 0;
      setWidgetStat(widgetId, 0, total);
      setWidgetStat(widgetId, 1, (total / days).toFixed(1) + (gran === "hour" ? "/hr" : "/day"));
      setWidgetStat(widgetId, 2, peak);
    }, function () { setWidgetStat(widgetId, 0, "Error"); });
  });
}

// Maintenance upcoming widget — urgency buckets as bar chart
function loadMaintenanceUpcomingWidget(widgetId, chart) {
  apiCall("Get", { typeName: "MaintenanceReminder", search: {} }, function (reminders) {
    var today = new Date(); today.setHours(0, 0, 0, 0);
    var overdue = 0, due7 = 0, due30 = 0, due90 = 0;
    reminders.forEach(function (r) {
      var d = r.nextServiceDate ? new Date(r.nextServiceDate) : null; if (!d) return;
      var days = Math.round((d - today) / 86400000);
      if (days < 0)        overdue++;
      else if (days <= 7)  due7++;
      else if (days <= 30) due30++;
      else if (days <= 90) due90++;
    });
    chart.data.labels = ["Overdue", "0-7d", "8-30d", "31-90d"];
    chart.data.datasets[0].data = [overdue, due7, due30, due90];
    chart.data.datasets[0].backgroundColor = [
      hexToRgba("#DC2626", 0.8), hexToRgba("#D97706", 0.8),
      hexToRgba("#059669", 0.7), hexToRgba("#0078D4", 0.6)
    ];
    chart.update();
    setWidgetStat(widgetId, 0, overdue);
    setWidgetStat(widgetId, 1, due7);
    setWidgetStat(widgetId, 2, due30);
  }, function () { setWidgetStat(widgetId, 0, "N/A"); });
}

// Maintenance spend widget — monthly completed events YTD
function loadMaintenanceSpendWidget(widgetId, chart) {
  var now = new Date();
  apiCall("Get", { typeName: "MaintenanceReminder", search: {} }, function (reminders) {
    var buckets = {};
    for (var i = 0; i < 12; i++) {
      var d = new Date(now.getFullYear(), i, 1);
      buckets[d.toISOString().slice(0, 7)] = 0;
    }
    var ytdTotal = 0, thisMonth = 0, pending = 0;
    var currentMo = now.toISOString().slice(0, 7);
    reminders.forEach(function (r) {
      var d = r.nextServiceDate ? new Date(r.nextServiceDate) : null; if (!d) return;
      var mo = d.toISOString().slice(0, 7);
      if (buckets[mo] !== undefined && d <= now) {
        buckets[mo]++;
        ytdTotal++;
        if (mo === currentMo) thisMonth++;
      }
      if (d > now) pending++;
    });
    var labels = Object.keys(buckets).sort().map(function (k) {
      var d = new Date(k + "-01"); return d.toLocaleString(undefined, { month: "short" });
    });
    var data = Object.keys(buckets).sort().map(function (k) { return buckets[k]; });
    chart.data.labels = labels; chart.data.datasets[0].data = data; chart.update();
    setWidgetStat(widgetId, 0, ytdTotal);
    setWidgetStat(widgetId, 1, thisMonth);
    setWidgetStat(widgetId, 2, pending);
  }, function () { setWidgetStat(widgetId, 0, "N/A"); });
}

// Fuel economy widget — per-vehicle daily MPG as separate lines (NOT stacked:
// MPG is a ratio/average, so summing across vehicles is meaningless). Click a
// vehicle's point to drill into that vehicle.
function loadFuelEconomyWidget(widgetId, widgetDef, chart) {
  var toDate = new Date(), fromDate = new Date();
  fromDate.setDate(toDate.getDate() - 30);
  var gran    = bucketGranularity(fromDate, toDate);
  var buckets = buildBuckets(fromDate, toDate, gran);

  // Averages must not stack — override the shared stacked scales for this chart.
  chart.options.scales.x.stacked = false;
  chart.options.scales.y.stacked = false;

  ensureDeviceMap(function () {
    apiCall("Get", { typeName: "Trip", search: { fromDate: fromDate.toISOString(), toDate: toDate.toISOString() } }, function (trips) {
      var fuelTrips = trips.filter(function (t) { return (t.fuelUsed || 0) > 0; });
      // Per vehicle per bucket accumulate fuel + distance, then derive MPG
      var byVehicle = {};
      fuelTrips.forEach(function (t) {
        var vid = t.device && t.device.id; if (!vid) return;
        var key = bucketKeyFor(t.start, gran); if (!key) return;
        if (!byVehicle[vid]) byVehicle[vid] = { name: state.deviceMap[vid] || vid, acc: {} };
        if (!byVehicle[vid].acc[key]) byVehicle[vid].acc[key] = { fuel: 0, dist: 0 };
        byVehicle[vid].acc[key].fuel += (t.fuelUsed || 0);
        byVehicle[vid].acc[key].dist += (t.distance || 0);
      });
      chart.data.labels = buckets.map(function (b) { return b.label; });
      chart.data.datasets = Object.keys(byVehicle).map(function (vid, i) {
        var clr = vehicleColor(i);
        return {
          label: byVehicle[vid].name, vehicleId: vid,
          data: buckets.map(function (b) {
            var a = byVehicle[vid].acc[b.key];
            return a && a.fuel > 0 ? parseFloat(calcMpg(a.dist, a.fuel)) : null;
          }),
          borderColor: clr, backgroundColor: hexToRgba(clr, 0.08),
          borderWidth: 2, tension: 0.4, fill: false, spanGaps: true,
          pointRadius: 2, pointHoverRadius: 5, pointBackgroundColor: clr
        };
      });
      enableVehicleLegend(chart);
      chart._siMeta = { widgetDef: widgetDef, buckets: buckets, granularity: gran };
      chart.update();

      var allFuel = fuelTrips.reduce(function (s, t) { return s + (t.fuelUsed || 0); }, 0);
      var allDist = fuelTrips.reduce(function (s, t) { return s + (t.distance || 0); }, 0);
      var allMpg  = chart.data.datasets.reduce(function (arr, ds) {
        ds.data.forEach(function (v) { if (v != null) arr.push(v); }); return arr;
      }, []);
      setWidgetStat(widgetId, 0, calcMpg(allDist, allFuel) + " mpg");
      setWidgetStat(widgetId, 1, allMpg.length ? Math.max.apply(null, allMpg).toFixed(1) + " mpg" : "—");
      setWidgetStat(widgetId, 2, allMpg.length ? Math.min.apply(null, allMpg).toFixed(1) + " mpg" : "—");
    }, function () { setWidgetStat(widgetId, 0, "N/A"); });
  });
}
// ─── Drill-down modal ─────────────────────────────────────────────────────────
function setupModal() {
  document.getElementById("modal-close").addEventListener("click", closeDrilldown);
  document.getElementById("drilldown-modal").addEventListener("click", function (e) { if (e.target === this) closeDrilldown(); });
  document.getElementById("dd-apply").addEventListener("click", applyDrilldownFilters);
  document.getElementById("dd-export-excel").addEventListener("click", exportDrilldownExcel);
  document.getElementById("dd-export-pdf").addEventListener("click", exportDrilldownPdf);
  document.getElementById("dd-view-table").addEventListener("click", function () { switchDrilldownView("table"); });
  document.getElementById("dd-view-map").addEventListener("click", function () { switchDrilldownView("map"); });
}

function openDrilldown(widgetDef, prefilter) {
  var type    = widgetDef.type || "exception";
  var typeDef = PRESET_WIDGET_TYPES[type];
  state.drilldown.widgetType = type;
  state.drilldown.widgetDef  = widgetDef;
  state.drilldown.data       = [];
  state.drilldown.filtered   = [];
  state.drilldown.meta       = {};
  state.drilldown.prefilter  = prefilter || null; // {vehicleId, from, to} from a chart click

  // Header
  var title, subtitle, dotClr;
  if (type === "exception") {
    var sev = severity(widgetDef.suggestion ? widgetDef.suggestion.count || 0 : 0);
    dotClr   = sev === "critical" ? "var(--critical)" : sev === "warning" ? "var(--warning)" : "var(--low)";
    title    = widgetDef.suggestion && widgetDef.suggestion.rule ? widgetDef.suggestion.rule.name : "Rule Detail";
    subtitle = "Violation detail — last 30 days";
  } else {
    dotClr   = typeDef ? typeDef.color : "var(--accent)";
    title    = typeDef ? typeDef.label : type;
    subtitle = typeDef ? typeDef.description : "";
  }
  document.getElementById("modal-title").textContent    = title;
  document.getElementById("modal-subtitle").textContent = subtitle;
  document.getElementById("modal-sev-dot").style.background = dotClr;

  // Reset UI
  document.getElementById("dd-summary").classList.add("hidden");
  document.getElementById("dd-table-wrap").innerHTML = "<p class='placeholder'>Loading...</p>";
  document.getElementById("dd-export-excel").disabled = true;
  document.getElementById("dd-export-pdf").disabled   = true;
  document.getElementById("drilldown-modal").classList.remove("hidden");

  // Build type-specific filters, then load initial data
  buildDrilldownFilters(type, widgetDef);
}

function closeDrilldown() {
  document.getElementById("drilldown-modal").classList.add("hidden");
  if (state.ddMap) { state.ddMap.remove(); state.ddMap = null; }
  state.drilldown = { widgetType: null, widgetDef: null, data: [], filtered: [], meta: {}, prefilter: null };
}

// Set the vehicle dropdown from an active prefilter (after it's been populated).
function applyPrefilterVehicle() {
  var pf = state.drilldown.prefilter;
  if (pf && pf.vehicleId) {
    var sel = document.getElementById("dd-vehicle");
    if (sel) sel.value = pf.vehicleId;
  }
}

// ── Filter builders ───────────────────────────────────────────────────────────
function filterFieldHtml(id, label, type) {
  if (type === "date")
    return "<div class='filter-group'><label class='filter-label'>" + label + "</label><input type='date' id='" + id + "' class='input-date' /></div>";
  if (type === "select")
    return "<div class='filter-group'><label class='filter-label'>" + label + "</label><select id='" + id + "' class='select'></select></div>";
  return "";
}

function buildDrilldownFilters(type, widgetDef) {
  var pf = state.drilldown.prefilter || null;
  var toDate = new Date(), fromDate = new Date();
  fromDate.setDate(toDate.getDate() - 30);
  if (pf && pf.from && pf.to) { fromDate = new Date(pf.from); toDate = new Date(pf.to); }
  var fields = document.getElementById("dd-filter-fields");

  if (type === "exception") {
    fields.innerHTML =
      filterFieldHtml("dd-from",     "From",    "date") +
      filterFieldHtml("dd-to",       "To",      "date") +
      filterFieldHtml("dd-vehicle",  "Vehicle", "select") +
      filterFieldHtml("dd-driver",   "Driver",  "select");
    document.getElementById("dd-from").value = fmtDateInput(fromDate);
    document.getElementById("dd-to").value   = fmtDateInput(toDate);
    var pending = 2, devices = [], drivers = [];
    function onDone() {
      if (--pending > 0) return;
      populateDropdown("dd-vehicle", devices.map(function (d) { return { id: d.id, label: d.name || d.id }; }), "All vehicles");
      populateDropdown("dd-driver",  drivers.map(function (u) { return { id: u.id, label: ((u.firstName || "") + " " + (u.lastName || "")).trim() }; }).filter(function (u) { return u.label; }), "All drivers");
      applyPrefilterVehicle();
      fetchDrilldownData(widgetDef, fromDate, toDate);
    }
    apiCall("Get", { typeName: "Device", search: {} }, function (d) { devices = d; onDone(); }, function () { onDone(); });
    apiCall("Get", { typeName: "User",   search: {} }, function (u) { drivers = u; onDone(); }, function () { onDone(); });

  } else if (type === "carbon-monthly") {
    if (!pf) { fromDate = new Date(); fromDate.setFullYear(fromDate.getFullYear() - 1); fromDate.setDate(1); }
    fields.innerHTML =
      filterFieldHtml("dd-from",    "From",    "date") +
      filterFieldHtml("dd-to",      "To",      "date") +
      filterFieldHtml("dd-vehicle", "Vehicle", "select");
    document.getElementById("dd-from").value = fmtDateInput(fromDate);
    document.getElementById("dd-to").value   = fmtDateInput(toDate);
    apiCall("Get", { typeName: "Device", search: {} }, function (devices) {
      populateDropdown("dd-vehicle", devices.map(function (d) { return { id: d.id, label: d.name }; }), "All vehicles");
      applyPrefilterVehicle();
      fetchDrilldownData(widgetDef, fromDate, toDate);
    }, function () { fetchDrilldownData(widgetDef, fromDate, toDate); });

  } else if (type === "speeding") {
    var thresh = (widgetDef.params && widgetDef.params.thresholdMph) || 80;
    fields.innerHTML =
      filterFieldHtml("dd-from",    "From",    "date") +
      filterFieldHtml("dd-to",      "To",      "date") +
      filterFieldHtml("dd-vehicle", "Vehicle", "select") +
      filterFieldHtml("dd-driver",  "Driver",  "select") +
      "<div class='filter-group'><label class='filter-label'>Threshold</label>" +
      "<input type='number' id='dd-speed-thresh' class='input-date' style='width:70px' value='" + thresh + "' min='1' max='200' /> mph</div>";
    document.getElementById("dd-from").value = fmtDateInput(fromDate);
    document.getElementById("dd-to").value   = fmtDateInput(toDate);
    var p2 = 2, d2 = [], u2 = [];
    function onDone2() {
      if (--p2 > 0) return;
      populateDropdown("dd-vehicle", d2.map(function (d) { return { id: d.id, label: d.name }; }), "All vehicles");
      populateDropdown("dd-driver",  u2.map(function (u) { return { id: u.id, label: ((u.firstName || "") + " " + (u.lastName || "")).trim() }; }).filter(function (u) { return u.label; }), "All drivers");
      applyPrefilterVehicle();
      fetchDrilldownData(widgetDef, fromDate, toDate);
    }
    apiCall("Get", { typeName: "Device", search: {} }, function (d) { d2 = d; onDone2(); }, function () { onDone2(); });
    apiCall("Get", { typeName: "User",   search: {} }, function (u) { u2 = u; onDone2(); }, function () { onDone2(); });

  } else if (type === "maintenance-upcoming") {
    fields.innerHTML =
      filterFieldHtml("dd-vehicle",  "Vehicle", "select") +
      "<div class='filter-group'><label class='filter-label'>Urgency</label>" +
      "<select id='dd-urgency' class='select'><option value=''>All</option><option value='overdue'>Overdue</option><option value='7'>Due in 7 days</option><option value='30'>Due in 30 days</option></select></div>";
    apiCall("Get", { typeName: "Device", search: {} }, function (devices) {
      populateDropdown("dd-vehicle", devices.map(function (d) { return { id: d.id, label: d.name }; }), "All vehicles");
      fetchDrilldownData(widgetDef, null, null);
    }, function () { fetchDrilldownData(widgetDef, null, null); });

  } else if (type === "maintenance-spend") {
    var yearStart = new Date(new Date().getFullYear(), 0, 1);
    fields.innerHTML =
      filterFieldHtml("dd-from",    "From",    "date") +
      filterFieldHtml("dd-to",      "To",      "date") +
      filterFieldHtml("dd-vehicle", "Vehicle", "select");
    document.getElementById("dd-from").value = fmtDateInput(yearStart);
    document.getElementById("dd-to").value   = fmtDateInput(toDate);
    apiCall("Get", { typeName: "Device", search: {} }, function (devices) {
      populateDropdown("dd-vehicle", devices.map(function (d) { return { id: d.id, label: d.name }; }), "All vehicles");
      fetchDrilldownData(widgetDef, yearStart, toDate);
    }, function () { fetchDrilldownData(widgetDef, yearStart, toDate); });

  } else if (type === "fuel-economy-daily") {
    fields.innerHTML =
      filterFieldHtml("dd-from",    "From",    "date") +
      filterFieldHtml("dd-to",      "To",      "date") +
      filterFieldHtml("dd-vehicle", "Vehicle", "select");
    document.getElementById("dd-from").value = fmtDateInput(fromDate);
    document.getElementById("dd-to").value   = fmtDateInput(toDate);
    apiCall("Get", { typeName: "Device", search: {} }, function (devices) {
      populateDropdown("dd-vehicle", devices.map(function (d) { return { id: d.id, label: d.name }; }), "All vehicles");
      applyPrefilterVehicle();
      fetchDrilldownData(widgetDef, fromDate, toDate);
    }, function () { fetchDrilldownData(widgetDef, fromDate, toDate); });
  }
}

function applyDrilldownFilters() {
  var type = state.drilldown.widgetType;
  var fromEl = document.getElementById("dd-from");
  var toEl   = document.getElementById("dd-to");
  var from   = fromEl ? fromEl.value : null;
  var to     = toEl   ? toEl.value   : null;

  document.getElementById("dd-table-wrap").innerHTML = "<p class='placeholder'>Loading...</p>";
  document.getElementById("dd-summary").classList.add("hidden");
  document.getElementById("dd-export-excel").disabled = true;
  document.getElementById("dd-export-pdf").disabled   = true;

  if (type === "maintenance-upcoming") {
    renderDrilldownTable(state.drilldown.data);
    return;
  }
  if (!from || !to) { alert("Please select a date range."); return; }
  fetchDrilldownData(state.drilldown.widgetDef, new Date(from + "T00:00:00"), new Date(to + "T23:59:59"));
}
// ─── Drill-down data fetchers ─────────────────────────────────────────────────
function fetchDrilldownData(widgetDef, fromDate, toDate) {
  var type = widgetDef.type || "exception";
  if (type === "exception") {
    apiCall("Get", { typeName: "ExceptionEvent", search: { fromDate: fromDate.toISOString(), toDate: toDate.toISOString(), ruleSearch: { id: widgetDef.suggestion.rule.id }, includeInvalidated: false } },
      function (events) { state.drilldown.data = events; renderDrilldownTable(events); },
      function (err) { drilldownError(err); });

  } else if (type === "carbon-monthly") {
    apiCall("Get", { typeName: "Trip", search: { fromDate: fromDate.toISOString(), toDate: toDate.toISOString() } }, function (trips) {
      var fuelTrips = trips.filter(function (t) { return (t.fuelUsed || 0) > 0; });
      state.drilldown.data = fuelTrips;
      renderDrilldownTable(fuelTrips);
    }, function (err) { drilldownError(err); });

  } else if (type === "speeding") {
    var thEl  = document.getElementById("dd-speed-thresh");
    var thMph = thEl ? (parseFloat(thEl.value) || 80) : ((widgetDef.params && widgetDef.params.thresholdMph) || 80);
    var thKmh = thMph / 0.621371;
    apiCall("Get", { typeName: "Trip", search: { fromDate: fromDate.toISOString(), toDate: toDate.toISOString() } }, function (trips) {
      var speeding = trips.filter(function (t) { return (t.maximumSpeed || 0) > thKmh; });
      state.drilldown.meta.thresholdMph = thMph;
      state.drilldown.data = speeding;
      renderDrilldownTable(speeding);
    }, function (err) { drilldownError(err); });

  } else if (type === "maintenance-upcoming") {
    apiCall("Get", { typeName: "MaintenanceReminder", search: {} }, function (reminders) {
      state.drilldown.data = reminders;
      renderDrilldownTable(reminders);
    }, function () {
      document.getElementById("dd-table-wrap").innerHTML = "<p class='dd-empty' style='color:var(--warning)'>Maintenance reminders not configured in this database.</p>";
    });

  } else if (type === "maintenance-spend") {
    apiCall("Get", { typeName: "MaintenanceReminder", search: {} }, function (reminders) {
      state.drilldown.data = reminders;
      renderDrilldownTable(reminders);
    }, function () {
      document.getElementById("dd-table-wrap").innerHTML = "<p class='dd-empty' style='color:var(--warning)'>Maintenance reminders not configured in this database.</p>";
    });

  } else if (type === "fuel-economy-daily") {
    apiCall("Get", { typeName: "Trip", search: { fromDate: fromDate.toISOString(), toDate: toDate.toISOString() } }, function (trips) {
      var fuelTrips = trips.filter(function (t) { return (t.fuelUsed || 0) > 0; });
      state.drilldown.data = fuelTrips;
      renderDrilldownTable(fuelTrips);
    }, function (err) { drilldownError(err); });
  }
}

function drilldownError(err) {
  document.getElementById("dd-table-wrap").innerHTML = "<p class='dd-empty' style='color:var(--critical)'>Error: " + esc((err && err.message) || String(err)) + "</p>";
}

// ─── Drill-down table renderers ───────────────────────────────────────────────
function renderDrilldownTable(rawData) {
  var type    = state.drilldown.widgetType;
  var vidSel  = document.getElementById("dd-vehicle");
  var drvSel  = document.getElementById("dd-driver");
  var urgSel  = document.getElementById("dd-urgency");
  var vidFilter = vidSel  ? vidSel.value  : "";
  var drvFilter = drvSel  ? drvSel.value  : "";
  var urgFilter = urgSel  ? urgSel.value  : "";
  var today = new Date(); today.setHours(0, 0, 0, 0);

  var filtered, headers, rows, summaryItems;

  if (type === "exception") {
    filtered = rawData.filter(function (e) {
      if (vidFilter && (!e.device || e.device.id !== vidFilter)) return false;
      if (drvFilter && (!e.driver || e.driver.id !== drvFilter))  return false;
      return true;
    });
    state.drilldown.filtered = filtered;
    var totalVeh = {}, totalMins = 0;
    filtered.forEach(function (e) { if (e.device) totalVeh[e.device.id] = 1; totalMins += durationMins(e.activeFrom, e.activeTo); });
    summaryItems = [ddSummaryItem(filtered.length, "Total Violations"), ddSummaryItem(Object.keys(totalVeh).length, "Vehicles Involved"), ddSummaryItem(fmtMins(totalMins), "Total Duration")];
    headers = ["Vehicle", "Driver", "Date", "Time", "Duration", ""];
    rows = filtered.slice().sort(function (a, b) { return new Date(b.activeFrom) - new Date(a.activeFrom); }).map(function (e) {
      var vehicle = state.deviceMap[e.device && e.device.id] || "Unknown";
      var driver  = e.driver ? (state.driverMap[e.driver.id] || "Unknown driver") : "No driver";
      var replay  = e.id ? "<button class='btn btn-sm btn-secondary inc-btn' data-event-id='" + esc(e.id) + "'>Replay</button>" : "";
      return "<tr><td class='td-vehicle'>" + esc(vehicle) + "</td><td class='td-driver'>" + esc(driver) + "</td><td>" + fmtDateReadable(e.activeFrom) + "</td><td>" + fmtTime(e.activeFrom) + "</td><td class='td-dur'>" + fmtMins(durationMins(e.activeFrom, e.activeTo)) + "</td><td class='td-action'>" + replay + "</td></tr>";
    });

  } else if (type === "carbon-monthly") {
    filtered = vidFilter ? rawData.filter(function (t) { return t.device && t.device.id === vidFilter; }) : rawData;
    state.drilldown.filtered = filtered;
    var totalCo2 = 0, totalFuel2 = 0, totalDist2 = 0;
    filtered.forEach(function (t) { totalCo2 += (t.fuelUsed || 0) * 2.4; totalFuel2 += (t.fuelUsed || 0); totalDist2 += (t.distance || 0); });
    summaryItems = [ddSummaryItem(Math.round(totalCo2) + " kg", "Total CO\u2082"), ddSummaryItem(calcMpg(totalDist2, totalFuel2), "Fleet MPG"), ddSummaryItem((totalDist2 * 0.621371).toFixed(0) + " mi", "Total Distance")];
    headers = ["Vehicle", "Date", "Distance", "Fuel Used", "MPG", "CO\u2082"];
    rows = filtered.slice().sort(function (a, b) { return new Date(b.start) - new Date(a.start); }).map(function (t) {
      var vName = state.deviceMap[t.device && t.device.id] || "Unknown";
      return "<tr><td class='td-vehicle'>" + esc(vName) + "</td><td>" + fmtDateReadable(t.start) + "</td><td>" + ((t.distance || 0) * 0.621371).toFixed(1) + " mi</td><td>" + (t.fuelUsed || 0).toFixed(1) + " L</td><td>" + calcMpg(t.distance || 0, t.fuelUsed || 0) + "</td><td>" + Math.round((t.fuelUsed || 0) * 2.4) + " kg</td></tr>";
    });

  } else if (type === "speeding") {
    var usedThresh = state.drilldown.meta.thresholdMph || ((state.drilldown.widgetDef && state.drilldown.widgetDef.params && state.drilldown.widgetDef.params.thresholdMph) || 80);
    filtered = rawData.filter(function (t) {
      if (vidFilter && (!t.device || t.device.id !== vidFilter)) return false;
      return true;
    });
    state.drilldown.filtered = filtered;
    var vehInvolved = {}, maxSpd = 0;
    filtered.forEach(function (t) { if (t.device) vehInvolved[t.device.id] = 1; if ((t.maximumSpeed || 0) > maxSpd) maxSpd = t.maximumSpeed; });
    summaryItems = [ddSummaryItem(filtered.length, "Total Incidents"), ddSummaryItem(Object.keys(vehInvolved).length, "Vehicles Involved"), ddSummaryItem(toMph(maxSpd).toFixed(0) + " mph", "Highest Speed")];
    headers = ["Vehicle", "Driver", "Date", "Start Time", "Speed (&gt;" + usedThresh + "mph)", "Distance"];
    rows = filtered.slice().sort(function (a, b) { return (b.maximumSpeed || 0) - (a.maximumSpeed || 0); }).map(function (t) {
      var vName = state.deviceMap[t.device && t.device.id] || "Unknown";
      return "<tr><td class='td-vehicle'>" + esc(vName) + "</td><td class='td-driver'>" + esc(t.driverName || "Unassigned") + "</td><td>" + fmtDateReadable(t.start) + "</td><td>" + fmtTime(t.start) + "</td><td><span class='speed-badge'>" + toMph(t.maximumSpeed || 0).toFixed(0) + " mph</span></td><td>" + ((t.distance || 0) * 0.621371).toFixed(1) + " mi</td></tr>";
    });

  } else if (type === "maintenance-upcoming") {
    var processed = rawData.map(function (r) {
      var dueDate = r.nextServiceDate ? new Date(r.nextServiceDate) : null;
      var days    = dueDate ? Math.round((dueDate - today) / 86400000) : null;
      return { vehicle: state.deviceMap[r.device && r.device.id] || "Unknown", what: r.comment || r.description || "Maintenance", dueDate: dueDate, dueDateStr: dueDate ? fmtDateReadable(dueDate.toISOString()) : "—", odometer: r.nextOdometerReading ? (r.nextOdometerReading / 1000).toFixed(0) + " km" : "—", days: days, deviceId: r.device && r.device.id };
    }).filter(function (r) { return r.dueDate; });
    if (vidFilter) processed = processed.filter(function (r) { return r.deviceId === vidFilter; });
    if (urgFilter === "overdue")  processed = processed.filter(function (r) { return r.days < 0; });
    else if (urgFilter === "7")   processed = processed.filter(function (r) { return r.days >= 0 && r.days <= 7; });
    else if (urgFilter === "30")  processed = processed.filter(function (r) { return r.days >= 0 && r.days <= 30; });
    state.drilldown.filtered = processed;
    summaryItems = [ddSummaryItem(processed.filter(function (r) { return r.days < 0; }).length, "Overdue"), ddSummaryItem(processed.filter(function (r) { return r.days >= 0 && r.days <= 7; }).length, "Due 7 Days"), ddSummaryItem(processed.filter(function (r) { return r.days >= 0 && r.days <= 30; }).length, "Due 30 Days")];
    headers = ["Vehicle", "What Is Due", "Date Due", "Odometer", "Status"];
    rows = processed.slice().sort(function (a, b) { return a.days - b.days; }).map(function (r) {
      var sHtml = r.days < 0 ? "<span class='badge badge-critical'>Overdue " + Math.abs(r.days) + "d</span>" : r.days <= 7 ? "<span class='badge badge-warning'>Due in " + r.days + "d</span>" : "<span class='badge badge-low'>Due in " + r.days + "d</span>";
      return "<tr><td class='td-vehicle'>" + esc(r.vehicle) + "</td><td>" + esc(r.what) + "</td><td>" + r.dueDateStr + "</td><td>" + r.odometer + "</td><td>" + sHtml + "</td></tr>";
    });

  } else if (type === "maintenance-spend") {
    var fromEl = document.getElementById("dd-from"), toEl = document.getElementById("dd-to");
    var spendFrom = fromEl ? new Date(fromEl.value + "T00:00:00") : new Date(new Date().getFullYear(), 0, 1);
    var spendTo   = toEl   ? new Date(toEl.value   + "T23:59:59") : new Date();
    var spentRows = rawData.map(function (r) {
      var d = r.nextServiceDate ? new Date(r.nextServiceDate) : null;
      return { vehicle: state.deviceMap[r.device && r.device.id] || "Unknown", what: r.comment || r.description || "Maintenance", date: d, dateStr: d ? fmtDateReadable(d.toISOString()) : "—", cost: r.cost != null ? r.cost : null, deviceId: r.device && r.device.id };
    }).filter(function (r) { return r.date && r.date >= spendFrom && r.date <= spendTo; });
    if (vidFilter) spentRows = spentRows.filter(function (r) { return r.deviceId === vidFilter; });
    state.drilldown.filtered = spentRows;
    var totalCost2 = spentRows.filter(function (r) { return r.cost != null; }).reduce(function (s, r) { return s + r.cost; }, 0);
    var hasCost2   = spentRows.some(function (r) { return r.cost != null; });
    summaryItems = [ddSummaryItem(spentRows.length, "Total Events"), ddSummaryItem(hasCost2 ? "\u00a3" + totalCost2.toFixed(2) : "—", "Total Spend"), ddSummaryItem([... new Set(spentRows.map(function (r) { return r.vehicle; }))].length, "Vehicles")];
    headers = ["Vehicle", "Description", "Date", "Cost", "Status"];
    rows = spentRows.slice().sort(function (a, b) { return b.date - a.date; }).map(function (r) {
      var status = r.date < today ? "<span class='badge badge-warning'>Overdue/Past</span>" : "<span class='badge badge-low'>Upcoming</span>";
      return "<tr><td class='td-vehicle'>" + esc(r.vehicle) + "</td><td>" + esc(r.what) + "</td><td>" + r.dateStr + "</td><td>" + (r.cost != null ? "\u00a3" + r.cost.toFixed(2) : "<span style='color:var(--text-muted)'>Not logged</span>") + "</td><td>" + status + "</td></tr>";
    });

  } else if (type === "fuel-economy-daily") {
    filtered = vidFilter ? rawData.filter(function (t) { return t.device && t.device.id === vidFilter; }) : rawData;
    state.drilldown.filtered = filtered;
    var tFuel = filtered.reduce(function (s, t) { return s + (t.fuelUsed || 0); }, 0);
    var tDist = filtered.reduce(function (s, t) { return s + (t.distance || 0); }, 0);
    var mpgVals = filtered.map(function (t) { return parseFloat(calcMpg(t.distance, t.fuelUsed)) || 0; }).filter(function (v) { return v > 0; });
    summaryItems = [ddSummaryItem(calcMpg(tDist, tFuel), "Fleet MPG"), ddSummaryItem(tFuel.toFixed(1) + " L", "Total Fuel"), ddSummaryItem((tDist * 0.621371).toFixed(0) + " mi", "Total Distance")];
    headers = ["Vehicle", "Date", "Start Time", "Distance", "Fuel Used", "MPG"];
    rows = filtered.slice().sort(function (a, b) { return new Date(b.start) - new Date(a.start); }).map(function (t) {
      var vName = state.deviceMap[t.device && t.device.id] || "Unknown";
      return "<tr><td class='td-vehicle'>" + esc(vName) + "</td><td>" + fmtDateReadable(t.start) + "</td><td>" + fmtTime(t.start) + "</td><td>" + ((t.distance || 0) * 0.621371).toFixed(1) + " mi</td><td>" + (t.fuelUsed || 0).toFixed(1) + " L</td><td><strong>" + calcMpg(t.distance, t.fuelUsed) + "</strong></td></tr>";
    });
  } else {
    return;
  }

  // Render summary
  var summaryEl = document.getElementById("dd-summary");
  summaryEl.innerHTML = summaryItems.join("");
  summaryEl.classList.remove("hidden");

  // Show the Table/Map toggle only for location-capable types (event has device + time)
  var mapCapable = (type === "exception" || type === "speeding");
  var toggle = document.getElementById("dd-view-toggle");
  if (mapCapable && rows && rows.length) { toggle.classList.remove("hidden"); }
  else { toggle.classList.add("hidden"); }
  switchDrilldownView("table"); // always reset to table on new data

  // Render table
  if (!rows || !rows.length) {
    document.getElementById("dd-table-wrap").innerHTML = "<p class='dd-empty'>No data found for the selected filters.</p>";
    return;
  }
  document.getElementById("dd-table-wrap").innerHTML =
    "<table class='dd-table'><thead><tr>" +
    headers.map(function (h) { return "<th>" + h + "</th>"; }).join("") +
    "</tr></thead><tbody>" + rows.join("") + "</tbody></table>";

  // Wire per-row Replay buttons (exception type only)
  if (type === "exception") {
    document.querySelectorAll("#dd-table-wrap .inc-btn").forEach(function (btn) {
      btn.addEventListener("click", function (e) {
        e.stopPropagation();
        var id = btn.getAttribute("data-event-id");
        var ev = (state.drilldown.filtered || []).find(function (x) { return String(x.id) === String(id); });
        if (ev) openIncident(ev);
      });
    });
  }

  document.getElementById("dd-export-excel").disabled = false;
  document.getElementById("dd-export-pdf").disabled   = false;
}

function ddSummaryItem(value, label) {
  return "<div class='dd-summary-item'><div class='dd-summary-label'>" + label + "</div><div class='dd-summary-value'>" + value + "</div></div>";
}

function populateDropdown(id, items, allLabel) {
  var sel = document.getElementById(id); if (!sel) return;
  sel.innerHTML = "<option value=''>" + allLabel + "</option>";
  items.forEach(function (item) {
    var opt = document.createElement("option");
    opt.value = item.id; opt.textContent = item.label;
    sel.appendChild(opt);
  });
}
// ─── Excel export ─────────────────────────────────────────────────────────────
function exportDrilldownExcel() {
  var type    = state.drilldown.widgetType;
  var data    = state.drilldown.filtered;
  if (!data || !data.length) return;
  var from = document.getElementById("dd-from")  ? document.getElementById("dd-from").value  : "";
  var to   = document.getElementById("dd-to")    ? document.getElementById("dd-to").value    : "";
  var typeDef  = PRESET_WIDGET_TYPES[type];
  var reportName = typeDef ? typeDef.label : (state.drilldown.widgetDef && state.drilldown.widgetDef.suggestion ? state.drilldown.widgetDef.suggestion.rule.name : "Report");

  var summaryRows = [
    ["Smart Insights \u2014 " + reportName],
    ["Fleet:", state.dbName || ""],
    ["Period:", from && to ? from + " to " + to : "All available"],
    ["Generated:", fmtDateReadable(new Date().toISOString())],
    [],
    ["SUMMARY"]
  ];

  var detailRows;
  if (type === "exception") {
    var totalVeh = {}, totalMins = 0;
    data.forEach(function (e) { if (e.device) totalVeh[e.device.id] = 1; totalMins += durationMins(e.activeFrom, e.activeTo); });
    summaryRows.push(["Total Violations", data.length], ["Vehicles Involved", Object.keys(totalVeh).length], ["Total Duration", fmtMins(totalMins)]);
    detailRows = [["Vehicle","Driver","Date","Time","Duration"]];
    data.forEach(function (e) { detailRows.push([state.deviceMap[e.device && e.device.id] || "Unknown", e.driver ? "Unknown" : "No driver", fmtDateReadable(e.activeFrom), fmtTime(e.activeFrom), fmtMins(durationMins(e.activeFrom, e.activeTo))]); });

  } else if (type === "carbon-monthly") {
    var co2Tot = data.reduce(function (s, t) { return s + (t.fuelUsed||0)*2.4; }, 0);
    var fTot   = data.reduce(function (s, t) { return s + (t.fuelUsed||0); }, 0);
    var dTot   = data.reduce(function (s, t) { return s + (t.distance||0); }, 0);
    summaryRows.push(["Total CO2 (kg)", Math.round(co2Tot)], ["Fleet MPG", calcMpg(dTot, fTot)], ["Total Distance (mi)", (dTot*0.621371).toFixed(1)]);
    detailRows = [["Vehicle","Date","Distance (mi)","Fuel (L)","MPG","CO2 (kg)"]];
    data.forEach(function (t) { detailRows.push([state.deviceMap[t.device && t.device.id] || "Unknown", fmtDateReadable(t.start), ((t.distance||0)*0.621371).toFixed(1), (t.fuelUsed||0).toFixed(1), calcMpg(t.distance, t.fuelUsed), Math.round((t.fuelUsed||0)*2.4)]); });

  } else if (type === "speeding") {
    var maxS = data.reduce(function (m, t) { return (t.maximumSpeed||0) > m ? t.maximumSpeed : m; }, 0);
    summaryRows.push(["Total Incidents", data.length], ["Highest Speed (mph)", toMph(maxS).toFixed(0)]);
    detailRows = [["Vehicle","Driver","Date","Start Time","Max Speed (mph)","Distance (mi)"]];
    data.forEach(function (t) { detailRows.push([state.deviceMap[t.device && t.device.id] || "Unknown", t.driverName || "Unassigned", fmtDateReadable(t.start), fmtTime(t.start), toMph(t.maximumSpeed||0).toFixed(0), ((t.distance||0)*0.621371).toFixed(1)]); });

  } else if (type === "maintenance-upcoming") {
    summaryRows.push(["Total Items", data.length], ["Overdue", data.filter(function (r) { return r.days < 0; }).length]);
    detailRows = [["Vehicle","What Is Due","Date Due","Odometer","Days Until Due"]];
    data.forEach(function (r) { detailRows.push([r.vehicle, r.what, r.dueDateStr, r.odometer, r.days]); });

  } else if (type === "maintenance-spend") {
    var hasCost3 = data.some(function (r) { return r.cost != null; });
    var costTot3 = data.filter(function (r) { return r.cost != null; }).reduce(function (s, r) { return s + r.cost; }, 0);
    summaryRows.push(["Total Events", data.length], ["Total Spend", hasCost3 ? "\u00a3" + costTot3.toFixed(2) : "Not logged"]);
    detailRows = [["Vehicle","Description","Date","Cost","Status"]];
    data.forEach(function (r) { detailRows.push([r.vehicle, r.what, r.dateStr, r.cost != null ? r.cost.toFixed(2) : "", r.status || ""]); });

  } else if (type === "fuel-economy-daily") {
    var fT = data.reduce(function (s, t) { return s + (t.fuelUsed||0); }, 0);
    var dT = data.reduce(function (s, t) { return s + (t.distance||0); }, 0);
    summaryRows.push(["Fleet MPG", calcMpg(dT, fT)], ["Total Fuel (L)", fT.toFixed(1)], ["Total Distance (mi)", (dT*0.621371).toFixed(1)]);
    detailRows = [["Vehicle","Date","Start Time","Distance (mi)","Fuel (L)","MPG"]];
    data.forEach(function (t) { detailRows.push([state.deviceMap[t.device && t.device.id] || "Unknown", fmtDateReadable(t.start), fmtTime(t.start), ((t.distance||0)*0.621371).toFixed(1), (t.fuelUsed||0).toFixed(1), calcMpg(t.distance, t.fuelUsed)]); });
  } else { return; }

  var wb = XLSX.utils.book_new();
  var wsSummary = XLSX.utils.aoa_to_sheet(summaryRows);
  var wsDetail  = XLSX.utils.aoa_to_sheet(detailRows);
  wsDetail["!cols"] = detailRows[0].map(function () { return { wch: 20 }; });
  XLSX.utils.book_append_sheet(wb, wsSummary, "Summary");
  XLSX.utils.book_append_sheet(wb, wsDetail,  "Data");
  XLSX.writeFile(wb, sanitiseFilename(reportName) + "_report.xlsx");
}

// ─── PDF branding helper ───────────────────────────────────────────────────────
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

// ─── PDF export ───────────────────────────────────────────────────────────────
function exportDrilldownPdf() {
  var type     = state.drilldown.widgetType;
  var data     = state.drilldown.filtered;
  if (!data || !data.length) return;
  var from = document.getElementById("dd-from") ? document.getElementById("dd-from").value : "";
  var to   = document.getElementById("dd-to")   ? document.getElementById("dd-to").value   : "";
  var typeDef  = PRESET_WIDGET_TYPES[type];
  var reportName = typeDef ? typeDef.label : (state.drilldown.widgetDef && state.drilldown.widgetDef.suggestion ? state.drilldown.widgetDef.suggestion.rule.name : "Report");

  var doc = new jspdf.jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
  var pageW = doc.internal.pageSize.getWidth(), margin = 16;

  doc.setFillColor(GEOTAB_NAVY[0], GEOTAB_NAVY[1], GEOTAB_NAVY[2]);
  doc.rect(0, 0, pageW, 28, "F");
  drawPdfHeaderLogo(doc, pageW, margin, 28);
  doc.setTextColor(255, 255, 255); doc.setFontSize(14); doc.setFont("helvetica", "bold");
  doc.text(reportName, margin, 12);
  doc.setFontSize(9); doc.setFont("helvetica", "normal");
  doc.text("Generated " + fmtDateReadable(new Date().toISOString()), margin, 20);

  var yPos = 36;
  doc.setTextColor(60, 60, 60); doc.setFontSize(9);
  if (state.dbName) { doc.setFont("helvetica", "bold"); doc.text("Fleet:", margin, yPos); doc.setFont("helvetica", "normal"); doc.text(state.dbName, margin + 14, yPos); yPos += 6; }
  if (from && to)   { doc.setFont("helvetica", "bold"); doc.text("Period:", margin, yPos); doc.setFont("helvetica", "normal"); doc.text(from + " \u2014 " + to, margin + 16, yPos); yPos += 10; }
  else { yPos += 4; }

  // Summary boxes
  var boxes;
  if (type === "exception") {
    var tV = {}, tM = 0; data.forEach(function (e) { if (e.device) tV[e.device.id]=1; tM+=durationMins(e.activeFrom,e.activeTo); });
    boxes = [{ label:"Total Violations",value:String(data.length) },{ label:"Vehicles Involved",value:String(Object.keys(tV).length) },{ label:"Total Duration",value:fmtMins(tM) }];
  } else if (type === "carbon-monthly") {
    var co2B = data.reduce(function(s,t){return s+(t.fuelUsed||0)*2.4;},0);
    var fB   = data.reduce(function(s,t){return s+(t.fuelUsed||0);},0);
    var dB   = data.reduce(function(s,t){return s+(t.distance||0);},0);
    boxes = [{ label:"Total CO2",value:Math.round(co2B)+" kg" },{ label:"Fleet MPG",value:calcMpg(dB,fB) },{ label:"Distance",value:(dB*0.621371).toFixed(0)+" mi" }];
  } else if (type === "speeding") {
    var mS = data.reduce(function(m,t){return(t.maximumSpeed||0)>m?t.maximumSpeed:m;},0);
    var vB = {}; data.forEach(function(t){if(t.device)vB[t.device.id]=1;});
    boxes = [{ label:"Total Incidents",value:String(data.length) },{ label:"Vehicles",value:String(Object.keys(vB).length) },{ label:"Highest Speed",value:toMph(mS).toFixed(0)+" mph" }];
  } else if (type === "maintenance-upcoming") {
    boxes = [{ label:"Overdue",value:String(data.filter(function(r){return r.days<0;}).length) },{ label:"Due 7 Days",value:String(data.filter(function(r){return r.days>=0&&r.days<=7;}).length) },{ label:"Due 30 Days",value:String(data.filter(function(r){return r.days>=0&&r.days<=30;}).length) }];
  } else if (type === "maintenance-spend") {
    var cT = data.filter(function(r){return r.cost!=null;}).reduce(function(s,r){return s+r.cost;},0);
    boxes = [{ label:"Total Events",value:String(data.length) },{ label:"Total Spend",value:data.some(function(r){return r.cost!=null;})? "\u00a3"+cT.toFixed(2):"Not logged" },{ label:"Vehicles",value:String([...new Set(data.map(function(r){return r.vehicle;}))].length) }];
  } else {
    var fP=data.reduce(function(s,t){return s+(t.fuelUsed||0);},0), dP=data.reduce(function(s,t){return s+(t.distance||0);},0);
    boxes = [{ label:"Fleet MPG",value:calcMpg(dP,fP) },{ label:"Total Fuel",value:fP.toFixed(1)+" L" },{ label:"Distance",value:(dP*0.621371).toFixed(0)+" mi" }];
  }
  var boxW = (pageW - margin*2 - 8) / 3;
  boxes.forEach(function (b, i) {
    var bx = margin + i * (boxW + 4);
    doc.setFillColor(232, 244, 253); doc.roundedRect(bx, yPos, boxW, 18, 2, 2, "F");
    doc.setTextColor(0, 90, 158); doc.setFontSize(16); doc.setFont("helvetica", "bold");
    doc.text(b.value, bx + boxW/2, yPos+10, { align: "center" });
    doc.setFontSize(7); doc.setFont("helvetica", "normal"); doc.setTextColor(100,100,100);
    doc.text(b.label, bx + boxW/2, yPos+15, { align: "center" });
  });
  yPos += 26;

  // Table
  var tableHead, tableBody;
  if (type === "exception") {
    tableHead = [["Vehicle","Driver","Date","Time","Duration"]];
    tableBody = data.map(function(e){return[state.deviceMap[e.device&&e.device.id]||"Unknown",e.driver?"Unknown":"No driver",fmtDateReadable(e.activeFrom),fmtTime(e.activeFrom),fmtMins(durationMins(e.activeFrom,e.activeTo))];});
  } else if (type === "carbon-monthly") {
    tableHead = [["Vehicle","Date","Dist (mi)","Fuel (L)","MPG","CO2 (kg)"]];
    tableBody = data.map(function(t){return[state.deviceMap[t.device&&t.device.id]||"Unknown",fmtDateReadable(t.start),((t.distance||0)*0.621371).toFixed(1),(t.fuelUsed||0).toFixed(1),calcMpg(t.distance,t.fuelUsed),Math.round((t.fuelUsed||0)*2.4)];});
  } else if (type === "speeding") {
    tableHead = [["Vehicle","Driver","Date","Time","Max Speed","Distance"]];
    tableBody = data.map(function(t){return[state.deviceMap[t.device&&t.device.id]||"Unknown",t.driverName||"Unassigned",fmtDateReadable(t.start),fmtTime(t.start),toMph(t.maximumSpeed||0).toFixed(0)+" mph",((t.distance||0)*0.621371).toFixed(1)+" mi"];});
  } else if (type === "maintenance-upcoming") {
    tableHead = [["Vehicle","What Is Due","Date Due","Odometer","Days Until"]];
    tableBody = data.map(function(r){return[r.vehicle,r.what,r.dueDateStr,r.odometer,r.days!=null?String(r.days):""];});
  } else if (type === "maintenance-spend") {
    tableHead = [["Vehicle","Description","Date","Cost","Status"]];
    tableBody = data.map(function(r){return[r.vehicle,r.what,r.dateStr,r.cost!=null?"\u00a3"+r.cost.toFixed(2):"Not logged",r.status||""];});
  } else {
    tableHead = [["Vehicle","Date","Time","Distance","Fuel","MPG"]];
    tableBody = data.map(function(t){return[state.deviceMap[t.device&&t.device.id]||"Unknown",fmtDateReadable(t.start),fmtTime(t.start),((t.distance||0)*0.621371).toFixed(1)+" mi",(t.fuelUsed||0).toFixed(1)+" L",calcMpg(t.distance,t.fuelUsed)];});
  }

  doc.autoTable({ startY: yPos, head: tableHead, body: tableBody, margin: { left: margin, right: margin }, styles: { fontSize: 9, cellPadding: 3, textColor: [45,55,72] }, headStyles: { fillColor: [0,120,212], textColor: 255, fontStyle: "bold", fontSize: 9 }, alternateRowStyles: { fillColor: [248,249,251] } });
  var pageCount = doc.internal.getNumberOfPages();
  for (var p = 1; p <= pageCount; p++) {
    doc.setPage(p); doc.setFontSize(8); doc.setTextColor(160,160,160);
    doc.text("Confidential", margin, doc.internal.pageSize.getHeight()-8);
    doc.text("Page "+p+" of "+pageCount, pageW-margin, doc.internal.pageSize.getHeight()-8, { align: "right" });
  }
  doc.save(sanitiseFilename(reportName) + "_report.pdf");
}

// ─── Legacy Trip History PDF export ────────────────────────────────────────────
// Mirrors the structure of the reference "Daily Report" export: header, a dark
// "Report Totals for" bar, a 2x3 KPI grid, then one table per vehicle per day
// with a Starting-from lead-in row, an Ignition-On row, trip rows, and a daily
// total row. Uses Smart Insights' own blue branding rather than reproducing a
// third-party vendor's logo/colours — the FORMAT is matched, not the brand.
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

  // Report-wide KPIs across all vehicles
  var allTrips = [];
  Object.keys(byDevice).forEach(function (vid) { allTrips = allTrips.concat(byDevice[vid].trips); });
  var totalDistM = allTrips.reduce(function (s, t) { return s + (t.distance || 0); }, 0);
  var totalDrive = allTrips.reduce(function (s, t) { var dm = parseDurationToMins(t.drivingDuration); return s + (dm != null ? dm : durationMins(t.start, t.stop)); }, 0);
  var totalIdle  = allTrips.reduce(function (s, t) { return s + (parseDurationToMins(t.idlingDuration) || 0); }, 0);
  var totalStop  = allTrips.reduce(function (s, t) { return s + (parseDurationToMins(t.stopDuration) || 0); }, 0);
  var numStops   = allTrips.length;
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
        var stopMins  = parseDurationToMins(t.stopDuration);
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
      var dayStop  = block.rows.reduce(function (s, r) { return s + (parseDurationToMins(r.trip.stopDuration) || 0); }, 0);
      var totalStyles = { fontStyle: "bold", fillColor: [245, 245, 245] };
      body.push([
        { content: block.date + " Total", styles: totalStyles },
        { content: milesFromDistance(dayDistM).toFixed(2) + " mi in " + fmtDurWhole(dayDrive), styles: totalStyles },
        { content: "", styles: totalStyles },
        { content: "", styles: totalStyles },
        { content: fmtDurWhole(dayIdle), styles: totalStyles },
        { content: fmtDurWhole(dayStop), styles: totalStyles },
        { content: block.rows.length + " stops", styles: totalStyles }
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

// ─── Map / GPS correlation ────────────────────────────────────────────────────
// ExceptionEvent/Trip records carry no coordinates, so we correlate each event to
// the vehicle's GPS trail (LogRecord) around the event time. Capped + batched.
function correlateEventLocations(events, cb, onProgress) {
  // Most-recent-first, capped
  var capped = events.slice().sort(function (a, b) {
    return new Date(b.activeFrom) - new Date(a.activeFrom);
  }).slice(0, MAP_EVENT_CAP);
  var truncated = events.length > capped.length;
  var located = [];
  if (!capped.length) { cb(located, truncated); return; }

  var BATCH = 50, idx = 0;

  function runBatch() {
    if (idx >= capped.length) { cb(located, truncated); return; }
    var sliceEvents = capped.slice(idx, idx + BATCH);
    var calls = sliceEvents.map(function (e) {
      var t = new Date(e.activeFrom).getTime();
      return ["Get", { typeName: "LogRecord", search: {
        deviceSearch: { id: e.device.id },
        fromDate: new Date(t - 180000).toISOString(), // -3 min
        toDate:   new Date(t + 180000).toISOString()  // +3 min
      } }];
    });
    state.api.multiCall(calls, function (results) {
      results.forEach(function (recs, i) {
        var e = sliceEvents[i];
        if (!recs || !recs.length) return;
        var t = new Date(e.activeFrom).getTime();
        var best = null;
        recs.forEach(function (r) {
          if (r.latitude == null || r.longitude == null) return;
          if (r.latitude === 0 && r.longitude === 0) return;
          var dd = Math.abs(new Date(r.dateTime).getTime() - t);
          if (!best || dd < best.dd) best = { r: r, dd: dd };
        });
        if (best) {
          var vid = e.device.id;
          located.push({
            lat: best.r.latitude, lng: best.r.longitude,
            vehicleId: vid, vehicleName: state.deviceMap[vid] || vid,
            when: e.activeFrom,
            ruleName: (e.rule && state.ruleNameMap) ? state.ruleNameMap[e.rule.id] : null
          });
        }
      });
      idx += BATCH;
      if (onProgress) onProgress(Math.min(idx, capped.length), capped.length);
      runBatch();
    }, function () {
      // Skip failed batch, keep going
      idx += BATCH;
      if (onProgress) onProgress(Math.min(idx, capped.length), capped.length);
      runBatch();
    });
  }
  runBatch();
}

// Assign a stable colour per vehicle id across a located set
function colourByVehicle(located) {
  var vids = [], map = {};
  located.forEach(function (l) { if (vids.indexOf(l.vehicleId) < 0) vids.push(l.vehicleId); });
  vids.forEach(function (vid, i) { map[vid] = vehicleColor(i); });
  return map;
}

function plotMarkers(leafletMap, layerGroup, located, onVehicleClick) {
  layerGroup.clearLayers();
  if (!located.length) { leafletMap.invalidateSize(); return null; }
  var colourFor = colourByVehicle(located);
  var bounds = [];
  located.forEach(function (l) {
    var m = L.circleMarker([l.lat, l.lng], {
      radius: 6, color: "#fff", weight: 1, fillColor: colourFor[l.vehicleId], fillOpacity: 0.9
    });
    m.bindPopup(
      "<strong>" + esc(l.vehicleName) + "</strong><br>" +
      (l.ruleName ? esc(l.ruleName) + "<br>" : "") +
      fmtDateReadable(l.when) + " " + fmtTime(l.when)
    );
    if (onVehicleClick) m.on("click", function () { onVehicleClick(l); });
    layerGroup.addLayer(m);
    bounds.push([l.lat, l.lng]);
  });
  if (bounds.length) leafletMap.fitBounds(bounds, { padding: [30, 30], maxZoom: 14 });
  setTimeout(function () { leafletMap.invalidateSize(); }, 60);
  return colourFor;
}

// ── Map tab ─────────────────────────────────────────────────────────────────
function setupMap() {
  var today = new Date(), weekAgo = new Date(); weekAgo.setDate(today.getDate() - 7);
  document.getElementById("map-from").value = fmtDateInput(weekAgo);
  document.getElementById("map-to").value   = fmtDateInput(today);

  apiCall("Get", { typeName: "Rule", search: {} }, function (rules) {
    state.ruleNameMap = {};
    var sel = document.getElementById("map-rule");
    rules.slice().sort(function (a, b) { return (a.name || "").localeCompare(b.name || ""); }).forEach(function (r) {
      state.ruleNameMap[r.id] = r.name || r.id;
      var opt = document.createElement("option");
      opt.value = r.id; opt.textContent = r.name || r.id;
      sel.appendChild(opt);
    });
  }, function () {});

  document.getElementById("map-load").addEventListener("click", loadMapData);
  var navBtn = document.querySelector("[data-tab='map']");
  if (navBtn) navBtn.addEventListener("click", function () {
    setTimeout(function () { if (state.mapObj) state.mapObj.invalidateSize(); }, 60);
  });
}

function ensureMap() {
  if (state.mapObj) return state.mapObj;
  state.mapObj = L.map("map-canvas").setView([54.5, -3], 5); // UK default view
  L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    maxZoom: 19, attribution: "&copy; OpenStreetMap contributors"
  }).addTo(state.mapObj);
  state.mapMarkers = L.layerGroup().addTo(state.mapObj);
  return state.mapObj;
}

function loadMapData() {
  var ruleId = document.getElementById("map-rule").value;
  var from   = document.getElementById("map-from").value;
  var to     = document.getElementById("map-to").value;
  if (!from || !to) { alert("Please select a date range."); return; }

  var status = document.getElementById("map-status");
  status.classList.remove("hidden");
  status.textContent = "Loading events\u2026";
  ensureMap();

  var search = {
    fromDate: new Date(from + "T00:00:00").toISOString(),
    toDate:   new Date(to   + "T23:59:59").toISOString(),
    includeInvalidated: false
  };
  if (ruleId) search.ruleSearch = { id: ruleId };

  ensureDeviceMap(function () {
    apiCall("Get", { typeName: "ExceptionEvent", search: search }, function (events) {
      if (!events.length) { status.textContent = "No events found for this range."; plotMarkers(state.mapObj, state.mapMarkers, []); document.getElementById("map-legend").classList.add("hidden"); return; }
      status.textContent = "Locating " + Math.min(events.length, MAP_EVENT_CAP) + " of " + events.length + " events\u2026";
      correlateEventLocations(events, function (located, truncated) {
        // Markers show a popup (vehicle, rule, timestamp) on click.
        var colourFor = plotMarkers(state.mapObj, state.mapMarkers, located, null);
        var msg = "Plotted " + located.length + " location" + (located.length !== 1 ? "s" : "");
        if (truncated) msg += " (capped at " + MAP_EVENT_CAP + " most recent)";
        if (!located.length) msg += " \u2014 no GPS positions were found for these events.";
        status.textContent = msg + ".";
        renderMapLegend(colourFor);
      }, function (done, total) {
        status.textContent = "Locating events\u2026 " + done + "/" + total;
      });
    }, function (err) { status.textContent = "Error loading events: " + esc((err && err.message) || String(err)); });
  });
}

function renderMapLegend(colourFor) {
  var legend = document.getElementById("map-legend");
  if (!colourFor || !Object.keys(colourFor).length) { legend.classList.add("hidden"); return; }
  legend.innerHTML = "<div class='legend-title'>Vehicles</div>" + Object.keys(colourFor).map(function (vid) {
    return "<div class='legend-row'><span class='legend-dot' style='background:" + colourFor[vid] + "'></span>" + esc(state.deviceMap[vid] || vid) + "</div>";
  }).join("");
  legend.classList.remove("hidden");
}

// ── Drilldown mini-map ────────────────────────────────────────────────────────
function switchDrilldownView(view) {
  var tableWrap = document.getElementById("dd-table-wrap");
  var mapWrap   = document.getElementById("dd-map");
  var bt = document.getElementById("dd-view-table");
  var bm = document.getElementById("dd-view-map");
  if (view === "map") {
    tableWrap.classList.add("hidden"); mapWrap.classList.remove("hidden");
    bt.classList.remove("active"); bm.classList.add("active");
    renderDrilldownMap();
  } else {
    mapWrap.classList.add("hidden"); tableWrap.classList.remove("hidden");
    bm.classList.remove("active"); bt.classList.add("active");
  }
}

function renderDrilldownMap() {
  var raw = state.drilldown.filtered || [];
  // Normalise: exception uses activeFrom, speeding uses start
  var norm = raw.filter(function (e) { return e.device && e.device.id; }).map(function (e) {
    return { device: e.device, activeFrom: e.activeFrom || e.start, rule: e.rule };
  });

  if (!state.ddMap) {
    state.ddMap = L.map("dd-map").setView([54.5, -3], 5);
    L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
      maxZoom: 19, attribution: "&copy; OpenStreetMap contributors"
    }).addTo(state.ddMap);
    state.ddMapMarkers = L.layerGroup().addTo(state.ddMap);
  }
  setTimeout(function () { state.ddMap.invalidateSize(); }, 60);

  if (!norm.length) { plotMarkers(state.ddMap, state.ddMapMarkers, []); return; }
  correlateEventLocations(norm, function (located) {
    plotMarkers(state.ddMap, state.ddMapMarkers, located);
  });
}

// ─── Incident replay (cause → infraction → aftermath) ─────────────────────────
// Per-event mini reconstruction: LogRecord path on a map + synced speed line.
// The infraction window (activeFrom→activeTo) is red; ±60s padding is cause/aftermath.
// Hovering the speed graph scrubs a marker along the route.
function setupIncident() {
  var closeBtn = document.getElementById("inc-close");
  if (closeBtn) closeBtn.addEventListener("click", closeIncident);
  var overlay = document.getElementById("incident-modal");
  if (overlay) overlay.addEventListener("click", function (e) { if (e.target === overlay) closeIncident(); });
}

function closeIncident() {
  document.getElementById("incident-modal").classList.add("hidden");
  if (state.incChart) { state.incChart.destroy(); state.incChart = null; }
  if (state.incMap)   { state.incMap.remove();    state.incMap = null; }
  state.incMarker = null; state.incPoints = null;
}

function openIncident(ev) {
  var deviceId = ev.device && ev.device.id;
  if (!deviceId) return;
  var evFrom = new Date(ev.activeFrom).getTime();
  var evTo   = new Date(ev.activeTo || ev.activeFrom).getTime();
  if (isNaN(evTo) || evTo < evFrom) evTo = evFrom;
  var PAD  = 60000; // 60s cause + aftermath
  var from = new Date(evFrom - PAD), to = new Date(evTo + PAD);

  var vName    = state.deviceMap[deviceId] || deviceId;
  var ruleName = (ev.rule && state.ruleNameMap && state.ruleNameMap[ev.rule.id]) ||
                 (state.drilldown.widgetDef && state.drilldown.widgetDef.suggestion && state.drilldown.widgetDef.suggestion.rule && state.drilldown.widgetDef.suggestion.rule.name) ||
                 "Event";
  document.getElementById("inc-title").textContent    = vName + " \u2014 " + ruleName;
  document.getElementById("inc-subtitle").textContent = fmtDateReadable(ev.activeFrom) + " " + fmtTime(ev.activeFrom) + " \u2192 " + fmtTime(ev.activeTo || ev.activeFrom);
  document.getElementById("inc-readout").textContent  = "Loading\u2026";
  document.getElementById("incident-modal").classList.remove("hidden");

  apiCall("Get", { typeName: "LogRecord", search: { deviceSearch: { id: deviceId }, fromDate: from.toISOString(), toDate: to.toISOString() } }, function (recs) {
    var pts = (recs || [])
      .filter(function (r) { return r.latitude != null && r.longitude != null && !(r.latitude === 0 && r.longitude === 0); })
      .map(function (r) { return { t: new Date(r.dateTime).getTime(), lat: r.latitude, lng: r.longitude, mph: toMph(r.speed || 0) }; })
      .sort(function (a, b) { return a.t - b.t; });
    if (pts.length < 2) {
      document.getElementById("inc-readout").textContent = "Not enough GPS data for this event to build a replay.";
      renderIncident([], evFrom, evTo);
      return;
    }
    document.getElementById("inc-readout").textContent = "Hover the graph to trace the route.";
    renderIncident(pts, evFrom, evTo);
  }, function (err) {
    document.getElementById("inc-readout").textContent = "Error loading GPS: " + esc((err && err.message) || String(err));
  });
}

function renderIncident(pts, evFrom, evTo) {
  if (state.incChart) { state.incChart.destroy(); state.incChart = null; }
  if (state.incMap)   { state.incMap.remove();    state.incMap = null; }
  state.incPoints = pts;

  state.incMap = L.map("inc-map");
  L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", { maxZoom: 19, attribution: "&copy; OpenStreetMap contributors" }).addTo(state.incMap);

  if (!pts.length) { state.incMap.setView([54.5, -3], 5); setTimeout(function () { state.incMap.invalidateSize(); }, 60); return; }

  var isEvent = pts.map(function (p) { return p.t >= evFrom && p.t <= evTo; });
  var latlngs = pts.map(function (p) { return [p.lat, p.lng]; });

  L.polyline(latlngs, { color: "#94A3B8", weight: 4, opacity: 0.9 }).addTo(state.incMap);
  var evLatlngs = [];
  pts.forEach(function (p, i) { if (isEvent[i]) evLatlngs.push([p.lat, p.lng]); });
  if (evLatlngs.length >= 2) L.polyline(evLatlngs, { color: "#DC2626", weight: 6, opacity: 0.95 }).addTo(state.incMap);

  L.circleMarker(latlngs[0], { radius: 5, color: "#fff", weight: 2, fillColor: "#334155", fillOpacity: 1 }).bindTooltip("Start").addTo(state.incMap);
  L.circleMarker(latlngs[latlngs.length - 1], { radius: 5, color: "#fff", weight: 2, fillColor: "#334155", fillOpacity: 1 }).bindTooltip("End").addTo(state.incMap);

  var startIdx = isEvent.indexOf(true); if (startIdx < 0) startIdx = 0;
  state.incMarker = L.circleMarker(latlngs[startIdx], { radius: 8, color: "#fff", weight: 2, fillColor: "#DC2626", fillOpacity: 1 }).addTo(state.incMap);

  state.incMap.fitBounds(latlngs, { padding: [30, 30], maxZoom: 16 });
  setTimeout(function () { state.incMap.invalidateSize(); }, 60);

  var labels = pts.map(function (p) { return fmtTime(new Date(p.t).toISOString()); });
  var data   = pts.map(function (p) { return parseFloat(p.mph.toFixed(1)); });
  state.incChart = new Chart(document.getElementById("inc-chart"), {
    type: "line",
    data: { labels: labels, datasets: [{
      label: "Speed (mph)", data: data,
      borderColor: "#94A3B8", borderWidth: 2, fill: false, tension: 0.3,
      pointRadius: 0, pointHoverRadius: 5, pointBackgroundColor: "#DC2626",
      segment: { borderColor: function (c) { return (isEvent[c.p0DataIndex] && isEvent[c.p1DataIndex]) ? "#DC2626" : "#94A3B8"; } }
    }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { mode: "index", intersect: false },
      plugins: {
        legend: { display: false },
        tooltip: { enabled: true, callbacks: { label: function (c) { return c.parsed.y + " mph"; } } }
      },
      scales: {
        x: { ticks: { color: "#9CA3AF", font: { size: 10 }, maxRotation: 0, autoSkip: true, maxTicksLimit: 8 }, grid: { color: "rgba(0,0,0,.04)" } },
        y: { beginAtZero: true, title: { display: true, text: "mph", color: "#9CA3AF" }, ticks: { color: "#9CA3AF", font: { size: 10 } }, grid: { color: "rgba(0,0,0,.04)" } }
      },
      onHover: function (evt, els) {
        if (!els || !els.length) return;
        var i = els[0].index;
        var p = pts[i]; if (!p || !state.incMarker) return;
        state.incMarker.setLatLng([p.lat, p.lng]);
        document.getElementById("inc-readout").textContent =
          fmtTime(new Date(p.t).toISOString()) + " \u00b7 " + p.mph.toFixed(0) + " mph" + (isEvent[i] ? " \u00b7 INFRACTION" : "");
      }
    }
  });
}

// ─── Persistence (AddInData — server-side, shared across all users) ───────────
// Data is stored in MyGeotab's AddInData entity so the dashboard is consistent
// for every user who has access to the add-in, across browser sessions and devices.
// Falls back to localStorage when running outside MyGeotab (local dev).
var ADDIN_ID = "SmartInsightsDashboard";

function buildSavePayload() {
  return state.widgets.map(function (w) {
    var def  = w.widgetDef;
    var type = def ? def.type || "exception" : "exception";
    var entry = { id: w.id, title: w.title, type: type, params: def ? (def.params || {}) : {}, x: w.x, y: w.y, w: w.w, h: w.h };
    if (type === "exception" && def && def.suggestion) {
      entry.ruleId   = def.suggestion.rule && def.suggestion.rule.id;
      entry.ruleName = def.suggestion.rule && def.suggestion.rule.name;
      entry.count    = def.suggestion.count;
    }
    return entry;
  });
}

function saveDashboard() {
  var details = JSON.stringify(buildSavePayload());

  if (!state.api) {
    // Local dev fallback
    try { localStorage.setItem("smartinsights-dashboard", details); } catch (e) {}
    return;
  }

  if (state.dashboardDataId) {
    // Record already exists — update it
    state.api.call("Set", {
      typeName: "AddInData",
      entity: { id: state.dashboardDataId, addInId: ADDIN_ID, details: details }
    }, function () {}, function (err) {
      console.warn("SmartInsights: AddInData Set failed", err);
    });
  } else {
    // First save — create the shared record and cache its id
    state.api.call("Add", {
      typeName: "AddInData",
      entity: { addInId: ADDIN_ID, details: details }
    }, function (id) {
      state.dashboardDataId = id;
    }, function (err) {
      console.warn("SmartInsights: AddInData Add failed, falling back to localStorage", err);
      try { localStorage.setItem("smartinsights-dashboard", details); } catch (e) {}
    });
  }
}

// Inflate saved widget objects into live dashboard widgets.
function restoreWidgets(list) {
  if (!list || !list.length) return;
  list.forEach(function (w) {
    var type = w.type || (w.ruleId ? "exception" : null);
    if (!type) return;
    var widgetDef;
    if (type === "exception") {
      if (!w.ruleId) return;
      widgetDef = { type: "exception", suggestion: { rule: { id: w.ruleId, name: w.ruleName }, count: w.count || 0 } };
    } else {
      var typeDef2 = PRESET_WIDGET_TYPES[type];
      var defaultP = typeDef2 && typeDef2.defaultParams ? JSON.parse(JSON.stringify(typeDef2.defaultParams)) : {};
      widgetDef = { type: type, params: w.params || defaultP };
    }
    addWidget(w.id, w.title, widgetDef, w.x, w.y, w.w, w.h);
  });
}

// Async — calls onDone() when widgets are ready (or on error).
function restoreDashboard(onDone) {
  onDone = onDone || function () {};

  if (!state.api) {
    // Local dev: use localStorage directly
    try { restoreWidgets(JSON.parse(localStorage.getItem("smartinsights-dashboard"))); } catch (e) {}
    onDone();
    return;
  }

  state.api.call("Get", {
    typeName: "AddInData",
    search: { addInId: ADDIN_ID }
  }, function (results) {
    if (results && results.length > 0) {
      // Shared record found — restore from it
      state.dashboardDataId = results[0].id;
      try { restoreWidgets(JSON.parse(results[0].details)); } catch (e) {}
      onDone();
    } else {
      // No AddInData yet — migrate any existing localStorage data then save it
      var lsRaw = null;
      try { lsRaw = JSON.parse(localStorage.getItem("smartinsights-dashboard")); } catch (e) {}
      if (lsRaw && lsRaw.length) {
        restoreWidgets(lsRaw);
        saveDashboard(); // writes to AddInData and caches the new id
        try { localStorage.removeItem("smartinsights-dashboard"); } catch (e) {}
      }
      onDone();
    }
  }, function (err) {
    // AddInData unavailable — fall back to localStorage silently
    console.warn("SmartInsights: AddInData Get failed, using localStorage", err);
    try { restoreWidgets(JSON.parse(localStorage.getItem("smartinsights-dashboard"))); } catch (e) {}
    onDone();
  });
}

// ─── Helpers ──────────────────────────────────────────────────────────────────
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
// distM is Trip.Distance, which is in KILOMETRES (see milesFromDistance note).
// Returns UK imperial mpg.
function calcMpg(distM, fuelL) {
  if (!fuelL || fuelL <= 0 || !distM) return "—";
  var miles   = distM * 0.621371;
  var gallons = fuelL / 4.54609;  // UK imperial gallons
  return (miles / gallons).toFixed(1);
}
function toMph(kmh)         { return kmh * 0.621371; }
function mphStr(kmh)        { return kmh ? toMph(kmh).toFixed(0) + " mph" : "—"; }
function hexToRgba(hex, a)  { var r=parseInt(hex.slice(1,3),16),g=parseInt(hex.slice(3,5),16),b=parseInt(hex.slice(5,7),16); return "rgba("+r+","+g+","+b+","+a+")"; }
function fmtDateInput(d)    { return d.toISOString().slice(0, 10); }
function fmtDateShort(iso)  { return iso ? new Date(iso).toISOString().slice(0, 10) : ""; }
function fmtDateReadable(iso) { if (!iso) return ""; var d = new Date(iso); return d.toLocaleDateString(undefined, { day: "2-digit", month: "short", year: "numeric" }); }
function fmtTime(iso)       { return iso ? new Date(iso).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }) : "--"; }
function fmtMins(m)         { return m < 60 ? Math.round(m) + "m" : Math.floor(m / 60) + "h " + Math.round(m % 60) + "m"; }
function durationMins(a, b) { return (!a || !b) ? 0 : (new Date(b) - new Date(a)) / 60000; }
function esc(str)           { return String(str).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"); }
function sanitiseFilename(s){ return s.replace(/[^a-z0-9_\-]/gi, "_").toLowerCase(); }
function showError(msg)     { document.getElementById("loading").innerHTML = "<p style='color:var(--critical);padding:40px'>" + msg + "</p>"; }
// ─── Inline name edit ─────────────────────────────────────────────────────────
function enableNameEdit(widgetId) {
  var titleSpan = document.getElementById("title-" + widgetId);
  if (!titleSpan || titleSpan.querySelector("input")) return; // already editing
  var currentTitle = titleSpan.textContent;
  var input = document.createElement("input");
  input.type = "text";
  input.value = currentTitle;
  input.className = "widget-title-input";
  titleSpan.textContent = "";
  titleSpan.appendChild(input);
  input.focus();
  input.select();

  function save() {
    var newTitle = input.value.trim() || currentTitle;
    titleSpan.textContent = newTitle;
    var wState = state.widgets.find(function (w) { return w.id === widgetId; });
    if (wState) { wState.title = newTitle; saveDashboard(); }
  }
  input.addEventListener("blur", save);
  input.addEventListener("keydown", function (e) {
    if (e.key === "Enter") { input.blur(); }
    if (e.key === "Escape") { input.value = currentTitle; input.blur(); }
    e.stopPropagation();
  });
}

// ─── Param editor modal ───────────────────────────────────────────────────────
var _paramEditingId = null;

function openParamEditor(widgetId) {
  var wState = state.widgets.find(function (w) { return w.id === widgetId; });
  if (!wState) return;
  _paramEditingId = widgetId;
  var type    = wState.widgetDef.type;
  var params  = wState.widgetDef.params || {};
  var title   = document.getElementById("title-" + widgetId);
  document.getElementById("param-editor-title").textContent = (title ? title.textContent : wState.title) + " \u2014 Parameters";

  var body = document.getElementById("param-editor-body");
  if (type === "speeding") {
    body.innerHTML =
      "<div class='param-row'>" +
        "<label class='param-label'>Speed threshold</label>" +
        "<div class='param-input-wrap'>" +
          "<input type='number' id='pe-threshold' class='input-date' style='width:80px' min='1' max='200' value='" + (params.thresholdMph || 80) + "' />" +
          "<span class='param-unit'>mph</span>" +
        "</div>" +
        "<p class='param-hint'>Trips where maximum speed exceeds this value will be flagged.</p>" +
      "</div>";
  } else {
    body.innerHTML = "<p class='param-hint'>No configurable parameters for this widget type.</p>";
  }

  document.getElementById("param-editor").classList.remove("hidden");
}

function closeParamEditor() {
  document.getElementById("param-editor").classList.add("hidden");
  _paramEditingId = null;
}

function saveParamEditor() {
  if (!_paramEditingId) return;
  var wState = state.widgets.find(function (w) { return w.id === _paramEditingId; });
  if (!wState) { closeParamEditor(); return; }
  var type = wState.widgetDef.type;
  if (type === "speeding") {
    var thEl = document.getElementById("pe-threshold");
    var val  = thEl ? (parseFloat(thEl.value) || 80) : 80;
    wState.widgetDef.params = { thresholdMph: val };
    // Update title dot label hint in header (keep user title, don't rename)
  }
  saveDashboard();
  closeParamEditor();
  // Reload widget chart with new params
  loadWidgetData(wState.id, wState.widgetDef, wState.chart);
}

function setupParamEditor() {
  document.getElementById("param-editor-close").addEventListener("click", closeParamEditor);
  document.getElementById("param-editor-cancel").addEventListener("click", closeParamEditor);
  document.getElementById("param-editor-save").addEventListener("click", saveParamEditor);
  document.getElementById("param-editor").addEventListener("click", function (e) {
    if (e.target === this) closeParamEditor();
  });
}