"use strict";

// ─── State ─────────────────────────────────────────────────────────────────
const state = {
  api: null,
  reportData: [],
  widgets: [],
  suggestions: [],
};

// ─── MyGeotab addin lifecycle ──────────────────────────────────────────────
if (typeof geotab === "undefined") { var geotab = { addin: {} }; }

geotab.addin.smartInsights = function () {
  return {
    initialize: function (api, freshState, callback) {
      try {
        state.api = api;

        if (freshState && freshState.database) {
          document.getElementById("db-name").textContent = freshState.database;
        }

        setupNav();
        setupReports();
        restoreDashboard();
        document.getElementById("loading").classList.add("hidden");
        document.getElementById("main").classList.remove("hidden");
        loadSuggestions();
      } catch (err) {
        showError("Init error: " + err.message);
      }

      if (callback) callback();
    },
    focus: function (api, freshState) {},
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

// ─── Reports ────────────────────────────────────────────────────────────────
function setupReports() {
  var today = new Date();
  var week  = new Date(today);
  week.setDate(today.getDate() - 7);
  document.getElementById("filter-from").value = fmtDateInput(week);
  document.getElementById("filter-to").value   = fmtDateInput(today);
  document.getElementById("run-report").addEventListener("click", runReport);
  document.getElementById("export-csv").addEventListener("click", exportCsv);
}

function runReport() {
  var type   = document.getElementById("report-type").value;
  var from   = document.getElementById("filter-from").value;
  var to     = document.getElementById("filter-to").value;
  var output = document.getElementById("report-output");
  var btn    = document.getElementById("run-report");

  if (!from || !to) { alert("Please select a date range."); return; }

  btn.disabled = true;
  btn.textContent = "Loading...";
  output.innerHTML = "<p class='placeholder'>Fetching data...</p>";
  document.getElementById("report-summary").classList.add("hidden");
  document.getElementById("export-csv").classList.add("hidden");

  if (type === "trips") {
    runTripsReport(from, to, function () {
      btn.disabled = false;
      btn.textContent = "Run";
    });
  }
}

function runTripsReport(from, to, done) {
  var fromDate = new Date(from + "T00:00:00").toISOString();
  var toDate   = new Date(to   + "T23:59:59").toISOString();

  apiCall("Get", { typeName: "Device", search: {} }, function (devices) {
    var deviceMap = {};
    devices.forEach(function (d) { deviceMap[d.id] = d.name; });

    apiCall("Get", {
      typeName: "Trip",
      search: { fromDate: fromDate, toDate: toDate }
    }, function (trips) {
      state.reportData = trips;

      var totalKm   = trips.reduce(function (s, t) { return s + (t.distance || 0); }, 0) / 1000;
      var totalMins = trips.reduce(function (s, t) { return s + durationMins(t.start, t.stop); }, 0);
      var vehicles  = {};
      trips.forEach(function (t) { if (t.device) vehicles[t.device.id] = 1; });

      var summaryEl = document.getElementById("report-summary");
      summaryEl.innerHTML =
        summaryCard("Total Distance", totalKm.toFixed(1), "km") +
        summaryCard("Drive Time",     fmtMins(totalMins), "") +
        summaryCard("Trips",          trips.length, "") +
        summaryCard("Active Vehicles", Object.keys(vehicles).length, "");
      summaryEl.classList.remove("hidden");

      var output = document.getElementById("report-output");
      if (trips.length === 0) {
        output.innerHTML = "<p class='placeholder'>No trips found for this period.</p>";
        done(); return;
      }

      var sorted = trips.slice().sort(function (a, b) { return new Date(b.start) - new Date(a.start); });

      output.innerHTML = sorted.map(function (t) {
        var vName = deviceMap[t.device && t.device.id] || (t.device && t.device.id) || "Unknown";
        var dist  = ((t.distance || 0) / 1000).toFixed(1);
        var mins  = durationMins(t.start, t.stop);
        return (
          "<div class='trip-card'>" +
            "<div class='trip-top'>" +
              "<div><div class='trip-vehicle'>" + vName + "</div>" +
              "<div class='trip-driver'>" + (t.driverName || "No driver assigned") + "</div></div>" +
              "<div class='trip-time'>" + fmtDateShort(t.start) + "<br>" + fmtTime(t.start) + " &rarr; " + fmtTime(t.stop) + "</div>" +
            "</div>" +
            "<div class='trip-meta'>" +
              tripStat("Distance", dist + " km") +
              tripStat("Duration", fmtMins(mins)) +
              tripStat("Max Speed", (t.maximumSpeed || 0).toFixed(0) + " km/h") +
              tripStat("Avg Speed", (t.averageSpeed  || 0).toFixed(0) + " km/h") +
            "</div>" +
          "</div>"
        );
      }).join("");

      document.getElementById("export-csv").classList.remove("hidden");
      done();
    }, function (err) {
      document.getElementById("report-output").innerHTML = "<p class='placeholder' style='color:var(--red)'>Error: " + (err.message || err) + "</p>";
      done();
    });
  }, function (err) {
    document.getElementById("report-output").innerHTML = "<p class='placeholder' style='color:var(--red)'>Error loading devices: " + (err.message || err) + "</p>";
    done();
  });
}

function summaryCard(label, value, unit) {
  return "<div class='summary-card'><div class='summary-label'>" + label + "</div><div class='summary-value'>" + value + "<span class='summary-unit'>" + unit + "</span></div></div>";
}

function tripStat(label, value) {
  return "<div class='trip-stat'><span class='trip-stat-label'>" + label + "</span><span class='trip-stat-value'>" + value + "</span></div>";
}

function exportCsv() {
  if (!state.reportData.length) return;
  var rows = [["Vehicle ID", "Driver", "Date", "Start", "Stop", "Distance (km)", "Duration", "Max Speed (km/h)", "Avg Speed (km/h)"]];
  state.reportData.forEach(function (t) {
    rows.push([
      (t.device && t.device.id) || "",
      t.driverName || "",
      fmtDateShort(t.start),
      fmtTime(t.start),
      fmtTime(t.stop),
      ((t.distance || 0) / 1000).toFixed(1),
      fmtMins(durationMins(t.start, t.stop)),
      (t.maximumSpeed || 0).toFixed(0),
      (t.averageSpeed  || 0).toFixed(0),
    ]);
  });
  var csv  = rows.map(function (r) { return r.join(","); }).join("\n");
  var blob = new Blob([csv], { type: "text/csv" });
  var url  = URL.createObjectURL(blob);
  var a    = document.createElement("a");
  a.href = url; a.download = "trips.csv"; a.click();
  URL.revokeObjectURL(url);
}

// ─── Suggestions ─────────────────────────────────────────────────────────────
function loadSuggestions() {
  document.getElementById("refresh-suggestions").addEventListener("click", loadSuggestions);
  var list = document.getElementById("suggestions-list");
  list.innerHTML = "<p class='placeholder'>Scanning your fleet data...</p>";

  var toDate   = new Date();
  var fromDate = new Date();
  fromDate.setDate(toDate.getDate() - 30);

  apiCall("Get", {
    typeName: "ExceptionEvent",
    search: { fromDate: fromDate.toISOString(), toDate: toDate.toISOString(), includeInvalidated: false }
  }, function (exceptions) {
    apiCall("Get", { typeName: "Rule", search: {} }, function (rules) {
      var counts = {};
      exceptions.forEach(function (e) {
        var id = e.rule && e.rule.id;
        if (id) counts[id] = (counts[id] || 0) + 1;
      });

      var ruleMap = {};
      rules.forEach(function (r) { ruleMap[r.id] = r; });

      var ranked = Object.keys(counts)
        .map(function (id) { return { rule: ruleMap[id], count: counts[id] }; })
        .filter(function (s) { return s.rule; })
        .sort(function (a, b) { return b.count - a.count; })
        .slice(0, 8);

      state.suggestions = ranked;

      if (ranked.length === 0) {
        list.innerHTML = "<p class='placeholder'>No violations found in the last 30 days.</p>";
        return;
      }

      list.innerHTML = ranked.map(function (s, i) { return suggestionCard(s, i); }).join("");

      list.querySelectorAll(".btn-add").forEach(function (btn) {
        btn.addEventListener("click", function () {
          addSuggestionToDashboard(state.suggestions[parseInt(btn.dataset.idx)]);
        });
      });
    }, handleErr(list));
  }, handleErr(list));
}

function suggestionCard(s, idx) {
  var icon = ruleIcon(s.rule.name || "");
  return (
    "<div class='suggestion-card'>" +
      "<div class='suggestion-icon'>" + icon + "</div>" +
      "<div class='suggestion-body'>" +
        "<div class='suggestion-title'>" + (s.rule.name || "Unknown Rule") + "</div>" +
        "<div class='suggestion-desc'>" + s.count + " violation" + (s.count !== 1 ? "s" : "") + " in the last 30 days</div>" +
      "</div>" +
      "<div class='suggestion-count'>" + s.count + "<span>violations</span></div>" +
      "<button class='btn btn-add btn-sm' data-idx='" + idx + "'>+ Dashboard</button>" +
    "</div>"
  );
}

function ruleIcon(name) {
  var n = name.toLowerCase();
  if (n.includes("speed"))   return "🚀";
  if (n.includes("idle"))    return "⏱️";
  if (n.includes("harsh") || n.includes("brake")) return "⚠️";
  if (n.includes("seat")  || n.includes("belt"))  return "🔒";
  if (n.includes("fuel"))    return "⛽";
  if (n.includes("after hours") || n.includes("curfew")) return "🌙";
  return "📊";
}

// ─── Dashboard ───────────────────────────────────────────────────────────────
function addSuggestionToDashboard(suggestion) {
  var id    = "widget-" + Date.now();
  var title = suggestion.rule.name + " (30 days)";

  document.querySelectorAll(".nav-btn").forEach(function (b) { b.classList.remove("active"); });
  document.querySelectorAll(".tab").forEach(function (t) { t.classList.remove("active"); });
  document.querySelector("[data-tab='dashboard']").classList.add("active");
  document.getElementById("tab-dashboard").classList.add("active");

  addWidget(id, title, suggestion);
  saveDashboard();
}

function addWidget(id, title, suggestion) {
  var grid  = document.getElementById("dashboard-grid");
  var empty = grid.querySelector(".grid-empty");
  if (empty) empty.remove();

  var widget = document.createElement("div");
  widget.className = "widget";
  widget.id = id;
  widget.innerHTML =
    "<div class='widget-header'>" +
      "<span class='widget-title'>" + title + "</span>" +
      "<button class='widget-remove' title='Remove'>&#x2715;</button>" +
    "</div>" +
    "<div class='widget-body'><canvas id='canvas-" + id + "'></canvas></div>";
  grid.appendChild(widget);

  widget.querySelector(".widget-remove").addEventListener("click", function () {
    var inst = state.widgets.find(function (w) { return w.id === id; });
    if (inst && inst.chart) inst.chart.destroy();
    state.widgets = state.widgets.filter(function (w) { return w.id !== id; });
    widget.remove();
    if (!grid.querySelector(".widget")) {
      grid.innerHTML = "<div class='grid-empty'><p>No widgets yet.</p><p>Head to <strong>Suggestions</strong> to add your first chart.</p></div>";
    }
    saveDashboard();
  });

  var canvas = document.getElementById("canvas-" + id);
  var chart  = new Chart(canvas, {
    type: "bar",
    data: { labels: ["Loading..."], datasets: [{ data: [0], backgroundColor: "rgba(99,102,241,.7)", borderRadius: 4 }] },
    options: { responsive: true, plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true, ticks: { precision: 0 } } } }
  });

  state.widgets.push({ id: id, title: title, chart: chart, suggestion: suggestion });
  loadWidgetData(id, suggestion, chart);
}

function loadWidgetData(widgetId, suggestion, chart) {
  var toDate   = new Date();
  var fromDate = new Date();
  fromDate.setDate(toDate.getDate() - 30);

  apiCall("Get", {
    typeName: "ExceptionEvent",
    search: {
      fromDate: fromDate.toISOString(),
      toDate:   toDate.toISOString(),
      ruleSearch: { id: suggestion.rule.id },
      includeInvalidated: false
    }
  }, function (exceptions) {
    var buckets = {};
    exceptions.forEach(function (e) {
      var day = fmtDateShort(e.activeFrom);
      buckets[day] = (buckets[day] || 0) + 1;
    });

    var labels = [], data = [];
    for (var i = 29; i >= 0; i--) {
      var d = new Date();
      d.setDate(d.getDate() - i);
      var key = fmtDateShort(d.toISOString());
      labels.push(key.slice(5));
      data.push(buckets[key] || 0);
    }

    chart.data.labels = labels;
    chart.data.datasets = [{ label: "Violations", data: data, backgroundColor: "rgba(99,102,241,.7)", borderRadius: 4 }];
    chart.options.plugins.legend.display = true;
    chart.update();
  }, function () {});
}

function saveDashboard() {
  var saved = state.widgets.map(function (w) {
    return { id: w.id, title: w.title, ruleId: w.suggestion && w.suggestion.rule && w.suggestion.rule.id, ruleName: w.suggestion && w.suggestion.rule && w.suggestion.rule.name };
  });
  try { localStorage.setItem("smartinsights-dashboard", JSON.stringify(saved)); } catch (e) {}
}

function restoreDashboard() {
  try {
    var raw = localStorage.getItem("smartinsights-dashboard");
    if (!raw) return;
    JSON.parse(raw).forEach(function (w) {
      if (!w.ruleId) return;
      addWidget(w.id, w.title, { rule: { id: w.ruleId, name: w.ruleName } });
    });
  } catch (e) {}
}

// ─── Helpers ──────────────────────────────────────────────────────────────────
function apiCall(method, params, onSuccess, onError) {
  state.api.call(method, params, onSuccess, onError || function (err) { console.error(method, err); });
}

function handleErr(container) {
  return function (err) {
    container.innerHTML = "<p class='placeholder' style='color:var(--red)'>Error: " + (err.message || err) + "</p>";
  };
}

function fmtDateInput(d) { return d.toISOString().slice(0, 10); }
function fmtDateShort(iso) { return iso ? new Date(iso).toISOString().slice(0, 10) : ""; }
function fmtTime(iso) { return iso ? new Date(iso).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }) : "--"; }
function fmtMins(m) { return m < 60 ? Math.round(m) + "m" : Math.floor(m / 60) + "h " + Math.round(m % 60) + "m"; }
function durationMins(a, b) { return (!a || !b) ? 0 : (new Date(b) - new Date(a)) / 60000; }
function showError(msg) { document.getElementById("loading").innerHTML = "<p style='color:var(--red);padding:40px'>" + msg + "</p>"; }
