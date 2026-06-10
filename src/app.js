"use strict";

// ─── State ─────────────────────────────────────────────────────────────────
const state = {
  api: null,
  session: null,
  widgets: [],       // { id, title, type, chartInstance }
  reportData: [],    // raw rows from last report run
  suggestions: [],   // ranked suggestion objects
};

// ─── MyGeotab addin entry point ────────────────────────────────────────────
geotab.addin.smartInsights = (elt, service) => {
  init(service);
};

async function init(service) {
  try {
    state.session = await service.api.getSession();
    state.api = service.api;

    document.getElementById("db-name").textContent = state.session.database;

    setupNav();
    setupReports();
    restoreDashboard();

    document.getElementById("loading").classList.add("hidden");
    document.getElementById("main").classList.remove("hidden");

    loadSuggestions();
  } catch (err) {
    showError("Failed to connect to MyGeotab: " + err.message);
  }
}

// ─── Navigation ─────────────────────────────────────────────────────────────
function setupNav() {
  document.querySelectorAll(".nav-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      const tab = btn.dataset.tab;
      document.querySelectorAll(".nav-btn").forEach(b => b.classList.remove("active"));
      document.querySelectorAll(".tab").forEach(t => t.classList.remove("active"));
      btn.classList.add("active");
      document.getElementById(`tab-${tab}`).classList.add("active");
    });
  });
}

// ─── Reports ────────────────────────────────────────────────────────────────
function setupReports() {
  // Default date range: last 7 days
  const today = new Date();
  const week  = new Date(today);
  week.setDate(today.getDate() - 7);

  document.getElementById("filter-from").value = fmtDateInput(week);
  document.getElementById("filter-to").value   = fmtDateInput(today);

  document.getElementById("run-report").addEventListener("click", runReport);
  document.getElementById("export-csv").addEventListener("click", exportCsv);
}

async function runReport() {
  const type   = document.getElementById("report-type").value;
  const from   = document.getElementById("filter-from").value;
  const to     = document.getElementById("filter-to").value;
  const output = document.getElementById("report-output");
  const btn    = document.getElementById("run-report");

  if (!from || !to) { alert("Please select a date range."); return; }

  btn.disabled = true;
  btn.textContent = "Loading...";
  output.innerHTML = "<p class='report-placeholder'>Fetching data...</p>";
  document.getElementById("report-summary").classList.add("hidden");
  document.getElementById("export-csv").classList.add("hidden");

  try {
    if (type === "trips") await runTripsReport(from, to);
  } catch (err) {
    output.innerHTML = `<p class='report-placeholder' style='color:var(--red)'>Error: ${err.message}</p>`;
  } finally {
    btn.disabled = false;
    btn.textContent = "Run";
  }
}

async function runTripsReport(from, to) {
  const fromDate = new Date(from + "T00:00:00");
  const toDate   = new Date(to   + "T23:59:59");

  const [trips, devices] = await Promise.all([
    apiCall("Get", {
      typeName: "Trip",
      search: {
        fromDate: fromDate.toISOString(),
        toDate:   toDate.toISOString(),
      }
    }),
    apiCall("Get", { typeName: "Device", search: {} }),
  ]);

  state.reportData = trips;

  const deviceMap = {};
  devices.forEach(d => { deviceMap[d.id] = d.name; });

  // Summary stats
  const totalKm   = trips.reduce((s, t) => s + (t.distance || 0), 0) / 1000;
  const totalMins = trips.reduce((s, t) => s + durationMins(t.start, t.stop), 0);
  const vehicles  = new Set(trips.map(t => t.device?.id)).size;

  const summaryEl = document.getElementById("report-summary");
  summaryEl.innerHTML = `
    ${summaryCard("Total Distance", totalKm.toFixed(1), "km")}
    ${summaryCard("Drive Time",     fmtMins(totalMins), "")}
    ${summaryCard("Trips",          trips.length, "")}
    ${summaryCard("Active Vehicles", vehicles, "")}
  `;
  summaryEl.classList.remove("hidden");

  // Trip cards
  const output = document.getElementById("report-output");
  if (trips.length === 0) {
    output.innerHTML = "<p class='report-placeholder'>No trips found for this period.</p>";
    return;
  }

  // Sort newest first
  const sorted = [...trips].sort((a, b) => new Date(b.start) - new Date(a.start));

  output.innerHTML = sorted.map(t => {
    const vName  = deviceMap[t.device?.id] || t.device?.id || "Unknown";
    const dist   = ((t.distance || 0) / 1000).toFixed(1);
    const mins   = durationMins(t.start, t.stop);
    const start  = fmtTime(t.start);
    const stop   = fmtTime(t.stop);
    const date   = fmtDateShort(t.start);

    return `
      <div class="trip-card">
        <div>
          <div class="trip-vehicle">${vName}</div>
          <div class="trip-driver">${t.driverName || "No driver assigned"}</div>
        </div>
        <div class="trip-time">${date}<br>${start} &rarr; ${stop}</div>
        <div class="trip-meta">
          ${tripStat("Distance", dist + " km")}
          ${tripStat("Duration", fmtMins(mins))}
          ${tripStat("Max Speed", (t.maximumSpeed || 0).toFixed(0) + " km/h")}
          ${tripStat("Avg Speed", (t.averageSpeed || 0).toFixed(0) + " km/h")}
        </div>
      </div>`;
  }).join("");

  document.getElementById("export-csv").classList.remove("hidden");
}

function summaryCard(label, value, unit) {
  return `<div class="summary-card">
    <div class="summary-label">${label}</div>
    <div class="summary-value">${value}<span class="summary-unit">${unit}</span></div>
  </div>`;
}

function tripStat(label, value) {
  return `<div class="trip-stat">
    <span class="trip-stat-label">${label}</span>
    <span class="trip-stat-value">${value}</span>
  </div>`;
}

function exportCsv() {
  if (!state.reportData.length) return;
  const rows = [["Vehicle", "Driver", "Date", "Start", "Stop", "Distance (km)", "Duration", "Max Speed (km/h)", "Avg Speed (km/h)"]];
  state.reportData.forEach(t => {
    rows.push([
      t.device?.id || "",
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
  const csv  = rows.map(r => r.join(",")).join("\n");
  const blob = new Blob([csv], { type: "text/csv" });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a");
  a.href = url; a.download = "trips.csv"; a.click();
  URL.revokeObjectURL(url);
}

// ─── Suggestions ─────────────────────────────────────────────────────────────
async function loadSuggestions() {
  document.getElementById("refresh-suggestions").addEventListener("click", loadSuggestions);
  const list = document.getElementById("suggestions-list");
  list.innerHTML = "<p class='report-placeholder'>Scanning your fleet data...</p>";

  try {
    const toDate   = new Date();
    const fromDate = new Date();
    fromDate.setDate(toDate.getDate() - 30);

    const [exceptions, rules] = await Promise.all([
      apiCall("Get", {
        typeName: "ExceptionEvent",
        search: {
          fromDate: fromDate.toISOString(),
          toDate:   toDate.toISOString(),
          includeInvalidated: false,
        }
      }),
      apiCall("Get", { typeName: "Rule", search: {} }),
    ]);

    // Count violations per rule
    const counts = {};
    exceptions.forEach(e => {
      const id = e.rule?.id;
      if (id) counts[id] = (counts[id] || 0) + 1;
    });

    const ruleMap = {};
    rules.forEach(r => { ruleMap[r.id] = r; });

    // Rank by violation count, take top 8
    const ranked = Object.entries(counts)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 8)
      .map(([id, count]) => ({ rule: ruleMap[id], count }))
      .filter(s => s.rule);

    state.suggestions = ranked;

    if (ranked.length === 0) {
      list.innerHTML = "<p class='report-placeholder'>No violations found in the last 30 days. Your fleet is clean!</p>";
      return;
    }

    list.innerHTML = ranked.map((s, i) => suggestionCard(s, i)).join("");

    list.querySelectorAll(".btn-success").forEach(btn => {
      btn.addEventListener("click", () => {
        const idx = parseInt(btn.dataset.idx);
        addSuggestionToDashboard(state.suggestions[idx]);
      });
    });

  } catch (err) {
    list.innerHTML = `<p class='report-placeholder' style='color:var(--red)'>Error loading suggestions: ${err.message}</p>`;
  }
}

function suggestionCard(s, idx) {
  const name = s.rule.name || "Unknown Rule";
  const icon = ruleIcon(name);
  return `
    <div class="suggestion-card">
      <div class="suggestion-icon">${icon}</div>
      <div class="suggestion-body">
        <div class="suggestion-title">${name}</div>
        <div class="suggestion-desc">Violated ${s.count} time${s.count !== 1 ? "s" : ""} in the last 30 days</div>
      </div>
      <div class="suggestion-count">${s.count}<span>violations</span></div>
      <button class="btn btn-success btn-sm" data-idx="${idx}">+ Dashboard</button>
    </div>`;
}

function ruleIcon(name) {
  const n = name.toLowerCase();
  if (n.includes("speed"))   return "🚀";
  if (n.includes("idle"))    return "⏱️";
  if (n.includes("harsh") || n.includes("brake")) return "⚠️";
  if (n.includes("seat") || n.includes("belt"))   return "🔒";
  if (n.includes("fuel"))    return "⛽";
  if (n.includes("after hours") || n.includes("curfew")) return "🌙";
  return "📊";
}

// ─── Dashboard ───────────────────────────────────────────────────────────────
function addSuggestionToDashboard(suggestion) {
  const id    = "widget-" + Date.now();
  const title = suggestion.rule.name + " (last 30 days)";

  // Switch to dashboard tab
  document.querySelectorAll(".nav-btn").forEach(b => b.classList.remove("active"));
  document.querySelectorAll(".tab").forEach(t => t.classList.remove("active"));
  document.querySelector('[data-tab="dashboard"]').classList.add("active");
  document.getElementById("tab-dashboard").classList.add("active");

  addWidget(id, title, "bar", suggestion);
  saveDashboard();
}

function addWidget(id, title, chartType, suggestion) {
  const grid  = document.getElementById("dashboard-grid");
  const empty = grid.querySelector(".grid-empty");
  if (empty) empty.remove();

  const widget = document.createElement("div");
  widget.className = "widget";
  widget.id = id;
  widget.innerHTML = `
    <div class="widget-header">
      <span class="widget-title">${title}</span>
      <div class="widget-actions">
        <button class="widget-remove" title="Remove">✕</button>
      </div>
    </div>
    <div class="widget-body">
      <canvas id="canvas-${id}"></canvas>
    </div>`;
  grid.appendChild(widget);

  widget.querySelector(".widget-remove").addEventListener("click", () => {
    const inst = state.widgets.find(w => w.id === id);
    if (inst?.chartInstance) inst.chartInstance.destroy();
    state.widgets = state.widgets.filter(w => w.id !== id);
    widget.remove();
    if (!grid.querySelector(".widget")) {
      grid.innerHTML = `<div class="grid-empty"><p>No widgets yet.</p><p>Head to <strong>Suggestions</strong> to add your first chart.</p></div>`;
    }
    saveDashboard();
  });

  // Render chart with placeholder — real async data load
  const canvas = document.getElementById(`canvas-${id}`);
  const chart  = new Chart(canvas, {
    type: chartType,
    data: { labels: ["Loading..."], datasets: [{ data: [0] }] },
    options: { responsive: true, plugins: { legend: { display: false } } }
  });

  state.widgets.push({ id, title, chartType, chartInstance: chart, suggestion });

  // Load real data for the chart
  loadWidgetData(id, suggestion, chart);
}

async function loadWidgetData(widgetId, suggestion, chart) {
  try {
    const toDate   = new Date();
    const fromDate = new Date();
    fromDate.setDate(toDate.getDate() - 30);

    const exceptions = await apiCall("Get", {
      typeName: "ExceptionEvent",
      search: {
        fromDate: fromDate.toISOString(),
        toDate:   toDate.toISOString(),
        ruleSearch: { id: suggestion.rule.id },
        includeInvalidated: false,
      }
    });

    // Bucket by day
    const buckets = {};
    exceptions.forEach(e => {
      const day = fmtDateShort(e.activeFrom);
      buckets[day] = (buckets[day] || 0) + 1;
    });

    // Fill last 30 days in order
    const labels = [];
    const data   = [];
    for (let i = 29; i >= 0; i--) {
      const d = new Date();
      d.setDate(d.getDate() - i);
      const key = fmtDateShort(d.toISOString());
      labels.push(key.slice(5)); // MM-DD
      data.push(buckets[key] || 0);
    }

    chart.data.labels = labels;
    chart.data.datasets = [{
      label: "Violations",
      data,
      backgroundColor: "rgba(0, 115, 230, 0.7)",
      borderRadius: 4,
    }];
    chart.options.plugins.legend.display = true;
    chart.update();
  } catch (err) {
    console.warn("Widget data load failed:", err);
  }
}

// ─── Dashboard persistence ────────────────────────────────────────────────────
function saveDashboard() {
  const saved = state.widgets.map(w => ({
    id:        w.id,
    title:     w.title,
    chartType: w.chartType,
    ruleId:    w.suggestion?.rule?.id,
    ruleName:  w.suggestion?.rule?.name,
  }));
  localStorage.setItem("smartinsights-dashboard", JSON.stringify(saved));
}

function restoreDashboard() {
  const raw = localStorage.getItem("smartinsights-dashboard");
  if (!raw) return;
  try {
    const saved = JSON.parse(raw);
    saved.forEach(w => {
      if (!w.ruleId) return;
      addWidget(w.id, w.title, w.chartType || "bar", {
        rule: { id: w.ruleId, name: w.ruleName }
      });
    });
  } catch (e) {
    console.warn("Failed to restore dashboard:", e);
  }
}

// ─── API helper ──────────────────────────────────────────────────────────────
function apiCall(method, params) {
  return new Promise((resolve, reject) => {
    state.api.call(method, params, resolve, reject);
  });
}

// ─── Utilities ───────────────────────────────────────────────────────────────
function fmtDateInput(d) {
  return d.toISOString().slice(0, 10);
}

function fmtDateShort(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  return d.toISOString().slice(0, 10);
}

function fmtTime(iso) {
  if (!iso) return "--";
  return new Date(iso).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
}

function fmtMins(mins) {
  if (mins < 60) return `${Math.round(mins)}m`;
  return `${Math.floor(mins / 60)}h ${Math.round(mins % 60)}m`;
}

function durationMins(start, stop) {
  if (!start || !stop) return 0;
  return (new Date(stop) - new Date(start)) / 60000;
}

function showError(msg) {
  document.getElementById("loading").innerHTML = `<p style="color:var(--red);padding:40px">${msg}</p>`;
}
