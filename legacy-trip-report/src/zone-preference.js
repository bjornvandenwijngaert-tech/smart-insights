"use strict";

// AddInData is scoped to a MyGeotab database, NOT intrinsically to a user.
// Use a separate add-in namespace and validate the signed-in User.id on reads.
// This boolean is logically per-user; it is not a security-private record.
var LEGACY_ZONE_PREFERENCE_ID = "aDFiMDc2OGMtMmI5OS00NDI";

function legacyPreferenceOwner() {
  return state.legacyFilters && state.legacyFilters.preference;
}
function legacyPreferenceCurrent(pref) {
  return !state.legacyIdentityStale && state.legacyFilters === pref.filters &&
    state.initializationEpoch === pref.epoch && state.dbName === pref.database &&
    state.legacyViewerUserId === pref.userId && state.legacyViewerLogin === pref.login;
}
function legacyPreferenceStatus() {
  var pref = legacyPreferenceOwner(), toggle = document.getElementById("legacy-zones");
  if (toggle) toggle.disabled = !!(state.legacyIdentityStale || state.legacyFocusChecking ||
    !pref || pref.status === "loading" || pref.status === "saving");
  var status = document.getElementById("legacy-zone-preference-status");
  if (!status) return;
  status.textContent = state.legacyIdentityStale ? "Refresh the add-in after changing database or account." :
    state.legacyFocusChecking ? "Checking account..." : !pref || pref.status === "loading" ? "Loading Zones preference..." :
    pref.status === "saving" ? "Saving Zones preference..." : pref.status === "unsaved" ?
      "Zones preference not saved; refresh to retry." : "";
}
function legacyPreferenceSuspend() {
  if (!legacyIsActive()) { legacyPreferenceStatus(); return; }
  document.getElementById("legacy-filters").classList.add("hidden");
  document.getElementById("report-output").classList.add("hidden");
  document.getElementById("export-csv").classList.add("hidden");
  document.getElementById("export-pdf-legacy").classList.add("hidden");
  document.getElementById("run-report").disabled = true;
  legacyPreferenceStatus();
}
function legacyPreferenceResume() {
  if (state.legacyIdentityStale) {
    if (legacyIsActive()) {
      document.getElementById("legacy-filters").classList.add("hidden");
      document.getElementById("report-output").classList.remove("hidden");
      document.getElementById("report-output").innerHTML =
        "<p class='placeholder'>Database or account changed. Refresh the add-in to load this database safely.</p>";
      document.getElementById("run-report").disabled = true;
    }
    legacyPreferenceStatus(); return;
  }
  if (!legacyIsActive() || state.legacyFocusChecking) { legacyPreferenceStatus(); return; }
  var pref = legacyPreferenceOwner();
  if (pref && (pref.status === "loading" || pref.status === "saving")) { legacyPreferenceSuspend(); return; }
  document.getElementById("legacy-filters").classList.remove("hidden");
  document.getElementById("report-output").classList.remove("hidden");
  document.getElementById("run-report").disabled = !!state.legacyLoading;
  if (state.legacyVisibleResult && state.legacyByDevice && Object.keys(state.legacyByDevice).length) {
    ["export-csv", "export-pdf-legacy"].forEach(function (id) { document.getElementById(id).classList.remove("hidden"); });
  }
  legacyPreferenceStatus();
}
function legacyPreferenceStale() {
  var showingLegacy = legacyIsActive();
  state.legacyIdentityStale = true;
  state.legacyFocusChecking = false;
  state.legacyPendingFocus = null;
  if (showingLegacy) cancelLegacyReport();
  else cancelLegacyZones(); // do not cancel an unrelated pending parent report
  if (showingLegacy) clearLegacyResults();
  else {
    state.legacyFullByDevice = null; state.legacyByDevice = null;
    state.legacyAddrMap = null; state.legacyStartLocations = null;
    state.legacyStopInfo = null; state.legacyZoneSnapshot = null;
    state.legacyReportRange = null;
  }
  state.legacyFilters = null;
  if (showingLegacy) { state.reportRaw = []; state.reportData = []; state.reportDriverMap = {}; }
  state.legacyViewerUserId = null; state.legacyViewerLogin = "";
  document.getElementById("legacy-filters").classList.add("hidden");
  legacyPreferenceResume();
}
function legacyPreferenceCall(pref, method, params, done) {
  var settled = false;
  var timer = setTimeout(function () { finish(new Error("Preference request timed out")); }, 15000);
  function finish(err, value) {
    if (settled) return;
    settled = true; clearTimeout(timer);
    if (legacyPreferenceCurrent(pref)) done(err, value);
  }
  try { pref.api.call(method, params, function (value) { finish(null, value); }, finish); }
  catch (err) { if (settled) throw err; finish(err); }
}
function legacyPreferenceRead(pref) {
  if (!legacyPreferenceCurrent(pref) || !pref.verified || pref.writing) return;
  pref.status = "loading";
  var token = ++pref.readToken;
  legacyPreferenceSuspend();
  legacyPreferenceCall(pref, "Get", { typeName: "AddInData", search: {
    addInId: LEGACY_ZONE_PREFERENCE_ID, whereClause: "userId = " + JSON.stringify(pref.userId)
  } }, function (err, records) {
    if (!pref.verified || token !== pref.readToken) return;
    if (err || !Array.isArray(records)) { pref.status = "unsaved"; legacyPreferenceResume(); return; }
    var matched = [], invalid = false;
    records.forEach(function (record) {
      if (!record || record.addInId !== LEGACY_ZONE_PREFERENCE_ID) { invalid = true; return; }
      var details = record.details;
      if (!details || typeof details !== "object" || Array.isArray(details)) { invalid = true; return; }
      if (details.userId !== pref.userId) return; // never accept another user's setting
      if (details.version !== 1 || typeof details.zonesEnabled !== "boolean" || typeof record.id !== "string") { invalid = true; return; }
      matched.push(record);
    });
    if (invalid) { pref.status = "unsaved"; legacyPreferenceResume(); return; }
    // Client clocks cannot order concurrent first-saves from different devices.
    // Never silently choose a potentially older choice when duplicates disagree.
    if (matched.some(function (record) { return record.details.zonesEnabled !== matched[0].details.zonesEnabled; })) {
      pref.record = null; pref.status = "unsaved"; legacyPreferenceResume(); return;
    }
    matched.sort(function (a, b) { return a.id.localeCompare(b.id); });
    pref.record = matched[0] || null;
    var restored = pref.record ? pref.record.details.zonesEnabled : false;
    if (pref.filters.zones !== restored && state.legacyZoneSnapshot && state.legacyZoneSnapshot.status === "loading") {
      cancelLegacyZones();
      state.legacyZoneSnapshot = null; // re-enable must start a fresh membership lookup
    }
    pref.filters.zones = restored;
    pref.status = "ready";
    updateLegacyZoneControls();
    if (legacyIsActive() && state.legacyVisibleResult && state.legacyFullByDevice) applyLegacyFilters();
    legacyPreferenceResume();
  });
}
function legacyPreferenceSave() {
  var pref = legacyPreferenceOwner();
  if (!pref || !legacyPreferenceCurrent(pref) || !pref.verified || state.legacyFocusChecking || pref.status !== "ready") return;
  var enabled = pref.filters.zones;
  var details = Object.assign({}, pref.record && pref.record.details, {
    version: 1, userId: pref.userId, zonesEnabled: enabled, updatedAt: new Date().toISOString()
  });
  var record = pref.record, entity = { addInId: LEGACY_ZONE_PREFERENCE_ID, details: details };
  if (record) { entity.id = record.id; if (record.groups) entity.groups = record.groups; }
  pref.status = "saving";
  pref.writing = true;
  legacyPreferenceStatus();
  legacyPreferenceCall(pref, record ? "Set" : "Add", { typeName: "AddInData", entity: entity }, function (err, id) {
    // A focus check may complete before this write. Keep its outcome private
    // until the same user is verified again; never re-read over a pending write.
    pref.writing = false;
    pref.writeOutcome = { err: err, id: id, entity: entity, added: !record };
    if (pref.verified && !state.legacyFocusChecking) legacyPreferenceApplyWrite(pref);
  });
}
function legacyPreferenceApplyWrite(pref) {
  if (!legacyPreferenceCurrent(pref) || !pref.verified || !pref.writeOutcome) return;
  var outcome = pref.writeOutcome;
  pref.writeOutcome = null;
  if (outcome.err || (outcome.added && (typeof outcome.id !== "string" || !outcome.id))) {
    // An Add timeout might have committed. Never automatically Add again.
    pref.status = "unsaved";
  } else {
    pref.record = { id: outcome.added ? outcome.id : outcome.entity.id,
      addInId: LEGACY_ZONE_PREFERENCE_ID, groups: outcome.entity.groups || [], details: outcome.entity.details };
    pref.status = "ready";
  }
  legacyPreferenceResume();
}
function legacyPreferenceStart() {
  if (state.legacyIdentityStale || state.legacyFocusChecking) { legacyPreferenceSuspend(); return; }
  var filters = legacyFilters();
  if (filters.preference) { legacyPreferenceResume(); return; }
  var userId = state.legacyViewerUserId, login = state.legacyViewerLogin;
  if (!userId || !login || !state.dbName || !state.api) {
    if (!state.legacyUnitsResolved) { legacyPreferenceSuspend(); return; }
    filters.preference = { status: "unsaved" };
    legacyPreferenceResume(); return;
  }
  var pref = filters.preference = { filters: filters, api: state.api, database: state.dbName,
    userId: userId, login: login, epoch: state.initializationEpoch, verified: true,
    readToken: 0, record: null, writing: false, writeOutcome: null, status: "loading" };
  legacyPreferenceRead(pref);
}
function legacyPreferenceAfterUser() {
  state.legacyUnitsResolved = true;
  if (state.legacyPendingFocus) { legacyPreferenceVerifyFocus(state.legacyPendingFocus); return; }
  if (legacyIsActive()) legacyPreferenceStart();
}
function legacyPreferenceFocus(api, freshState) {
  if (state.legacyIdentityStale) return;
  if (freshState && freshState.database && state.dbName &&
      freshState.database.toLowerCase() !== state.dbName.toLowerCase()) { legacyPreferenceStale(); return; }
  var ticket = { api: api || state.api, epoch: state.initializationEpoch, database: state.dbName,
    requestedDatabase: freshState && freshState.database };
  state.legacyPendingFocus = ticket;
  state.legacyFocusChecking = true; // also when another report is selected
  var pref = legacyPreferenceOwner();
  if (pref && pref.filters) { pref.verified = false; pref.readToken++; }
  legacyPreferenceSuspend();
  if (state.legacyUnitsResolved) legacyPreferenceVerifyFocus(ticket);
}
function legacyPreferenceVerifyFocus(ticket) {
  if (state.legacyPendingFocus !== ticket || state.initializationEpoch !== ticket.epoch || state.legacyIdentityStale) return;
  if (!state.legacyViewerLogin) { legacyPreferenceStale(); return; }
  var settled = false;
  var timer = setTimeout(function () { finish(null); }, 8000);
  function finish(session) {
    if (settled) return;
    settled = true; clearTimeout(timer);
    if (state.legacyPendingFocus !== ticket || state.initializationEpoch !== ticket.epoch || state.legacyIdentityStale) return;
    var creds = (session && session.credentials) || session || {};
    var login = creds.userName || creds.username;
    var sessionDatabase = creds.database;
    if (!login || login.toLowerCase() !== state.legacyViewerLogin ||
        typeof sessionDatabase !== "string" || sessionDatabase.toLowerCase() !== ticket.database.toLowerCase() ||
        (ticket.requestedDatabase && sessionDatabase.toLowerCase() !== ticket.requestedDatabase.toLowerCase())) {
      legacyPreferenceStale(); return;
    }
    state.legacyPendingFocus = null;
    state.legacyFocusChecking = false;
    var pref = legacyPreferenceOwner();
    if (!pref || !pref.filters) {
      if (legacyIsActive()) { legacyPreferenceStart(); setupLegacyFilters(); }
      else legacyPreferenceStatus();
      return;
    }
    pref.verified = true;
    pref.api = ticket.api;
    if (pref.writeOutcome) legacyPreferenceApplyWrite(pref);
    else if (pref.writing) legacyPreferenceSuspend();
    else if (pref.status === "ready" || pref.status === "loading") legacyPreferenceRead(pref);
    else legacyPreferenceResume();
    if (legacyIsActive()) setupLegacyFilters();
  }
  try { ticket.api.getSession(finish); } catch (err) { finish(null); }
}
