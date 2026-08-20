# Smart Insights — Session Context

## What this is
A MyGeotab add-in hosted on GitHub Pages, registered via `config.json`.
Hosted at: https://bjornvandenwijngaert-tech.github.io/smart-insights/index.html
Registered under: `ActivityLink/` path in MyGeotab.
Support email: bjornvandenwijngaert@geotabsmb.com

---

## File structure
```
smart-insights/
  index.html        — App shell, modal HTML, CDN scripts
  config.json       — MyGeotab add-in registration
  src/
    app.js          — All logic (state, API calls, rendering, exports)
    styles.css      — All styling
    logo-data.js    — Auto-generated: Geotab logo as a base64 data URL constant
                      (GEOTAB_LOGO_DATAURL) used in PDF exports. Loaded before app.js.
    geotab-logo.png — Source Geotab wordmark (707x130, navy #27325D bg). Regenerate
                      logo-data.js from this if the logo is ever swapped.
  CONTEXT.md        — This file
```

---

## Tech stack
- **MyGeotab Add-in API** — `geotab.addin.smartInsights` lifecycle (initialize / focus / blur)
- **Chart.js 4.4.0** — Stacked-per-vehicle bars + per-vehicle lines in dashboard widgets
- **Leaflet 1.9.4 + OpenStreetMap tiles** — Infraction map (tab + drill-down mini-map)
- **SheetJS (xlsx 0.18.5)** — Excel export
- **jsPDF 2.5.1 + AutoTable 3.8.2** — PDF export
- All via CDN, no build step, no npm

---

## Current version: v2.7.0

### Changelog since v2.2.0
- **v2.7.0** — **Legacy Trip History Report** (Reports tab, `legacy-trip-history`).
  Recreates a classic per-vehicle/per-day trip-log export (start time, distance,
  driving duration, reverse-geocoded stop address, arrival time, idle duration,
  stop duration, daily totals). Customer sets the range via a **length selector**
  (Daily / Weekly / Monthly / Custom) that auto-fills the existing from/to date
  inputs. Addresses are resolved via MyGeotab's own **`GetAddresses`** API
  (`Trip.stopPoint` \u2192 coordinates \u2192 `GetAddresses` \u2192 `ReverseGeocodeAddress`) \u2014
  **not** a third-party geocoder, so coordinates never leave Geotab's
  infrastructure and there's no external rate-limit bottleneck (450 req/min,
  batched up to 400 unique coordinates per call, deduped by ~1m precision so
  repeat stops like depots only cost one lookup). Exports as a branded PDF
  (`Export PDF` button, jsPDF + AutoTable) matching the reference format's
  structure: header, dark "Report Totals for" bar, 2\u00d73 KPI grid, then one
  table per vehicle per day with a "Starting from" lead-in row, an
  "(Ignition On)" gap row, trip rows, and a daily total row. Also has a flat
  CSV export and an in-app HTML preview.
  **Geotab logo branding**: both PDF exports (Legacy Trip History + drill-down)
  carry the Geotab wordmark top-right in the header. The logo PNG has a navy
  (#27325D) background, so both PDF header bars are drawn in that same navy
  (`GEOTAB_NAVY`) for a seamless fit (previously bright blue #0078D4; KPI boxes
  and table header rows stay accent-blue). Logo is embedded as a base64 data URL
  in `src/logo-data.js` (no external fetch, works offline). Helper:
  `drawPdfHeaderLogo()`. Excel export is intentionally left unbranded \u2014 the
  SheetJS Community build in use cannot embed images.
  **Known simplification**: the reference format splits a "Working Total" row
  (work-hours only) from a separate "Total" row (whole day incl. after-hours
  idling), using Trip's `Work*`/`AfterHours*` fields. v1 collapses this into a
  single combined "Total" row per day. Revisit if the customer needs that split.
  **Bug fixes (post-initial-build):**
  - `parseDurationToMins()` originally only handled ISO-8601 durations, but
    MyGeotab serializes TimeSpan as .NET strings ("hh:mm:ss" / "d.hh:mm:ss").
    That made every `stopDuration`/`idlingDuration` fail to parse \u2192 stop durations
    showed "\u2014" (looked like "no stops") and idle showed 0m everywhere. Parser now
    handles .NET TimeSpan, ISO-8601, and numeric-seconds formats.
  - The "(Ignition On)" lead-in row used to print the gap between the previous
    trip's stop and the day's first drive in the Idle column \u2014 which for an
    overnight park showed as e.g. "17h 08m idle". It now shows "\u2014": that gap is
    parked/overnight time, not idle, and Trip data exposes no pre-drive idle value.
  - NOTE: because Geotab's `StopDuration` runs to the *next* trip's start, the
    last trip before an overnight legitimately carries a very large stop duration,
    which inflates the "Average Stop Duration" KPI. Left as-is (accurate, matches
    the reference tool); revisit if the customer wants cross-day gaps excluded.

  **Distance / MPG units fix (app-wide).** `Trip.Distance` is in KILOMETRES per
  the MyGeotab SDK, but the whole codebase had been treating it as metres
  (`/1609.344` for miles, `/1000` for km, and `calcMpg` dividing by 1609.344).
  That made every distance and MPG figure ~1000x too small (a multi-hour drive
  showed ~0.01 mi). Fixed across ALL reports (Daily Trips, Monthly Carbon,
  Speeding, Daily Fuel Economy, Legacy Trip History) and every surface
  (in-app tables, CSV, PDF, Excel, drill-downs, dashboard widgets): km->mi now
  uses `* 0.621371`, `calcMpg()` takes km, and Daily Trips no longer divides by
  1000. Odometer conversions (`nextOdometerReading / 1000`) are left as-is \u2014
  Odometer genuinely is in metres. Speed (`maximumSpeed`->mph via `toMph`) was
  always correct and is unchanged.

  **Standalone extraction.** The Legacy Trip History Report also exists as a
  self-contained standalone add-in under `legacy-trip-report/` (its own
  config.json, index.html, src/app.js, styles.css, logo). It's a lift of the
  legacy-report functions from this file for pitching to Geotab / installing on
  its own. See `legacy-trip-report/CONTEXT.md`. KEEP THE TWO IN SYNC until they
  are formally split into separate repos.

  The **Activity Report** now has the same treatment under `activity-report/`
  (own config.json, index.html, src/app.js, styles.css, logo). Same rule: KEEP IT
  IN SYNC with the Activity Report functions in `src/app.js`. See
  `activity-report/CONTEXT.md`.

  **Scheduled email reports \u2014 investigated, not built (deferred).** A MyGeotab
  add-in is client-side only: it runs solely while a user has the page open, with
  no server, background process, or cron. Scheduled emailing is therefore
  impossible from the add-in alone \u2014 it needs a server-side backend (e.g. a Cloud
  Run Function on Cloud Scheduler that authenticates to the MyGeotab API, renders
  the PDF server-side, and emails via SendGrid), with the add-in optionally
  providing a schedule-config UI stored in AddInData. Deferred by customer choice;
  only the bug fixes above were shipped this round.
- **v2.3.0** — Dashboard widgets snap to avoid overlap while dragging (live snap on
  `mousemove`). `computeSnapPosition` is a pure function that reads every other
  widget's real geometry via `getBoundingClientRect()` (grid-local), finds the
  nearest free slot along the axis closest to the dragged tile's centre, 8px gap.
- **v2.4.0** — Dashboard layout persisted in MyGeotab **AddInData** (`addInId:
  "SmartInsightsDashboard"`) instead of localStorage, so the layout is shared across
  all users with add-in access and is per-database. First load migrates any old
  localStorage layout then clears it. localStorage remains a fallback for local dev
  (no `state.api`).
  **This did not fix the cross-database bleed.** The claim above was wrong. AddInData
  was put in front of localStorage but localStorage was never removed from the code
  paths that matter, and the migration path actively spreads the bleed. Still broken as
  of 2026-08-19. See "Open bugs" below before touching persistence.
- **v2.5.0** — Charts + map. All event widgets now render **stacked-per-vehicle bars**
  (distinct colour per vehicle) with **adaptive granularity** (per-day, or per-hour when
  the range ≤ 24h). Clicking a vehicle segment opens the drill-down pre-filtered to that
  vehicle + time bucket. Fuel economy is **per-vehicle lines** (MPG is an average, cannot
  stack). New **Map tab** plots infraction locations, and the drill-down has a **Table/Map
  toggle** for exception/speeding. Map uses **Leaflet + OpenStreetMap** (CDN).
- **v2.6.0** — **Incident replay** modal (`#incident-modal`). A "Replay" button on each
  exception drill-down row opens a per-event reconstruction: LogRecord path on a Leaflet
  map (grey cause/aftermath + red `activeFrom→activeTo` infraction segment, start/end +
  scrub marker) synced to a Chart.js **speed line** (event portion red). Hovering/tracing
  the speed graph moves the marker along the route; readout shows time · mph · INFRACTION.
  Window = event ±60s. Needs ≥2 GPS points, else shows a "not enough GPS data" message.

---

## Open bugs

### Dashboard layout bleeds between databases (open, high priority)

**Reported 2026-08-19, confirmed still present.** A dashboard layout built in one
database shows up in another. v2.4.0 claimed to fix this and did not. Treat the
changelog entry as unreliable and this section as the current truth.

Not yet reproduced against two live databases. What follows is read from the code in
`src/app.js:3515-3622`, so the mechanisms are real but which one is firing has not
been confirmed.

**Why AddInData alone did not fix it.** AddInData genuinely is per-database, so the
happy path is correct. The problem is that localStorage was left in place underneath
it, and the localStorage key carries no database identity:

```js
localStorage.setItem("smartinsights-dashboard", details);
```

Add-in pages are injected into the MyGeotab page, not iframed, so localStorage is
scoped to the MyGeotab origin (`my.geotab.com`). Every database on that server shares
one origin, therefore one key. Anything written there is visible to every database the
user opens on that server. This is the bleed vector and it is still live.

**Four candidate causes, in the order worth checking.**

1. **Stale `state.dashboardDataId` across a database switch** (`3544`, `3602`).
   The AddInData record id is cached on `state`. Switching database in the MyGeotab UI
   is client-side navigation and may not tear down the add-in's JS scope, which is
   shared across pages by design. If the id survives the switch, `saveDashboard` takes
   the `Set` branch and writes database B's layout into database A's record. This is
   the strongest suspect because it needs no localStorage involvement at all and it
   matches "persistent between databases" exactly.

2. **The migration path plants one database's layout into another** (`3606-3613`).
   On finding no AddInData record, the code restores whatever is in localStorage and
   then saves it as *this* database's record. One stale localStorage entry from
   database A therefore becomes database B's permanent saved layout. The
   `removeItem` on `3612` fires after the damage, and it is also ordered wrong: it runs
   synchronously while the `Add` on `3554` is still in flight, and if that `Add` fails
   the error handler on `3561` writes the key straight back.

3. **Silent localStorage fallbacks on any API failure** (`3561`, `3619`).
   A failed `Add` or `Get` drops to the shared key with only a `console.warn`. The user
   sees a layout that looks saved and has in fact been written somewhere every other
   database can read. Failures are invisible in normal use.

4. **`Get` takes `results[0]` unconditionally** (`3600-3602`).
   Concurrent first-saves by different users can create more than one record for the
   same `addInId`. Whichever comes back first wins, silently. Not the reported symptom
   but it will produce divergent layouts between users in the same database.

**Direction when this gets worked on** (not decided, do not treat as a plan):
namespace any localStorage key by database and clear the cached record id whenever the
database changes, rather than adding more fallbacks. The database name is already read
for the db badge. Bear in mind that removing the fallback entirely means a failed
AddInData call loses the layout, so the failure needs to become visible to the user
instead of a `console.warn`.

**Verification this needs, which no previous fix had.** Two real databases, a distinct
layout saved in each, then switch between them both with and without a full page
reload, and confirm each keeps its own. The v2.4.0 fix was reasoned about rather than
tested across databases, which is how a wrong claim reached the changelog.

---

## What's been built

### Three tabs
1. **Dashboard** — Draggable/resizable widgets. Exception-rule widgets (added from Suggestions) + 5 preset widget types (added via "+ Add Widget" picker)
2. **Reports** — 5 report types: Daily Trips, Monthly Carbon, Speeding >80mph, Daily Fuel Economy, **Legacy Trip History Report**. Each with summary KPIs, filterable table, CSV export, vehicle filter. Legacy Trip History additionally has a length selector (Daily/Weekly/Monthly/Custom) and a branded PDF export. (Upcoming Maintenance + Maintenance Spend removed — see Maintenance data constraint below.)
3. **Suggestions** — Top 8 exception rules ranked by violation count, with intelligence overlays

### Three preset dashboard widget types
| Type | Chart | Stats | Data source |
|------|-------|-------|-------------|
| Monthly Carbon Report | Bar (12 months CO₂ kg) | Total CO₂ / Fleet MPG / Best MPG | Trip.fuelUsed × 2.4 kg/L |
| Speeding >80mph | Bar (daily count, 30d) | Total / Per Day / Worst Day | Trip.maximumSpeed > 128.747 km/h |
| Daily Fuel Economy | Line (daily MPG, 30d) | Fleet MPG / Best Day / Worst Day | Trip.fuelUsed + Trip.distance |

(Upcoming Maintenance + Maintenance Spend removed from the picker — see Maintenance data constraint below.)

### Widget picker
- "+ Add Widget" button in dashboard toolbar opens a picker modal
- All 5 preset types available as cards with icon, name, description
- Existing Suggestions → "+ Dashboard" flow unchanged

### Suggestions tab intelligence
- Fetches 4 resources in parallel: devices, rules, current 30-day exceptions, previous 30-day exceptions
- **Fleet health score** (0–100) displayed at top: computed from violations-per-vehicle rate, colour-coded green/amber/red
- Each suggestion card shows:
  - SVG icon (category-matched, no emojis)
  - Severity badge: Critical (≥30 violations) / Warning (≥10) / Low (<10)
  - Trend badge: % change vs previous 30-day period (↑ bad / ↓ good / Stable)
  - Top 3 offending vehicles with mini bar chart
  - "+ Dashboard" button to pin to dashboard
- Colour-coded left border per severity

### Dashboard
- Widgets persist via `localStorage` (key: `smartinsights-dashboard`)
- **Edit Layout mode** (button in toolbar):
  - Activates drag and resize on widgets
  - Full-screen dot grid appears on `body` (`body.edit-mode-active`)
  - Remove (×) button on widgets only visible in edit mode
  - Dashed blue outline on dashboard canvas
- **View mode** (default):
  - Widgets are locked, no handles visible
  - Clicking the chart/stats area of a widget opens the drill-down modal
  - Header area click does NOT trigger drill-down
- Each widget has a stats bar: Total / Daily Avg / Worst Day

### Drill-down modal
Opens when any widget is clicked in view mode. Filters and columns are type-aware.
- **Exception type**: date range / vehicle / driver filters → Vehicle, Driver, Date, Time, Duration table
- **Carbon type**: date range / vehicle filters → Vehicle, Date, Distance, Fuel, MPG, CO₂ table
- **Speeding type**: date range / vehicle / driver filters → Vehicle, Driver, Date, Time, Max Speed, Distance table
- **Fuel Economy**: date range / vehicle filters → Vehicle, Date, Time, Distance, Fuel, MPG table
- (Maintenance Upcoming / Maintenance Spend drill-downs no longer reachable — options removed)
- **Export options** (enabled after data loads): Excel (.xlsx) two-tab workbook + PDF with blue header, KPI boxes, AutoTable, "Confidential" footer

---

## Design decisions & philosophy

### Colour palette (Geotab-aligned)
```
--accent:       #0078D4   (Geotab blue)
--accent-dark:  #005A9E
--accent-dim:   #E8F4FD
--critical:     #DC2626   (red)
--warning:      #D97706   (amber)
--low:          #059669   (green)
```

### Target customer
Small-to-mid fleet customers (Verizon Connect migration context).
**Key product philosophy**: These customers have no BI tools, no data analysts.
The add-in must do the interpretation work FOR them.
- Always show vehicle names, never device IDs
- Always show readable dates/times, never ISO strings
- PDF and Excel exports are formatted REPORTS, not raw data dumps
- Two layers: summary (at a glance) → detail (on demand, via drill-down)
- CSV intentionally excluded — it serves data pipeline use cases, not this audience

---

## Maintenance data constraint (why the two maintenance reports were removed)
Investigated 2026-07 against live DB `TRAI02`:
- MyGeotab's classic SDK has **no** `WorkOrder`, `MaintenanceReminder`, or
  `MaintenanceRecord` entity/method. The old code called `Get typeName:"MaintenanceReminder"`,
  which does not exist — so those reports never returned data on any real DB.
- **Work Order Management** is a next-gen (GreatLakes) first-party app
  (`solutionId: geotabWorkOrderManagement`, `app.geotab.com/apps/maintenance-work-request`,
  appMenuId `maintenanceWorkRequest`). Its records live in a next-gen maintenance
  service, not in classic `AddInData`. Querying `AddInData` as a normal user returns
  only UI/map-view prefs — the work order records are not exposed to a classic add-in
  via `api.call("Get", ...)`, GetFeed, or MultiCall.
- **Conclusion**: a classic add-in cannot read Work Order Management data today.
  Both maintenance widget/report options were removed (handler functions retained as
  dead code in app.js for quick restore). Revisit if Geotab ships a public Maintenance
  Center API. Proper support would warrant its own dedicated page — out of scope for
  the current (12-unit) customer.

## Known placeholders / next logical features
- Reports tab: Exceptions and Fuel reports marked `disabled` / "coming soon"
- Drill-down filters currently filter client-side after fetching all events for the date range
- Driver map (`state.driverMap`) is populated from `User` fetch but display name logic may need refinement depending on how driver names are stored in the customer's database
- The health score formula is simple (violations / devices / 50 * 100) — could be refined with weightings per severity

---

## API calls in use
All via `state.api.call(method, params, onSuccess, onError)` and `state.api.multiCall`.
```
Get → Device         (device list / names)
Get → User           (driver list for drill-down filter)
Get → Rule           (exception rule list + map rule filter; names cached in state.ruleNameMap)
Get → ExceptionEvent (violations — current/previous period, widget data, drill-down, map)
Get → Trip           (reports tab, speeding/carbon/fuel widgets)
Get → LogRecord      (GPS correlation for the map — batched via multiCall, 50/batch)
GetAddresses         (reverse geocode Trip.stopPoint coordinates for Legacy Trip
                       History Report; official MyGeotab API, not a third-party
                       geocoder — see v2.7.0 changelog)
Add/Set/Get → AddInData (dashboard layout persistence; addInId "SmartInsightsDashboard")
```
**Map data constraint**: ExceptionEvent/Trip carry NO coordinates. The map correlates
each event to the vehicle's GPS trail by querying LogRecord for that device in a ±3 min
window around the event time and taking the nearest point. Capped at 500 most-recent
events (`MAP_EVENT_CAP`), batched 50 per multiCall.

---

## Repo / deployment
- GitHub repo: bjornvandenwijngaert-tech/smart-insights
- Deployed via GitHub Pages
- **Push is done MANUALLY by Bjorn** each time a new version is created. GIA does not
  push; just tell GIA when a push has been done so context stays accurate. (This local
  folder has no git remote configured — the push happens outside it.)
- config.json points directly to the Pages URL
- Cache-bust: all asset URLs include a `?v=X.Y.Z` query string — **increment on each release**
  (this maintenance-removal change needs a version bump + cache-bust before it takes effect)
