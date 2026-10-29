# Activity Report — standalone add-in

A self-contained MyGeotab add-in containing only the Activity Report. It is a lift
of the Activity Report functions from the parent Smart Insights `src/app.js`, so it
can be installed on its own by customers who don't want the full dashboard.

The report is grouped as **vehicle -> local start day totals only**.
Each day combines every trip for that vehicle, regardless of driver allocation.
There are no driver or week sections, and Activity does not resolve driver users.
Vehicle grouping uses stable device ids, not display names.

Sibling of `../legacy-trip-report/`, which does the same for the Legacy Trip
History Report.

## Install

Register this URL in MyGeotab (Administration → System → System Settings →
Add-Ins → New Add-In, paste the JSON or the URL):

```
https://bjornvandenwijngaert-tech.github.io/smart-insights/activity-report/config.json
```

It appears in the MyGeotab menu under **Activity** as "Activity Report".

## Files

| File | Notes |
|---|---|
| `config.json` | Add-in manifest. `path: "ActivityLink/"`, menu name "Activity Report". |
| `index.html` | Single page. Loads Chart.js, ExcelJS, jsPDF + AutoTable from CDN. No Leaflet, no XLSX (this report has no map and uses ExcelJS, not SheetJS). |
| `src/app.js` | The whole add-in. |
| `src/styles.css` | Copied verbatim from the parent `src/styles.css`. |
| `src/activity.css` | Scoped horizontal daily-row and vehicle-footer styles, also loaded by the parent. |
| `src/logo-data.js` | Base64 Geotab logo for the PDF header. Copied verbatim. |
| `src/geotab-logo.png` | Source of the above, kept for regeneration. |

## Deliberate differences from the parent

- **Wrapped in an IIFE.** MyGeotab injects add-in pages into its own document
  rather than iframing them, so every installed add-in shares one JS global scope.
  Without the wrapper, `var state` / `apiCall` / `fmtDist` here would collide with
  the parent Smart Insights add-in when both are installed on the same database.
  The registry is reached as `window.geotab.addin.activityReport` — a bare
  `var geotab` fallback inside the IIFE would hoist, shadow the real global, and
  silently register the add-in on a throwaway object.
- **Own localStorage key**, `activity_report_standalone_settings`, for the same
  shared-origin reason.
- **Separate Activity builder:** `buildActivityByDevice` is identical in both
  Activity apps. The parent's driver-based Legacy builder is independent.
  Toolbar element id `legacy-length` → `range-length`.
- **Omitted:** `OVERNIGHT_MIN_HOURS`, `buildLegacyStopInfo`, `legacyStopMins`. Those
  suppress overnight parks for the Legacy report only; the Activity Report reports
  raw `Trip.stopDuration` by design, so the last stop of a day carries the full
  gap until the next morning's first trip.
- **Units trimmed** to distance only (`MI_PER_KM`, `loadUserUnits`, `isMetric`,
  `distUnit`, `distVal`, `fmtDist`). The Activity Report shows no speed, fuel
  economy or threshold, so those helpers aren't carried over.

## Things worth knowing before editing

- **Units follow the logged-in user, not the database.** MyGeotab returns
  everything metric; `User.isMetric` decides display. `loadUserUnits()` reads it via
  `getSession` → `Get User`, with an 8s timeout guard falling back to
  `SystemSettings.measurementSystem` (database-wide, so it can disagree with the
  user's own profile — it's a last resort, not the source of truth).
- **A trip has no `startPoint`.** Reuse a valid previous fetched stop only when
  it precedes departure. Missing daily starts (including the first fetched trip)
  and missing final stop points are looked up with `Get LogRecord`, scoped to the
  real Device.id and endpoint UTC instant. The positive 1ms search requests an
  interpolated boundary point; accept only a valid coordinate at that instant,
  never an arbitrary older/other-device point. Interpolated records can lack id.
- **Addresses** are resolved with MyGeotab's own `GetAddresses` — coordinates never
  leave Geotab's infrastructure. Only displayed daily endpoints are geocoded,
  deduped to ~1m via `coordKey()` and chunked 400 coordinates/call. GPS requests are
  paced at most 10/sec (<1000/min); geocoder requests at most one/150ms (<450/min).
  Both use a 15s once-only timeout/error guard. Empty GPS results remain unknown;
  actual GPS failures show "location lookup failed". Address failures preserve
  coordinates with an explicit "address lookup failed" label in every output.
- **Durations** are MyGeotab `TimeSpan` strings (`"1.02:28:55"`), parsed by
  `parseDurationToMins`, which falls back to timestamp math when unparseable.
- **Driver handovers:** days use the whole chronological vehicle timeline.
  Retain every row's original `prevTrip` through handovers and across days.
- **Preview:** each vehicle has a horizontal daily table and one selected-period
  total row underneath. Exactly one row per vehicle/day; no trip itemisation,
  disclosure controls, hidden detail, fleet cards or driver/week folds.
- **Daily endpoints:** earliest trip start, latest trip stop and their resolved
  locations. Times use 24-hour `hh:mm`; dates `dd/mm/yyyy`.
- **Total Hours:** first departure to latest arrival, using actual elapsed UTC
  time, including gaps between trips. Vehicle Total Hours sums daily spans, not
  the elapsed time across the whole report range. Overnight trips belong to their
  local start day. Stop duration remains the raw post-trip duration, which can
  include overnight parking; engine-on remains driving + idle.
- **Exports:** CSV/PDF contain daily rows and one vehicle footer, no driver column.
  Excel opens on Activity Report with those same daily rows and vehicle totals
  below. No individual-trip rows exist anywhere in the workbook, including hidden
  rows or sheets. Outline levels are vehicle / day (1/2), initially expanded.
  The workbook contains only Activity Report; there is no Summary worksheet or
  generated chart. This export no longer needs Chart.js or a canvas.
  Hours/stop/idle/driving/work figures use decimal hours; selected engine-on,
  driving-only and idle-only reporting columns use `[h]:mm`;
  distance follows the user's profile. Excel dates/times are native date values.
- **Row emphasis:** daily data is plain; only vehicle-total data rows are bold.
  Excel's column header retains its existing bold blue style. This applies to
  every reporting-selector combination.
- **Excel formulas:** daily metrics are the aggregated source facts. Daily elapsed
  hours use `24*(day end UTC - day start UTC)`; engine-on uses `(idle + driving)/24`.
  Vehicle totals SUM only the contiguous daily rows. Two hidden columns hold each
  day's first departure and final arrival in UTC, not individual-trip data, so
  overnight/DST remains correct. All exported rows are visible initially, and
  only vehicle groups can collapse. Verify OOXML and Excel on library upgrades.
- **Column controls:** preserve the original Engine On (total), Driving Time,
  Idle Time and Work/After-hours selectors. They choose additional aggregate
  reporting columns, independently of the daily-only row grain. The base daily
  stop/idle/driving-hour totals remain base summary facts. All renderers and
  exports share `activitySummaryColumns` for order, values, formats and widths.
  Engine On is checked by default when no saved array exists; saved arrays are
  respected exactly, including empty and work-only selections. Column choices
  apply on Run, as before. Daily-only cleanup does not remove reporting choices.
- **Filtering:** vehicle filters update every scoped summary and export.
- **Result snapshot:** starting a run clears previous exportable Activity data and
  hides exports. Publish trips, groups, the endpoint/address snapshot and
  `activityRange` together after resolution completes. GPS jobs use device + UTC
  instant keys; endpoint keys add start/end role. No enrichment is attached to
  source trips or rebuilt day/row objects. Superseding runs, report-type changes
  and leaving a pending add-in cancel obsolete callbacks and queued work.
  CSV/PDF/Excel dates and filenames
  summaries use that completed range, not subsequently edited date inputs.
- **Calendar dates:** date inputs use `localDayKey`, matching displayed dates.

## Keep in sync

Every fix to the Activity Report belongs in **both** this file's `src/app.js` and
the parent `../src/app.js`, until the two are formally split into separate repos.
The six historical-driver helpers are identical in the parent, this add-in and
the Legacy Trip History standalone add-in. The regression suite checks this.

## Version

v1.3.5 — recover missing displayed endpoint coordinates using native
LogRecord boundary searches before GetAddresses. One completed snapshot supplies
HTML/CSV/PDF/Excel; source trip facts and report layout are preserved. Mirrored to
Smart Insights v2.12.5. Local implementation tests pass. Published at the user's
request for live testing; customer/native verification remains pending.

## Data validation status — 29 October 2026

Automated/synthetic fixtures, actual browser rendering and desktop Excel checks
verify implementation, formulas, selector combinations and formatting. They do
not establish parity with this user's live demo data or native Trips History.
The live demo database, request window, browser timezone and units were identified.
The report includes complete overnight trips attributed to their local start day.
The logged-in MyGeotab profile timezone and unrounded Trip/native-export data
still need verification. Do not infer profile timezone from a database name or
addresses. Case-specific source observations are kept in private scratch notes,
not in this public-hosting repository.

The browser pairing was repaired and the live add-in was observed, but subsequent
source reads failed with disconnections/timeouts; CLI credential import also
failed. The existing CLI session was expired and pointed elsewhere. Native Trips
History/source parity was not completed and was deferred by the user. Resume
with a stable signed-in connection or the Activity workbook plus native Trips
History export for the same window, vehicles and units.

Location correction — 29 October 2026: the customer single-day page confirmed
resolved destination addresses and unknown first departures, matching the missing
predecessor context. Version 1.3.5 adds GPS lookup for missing endpoints.
Its regression fails behaviorally against v1.3.4 and passes in both candidate apps.
The new fixtures remain synthetic control-flow checks. Real LogRecord producer
output -> resolver -> real GetAddresses -> customer/native comparison
is still required, including verifying the 1ms boundary query in the target runtime.
Browser access subsequently timed out and MyGeotab authentication requires user
action. The user explicitly requested publication and will perform the live test.
Do not present the release as customer/native-validated until that check completes.

Review also needs whole overlapping/overnight trips, post-trip stop attribution,
rounding, browser-local dates versus MyGeotab profile timezone, derived engine-on
versus the cumulative engineHours field, and whether work-idling fields actually
exist in the database's Trip responses. Do not claim source/native parity until
those live comparisons have been performed. No data-fetch behavior changed for
the row-emphasis correction.

Displayed-number spot checks are not source validation: sums of rounded daily
values may differ from a vehicle total calculated from full-precision values.

Regression checks: `node --test tests/trip-report-drivers.test.cjs` from the repo root.

Bump `config.json`, the `version-badge` in
`index.html`, and the `?v=` cache-busting query on all three asset tags together.
