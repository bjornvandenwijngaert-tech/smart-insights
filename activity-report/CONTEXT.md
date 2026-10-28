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
- **A trip has no `startPoint`.** A vehicle starts where it last stopped, so the
  Start Location is inferred from the *previous* fetched trip's `stopPoint`. The
  first fetched trip per vehicle has no predecessor in this result and shows
  "(location unknown)" even if an earlier trip outside the query has coordinates.
  This is a lookup-context limitation, not evidence of missing source GPS.
- **Addresses** are resolved with MyGeotab's own `GetAddresses` — coordinates never
  leave Geotab's infrastructure. Deduped to ~1m via `coordKey()` and chunked 400 per
  call to stay under the 450/min limit. Failures fall back to raw coordinates.
- **Durations** are MyGeotab `TimeSpan` strings (`"1.02:28:55"`), parsed by
  `parseDurationToMins`, which falls back to timestamp math when unparseable.
- **Driver handovers:** days use the whole chronological vehicle timeline.
  Retain every row's original `prevTrip` through handovers and across days.
- **Preview:** each vehicle has a horizontal daily table and one selected-period
  total row underneath. Exactly one row per vehicle/day; no trip itemisation,
  disclosure controls, hidden detail, fleet cards or driver/week folds.
- **Daily endpoints:** earliest trip start and its actual predecessor location,
  latest trip stop and its location. Times use 24-hour `hh:mm`; dates `dd/mm/yyyy`.
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
  hides exports. Publish trips, groups, addresses and `activityRange` together
  after address resolution completes. CSV/PDF/Excel dates and filenames
  summaries use that completed range, not subsequently edited date inputs.
- **Calendar dates:** date inputs use `localDayKey`, matching displayed dates.

## Keep in sync

Every fix to the Activity Report belongs in **both** this file's `src/app.js` and
the parent `../src/app.js`, until the two are formally split into separate repos.
The six historical-driver helpers are identical in the parent, this add-in and
the Legacy Trip History standalone add-in. The regression suite checks this.

## Version

v1.3.4 — remove the Summary worksheet and its chart generation from
Excel export. The workbook contains only Activity Report with daily rows and
vehicle totals. Mirrored to Smart Insights v2.12.4.

## Data validation status — 28 October 2026

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

Known location semantics: the app produces "(location unknown)" for an absent or
incomplete endpoint coordinate; valid coordinates with a failed GetAddresses
lookup fall back to latitude/longitude. First fetched starts lack predecessor
context. Unknown final locations require checking that final Trip.stopPoint.

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
