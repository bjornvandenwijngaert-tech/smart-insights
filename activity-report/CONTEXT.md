# Activity Report — standalone add-in

A self-contained MyGeotab add-in containing only the Activity Report. It is a lift
of the Activity Report functions from the parent Smart Insights `src/app.js`, so it
can be installed on its own by customers who don't want the full dashboard.

The report is grouped as **vehicle -> local start day -> optional trip detail**.
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
  Start Location column comes from the *previous* trip's `stopPoint`. The first
  trip of the fetched range genuinely has no origin and shows "(location unknown)".
- **Addresses** are resolved with MyGeotab's own `GetAddresses` — coordinates never
  leave Geotab's infrastructure. Deduped to ~1m via `coordKey()` and chunked 400 per
  call to stay under the 450/min limit. Failures fall back to raw coordinates.
- **Durations** are MyGeotab `TimeSpan` strings (`"1.02:28:55"`), parsed by
  `parseDurationToMins`, which falls back to timestamp math when unparseable.
- **Driver handovers:** days use the whole chronological vehicle timeline.
  Retain every row's original `prevTrip` through handovers and across days.
- **Preview:** each vehicle has a horizontal daily table and one selected-period
  total row underneath. All daily summaries are visible; native `details` expands
  an individual day's trips. No fleet cards, driver/week folds or driver labels.
- **Daily endpoints:** earliest trip start and its actual predecessor location,
  latest trip stop and its location. Times use 24-hour `hh:mm`; dates `dd/mm/yyyy`.
- **Total Hours:** first departure to latest arrival, using actual elapsed UTC
  time, including gaps between trips. Vehicle Total Hours sums daily spans, not
  the elapsed time across the whole report range. Overnight trips belong to their
  local start day. Stop duration remains the raw post-trip duration, which can
  include overnight parking; engine-on remains driving + idle.
- **Exports:** CSV/PDF contain daily rows and one vehicle footer, no driver column.
  Excel opens on Activity Report with those same daily rows visible, raw trip
  rows hidden underneath, and vehicle totals below. Outline levels are vehicle /
  day / trip (1/2/3). The previous Summary/chart remains a secondary worksheet.
  Hours/stop/idle/driving/work figures use decimal hours; engine-on uses `[h]:mm`;
  distance follows the user's profile. Excel dates/times are native date values.
- **Excel formulas:** other metric totals use `SUBTOTAL(9,...)` over raw descendants.
  Daily elapsed hours use `24*(MAX(end UTC)-MIN(start UTC))`; vehicle elapsed hours
  use SUMIF on daily rows only. Three hidden helper columns retain raw UTC trip
  starts, UTC stops and row type so DST stays correct without double-counting.
  Per-row `collapsed` overrides ExcelJS 4.4's sheet-wide getter: days start
  collapsed, vehicles start expanded. Verify OOXML and desktop Excel on upgrades.
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

v1.3.0 — clean vehicle/day layout with a horizontal row per day and
one vehicle-period footer; no driver allocation or intermediate weekly totals.
Expandable trip detail, matching CSV/PDF, and Excel daily summaries visible by
default. Completed-range export locking retained. Mirrored to Smart Insights v2.12.0.

Regression checks: `node --test tests/trip-report-drivers.test.cjs` from the repo root.

Bump `config.json`, the `version-badge` in
`index.html`, and the `?v=` cache-busting query on all three asset tags together.
