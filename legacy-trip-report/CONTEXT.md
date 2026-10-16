# Legacy Trip History — Standalone Add-in — Session Context

## What this is
A **single-purpose** MyGeotab add-in that produces one thing: the Legacy Trip
History Report (a classic per-vehicle/per-driver/per-day trip log with reverse-geocoded
stop addresses, CSV + branded PDF export). It is a self-contained extraction of
the Legacy Trip History feature from the parent **Smart Insights** add-in (one
level up), created so the report can be pitched to Geotab and installed on its
own without the Dashboard / Suggestions / Map surface.

Hosted (intended): https://bjornvandenwijngaert-tech.github.io/smart-insights/legacy-trip-report/index.html
Registered under: `ActivityLink/` path in MyGeotab.
Add-in lifecycle key: `geotab.addin.legacyTripHistory`.
Support email: bjornvandenwijngaert@geotabsmb.com

---

## File structure
```
legacy-trip-report/
  config.json       — MyGeotab add-in registration (name "Legacy Trip History")
  index.html        — Minimal shell: report toolbar + output. Loads jsPDF +
                      jsPDF-autotable (NO Chart.js / Leaflet / SheetJS — not needed)
  src/
    app.js          — Only the legacy-report logic + its dependencies
    styles.css      — Copied whole from the parent app (contains unused rules for
                      widgets/modals/map; safe, can be trimmed later)
    logo-data.js    — Geotab logo as a base64 data URL (GEOTAB_LOGO_DATAURL)
    geotab-logo.png — Source Geotab wordmark (707x130, navy #27325D bg)
  CONTEXT.md        — This file
```

---

## Tech stack
- **MyGeotab Add-in API** — `geotab.addin.legacyTripHistory` lifecycle (initialize)
- **jsPDF 2.5.1 + AutoTable 3.8.2** — branded PDF export
- Plain JS for CSV export; no build step, no npm, all via CDN

---

## How the report works
1. Fetch `Device` (name map) + `Trip` for the selected range.
2. Resolve each unique historical `Trip.driver.id` with `Get User`, using
   `UserSearch.userIds` in batches of 100. Do not filter to currently active users
   or currently enabled drivers; archived users may own trips in this range.
3. Group by `Device.id`, then driver identity, then local day. Device names are
   display labels only. Build the chronological vehicle timeline before grouping,
   so start locations and overnight detection still use the actual prior trip.
4. Reverse-geocode every unique `Trip.stopPoint` via MyGeotab's **`GetAddresses`**
   method (coordinates never leave Geotab; deduped to ~1m precision; batched up
   to 400/call under the 450/min limit). No third-party geocoder.
5. Render per-vehicle/per-driver/per-day tables: "Starting from" lead-in, "(Ignition On)"
   marker, trip rows, daily total. KPIs across the whole range.
6. Export CSV (one trip per row, with that trip's singular `Driver`) or branded
   PDF (Geotab logo, navy header, KPI grid, vehicle/driver/day sections). PDF table
   headers repeat the vehicle, driver and date on continuation pages.

Length selector (Daily / Weekly / Monthly / Custom) auto-fills the from/to date
inputs; Custom leaves them for manual entry. Optional vehicle filter.

## Current version: v1.3.0

- Driver names now come from the `User` referenced by each trip, not the
  non-standard `Trip.driverName` field or the asset's current driver.
- Unknown/unassigned trips have their own section. A known driver id whose name
  cannot be read is labelled `Driver name unavailable (...)`, not unassigned.
- Grouping is by driver id, so different users with the same display name remain
  separate. Their login/id disambiguates their headings when necessary.
- Vehicle filtering now updates summary cards as well as tables and exports.
- Date inputs use local calendar dates, avoiding UTC/BST preset shifts.
- Regression checks: `node --test tests/trip-report-drivers.test.cjs` from the
  repository root. Tests cover both standalone reports and both parent reports.

---

## Important shared-code note
This is a **lift** of the legacy-report functions from the parent
`../src/app.js`. If you fix a bug in one, mirror it in the other until they are
formally split into separate repos. Functions kept in sync include:
`parseDurationToMins`, `fmtDurWhole`, `fmtDurPrecise`, `milesFromDistance`,
`buildLegacyByDevice`, `resolveStopAddresses`, `formatReverseGeocodeAddress`,
`addressForPoint`, `buildLegacyDayBlocks`, `renderLegacyTripHistoryOutput`,
`exportLegacyTripHistoryPdf`, `drawPdfHeaderLogo`, `localDayKey`,
`buildLegacyStopInfo`, `legacyStopMins`, `OVERNIGHT_MIN_HOURS`,
`resolveHost`, `cardDatePart`, `tripHistoryUrl`, `tripHistoryCell`, `esc`.
The historical driver helpers (`reportDriverId`, `isUnassignedDriver`,
`reportUserName`, `reportDriverInfo`, `resolveReportDrivers`,
`groupReportDriverDays`) are also copied into the standalone Activity Report.
Keep those three copies identical; the regression suite checks this.

**`esc()` must escape single quotes.** Every attribute here is written with
single-quoted delimiters and the rison Trip History URLs quote their dates with
apostrophes, so an `esc()` that leaves `'` alone truncates every href at
`dateRange:(endDate:` and the link opens a bare map. All three copies of this
code have now hit that bug; do not "simplify" the escaper.

The Trip History deep-link helpers have a **third** copy, in the Activity Report
add-in at `~/repos/vehicle-activity-log` (`replayUrl`, `resolveHost`,
`cardDatePart`), which is where the URL shape was reverse engineered and is the
one to check first if links stop working. Two deliberate differences here:
that repo builds times from the MyGeotab **profile** timezone, this report uses
browser-local (matching `fmtTime` / `localDayKey`, which it has always used);
and rows here always have a Trip, so there is no ±15 min padded fallback window.

The parent also has an **Activity Report** that shares `buildLegacyByDevice`,
`buildLegacyDayBlocks`, `resolveStopAddresses`, `addressForPoint` and
`parseDurationToMins`. Anything put into those helpers changes both reports.
The overnight rule below is deliberately kept out of them: it lives in
`legacyStopMins`, which only the legacy call sites read.

---

## Data notes / known behaviour (inherited from the parent)
- **Distance is in KILOMETRES** (`Trip.Distance`). km->mi via `* 0.621371`.
  (The parent app originally treated it as metres, making distances ~1000x too
  small; corrected here from the start.)
- **Durations** (`idlingDuration`, `stopDuration`, `drivingDuration`) come as
  .NET TimeSpan strings ("hh:mm:ss" / "d.hh:mm:ss"); `parseDurationToMins`
  handles that plus ISO-8601 and numeric-seconds.
- **"(Ignition On)" row** shows no duration on purpose — the gap to the previous
  trip is overnight/parked time, not idle, and Trip data has no pre-drive idle.
- **Overnight parks are excluded** (v1.1.0). `Trip.StopDuration` runs to the next
  trip's start, so a van parked at 18:00 and driven at 07:00 arrived as a single
  13-hour stop on the day's last trip and inflated Total and Average Stop
  Duration. `buildLegacyStopInfo` now drops a stop when all three hold: it is
  that vehicle's last trip of the **local** day, it runs longer than
  `OVERNIGHT_MIN_HOURS` (3), and it ends on a later local day. A long mid-shift
  stop that ends the same day still counts. Excluded stops show an em-dash and
  are left out of the day total, Total Stop Duration, Average Stop Duration and
  Number of Stops — so average x count still equals the total. A trip with no
  `stopDuration` (usually the last in the range) is excluded for the same reason.
  Removal is silent by design: no separate "Parked" column.
- **Days are bucketed in local time** (v1.1.0), via `localDayKey`. Previously
  `fmtDateShort` (`toISOString`) grouped in UTC while every displayed time was
  local, so a 00:30 BST trip filed under the previous day. `fmtDateShort` itself
  is unchanged — the parent still uses it for `bucketKeyFor` and the trips list.
- **Simplification**: the reference format's "Working Total" vs after-hours
  "Total" split is collapsed into one daily total (Trip `Work*`/`AfterHours*`
  fields not yet used).
- **Idle sits inside stop, not beside it.** Geotab's `StopDuration` "also
  includes any idling done at the end of a trip", so "Total Stop Duration" and
  "Total Idle Time" must not be added together. Not yet reflected in the
  labelling — open cosmetic item.

---

## Deployment
- Intended to be hosted under the same GitHub Pages as Smart Insights, at the
  `legacy-trip-report/` subpath. Fully self-contained, so it can be lifted into
  its own repo with no code changes (only the config.json `url` would change).
- Cache-bust: asset URLs use `?v=1.3.0` — increment on each release, and keep the
  `version-badge` in index.html and `version` in config.json in step.
- Push is manual (same workflow as the parent app).
