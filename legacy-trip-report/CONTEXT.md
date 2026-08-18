# Legacy Trip History — Standalone Add-in — Session Context

## What this is
A **single-purpose** MyGeotab add-in that produces one thing: the Legacy Trip
History Report (a classic per-vehicle/per-day trip log with reverse-geocoded
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
2. Group trips by device, sorted chronologically; split into per-day blocks.
3. Reverse-geocode every unique `Trip.stopPoint` via MyGeotab's **`GetAddresses`**
   method (coordinates never leave Geotab; deduped to ~1m precision; batched up
   to 400/call under the 450/min limit). No third-party geocoder.
4. Render per-vehicle/per-day tables: "Starting from" lead-in, "(Ignition On)"
   marker, trip rows, daily total. KPIs across the whole range.
5. Export CSV (flat) or branded PDF (Geotab logo, navy header, KPI grid,
   per-day tables).

Length selector (Daily / Weekly / Monthly / Custom) auto-fills the from/to date
inputs; Custom leaves them for manual entry. Optional vehicle filter.

---

## Important shared-code note
This is a **lift** of the legacy-report functions from the parent
`../src/app.js`. If you fix a bug in one, mirror it in the other until they are
formally split into separate repos. Functions kept in sync include:
`parseDurationToMins`, `fmtDurWhole`, `fmtDurPrecise`, `milesFromDistance`,
`buildLegacyByDevice`, `resolveStopAddresses`, `formatReverseGeocodeAddress`,
`addressForPoint`, `buildLegacyDayBlocks`, `renderLegacyTripHistoryOutput`,
`exportLegacyTripHistoryPdf`, `drawPdfHeaderLogo`, `localDayKey`,
`buildLegacyStopInfo`, `legacyStopMins`, `OVERNIGHT_MIN_HOURS`.

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
- Cache-bust: asset URLs use `?v=1.1.0` — increment on each release, and keep the
  `version-badge` in index.html and `version` in config.json in step.
- Push is manual (same workflow as the parent app).
