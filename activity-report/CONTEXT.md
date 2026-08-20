# Activity Report — standalone add-in

A self-contained MyGeotab add-in containing only the Activity Report. It is a lift
of the Activity Report functions from the parent Smart Insights `src/app.js`, so it
can be installed on its own by customers who don't want the full dashboard.

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
- **Two renames** for readability in a single-report file. Upstream these are
  shared with the Legacy Trip History report and carry "Legacy" in the name:
  `buildLegacyByDevice` → `buildByDevice`, `buildLegacyDayBlocks` → `buildDayBlocks`.
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

## Keep in sync

Every fix to the Activity Report belongs in **both** this file's `src/app.js` and
the parent `../src/app.js`, until the two are formally split into separate repos.

## Version

v1.0.0 — initial extraction. Bump `config.json`, the `version-badge` in
`index.html`, and the `?v=` cache-busting query on all three asset tags together.
