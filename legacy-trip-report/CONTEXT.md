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
    filters.js      — Shared page-memory catalog, selectors and guarded runner;
                      also loaded by the parent app
    filters.css     — Scoped selectors and wrapped preview address cells
    zones.js        — Optional native zone containment, lookup lifecycle and columns
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
1. Load `Device`, `User` and `Group` catalogs for pre-run selectors; fetch `Trip`
   for the selected range when Run is clicked. Failed/incomplete catalogs are
   shown with Retry, never treated as an unrestricted filter.
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
   marker, trip rows and daily/driver/vehicle totals. No on-screen KPI tiles;
   the PDF retains its report-wide summary.
6. Export CSV (one trip per row, with that trip's singular `Driver`) or branded
   PDF (Geotab logo, navy header, KPI grid, vehicle/driver/day sections). PDF table
   headers repeat the vehicle, driver and date on continuation pages.

Length selector (Daily / Weekly / Monthly / Custom) auto-fills the from/to date
inputs; Custom leaves them for manual entry. Group selection includes descendants
and combines selected groups as a union. Assets uses `Device.groups`; Drivers
uses driver home `driverGroups` (not `companyGroups`, `reportGroups` or security
groups). Unknown drivers have no group membership, so appear only under All
groups. Driver trips can span multiple assets, independent of asset group.

## Current version: v1.6.1

**Why:** The portrait report still reserves 14 mm on both sides of the trip
tables, leaving less space than necessary for addresses and optional zone names.

**Approach:** Use 5 mm side margins for the Legacy trip tables only, expanding
their total width from 182 to 200 mm. Redistribute that space among the columns
while keeping portrait A4, text wrapping, all numerical results and links. The
navy header/logo, black totals bar, KPI boxes, per-vehicle/driver heading inset,
and page footers stay at their original coordinates. Top/bottom AutoTable margins
and pagination rules are unchanged; wider cells may naturally alter wrapping and
the resulting number of pages. A printer with a non-printable area larger than
5 mm can still clip the table; this change does not move the original edge-to-edge
navy header. The parent and standalone Legacy copies remain aligned. No changes
to Activity, CSV, or preview. GitHub publication requested on 2026-11-05.

Verification: all 228 report regression checks and four real jsPDF/AutoTable
rendering smokes pass, both with and without Zones in both apps. The rendered
fixture embeds the real logo; tests hold header/totals/KPI/footer positions and
table content at 5–205 mm, including wrapped addresses, zone names and links.
Baseline regression tests failed at the old side-margin/width assertions.
Independent plan r3 and local code review PASS (GPT 6 Sol); physical-printer
verification remains pending. GitHub publication requested by the user on
2026-11-05; deployment evidence follows the push.

### v1.6.0

**Why:** Customers using existing Geotab zones want the zone name beside stops
inside their boundaries; customers without zones should not have an empty column.

**Approach:** A labelled native checkbox styled as a slider/switch defaults off.
When enabled, a Zones column appears after Stop Location in preview, CSV and PDF.
`GetAddresses.zones` supplies containing zone IDs; a read-only `Get Zone` lookup
supplies existing current zone names. No zone creation or proximity guesses.
Overlapping zones are sorted by name and listed together, deduplicated by ID.

- Membership keys use exact stop longitude/latitude, not the rounded address
  cache: points either side of a boundary must remain distinct. Zone lookup uses
  currently active, accessible zone definitions, not historical zone geometry.
  Missing GPS and lookup/name failures remain explicit; an actual no-match is
  blank. The 10,000-zone response cap fails visibly rather than silently omitting
  records. There are no Zone requests when the option is off.
- One guarded lookup snapshot serves preview and both exports. The option stays
  in page memory across runs; new runs refresh memberships. Ready results can be
  reused when toggling off/on. While loading, enabled-zone exports are disabled;
  turning off immediately restores them. Cancellation clears queued timers,
  releases shared controls and exposes Retry instead of leaving Loading forever.
  Failed lookups preserve the trip report with labelled zone-error cells.
- The optional CSV has 11 columns; off retains the original 10-column schema.
  PDF layouts use 182mm total width inside portrait A4: 8 columns off, 9 on. On
  uses 30mm address and 26mm zone columns, with natural wrapping and growing rows.
  Totals, departure recovery and Trip History links remain consistent in both.
- Plan r2 independently PASS (GPT 6 Astra). All 228 regression tests and four
  real jsPDF/AutoTable smokes pass (parent/standalone, Zones off/on).
  An authenticated isolated browser tab ran the exact local source against 195
  real trips: 52 stops matched customer zones, including 11 overlapping matches.
  Preview/CSV/PDF matched native membership/name responses; totals were unchanged,
  all 195 links remained, and the Zones-on PDF stayed portrait (13 pages).
  The final corrected source was re-tested in the authenticated tab, including
  on/off CSV/PDF audits. The native keyboard switch toggled both ways with trusted
  Space events; background relay
  key dispatch did not activate it, so the keyboard test used the foreground tab.
  The temporary verification tab was closed; the original user tab was preserved.
- Initial code review found two Medium issues: cancellation could restore stale
  rows during a new run, and malformed individual address entries could appear
  as no zone. New regression assertions reproduced both before their fixes;
  the runner now clears the old preview and membership responses require objects.
  Final independent re-review PASS (GPT 6 Astra): both findings accepted fixed,
  all 228 tests/four PDF smokes independently rerun, plus cancellation and snapshot
  probes. Reviewed diff SHA-256:
  `4b20a4efb650fd138a31bc6125a4d6cddd0c5b79bb109296c4f8823abfd593c3`
  (before this documentation receipt; implementation unchanged). This is local
  review and live-source verification, not deployment or CI. GitHub publication
  explicitly requested by the user on 2026-11-05.

### v1.5.2

**Why:** Live inspection found 184 resolved stop rows but two unknown Starting
from rows: each vehicle's first fetched trip had no predecessor. Both departure
GPS positions and formatted addresses were available through the logged-in API.

**Approach:** When the full vehicle timeline has no valid, non-overlapping
previous stop location, query `LogRecord` for that device at `Trip.start` using a
1ms positive boundary window. Accept only exact-time, matching-device finite
coordinates; interpolate through the API, never guess a nearby location. Send
recovered departure coordinates through the existing deduplicated address batch.

- Shared helpers in `filters.js` keep one in-memory start snapshot keyed by
  asset ID and start instant, published alongside completed trips/addresses/range.
  The same predecessor predicate serves lookup and rendering. HTML/PDF Starting
  from rows use it; CSV retains its existing stop-only schema. No raw Trip
  mutation, synthetic predecessor, new trip, totals or link changes.
- GPS requests are paced at 100ms, geocoding at 150ms. Once-only guards handle
  errors, timeout, throws and duplicate/late callbacks; cancellation checks also
  cover queued dispatch. Missing GPS is distinct from a failed lookup. Failed
  departure geocoding retains labelled coordinates; ordinary stop fallback stays
  unchanged. New runs/initialization clear the prior start snapshot.
- Regression: the first-departure test failed on both previous implementations
  at the missing resolved Starting from assertion before the fix. Current tests
  include exact/same-time cross-asset identity, interpolation, invalid GPS,
  handovers, overlap, address deduplication, filtering, rerun, failures and cancel.
- The user requested implementation and GitHub publication on 2026-11-04.
  Independent plan r1 PASS (GPT 6 Astra). 183 local regression checks and both
  real PDF rendering smokes pass. Authenticated browser comparison ran the exact
  new standalone source and the previous source in detached DOMs against one
  freshly fetched Trip response: 187 trips, 13 Starting from rows, two unknown
  departures before and zero after (HTML and actual 12-page PDF). Two GPS reads
  recovered both departures; stop labels, canonical totals and selections were
  identical. The previously displayed report had 184 trips, so sharing one fresh
  response avoided comparing different ongoing-trip snapshots. No real GPS or
  address data was committed. Final reviewed-head/deployment verification is
  recorded separately when complete.
- Independent local code review PASS (GPT 6 Astra), including all 183 tests,
  both PDF smokes and 16 additional asynchronous/pacing/cancellation probes.
  Reviewed diff SHA-256: `e2848ba1ba53cc9de7a07b3a7e76569b0f5675fd572007827da004bbb80bd329`
  before adding this documentation receipt; implementation unchanged.

### v1.5.1

**Why:** System-created groups cluttered the primary list, dropdowns stayed open
after outside clicks, and the inline hint made the new controls misalign. The
customer also requested removing the six on-screen summary tiles.

**Approach:** Present customer-created groups first and collapse system branches
behind a Show/Hide button. Classification uses documented SDK KnownIds and their
descendants, never localized names. Company is a system option itself, not a
system subtree: its customer-created descendants remain in the main list, and
Company retains its distinct filtering semantics from unrestricted All groups.

- Outside clicks close either picker without changing selections, suppressing
  clicks, or interfering with native checkboxes. One document listener per
  wrapper survives reinitialization; Escape still returns focus to the summary.
- Hints/Retry stay inside the group panel. A visible Groups unavailable summary
  plus a visually-hidden live status outside collapsed panels expose failures
  without adding a toolbar row. Existing controls and summaries align at 34px;
  narrow screens still wrap naturally.
- Legacy preview generates no summary-card tiles. Table footers and PDF summary
  totals remain; parent non-Legacy report tiles are untouched.
- Verification: 159 local checks plus both real PDF rendering smokes. Old-source
  probes reproduce the outside-click, hidden failure and summary-tile assertions.
  Synthetic event tests are not native browser/keyboard or visual alignment
  proof: Chrome relay is disconnected, so those live checks remain pending.
- Plan r2 independently adopted; final local code review PASS by an independent
  GPT 6.1 Sol session, with 159 tests and both PDF smokes independently rerun.
  Reviewed diff SHA-256: `dc84927a64de8d23586b163e9452e8d1ddcbefdf8eeb7ad520a553a59a989b7c`
  (before adding this receipt; implementation unchanged). This is not native
  browser verification or a committed-head/CI receipt. GitHub publication was
  explicitly requested by the user on 2026-11-04; live visual verification is
  still pending.

### v1.5.0

**Why:** Landscape PDFs are unwanted, and the old post-run single-vehicle
selector reset whenever the report was rerun.

**Approach:** Keep all eight PDF columns within portrait A4 using fixed widths
(address 33mm), line wrapping and automatic row height. A shared controller
offers group filtering, an Assets/Drivers mode (Assets default), searchable
checkbox selection and Select all filtered. Individual exclusions remain set
on subsequent runs. Choices are not written to storage: refresh/close or a new
initialization resets them; changing report type and returning does not.

- Build the complete vehicle timeline first, then select driver groups without
  rebuilding predecessors or overnight-stop classification. Preview, totals,
  CSV and PDF all consume the same selected snapshot. Empty/failed/loading runs
  cannot export an earlier report. Completed exports retain their own range.
- Initialization clears catalogs/results/selection before replacing the API;
  callbacks from obsolete runs/identities cannot publish or dispatch more work.
  The same generation guard protects regional settings and initialization
  continuations. Identity-dependent preview/summary/filter content is cleared
  immediately. Cancellation restores the shared Run button for every report type.
- Each Run refreshes Device/User/Group membership records while retaining
  selected IDs/exclusions, so new assets and group-tree changes are reflected.
  Long asset headings wrap inside the portrait PDF margins as well as addresses.
  Parent initialization also loads pre-run catalogs when Legacy remains selected
  across reinitialization; it does not depend on a report-type change event.
- Tests: `node --test tests/trip-report-drivers.test.cjs`. Real PDF renderer:
  `node tests/legacy-pdf-layout.cjs <pinned-library-directory>` using the exact
  jsPDF 2.5.1 / AutoTable 3.8.2 scripts from index.html. This checks portrait
  bounds, address wrapping/height, oversize pagination, URI links and totals.
- Verification boundary: synthetic API tests are not authenticated API proof.
  Live MyGeotab/native driver-group matching and rendered keyboard/UI checks
  remain pending; the Chrome relay was disconnected. GitHub publication requested
  by the user on 2026-11-04; this is not authenticated MyGeotab verification.
- Regression evidence: 146 local tests pass. The real renderer smoke passes for
  both copies (four portrait pages, 40 links) and fails against base `809c185`
  at `Export must be portrait`. Executing the old vehicle-filter/rerun sequence
  resets one selected trip to all six; the new runner keeps the one selected trip.
- Review: large/risky plan r2 PASS by an independent GPT 6.1 Sol session. Local
  implementation review initially found four Medium issues (Run cancellation,
  identity DOM reset, stale unit callbacks, cached membership) and one Low
  (long PDF asset heading). All five were reproduced by failing regression
  assertions and fixed, then independently accepted. Re-review found one further
  Medium initialization gap (selected parent Legacy catalogs were empty without
  a change event); it was likewise reproduced and fixed. Final local review PASS
  by independent GPT 6.1 Sol (`geotab-azure/gpt-6.1-sol`), with all findings
  accepted fixed and 146 tests plus both PDF smokes independently rerun.
  Reviewed snapshot SHA-256:
  `6206a567668d442c641fad0ae8414a199742e01f96d0b470c9089e851edcbdd5`
  (before adding this receipt; implementation unchanged). This is a local-only
  review, not a committed-head/MR/CI or release-readiness receipt.

### v1.4.3

- Dates in the preview, CSV, PDF headings/range/created date and PDF filename use
  DD-MM-YYYY. `fmtLegacyDate` is presentation-only; canonical YYYY-MM-DD day keys,
  date input values, sorting, API ranges and Trip History timestamps stay intact.

### v1.4.2

- PDF now has eight columns: Trip History is the final column, with trip links
  on detail rows and plain "X stops" counts on daily, driver and vehicle totals.
  No separate Stops column; landscape A4, numerical values and preview unchanged.

### v1.4.1

- PDF uses landscape A4 and nine columns, with a separate Trip History column
  after Stops. Each trip has a real PDF URI annotation pointing to the same
  vehicle/day/trip deep link as the preview. Missing-host, totals, synthetic and
  blank continuation cells do not create annotations. Numerical fields unchanged.

- The viewer's `User.isMetric` is resolved during initialization before the report
  becomes available, matching Smart Insights. Canonical `Trip.distance` stays in
  kilometres: metric displays km, imperial converts numerically to miles in the
  preview, summary cards, filtered results, CSV and PDF (50 km = 31.07 mi).
- Preview and PDF have separate Distance and Duration (driving time) columns.
  Footers show each driver/day, each driver over the selected period within that
  vehicle, and the entire vehicle period. Totals sum unrounded source values and
  retain seconds. CSV keeps its flat one-trip-per-row format.

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
`parseDurationToMins`, `fmtDurWhole`, `fmtDurPrecise`, `distVal`, `distUnit`, `fmtDist`, `fmtLegacyDate`,
`legacyTotals`, `legacyTotalCells`, `legacyHtmlTotal`,
`buildLegacyByDevice`, `resolveStopAddresses`, `formatReverseGeocodeAddress`,
`addressForPoint`, `buildLegacyDayBlocks`, `renderLegacyTripHistoryOutput`,
`exportLegacyTripHistoryPdf`, `drawPdfHeaderLogo`, `localDayKey`,
`buildLegacyStopInfo`, `legacyStopMins`, `OVERNIGHT_MIN_HOURS`,
`resolveHost`, `cardDatePart`, `tripHistoryUrl`, `tripHistoryCell`, `esc`.
The historical driver helpers (`reportDriverId`, `isUnassignedDriver`,
`reportUserName`, `reportDriverInfo`, `resolveReportDrivers`,
`groupReportDriverDays`) are also copied into the standalone Activity Report.
Keep those three copies identical; the regression suite checks this.

The Legacy runner/filter controller is now shared, not copied: both pages load
`legacy-trip-report/src/filters.js` and `filters.css`. Do not reintroduce a
per-run selector rebuild or rebuild the timeline from driver-filtered trips.

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
- Cache-bust: asset URLs use `?v=1.6.1` — increment on each release, and keep the
  `version-badge` in index.html and `version` in config.json in step.
- Publish only on explicit user request (same workflow as the parent app).
