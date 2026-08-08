# AI / Technical Reference — Excel Timeplan HTML Dashboard

Complete reference for anyone (human or AI) maintaining, extending or debugging this project.
`README.md` is the short user-facing version.

**Read §0 first.** As of 2026-08-07 the dashboard reads **many workbooks (one per timeplan)** instead
of one master workbook with five input sheets. §0 is authoritative wherever it contradicts a later
section; the later sections are still correct about *the shape of one input sheet* and about the zip /
XML / rendering machinery, which did not change.

---

## 0.0 Distribution folder layout (v2.2)

The shipped folder is numbered so it sorts in workflow order. This is the layout every path in
this document and in `README.md` refers to:

```
README.md                      user guide / GitHub landing page, at the pack root
01_Documentation/              README-AI.md (this file)
docs/                          README screenshots
02_Templates/                  L1_TEMPLATE.xlsx … L5_TEMPLATE.xlsx — blank two-sheet starters
03_Time_Plans/                 the live plans, one .xlsx per timeplan — THIS is the folder users load
03_Time_Plans/_baselines/      YYYY-MM_n.json snapshots (see §0.8)
03_Time_Plans/04_Archived_Timeplans/  pre-overwrite copies, <date>_<time>_<file>.xlsx (see §0.10)
04_Exports/                    suggested destination for SVG / PNG / PDF
05_Master_Time_Plan_Dashboard/ Time Plan Dashboard.html — the single-file app
```

Renamed at 2026-08-07 from `Documentation/`, `Workbook/Templates/`, `Timeplans/`, `Saved Plans/`,
`Exports/` and a root-level HTML. **Nothing in the code hardcodes any of these names.** The dashboard
only ever knows the folders the user picked; the archive folder is a *picked handle* (`archiveDirHandle`,
persisted in IDB under `archiveDir`) and falls back to `ARCHIVE_DIR` created inside the timeplan folder;
`_baselines` is created relative to that folder (`BASELINE_DIR`), so the numbered names are convention and can be changed
again without touching the app. The legacy seven-sheet master workbook was dropped from the
distribution (it still lives at the project root as `Time_Plan_L1-L5_Master.xlsx` and still loads —
see §0.3 case 3).

---

## 0. Multi-file architecture (current)

### 0.1 Model

`state.levels` (a `{L1…L5}` object) is gone. The dashboard now holds `state.plans` — an **ordered
array**, one entry per loaded `.xlsx`:

```js
plan = {
  id, order,            // 'l2_component', index after sorting
  fileName,             // 'L2_Component.xlsx'
  key,                  // 'L2_Component' — the dependency prefix + display label
  sortName,             // 'Component'
  level, levelKey,      // 2, 'L2'  (unknownLevel plans get level 9 / 'L?')
  sheetName, sheetPath, // 'L2 input', 'xl/worksheets/sheet3.xml'
  color,                // level colour, lightened per duplicate at the same level
  meta, activities, layout,
}
```

Original file bytes live **outside** state in `this.planBytes[planId]` (ArrayBuffer), the way
`this.originalAB` used to.

**Ordering** (`finalizePlans`): `level` asc → `sortName` alphabetical (locale-aware, case-insensitive)
→ `fileName`. Duplicate `key`s get ` (2)` appended; duplicate `id`s get `_2`.

**Colours**: `levelColor(level, dupIndex)` — `LEVEL_COLORS[level-1]`, mixed toward white by
`dupIndex*17%` (capped 50%) so the 2nd and 3rd plan at the same level are tellable apart.

### 0.2 Activity codes

The old `L1-3` code string is replaced by `plan.id + '|' + no` (e.g. `l2_component|3`), assigned in
`finalizePlans` and stored on `a.code`. Blank `No.` becomes `r{row}`; a duplicate within one file gets
`#2` appended. **Nothing parses a code string any more** — `computeActivityIndex` returns
`{planId, planKey, no, display, rank, …}` per code and `rank = plan.order*100000 + numeric(no)` drives
arrow direction. `display` is `plan.key + '-' + no`, which is what the UI shows and what the user
types in Excel.

`state.edits`, `state.activityVisibility`, `selectedCode` and `hoveredArrowKey` are all keyed by code.
`state.visible` and `state.planNameVisible` / `planNameOverrides` are keyed by `plan.id`.

### 0.3 Which sheet a file contributes

**Current shipped shape: two sheets per file** — `L{n} input` (data, the only sheet the dashboard
touches) and `L{n} time plan` (Excel-side Gantt, formulas pointing at `'L{n} input'`). The demo files
and `02_Templates/L{1-5}_TEMPLATE.xlsx` are built this way. Both are produced by rebuilding the
master into a fresh 10-part package: `[Content_Types].xml`, `_rels/.rels`, `docProps/{app,core}.xml`,
`xl/workbook.xml` (2 `<sheet>` entries, `calcPr fullCalcOnLoad="1"`), `xl/_rels/workbook.xml.rels`,
`xl/theme/theme1.xml`, `xl/styles.xml`, `xl/worksheets/sheet1.xml` (input), `xl/worksheets/sheet2.xml`
(time plan). The time plan sheet is the master's `L3 time plan` with `'L3 input'` string-replaced to
`'L{n} input'` — it references no other sheet, so that is a complete repoint. Dropping the four other
input sheets and the 5.4 MB `master time plan` sheet takes each file from ~884 KB to ~110 KB.

`parseWorkbookFile` (per file):

1. `parsePlanFileName` reads a leading `L{1-9}` off the base name → `{base, level, name}`.
2. **If the name declares a level**, the file contributes exactly **one** plan: the sheet named
   `L{n} input`, else the only `L{n} input`-shaped sheet, else the first sheet matching `/input/i`,
   else sheet 1.
   *This is deliberate and load-bearing:* it makes "copy the old master five times and rename the
   copies" a valid migration — each copy still contains all five input sheets, but only its own level
   is read, so nothing duplicates.
3. **If the name declares no level**, every `L{n} input` sheet in the file becomes its own plan
   (`key = 'L{n}_' + base`) — this is how the untouched master workbook still loads as five plans.
4. No level and no `L{n} input` sheets → one plan at level 9, `unknownLevel: true`, warned about.

A file that throws is skipped with a note; the load only fails if *no* file parsed.

### 0.4 Dependency references

Canonical user syntax is **`{file base name}{sep}{activity name or No.}`**, e.g.
`L1_Project-Gate 2 review`. `resolveDependencies(plans)` tries, in order:

1. **Plan prefix** — every plan `key`, **longest first**, tested as a prefix of the normalised ref;
   the remainder must start with a separator (`- – — ! : / | >`). Longest-first is what stops
   `L2_Assembly` from shadowing `L2_Assembly Line`.
2. **Level shorthand** `L{n}{sep}{token}` — every plan at that level, in display order, first hit
   wins. This is what keeps pre-existing `L1-Risk assessment` data working.
3. **Bare token** — looked up in the activity's own plan.

Normalisation is `depNameKey` (collapse whitespace, lowercase); a trailing `.xlsx` in the ref is
stripped. Numeric tokens match `No.` first, then names. Unresolved refs go to `a.depsUnresolved`, the
load banner, and a red chip. The `resolveDependencies`-after-every-reparse trap from §3.5 still
applies.

### 0.5 Persistence

localStorage cannot hold several workbooks, so the cache moved to **IndexedDB** (`tpDashboard` /
`files` store): key `planFiles` = `[{name, ab}]` (+ a `__source` label), key `outDir` = the saved
output **directory handle** (structured-cloneable; permission is re-queried on save). `Clear` deletes
`planFiles` and also removes the two dead v1 localStorage keys.

### 0.6 Save

`onSaveExcel` groups pending edits by `planId`, unions that with the plans whose Timeplan Name box is
dirty, and for each **touched** plan: rebuilds that plan's zip from `this.planBytes[id]`, re-parses,
re-keys the reparsed rows onto the existing codes by `row`, verifies every edited value, and only then
queues the output. Nothing is written until every file has verified.

Output goes to the stored `saveDirHandle` (`getFileHandle(name,{create:true})` → `createWritable`),
asking for one if none is set; anything other than an `AbortError` from the picker degrades to one
`downloadBlob` per file, 350 ms apart. File names come from `outputFileName(plan)`:
`{key}_edited_{YYYY-MM-DD}.xlsx` by default, `plan.fileName` when the `saveNaming` prop is
`sameName`. After writing, `planBytes` is refreshed, `resolveDependencies` re-run, and the IndexedDB
cache rewritten.

### 0.7 Superseded by the above

§1's file table (below), §3.3's `levels.L1` reads, §3.5's `L{n}-{No.}` codes, §4.1's
`computeDefaultRange(levels)` signature, all of §5, §6's localStorage keys. `LEVEL_META` no longer
exists (`LEVEL_COLORS` + `levelColor()` replace it). `computeDefaultRange`, `computeAutoColumnWidths`,
`computeActivityIndex`, `detectCycles` and `resolveDependencies` all take the plans **array**.

---

## 0.8 UI revision v2.2 (2026-08-07)

A usability pass. Nothing about parsing, write-back or geometry changed; all of it is in the
template and in `renderVals`.

### "Checks" is now "Warnings"

User-visible strings only. `healthBtnLabel` is `'Warnings '+n` / `'No warnings'`; `healthTitle` is
`'Warnings'` (or `'Warnings — '+key` in single-file check mode); the load-banner notes, the per-plan
badge tooltip and the two `data-props` sections (`strictFileNames`, `staleDays`) all say Warnings.
The internal names are unchanged — `buildHealth`, `health`, `healthRows`, `panel==='health'`,
`healthByPlan`. **Do not rename them** unless you rename them everywhere; the panel key `'health'` is
also what `openPanel` stores in state.

The panel now also carries a permanent **How to write a dependency** block below the issue list
(syntax, the `L1_Programme-14; L3_Tooling-7` example, and the "use the No." rule). It is static
markup, not data — keep it in sync with `resolveDependencies` (§0.4) if the accepted syntax changes.

### Save is a two-step

The red **Save to Excel (n)** button no longer writes. It calls `onOpenSavePanel` → `openPanel('save')`
and the new `panelSave` popup lists, per changed timeplan: the plan key, a breakdown
(`3 activities + timeplan name`), the total, the exact output path, and one line per edited activity
showing `Ref · name · what changed` (`start → Wk 12 2027`, `status → Risk`, …). The write button
(`onSaveExcel`) and a **Discard changes** button (`onDiscardEdits`, confirms, clears `edits` and
`planNameOverrides`) sit in its footer. `onSaveExcel` sets `panel:null` when it starts.

`saveRows` is built in `renderVals` by grouping `state.edits` by `planId`, diffing each edit against
the activity's parsed values to name the fields that actually changed, then unioning with
`dirtyPlanNamePlans()`. It is display only — the real grouping still happens inside `onSaveExcel`
(§0.6), so **if you change how edits are grouped for writing, change `saveRows` too** or the panel
will promise a different set of files from the one that gets written.

### Baselines: every capture is kept

Ids moved from `YYYY-MM` (one per month, replace-with-confirm) to `YYYY-MM_n` — `nextBaselineId()`
scans the loaded baselines for the current month and returns the next free sequence number, so
capturing never overwrites. `monthLabel` renders both forms (`August 2026`, `August 2026 · 2`) and
`baselineSortKey` zero-pads the sequence so `_10` sorts after `_9` (plain `localeCompare` did not).
`baselineConfirm` state is now dead but still initialised; `currentBaselineId` was replaced by
`currentBaselineMonth` + `nextBaselineId`. Legacy `YYYY-MM.json` files still load and still label.

Capturing sets `showBaseline:true` so the comparison is visible immediately. The **Show** checkbox
became a pill toggle (`baseToggleBtnStyle` / `baseKnobTrackStyle` / `baseKnobStyle` / `baseToggleLabel`,
modelled on the Compact toggle) reading *Baseline view on* / *Baseline view off*. The panel leads with
two paragraphs explaining what capture does and does not do — written because the button's purpose
was genuinely unreadable from `+3w` alone.

### Timeplans panel has column headings

A sticky header row (Show · file | Timeplan name | Gutter | Age | Acts) over the list and a legend
below it. The header widths are hand-matched to the row widths (`170px` / flex / `47px` / `28px` /
`26px`) — **if you change a row column width, change the header's too.** The legend interpolates
`staleDaysLabel` and `staleDaysDoubleLabel` from the `staleDays` prop.

### Ref column on plan tabs

Each plan tab's table gained a monospace **Ref** column (`detailPlan.key + '-' + act.no`,
`user-select:all`) so the exact dependency string can be copied instead of typed. This is the main
"foolproofing" of the dependency syntax — the grammar itself did not change.

`APP_VERSION` is `v2.2`.

---

## 0.9 UI revision v2.7 (2026-08-08)

`APP_VERSION` is now the string `'Version 2.8'` (rendered verbatim in the top bar pill, so it must
stay human-readable — do not put a bare `v` back).

### Programme heading is dashboard-owned, not file-owned

The strip above the activity table used to read `projectName`/`code`/`sop` out of whichever plan
happened to have them. With one file per timeplan that is arbitrary. It is now an override stored in
`localStorage` under `tpDashboardHeadingV1`:

- `HDR_FIELDS` declares the six fields (title, description, code, sop, revision, issuedBy) with their
  labels, placeholders and flex widths. Add a field here and the edit form grows automatically.
- `loadHeading()` / `saveHeading()` wrap the storage; `this.state.hdr` holds it.
- `effHeading()` merges override over file fallback. **An override of `''` is meaningful** (the user
  cleared the field), so only `undefined` falls through to the file — do not switch this to `||`.
- `onHdrField` writes through to storage on every keystroke; `onHdrReset` clears the whole override.
- `infoBarHeight` is 72 normally, 152 while editing. `hdrView`/`hdrEdit` swap the two layouts inside
  the same sticky strip.

The SVG export header (`buildExportSVG`) reads `effHeading()` too and now draws three lines
(title / description + plan count + baseline / `factLine`) in a 54px band with the accent stripe,
matching the on-screen strip. `infoH` went 36 → 54; if you change the strip, change both.
`baseName()` (export filenames) also comes from `effHeading().title`.

### Top bar

"Time Plan Dashboard" → **"Master Time Plan Dashboard"**. The plain text `headerContext`
("12 timeplans · 60 activities") is replaced by two pill stats with small inline SVG marks
(`statPlans`/`statActs`), shown only when a file is loaded; `headerContext` still renders in the
not-loaded state. The version pill got a green dot.

### Warnings filters

`state.healthSev` (`all`/`errors`/`warnings`) and `state.healthPlan` (`all`/`__cross`/planId) filter
`health` into `healthFiltered` **before** grouping. `sevTabs` and `healthPlanOptions` drive the filter
bar. The counts on the tabs and on the toolbar button always come from the **unfiltered** `health` —
filtering must never change the reported totals.

### Baseline: compare two captures

`state.baselineTo` is `'today'` or a second baseline id. `compareDates(plan, act)` returns the
"to" side — live effective dates, or the entry from that second capture — and every consumer
(`baselineSlip` on screen, in the export, and in the summary counts) goes through it. When comparing
two captures the **bars still show today's plan**; only the ghost outline and the badge describe the
snapshot-to-snapshot move. That is deliberate (the plan you are looking at is still the real one) and
the panel says so in an amber note.

### Knock-on / impact analysis

New toolbar button and panel. Two pure functions near `slipText`:

- `buildSuccessorMap(plans)` inverts `dependsOn` into predecessor → successors.
- `computeImpact(plans, changes, effDates)` BFS-walks successors from each change, depth-capped at
  `IMPACT_MAX_DEPTH` (8) with a per-change `localSeen` set so a diamond dependency is reported once.
  For each successor it computes `slack = successorStart - predecessorEnd`; `slack <= 0` is
  `must` (shift `1 - slack` weeks), `slack <= 2` is `review`, otherwise `ok`. The queue carries the
  **shifted** end so knock-on accumulates down the chain.

`impactChanges()` defines what counts as "a change": an unsaved edit that moved an end date, or an
end date that differs from the selected baseline. Edits win over baseline for the same activity.

Reports are plain print-styled HTML opened in a new window (`reportShell` + `reportHeader`, A4
portrait, `@page` margin 14mm). `onPrintImpactReport` builds then `window.open`s; `onDownloadImpactReport`
saves through `saveBlobWithPicker`. **`buildImpactReport()` runs before `window.open`**, so a blocked
pop-up cannot hide a build error.

### Status sheets

`buildStatusSheets(planIds)` emits one `.page` per plan, each with its own `reportHeader()` (page 1
only would be wrong for a document that gets split up in a review) — hence `reportShell(..., {bare:true})`.
Per page: level/activity-count/freshness/baseline/warning chips, the file + sheet + Plan ID +
dependency-prefix line, status mix, the full activity table (with `Ref` and raw `dependsOnRaw`), that
plan's warnings, and a sign-off line. Scope `'current'` uses `state.page` when it is not `'master'`.

### Capacity

The activity scan loop no longer stops early on row count: it reads to row 2000 and gives up only
after 80 consecutive blank rows (was: 30 blanks after row 107, which would truncate a 200-row plan
with spacer rows). The shipped templates are formatted/formula-filled to about row 206–216, so
`buildHealth` warns above 200 activities telling the owner to copy the last row of the *time plan*
sheet down. Nothing in the dashboard caps at 200.

---

## 1. Files

| File | Role |
|---|---|
| `dist/03_Time_Plans/L*.xlsx` | The 12 demo timeplans (1×L1, 2×L2, 3×L3, 4×L4, 2×L5), each two sheets. |
| `Time_Plan_L1-L5_Master.xlsx` | The original single workbook. Kept as the template to split from; still loads (see §0.3 case 3). |
| `Time Plan Dashboard.dc.html` | **Source of truth for the UI.** A Design Component: template + `class Component extends DCLogic`. Edit this. |
| `Time Plan Dashboard (Standalone).html` | Self-contained bundle of the DC — this is what ships (copied to `dist/06_Master_Time_Plan_Dashboard/Time Plan Dashboard.html`). Regenerate after every DC change; never hand-edit. |
| `README.md` / `README-AI.md` | User guide / this file. |
| `support.js` | DC runtime. Generated. Never edit. |

The distribution folder (`dist/`) is laid out per §0.0.

No server, no build step, no npm, no XLSX library. The dashboard opens from `file://` and does
everything in the browser.

**There is no `read me` sheet in the workbook.** An earlier revision had one as sheet 1; it was
deleted (the local part `xl/worksheets/sheet1.xml` was dropped and its entries removed from
`xl/workbook.xml`, `xl/_rels/workbook.xml.rels` and `[Content_Types].xml`). Do not reintroduce it —
the usage instructions live in `README.md`. Note the remaining worksheet parts are still named
`sheet2.xml`…`sheet8.xml`; they were deliberately **not** renumbered, because part filenames are
resolved through the rels graph, not by number.

**No text wrapping anywhere in the workbook.** Every `wrapText` attribute was stripped from
`xl/styles.xml`, and the literal newlines that used to sit inside header labels (`"Start\nWk"`) were
flattened to spaces (`"Start Wk"`). If you regenerate the workbook from a script, do not re-emit
`wrapText="1"` or embedded `\n` in any cell value.

---

## 2. Workbook construction

### 2.1 Sheet inventory

| Tab | Part | Purpose | Protected |
|---|---|---|---|
| `L1 input` … `L5 input` | `sheet2.xml`…`sheet6.xml` | Data entry, one sheet per hierarchy level | no |
| `L3 time plan` | `sheet7.xml` | Native Excel Gantt of L3 only | yes (`password="DC2F"`) |
| `master time plan` | `sheet8.xml` | Native Excel Gantt of all five levels stacked | yes (`password="DC2F"`) |

The two time-plan sheets are **derived views**. They contain no typed data — every visible value is a
formula pointing at an input sheet, and every bar is conditional formatting. The dashboard does not
read them at all; they exist for people who never open the HTML.

The tab names are load-bearing: the dashboard locates sheets by name through
`xl/workbook.xml` → `xl/_rels/workbook.xml.rels`. Sheet **order** may change; the names
`L1 input`…`L5 input` may not.

### 2.2 `L{n} input` layout

Header block, rows 2–5:

| Cell | Content |
|---|---|
| `E2` | Project Name (merged `E2:N2`) |
| `E3` | Project Code |
| `F3` / `G3` | "SoP" label / SoP value, e.g. `Wk 1, 2026` |
| `E4` | Description (merged `E4:N4`) |
| `E5` | **Timeplan Name** (merged `E5:N5`) — this level's plan name, e.g. "L1 Programme Master Plan" |
| `B2:D2`, `B3:D3`, `B4:D4`, `B5:D5` | Row labels (merged); `B5` reads "Timeplan Name" |

> The Timeplan Name row was added late. An earlier revision had the dashboard reading `E5` while the
> workbook left row 5 completely blank and unlabelled — so the five toolbar boxes always loaded empty
> and nothing typed into them ever reached Excel. The label in `B5` is load-bearing: the dashboard
> locates the value cell by finding it (see §3.5), so do not delete or reword it.

Timeline-window block (present on every input sheet, but **only `L1 input`'s is used**):

| Cell | Content |
|---|---|
| `P2:S2` | Labels: Start Wk, Start Yr, End Wk, End Yr |
| `P3` | Timeline start week (default `1`) |
| `Q3` | Timeline start year (default `2026`) |
| `R3` | Timeline end week (default `52`) |
| `S3` | Timeline end year (default `2030`) |
| `U2:V2` | Labels: Current Wk, Current Yr |
| `U3` | `=WEEKNUM(TODAY(),21)` — ISO week number of today |
| `V3` | `=YEAR(TODAY())` |
| `P4` | "Timeplan range" caption (merged `P4:S4`) |
| `U4` | "Today" caption (merged `U4:V4`) |

Activity table: **header row 6, data rows 7–106** (100 pre-formatted rows).

| Col | Field | Notes |
|---|---|---|
| C | Colour | Cell **fill** = the activity's colour everywhere in the dashboard. No fill → palette fallback by row index. The cell holds no value. |
| D | No. | Short ID. Canonical dependency code is `L{level}-{No.}`. |
| E | Activity / Event | Name. An empty E cell ends the row for every downstream formula and for the parser. |
| F | Responsible | Free text. |
| G | Location | Free text. |
| H / I | Start Wk / Start Yr | ISO week + year. **Primary date entry.** Validation: whole 1–53 / whole 1000–9999. |
| J / K | End Wk / End Yr | ISO week + year. Same validation. |
| L | Tracker | Formula, see below. Not read by the dashboard. |
| M | Status | Data-validation dropdown: `TBD, Planned, Aligned, Dependency, Risk, Blocker, Completed`. |
| N | Comments / Notes | Free text, not rendered anywhere. |
| O | Dependencies | `L{level}-{activity name}`, several separated by `;` or `,`. |

There is **no Start Date / End Date column**. An earlier revision added real date serials at L/M and
shifted Status/Comments/Dependencies to O/P/Q; that was reverted. If you find a workbook shaped that
way it is stale — do not reintroduce it.

### 2.3 The `Tracker` formula (column L)

Row 7, filled down to row 106:

```
=IF(E7="","",
   IF(M7="Completed","On Track",
     IF(OR(J7="",K7=""),"",
       IF(($V$3*100+$U$3)>(K7*100+J7),"Delayed","On Track"))))
```

Reading it: blank row → blank. Explicitly completed → "On Track". No end date → blank. Otherwise
compare **today's encoded week** against the activity's **encoded end week**, where the encoding is
`year*100 + week`. That single integer encoding is used everywhere in this project (Excel and JS
alike) to make week/year pairs orderable with one comparison.

Two conditional-format rules on `L7:L106` colour "On Track" green and "Delayed" red.

### 2.4 Status colouring

`M7:M106` carries seven `containsText` conditional-format rules, one per status value, each pointing
at a `dxf` in `xl/styles.xml`. The same seven dxfs are reused on the time-plan sheets' status columns
(`'L3 time plan'!E7:E106`, `'master time plan'!F7:F506`), so status colours stay consistent.

The dashboard hardcodes an equivalent palette in `STATUS_STYLES` — if you change the Excel dxfs,
change `STATUS_STYLES` too, or the two views will disagree.

### 2.5 `L3 time plan` (sheet7)

Frozen pane at `F7` (`xSplit=5`, `ySplit=6`). Columns A–E are the row header, F onwards is the week
grid (to column `LR`).

Title block:

```
A1 = IF('L3 input'!E2<>"",'L3 input'!E2,"")
A2 = "Project Code: "&'L3 input'!E3&"      |      SoP Date: "&IF('L3 input'!G3<>"",'L3 input'!G3,"—")
A3 = IF('L3 input'!E4<>"",'L3 input'!E4,"")
```

`master time plan!A1:A3` carry the same three formulas against `'L1 input'`.

Row-header labels are literal text in `A4:E6` (merged vertically): No. / Activity / Event /
Responsible / Location / Status.

Row header body, row 7 filled down — every cell is guarded on the activity name being non-empty:

```
A7 = IF('L3 input'!E7<>"",'L3 input'!D7,"")     ' No.
B7 = IF('L3 input'!E7<>"",'L3 input'!E7,"")     ' Activity
C7 = IF('L3 input'!E7<>"",'L3 input'!F7,"")     ' Responsible
D7 = IF('L3 input'!E7<>"",'L3 input'!G7,"")     ' Location
E7 = IF('L3 input'!E7<>"",'L3 input'!M7,"")     ' Status
```

**The week grid header — this is the clever part.** Three rows: year (4), month (5), week (6). The
first grid column seeds from the L1-style range cells on `L3 input`, then every subsequent column
increments itself:

```
F4 = 'L3 input'!$Q$3                            ' seed year
F6 = 'L3 input'!$P$3                            ' seed week

G6 = IF(F6 = IF(WEEKNUM(DATE(F4,12,28),21)=53, 53, 52), 1, F6+1)
G4 = IF(F6 = IF(WEEKNUM(DATE(F4,12,28),21)=53, 53, 52), F4+1, F4)

F5 = TEXT(DATE(F4,1,1)+(F6-1)*7,"MMM")          ' month label
```

`WEEKNUM(DATE(y,12,28),21)` is the ISO trick for "does year *y* have 53 weeks?" — 28 December always
falls in the last ISO week. So the week counter rolls over to 1 and the year counter increments at
week 52 or 53 as appropriate. The month label is approximated as "the month containing the Monday
`(week-1)*7` days after 1 January", which is what the dashboard's `monthAbbrev()` reproduces.

**Bars are conditional formatting, not cell content.** The grid cells `F7:LR106` are empty. Fifty
`x14:cfRule` expression rules are applied to that range, one per row-within-a-block:

```
AND( MOD(ROW()-7,50)=k,
     'L3 input'!$E7<>"",
     (F$4*100+F$6) >= ('L3 input'!$I7*100+'L3 input'!$H7),
     (F$4*100+F$6) <= ('L3 input'!$K7*100+'L3 input'!$J7) )
```

for k = 0…49, each with its own fill colour. The `MOD(ROW()-7,50)=k` term is how 50 distinct bar
colours are cycled down the sheet without needing one rule per row. The last two terms are the
encoded-week range test described in §2.3.

Header-row decoration rules on `F4:LR6`:

- `(F$4*100+F$6) > ('L3 input'!$S$3*100+'L3 input'!$R$3)` → grey out columns past the plan end.
- `MOD(F$4-'L3 input'!$Q$3,6)=0` … `=5` → six alternating year-band colours, so consecutive years are
  visually distinguishable.
- `MOD(F$6-1,13)=0` (standard `dxfId=54` rule) → quarter-boundary emphasis every 13 weeks.

Today marker: `AND(F$4='L3 input'!$V$3, F$6='L3 input'!$U$3)` applied separately to `F4:LR4` (thick
red top/left/right border), `F5:LR105` (red side borders) and `F106:LR106` (closing bottom border) —
three ranges so the marker draws as one continuous vertical box around the current week.

Body greyout: the same "past plan end" rule is repeated on `F7:LR106`.

### 2.6 `master time plan` (sheet8)

Same construction, one level deeper. Frozen pane at `G7` (`xSplit=6`). Columns A–F are the row
header (A = Level, B = No., C = Activity, D = Responsible, E = Location, F = Status), G onwards is the
week grid to column `LS`.

Rows are allocated in five fixed blocks of 100:

| Rows | Level | `A` column (merged over the whole block) |
|---|---|---|
| 7–106 | L1 | `A7:A106` = "L1" |
| 107–206 | L2 | `A107:A206` = "L2" |
| 207–306 | L3 | `A207:A306` |
| 307–406 | L4 | `A307:A406` |
| 407–506 | L5 | `A407:A506` |

Each block's row-header formulas point at the matching input sheet with a row offset of 100 per
level, and each block has its own 50-rule `x14` conditional-formatting set referencing that level's
sheet. **This is why the input tables are exactly 100 rows.** Growing an input table past 100 rows
means rebuilding sheet8's block offsets and all 250 conditional-format rules — the dashboard has no
such limit (see §3.3), so prefer leaving Excel alone and treating the dashboard as the view of record
for oversized plans.

Master-range helper cells live off to the right, out of the printed area:

| Cell | Content |
|---|---|
| `LV1` | "Master Timeplan Range" caption (merged `LV1:LY1`) |
| `LV3` | `='L1 input'!P3` — start week |
| `LW3` | `='L1 input'!Q3` — start year |
| `LX3` | `='L1 input'!R3` — end week |
| `LY3` | `='L1 input'!S3` — end year |
| `LV4` | "Today" caption (merged `LV4:LY4`) |
| `LV5` | `='L1 input'!U3` — current week |
| `LW5` | `='L1 input'!V3` — current year |

The grid header then reads `G4 = $LW$3`, `G6 = $LV$3`, and the today marker uses
`AND(G$4=$LW$5, G$6=$LV$5)`. `LV3:LY3` are not editable copies — they are formulas hard-linked to
`L1 input!P3:S3`, so the Excel Gantt and the HTML Gantt (which reads `L1 input!P3:S3` directly) can
never drift apart. Change the window in `L1 input`, never in `LV3:LY3`.

### 2.7 Sheet protection

Both time-plan sheets carry `<sheetProtection sheet="1" password="DC2F" …>`. This is Excel's legacy
16-bit hash — trivially removable and intended only to stop accidental typing, not as security. The
input sheets are unprotected.

---

## 3. Dashboard: reading the workbook

### 3.1 Unzip and XML

There is no XLSX library (nothing to bundle one with, and the file must work offline from `file://`).

- `unzip(ab)` walks the zip central directory by hand; `extractFile(zip, name)` inflates one entry
  with `DecompressionStream('deflate-raw')`. If `DecompressionStream` is missing the loader throws a
  readable "use a recent browser" error.
- Sheet XML is parsed with regex, not DOMParser: `parseCells` matches
  `<c r="A1" …>…</c>` (and the self-closing form), reading `t="inlineStr"`, `t="s"` (shared strings)
  and multi-run rich text via `extractRunsText`, which concatenates **all** `<t>` runs in the node.
  Real Excel switches to shared strings on save; ignoring that broke the parser once.
- **Formula cells are read from their cached `<v>`** — there is no formula engine. Any cell the
  dashboard must read has to ship with a correct cached value, not an empty `<v>`. This is why the
  dashboard reads the input sheets (typed values) and never the time-plan sheets (all formulas, and
  openpyxl writes them with empty caches).
- `parseStyles` builds `fillColors` (from `<fills>`) and `xfFill` (cellXfs → fillId);
  `getCellColor(cells, ref, stylesInfo, fallback)` resolves the column-C fill through theme colour +
  tint + indexed palette, falling back to `PALETTE_FALLBACK[rowIndex % 50]`.
- `decodeXmlEntities` handles `&#x…;`, `&#…;` and the five named entities.

### 3.2 Layout discovery

The activity table is **discovered, not hardcoded**. `detectLayout(cells)`:

1. Scans every parsed cell for one whose `normHeader()` text maps to `'activity'` in
   `HEADER_ALIASES`; the lowest row number wins → that is the header row.
2. Maps every other cell in that row through `HEADER_ALIASES` to a field key.
3. Fills any missing key from `DEFAULT_LAYOUT.cols` (`{no:'D', activity:'E', responsible:'F',
   location:'G', startWk:'H', startYr:'I', endWk:'J', endYr:'K', tracker:'L', status:'M',
   comments:'N'}`).
4. If no activity header is found at all, returns `DEFAULT_LAYOUT` wholesale (header row 6, first data
   row 7).

`normHeader` collapses all whitespace runs to single spaces and lowercases, so `"Start\nWk"` and
`"Start Wk"` are the same key — which is why flattening the header newlines in §1 was safe.

This is what makes moved columns and a shifted table survivable, but it also means **renaming a header
in Excel silently breaks that column**. Add the new spelling to `HEADER_ALIASES` rather than renaming
things in the workbook.

### 3.3 Row scanning

`parseLevelSheet` walks from `layout.firstDataRow` to row 2000. An empty Activity cell increments a
blank streak; the scan stops once the streak reaches 30 **and** the row index is past
`firstDataRow + 100`. So the 100 pre-formatted rows are always fully scanned, a plan can grow past
them, and a genuine gap of up to 29 blank rows inside the table doesn't truncate it.

`parseLevelSheet` also reads the header block (`E2`, `E3`, `G3`, `E4`) into `level.meta`, and — for
`L1 input` — the `P3:S3` window into `meta.startWk/startYr/endWk/endYr`.

### 3.4 The Timeplan Name cell

`detectPlanNameRef(cells, layout)` scans columns A–D of every header row (rows 2 …
`layout.headerRow - 1`) for a label matching `/^(timeplan|time plan|plan)\s*name$/` after
whitespace collapsing and lowercasing, and returns `'E' + thatRow`. Falls back to `'E5'`. The result
is stored on `level.meta.planNameRef` and is what the save path writes to — never a hardcoded `E5`.
Because the match is anchored, `"Project Name"` in `B2` cannot win it.

### 3.5 Dependencies

Two-stage, and the split matters:

1. `parseDependencies(raw)` splits column O on `[;,]+`, trims, and strips surrounding straight or
   curly quotes. `parseLevelSheet` stores the result as `a.dependsOnRaw` only. It **cannot** resolve
   anything at this point — the target level may not be parsed yet.
2. `resolveDependencies(levels)` runs **after all five levels are parsed** and rewrites each
   activity's `dependsOn` into canonical `L{n}-{No.}` codes.

Resolution accepts `L{level}-{activity name}` (matched through `depNameKey`, i.e. case- and
whitespace-insensitively, within that level, first match wins on duplicate names) and the shorthand
`L{level}-{No.}`. Self-references and duplicates are dropped. Anything unmatched goes into
`a.depsUnresolved`, is listed in the load warning banner, and renders as a `"… (not found)"` chip —
never silently discarded.

> **Trap:** `resolveDependencies` must be called on *every* code path that produces a `levels` object.
> Missing it in the post-save reparse made all arrows vanish after "Save to Excel". If you add another
> path that reparses sheets, call it there too.

`detectCycles(levels)` DFSs the cross-level graph on load; cycles surface as an amber warning and do
not block loading.

---

## 4. Dashboard: rendering

### 4.1 Week axis

`getISOWeekYear(date)` implements the ISO-8601 week number (Thursday rule) in UTC.
`lastIsoWeek(year)` = `getISOWeekYear(Dec 28)` — the same trick sheet7 uses in Excel.
`buildWeeks(startWk,startYr,endWk,endYr)` walks week by week, rolling the year over at
`lastIsoWeek()`, guarded at 700 iterations. `decorateWeeks` tags each week with its month abbrev and
whether it is the current week. `buildSegments(weeks, colW, keyFn)` collapses consecutive weeks
sharing a key into the merged year and month header bands.

Everywhere a week/year pair is compared, it is encoded as `yr*100 + wk` — identical to the Excel
formulas. Keep it that way; mixed conventions is how off-by-one-week bugs get in.

`computeBar(act, weeks, colW)` maps an activity onto pixel `left`/`width` by finding the first week
index ≥ start and the last index ≤ end, clamping to the visible window and returning
`{visible:false}` when the activity falls entirely outside it.

`computeDefaultRange(levels)` returns the **global extent of the data**: the minimum encoded start
and the maximum encoded end across every activity on all five levels. A row with only one of the two
dates counts at both ends, so a half-filled activity can never sit outside the default window. Only
if no activity carries any date at all does it fall back to `L1 input!P3:S3`, and then to "this week
→ this week next year".

> The precedence used to be the other way round — `L1 input!P3:S3` first, activity extremes only as a
> fallback — which meant the shipped workbook (P3:S3 = Wk 1 2026 → Wk 52 2030, data ending Wk 48 2026)
> opened on four years of empty grid. The Excel `master time plan` sheet still draws `P3:S3`; the two
> views are *deliberately* allowed to differ here, because Excel conditional formatting cannot compute
> a min/max across five sheets but JS can. Snapshot (`state.crop`) is seeded from this on load and
> overrides it thereafter; **Reset** re-seeds from it.
`computeMasterRange(levels, crop)` — the user's Snapshot crop always wins for the live view.

### 4.2 Geometry constants

| Where | Row height | Section gap | Column width |
|---|---|---|---|
| DOM timeline (`renderVals`) | **28px** | **3px** between level sections | `density==='compact' ? 20 : 26`, ÷4 in compact-timeline mode |
| SVG export (`buildExportSVG`) | **24px** (`rowH`) | none | `baseColW = 22`, ÷4 in compact-timeline mode |

The DOM arrow overlay builds `posMap` with its own accumulator
(`yCursor += 28` per row, `+= 3` per section boundary — around line 1815). **If you change the row
height or the section divider in the template you must change that accumulator too**, or every arrow
misaligns with its bar. The export builder keeps a separate accumulator for the same reason.

`computeAutoColumnWidths(levels)` measures the real strings with a canvas 2D context and clamps:
No 28–60, Activity 140–380, Responsible 70–180, Location 70–180, Status ~measured. So the left table
sizes itself to the content instead of guessing.

### 4.3 Dependency arrows

Clicking a swatch sets `selectedCode`. `computeActivityIndex` flattens all levels into a
`code → {name, color, level, no}` map. `computeClosureNodes` BFSs the full transitive closure of the
selection (both directions), skipping hidden levels/rows. `computeClosureEdges` returns every edge
inside that closure, deduped by `from>to`.

- `direct` = the edge touches `selectedCode` → **solid, 2px**. Everything else → **dashed `5,4`,
  1.4px**.
- Both strokes and both arrowhead markers are **black (`#000000`)**.
- Markers use `markerUnits="userSpaceOnUse"` (11×11, `refX=9`). Without this the head scales with
  stroke width and solid arrows render visibly larger than dashed ones — that was a real bug.
- Only `marker-end` is ever set. An origin dot used to be drawn at each line's start and read as a
  second arrowhead; it was removed from both the DOM overlay and the SVG export. Do not add it back.
- Direction is forced top-to-bottom by rank (`level*100000 + No.`) regardless of which side of the
  Dependencies column named which activity, so an arrow never points upward.
- `pickAnchors(fromPos, toPos)` chooses horizontal side-to-side anchors when there is at least
  `max(24, |Δcy| * 0.25)` of horizontal gap, otherwise bottom-to-top vertical anchors. The gap scales
  with vertical distance because flat horizontal curves kink at the arrowhead when the rows are far
  apart.
- A transparent 14px-wide copy of each path sits underneath for hover hit-testing;
  `hoveredArrowKey` dims all other arrows to 12%.

`buildExportSVG` also draws the Timeplan Name: when `this.getEffectivePlanName(lm.key)` is non-empty
the level gutter carries two rotated `<text>` lines (key at 13px, name at 9px, offset ±6/7px about the
gutter centre) instead of one. The class method — distinct from the local `getEffectivePlanName` closure
inside `renderVals` — applies the `planNameVisible` filter, because a hidden name must not appear in
an export; the closure does not, because the toolbar box must keep showing the text while it is hidden.

`buildExportSVG` redraws the arrows independently as a string. **Any arrow-styling change has to be
made in both places** — the DOM overlay in `renderVals` and the export string builder.

### 4.4 Exports and file output

A web page cannot choose a download folder. `saveBlobWithPicker(blob, filename, description, mime,
ext)` is the single choke point for every file the dashboard produces: it uses
`window.showSaveFilePicker` where available so the user picks the destination, and falls back to
`downloadBlob` (an `<a download>` click) otherwise. It returns `'picked'`, `'downloaded'` or
`'cancelled'`; callers must handle `'cancelled'` rather than assuming success.

- `onDownloadSVG` — serialises `buildExportSVG()` and saves it through the picker.
- `onDownloadPNG` — rasterises that SVG through an `Image` + `canvas` at 2×, then the picker.
- `onExportPDF` — opens a print window containing the SVG on a single landscape Letter page. No picker
  needed: the browser's own "Save as PDF" flow already chooses the folder. For very large plans,
  rescale at print time.
- `onSaveExcel` — see §5.

---

## 5. Editing and write-back

Double-clicking a bar or a table row opens a two-step popup (`openEditPopup` → `goToDetailsStep` →
`applyDetailsStep`): weeks/years first, then optional status / activity / location / responsible
(skippable via `skipDetailsStep`). Edits live in `state.edits` keyed `L{n}-{No.}` and are applied on
read through `getEffectiveDates(levelKey, act)`, so they render red across the master timeline, the
level tables and all exports until saved.

`onSaveExcel`:

1. Applies edits to each affected sheet's XML via `setCellValue` (numbers) / `setCellString`
   (strings), using **that sheet's detected column letters** — never hardcoded ones.
2. Strings are written as `t="inlineStr"` with `xml:space="preserve"`, so the shared-string table is
   never touched and no other sheet is disturbed. Cells that don't exist yet are inserted into the row
   in the correct column order (`colLetterToNum` comparison).
3. Rebuilds the zip in place (`rebuildZip`): untouched entries are copied **byte-identical** with
   their original compressed data; only modified parts are recompressed (stored uncompressed, method
   0, with a fresh CRC32).
4. Re-parses the rebuilt workbook and **verifies the edits actually stuck** before offering it.
5. **Never overwrites the loaded workbook.** The result always goes through `saveBlobWithPicker`,
   suggesting `{baseName}_edited_{YYYY-MM-DD}.xlsx`. On Chrome/Edge that is `showSaveFilePicker`
   (the dialog reopens on the last folder used, so the user points it at `Saved Plans/` once); on
   Firefox/Safari it degrades to a plain download. A cancelled dialog aborts the save and leaves the
   pending edits intact. The old "write back in place through `this.fileHandle`" path was removed
   deliberately — an edited plan is a new revision and overwriting the master is how originals get
   lost. `this.fileHandle` is still kept from the open picker, but is no longer used for writing.
6. Re-parses into new state — and calls `resolveDependencies` (see the trap in §3.4).

The Save button outlines solid red while edits are pending, and its count is
`activity edits + dirty plan names`.

### 5.1 Timeplan Name write-back

`state.planNameOverrides[lk]` holds what the user typed in the toolbar box. `dirtyPlanNameLevels()`
returns the levels where that string differs from `levels[lk].meta.planName` (the workbook's current
value); those levels are unioned with the levels carrying activity edits into `touched`, and each
touched sheet gets both treatments in one pass before the zip is rebuilt. The name is written with
`setCellString` at `meta.planNameRef`, verified on the reparse alongside the activity edits, and its
override is deleted afterwards so the box stops reading dirty.

Both control paths matter:

- the **text box** edits `planNameOverrides` and is saved back to Excel;
- the **tick box** edits `planNameVisible` and only hides the name in the master plan and exports.

The two must not share a `<label>`. They previously did — one `<label>` wrapping the checkbox *and*
the text input, which makes the label's implicit control ambiguous. The checkbox now sits in its own
small `<label>` inside a plain `<div>`.

Because the write-back only ever touches individual `<c>` elements on input sheets, the formulas,
conditional formatting and protection on the time-plan sheets survive a save untouched. Excel
recalculates them on open (`fullCalcOnLoad="1"` is set in `xl/workbook.xml`).

---

## 6. State, props, persistence

`localStorage` keys: `tpDashboardXlsxB64` (last workbook, base64) and `tpDashboardXlsxName`. The
**Clear** button (left of Load Excel, always visible) removes both and resets `fileLoaded`, `levels`,
`edits`, `crop`, `selectedCode`, visibility toggles and the file handle. Without clearing, a stale
cached workbook is restored on load — the cause of several "it's still broken" reports after a fix.

Row visibility toggles, column toggles, the plan-name *visibility* tick boxes and the crop window are
dashboard-only. Plan-name **text** is not: `state.planNameOverrides` is savable (see §5.1).

DC props (`data-props`, surfaced as the host Tweaks panel):

| Prop | Editor | Default | Effect |
|---|---|---|---|
| `density` | enum `comfortable` / `compact` | `compact` | Week column width 26 vs 20; level gutter 44 vs 40 |
| `showCurrentWeekMarker` | boolean | `true` | Draws the today line |
| `accentColor` | color (5 swatches) | `#000000` | Selection ring colour |

`$preview` is 1400×900.

---

## 7. Accepted limitations (not bugs)

- PDF export is a single landscape Letter page; rescale at print time for very large plans.
- Row scan caps at 2000 with the 30-blank-row heuristic — roughly 150 activities per level in
  practice. The Excel time-plan sheets cap at exactly 100 per level (§2.6).
- Duplicate activity names within one level make name-based dependency references ambiguous; the
  first match wins.
- Status strings are constrained by Excel data validation, so the dashboard doesn't validate them.
- The dashboard reads cached formula results only, so a workbook written by a tool that omits caches
  will show blanks for any formula cell the dashboard depends on (currently none in the input path).

---

## 8. After any change

1. Rebuild `Time Plan Dashboard (Standalone).html` from the DC and copy it over
   `dist/06_Master_Time_Plan_Dashboard/Time Plan Dashboard.html`.
2. Re-test the full loop: load the workbook → click a swatch with multi-level dependencies → edit an
   activity → Save to Excel → confirm arrows and dependency chips survive the save.
3. If you touched the workbook, also open it in Excel and check: no `read me` sheet, no wrapped text,
   both time-plan sheets still draw bars and the today marker.
4. **Test every toolbar control against a real workbook, not by reading the code.** The plan-name bug
   survived a review because the HTML side was correct in isolation — the boxes bound, typed and
   re-rendered fine; the defect was that the *workbook* had no cell for them to read or write. Load
   the shipped `.xlsx` and confirm each control shows a real value and round-trips through Save.

## 9. Audit log

| Found | Defect | Fix |
|---|---|---|
| 2026-08 | Dashboard read the Timeplan Name from `E5`; the workbook had no such field (row 5 blank, unlabelled, unmerged). Boxes always empty, typed names never persisted. | Added the `Timeplan Name` row (`B5` label, `E5:N5` value, merged) to all five input sheets; added `detectPlanNameRef` and Excel write-back. |
| 2026-08 | Checkbox and text input shared one `<label>`, making the label's implicit control ambiguous. | Split into a `<div>` with the checkbox in its own `<label>`. |
| 2026-08 | `computeDefaultRange` preferred `L1 input!P3:S3` over the data, so the shipped workbook opened on four years of empty grid. | Global min/max across L1–L5 is now primary; `P3:S3` is the fallback. |
| 2026-08 | Timeplan Names appeared on screen but not in SVG/PNG/PDF exports. | `buildExportSVG` draws them in the level gutter. |
| 2026-08-07 | One master workbook could not support per-level ownership. | Rebuilt around one file per timeplan (§0): filename-derived level + order, plan-prefixed dependencies, IndexedDB cache, per-plan verified write-back into a chosen folder. |
| 2026-08-07 | Dependency by name hit the first match silently; a multi-plan file was rebuilt once per plan, losing edits. | No. matched first; duplicate names warned + chipped. Save groups edits by source file and rebuilds each file once. Added source-freshness check, Reload folder, PNG scale cap. |
| 2026-08-07 | Each timeplan file still carried all five input sheets plus the master Gantt. | Files rebuilt as exactly two sheets (§0.3); blank per-level templates added. |
| 2026-08-07 | Distribution folder names gave no reading order and the legacy master cluttered it. | Renumbered `01_Documentation` … `06_Master_Time_Plan_Dashboard` (§0.0); legacy master dropped from `dist/`. |
| 2026-08-08 | The heading strip was read from an arbitrary file; the top bar spent its width on text stats; no way to see knock-on effects, filter warnings, compare two captures, or print a per-plan review sheet. | v2.7 (§0.9): dashboard-owned heading + redesigned strip, pictorial top-bar stats, Warnings filters, baseline-to-baseline compare, knock-on panel + impact report, printable status sheets, 200-activity capacity. |
| 2026-08-07 | "Checks" and "Baseline" were unreadable as labels; the Timeplans panel's numeric columns had no headings; Save wrote blind. | v2.2 UI pass (§0.8): Warnings rename, explained baseline panel + pill toggle + kept-forever captures, Timeplans column headings + legend, two-step Save with a per-file change list, Ref column, dependency help block. |


---

## 0.10 Archiving before overwrite (v2.8, 2026-08-08)

`ARCHIVE_DIR = '04_Archived_Timeplans'`.

`archivePrevious(srcDir, fileName)` runs inside the write loop of `onSaveExcel`, once per output
file, **only** when `naming === 'inPlace'` **and** `state.archiveOn !== false`. It reads the bytes
currently on disk, then writes them into `archiveTargetDir(srcDir)` as
`YYYY-MM-DD_HHMM_<file>.xlsx` (date first so the folder sorts chronologically), and returns the
label+name for the confirmation message. If the source file does not exist yet it returns `null`
(nothing to keep). Any other failure **throws**, which aborts the save before `getFileHandle(…,
{create:true})` runs — the original is never destroyed by a failed archive.

`archiveTargetDir(srcDir)`:
1. `this.archiveDirHandle` if set (restored from IDB key `archiveDir` in `componentDidMount`).
   Permission is queried/requested; refusal throws with a message naming the folder.
2. otherwise `srcDir.getDirectoryHandle(ARCHIVE_DIR, {create:true})`.

The same-folder check is made **twice**: at pick time in `onChooseArchiveFolder`, and again inside
`archiveTargetDir` (`picked.isSameEntry(srcDir)` -> drop the pick and fall back to step 2). The
pick-time guard alone is not enough — `loadDirHandle` is null when the archive folder is chosen
before any timeplans are loaded, and the handle persists in IDB across sessions.

State: `archiveDirName` (display), `archiveOn` (default `true`, persisted under IDB key
`archiveOn`; only an explicit `false` is restored). Handlers `onChooseArchiveFolder` (rejects a
pick that `isSameEntry` the load folder — dated copies in the timeplan folder would be re-loaded
as timeplans) and `onToggleArchive`.

Folder numbering changed at v2.8.1: the archive lives INSIDE the timeplan folder, so the old
top-level 04 folder is gone and Exports/Dashboard moved down to 04/05. The picker button now reads
"Archive elsewhere…" — picking a folder is the exception, not the setup step.

UI: a checkbox + folder button in **Files → Output** and again as a strip above the confirm button
in the **Save** panel (`showArchiveRow` = `saveInPlace`). `saveRows[].target` and the post-save
warning both name the archive folder, and say plainly when archiving is off.

### Destructive actions are two-click, not confirm()

`onDiscardEdits` and `onClearFile` no longer call `window.confirm` (a suppressed dialog made the
button look dead). Each arms on the first click — `state.discardArmed` / `state.clearArmed`, label
swaps to "Click again to …", auto-disarms after 5 s / 6 s — and acts on the second. `onClearFile`
also resets the baseline state (`baselines`, `baselineId`, `baselineTo`, `baselineConfirm`), which
it used to leave behind. `onReloadFolder` still uses `confirm()`.
