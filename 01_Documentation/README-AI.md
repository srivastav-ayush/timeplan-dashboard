# AI / Technical Reference — Excel Timeplan HTML Dashboard

Complete reference for anyone (human or AI) maintaining, extending or debugging this project.
`README.md` is the short user-facing version.

---

# Part A — Orientation

Everything in Part A is the current, authoritative picture. Part B (§0 onward) is the detailed
reference and the dated revision log; where the log contradicts Part A, Part A wins.

## A.0 Current workbook revision (2026-08-18)

The workbooks were rebuilt on this date. Where §2 and the older revision log below describe a
different layout, THIS section is what is in the files.

**Input sheet — every `L{n} input` sheet in every workbook**

| | |
|---|---|
| Row 5 | `B5` Timeplan Name label / `E5:F5` value · `G5` **Timeplan ID** label / `H5:I5` value · `J5:L5` **Timeplan Owner** label / `M5:N5` value · `O5` Last updated label / `P5` value |
| Row 6 | blank spacer, height 9 — separates the header block from the table |
| Row 7 | column headings |
| Rows 8–207 | **200 activity rows**, `No.` pre-filled 1–200 |
| `R2:U4` | Timeplan range box (Start Wk / Start Yr / End Wk / End Yr) — moved two columns right off the Dependencies column |
| `W2:X4` | Today box: `W3 = WEEKNUM(TODAY(),21)`, `X3 = YEAR(TODAY())` |

Columns: `C` Colour · `D` No. · `E` Activity / Event · `F` Responsible · `G` Department ·
`H` Location · `I` Start Wk · `J` Start Yr · `K` End Wk · `L` End Yr · **`M` Status** ·
**`N` Tracker** · `O` Comments / Notes · `P` Dependencies. Status now sits *before* Tracker.

**Status** is the only typed column — a 14-value dropdown on `M8:M207`, in dropdown order:
Completed, Cancelled, Skipped, Risk, Blocker, Dependency, Aligned, Planned, To be discussed,
Not done, Deferred, Not Achieved, Pending, Not Completed. Each has its own conditional format
(dxf ids in `STATUS_DXF`, `tools/xlsxkit.js`), matched with `operator="equal"` so
`Not Completed` cannot trip the `Completed` rule. The dashboard's `STATUS_STYLES` uses the same
colours.

**Tracker** (`N`) is a formula, never typed — `trackerFormula()` in `tools/xlsxkit.js`, self-contained
per row, reading `TODAY()` and the row's ISO end-of-week directly:

- `Done` — Status is Completed.
- `Closed` — Status is Cancelled or Skipped (off the plan, so never late).
- `Incomplete` — the end week has passed and the row is not closed out.
- `Due this week` — the current week IS the end week.
- `On Track` — still ahead of the end week.
- blank — no activity name, or no end week/year.

Every other Status says *why* and is judged on the calendar like any other row. `TRACKER_STYLES`
in the dashboard mirrors the five outcomes.

**No. column** uses two dedicated cell formats (one per row band) with an explicit black,
non-bold 9pt font — it no longer inherits the band's coloured/bold text.

**Last updated (`P5`) is automatic.** The cell is `=TEXT(NOW(),"yyyy-mm-dd")` with a cached value,
so Excel refreshes it whenever the workbook recalculates and nobody types a date. Two consequences:
`NOW()` is volatile, so merely opening the file marks it dirty; and the dashboard's save-time stamp
(§0.6) now **skips a Last updated cell that holds a formula** rather than overwriting the automation
with a literal date.

**No sheet protection anywhere.** The time-plan sheets used to carry
`<sheetProtection sheet="1" password="DC2F" …>` (Excel's legacy 16-bit hash, `DC2F` — not a
recoverable password, and not security). It is removed from every sheet in every workbook; §2.7 is
superseded.

**Time plan sheets**: rows 7–206 (200 rows), each row's five header cells pointing at input row
`r+1` to account for the new blank row 6. The master workbook's stacked `master time plan` sheet is
five 200-row blocks, rows 7–1006.

### A.0.1 Audit, 2026-08-18

All 14 workbooks (root master, `dist` master, 5 templates, 6 demo timeplans, `tools/_test_L1.xlsx`)
were checked programmatically: row/cell tag balance, style and dxf ids in range, sheet refs resolve,
Content_Types covers every part, declared `cellXfs`/`dxfs`/`mergeCells`/`dataValidations` counts match
the elements present, no `sheetProtection`, input rows 8–207 complete with all 14 columns,
`No.` = 1–200, per-row Tracker formula identical to `trackerFormula(r)`, every typed Status inside
the 14-value list, every Colour inside the five, the dropdown string under Excel's 255-char limit,
CF rule counts and dxf ids, row-5 labels, `P5` auto formula, `W3`/`X3` today formulas, time-plan
rows 7–206 with the five header formulas pointing at input row `r+1`, grid F..LR present, the six
`extLst` range refs, bar sqref `F7:LR206`, and the master's five 200-row blocks (rows 7–1006).
Clean. `calcPr fullCalcOnLoad="1"` is present in every workbook, so the empty cached `<v>` on the
Tracker and header formulas is filled the moment Excel opens the file.

Two changes came out of the audit:

1. **Input sheets now freeze at row 7** (`<pane ySplit="7" topLeftCell="A8" state="frozen"/>`). With
   200 rows the column names scrolled away. `buildInputSheet` adds it.
2. **`detectLabeledRef` scanned only columns A–N** (`col > 14` → `continue`), so the "Last updated"
   label at **O5** was never seen and every plan reported *No Last updated date*. The bound is now
   `col > 16` — the header block ends at P, and R onward is the week/range box, which must stay
   outside the scan.

An end-to-end load of the six demo timeplans through the shipped standalone was verified in-browser:
6 plans / 30 activities parsed, header block read, cross-file dependencies resolved
(`L3_Tooling-1 → L2_Bodyshop-1`), Tracker computed (`Incomplete`, `Done`, `On Track` with week
counts), Risks/Dependency-risk/File-checks panels populated, per-plan tabs render every column.

The formulas that feed the time-plan sheets test their **own** source column, not the Activity
column: `C7 = IF('L3 input'!F8<>"",'L3 input'!F8,"")`. Only `A` (No.) and `B` (Activity) gate on
Activity. That is the original workbook's design and is intentional.


## A.1 How to read this document

| If you are… | Start at |
|---|---|
| new to the project | A.2, A.3, A.4 — then §2 for the workbook and §3 for the parser |
| changing the Excel side | §2 (workbook construction), then §3.2 layout discovery |
| changing how it draws | §4 (rendering), §4.2 geometry constants |
| changing save / write-back | §5, then §0.6 and §0.10 |
| changing what persists | §0.14 and A.7 |
| debugging "my file did not load" | A.6 (the skip rules), then §3.3 |
| adding a column | A.8 (the checklist) |
| deploying to a team | A.9 |

## A.2 What the thing is

One HTML file, no build step, no dependencies, no network. It:

1. reads N Excel workbooks (one per timeplan) out of one folder,
2. parses the xlsx zip itself — no SheetJS, no library at all,
3. stacks every plan into one Gantt ordered by the `L{n}` in the file name,
4. resolves cross-file dependency references and reports the ones that cannot work,
5. lets a user edit dates and writes each changed workbook back **rebuilt from its own bytes**,
6. keeps a dated copy of every version it replaces.

It is a Design Component (`Time Plan Dashboard.dc.html`) that is inlined into a single
standalone HTML file for distribution. The standalone is the artefact users get; the DC is the
source. **Edit the DC, then re-inline — never hand-edit the standalone.**

## A.3 The five invariants

Break any of these and something downstream breaks silently. They are the load-bearing rules.

1. **One .xlsx = one timeplan.** The level and the sort order come from the *file name*
   (`L{1-9}_Name.xlsx`), never from anything inside the file.
2. **A file is rebuilt from its own bytes on save.** Edits to `L2_Bodyshop` can never reach
   `L1_Programme`'s file. Every rebuild is re-parsed and verified before a byte is written; a
   value that does not round-trip aborts the whole save.
3. **Excel owns colour, status and structure. The dashboard owns the view.** There is no colour
   picker in the dashboard — it reads what the workbook says. Milestones, markers, heading,
   logo, ticks and week window are dashboard-owned and never written into a workbook.
4. **Tracker is recomputed, never read.** The identical rule lives as a self-contained formula in
   the xlsx and as JS in the dashboard, so the two can never disagree. The cell's cached value is
   empty until Excel recalculates, which is why the file's value is ignored.
5. **Identity is Plan ID first, file name second, activity No. third.** This is what survives a
   file rename in baselines and dependency resolution.

## A.4 Data flow, end to end

```
  folder picker / drag / plans.json over http
        |
        v
  readFileList / directory scan        skip rules -> "skipped" list -> Risks > File checks
        |
        v
  unzip (inflate raw, no library)  ->  xl/workbook.xml, sharedStrings, styles, sheetN.xml
        |
        v
  parsePlanFileName(name)          ->  {base, level, name}
  layout discovery on "L{n} input" ->  header row, column map, first activity row
        |
        v
  row scan (up to 2000 rows)       ->  activities[]  {no, name, resp, dept, loc, s/e wk+yr,
        |                                             status, colour, comments, deps[]}
        v
  resolveRef() per dependency      ->  edges[]  +  unresolved[] -> File checks
        |
        v
  week axis (min..max, <=700)      ->  render master / per-plan tabs / SVG / PNG / PDF / A4
        |
        v
  edits -> pendingEdits (red)      ->  Save to Excel: archive old, rebuild, verify, write
```

## A.5 Where state lives

| Store | Key | Holds | Cleared by |
|---|---|---|---|
| `localStorage` | `tpDashboardSessionV1` | the whole dashboard view record | Files ▸ Reset view |
| `localStorage` | `tpDashboardSavedBy` | the name stamped into `Dashboard_Settings.json` | not cleared by Reset view |
| `IndexedDB` | `planFiles` | last-loaded workbook bytes, for offline reopen | Clear |
| `IndexedDB` | `outDir` | the save-target directory handle | Clear |
| the folder | `Dashboard_Settings.json` | the team's shared view record | deleting the file |
| the folder | `_baselines/*.json` | dated snapshots | deleting the files |

Pending Excel edits are deliberately in **none** of these. A changed date is a change to a
timeplan, so it stays in the red *Save to Excel* count until it is written into the file.

## A.6 Why a file did not load

The folder scan skips, and reports, in this order. All of it surfaces under Risks ▸ File checks —
nothing is dropped in silence.

| Skipped | Reason given |
|---|---|
| `~$Anything.xlsx` | Excel lock file — filtered before the list is even built |
| not `.xlsx` | "it is not an .xlsx file" |
| in a subfolder | "only files at the top level of the folder are loaded" |
| `… (1).xlsx`, `conflicted copy`, `Copy of …` | "it looks like a sync conflict or a duplicate copy" |
| not `L{1-9}_Name` (strict mode) | "the name does not follow L{level}_Name.xlsx" |

The conflict pattern matters on a synced drive: Google Drive, Dropbox and OneDrive all resolve a
double-write by keeping both files under a decorated name. Loading both would silently duplicate a
plan, so they are reported instead.

## A.7 The shared settings file, and the overwrite guard

`Dashboard_Settings.json` is **one file per folder for the whole team**, which makes Save Dashboard
a multi-user write. The record carries two fields that are excluded from the change-detection body
(`sessionBody()` strips both, or the record would rewrite itself every render):

- `savedAt` — ISO 8601. Written on every save.
- `savedBy` — free text, prompted once per browser and kept in `tpDashboardSavedBy`.

On `onSaveSessionFile` the folder's existing file is re-read *before* the write. If its `savedAt`
is greater than `state.sessionSavedAt` (the stamp of the record this browser is working from),
somebody else has saved since — a `confirm()` names them and the time and asks before replacing.
Declining writes nothing and points at Files ▸ Load view…. ISO 8601 strings compare correctly
lexicographically, so no date parsing is involved.

On load, the folder record is applied only when its `savedAt` is newer than the browser's own.
The Files panel stamp reads "Saved by X — <time>" when a name is present.

This is a *guard*, not a lock: it cannot stop two people writing within the same read window. The
documented mitigation is organisational — nominate one owner of the team view (README §8).

## A.8 Adding a column: the checklist

A new activity column touches seven places. Miss one and it half-works.

1. `02_Templates/*.xlsx` — the column, its width, its formatting, rows 7–106.
2. The header alias map (`'dependencies':'dependencies', 'depends on':'dependencies'` etc.) so
   layout discovery finds it whatever the sheet calls it.
3. The row scanner — read it into the activity object.
4. `View ▸ Columns` — the toggle, and its default.
5. The master-plan and per-plan tab renderers — heading cell **and** body cell, both `flex:none`
   with matching widths (see the alignment note in the 2026-08-15 log entry).
6. The SVG / PNG / PDF exporters and the A4 status sheets.
7. Write-back, if it is editable — the cell write, and the round-trip verification.

## A.9 Deployment on a synced drive

The intended deployment is a shared folder mirrored to every machine (Google Drive for desktop,
OneDrive, Dropbox), each person opening their own local copy of the HTML.

- **Mirror, not stream.** Streaming mode leaves placeholder files; a directory scan then triggers N
  on-demand downloads, which is slow and fails with no connection. Mirror mode is a hard
  requirement for the offline promise.
- **The HTML cannot be run from the provider's web viewer.** Drive serves it as text. It must be
  opened from the local mirror, over `file://`.
- **`file://` means no `plans.json`.** `fetch` is blocked on `file://`, so `autoLoadManifest()`
  quietly no-ops and the folder picker is the only way in. `plans.json` is dead weight in every
  folder-based deployment; it exists for the http case only.
- **Permissions belong per subfolder** — edit on `03_Time_Plans`, read-only on `02_Templates` and
  the dashboard folder.
- **The archive folder syncs.** `04_Archived_Timeplans/` grows one workbook per save and
  replicates to every machine. *Archive elsewhere…* points it outside the synced tree.
- **Excel holds a lock.** An open workbook cannot be cleanly overwritten or synced; the
  stale-check (size + mtime compared against load time) catches the read side, not the lock.

## A.10 After any change

Non-negotiable, in this order:

1. Edit the DC, never the standalone.
2. Re-inline to `Time Plan Dashboard (Standalone).html`.
3. Copy that to `dist/05_Master_Time_Plan_Dashboard/Time Plan Dashboard.html`.
4. Load the six example timeplans and check: master plan draws, every tab draws, Risks has all
   three tabs, an edit goes red and saves, SVG exports.
5. Append what changed to §9, dated.

---

# Part B — Reference and revision log

## 0.0 Distribution folder layout (v2.2)

The shipped folder is numbered so it sorts in workflow order. This is the layout every path in
this document and in `README.md` refers to:

```
README.md                      user guide / GitHub landing page, at the pack root
Time_Plan_L1-L5_Master.xlsx    the legacy all-in-one workbook, at the pack root
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

1. `parsePlanFileName` reads a leading `L{1-9}` (L1–L5 is the convention; L6–L9 work and are lightened per palette cycle) off the base name → `{base, level, name}`.
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

### "Checks" → "Warnings" → **Problems** (superseded twice)

User-visible strings only, and now superseded: the standalone Warnings button no longer exists. The
health list is the **File checks** tab of the **Problems** panel (`panel==='issues'`) — see
*The seven-button toolbar* at the end of this file. The internal names are unchanged —
`buildHealth`, `health`, `healthRows`, `healthByPlan`, `healthSev`, `healthPlan`, `sevTabs`.
**Do not rename them** unless you rename them everywhere. `healthBtnLabel`, `healthTitle` and the
panel key `'health'` are gone.

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

`APP_VERSION` was the string `'Version 2.8'` at this revision; it is now pinned to `'Version 1'`
(rendered verbatim in the top bar pill, so it must stay human-readable — do not put a bare `v` back).

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

Two pure functions near `slipText` (surfaced as the **Delays** tab of the **Problems** panel; the
separate "Knock-on" button it was introduced with is gone):

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
with spacer rows). The shipped files carry *row formatting* to row 216 (input) / 206 (time plan), but the Tracker
formula, the Colour and Status dropdowns, the status colouring and the Gantt bar rules only cover
**rows 7–106 — 100 activities**. `buildHealth` therefore warns above `TEMPLATE_ROWS` (100), naming
both sheets to copy the last row down. Nothing in the dashboard itself caps at 100 or 200.

---

## 1. Files

| File | Role |
|---|---|
| `dist/03_Time_Plans/L*.xlsx` | The 6 demo timeplans (1×L1, 1×L2, 1×L3, 2×L4, 1×L5), each two sheets, listed in `plans.json`. |
| `Time_Plan_L1-L5_Master.xlsx` | The original single workbook (five input sheets + `L3 time plan` + `master time plan`). Migrated 2026-08-14 to the same shape as the per-plan files: Department at G, Plan ID (`L{n}_Master`) and Last updated on row 5, tracker/status/dependency refs shifted, all `master time plan` block formulas and bar rules repointed. Loads as five plans (§0.3 case 3) — but only with the `strictFileNames` prop set to `false`, because the file name does not start with `L{level}`. Shipped at the pack root as well as the project root. |
| `Time Plan Dashboard.dc.html` | **Source of truth for the UI.** A Design Component: template + `class Component extends DCLogic`. Edit this. |
| `Time Plan Dashboard (Standalone).html` | Self-contained bundle of the DC — this is what ships (copied to `dist/05_Master_Time_Plan_Dashboard/Time Plan Dashboard.html`). Regenerate after every DC change; never hand-edit. |
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
| `E5` | **Timeplan Name** (merged `E5:F5`) — this level's plan name, e.g. "L1 Programme Master Plan" |
| `G5` / `H5` | "Plan ID" label / the stable id, e.g. `L3_Tooling` (value merged `H5:I5`) — set once, never changed |
| `J5` / `K5` | "Last updated" label / `YYYY-MM-DD` (value merged `K5:N5`); a save re-stamps it (`meta.updatedRef`) |

> Row 5 used to be one merge, `E5:N5`, which covered the Plan ID and Last updated cells: the
> dashboard read them but nobody could see or type them in Excel. The merge is now `E5:F5` +
> `H5:I5` + `K5:N5` in all eleven pack workbooks and in the master.
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
| C | Colour | Data-validation **dropdown** of colour names (Red/Yellow/Green/Blue/Purple). The name drives the cell fill, the Excel Gantt bars and the dashboard bars. Empty = grey "not chosen". See *Activity colour is an Excel dropdown*, below. |
| D | No. | Short ID. Canonical dependency reference is `{file base name}-{No.}` (§0.4). |
| E | Activity / Event | Name. An empty E cell ends the row for every downstream formula and for the parser. |
| F | Responsible | Free text. |
| G | Department | Free text. Added 2026-08 — everything from the old G rightwards moved one column. |
| H | Location | Free text. |
| I / J | Start Wk / Start Yr | ISO week + year. **Primary date entry.** Validation: whole 1–53 / whole 1000–9999. |
| K / L | End Wk / End Yr | ISO week + year. Same validation. |
| M | Tracker | Formula, see §2.3. Derived, never typed. Not read by the dashboard — recomputed in JS. |
| N | Status | Data-validation dropdown: `TBD, Planned, Aligned, Dependency, Risk, Blocker, Completed`. |
| O | Comments / Notes | Free text, not rendered anywhere. |
| P | Dependencies | `{file base name}-{No.}`, several separated by `;` or `,`. |

There is **no Start Date / End Date column**. An earlier revision added real date serials at L/M and
shifted Status/Comments/Dependencies to O/P/Q; that was reverted. If you find a workbook shaped that
way it is stale — do not reintroduce it.

### 2.3 The `Tracker` formula (column M)

Row 7, filled down to row 106. Four outcomes: `Completed`, `Delayed Nw`, `Due this week`,
`On Track`, plus blank when there is nothing to judge.

```
=IF($E7="","",
   IF($N7="Completed","Completed",
     IF(OR($K7="",$L7=""),"",
       IF(TODAY()>SUN, "Delayed "&ROUNDUP((TODAY()-SUN)/7,0)&"w",
         IF(TODAY()>=MON, "Due this week", "On Track")))))

   where MON = DATE($L7,1,4)-WEEKDAY(DATE($L7,1,4),3)+($K7-1)*7   Monday of the end week
         SUN = MON+6                                              Sunday of the end week
```

(MON and SUN are written out in full in the file — Excel has no LET in the baseline target.)

Reading it: blank row or no end date → blank. `Completed` in the Status column → `Completed`,
early or late. Otherwise the activity's end week is turned into a **real date range** and compared
against `TODAY()`: past its Sunday → `Delayed Nw`; inside the week → `Due this week`; before it
→ `On Track`.

`DATE(yr,1,4)` is the ISO anchor (Jan 4 is always in ISO week 1) and `WEEKDAY(...,3)` returns
0 = Monday, so `DATE(yr,1,4)-WEEKDAY(DATE(yr,1,4),3)` is the Monday of ISO week 1 of that year.

#### Why it was rewritten (2026-08-14)

The previous formula was `($V$3*100+$U$3) > (L7*100+K7)` with `U3 = WEEKNUM(TODAY(),21)` and
`V3 = YEAR(TODAY())`. Two defects:

1. **It depended on two helper cells in plain sight.** `U3`/`V3` sit unprotected on row 3 of the
   input sheet. Clear, overwrite or paste-values over them and the left-hand side collapses to
   `0`, so `0 > 202612` is FALSE and **every row reads "On Track" forever**, no matter how late.
   This was the reported bug. The formula is now self-contained per row and reads `TODAY()`
   directly — nothing outside the row can break it.
2. **`YEAR(TODAY())` is the calendar year, not the ISO week-year.** In the days around New Year the
   two differ, so `year*100+week` mis-ordered: 1 Jan 2027 is ISO week 53 of 2026, but the formula
   encoded it as `202753`, marking activities due in early 2027 as delayed. Comparing real dates
   removes the boundary entirely.

Also: `"Completed"` used to render as `"On Track"`, which lost the distinction between *done* and
*not due yet*; and there was no signal at all for *due right now*.

Four conditional-format rules on `M7:M106` (reusing existing dxfs, no styles.xml change):
`Delayed` → dxf 72 (bold italic dark red on pink), `Due this week` → dxf 59 (bold amber),
`Completed` → dxf 64 (bold green), `On Track` → dxf 61 (teal). `Delayed 21w` matches the
`Delayed` rule by `containsText`, so the week count needs no extra rule.

#### The dashboard mirror

`trackerFor(act)` in the DC implements the identical rule and `TRACKER_STYLES` the identical four
colours. It is **computed, never read from column M** — a formula's cached `<v>` in the file is
empty until Excel recalculates, so the cell is worthless to a parser. The comparison uses the
existing `absWeek()` / `getISOWeekYear()` pair (the same absolute-week scale the Risks view uses),
which is ISO-correct by construction, so the two views agree row for row.

Surfaced in: the master plan's left grid (last column, after Status), each timeplan tab's table,
the SVG/PNG/PDF export, and the printed status sheets. Toggle under **View ▸ Columns**
(`columnVisible.tracker`, default on).

If you change the four labels or colours, change them in **both** places — the Excel formula and
`TRACKER_STYLES`/`trackerFor` — or the sheet and the dashboard will disagree.

### 2.4 Status colouring

`N7:N106` carries seven `containsText` conditional-format rules, one per status value, each pointing
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
D7 = IF('L3 input'!E7<>"",'L3 input'!H7,"")     ' Location
E7 = IF('L3 input'!E7<>"",'L3 input'!N7,"")     ' Status
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

**Bars are conditional formatting, not cell content.** The grid cells `F7:LR106` are empty.
*Superseded:* the fifty `MOD(ROW()-7,50)` rules described here were replaced by six colour-name
rules per block when the colour column became a dropdown — see *Activity colour is an Excel
dropdown* at the end of this file. The original construction was fifty `x14:cfRule` expression
rules applied to that range, one per row-within-a-block:

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
blank streak; the scan gives up only once the streak passes **80** — there is no row-count condition
any more. So the pre-formatted rows are always fully scanned, a plan can grow well past them, and a
genuine gap of up to 80 blank rows inside the table does not truncate it.

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
- Row scan caps at row 2000 with the 80-blank-row heuristic. The shipped .xlsx files are only
  formula-filled to row 106 (100 activities) — past that the Excel side needs the last row copied
  down; the dashboard is unaffected and warns (`TEMPLATE_ROWS`).
- Duplicate activity names within one level make name-based dependency references ambiguous; the
  first match wins.
- Status strings are constrained by Excel data validation, so the dashboard doesn't validate them.
- The dashboard reads cached formula results only, so a workbook written by a tool that omits caches
  will show blanks for any formula cell the dashboard depends on (currently none in the input path).

---

## 8. After any change

1. Rebuild `Time Plan Dashboard (Standalone).html` from the DC and copy it over
   `dist/05_Master_Time_Plan_Dashboard/Time Plan Dashboard.html`.
1b. Re-capture `dist/docs/dashboard-*.png` if the change is visible — the README shows them.
2. Re-test the full loop: load the workbook → click a swatch with multi-level dependencies → edit an
   activity → Save to Excel → confirm arrows and dependency chips survive the save.
3. If you touched the workbook, also open it in Excel and check: no `read me` sheet, no wrapped text,
   both time-plan sheets still draw bars and the today marker.
4. **Test every toolbar control against a real workbook, not by reading the code.** The plan-name bug
   survived a review because the HTML side was correct in isolation — the boxes bound, typed and
   re-rendered fine; the defect was that the *workbook* had no cell for them to read or write. Load
   the shipped `.xlsx` and confirm each control shows a real value and round-trips through Save.

## Risks, and what moved where (2026-08-14, later the same day)

**`impactChanges()` gained a third trigger, and it is the important one.** It used to fire only on
an unsaved edit or a move away from the baseline, so a dashboard opened on a folder of stale files
reported nothing at all — the single most common real situation. It now also treats an activity as a
cause when **its end week has passed and its status is not `Completed`**:

```js
const done = /^\s*completed\s*$/i.test(String(eff.status||''));
if(!done && nowAbs!=null && newEndAbs < nowAbs) out.push({… source:'Late', late:true,
  delta: nowAbs-newEndAbs, newEndAbs: nowAbs });
```

`newEndAbs` is set to **this week**, not the missed week: the earliest something unfinished can now
finish is now, and `computeImpact` walks the chain from that. `nowAbs` comes from
`absWeek(getISOWeekYear(new Date()))`. Precedence per activity is edit → late → baseline, one cause
each, so the same activity never appears twice.

**The panel is cause-above-effect, not two tables.** `riskRows` is a *flat* list of two row kinds —
one `cause` row per `impactChange` that has successors, then its `kid` rows indented under it — and
each row carries its own `rowStyle` / `refStyle` / `nameStyle` / `metaStyle` / `pillStyle` so a
single `<sc-for>` renders both. Causes are ordered: has-a-must-move kid, then has-kids, then by
delta. Causes with no successors are counted into `orphanCauses` and summarised in one line
(`riskOrphanText`) instead of padding the list.

**Counts are distinct activities, not rows.** `impMust`/`impReview` come from `Set`s of successor
`code`s, and the review set has the must set subtracted (an activity can be `must` under one cause
and `review` under another). `impRows.filter(...).length` over-counts, because a shared successor
appears once per cause — right for `riskRows`, wrong for the number quoted in a review. Latent
before; the `late` trigger produces many more causes, so it showed on screen (8 rows vs 6
activities in the demo folder).

Renames: button **Problems → Risks** (`'Risks '+n` / `'No risks'`), tab **Delays → At risk**,
hole `issuesTabDelays → issuesTabRisks`, state `issuesTab: 'delays' → 'risks'`. Removed holes:
`impactHasChanges`, `impactChangeRows`, `impactHasRows`, `impactRowsView`, `impactEmpty`,
`impactEmptyText`. Added: `riskRows`, `riskEmpty`, `riskEmptyText`, `riskHasOrphans`,
`riskOrphanText`, `goToChecks`, `checksNudgeLabel`, `checksNudgeStyle`. The status sheet's per-plan
section and chip are now headed **File checks**.

`checksNudge*` is the answer to "nobody will find the second tab": a tinted full-width button at the
foot of the At risk tab that names the file-check count and switches tab.

### Heading strip owns the logo; Files got shorter

The logo moved out of the Files panel into the **Edit** strip (`hdrEdit`), next to the fields it
belongs with — same `onPickLogo` / `onResetLogo` / `onToggleLogo` / `logoPreviewStyle` /
`logoSourceLabel` / `logoToggleLabel` / `hasLogoNote` values, just mounted there instead.
`infoBarHeight` is `hdrEditing ? 240 : 72` (was 152) to fit the extra row.

Files now has three sections instead of five: **Source** (with `Reload folder` always shown),
**Where saves go** (one sentence; the archive tick and the archive-folder picker were duplicates of
the ones in the Save panel and are gone from here), **Dashboard view** (was "Dashboard settings" —
renamed and re-explained around *it comes back on its own; use Load view only if it did not*), plus
Clear. `Save settings file` is gone from the panel; the top-bar **Save Dashboard** is the one way.

`onReloadFolder` no longer needs a directory handle: with no `loadDirHandle` it re-runs
`autoLoadManifest()`, so Reload works for a dashboard served over http as well as one whose folder
was picked. Only a set loaded file-by-file cannot reload, and it says so.

## The seven-button toolbar (2026-08-14)

The bar was nine buttons of white-with-a-border, three of which opened panels the user read as one
job. It is now seven filled buttons in two groups, split by a hairline:

| Button | Panel key | Fill | Holds |
|---|---|---|---|
| **Problems** / **No problems** | `issues` | `#B3261E` if `impMust` or `healthErrors`, else `#B26A00` if anything at all, else `#1E8449` | Both former panels. Two pill tabs — **Delays** (`impactChanges` + `impactRows`) and **File checks** (`buildHealth`) |
| **Baseline** | `baseline` | `#2E5C8A` | Unchanged behaviour, shorter copy |
| **View** | `view` | `#1F3864` | The former Timeplans panel *and* the former View panel in one scroller: filter, show/hide, rename, gutter label, Age, Acts, badge, then Columns, Timeline scale, Week window |
| **Files** | `files` | `#6E6A5F` | Source, output, archive, logo, settings load/reset, clear |
| **Save Dashboard** | — | `#166534` | Calls `onSaveSessionFile` directly (no panel). The duplicate button inside Files is gone |
| **Save to Excel (n)** | `save` | `#B3261E` when dirty, `#6E6A5F` when clean (white on `#7D7A70` was 4.29:1, under AA) | Unchanged two-step |
| **Export** | `export` | `#9A3412` | Unchanged |

`fillBtn(fill, active)` builds them all: white text, `mixHex(fill,'#000',0.12)` border, and while
that button's panel is open the fill darkens 24% with an inset shadow. It is declared partway down
`renderVals`, so `saveBtnStyle` — computed earlier in the same method — repeats the literals rather
than calling it (a TDZ error otherwise).

The pre-load empty state has the same constraint from the other direction: its wrapper is
`flex:1;min-height:0;overflow:auto` with `align-items:flex-start` and `margin:auto` on the card.
Centring it with `align-items:center` clips a too-tall card at the **top** as well, which
`overflow:auto` cannot scroll back to — `margin:auto` keeps it centred when it fits and lets it
scroll when it does not.

Panel heights are `max-height: calc(100vh - 170px)`, not a `vh` share: the panels are anchored at
`top:100%` of the toolbar (y≈158px), so `78vh` overflowed the window bottom and put the Problems
footer buttons out of reach. Any new panel here must subtract the anchor the same way (all six panels — issues, view, baseline,
save, files and export — now use the same `calc(100vh - 170px)`; `panelFiles` had been subtracting
130 and overflowed by 28px at a 540px window).

Removed keys: `healthBtnStyle`, `healthBtnLabel`, `healthTitle`, `plansBtnStyle`, `plansBtnLabel`,
`impactBtnStyle`, `impactBtnLabel`, `barBtn`, `panelPlans`, `panelHealth`, `panelImpact`,
`onOpenPlansPanel`, `onOpenHealthPanel`, `onOpenImpactPanel`. Added: `fillBtn`, `panelIssues`,
`onOpenIssuesPanel`, `issuesBtnStyle`, `issuesBtnLabel`, `issueTabs`, `issuesTabDelays`,
`issuesTabChecks`, `issuesIntro`, `viewBtnLabel`, `viewPanelSub`, `dashSaveBtnStyle`,
`dashSaveBtnLabel`. New state: `issuesTab` (`'delays'` | `'checks'`, defaulting to `'checks'` when
there are no delays but there are file checks).

Wording, all user-facing: verdicts are `Move 8w later` / `Tight — 2w spare` / `5w spare` (were
`Must move 8w` / `Review — only 2w float left` / `5w float left`); the checks severity pills are
**All** / **Must fix** / **Worth a look**; column heads in the delays table are **Knocked on** /
**Activity** / **Runs** / **What to do**; the baseline panel opens with "A baseline is a photo of
today's dates" and its \+3w/−2w legend now sits next to the toggle it describes; the load banner and
the per-file badge tooltip point at **Problems**; the printed status sheet's per-plan section is
headed **Problems**.

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
| 2026-08-15 | The toolbar was eight solid-filled colour buttons in one undifferentiated row: everything shouted, so nothing did, and there was no way to tell a tool from an action. | Regrouped into four bordered clusters (Files/Export · Save Dashboard/Save to Excel · View · Risks/Baseline) separated by rules. Buttons are now outline-only at rest (hue in the border and label, white ground) and fill with their hue for any of three reasons — hovered, their panel is open, or they are *asking* — via `outlineBtn(fill, active, attention, key)`. **Hover is state-driven** (`btnHover` + `onBtnEnter(key)`/`onBtnLeave`), not CSS: `style-hover` is compiled at template-parse time and so accepts only a literal declaration string, silently producing an empty rule when given a `{{ }}` hole. State also lets hover compose correctly with the darkened panel-open look, which an `!important` CSS rule would have overridden. |
| 2026-08-15 | **Save Dashboard gave no indication that the view had drifted from the settings file**, unlike Save to Excel which has always carried a pending count. | Added `sessionMark` (the record body at the last point screen and file agreed) plus `sessionDirtyParts()`, which diffs the live record against it and returns changed buckets named the way a person would say them (heading, milestones, markers, columns, logo, timeline, week window, what is shown, baseline, open tab). The button shows `Save Dashboard · N` and fills green whenever N > 0, and its tooltip lists what changed. `sessionMark: null` means "clean from here" and is re-stamped on the next update — set by every path that puts the two back in sync (restore from storage, folder pickup, file write, file load, reset). |
| 2026-08-15 | **The Risks list had unlabelled columns**, and its one number column mixed two different meanings, so the reader could not tell what any figure was measuring. | Rebuilt on a shared 5-column grid with a sticky header: Timeplan · Activity · Planned · Late by / gap was · Must move. Cause rows are shaded (red for late, blue for moved) with the count waiting as a chip; effect rows sit under them on the same grid. |
| 2026-08-15 | **The "gap" figure was wrong by construction.** `computeImpact` only carried the predecessor's *projected* end down the chain, so "slack" already had the lateness baked into it — a 21w-late cause made a comfortable 3w gap read as a 24w overlap. | The walk now carries two ends per predecessor: `endAbs` (where it will actually finish, after shifts) and `pEndAbs` (where the plan said it would). `slack` still drives must-move/shift; the new `plannedSlack` answers "how much room was this relying on" and is what the gap column shows. `impactChanges()` records `plannedEndAbs` on all three change kinds. |
| 2026-08-15 | The printed risk list was two disconnected tables (what changed / what it affects), leaving the reader to join them up by reference number by hand. | `buildImpactReport()` rewritten as one cause-then-effects table matching the panel, with a three-part legend defining the two number columns, summary chips, repeating table header, `page-break-inside:avoid` per row, an owner column, and a separate closing table for late activities nothing links to. |
| 2026-08-15 | The **timeplans** and **activities** pills used multi-line `title` attributes for their file lists — browser-timed, line breaks dropped on some platforms, silently truncated. | Both now use the same custom hover card as the version pill (`pillHover` state, `statPlansRows` / `statActsRows`), with colour swatches, per-file activity counts and a hidden marker. |
| 2026-08-14 | **Tracker read "On Track" on activities that were weeks overdue.** The formula compared `$V$3*100+$U$3` against the end week, and `U3`/`V3` are ordinary unprotected cells on row 3 — once cleared or pasted over, the left side was `0`, so nothing ever compared as late. It also used `YEAR(TODAY())` rather than the ISO week-year, mis-ordering the days around New Year, and collapsed *Completed* into *On Track*. | Replaced in all 5 templates, all 6 demo plans and the 5 input sheets of the master workbook (1800 cells) with a self-contained per-row formula that converts the end week to a real date range and compares `TODAY()` against it: `Completed` / `Delayed Nw` / `Due this week` / `On Track` / blank. Four CF rules on `M7:M106`. Mirrored in the DC as `trackerFor()` + `TRACKER_STYLES` and surfaced as a Tracker column on the master plan, the timeplan tabs, the SVG/PNG/PDF export and the status sheets, toggleable under View ▸ Columns. |
| 2026-08-14 | `setCellValue` kept a blank template cell's `t="inlineStr"` when writing a number into it, so Excel read the written week as empty *text* — quietly breaking the Tracker formula and the Gantt bar rules for that row. | The stale `t="inlineStr"`/`"s"`/`"str"` is now stripped on numeric writes. |
| 2026-08-14 audit | A timeplan row carrying a warning badge pushed **Age** and **Acts** out of line with the column headings (the badge was an `sc-if` with no reserved slot; wrapping it in a plain spacer span does not work — the runtime drops an element whose only child is a false `sc-if`). | The badge span is always rendered, transparent when the file is clean, and the header row gained a matching 16px spacer. |
| 2026-08-14 audit | After a save the reparsed plan kept the **old** `updatedAt`, so the Age column, the freshness warning and the status sheets still showed the pre-save date even though the save had just stamped Last updated in the file. | The post-save plan recomputes `updatedAt` with `parseUpdatedValue(reparsed.meta.updatedRaw)`. |
| 2026-08-14 audit | The capacity warning fired above 200 activities, but the shipped .xlsx files are only formula-filled / validated / bar-ruled to row 106. A plan of 150 activities silently lost its Excel Gantt, dropdowns and Tracker column with no warning. | `TEMPLATE_ROWS = 100`; the warning quotes the real limit and names both sheets to copy down. Docs corrected (§0.9, §7, README *Limits*). |
| 2026-08-14 audit | `dist/docs/dashboard-*.png` still showed v2.8: twelve timeplans, no Department column, orange bars, no Milestones/Markers tabs. README leads with them. | All three re-captured from the current build. |
| 2026-08-14 | Every dropdown panel was sized as a `vh` share although all of them hang off the toolbar at y≈158, so at a 540px window the three merged panels overflowed and the Problems footer buttons could not be reached; Files was subtracting 130 not 170; Export had no cap at all; and the pre-load instructions card was clipped by 125px with nothing scrollable. | All six panels on `calc(100vh - 170px)`; empty-state wrapper made scrollable with a `margin:auto` card. |
| 2026-08-14 | The risk headline, the tab pill and the button counted cause→effect ROWS, so an activity waiting on two late things was counted twice ("8 activities" when 6 were affected). | `impMust`/`impReview` are now distinct-code sets, review minus must. |
| 2026-08-14 | The knock-on analysis only fired on an unsaved edit or a baseline move, so a folder of stale timeplans — the normal case — reported nothing at risk; the panel presented cause and effect as two disconnected tables; the logo sat in Files, away from the heading fields it belongs with; Files had five sections, two of them duplicating the Save panel; Reload folder was hidden whenever the dashboard was served over http. | Overdue activities are now a cause (`late`); the list is cause-above-effect; **Problems → Risks**; logo moved into **Edit**; Files cut to three sections; Reload folder falls back to re-fetching the manifest. |
| 2026-08-14 | Nine outline buttons, with Warnings/Knock-on/Timeplans reading as three jobs the user treats as one, and Save-settings buried in Files. | Seven filled buttons; Problems merges Warnings + Knock-on behind two tabs; View merges Timeplans + View; Save Dashboard promoted to the bar (see *The seven-button toolbar*). |
| 2026-08-14 audit | Row 5 was merged `E5:N5`, hiding the **Plan ID** and **Last updated** cells behind the Timeplan Name band — the dashboard read them, but in Excel they were invisible and untypable, while the README told owners to fill them in. | Row 5 is now three merges (`E5:F5`, `H5:I5`, `K5:N5`) in all five templates, all six timeplans and the master. |
| 2026-08-14 audit | The duplicate-file-name check counted **plans**, not source files, so the legacy master (one file, five plans) always warned "Two loaded files are called Time_Plan_L1-L5_Master.xlsx. Rename one of them." | The check now compares `srcUid`, so several plans out of one workbook are not a clash. |
| 2026-08-14 audit | The root master workbook still had the pre-Department layout and no Plan ID / Last updated, so it loaded on a different lettering from every other file. | Migrated in place (§1): Department inserted at G on all five input sheets, row-5 fields added, tracker formula, validations, conditional formatting, column widths, `L3 time plan` and all five `master time plan` blocks repointed. Verified by loading it into the dashboard: 5 plans × 10 activities, Department/Status/Dependencies detected at G/N/P, dependencies resolved, no warnings. |
| 2026-08-14 audit | Doc drift: "12 demo timeplans" (six ship), `dist/06_Master_Time_Plan_Dashboard` (it is 05), two logo files in `01_Documentation` (one), `APP_VERSION 'Version 2.8'` (pinned to `'Version 1'`), the §2.2 column table and §2.3/§2.4/§2.5 cell references still on the pre-Department lettering, the 30-blank-row scan (80), markers "six colours" (five). | All corrected in place; §2.5's fifty `MOD(ROW())` bar rules marked superseded. |


---

## 2026-08-15 audit and stress test

A synthetic folder was generated to load the dashboard the way a real programme changes shape:
ten workbooks across levels L1, L2, L3 (×2), L4 (×2), L5, L6, L9 and one file with no level prefix;
activity counts of 0, 1, 2, 3 and 250; a six-row blank gap inside a table; a duplicate No.; a missing
Plan ID; a missing Last updated; an unresolvable dependency; a two-activity dependency cycle; and one
activity running to 2049 (a 1,200-week range). All of it loaded — 9 plans, 262 activities — with no
console error, and every page (master, each plan tab, Milestones, Markers, Risks, Export) rendered.
Four real defects came out of it.

| Found | Defect | Fix |
|---|---|---|
| 2026-08-15 audit | **A blank Plan ID cell made the words "Last updated" the plan's key.** `detectLabeledRef` took "the first non-empty cell within four columns to the right of the label", so with H5 empty the scan ran on to J5 and returned the *next field's label*. That string then became the plan key, the tab name, the dependency prefix, the Ref column and the baseline match key — and the "No Plan ID" warning never fired, because the plan looked like it had one. **Both shipped templates and every new file start with H5 empty, so this hit anyone who copied a template and did not fill the Plan ID in.** | `HEADER_FIELD_LABEL_RES` + `isHeaderFieldLabel()` list every label the header block can carry; the value scan now stops when it meets one and returns the empty cell next to the label instead. The write-back target is unchanged. |
| 2026-08-15 audit | **L6–L9 files drew in L1–L4's colours.** `levelColor` is `LEVEL_COLORS[(level-1) % 5]`, so an L6 tab and gutter were pixel-identical to L1's. The palette is deliberately five colours, so the hue has to repeat — the level did not. | Each further cycle past L5 is mixed 26% toward white (on top of the existing per-duplicate lightening), so L6 reads as a lighter L1 rather than as an L1. |
| 2026-08-15 audit | **An unlevelled file was given level 9**, the same rank as a real `L9_*.xlsx`, so the two interleaved in the stack. | Unlevelled plans are level 99. `levelKey` stays `L?` and the View panel still groups them as *Unlevelled*. |
| 2026-08-15 audit | **A file listed in `plans.json` that no longer fetches was dropped in silence** — a renamed or deleted timeplan simply stopped appearing, with nothing anywhere to say so. | `autoLoadManifest` records every miss in `this._manifestMissing`; `loadFileEntries` turns each into a File-checks row naming the file and telling the owner to update `plans.json`. |
| 2026-08-15 audit | A cell marker whose stored `planId` is no longer loaded showed the **first** timeplan in its dropdown next to an activity from a different file, with "nothing is drawn yet". | `mkRows` falls back to the plan that owns the activity code, so both dropdowns describe the same row and the marker draws. |
| 2026-08-15 audit | Wording drift now that `parsePlanFileName` accepts `L[1-9]`: the empty-state naming card, the header subtitle, the Master Plan tab tooltip and the unlevelled-file warning all still said L1–L5. The 700-week note said what the limit was but not what happens at it. | All say L1–L9 (or "in level order"); the 700-week note now states that weeks past the 700th are not drawn. |
| 2026-08-15 audit | The shipped demo baseline `_baselines/2026-07.json` still held all **twelve** plans of the old demo set — six of its entries pointed at files that no longer ship. | Trimmed to the six shipped plans. |

**Verified unchanged by the stress set:** 250 activities in one file render, warn once about
`TEMPLATE_ROWS` and cost nothing in the grid; a 0-activity plan gets a tab and an empty table rather
than an error; a 6-row blank gap inside a table does not truncate the scan (the 80-blank rule);
duplicate No. within a file is an error and the second row still loads; a dependency cycle loads with
an amber error; an unresolved ref draws as a red "(not found)" chip; a file whose name does not start
with `L{level}` is reported in File checks and not loaded (`strictFileNames`); Risks counted 15
delayed causes, 5 dependency risks and 18 file checks across the set.

**Known and accepted:** the grid draws at most 700 weeks (~13 years) and anything past that is not
drawn — the banner says so; the Excel-side `master time plan` sheet in the legacy workbook still only
has five 100-row blocks, so L6+ and >100 rows are dashboard-only (the per-plan files are unaffected);
`plans.json` still has to be edited by hand when the set of files changes, because a browser cannot
list a folder over http.


---

## 2026-08-15 alignment fix and full re-audit

| Found | Defect | Fix |
|---|---|---|
| 2026-08-15 | **The master grid's column headings sat 1–2px left of the rows below them**, so every dashed vertical divider in the heading row missed the one in the body, the error growing left to right (Status and Tracker were the worst). The sticky left column carried `border-right:2px`, which under the global `box-sizing:border-box` took 2px *out of* its content box. The body cells are `flex:none` and simply overflowed by 2px; the heading cells had no `flex` declaration, so flex-shrink distributed those 2px across them proportionally. | The 2px divider is now `box-shadow:2px 0 0 #D8D4CB` — drawn outside the content box, so the content box is exactly `leftTotal` and matches `gridWidth = leftTotal + timelineWidthPx` (verified: 632 + 1586 = 2218, no overflow). Every heading cell is explicitly `flex:none`. **Do not put a border back on that element** — use a box-shadow for any divider on a column whose width is arithmetic the rest of the layout depends on. |
| 2026-08-15 | A hidden column still drew a 1px dashed divider in the heading row (`width:0` + `border-right:1px` = 1px under border-box), which the body row did not, so hiding Location pushed the heading half a pixel out and left a stray line. | Responsible / Department / Location heading cells moved out of the template into `leftHdrResponsibleStyle` / `leftHdrDepartmentStyle` / `leftHdrLocationStyle`, built from `leftHdrBase` like Status and Tracker already were, each dropping its `borderRight` when its width is 0. |
| 2026-08-15 | The Activity heading sat at `padding-left:6px` where the body cell uses `0 8px`, a 2px text offset. | Heading padding is `0 8px 6px`. |
| 2026-08-15 | **View ▸ Columns had no stated scope.** It filters the master grid and `buildExportSVG` only — the timeplan tabs (`levelDetail.activities`, own style objects) and `buildStatusSheets` always render every column. Hiding Location and then seeing it on a tab reads as a bug. | Left as designed (a tab is the full record of one file) and documented: a note under the Columns heading in the View panel, and the same sentence in `README.md`. |

**Re-audited this turn, unchanged:** all five templates (2 sheets, `Colour` at C6, `Department` at
G6, `Timeplan Name` at B5, Plan ID / Last updated on row 5, self-contained Tracker formula, 4 data
validations, 3 conditional-formatting blocks, no `wrapText`, 10 zip parts each); all six shipped
timeplans (same shape, Plan ID present, 5 activities each, 0 unresolved dependencies, 34 `x14`
Gantt rules, `fullCalcOnLoad`); `plans.json` lists exactly the six files present; the demo baseline
`_baselines/2026-07.json` holds exactly those six plan ids; the root/pack master workbook (7 sheets,
all five input sheets migrated, Plan IDs `L1_Master`…`L5_Master`, `Last updated` 2026-08-14, 34 and
82 `x14` rules on the two Gantt sheets). Every page and panel rendered (Master, Milestones,
Markers, all six tabs, Risks × 3 tabs, Baseline, View, Files, Export, Save); an injected edit
produced the right Save panel entry (file, archive path, `end → Wk 41 2026`, `status → Risk`), a
Delayed cause and a red row; `buildExportSVG` (46.8 KB, logo + Department + Tracker present),
`buildImpactReport` and `buildStatusSheets` (6 pages) all build.

**Left as data, not defects:** `L4_Quality.xlsx` is stamped `2026-07-28` and so trips the
18-day freshness warning — that is the demo of the staleness check, and every file goes stale as
the calendar moves. The demo set also ships 2 delayed activities and 7 dependency risks on
purpose, so the Risks panel has something to show.


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


---

## 0.11 Brand logo (v2.9, 2026-08-08)

The heading strip and every export now carry a logo. `APP_VERSION` is `'Version 2.9'`.

**Precedence** (`initLogo()`, run from `componentDidMount`):

1. `idbGet('brandLogo')` — an image the user picked with **Files -> Branding -> Replace logo...**,
   stored as `{src (data URL), name, source:'picked', ratio}`.
2. The first URL in `LOGO_FOLDER_CANDIDATES` that `fetch` resolves —
   `../01_Documentation/logo.svg`, `.../logo.png`, then flatter fallbacks. This is the "swap the
   file, everyone gets the new logo" route. **It only works when the pack is served over http(s);**
   `file://` blocks the fetch, which is why step 3 exists and is the normal offline case.
3. `LOGO_BUILTIN` — the Volvo word mark inlined as an SVG data URL in the source.

`state.logoOn` (persisted inverted under IDB key `brandLogoOff`) hides it everywhere without
discarding it. `onResetLogo` deletes `brandLogo` and re-runs `initLogo`, so "Use default" falls
back to the folder file before the built-in.

`logoForExport()` is the single accessor every non-DOM consumer uses — it returns `null` when the
logo is off, so a hidden logo can never leak into an export:

- `buildExportSVG` draws an `<image href="dataURL">` right-aligned in the 54px `infoH` band,
  24px tall, width `24 * ratio` capped at 200px. Data URLs keep the SVG self-contained and keep the
  PNG rasterisation untainted.
- `reportHeader()` (impact report, status sheets) absolutely-positions a 24px `<img>` at the top
  right and reserves `padding-right:190px` on `.hd`.

`ratio` is measured once on load with `imageRatio()` (an `Image` probe, falling back to the
built-in 524.06/45.51) so an arbitrary user logo is never squashed.

The **Edit heading** button was shortened to **Edit** to make room; the logo sits to its right behind
a divider, at the far right of the heading strip.

`dist/01_Documentation/` ships exactly one logo file, `logo.svg` — the only name
`LOGO_FOLDER_CANDIDATES` looks for (see §0.13).


---

## 0.12 Milestones, PDF sizes, print footer, black accent (v2.9, 2026-08-08)

### Milestones

Dashboard-owned markers on the week axis. They are NOT read from or written to any workbook —
they live in `localStorage` under `tpDashboardMilestonesV1` as
`[{id, shape, color, label, wk, yr, on}]` (`loadMilestones` / `saveMilestones`;
`writeMilestones(list)` is the single write path and always does both storage + setState).

- `MS_SHAPES` — diamond (key deliverable), triangle (event deadline / review), circle (minor
  check-in), square (major start or closure gate). Add one here and it appears in the picker, the
  legend and both renderers automatically.
- `MS_COLORS` — five fixed swatches. Not a free colour picker on purpose.
- `msBoxStyle(shape, color, size)` returns a React style object; `msSvgShape(...)` returns the
  equivalent SVG node string. **Both must agree** — the on-screen marker and the exported marker are
  drawn by different code from the same record.
- `activeMilestones()` is the only thing renderers call: switched on, week 1–53, year > 1000.

New page `state.page === 'milestones'` (nav tab between Master Plan and the plan tabs, labelled
with the count). One row per milestone: shape buttons, colour swatches, name, week, year, Show
checkbox, delete. `msRows` in `renderVals` builds it all, including the per-option click handlers.

**Master plan rendering.** `weekHeaderH` is **74** when any milestone is active, **54** otherwise
— the extra 20px is the milestone band between the month row (top 19) and the week numbers.
The template's three former literal `54px` values (left header block, week header, arrow-overlay
`top`) are now `{{ weekHeaderH }}`; **if you add a fourth consumer of the header height, use the
same value.** `msMarkers` positions the **name above the symbol** (black `#1F2430`, 9px) centred on the week column; `msLines` draws a
1.5px line from `weekHeaderH` down `msContentH` = `sum(section.activities.length*28) +
(sections.length-1)*3` — the same accumulator §4.2 warns about, so a row-height change means three
edits now, not two. Milestones outside the Snapshot window are counted into `msOffAxisNote`
rather than silently dropped.

**Export rendering.** `buildExportSVG` sets `headerH = msOn ? 72 : 52` and derives
`weekTextY = infoH + headerH - 6` (the old hardcoded `infoH+46` in two places). Markers draw at `infoH+52` with the label baseline at `infoH+40` (name above symbol, black),
lines from `topH` to `height`, both pushed after the bars and before the
current-week rect so they read on top.

### PDF page sizes and footer

`Component.PDF_SIZES` — `letter` (11×8.5in), `a3` (16.54×11.69), `a2` (23.39×16.54), all
landscape, all in inches; `css` is the `@page size` value. `onExportPDF(sizeKey)` takes the key
(defaulting to letter) and reserves 22px at the foot of the sheet for the footer, so the artwork
scale is computed against `availH - footH`. The Export panel renders the three as a row from
`pdfSizes`.

`footerLine()` returns the footer parts: the `pdfFooterNote` prop (default
"Confidential — internal use only"), programme title, code, revision, date. `reportFooter(n, total)`
wraps them in `.pfoot` for the print reports — one per `.page` in `buildStatusSheets` (with real
page numbers), one at the end of `reportShell` for the impact report.

### Accent

`accentColor` defaults to `#000000` and the two `accent==='#000000' ? '#1F3864' : accent`
substitutions (heading stripe on screen, stripe in the SVG export) are **gone** — black now means
black. `reportShell`'s `.hd` border and the Export panel's PDF buttons read the prop too.

### Empty state

The no-file card leads with the logo at 40px above a divider (`logoEmptyStyle`, same
background-image technique as §0.11).


### Verified in exports (2026-08-08)

The SVG export was captured and inspected: milestone symbols, their black labels and their full-height
colour lines all render, together with the brand logo in the top-right of the header band. PNG is a
canvas raster of that same SVG and the PDF print window embeds the same string, so all three share one
renderer — a milestone that appears on screen appears in every export.


### Milestone line stacking (fix, 2026-08-08)

`msLines` sits at `zIndex: 30`, **below** the sticky week header's `zIndex: 40`. At 44 it drew
over the header while the grid scrolled, so the line appeared to run up through the month and year
bands. Do not raise it above 39. (The current-week marker is deliberately different — it is
`zIndex: 45` with `top:0;bottom:0` so it *does* span the header.)


---

## 0.13 Version pill (2026-08-08)

`APP_VERSION` is pinned to the string `'Version 1'` and is not to be bumped. Alongside it:
`APP_PUBLISHED` (`'13 Aug 2026'`),
`APP_AUTHOR`, `APP_AUTHOR_EMAIL` and `APP_VERSION_TOOLTIP` (a newline-joined native `title`
fallback). The pill in the top bar is `position:relative` and opens a hover card
(`state.verHover`, set by `onVerEnter`/`onVerLeave`) listing version, publish date, author and
email. `APP_VERSION` is still what the report headers print, so changing it changes both.

`dist/01_Documentation/` holds exactly one logo file, `logo.svg` — the duplicate copy under the
original Volvo file name was removed. That name is the only one `LOGO_FOLDER_CANDIDATES` looks for.

---

## 0.14 Session persistence (v2.8, 2026-08-10)

Everything the user authors **on** the dashboard rather than in a workbook now survives a close.
Before this, only the milestones (`tpDashboardMilestonesV1`) and the heading
(`tpDashboardHeadingV1`) persisted; timeplan tick state, column choices, the crop and the open
page were rebuilt from scratch on every load.

### The record

`SESSION_KEY = 'tpDashboardSessionV1'`, `SESSION_FILE = 'Dashboard_Settings.json'`,
`SESSION_VERSION = 1`. Helpers `readSession()` / `writeSession(rec)` / `sessionBody(rec)` /
`sessionStamp(iso)` sit next to `saveHeading`. The record:

```
{v, savedAt, hdr, milestones, columnVisible, compactTimeline, logoOn, logoHeight,
 healthSev, healthPlan, showBaseline, baselineId,
 page, crop, visible, planNameVisible, activityVisibility}
```

The last five are **plan-scoped**. `sessionRecord(prev)` only refreshes them while
`fileLoaded && plans.length`; with an empty dashboard the previous values are copied through
untouched, so opening the page before loading files can never wipe them.

`hdr` and `milestones` are duplicated here on purpose — their own two keys stay the single source
for the existing `loadHeading`/`loadMilestones` paths, and `sessionPatch` writes back through
`saveHeading`/`saveMilestones` so the three keys cannot drift.

### Write path

`componentDidUpdate` debounces `persistSession` by 350 ms; `componentWillUnmount` flushes it.
`persistSession` compares `sessionBody(rec)` — the record **minus `savedAt`** — against
`this._sessionBody` and skips identical writes, so a hover re-render does not touch storage and
`savedAt` does not creep. It never calls `setState`, so there is no update loop.

### Read path

1. `componentDidMount` applies `sessionPatch(rec, null)` **synchronously, before** the async
   IndexedDB restore, so the defaults never flash.
2. `loadFileEntries` re-reads the record after parsing and lays it over the freshly computed
   defaults: `{...visible, ...patch.visible}`, likewise `planNameVisible`.

`sessionPatch(rec, plans)` validates every field by type and filters plan-scoped keys to ids that
exist in the current set (`activityVisibility` keys split on `|` first, per §0.2). A restored
`crop` is only applied when it **overlaps** `computeDefaultRange(plans)` — a window from another
programme would otherwise draw an empty grid. A restored `baselineId` is only applied when that
snapshot is actually present.

### The settings file

`onSaveSessionFile` writes the record as JSON into `loadDirHandle` (the timeplan folder) with the
same permission dance as `onTakeBaseline`, and falls back to `downloadBlob`. `readFolderSession(dir)`
is called from `triggerFolderPicker` and stashes any `Dashboard_Settings.json` on
`this._folderSession`; `loadFileEntries` prefers it over the local record **only when its
`savedAt` string is later** (ISO, so lexicographic compare is correct), then clears it. That is
what lets one person's setup follow the folder to another machine without ever clobbering newer
local work. `onLoadSessionFile` applies one by hand and rejects a baseline file (it has
`plans`/`activities`).

`onResetSession` is two-click like `onClearFile`, clears all three keys and restores the defaults;
plans and baselines are untouched.

### Deliberately not persisted

`edits` and `planNameOverrides`. Those are pending **workbook** writes, already surfaced by the
red Save to Excel count; restoring them silently would let a stale edit be written into a file
that has since changed on disk.

### UI

A *Dashboard settings* section at the foot of the Files panel (Save settings file · Load
settings… · Reset settings, plus `sessionTargetLabel` and `sessionStampLabel`). The panel gained
`max-height:calc(100vh - 130px); overflow:auto` — it is now ~674 px tall.

## Automatic folder loading (v2.9)

`autoLoadManifest()` runs before the IndexedDB cache in `componentDidMount`. It tries
`./03_Time_Plans/`, `../03_Time_Plans/` and `./dist/03_Time_Plans/`, fetching `plans.json`
(`{files:[...], baselines:[...]}`) and then each listed file, and hands the result to
`loadFileEntries`. Over `file://` every fetch throws, it returns false and the cached set
loads as before. The real folder therefore wins over the cache whenever it can be read.

The twelve example timeplans carry two dependency rows each — one link to another plan, one
inside the plan — instead of a dependency on nearly every activity.

## Department column, markers, up-triangle (v2.9)

The example folder holds six timeplans (L1, L2, L3, L4 × 2, L5) with `plans.json` listing them;
several files per level is still supported, the set is just kept small.

**Department** is a real Excel column, inserted as **G** between Responsible (F) and Location (H)
in all twelve timeplans and all five templates. Everything from the old G onward moved one column
right for rows 6+ only, so the helper block at P2:V5 stayed put: cell refs, the Tracker formula,
data validations, conditional formatting, column widths and all 1275 `'Lx input'!` references on
the Gantt sheet were shifted with it. `HEADER_ALIASES` gains department/dept/function/org, so the
column is found by header name wherever an owner moves it; `DEFAULT_LAYOUT` is the new lettering.
Department is parsed, rendered on the master grid and the plan tabs, editable in the double-click
popup, written back to Excel, included in the SVG/PNG export and the A4 status sheets, and has its
own checkbox under View ▸ Columns.

**Cell markers** (`MK_*`, `cellMarkers` state, `tpDashboardCellMarkersV1`) are symbols pinned to
one activity row at one week: `{planId, code, wk, yr, symbol, color, label, on}`. The Markers page
mirrors the Milestones page; `activeCellMarkers()` drops rows whose plan or activity is no longer
loaded. On render they are grouped into `cellMarkerIndex` by activity code and drawn inside that
row's timeline cell, name above the glyph. Symbols are text glyphs (with U+FE0E where a codepoint
would otherwise render as emoji) so they survive print and SVG export.

**Milestones** gain a `triangleUp` shape in `MS_SHAPES`, `msBoxStyle` and `msSvgShape`. The Name
field no longer carries a "FIG" placeholder.

## Tabs, marker drawing, fixed version name

Tab fills: `NAV_FIXED` holds the three page colours (master ink, milestones bronze, markers plum);
timeplan tabs use `plan.color`, i.e. `levelColor(level, dupIndex)`, so a level reads the same on the
tab and in the grid. Inactive tabs are the same hue mixed 42% into white; `readableOn(hex)` picks
black or white text by relative luminance. No counts in the tab labels.

Cell markers now draw glyph-first with the name to the right (`flexDirection:'row'`), 26px in the
DOM and 22px in the SVG export with a white `paint-order="stroke"` outline, at `zIndex:8` so they
stay legible over a bar. `APP_VERSION` is pinned to "Version 1" and is not to be bumped.

## Activity colour is an Excel dropdown (2026-08-13)

The colour column (the cell immediately left of **No.**, column `C` in the shipped files) used to
be a *painted* cell — the dashboard read its fill. It is now a **data-validation dropdown holding
a colour name**, and the name is the single source of truth for Excel and the dashboard alike.

**Palette** — five names. These five are the ONLY colours in the whole product: activity fills,
level/tab identity (`LEVEL_COLORS`), milestone symbols (`MS_COLORS`) and cell markers
(`MK_COLORS`) all draw from them. `PALETTE_5` is derived from `ACTIVITY_COLOURS`;
`snapToPalette()` pulls any stored older colour to the nearest of the five on load, so a legacy
milestone can always be re-picked. `LEVEL_COLORS` repeats the hexes literally because it is
declared above `ACTIVITY_COLOURS`. Defined in exactly two places; change both together:

| Name | Hex | Where |
|---|---|---|
| Red | `E74B3B` | `ACTIVITY_COLOURS` in the DC, `PAL` in the migration script |
| Yellow | `FFC000` | |
| Green | `70AD47` | |
| Blue | `4472C4` | |
| Purple | `9B59B6` | |
| *(unset)* | `98A2B3` | `COLOUR_UNSET` / `UNSET_HEX` |

**Per workbook (all five templates, all six timeplans and the root master):**

1. `xl/styles.xml` — seven `dxf`s appended (six palette fills with a readable bold font colour,
   plus the grey unset fill) and one `cellXf` appended: `fillId=0`, the colour column's original
   `borderId`, centred. Every colour cell now points at that one xf, so the cell carries **no**
   static fill; all colour comes from conditional formatting.
2. Input sheet — header cell labelled `Colour`; the column widened to 12 so a name fits; each data
   row's colour cell rewritten as `t="inlineStr"` carrying the name (seeded from the *nearest*
   palette colour to whatever fill the row used to have), blank rows left empty; a
   `conditionalFormatting` block on `C{first}:C{last}` with six `cellIs equal` rules (one per name)
   plus an `expression` rule `AND($C7="",$E7<>"")` for the grey unset state; one `dataValidation`
   `type="list"` with `"Red,Yellow,Green,Blue,Purple"`, an input prompt and an error alert.
   Priorities 200–206 (existing sheet rules use 1–10; ranges don't overlap). Orange was dropped in a
second pass: its `cfRule`, its bar rule and its validation entry were deleted and seeded cells
remapped to Yellow; its `dxf` is left in `styles.xml`, unreferenced.
3. Time-plan sheet — the **50 bar rules per block are gone**. Each `x14:conditionalFormatting`
   block that used `MOD(ROW()-7,50)=k` is rewritten as **6 rules**: five of
   `AND('L3 input'!$C7="Red", <the original guard + encoded-week range test verbatim>)` and one
   with `$C7=""` painting grey. Only the `MOD(...)` term was swapped, so the row anchoring, the
   sheet reference and the date columns are exactly as they were — which is why the root master's
   `master time plan` (five blocks at 100-row offsets) migrated with the same code path.

So colour no longer cycles by row position: two activities that pick Blue are both blue, and
colour becomes a workstream grouping the owner controls.

**Dashboard side** (`parseLevelSheet`): `colourFromName(g(colorCol+r))` wins; a legacy sheet with
a painted cell and no name still falls through to `getCellColor` and then `PALETTE_FALLBACK`. Each
activity keeps `colourName`. If **any** row in the sheet carries a name, rows without one are
forced to `COLOUR_UNSET` — otherwise a stale fill would outlive the migration. The dashboard has
no colour control and never writes the colour column; that is deliberate per the brief.

**Nav bar.** All three tool tabs are `NAV_INK` (`#1F2430`) with white text; timeplan tabs use
`plan.color` with `#1F2430` text. Inactive tabs mix 34% white when dark, 50% when coloured. The
first timeplan tab carries `sep:true`, which renders the "TIMEPLANS" rule before it in the
template. The master-plan level gutter and the SVG export gutter now take `readableOn(plan.color)`
instead of a hardcoded white — L3's yellow band needs black text.

The migration was done by a one-off script (zip walk + regex XML edit, the same machinery as
§3.1/§5). It is idempotent-unsafe: running it twice would append a second set of dxfs and a second
validation. If a new workbook has to be migrated, port the script from the git history of this
change rather than re-running it on an already-migrated file.


## 2026-08-15 — shared-folder deployment: settings guard and documentation

**Save Dashboard is a multi-user write, and now says so.** `Dashboard_Settings.json` is one file
per folder shared by the whole team; previously the last person to press the button silently
replaced everyone else's milestones, markers, heading, ticks and week window.

- `savedBy` added to the settings record, prompted once per browser and kept in
  `localStorage.tpDashboardSavedBy`. Excluded from `sessionBody()` alongside `savedAt`, so it
  cannot trigger a rewrite on every render.
- `onSaveSessionFile` re-reads the folder's existing settings file before writing. If its
  `savedAt` is newer than `state.sessionSavedAt`, a `confirm()` names the person and the time
  and asks before replacing. Declining writes nothing and points at Files ▸ Load view….
- `state.sessionSavedBy` threaded through the four places the record is applied (mount, folder
  load, Load view…, Reset view). The Files panel stamp reads "Saved by X — <time>" when a name
  is present, "Your view as of <time>" when it is not.
- Deliberately a guard, not a lock — see A.7.

**Documentation restructured.**

- `README.md` cut from ~18 KB to ~9 KB and rewritten for the plan owner, not the developer:
  ten numbered sections, open-it-in-three-steps first, a four-rule shared-folder section, a
  Google Drive section (mirror vs stream, per-folder permissions, sync conflicts, archive
  growth), and a closing section saying plainly that `plans.json` and the legacy
  `Time_Plan_L1-L5_Master.xlsx` can be ignored.
- `README-AI.md` given a Part A orientation layer: how to read the document, what the thing is,
  the five invariants, an end-to-end data-flow diagram, a state-location table, the file-skip
  rules, the settings-guard spec, an add-a-column checklist, and the synced-drive deployment
  notes. The existing reference and dated log are unchanged as Part B.

No change to parsing, rendering, write-back, baselines, risks or exports.
