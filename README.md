# Master Time Plan Dashboard

**Every owner keeps their own Excel timeplan. This reads all of them and draws one plan.**

An offline HTML dashboard that reads one Excel file per timeplan out of a shared folder and
stacks them into a single master Gantt — L1 at the top, L5 at the bottom. It writes changes
back into the original workbooks, keeps a dated copy of every version it replaces, and tells
you what has to move when something slips.

No install, no server, no account, no internet. One HTML file.

![The master plan — every timeplan stacked into one Gantt](docs/dashboard-master.png)

---

## Why

Programme plans get split across levels and owners: a programme master plan, a plan per
workstream, a plan per supplier. Each owner works in their own Excel file, which is right —
they own their dates. What nobody has is the whole picture, and by the time it is copied into
one deck by hand it is a week old and wrong.

This dashboard leaves every owner's file exactly where it is. It reads them all, draws them
together, and reports the four things that are hard to see across separate files:

- what is stale (nobody has touched L4_Logistics for 49 days)
- what is broken (a dependency pointing at an activity that does not exist)
- what has moved since the last review
- what else has to move because of it

## Quick start

1. Download the folder, or clone this repository.
2. Open `05_Master_Time_Plan_Dashboard/Time Plan Dashboard.html` in Chrome or Edge.
3. Press **Load folder** and pick `03_Time_Plans`.

If the folder is served over http (a published shared drive, or a local web server) step 3 is
skipped: the dashboard reads `03_Time_Plans/plans.json` and loads every file listed there by
itself. Keep that list in step with the folder when you add or remove a timeplan. Opened
straight from disk, the browser blocks folder reads, so **Load folder** stays the way in.

Six example timeplans are included (one L1, one L2, one L3, two L4, one L5) so it draws something
straight away. Delete them when you start your own. Any number of files can share a level — the
example set is deliberately small, not a limit.

## What is in the box

```
README.md                      this file
01_Documentation/              README-AI.md — the technical reference for the whole app
02_Templates/                  blank L1–L5 starter files — copy one to create a new timeplan
03_Time_Plans/                 the live timeplans, one .xlsx per plan — this is what you load
                               plans.json — the list used for automatic loading over http
04_Exports/                    somewhere to put SVG / PNG / PDF exports
05_Master_Time_Plan_Dashboard/ the dashboard itself — open this
docs/                          the screenshots this README shows
Time_Plan_L1-L5_Master.xlsx    the older all-in-one workbook (five input sheets in one file),
                               kept as a reference and as something to split new plans from
```

Two folders appear inside `03_Time_Plans` as you use it: `_baselines/` (dated snapshots for the
Baseline view) and `04_Archived_Timeplans/` (the version each file replaced, dated). A
`Dashboard_Settings.json` appears there too if you ask for one — see *Your settings are kept*.
Nothing in the code hardcodes any of these names — rename them freely.

## One file per timeplan

Copy the matching file out of `02_Templates` into `03_Time_Plans` and rename it:
`L3_TEMPLATE.xlsx` → `L3_Tooling.xlsx`. Don't rename the sheets inside.

**The file name sets the level and the order.** `L1` … `L9` puts the plan in the stack (L1 on
top); the word after it orders plans that share a level. Several files may share a level, and there is
no limit on how many files you load. L1–L5 is the usual set; L6–L9 work the same way and reuse the
five colours a shade lighter.

Each file has exactly two sheets: **`L3 input`** (where the owner types — the only sheet the
dashboard reads or writes) and **`L3 time plan`** (that plan's own Gantt in Excel, drawn by
formulas — don't type in it).

On the input sheet fill in the header block (Project Name, Code, SoP, Description, Timeplan
Name, **Plan ID**, **Last updated**) and then the activity table from row 7: Colour · No. ·
Activity / Event · Responsible · Department · Location · Start Wk/Yr · End Wk/Yr · Tracker ·
Status · Comments · Dependencies.

**Colour** is a dropdown. Click the cell left of **No.** and pick one of five colour names —
Red, Yellow, Green, Blue, Purple. The
cell fills with that colour, the bars for that activity on the `L3 time plan` sheet take the same
colour, and so do its bars in the HTML dashboard. Excel owns the colour: there is no colour
picker in the dashboard, it reads whatever the workbook says. An activity with no colour picked
yet shows grey in both views.

The buttons along the top sit in four groups: getting files in and pictures out
(**Files**, **Export**), writing your work back out (**Save Dashboard**, **Save to Excel**),
what the dashboard shows (**View**), and reading the plan (**Risks**, **Baseline**). They stay
white until you hover over them, so colour on that row always means something. Two of them
colour up on their own: **Save to Excel** turns red with a count as soon as you edit a date, and
**Save Dashboard** turns green with a count as soon as your view differs from the settings file
in the folder. Hover either one and it tells you exactly what is waiting.

**Risks** lists what is in trouble and why. Each shaded row is something that is late or has
moved; the activities that wait on it are listed underneath. Two of the columns are easy to mix
up, so read them like this: **Late by / gap was** is how many weeks late the shaded activity is,
and on the rows beneath it, how much room the plan originally left between the two. **Must move**
is how many weeks that activity has to shift to start after the shaded one actually finishes.
**Print list** gives you the same thing on paper.

**Tracker** is not typed by anyone — it is a formula, and it is the one column nobody should
touch. It answers "is this late?" from two things only: whether **Status** says *Completed*, and
where the activity's **End Wk/Yr** sits relative to the current week.

| Tracker says | when |
|---|---|
| **Completed** | Status is *Completed*. Early or late, the work is done. |
| **Delayed 21w** | the end week has passed and Status does **not** say *Completed*. The number is how many weeks past due. |
| **Due this week** | the current week *is* the end week, and it is not Completed yet. |
| **On Track** | the end week is still ahead. |
| *(blank)* | no activity name, or no end week to judge against. |

So an activity only reads **On Track** while its deadline is still in the future, or once it is
marked *Completed*. The moment the end week passes without *Completed*, it reads **Delayed** and
says by how much. The same rule runs in the HTML dashboard, in its own **Tracker** column
(switch it off under **View ▸ Columns**) and on the printed status sheets, so the sheet and the
dashboard can never disagree.

The formula is self-contained in each cell and compares real dates, so it keeps working across
the New Year and cannot be broken by editing anything else on the sheet.

**Department** sits next to Responsible and says which function that owner belongs to
(Body Engineering, Quality Assurance, Supplier Quality). It is read like any other column and
shown on the master plan and on each timeplan tab. **View ▸ Columns** switches it — or any other
column — off on the master plan and in the SVG, PNG and PDF exports; a timeplan tab and the
printed status sheets always show every column, because they are the full record of one file.

Those two sit to the right of the Timeplan Name on row 5, so all of row 5 is one line: name,
Plan ID, Last updated.

**Plan ID** is set once and never changed — dependencies keep working even if the file is
renamed. Fill it in on every new file: leave it empty and the dashboard falls back to the file name,
and Risks ▸ File checks says so. **Last updated** is the freshness stamp; the dashboard flags plans older than two
weeks and stamps it for you when it saves.

<!-- Drop a screenshot of one of the Excel input sheets in docs/excel-input.png and uncomment:
![The Excel input sheet each owner fills in](docs/excel-input.png)
-->

## Dependencies

In the Dependencies column write the **file name**, a hyphen, then the **No.** of the activity
this one waits for:

```
L1_Programme-14; L3_Tooling-7
```

Separate several with `;`. The activity name works instead of the No.
(`L1_Programme-Gate 2 review`), and a bare No. or name means "in this same file".

Use the No. — it is matched first and can never be ambiguous. Every activity's exact reference
is printed in the **Ref** column of its own tab, ready to copy. Anything that doesn't match is
listed under Risks ▸ File checks and drawn as a red chip; a typo is never silently dropped.

![A single timeplan tab, with the copy-ready Ref column](docs/dashboard-timeplan.png)

## What it does

| | |
|---|---|
| **Files** | Where the timeplans come from, **Reload folder** after an owner updates their file in Excel, where saves go, your dashboard view, and clear. |
| **Milestones** | Programme markers on the week axis — symbol, colour, name, week — drawn across every timeplan. They belong to the dashboard, not to any file. |
| **Markers** | A symbol pinned to a single cell: pick the timeplan, then the activity, then the week and year. Twelve symbols (star, hexagonal star, smiley, flag, tick, cross and more), five colours, and the name you type sits beside the symbol. The glyph fills the cell and stays readable where it crosses an activity bar. Nothing is written into the Excel files. |
| **Save to Excel (n)** | Lists exactly which timeplans changed and what changed in each, then writes them. Only changed files are rewritten. |
| **Export** | SVG, PNG and PDF of the current view, plus printable A4 status sheets — one page per timeplan, with a sign-off line. |
| **Risks** | What is in trouble right now, under three tabs: **Delayed** (past its end week, or moved), **Dependency risk** (a link that cannot work — the wrong way round, or waiting on something unfinished) and **File checks** (stale files, dates, names, IDs, dependency links). Green and reading *No risks* when there is nothing. |
| **Edit** (on the heading strip) | The programme name, code, SoP, revision, issued-by, the one-line description — and the logo. Everything shown here and on every export. |
| **Baseline** | A photo of today's dates — take one, then see how far things have moved since. |
| **View** | Which timeplans are shown, renaming, **Age** (days since the owner last saved that file), **Acts** (activity count), which columns to show, the compact timeline, and a week window that crops the view and every export. |
| **Save Dashboard** | Writes your dashboard view — milestones, markers, heading, logo, tick boxes, week window — into the timeplan folder, so anyone who loads that folder gets it automatically. |

The tabs are colour-coded: Master Plan, Milestones and Markers are black with white text — they are
dashboard tools, not files — and a small **Timeplans** rule separates them from the file tabs. Each
timeplan tab carries its level colour from the same five-colour palette (L1 blue, L2 green, L3 yellow,
L4 purple, L5 red) — every L1 the same, every L2 the same, and so on — so the tab
matches the band down the left of the master plan. Hover the **timeplans** pill to list the loaded
file names, or the **activities** pill for the activity count per timeplan.

Double-click a bar or its colour swatch on the master plan, or a row on any timeplan tab, to
change its start/end week and, optionally, its status, name, responsible, department or location.
Pending changes show in red until you save.

### Your settings are kept

Everything you set up on the dashboard rather than in Excel — the milestones, the markers, the heading
fields and logo (both under **Edit**, on the heading strip beside the logo), which timeplans and
columns are ticked, the week window and the tab you were on — is saved in the browser as you work. Close the dashboard, open it again, and it comes
back exactly as you left it. Nothing is written into the timeplans and nothing leaves your
machine.

That covers one person on one computer. To move a setup — to another machine, or to the rest of
the team — press **Save Dashboard** in the top bar. It writes a small `Dashboard_Settings.json` into
the timeplan folder (or downloads it if the browser cannot write there). Anyone who loads that
folder picks it up automatically when it is newer than what their own browser holds, and
**Files ▸ Load view…** applies one by hand from anywhere — you only need it if your view did not
come back on its own, or if you want to pick up someone else's.

**Files ▸ Reset view** clears the lot and goes back to the defaults. The timeplans, the baselines
and the archived copies are never touched by any of this.

Pending Excel edits are deliberately *not* part of it: a changed date is a change to a timeplan,
so it stays in the red *Save to Excel* count until you write it into the file.

### Baselines

**Capture baseline now** freezes the start and end week of every activity in every loaded file,
exactly as they are at that moment, and stores it as a dated snapshot. Nothing on the plan moves
and nothing is written into the timeplans.

Afterwards, switching **Baseline view** on shows the difference: a dashed outline where the
activity used to sit and a badge for how far it moved — `+3w` (red) is three weeks later, `−2w`
(green) is two weeks earlier. Every snapshot is kept and named by month, and the two dropdowns
set what is compared: from a capture, to today's live plan or to a second capture. So "what
moved between the June and the August review" is one question, not an afternoon.

Activities are matched on Plan ID first, file name second, then activity number, so renaming a
file or an activity does not break the comparison.

### Risks

One question: **something is late — what does that break?** The list is written that way, cause
above effect. Each late or moved activity is a heading, and the activities waiting on it are
indented underneath:

```
L2_Bodyshop-1  Bodyshop layout study     21w late · due W12 '26 · Aligned    1 waiting
  ↳ L3_Tooling-1  Tooling concept        W9 '26 – W12 '26         Move 25w later
```

Three things make an activity a cause:

- **It is late** — its end week has passed and nobody marked it *Completed*. No baseline and no
  editing needed; this is true the moment you open the folder.
- **You moved it** — an unsaved edit changed its end week.
- **It moved away from the baseline** — if you have one selected.

For each activity waiting on a cause you get one of:

- **Move Nw later** — it now starts on or before the thing it waits for finishes, so its dates
  are impossible as they stand.
- **Tight — Nw spare** — it still fits, but with two weeks of slack or less.

**Print list** / **Save list as file** produce a document with the owner of each affected
activity. Take it to the review.

The **Dependency risk** tab asks the same question from the other end: which links cannot work as
written. A pair of rows per case — the activity at risk, then the activity it waits on — for
anything that starts before what it depends on finishes, or that waits on something already past
its end week unfinished.

Nothing is changed automatically, and only the links people actually wrote in the Dependencies
column are followed — an activity nobody linked can never appear here.

![Risks: what is late, and what that breaks](docs/dashboard-knockon.png)

## Saving, and how your old versions are kept

**Save to Excel** writes one file per changed timeplan, rebuilt from that file's own bytes, so
an edit in `L2_Bodyshop` can never touch `L1_Programme`. Formulas, conditional formatting,
colours and sheet protection all survive.

Before anything is overwritten, the version being replaced is copied into
`03_Time_Plans/04_Archived_Timeplans/` as `2026-08-08_1432_L2_Bodyshop.xlsx` — date, time, then
the timeplan file name, so the copies sort oldest to newest. Only then is the new version
written. If the copy cannot be made, the save stops and the original is left alone.

The checkbox in the Save panel turns archiving off if you don't want the copies; **Archive
elsewhere…** points them at any other folder.

Every rebuilt file is re-parsed and verified before anything is written; if one value doesn't
round-trip, the whole save aborts. If any file changed on disk since you loaded it, the save
stops and asks you to **Reload folder** first — you can't overwrite somebody else's work by
accident.

## Limits

- **Chrome or Edge** for folder loading and folder saving. Firefox and Safari can load files
  and will download saved files instead.
- The templates are formatted and formula-filled for **100 activities per timeplan** — rows 7–106 on both sheets. Go
  past that and the dashboard still works — it reads up to 2000 rows — but Risks ▸ File checks tells you
  to copy the last row of that file's own *input* and *time plan* sheets down, or its dropdowns,
  Tracker column and Excel Gantt stop at row 106. Twelve files of 100 activities is 1200
  activities and is fine.
- The grid draws up to 700 weeks (~13 years). Past that the banner says so and the later weeks are
  not drawn — narrow the week window under **View**.
- Over http the dashboard loads what `plans.json` lists. A name in there that no longer exists is
  reported under Risks ▸ File checks rather than skipped in silence.

## Privacy

Nothing is uploaded anywhere. There is no server, no analytics, no network call of any kind.
All parsing, rendering and rebuilding happens in your browser, and the file works fully
offline — including from a USB stick or a network drive.

## For developers

`01_Documentation/README-AI.md` is the full technical reference: data model, the .xlsx parse
and rebuild path, dependency resolution, the baseline format, and every UI revision. The
dashboard is a single self-contained HTML file with no build step and no dependencies.
