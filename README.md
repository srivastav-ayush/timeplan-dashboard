# Master Time Plan Dashboard

Everyone keeps their own Excel timeplan in one shared folder. The dashboard reads them all
and draws one plan.

No install, no server, no account, no internet. One HTML file.

![The master plan](docs/dashboard-master.png)

---

## 1. Open the dashboard

1. Open `05_Master_Time_Plan_Dashboard/Time Plan Dashboard.html` in **Chrome or Edge**.
2. Press **Load folder**.
3. Pick the `03_Time_Plans` folder.

That is it. Every timeplan in the folder is drawn, L1 at the top, L5 at the bottom.

> **Google Drive users:** the dashboard cannot be opened from drive.google.com in the browser —
> Drive will only show you the code. Open it from the copy on your own machine, in the folder
> Drive for desktop keeps in sync. Make a desktop shortcut to it. See section 9.

## 2. What is in the folder

| | |
|---|---|
| `02_Templates/` | blank L1–L5 starter files. Copy one to start a new timeplan. |
| `03_Time_Plans/` | the live timeplans, one .xlsx each. **This is the folder you load.** |
| `04_Exports/` | somewhere to put the pictures and PDFs you export. |
| `05_Master_Time_Plan_Dashboard/` | the dashboard. Open this. |
| `01_Documentation/` | the full technical reference, for developers. |

## 3. Start your own timeplan

Copy the matching template into `03_Time_Plans` and rename it:

```
L3_TEMPLATE.xlsx   ->   L3_Tooling.xlsx
```

**The name matters.** `L1`…`L9` sets where the plan sits in the stack (L1 on top); the word
after it puts plans on the same level in order. Several plans can share a level.

Don't rename the sheets inside. Each file has two: **`L3 input`** — where you type — and
**`L3 time plan`** — that plan's own Gantt, drawn by formulas. Never type in the second one.

## 4. Filling in your file

Fill the header block at the top, then the activity table from row 7.

**On row 5, fill in Plan ID and Last updated.** Plan ID is set once and never changed —
dependencies keep working even if you rename the file.

The columns you type in: Colour · No. · Activity / Event · Responsible · Department ·
Location · Start Wk/Yr · End Wk/Yr · Status · Comments · Dependencies.

**Colour** is a dropdown — click the cell left of **No.** and pick Red, Yellow, Green, Blue or
Purple. The cell fills, and the bars for that activity take the same colour in Excel and in the
dashboard. No colour picked yet shows grey.

**Tracker is the one column nobody types in.** It is a formula. It reads:

| It says | when |
|---|---|
| **Completed** | Status is *Completed*. |
| **Delayed 21w** | the end week has passed and Status is not *Completed*. |
| **Due this week** | this week is the end week, and it is not Completed. |
| **On Track** | the end week is still ahead. |

**Dependencies** — write the file name, a hyphen, then the **No.** of the activity you are
waiting for. Separate several with `;`:

```
L1_Programme-14; L3_Tooling-7
```

A bare number means "in this same file". Every activity's exact reference is printed in the
**Ref** column of its own tab in the dashboard, ready to copy. Anything that does not match is
listed under **Risks ▸ File checks** — a typo is never silently dropped.

## 5. Rules for a shared folder

Four rules. They are the whole of it.

1. **You edit your own file. Nobody edits yours.** Everyone can see every plan in the
   dashboard; only the owner changes the dates in it. If you need someone else's activity to
   move, tell them — don't edit their file.
2. **Close Excel before you press Save to Excel.** An open workbook is locked by Excel: the
   save can fail, and the sync will not go through cleanly.
3. **Press Reload folder before you trust what you are looking at.** It re-reads every file.
   The **Age** column under **View** tells you how many days since each owner last saved.
4. **Keep the names right.** `L3_Tooling.xlsx`. Anything else is skipped and reported.

## 6. Reading the plan

| | |
|---|---|
| **Risks** | What is late, what that breaks, and what is wrong with the files. Cause on top, everything waiting on it underneath. |
| **Baseline** | Freeze today's dates, then see how far things have moved since. |
| **View** | Which timeplans to show, which columns, the week window, and Age. |
| **Export** | SVG, PNG and PDF of what you are looking at, plus printable A4 status sheets with a sign-off line. |
| **Milestones / Markers** | Programme markers across the week axis, or a symbol pinned to one cell. Neither is written into anyone's Excel file. |

Double-click a bar on the master plan, or a row on a timeplan tab, to change its dates.
Pending changes show in red until you save.

## 7. Saving

**Save to Excel** lists which timeplans changed and what changed in each, then writes only
those files. Every file is rebuilt from its own bytes, so an edit in `L2_Bodyshop` can never
touch `L1_Programme`. Formulas, colours and formatting all survive.

Before anything is overwritten, the old version is copied into
`03_Time_Plans/04_Archived_Timeplans/` with the date and time in its name. If that copy cannot
be made, the save stops.

If a file changed on disk since you loaded it, the save stops and asks you to **Reload folder**
first — you cannot overwrite somebody else's work by accident.

## 8. Your dashboard view, and Save Dashboard

The milestones, markers, heading, logo, ticked timeplans and week window are saved in your
browser as you work. Close the dashboard, open it again, it comes back as you left it.
Nothing is written into anyone's timeplan.

**Save Dashboard** writes that view into the timeplan folder as `Dashboard_Settings.json`, so
everyone who loads the folder gets it.

**There is only one of these files, and it belongs to the whole team.** Pressing Save Dashboard
replaces whatever is in there. If someone has saved theirs since you picked yours up, the
dashboard tells you who and when and asks before replacing it. The first time you press it,
it asks for your name so the next person knows whose view they are looking at.

In practice: **nominate one person to own the team view.** Everyone else can set up their own
however they like — it stays in their browser and is never written to the folder unless they
press the button.

**Files ▸ Reset view** clears yours and goes back to the defaults. Timeplans, baselines and
archives are never touched by any of this.

## 9. Running this on Google Drive

This works well, with two settings to get right.

**Drive for desktop must be in "Mirror files" mode, not "Stream files".** Mirror keeps real
files on your disk, so the dashboard reads them instantly and works with no internet. Streaming
leaves placeholders that have to be fetched one at a time, which is slow and fails offline.
Drive for desktop ▸ Settings ▸ your Drive ▸ **Mirror files**. If you must stay on streaming,
right-click the top folder and tick **Available offline**.

**Set the permissions per folder:** edit on `03_Time_Plans`, view-only on `02_Templates` and
`05_Master_Time_Plan_Dashboard`, so nobody edits a template or deletes the dashboard.

Two things worth knowing:

- **Sync conflicts.** If two people write the same file, Drive keeps both and names one
  something like `L2_Bodyshop (1).xlsx`. The dashboard does not load those — it lists them
  under **Risks ▸ File checks** so you can see it happened and delete the loser. Rule 1 in
  section 5 is what stops it.
- **The archive folder syncs too.** Every save uploads a copy of the old version to everyone.
  If the folder gets heavy, either delete old archives now and then, or use **Archive
  elsewhere…** in the Save panel to point them at a folder on your own machine instead.

## 10. Two things you can ignore

- **`plans.json`** in `03_Time_Plans` is a list of the timeplan files. It is only used when
  the dashboard is opened from a web address (`http://…`) instead of from a folder — then it
  can load the files without you pressing anything. **On Google Drive, or any normal folder,
  it is never read and you never need to touch it.** It does no harm if it goes out of date.
- **`Time_Plan_L1-L5_Master.xlsx`** is the older all-in-one workbook, all five levels in one
  file. It is kept for reference. It is **not** the dashboard and not something you fill in.

## Limits

- **Chrome or Edge** for loading and saving folders. Firefox and Safari can load files, and
  will download saved files instead of writing them back.
- The templates are set up for **100 activities per timeplan** (rows 7–106). Past that the
  dashboard still reads them, but Risks ▸ File checks tells you to copy the last row down or
  the dropdowns, Tracker and Excel Gantt stop at row 106.
- The grid draws up to **700 weeks**. Narrow the week window under **View** if you hit it.

## Privacy

Nothing is uploaded anywhere. No server, no analytics, no network call of any kind. All of it
happens in your browser, and it works fully offline.

---

For developers: `01_Documentation/README-AI.md` is the full technical reference — data model,
the .xlsx parse and rebuild path, dependency resolution, the baseline format, and the revision
log.
