# Time Plan Pack — read me

One Excel file per timeplan. One HTML dashboard that reads them all and draws a single Gantt.

## Folders

| Folder | What is in it |
|---|---|
| `02_Templates` | Empty L1–L5 workbooks. **Copy** one to start a new timeplan. |
| `03_Time_Plans` | The live timeplans. One `.xlsx` per plan, named `L{level}_Name.xlsx`. |
| `05_Master_Time_Plan_Dashboard` | `Time Plan Dashboard.html` — open in Chrome or Edge. |
| `04_Exports` | Where you save pictures and PDFs. |
| `01_Documentation` | Technical reference (you do not need it). |

## Using it

1. Copy a template out of `02_Templates` into `03_Time_Plans` and rename it `L2_MyPlan.xlsx`.
2. Open the **input** sheet. Fill in the header block, then one activity per row from row 8.
3. Save. Open the dashboard, click **Files ▸ Load folder** and pick `03_Time_Plans`.

## The input sheet

Header block (rows 2–5): Project Name, Project Code, SoP, Description, **Timeplan Name**,
**Timeplan ID**, **Timeplan Owner**. Last updated fills itself in — do not type a date there.

The **Timeplan range** box (R2:U3) sets the first and last week the Excel Gantt draws.
**Today** (W3:X3) fills itself in.

Activity table, rows 8 to 207 — 200 activities:

| Column | Fill in |
|---|---|
| Colour | Dropdown: Red / Yellow / Green / Blue / Purple. Colours the bar. |
| No. | Already numbered 1–200. Leave it. |
| Activity / Event | The name. **A row with no name is ignored everywhere.** |
| Responsible | Person. |
| Department, Location | Free text. |
| Start Wk / Yr, End Wk / Yr | ISO week numbers and years. |
| Status | Dropdown, 14 values. The one thing you choose. |
| Tracker | **Formula — do not type in it.** |
| Comments / Notes | Free text. |
| Dependencies | `L1_Programme-14; L3_Tooling-7` — see below. |

### Status

Completed · Cancelled · Skipped · Risk · Blocker · Dependency · Aligned · Planned ·
To be discussed · Not done · Deferred · Not Achieved · Pending · Not Completed

### Tracker (works itself out)

| | |
|---|---|
| Done | Status is Completed. |
| Closed | Status is Cancelled or Skipped — off the plan. |
| Incomplete | The end week has passed and it is not closed out. |
| Due this week | This week is the end week. |
| On Track | Still ahead of the end week. |

### Dependencies

Write the other file's name, a hyphen, then the activity's **No.**: `L1_Programme-14`.
Several: separate with `;`. A bare number means this same file. The exact reference for every
activity sits in the **Ref** column of its tab in the dashboard, ready to copy. Anything that
does not match is listed in the dashboard's file checks.

## ⚠ Rules for the Excel files

These workbooks are wired together by cell position. Breaking a position breaks the Gantt,
the Tracker and the dashboard — silently.

- **Do not insert or delete rows or columns.** Not in the input sheet, not in the time plan sheet.
- **Do not cut cells (Ctrl+X) and do not drag cells to move them.** Cutting takes the formulas
  that point at that cell with it. **Copy and paste, or type, only.**
- **Paste as values** (Paste Special ▸ Values) when pasting from another file. A normal paste
  brings the other file's formatting and can bring links to it.
- **Do not rename the sheets.** `L3 input` and `L3 time plan` are found by name.
- **Do not rename a file once its ID is in use** — every dependency pointing at it stops working.
- **Do not type in the Tracker column, or anywhere on the time plan sheet.** Both are formulas.
- **Do not touch the Timeplan range box or the Today box** beyond the four range numbers.
- **Do not use Sort.** It moves values out from under the formulas. Reorder by retyping.
- Only fill rows 8–207. Past row 207 the dropdowns and the Tracker formula stop.
- One file per timeplan, named `L{1-9}_Name.xlsx`, all in the same folder, no sub-folders.

Nothing is password protected any more — which is exactly why the rules above matter.

## The dashboard

Open `Time Plan Dashboard.html` from the folder on your machine (not from a preview in Drive or
SharePoint). **Files ▸ Load folder**, pick the timeplan folder, and it draws everything.
**Risks** lists what is late, what is waiting on something unfinished, and what is wrong in the
files. **Export** gives you SVG, PNG, PDF and one printable status sheet per plan.
**Save to Excel** writes date edits back into the original workbooks, keeping a dated copy of
what it replaced.
