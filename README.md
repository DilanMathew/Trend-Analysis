# Advanced Milestone Trend Analysis (Excel)

`Advanced_Milestone_Trend_Analysis.xlsx` is a macro-free Excel workbook for Milestone Trend Analysis (MTA). It tracks how forecast milestone completion dates move across successive reporting periods. Each milestone is drawn as a line: the X axis is the **reporting date** and the Y axis is the **forecast completion date**.

## Contents

| Sheet | Purpose |
|---|---|
| **Dashboard** | 8 KPI cards, the MTA chart with a milestone key and show/hide toggles, 5 supporting charts and 4 filters (Project, Reporting Period, Category, Status) |
| **Milestone Data** | `tblMilestoneData`: one row per milestone per reporting date, holding the full forecast history |
| **Project Setup** | `tblProjects`, the analysis thresholds (named cells) and the dropdown lists (`tblCategories`, `tblPriorities`, `tblDepartments`) |
| **Milestone Register** | `tblMilestones`: master list with a unique Milestone ID per milestone |
| **Reporting History** | Snapshot entry form with validation and a paste-ready block, plus a colour-coded forecast history matrix |
| **Delay Analysis** | Per-milestone baseline, first, previous, current and actual dates; variances; status; movement and risk flags. Includes a historical comparison panel |
| **Status Report** | One-project status report: summary, RAG status, performance, delays, latest changes and risks. Fits A4 portrait |
| **Instructions** | Data-entry guide and technical notes |
| **Engine** | Helper calculations behind the dropdowns and chart series. All formulas; do not edit |

Sample data: 3 projects with different reporting rhythms (bi-weekly, monthly, every 3 weeks), 40 milestones and 347 forecast observations. The data covers slips, pull-ins, stable forecasts, completed, early, late and not-yet-started milestones, a skipped report and a missing latest update.

## Recording a reporting cycle (no VBA)

1. **Reporting History**: pick the project and type the reporting date.
2. Type the new forecast for each milestone. If a milestone is complete, also type its actual date. Column K validates each row: ✔ ready, ⚠ warning, ✖ blocked. Duplicates are blocked.
3. Copy the green block `O10:S…`. Paste it as **values** into the first empty row under the table on **Milestone Data**. The table expands and every calculated column fills automatically.

You can also type new rows directly at the bottom of `tblMilestoneData`. Only 5 columns are inputs: Milestone ID, Reporting Date, Forecast, Actual and Comments. Dropdowns and validation block unknown IDs, invalid dates and duplicate milestone/date pairs.

The optional `vba/MTA_Snapshot.bas` (`PostSnapshot`) replaces the copy/paste step if you save the workbook as `.xlsm`. It is not required.

## Technical design

- **Tables and structured references.** All inputs live in Excel Tables. Lookup columns in `tblMilestoneData` are table *calculated columns*, so new rows fill themselves in. Dropdowns use names such as `lstMilestoneID = tblMilestones[Milestone ID]`, so they grow with the tables.
- **Observation key.** `Obs Key = Milestone ID & "|" & reporting-date serial`. A single INDEX/MATCH on this key returns the forecast for any milestone/date pair, or nothing. No values are interpolated or carried forward.
- **Previous forecast.** `MAXIFS(Reporting Date, Milestone = this, Reporting Date < this)`, then a `SUMIFS` on that date. This works for any reporting frequency and for skipped reports.
- **As-of logic.** Every Delay Analysis column counts only observations on or before the selected reporting period. `(Latest)` means each project's own latest report. Each column is exposed as a named range (`da_status`, `da_totdelay`…), which the Dashboard, Engine and Status Report read.
- **Status rules.** In order: Not Yet Started (no observation yet) → Completed (actual date recorded) → Delayed (baseline variance > `cfgTolerance`) → Ahead of Schedule (variance < −tolerance) → On Schedule. A completed milestone is never classed as Delayed.
- **MTA chart.** An XY scatter chart, so both axes are real date axes and the gaps between points reflect the actual reporting intervals. The Engine grid has one row per reporting date (sorted unique dates built with chained `MINIFS(date > previous)`) and one column per milestone slot. An empty cell returns `NA()`, which Excel does not plot, so a missing observation never appears as zero or as an invented date. Overlay series add red rings for significant moves (≥ `cfgSigMove`) and diamonds for achieved milestones. A dashed diagonal is the completion line.
- **Slots.** Delay Analysis numbers the milestones that pass the MTA filters (`MTA Slot`). Engine slot *k* picks milestone *k*, so the 15 chart series never need re-pointing.
- **Why INDEX/MATCH and MAXIFS rather than FILTER/SORT/UNIQUE/XLOOKUP.** Chart series and data-validation lists need fixed cell addresses, and spilled arrays cannot sit inside tables. Non-spilling formulas also work in Excel 2019+ and LibreOffice, which is what the automated tests use to recalculate. Nothing depends on volatile functions (OFFSET, INDIRECT, TODAY).
- **Protection.** Dashboard, Delay Analysis, Reporting History, Status Report and Engine are protected without a password, with input cells left unlocked. The data sheets stay unprotected because Excel cannot expand a table on a protected sheet.

Capacity: 200 milestones, 15 lines × 40 reporting periods per project on the MTA chart, 25 milestones per project on the entry form, 26 dates in the history matrix. Each limit is a block of copied formula rows and can be extended.

## Rebuilding and testing

```bash
python tools/build_workbook.py
```

```bash
python tools/qa_tests.py /path/to/recalc.py
```

`tools/qa_tests.py` recalculates scenario copies in LibreOffice and compares them with an independent Python model of the MTA rules. It covers:

- KPI and Delay Analysis accuracy under several filter combinations
- missing-observation handling
- hiding a line
- adding a milestone and a reporting period
- duplicate detection
- the entry form
