# MS Project → MTA workbook: automatic CSV import

`Advanced_Milestone_Trend_Analysis_MSProject.xlsx` is the MS Project edition of the workbook. Its Milestone Register and Milestone Data tables are fed by Power Query from a folder of MS Project CSV exports. Each export is one reporting snapshot, so the forecast history builds up as you add files.

```
\\Team\MTA\MSP Exports\
   PRJ-001_2026-11-16.csv      <- one file per project per reporting cycle
   PRJ-001_2026-11-30.csv
   PRJ-002_2026-11-09.csv
```

Each cycle: export from MS Project into the folder, open the workbook, then **Data → Refresh All**.

## 1. MS Project: create the export map (once)

1. **File → Save As**, choose **CSV (Comma delimited)**, and name the file `<ProjectID>_<yyyy-mm-dd>.csv`. The date is the reporting (status) date, for example `PRJ-001_2026-11-30.csv`.
2. In the Export Wizard, choose **Selected Data → New map → Tasks**, and tick **Export includes headers**.
3. Optional: set **Export filter = Milestones** to keep the file small. The query skips non-milestones anyway.
4. Add these fields. Their order and the destination names don't matter; the query recognises `Unique_ID`, `Unique ID` and so on.

   | Required | Recommended |
   |---|---|
   | Unique ID, Name, Milestone, Finish | Baseline Finish, Actual Finish, % Complete, Summary, Resource Names, Priority, Text1 (used as Category), Notes |

5. **Save Map** as `MTA Export`. Next cycle, pick that map and click Finish.

How each field is used:
- **Milestone ID** is `<ProjectID>-<Unique ID>`, for example `PRJ-001-108`. Unique ID never changes when tasks are re-ordered, so a milestone's history stays connected across cycles.
- **Forecast** is Finish. **Actual** is Actual Finish, or Finish when % Complete is 100.
- **Baseline** is Baseline Finish. If the project was never baselined, the first forecast is used instead.
- **Category** is Text1 (blank becomes "Milestone"). **Priority** is mapped from 0–1000: 800 and above Critical, 600 and above High, 400 and above Medium, otherwise Low.
- **Owner** is Resource Names, with units removed (`Tom Becker[50%]` becomes `Tom Becker`). **Comments** come from Notes.

## 2. Excel: install the queries (once, about 10 minutes)

1. Open `Advanced_Milestone_Trend_Analysis_MSProject.xlsx`.
2. For each block in `MSP_Import.pq`:
   - Choose **Data → Get Data → From Other Sources → Blank Query**.
   - Choose **Home → Advanced Editor**, replace the text with the block's code, and click OK.
   - Rename the query (Query Settings → Name) exactly as in the block header: `pMSPFolder`, `pDateCulture`, `pDelimiter`, `pEncoding`, `fnMSPDate`, `fnMSPHeader`, `MSP_Files`, `MSP_Tasks`, `MSP Observations`, `MSP Milestones`.
3. Set the parameter values:
   - `pMSPFolder` is your export folder.
   - `pDateCulture` is `en-GB` if MS Project writes 30/11/26, or `en-US` if it writes 11/30/26.
   - `pDelimiter` is `,` or `;`.
   - `pEncoding` is `1252` (MS Project's default) or `65001` for UTF-8.
4. Click **Home → Close & Load To… → Only Create Connection**.
5. In **Queries & Connections**, right-click **MSP Observations**, choose **Load To… → Table → Existing worksheet** and enter `='MSP Observations'!$A$1`. Do the same for **MSP Milestones** with `='MSP Milestones'!$A$1`.
6. Optional: open **Query Properties** and tick **Refresh data when opening the file**.
7. On Project Setup, make sure each Project ID used in the file names exists in `tblProjects`. Add any new Text1 categories to `tblCategories`.

Don't rename the two output sheets, and don't change the query's column order. The workbook reads columns A–E and A–I of those sheets by position.

## 3. Try it with the sample files first

`sample_msp_export/` holds 34 realistic exports: 3 projects, bi-weekly, monthly and 3-weekly cycles, and UK date format. They include summary tasks, ordinary tasks and resource units, so the filtering is exercised too. To try them:
1. Set `pMSPFolder` to that folder and `pDateCulture` to `en-GB`.
2. Click Refresh All.

The dashboard should show 40 milestones: 19 completed, 12 delayed, 3 ahead and 6 on schedule.

## Capacity

The MS Project edition holds 4,000 observation rows and 200 milestones. To extend it, use **Table Design → Resize Table** on `tblMilestoneData` or `tblMilestones`; the formulas fill the new rows automatically. The Delay Analysis limit is 200 milestones.

## Troubleshooting

| Symptom | Fix |
|---|---|
| Forecasts are blank or day and month are swapped | `pDateCulture` doesn't match MS Project's date format |
| Nothing loads | Check `pMSPFolder`, the `.csv` extension and the `<ProjectID>_<yyyy-mm-dd>.csv` naming |
| Everything lands in one column | Set `pDelimiter` to `;` |
| Accented names are garbled | Switch `pEncoding` between 1252 and 65001 |
| `Formula.Firewall` error | File → Options and settings → Query Options → Privacy → *Ignore the Privacy Levels* (current workbook) |
| Project shows as blank on the dashboard | Add the Project ID from the file name to `tblProjects` |
| The same date was exported twice | No action needed: the most recently saved file wins |

## How it was tested

`tools/msp_mirror.py` mirrors the query logic step for step in Python. `tools/qa_msp.py` then:
1. loads the mirror's output into the workbook exactly as Power Query would;
2. recalculates the workbook;
3. checks the result against an independent model of the MTA rules;
4. simulates the next cycle by dropping in a new export.

The M code itself hasn't been run inside Excel, because this build environment has no Excel or Power Query engine. The first refresh on your machine is the real check. If a step errors, the message names the step, and that is enough to fix it.
