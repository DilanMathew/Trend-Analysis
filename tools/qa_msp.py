"""End-to-end test of the MS Project edition.

1. writes sample MS Project CSV exports (tools/msp_sample.py)
2. runs the Python mirror of powerquery/MSP_Import.pq on that folder
3. loads the result into the 'MSP Observations' / 'MSP Milestones' sheets exactly as
   Power Query's "Load To > existing worksheet A1" would, recalculates with LibreOffice
4. checks the workbook against the independent MTA model in qa_tests.py
5. drops a new CSV into the folder ("next reporting cycle"), reloads, and checks the update

Usage:  python tools/qa_msp.py <path-to-recalc.py>
"""
import csv
import datetime as dt
import os
import shutil
import sys
import tempfile

import openpyxl

sys.path.insert(0, os.path.dirname(__file__))
import msp_mirror  # noqa: E402
import msp_sample  # noqa: E402
import qa_tests as q  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "Advanced_Milestone_Trend_Analysis_MSProject.xlsx")


def load(folder, name):
    obs, ms = msp_mirror.observations(folder), msp_mirror.milestones(folder)
    wb = openpyxl.load_workbook(SRC)
    for sheet, hdrs, rows in (("MSP Observations", msp_mirror.OBS_HEADERS, obs), ("MSP Milestones", msp_mirror.MS_HEADERS, ms)):
        ws = wb[sheet]
        ws.append(hdrs)
        for r in rows:
            ws.append(r)
    p = os.path.join(q.TMP, name + ".xlsx")
    wb.save(p)
    q.recalc(p)
    return obs, ms, openpyxl.load_workbook(p, data_only=True)


def main():
    folder = os.path.join(q.TMP, "exports")
    shutil.copytree(msp_sample.build(), folder)
    print(f"== 1. Import {len(os.listdir(folder))} MS Project CSV exports")
    obs, ms, wbv = load(folder, "msp")
    q.check(len(ms) == 40, f"Mirror finds 40 milestone tasks (summary + ordinary tasks excluded): {len(ms)}")
    q.check(all(int(r[1].rsplit("-", 1)[1]) < 500 for r in ms), "No summary or ordinary task leaked into the milestone list")
    q.check(not any("[" in (r[5] or "") for r in ms), "Resource units stripped from owner names ('Tom Becker[50%]' -> 'Tom Becker')")
    reg = wbv["Milestone Register"]
    reg_ids = [reg.cell(r, 3).value for r in range(5, reg.max_row + 1) if reg.cell(r, 3).value]
    q.check(len(reg_ids) == 40, f"Milestone Register mirrors the query: {len(reg_ids)} milestones")
    data = wbv["Milestone Data"]
    rows = [r for r in range(5, data.max_row + 1) if data.cell(r, 1).value]
    q.check(len(rows) == len(obs), f"Milestone Data mirrors the query: {len(rows)} observations")
    checks = {data.cell(r, 21).value for r in rows}
    q.check(checks == {"OK"}, f"All imported observations pass Data Check: {checks}")
    blank_checks = {data.cell(r, 21).value for r in range(5, data.max_row + 1) if not data.cell(r, 1).value}
    q.check(blank_checks <= {None, ""}, "Unused capacity rows stay blank (no false data-check flags)")
    q.compare(wbv, wbv, "MSP import")
    errs = [(ws.title, c.coordinate, c.value) for ws in wbv for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and c.value[:4] in ("#REF", "#VAL", "#DIV", "#NAM", "#NUM", "#NUL")]
    q.check(not errs, f"No formula errors after import {errs[:3]}")
    slots, grid = q.grid(wbv)
    q.check(len(grid) == 13 and slots[0].startswith("PRJ-001-"), f"MTA chart shows PRJ-001's 13 reports with MS Project IDs ({slots[0]})")
    i108 = next(i for i, s in enumerate(slots) if s and s.startswith("PRJ-001-108 "))
    q.check(q.d(grid[-1][1][i108]) == dt.date(2027, 1, 1), "MS Project date 'Fri 01/01/27' plotted as 01-Jan-2027 for SIT Complete")
    done = next(i for i, s in enumerate(slots) if s and s.startswith("PRJ-001-101 "))
    q.check(all(r[1][done] == "#N/A" for r in grid[2:]), "Completed milestone line ends at completion although MS Project keeps exporting it")

    print("== 2. Next cycle: drop PRJ-001_2026-11-30.csv into the folder and refresh")
    src = os.path.join(folder, "PRJ-001_2026-11-16.csv")
    with open(src, encoding="cp1252", newline="") as fh:
        rows_csv = list(csv.DictReader(fh))
    for r in rows_csv:
        if r["Unique_ID"] == "108":
            r["Finish"], r["Notes"] = "Fri 15/01/27", "SIT defects, +14 d"
        if r["Unique_ID"] == "112":
            r["Finish"], r["Notes"] = "Fri 12/02/27", "Recovery plan, -7 d"
    with open(os.path.join(folder, "PRJ-001_2026-11-30.csv"), "w", encoding="cp1252", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=list(rows_csv[0].keys()))
        w.writeheader()
        w.writerows(rows_csv)
    obs2, ms2, wbv2 = load(folder, "msp_next")
    slots2, grid2 = q.grid(wbv2)
    q.check(grid2[-1][0] == dt.date(2026, 11, 30), "New reporting date appears on the MTA chart after refresh")
    q.check(q.d(grid2[-2][1][i108]) == dt.date(2027, 1, 1) and q.d(grid2[-1][1][i108]) == dt.date(2027, 1, 15),
            "History kept: SIT Complete shows 01-Jan at 16-Nov and 15-Jan at 30-Nov")
    da = q.read_da(wbv2)
    q.check(da["PRJ-001-108"]["move"] == 14 and da["PRJ-001-112"]["move"] == -7, "Forecast movement picked up (+14 d, -7 d)")
    q.compare(wbv2, wbv2, "after refresh")

    shutil.rmtree(q.TMP, ignore_errors=True)
    print()
    print("ALL TESTS PASSED" if not q.fails else f"{len(q.fails)} FAILURE(S)")
    sys.exit(1 if q.fails else 0)


if __name__ == "__main__":
    main()
