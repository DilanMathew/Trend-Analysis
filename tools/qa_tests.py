"""QA harness: builds scenario copies of the workbook, recalculates them with
LibreOffice and checks the results against an independent Python
implementation of the MTA rules.

Usage:  python tools/qa_tests.py <path-to-recalc.py>
"""
import copy
import datetime as dt
import os
import shutil
import subprocess
import sys
import tempfile
from collections import defaultdict

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "Advanced_Milestone_Trend_Analysis.xlsx")
RECALC = sys.argv[1]
TMP = tempfile.mkdtemp()
TOL, SIG, SLIP, HIGHVAR, WINDOW = 3, 10, 3, 20, 21
fails = []


def check(cond, msg):
    print(("PASS  " if cond else "FAIL  ") + msg)
    if not cond:
        fails.append(msg)


def recalc(path):
    out = subprocess.run([sys.executable, RECALC, path, "240"], capture_output=True, text=True)
    assert '"status"' in out.stdout, out.stdout + out.stderr
    return out.stdout


def d(v):
    return v.date() if isinstance(v, dt.datetime) else v


def add_table_row(ws, tname, values):
    """Simulate typing a new row directly below an Excel table (auto-expansion)."""
    t = ws.tables[tname]
    first, last = t.ref.split(":")
    last_col = "".join(c for c in last if c.isalpha())
    last_row = int("".join(c for c in last if c.isdigit()))
    new = last_row + 1
    ncols = openpyxl.utils.column_index_from_string(last_col)
    for c in range(1, ncols + 1):
        src = ws.cell(last_row, c)
        tgt = ws.cell(new, c)
        if c in values:
            tgt.value = values[c]
        elif isinstance(src.value, str) and src.value.startswith("="):
            tgt.value = src.value
        tgt.number_format = src.number_format
    t.ref = f"{first}:{last_col}{new}"


# ---------------------------------------------------------------- reference model
def model(wb, sel_project="All Projects", sel_period=None, sel_cat="All Categories", sel_status="All Statuses"):
    reg = wb["Milestone Register"]
    ms = []
    for r in range(5, reg.max_row + 1):
        if reg.cell(r, 3).value:
            ms.append(dict(proj=reg.cell(r, 1).value, id=reg.cell(r, 3).value, base=d(reg.cell(r, 6).value),
                           cat=reg.cell(r, 10).value))
    pnames = {}
    ps = wb["Project Setup"]
    for r in range(6, 20):
        if ps.cell(r, 2).value and str(ps.cell(r, 2).value).startswith("PRJ"):
            pnames[ps.cell(r, 2).value] = ps.cell(r, 3).value
    data = wb["Milestone Data"]
    obs = defaultdict(list)
    proj_reps = defaultdict(set)
    projof = {m["id"]: m["proj"] for m in ms}
    for r in range(5, data.max_row + 1):
        mid = data.cell(r, 1).value
        if not mid:
            continue
        rep, fc, act = d(data.cell(r, 2).value), d(data.cell(r, 3).value), d(data.cell(r, 4).value)
        obs[mid].append((rep, fc, act))
        proj_reps[projof.get(mid)].add(rep)
    res = {}
    for m in ms:
        asof = sel_period or max(proj_reps[m["proj"]], default=None)
        o = sorted(x for x in obs[m["id"]] if asof and x[0] <= asof)
        st = "Not Yet Started"
        row = dict(nobs=len(o))
        if o:
            cur = o[-1][1]
            acts = [x[2] for x in o if x[2]]
            var = cur - m["base"]
            row.update(curfc=cur, basevar=var.days, move=(cur - o[-2][1]).days if len(o) > 1 else None)
            if acts:
                st = "Completed"
            elif var.days > TOL:
                st = "Delayed"
            elif var.days < -TOL:
                st = "Ahead of Schedule"
            else:
                st = "On Schedule"
            slips = sum(1 for a, b in zip(o, o[1:]) if b[1] > a[1])
            row["slips"] = slips
            row["totdelay"] = ((max(acts) if acts else cur) - m["base"]).days
        row["status"] = st
        inf = ((sel_project == "All Projects" or pnames[m["proj"]] == sel_project)
               and (sel_cat == "All Categories" or m["cat"] == sel_cat)
               and (sel_status == "All Statuses" or st == sel_status))
        row["infilter"] = inf
        res[m["id"]] = row
    return res


def read_da(wbv):
    da = wbv["Delay Analysis"]
    hdr = {da.cell(15, c).value: c for c in range(1, da.max_column + 1)}
    out = {}
    for r in range(16, 216):
        mid = da.cell(r, 2).value
        if mid:
            g = lambda h: da.cell(r, hdr[h]).value  # noqa: E731
            out[mid] = dict(status=g("Schedule Status"), basevar=g("Baseline Variance (d)"), move=g("Forecast Movement (d)"),
                            slips=g("Slip Count"), totdelay=g("Total Delay (d)"), infilter=g("In Dashboard Filter"),
                            nobs=g("Observations"), curfc=d(g("Current Forecast")))
    return out


def compare(wbf, wbv, label, **sel):
    exp = model(wbf, **sel)
    got = read_da(wbv)
    bad = []
    for mid, e in exp.items():
        g = got.get(mid)
        if g is None:
            bad.append((mid, "missing"))
            continue
        if g["status"] != e["status"] or int(g["infilter"]) != int(e["infilter"]) or g["nobs"] != e["nobs"]:
            bad.append((mid, e["status"], g["status"], e["infilter"], g["infilter"]))
        if e["nobs"]:
            if g["basevar"] != e["basevar"] or g["totdelay"] != e["totdelay"] or g["slips"] != e["slips"] or g["curfc"] != e["curfc"]:
                bad.append((mid, "values", e, g))
            if (e["move"] if e["move"] is not None else None) != (g["move"] if g["move"] != "" else None):
                bad.append((mid, "move", e["move"], g["move"]))
    check(not bad, f"{label}: Delay Analysis matches independent model for {len(exp)} milestones {bad[:3]}")
    dash = wbv["Dashboard"]
    kp = {dash.cell(7, c).value: dash.cell(8, c).value for c in range(2, 26, 3)}
    inf = [e for e in exp.values() if e["infilter"]]
    check(kp["TOTAL MILESTONES"] == len(inf), f"{label}: KPI total milestones = {len(inf)} (got {kp['TOTAL MILESTONES']})")
    for s, k in (("Completed", "COMPLETED"), ("Delayed", "DELAYED"), ("Ahead of Schedule", "AHEAD OF SCHEDULE"), ("On Schedule", "ON SCHEDULE")):
        n = sum(1 for e in inf if e["status"] == s)
        check(kp[k] == n, f"{label}: KPI {k.lower()} = {n} (got {kp[k]})")
    tds = [e["totdelay"] for e in inf if e["nobs"]]
    avg = sum(tds) / len(tds) if tds else 0
    check(abs((kp["AVERAGE DELAY"] or 0) - avg) < 1e-9, f"{label}: KPI average delay = {avg:.3f} (got {kp['AVERAGE DELAY']})")
    return exp, got


def scenario(name, mutate):
    p = os.path.join(TMP, name + ".xlsx")
    wb = openpyxl.load_workbook(SRC)
    mutate(wb)
    wb.save(p)
    recalc(p)
    return openpyxl.load_workbook(p), openpyxl.load_workbook(p, data_only=True)


def grid(wbv):
    e = wbv["Engine"]
    hdr_row = next(r for r in range(1, 200) if e.cell(r, 1).value == "Reporting date")
    slots = [e.cell(hdr_row, 2 + k).value for k in range(1, 16)]
    rows = []
    r = hdr_row + 1
    while e.cell(r, 1).value not in (None, ""):
        rows.append((d(e.cell(r, 1).value), [e.cell(r, 2 + k).value for k in range(1, 16)]))
        r += 1
    return slots, rows


# ---------------------------------------------------------------- tests
print("== 1. Base workbook")
wbf, wbv = scenario("base", lambda wb: None)
compare(wbf, wbv, "base")
slots, rows = grid(wbv)
exp_reps = sorted({d(wbf["Milestone Data"].cell(r, 2).value) for r in range(5, wbf["Milestone Data"].max_row + 1)
                   if wbv["Milestone Data"].cell(r, 6).value == "PRJ-001"})
check([r[0] for r in rows] == exp_reps, f"MTA grid has the {len(exp_reps)} PRJ-001 reporting dates in order")
# M-114 skipped the 16-Nov report -> must be #N/A (not plotted), not 0 or carried forward
i114 = next(i for i, s in enumerate(slots) if s and s.startswith("M-114"))
last = rows[-1]
check(last[0] == dt.date(2026, 11, 16) and last[1][i114] == "#N/A", "Missing observation (M-114 @ 16-Nov) is #N/A, not zero or invented")
# completed milestones stop after completion
i101 = next(i for i, s in enumerate(slots) if s and s.startswith("M-101"))
check(all(r[1][i101] == "#N/A" for r in rows[2:]), "Completed milestone line (M-101) ends at completion")
# grid values equal raw data
data = wbf["Milestone Data"]
raw = {(data.cell(r, 1).value, d(data.cell(r, 2).value)): d(data.cell(r, 3).value) for r in range(5, data.max_row + 1)}
mism = [(s, rep) for rep, vals in rows for s, v in zip(slots, vals) if s and v != "#N/A" and d(v) != raw.get((s.split()[0], rep))]
check(not mism, "Every plotted MTA point equals the recorded forecast for that milestone/date")
check(sum(1 for r in range(5, data.max_row + 1) if wbv["Milestone Data"].cell(r, 21).value != "OK") == 0, "All sample observations pass Data Check")
errs = [(ws.title, c.coordinate, c.value) for ws in wbv for row in ws.iter_rows() for c in row
        if isinstance(c.value, str) and c.value[:4] in ("#REF", "#VAL", "#DIV", "#NAM", "#NUM", "#NUL") and ws.title != "Instructions"]
check(not errs, f"No #REF/#VALUE/#DIV/#NAME/#NUM errors anywhere {errs[:3]}")
na_outside = [(ws.title, c.coordinate) for ws in wbv if ws.title not in ("Engine", "Instructions") for row in ws.iter_rows() for c in row if c.value == "#N/A"]
check(not na_outside, f"#N/A only appears in Engine chart series {na_outside[:3]}")

print("== 2. Project + category filter")
def m2(wb):
    wb["Dashboard"]["B5"] = "Customer Portal Redesign"
    wb["Dashboard"]["K5"] = "Build"
wbf, wbv = scenario("filter_proj_cat", m2)
compare(wbf, wbv, "PRJ-003/Build", sel_project="Customer Portal Redesign", sel_cat="Build")
slots, rows = grid(wbv)
ids = [s.split()[0] for s in slots if s]
check(ids == ["M-305", "M-306", "M-307", "M-311"], f"MTA slots follow project+category filter: {ids}")

print("== 3. Status filter + historical reporting period")
def m3(wb):
    wb["Dashboard"]["G5"] = dt.datetime(2026, 9, 14)
    wb["Dashboard"]["P5"] = "Delayed"
wbf, wbv = scenario("filter_period_status", m3)
compare(wbf, wbv, "as-of 14-Sep-2026 / Delayed", sel_period=dt.date(2026, 9, 14), sel_status="Delayed")
slots, rows = grid(wbv)
check(rows[-1][0] <= dt.date(2026, 9, 14), f"MTA grid stops at the selected reporting period (last {rows[-1][0]})")

print("== 4. Hide a milestone line")
def m4(wb):
    wb["Dashboard"]["Z22"] = "No"   # slot 8
wbf, wbv = scenario("hide", m4)
slots, rows = grid(wbv)
check(all(r[1][7] == "#N/A" for r in rows) and all(r[1][8] != "#N/A" for r in rows[:1]), "Show=No hides only that milestone's series")

print("== 5. Add a new milestone and a new reporting period")
def m5(wb):
    reg = wb["Milestone Register"]
    add_table_row(reg, "tblMilestones", {1: "PRJ-001", 3: "M-115", 4: "Data Archive Validated", 5: "Archive reconciliation", 6: dt.datetime(2027, 3, 26),
                                         7: "Raj Malhotra", 8: "Information Technology", 9: "Medium", 10: "Governance"})
    ws = wb["Milestone Data"]
    new = [("M-108", dt.datetime(2027, 1, 15), None, "Further SIT defects (+14 d)"), ("M-109", dt.datetime(2027, 1, 8), None, ""),
           ("M-110", dt.datetime(2027, 1, 8), None, ""), ("M-111", dt.datetime(2027, 1, 29), None, ""),
           ("M-112", dt.datetime(2027, 2, 12), None, "Recovery plan (-7 d)"), ("M-113", dt.datetime(2027, 3, 12), None, ""),
           ("M-114", dt.datetime(2027, 3, 5), None, ""), ("M-115", dt.datetime(2027, 3, 26), None, "First report")]
    for mid, fc, act, cm in new:
        add_table_row(ws, "tblMilestoneData", {1: mid, 2: dt.datetime(2026, 11, 30), 3: fc, 4: act, 5: cm})
wbf, wbv = scenario("add_rows", m5)
compare(wbf, wbv, "after additions")
slots, rows = grid(wbv)
check(rows[-1][0] == dt.date(2026, 11, 30), "New reporting period (30-Nov-2026) appears automatically on the MTA chart")
check(slots[14] and slots[14].startswith("M-115"), f"New milestone M-115 gets its own MTA series (slot 15 = {slots[14]})")
i112 = next(i for i, s in enumerate(slots) if s and s.startswith("M-112"))
check(d(rows[-1][1][i112]) == dt.date(2027, 2, 12) and d(rows[-2][1][i112]) == dt.date(2027, 2, 19), "History intact: M-112 keeps 19-Feb forecast at 16-Nov and shows 12-Feb at 30-Nov")
da = read_da(wbv)
check(da["M-112"]["move"] == -7 and da["M-108"]["move"] == 14, "Forecast movement recalculated for new period (M-112 -7 d, M-108 +14 d)")
check(da["M-114"]["nobs"] == 13, "M-114 now has its 13th observation")
dash = wbv["Dashboard"]
check(dash["B8"].value == 41, "KPI total milestones includes the new milestone (41)")

print("== 6. Duplicate detection")
def m6(wb):
    add_table_row(wb["Milestone Data"], "tblMilestoneData", {1: "M-108", 2: dt.datetime(2026, 11, 16), 3: dt.datetime(2027, 1, 1)})
wbf, wbv = scenario("dup", m6)
ws = wbv["Milestone Data"]
dups = [ws.cell(r, 21).value for r in range(5, ws.max_row + 1) if ws.cell(r, 1).value == "M-108" and d(ws.cell(r, 2).value) == dt.date(2026, 11, 16)]
check(dups == ["DUPLICATE observation"] * 2, f"Duplicate milestone/date pair flagged on both rows: {dups}")

print("== 7. Reporting History entry form")
def m7(wb):
    rh = wb["Reporting History"]
    rh["G17"] = dt.datetime(2027, 1, 8)     # M-108
    rh["G18"] = dt.datetime(2027, 1, 8)     # M-109
    rh["I17"] = "Test comment"
wbf, wbv = scenario("rh", m7)
rh = wbv["Reporting History"]
check(rh["K17"].value == "✔ Ready" and rh["J17"].value == 7, f"Form validates entry and shows movement (+7 d): {rh['K17'].value}, {rh['J17'].value}")
blk = [(rh.cell(r, 15).value, d(rh.cell(r, 16).value), d(rh.cell(r, 17).value), rh.cell(r, 19).value) for r in range(10, 13)]
check(blk[0] == ("M-108", dt.date(2026, 11, 30), dt.date(2027, 1, 8), "Test comment") and blk[1][0] == "M-109" and blk[2][0] in ("", None),
      f"Paste-ready block compacts ready rows: {blk}")
def m7b(wb):
    rh = wb["Reporting History"]
    rh["C5"] = dt.datetime(2026, 11, 16)
    rh["G17"] = dt.datetime(2027, 1, 8)
wbf, wbv = scenario("rh_dup", m7b)
check(wbv["Reporting History"]["K17"].value.startswith("✖ Duplicate"), "Form blocks a duplicate snapshot date")

shutil.rmtree(TMP, ignore_errors=True)
print()
print("ALL TESTS PASSED" if not fails else f"{len(fails)} FAILURE(S)")
sys.exit(1 if fails else 0)
