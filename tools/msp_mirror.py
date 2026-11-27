"""Python mirror of powerquery/MSP_Import.pq, step for step.

Used by tools/qa_msp.py to load the MS Project edition of the workbook exactly the way the
Power Query output would, so the workbook side can be tested without Excel.
"""
import csv
import os
import re
from datetime import date, datetime

ALIASES = {
    "uniqueid": ["uniqueid", "uid"], "name": ["name", "taskname"], "milestone": ["milestone"],
    "summary": ["summary"], "baselinefinish": ["baselinefinish", "baseline0finish"],
    "finish": ["finish", "finishdate"], "actualfinish": ["actualfinish"],
    "percentcomplete": ["percentcomplete", "complete", "pctcomplete"],
    "resourcenames": ["resourcenames", "resources"], "priority": ["priority"],
    "category": ["text1", "category"], "notes": ["notes"],
}
YES = {"yes", "true", "1", "-1", "ja", "oui", "si", "sí"}


def header(h):
    return re.sub(r"[^A-Za-z0-9]", "", h).lower()


def msp_date(v, culture="en-GB"):
    txt = (v or "").strip()
    words = [w for w in txt.split(" ") if w]
    w0 = words[0].replace(",", "").replace(".", "").lower() if words else ""
    if len(words) > 1 and any(w0.startswith(d) for d in ("mon", "tue", "wed", "thu", "fri", "sat", "sun")) and not re.search(r"\d", w0):
        words = words[1:]
    words = [w for w in words if ":" not in w and w.lower() not in ("am", "pm")]
    clean = " ".join(words)
    if not clean or clean.upper() == "NA":
        return None
    fmts = ["%d/%m/%y", "%d/%m/%Y", "%d %B %Y", "%d %b %Y"] if culture == "en-GB" else ["%m/%d/%y", "%m/%d/%Y", "%B %d %Y"]
    for fm in fmts:
        try:
            return datetime.strptime(clean.replace(",", ""), fm).date()
        except ValueError:
            pass
    return None


def files(folder):
    out = []
    for name in os.listdir(folder):
        if not name.lower().endswith(".csv") or name.startswith("~"):
            continue
        stem = name.rsplit(".", 1)[0]
        parts = stem.split("_")
        path = os.path.join(folder, name)
        try:
            rep = date.fromisoformat(parts[-1].strip()) if len(parts) > 1 else None
        except ValueError:
            rep = None
        rep = rep or datetime.fromtimestamp(os.path.getmtime(path)).date()
        with open(path, encoding="cp1252", newline="") as fh:
            rows = list(csv.reader(fh))
        hdr = [header(h) for h in rows[0]]
        data = [dict(zip(hdr, r)) for r in rows[1:]]
        out.append(dict(file=name, modified=os.path.getmtime(path), project=parts[0].strip(), rep=rep, data=data))
    return out


def tasks(folder):
    res = []
    for f in files(folder):
        cols = set(f["data"][0].keys()) if f["data"] else set()
        pick = {k: next((a for a in al if a in cols), None) for k, al in ALIASES.items()}
        for r in f["data"]:
            g = {k: (r.get(c) if c else None) for k, c in pick.items()}
            if (g["milestone"] or "").strip().lower() not in YES or (g["summary"] or "").strip().lower() in YES:
                continue
            if not (g["uniqueid"] or "").strip():
                continue
            fc = msp_date(g["finish"])
            try:
                pct = float(re.sub(r"[^0-9.]", "", g["percentcomplete"] or ""))
            except ValueError:
                pct = None
            act = msp_date(g["actualfinish"]) or (fc if pct is not None and pct >= 100 else None)
            owner = None if g["resourcenames"] is None else ", ".join(x.split("[")[0].strip() for x in g["resourcenames"].split(","))
            if fc is None:
                continue
            res.append(dict(id=f"{f['project']}-{g['uniqueid'].strip()}", project=f["project"], rep=f["rep"], mod=f["modified"],
                            fc=fc, actual=act, base=msp_date(g["baselinefinish"]), comment=(g["notes"] or "")[:250] or None,
                            owner=owner, name=g["name"], prio=g["priority"], cat=g["category"], uid=g["uniqueid"].strip()))
    return res


def observations(folder):
    t = sorted(tasks(folder), key=lambda x: (x["id"], x["rep"], -x["mod"]))
    seen, dedup = set(), []
    for x in t:
        if (x["id"], x["rep"]) not in seen:
            seen.add((x["id"], x["rep"]))
            dedup.append(x)
    first_done = {}
    for x in dedup:
        if x["actual"]:
            first_done[x["id"]] = min(first_done.get(x["id"], x["rep"]), x["rep"])
    kept = [x for x in dedup if x["id"] not in first_done or x["rep"] <= first_done[x["id"]]]
    kept.sort(key=lambda x: (x["rep"], x["id"]))
    return [[x["id"], x["rep"], x["fc"], x["actual"], x["comment"]] for x in kept]


def milestones(folder):
    t = tasks(folder)
    by = {}
    for x in t:
        by.setdefault(x["id"], []).append(x)
    out = []
    for mid, xs in by.items():
        latest = sorted(xs, key=lambda x: (x["rep"], x["mod"]))[-1]
        known = [x for x in sorted(xs, key=lambda x: x["rep"]) if x["base"]]
        base = known[-1]["base"] if known else sorted(xs, key=lambda x: x["rep"])[0]["fc"]
        try:
            p = float(latest["prio"])
        except (TypeError, ValueError):
            p = 500
        prio = "Critical" if p >= 800 else "High" if p >= 600 else "Medium" if p >= 400 else "Low"
        cat = (latest["cat"] or "").strip() or "Milestone"
        out.append([latest["project"], mid, latest["name"], f"MS Project task UID {latest['uid']}", base,
                    latest["owner"], None, prio, cat])
    out.sort(key=lambda r: (r[0], r[4], r[1]))
    return out


OBS_HEADERS = ["Milestone ID", "Reporting Date", "Forecast Completion Date", "Actual Completion Date", "Comments"]
MS_HEADERS = ["Project ID", "Milestone ID", "Milestone Name", "Milestone Description", "Baseline Date",
              "Responsible Person", "Department", "Priority", "Category"]
