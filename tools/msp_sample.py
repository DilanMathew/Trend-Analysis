"""Writes sample MS Project CSV exports (one per project per reporting date) to
powerquery/sample_msp_export/, in the shape MS Project's Save As > CSV export map produces:
UK date format ("Mon 30/11/26"), "NA" for empty dates, Yes/No flags, summary and
ordinary tasks mixed in with milestones, resource names with units ("Tom Becker[50%]").
"""
import csv
import os
import sys
from datetime import timedelta

sys.path.insert(0, os.path.dirname(__file__))
import sample_data  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "powerquery", "sample_msp_export")
HEADERS = ["ID", "Unique_ID", "Name", "Duration", "Milestone", "Summary", "Baseline_Finish", "Finish",
           "Actual_Finish", "Percent_Complete", "Resource_Names", "Priority", "Text1", "Notes"]
PRIO = {"Critical": 900, "High": 700, "Medium": 500, "Low": 300}


def msp_date(d):
    return "NA" if d is None else d.strftime("%a %d/%m/%y")


def build():
    os.makedirs(OUT, exist_ok=True)
    for f in os.listdir(OUT):
        os.remove(os.path.join(OUT, f))
    for p in sample_data.PROJECTS:
        rows = sample_data.MILESTONES[p["id"]]
        state = {m[0]: {"fc": m[4] + timedelta(days=m[9]), "actual": None, "note": ""} for m in rows}
        for i, rep in enumerate(p["reports"]):
            for (mid, name, desc, cat, base, owner, dept, prio, deltas, init, skip) in rows:
                s = state[mid]
                s["note"] = ""
                if s["actual"]:
                    continue
                d = deltas.get(i, 0)
                s["fc"] += timedelta(days=d)
                if d:
                    s["note"] = ("Slipped" if d > 0 else "Pulled in") + f" {d:+d} d, reported {rep:%d-%b}"
                if s["fc"] <= rep:
                    s["actual"] = s["fc"]
            out = []
            uid_task = 1000
            for ci, cat in enumerate(dict.fromkeys(m[3] for m in rows)):
                out.append(dict(Unique_ID=500 + ci, Name=f"{cat} phase", Duration="40 days", Milestone="No", Summary="Yes",
                                Baseline_Finish="NA", Finish=msp_date(rep + timedelta(days=60)), Actual_Finish="NA",
                                Percent_Complete="50%", Resource_Names="", Priority=500, Text1="", Notes=""))
                for (mid, name, desc, mcat, base, owner, dept, prio, deltas, init, skip) in rows:
                    if mcat != cat:
                        continue
                    uid = int(mid.split("-")[1])
                    s = state[mid]
                    uid_task += 1
                    out.append(dict(Unique_ID=uid_task, Name=f"Prepare: {name}", Duration="10 days", Milestone="No", Summary="No",
                                    Baseline_Finish=msp_date(base - timedelta(days=3)), Finish=msp_date(s["fc"] - timedelta(days=3)),
                                    Actual_Finish="NA", Percent_Complete="0%", Resource_Names=owner, Priority=500, Text1=cat, Notes=""))
                    owner_txt = owner + ("[50%]" if uid % 3 == 0 else "")
                    out.append(dict(Unique_ID=uid, Name=name, Duration="0 days", Milestone="Yes", Summary="No",
                                    Baseline_Finish=msp_date(base), Finish=msp_date(s["fc"]),
                                    Actual_Finish=msp_date(s["actual"]), Percent_Complete="100%" if s["actual"] else "0%",
                                    Resource_Names=owner_txt, Priority=PRIO[prio], Text1=cat, Notes=s["note"]))
            fn = os.path.join(OUT, f"{p['id']}_{rep:%Y-%m-%d}.csv")
            with open(fn, "w", newline="", encoding="cp1252") as fh:
                w = csv.DictWriter(fh, fieldnames=HEADERS)
                w.writeheader()
                for k, row in enumerate(out, 1):
                    row["ID"] = k
                    w.writerow(row)
    return OUT


if __name__ == "__main__":
    out = build()
    print(len(os.listdir(out)), "files in", out)
