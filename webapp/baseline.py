"""Monthly progress report (form-based data entry) and baseline change approval chain.

A baseline change request lists one or more milestones with a proposed new baseline and is
approved level by level (L1 -> L2 -> L3 ...). Only the next pending level is emailed. When the last
level approves, the new baselines are applied (old values kept in baseline_history). A rejection at
any level ends the request and leaves the baselines unchanged.
"""
import calendar
import datetime as dt
import secrets

SCHEMA = """
CREATE TABLE IF NOT EXISTS baseline_requests(id INTEGER PRIMARY KEY AUTOINCREMENT, project_id TEXT NOT NULL, period TEXT,
  requested_by TEXT, requester_email TEXT, created_at TEXT, status TEXT NOT NULL, current_level INTEGER, closed_at TEXT);
CREATE TABLE IF NOT EXISTS baseline_items(request_id INTEGER NOT NULL, milestone_id TEXT NOT NULL, old_baseline TEXT,
  new_baseline TEXT NOT NULL, reason TEXT);
CREATE TABLE IF NOT EXISTS baseline_approvals(request_id INTEGER NOT NULL, level INTEGER NOT NULL, approver TEXT NOT NULL,
  token TEXT UNIQUE NOT NULL, status TEXT NOT NULL, decided_by TEXT, decided_at TEXT, comment TEXT, PRIMARY KEY(request_id, level));
CREATE TABLE IF NOT EXISTS baseline_history(milestone_id TEXT NOT NULL, old_baseline TEXT, new_baseline TEXT, request_id INTEGER,
  changed_at TEXT);
CREATE TABLE IF NOT EXISTS project_approvers(project_id TEXT NOT NULL, level INTEGER NOT NULL, approver TEXT NOT NULL,
  PRIMARY KEY(project_id, level));
"""
PROGRESS = ("Completed", "In progress", "Yet to start")


def setup(con):
    con.executescript(SCHEMA)
    try:
        con.execute("ALTER TABLE observations ADD COLUMN progress TEXT")
    except Exception:
        pass
    con.commit()


def period_date(period):
    """'2026-11' -> reporting date: the 30th of the month (last day for shorter months)."""
    y, m = (int(x) for x in period.split("-"))
    return dt.date(y, m, min(30, calendar.monthrange(y, m)[1]))


def fmt(s):
    if not s:
        return "-"
    d = dt.date.fromisoformat(s) if isinstance(s, str) else s
    return d.strftime("%d-%b-%Y")


# ----------------------------------------------------------------------------- monthly report
def save_monthly(con, payload, now, emailer, url):
    """payload: project_id, period, submitted_by, submitter_email, rows[], baseline_changes[], approvers[].
    Returns dict(saved, errors, request_id)."""
    pid = payload["project_id"]
    if not con.execute("SELECT 1 FROM projects WHERE id=?", (pid,)).fetchone():
        return dict(errors=["Unknown project."])
    try:
        rep = period_date(payload["period"])
    except Exception:
        return dict(errors=["Choose a valid reporting period."])
    pr = con.execute("SELECT * FROM projects WHERE id=?", (pid,)).fetchone()
    if pr["start_date"] and pr["end_date"] and not (pr["start_date"][:7] <= payload["period"] <= pr["end_date"][:7]):
        return dict(errors=[f"The reporting period must be between the project start ({pr['start_date'][:7]}) and end ({pr['end_date'][:7]})."])
    ms = {r["id"]: dict(r) for r in con.execute("SELECT * FROM milestones WHERE project_id=?", (pid,))}
    errors, todo = [], []
    for i, r in enumerate(payload.get("rows", []), 1):
        mid = r.get("milestone_id")
        m = ms.get(mid)
        if not m:
            errors.append(f"Row {i}: unknown milestone {mid}")
            continue
        prog = r.get("progress") or ""
        if prog not in PROGRESS:
            errors.append(f"{mid}: choose a progress value")
            continue
        try:
            comp = dt.date.fromisoformat(r["completed"]) if r.get("completed") else None
            fc = dt.date.fromisoformat(r["forecast"]) if r.get("forecast") else None
        except ValueError:
            errors.append(f"{mid}: invalid date")
            continue
        if prog == "Completed":
            if not comp:
                errors.append(f"{mid}: enter the completed date")
                continue
            if comp > rep:
                errors.append(f"{mid}: completed date is after the reporting period ({fmt(rep)})")
                continue
            fc = comp
        elif comp:
            errors.append(f"{mid}: a completed date needs progress 'Completed'")
            continue
        if not fc:
            errors.append(f"{mid}: enter this month's forecast date (keep last month's date if it has not moved)")
            continue
        prev = con.execute("SELECT forecast FROM observations WHERE milestone_id=? AND rep_date<? ORDER BY rep_date DESC LIMIT 1",
                           (mid, rep.isoformat())).fetchone()
        reason = (r.get("reason") or "").strip()
        if prev and prev["forecast"] != fc.isoformat() and not reason:
            errors.append(f"{mid}: forecast moved from {fmt(prev['forecast'])} to {fmt(fc)} - enter a reason for variance")
            continue
        todo.append((mid, fc, comp, reason, prog))
    # baseline change request
    changes, chain = [], []
    bc = [c for c in payload.get("baseline_changes", []) if c.get("milestone_id") or c.get("new_baseline")]
    if bc:
        for c in bc:
            mid = c.get("milestone_id")
            if mid not in ms:
                errors.append("Baseline change: choose a milestone of this project")
                continue
            try:
                nb = dt.date.fromisoformat(c.get("new_baseline") or "")
            except ValueError:
                errors.append(f"Baseline change {mid}: enter the new baseline date")
                continue
            if nb.isoformat() == ms[mid]["baseline"]:
                errors.append(f"Baseline change {mid}: the new baseline equals the current one")
                continue
            if not (c.get("reason") or "").strip():
                errors.append(f"Baseline change {mid}: enter the reason for the change")
                continue
            if con.execute("""SELECT 1 FROM baseline_items i JOIN baseline_requests r ON r.id=i.request_id
                              WHERE i.milestone_id=? AND r.status LIKE 'Pending%'""", (mid,)).fetchone():
                errors.append(f"Baseline change {mid}: a change for this milestone is already awaiting approval")
                continue
            changes.append((mid, ms[mid]["baseline"], nb.isoformat(), c["reason"].strip()))
        chain = [a.strip() for a in payload.get("approvers", []) if a and a.strip()]
        if not chain:
            errors.append("Baseline change: enter at least the L1 approver's email")
        bad = [a for a in chain if "@" not in a or "." not in a.split("@")[-1] or " " in a]
        if bad:
            errors.append("Baseline change: approver emails are not valid: " + ", ".join(bad))
        if len(set(chain)) != len(chain):
            errors.append("Baseline change: the same approver is listed at two levels")
        if not (payload.get("submitted_by") or "").strip():
            errors.append("Baseline change: enter your name under Submitted by")
    if errors:
        return dict(errors=errors, saved=0)
    for mid, fc, comp, reason, prog in todo:
        con.execute("""INSERT INTO observations(milestone_id,rep_date,forecast,actual,comment,source,created_at,progress)
                       VALUES(?,?,?,?,?,?,?,?)
                       ON CONFLICT(milestone_id,rep_date) DO UPDATE SET forecast=excluded.forecast, actual=excluded.actual,
                         comment=excluded.comment, source=excluded.source, created_at=excluded.created_at, progress=excluded.progress""",
                    (mid, rep.isoformat(), fc.isoformat(), comp.isoformat() if comp else None, reason or None,
                     f"monthly report ({(payload.get('submitted_by') or 'admin').strip()})", now(), prog))
    req_id = None
    if changes:
        con.execute("DELETE FROM project_approvers WHERE project_id=?", (pid,))
        con.executemany("INSERT INTO project_approvers VALUES(?,?,?)", [(pid, i, a) for i, a in enumerate(chain, 1)])
        cur = con.execute("""INSERT INTO baseline_requests(project_id,period,requested_by,requester_email,created_at,status,current_level)
                             VALUES(?,?,?,?,?,?,1)""", (pid, payload["period"], payload["submitted_by"].strip(),
                                                         (payload.get("submitter_email") or "").strip(), now(), "Pending L1"))
        req_id = cur.lastrowid
        con.executemany("INSERT INTO baseline_items VALUES(?,?,?,?,?)", [(req_id,) + c for c in changes])
        for i, a in enumerate(chain, 1):
            con.execute("INSERT INTO baseline_approvals(request_id,level,approver,token,status) VALUES(?,?,?,?,?)",
                        (req_id, i, a, secrets.token_urlsafe(18), "Pending" if i == 1 else "Waiting"))
        notify_level(con, req_id, 1, emailer, url)
    con.commit()
    return dict(saved=len(todo), errors=[], request_id=req_id, rep_date=rep.isoformat())


# ----------------------------------------------------------------------------- approval chain
def request_detail(con, rid):
    r = con.execute("SELECT r.*, p.name AS project_name FROM baseline_requests r JOIN projects p ON p.id=r.project_id WHERE r.id=?",
                    (rid,)).fetchone()
    if not r:
        return None
    d = dict(r)
    d["items"] = [dict(x) for x in con.execute("""SELECT i.*, m.name AS milestone_name FROM baseline_items i
                                                  LEFT JOIN milestones m ON m.id=i.milestone_id WHERE request_id=?""", (rid,))]
    d["approvals"] = [dict(x) for x in con.execute("SELECT * FROM baseline_approvals WHERE request_id=? ORDER BY level", (rid,))]
    return d


def summary_text(d):
    lines = [f"  {i['milestone_id']} {i['milestone_name'] or ''}: {fmt(i['old_baseline'])} -> {fmt(i['new_baseline'])}"
             f" ({(dt.date.fromisoformat(i['new_baseline']) - dt.date.fromisoformat(i['old_baseline'])).days:+d} d). Reason: {i['reason']}"
             for i in d["items"]]
    return "\n".join(lines)


def notify_level(con, rid, level, emailer, url):
    d = request_detail(con, rid)
    a = next(x for x in d["approvals"] if x["level"] == level)
    n = len(d["approvals"])
    prior = "".join(f"  L{x['level']} {x['approver']}: approved by {x['decided_by']} on {x['decided_at']}"
                    + (f" - {x['comment']}" if x["comment"] else "") + "\n" for x in d["approvals"] if x["level"] < level)
    emailer(con, a["approver"], f"[MTA] Baseline change approval needed (L{level} of {n}): {d['project_id']} {d['project_name']}",
            f"A baseline change for {d['project_id']} {d['project_name']} needs your approval (level {level} of {n}).\n\n"
            f"Requested by: {d['requested_by']} on {d['created_at']} (reporting period {d['period']})\n\n"
            f"Changes:\n{summary_text(d)}\n\n" + (f"Already approved:\n{prior}\n" if prior else "") +
            f"Approve or reject here:\n{url}/baseline/{a['token']}\n", None)


def decide(con, token, decision, name, comment, now, emailer, url, on_applied):
    a = con.execute("SELECT * FROM baseline_approvals WHERE token=?", (token,)).fetchone()
    if not a:
        raise ValueError("This approval link is not valid.")
    d = request_detail(con, a["request_id"])
    if a["status"] != "Pending":
        raise ValueError(f"Level {a['level']} is {a['status'].lower()} - no decision is needed from this link.")
    if decision not in ("Approved", "Rejected") or not (name or "").strip():
        raise ValueError("Choose Approve or Reject and enter your name.")
    if decision == "Rejected" and not (comment or "").strip():
        raise ValueError("Give a reason when rejecting.")
    con.execute("UPDATE baseline_approvals SET status=?, decided_by=?, decided_at=?, comment=? WHERE request_id=? AND level=?",
                (decision, name.strip(), now(), (comment or "").strip(), a["request_id"], a["level"]))
    n = len(d["approvals"])
    everyone = ", ".join(dict.fromkeys([x["approver"] for x in d["approvals"]] + ([d["requester_email"]] if d["requester_email"] else [])))
    if decision == "Rejected":
        con.execute("UPDATE baseline_approvals SET status='Cancelled' WHERE request_id=? AND status='Waiting'", (a["request_id"],))
        con.execute("UPDATE baseline_requests SET status=?, closed_at=? WHERE id=?", (f"Rejected at L{a['level']}", now(), a["request_id"]))
        emailer(con, everyone, f"[MTA] Baseline change rejected at L{a['level']}: {d['project_id']} {d['project_name']}",
                f"{name.strip()} (L{a['level']}) rejected the baseline change requested by {d['requested_by']}.\nReason: {comment.strip()}\n\n"
                f"Changes (not applied):\n{summary_text(d)}\n", None)
    elif a["level"] < n:
        nxt = a["level"] + 1
        con.execute("UPDATE baseline_approvals SET status='Pending' WHERE request_id=? AND level=?", (a["request_id"], nxt))
        con.execute("UPDATE baseline_requests SET status=?, current_level=? WHERE id=?", (f"Pending L{nxt}", nxt, a["request_id"]))
        notify_level(con, a["request_id"], nxt, emailer, url)
    else:
        for i in d["items"]:
            con.execute("UPDATE milestones SET baseline=? WHERE id=?", (i["new_baseline"], i["milestone_id"]))
            con.execute("INSERT INTO baseline_history VALUES(?,?,?,?,?)", (i["milestone_id"], i["old_baseline"], i["new_baseline"],
                                                                            a["request_id"], now()))
        con.execute("UPDATE baseline_requests SET status='Approved', closed_at=? WHERE id=?", (now(), a["request_id"]))
        emailer(con, everyone, f"[MTA] Baseline change approved and applied: {d['project_id']} {d['project_name']}",
                f"All {n} approval level(s) approved. The new baselines are now in effect:\n{summary_text(d)}\n", None)
        on_applied()
    con.commit()
    return request_detail(con, a["request_id"])
