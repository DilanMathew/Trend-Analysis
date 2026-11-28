"""Monthly progress report and the project's approval workflows.

  * project    - creating a project. Approved by the Project Manager alone (one level: the PM
                 name/email entered on Project Setup) - no configurable chain. The project is
                 'Pending' until the PM approves it, and only then can its baseline be set.
  * baseline   - setting up the project's first (initial) baseline, once the project itself is
                 approved. Configurable Level 1, 2, 3 (name, email) on Project Setup.
  * addition   - adding further milestones into an existing baseline. New milestones are Draft
                 until submitted; they join the baseline only when every level has approved.
                 Configurable Level 1, 2, 3 (name, email).
  * rebaseline - changing the baseline date of milestones already in the baseline. Configurable
                 Level 1, 2, 3 (name, email). Each approved baseline (the initial one and every
                 rebaseline after it) is numbered in sequence - Baseline 1, Baseline 2, Baseline 3,
                 ... - tracked via projects.baseline_no / baseline_requests.baseline_no.
A request is approved level by level (L1 -> L2 -> L3). Only the level whose turn it is gets a link.
A rejection at any level ends the request.
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
CREATE TABLE IF NOT EXISTS project_workflows(project_id TEXT NOT NULL, workflow TEXT NOT NULL, level INTEGER NOT NULL,
  name TEXT, person_id TEXT, email TEXT, PRIMARY KEY(project_id, workflow, level));
"""
MIGRATIONS = [
    "ALTER TABLE observations ADD COLUMN progress TEXT",
    "ALTER TABLE baseline_requests ADD COLUMN kind TEXT DEFAULT 'rebaseline'",
    "ALTER TABLE baseline_approvals ADD COLUMN approver_name TEXT",
    "ALTER TABLE baseline_approvals ADD COLUMN approver_id TEXT",
    "ALTER TABLE milestones ADD COLUMN baseline_status TEXT DEFAULT 'Approved'",
    "ALTER TABLE milestones ADD COLUMN status_note TEXT",
    "ALTER TABLE milestones ADD COLUMN rebase_req INTEGER",
    "ALTER TABLE baseline_requests ADD COLUMN baseline_no INTEGER",
    "ALTER TABLE projects ADD COLUMN baseline_no INTEGER DEFAULT 0",
]
PROGRESS = ("Completed", "In progress", "Yet to start")
WORKFLOWS = {"baseline": "Setting up a baseline",
             "addition": "Addition of milestones into baseline", "rebaseline": "Rebaseline of milestones"}
KIND_LABELS = {"project": "Project creation", **WORKFLOWS}


def setup(con):
    con.executescript(SCHEMA)
    for m in MIGRATIONS:
        try:
            con.execute(m)
        except Exception:
            pass
    # the "creation" workflow was renamed to "baseline"; carry over any already-configured chain
    try:
        con.execute("UPDATE project_workflows SET workflow='baseline' WHERE workflow='creation'")
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


# ----------------------------------------------------------------------------- workflow configuration
def get_workflows(con, pid):
    out = {k: [] for k in WORKFLOWS}
    for r in con.execute("SELECT * FROM project_workflows WHERE project_id=? ORDER BY workflow, level", (pid,)):
        if r["workflow"] in out:
            out[r["workflow"]].append(dict(level=r["level"], name=r["name"] or "", person_id=r["person_id"] or "", email=r["email"] or ""))
    return out


def chain(con, pid, workflow):
    """Configured levels with a name, in order."""
    return [lv for lv in get_workflows(con, pid)[workflow] if lv["name"].strip()]


def save_workflows(con, pid, wf):
    """wf: {workflow: [{name, person_id, email}, ...]} (index = level - 1). Returns list of errors."""
    errors = []
    for key, levels in (wf or {}).items():
        if key not in WORKFLOWS:
            continue
        filled = [bool((lv.get("name") or "").strip()) for lv in levels]
        if any(filled) and not all(filled[:max(i for i, f in enumerate(filled) if f) + 1]):
            errors.append(f"{WORKFLOWS[key]}: fill the levels in order (Level 1 first, no gaps).")
        for i, lv in enumerate(levels, 1):
            if (lv.get("name") or "").strip():
                em = (lv.get("email") or "").strip()
                if em and ("@" not in em or "." not in em.split("@")[-1] or " " in em):
                    errors.append(f"{WORKFLOWS[key]} Level {i}: '{em}' is not a valid email address.")
    if errors:
        return errors
    for key, levels in (wf or {}).items():
        if key not in WORKFLOWS:
            continue
        con.execute("DELETE FROM project_workflows WHERE project_id=? AND workflow=?", (pid, key))
        for i, lv in enumerate(levels, 1):
            if (lv.get("name") or "").strip():
                con.execute("INSERT INTO project_workflows VALUES(?,?,?,?,?,?)",
                            (pid, key, i, lv["name"].strip(), (lv.get("person_id") or "").strip(), (lv.get("email") or "").strip()))
    return []


# ----------------------------------------------------------------------------- requests
def create_request(con, pid, kind, items, requested_by, requester_email, period, now, emailer, url, draft=False, levels=None):
    """items: [(milestone_id, old_baseline, new_baseline, reason)] - old_baseline None = new milestone.
    levels: explicit [{level,name,email}, ...] to approve with, instead of the project's configured
    chain for `kind` - used for kind='project', approved by the Project Manager alone.
    Returns request id. A draft gets no approvers until submit_request()."""
    cur = con.execute("""INSERT INTO baseline_requests(project_id,period,requested_by,requester_email,created_at,status,current_level,kind)
                         VALUES(?,?,?,?,?,?,1,?)""", (pid, period, requested_by, requester_email, now(), "Draft", kind))
    rid = cur.lastrowid
    con.executemany("INSERT INTO baseline_items VALUES(?,?,?,?,?)", [(rid,) + tuple(i) for i in items])
    if not draft:
        submit_request(con, rid, requested_by, requester_email, now, emailer, url, levels=levels)
    return rid


def submit_request(con, rid, requested_by, requester_email, now, emailer, url, levels=None):
    r = con.execute("SELECT * FROM baseline_requests WHERE id=?", (rid,)).fetchone()
    pid, kind = r["project_id"], r["kind"] or "rebaseline"
    levels = levels if levels is not None else chain(con, pid, kind)
    con.execute("DELETE FROM baseline_approvals WHERE request_id=?", (rid,))
    con.execute("UPDATE baseline_requests SET status='Pending L1', current_level=1, requested_by=?, requester_email=?, created_at=? WHERE id=?",
                (requested_by, requester_email, now(), rid))
    for i in con.execute("SELECT milestone_id, old_baseline FROM baseline_items WHERE request_id=?", (rid,)).fetchall():
        if kind == "addition" or i["old_baseline"] is None:
            con.execute("UPDATE milestones SET baseline_status='Pending', status_note='Awaiting Level 1 approval' WHERE id=?", (i["milestone_id"],))
    for lv in levels:
        con.execute("""INSERT INTO baseline_approvals(request_id,level,approver,token,status,approver_name,approver_id)
                       VALUES(?,?,?,?,?,?,?)""", (rid, lv["level"], lv["email"], secrets.token_urlsafe(18),
                                                  "Pending" if lv["level"] == 1 else "Waiting", lv["name"], lv.get("person_id", "")))
    notify_level(con, rid, 1, emailer, url)
    return rid


def request_detail(con, rid):
    r = con.execute("SELECT r.*, p.name AS project_name FROM baseline_requests r JOIN projects p ON p.id=r.project_id WHERE r.id=?",
                    (rid,)).fetchone()
    if not r:
        return None
    d = dict(r)
    d["kind"] = d.get("kind") or "rebaseline"
    d["kind_label"] = KIND_LABELS.get(d["kind"], d["kind"])
    d["items"] = [dict(x) for x in con.execute("""SELECT i.*, m.name AS milestone_name FROM baseline_items i
                                                  LEFT JOIN milestones m ON m.id=i.milestone_id WHERE request_id=?""", (rid,))]
    d["approvals"] = [dict(x) for x in con.execute("SELECT * FROM baseline_approvals WHERE request_id=? ORDER BY level", (rid,))]
    return d


def who(a):
    nm = a.get("approver_name") or a.get("approver") or "-"
    return f"{nm} ({a['approver_id']})" if a.get("approver_id") else nm


def project_text(con, pid):
    pr = con.execute("SELECT * FROM projects WHERE id=?", (pid,)).fetchone()
    if not pr:
        return ""
    return (f"  Project ID: {pr['id']}\n  Name: {pr['name']}\n"
            f"  Start: {fmt(pr['start_date'])}   End: {fmt(pr['end_date'])}\n"
            f"  Controls: {pr['controls_name'] or '-'}   Manager: {pr['manager'] or '-'}\n")


def summary_text(d):
    lines = []
    for i in d["items"]:
        if d["kind"] == "addition" or i["old_baseline"] is None:
            lines.append(f"  {i['milestone_id']} {i['milestone_name'] or ''}: NEW milestone, baseline {fmt(i['new_baseline'])}"
                         + (f". Reason: {i['reason']}" if i.get("reason") else ""))
        else:
            lines.append(f"  {i['milestone_id']} {i['milestone_name'] or ''}: {fmt(i['old_baseline'])} -> {fmt(i['new_baseline'])}"
                         f" ({(dt.date.fromisoformat(i['new_baseline']) - dt.date.fromisoformat(i['old_baseline'])).days:+d} d). "
                         f"Reason: {i['reason']}")
    return "\n".join(lines)


def notify_level(con, rid, level, emailer, url):
    d = request_detail(con, rid)
    a = next(x for x in d["approvals"] if x["level"] == level)
    n = len(d["approvals"])
    what = d["kind_label"]
    detail_label, detail = ("Project", project_text(con, d["project_id"])) if d["kind"] == "project" else ("Milestones", summary_text(d))
    prior = "".join(f"  L{x['level']} {who(x)}: approved by {x['decided_by']} on {x['decided_at']}"
                    + (f" - {x['comment']}" if x["comment"] else "") + "\n" for x in d["approvals"] if x["level"] < level)
    emailer(con, a["approver"], f"[MTA] {what} - approval needed (Level {level} of {n}): {d['project_id']} {d['project_name']}",
            f"Dear {who(a)},\n\n{what} for {d['project_id']} {d['project_name']} needs your approval (Level {level} of {n}).\n\n"
            f"Requested by: {d['requested_by']} on {d['created_at']}\n\n{detail_label}:\n{detail}\n\n"
            + (f"Already approved:\n{prior}\n" if prior else "") + f"Approve or reject here:\n{url}/baseline/{a['token']}\n", None)


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
    new_ids = [i["milestone_id"] for i in d["items"] if d["kind"] == "addition" or i["old_baseline"] is None]
    everyone = ", ".join(dict.fromkeys([x["approver"] for x in d["approvals"] if x["approver"]]
                                       + ([d["requester_email"]] if d["requester_email"] else [])))
    what = d["kind_label"]
    if decision == "Rejected":
        con.execute("UPDATE baseline_approvals SET status='Cancelled' WHERE request_id=? AND status='Waiting'", (a["request_id"],))
        con.execute("UPDATE baseline_requests SET status=?, closed_at=? WHERE id=?", (f"Rejected at L{a['level']}", now(), a["request_id"]))
        for mid in new_ids:
            con.execute("UPDATE milestones SET baseline_status='Rejected', rebase_req=NULL, status_note=? WHERE id=?",
                        (f"Rejected at Level {a['level']} by {name.strip()}: {comment.strip()}", mid))
        if d["kind"] == "project":
            con.execute("UPDATE projects SET status='Rejected' WHERE id=?", (d["project_id"],))
        detail = project_text(con, d["project_id"]) if d["kind"] == "project" else f"Milestones (not applied):\n{summary_text(d)}\n"
        emailer(con, everyone, f"[MTA] {what} rejected at Level {a['level']}: {d['project_id']} {d['project_name']}",
                f"{name.strip()} (Level {a['level']}) rejected the request from {d['requested_by']}.\nReason: {comment.strip()}\n\n"
                f"{detail}\n", None)
    elif a["level"] < n:
        nxt = a["level"] + 1
        con.execute("UPDATE baseline_approvals SET status='Pending' WHERE request_id=? AND level=?", (a["request_id"], nxt))
        con.execute("UPDATE baseline_requests SET status=?, current_level=? WHERE id=?", (f"Pending L{nxt}", nxt, a["request_id"]))
        for mid in new_ids:
            con.execute("UPDATE milestones SET status_note=? WHERE id=?", (f"Awaiting Level {nxt} approval", mid))
        notify_level(con, a["request_id"], nxt, emailer, url)
    else:
        pr = con.execute("SELECT baseline_no FROM projects WHERE id=?", (d["project_id"],)).fetchone()
        bno = (pr["baseline_no"] or 0) + (1 if d["kind"] in ("baseline", "rebaseline") else 0)
        for i in d["items"]:
            if d["kind"] in ("baseline", "addition") or i["old_baseline"] is None:
                con.execute("UPDATE milestones SET baseline_status='Approved', rebase_req=NULL, status_note=? WHERE id=?",
                            (f"Added to Baseline {bno} on {now()}", i["milestone_id"]))
            else:
                con.execute("UPDATE milestones SET baseline=? WHERE id=?", (i["new_baseline"], i["milestone_id"]))
                con.execute("INSERT INTO baseline_history VALUES(?,?,?,?,?)", (i["milestone_id"], i["old_baseline"], i["new_baseline"],
                                                                                a["request_id"], now()))
        con.execute("UPDATE projects SET baseline_no=? WHERE id=?", (bno, d["project_id"]))
        con.execute("UPDATE baseline_requests SET status='Approved', closed_at=?, baseline_no=? WHERE id=?", (now(), bno, a["request_id"]))
        if d["kind"] == "project":
            con.execute("UPDATE projects SET status='Approved' WHERE id=?", (d["project_id"],))
            detail = f"The project is now approved and ready to have its baseline set up:\n{project_text(con, d['project_id'])}"
        elif d["kind"] in ("baseline", "addition"):
            detail = f"These milestones are now in Baseline {bno}:\n{summary_text(d)}"
        else:
            detail = f"The new baselines are now in effect (Baseline {bno}):\n{summary_text(d)}"
        emailer(con, everyone, f"[MTA] {what} approved: {d['project_id']} {d['project_name']}",
                f"All {n} approval level(s) approved. {detail}\n", None)
        on_applied()
    con.commit()
    return request_detail(con, a["request_id"])


# ----------------------------------------------------------------------------- monthly report
def save_monthly(con, payload, now, emailer, url):
    """payload: project_id, period, submitted_by, submitter_email, rows[], baseline_changes[].
    Rebaseline requests go to the project's configured rebaseline workflow."""
    pid = payload["project_id"]
    pr = con.execute("SELECT * FROM projects WHERE id=?", (pid,)).fetchone()
    if not pr:
        return dict(errors=["Unknown project."])
    try:
        rep = period_date(payload["period"])
    except Exception:
        return dict(errors=["Choose a valid reporting period."])
    if pr["start_date"] and pr["end_date"] and not (pr["start_date"][:7] <= payload["period"] <= pr["end_date"][:7]):
        return dict(errors=[f"The reporting period must be between the project start ({pr['start_date'][:7]}) and end ({pr['end_date'][:7]})."])
    ms = {r["id"]: dict(r) for r in con.execute(
        "SELECT * FROM milestones WHERE project_id=? AND COALESCE(baseline_status,'Approved')='Approved'", (pid,))}
    errors, todo = [], []
    for i, r in enumerate(payload.get("rows", []), 1):
        mid = r.get("milestone_id")
        m = ms.get(mid)
        if not m:
            errors.append(f"Row {i}: {mid} is not an approved baseline milestone of this project")
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
    changes = []
    bc = [c for c in payload.get("baseline_changes", []) if c.get("milestone_id") or c.get("new_baseline")]
    if bc:
        if not chain(con, pid, "rebaseline"):
            errors.append("Baseline change: set the rebaseline approval workflow (Level 1 at least) on Project Setup first")
        if not (payload.get("submitted_by") or "").strip():
            errors.append("Baseline change: enter your name under Submitted by")
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
        req_id = create_request(con, pid, "rebaseline", changes, payload["submitted_by"].strip(),
                                (payload.get("submitter_email") or "").strip(), payload["period"], now, emailer, url)
    con.commit()
    return dict(saved=len(todo), errors=[], request_id=req_id, rep_date=rep.isoformat())
