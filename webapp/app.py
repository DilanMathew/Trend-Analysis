"""Milestone Trend Analysis web app with an escalation / approval workflow.

FastAPI + SQLite.  Environment variables:
  DATA_DIR         where mta.db lives (Railway volume mount, default ./data)
  ADMIN_PASSWORD   required for write actions (open if unset - local use only)
  PUBLIC_URL       base URL used in approval links (default: request URL)
  RESEND_API_KEY   send email through Resend (HTTPS)       } if neither is set,
  SMTP_HOST/SMTP_PORT/SMTP_USER/SMTP_PASSWORD               } emails stay in the
  EMAIL_FROM       sender address                           } Outbox (dry run)
  ESCALATION_CHECK_MINUTES  how often rules are re-evaluated (default 60)
"""
import datetime as dt
import html
import json
import os
import secrets
import smtplib
import sqlite3
import sys
import tempfile
import threading
import time
import urllib.request
from email.message import EmailMessage

from fastapi import Depends, FastAPI, File, Form, Header, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, Response
from pydantic import BaseModel

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "tools"))
sys.path.insert(0, os.path.join(HERE, "tools"))
import msp_mirror  # noqa: E402
import baseline as bl  # noqa: E402
import responsibility as rm  # noqa: E402
import sample_data  # noqa: E402

DATA_DIR = os.environ.get("DATA_DIR", os.path.join(HERE, "data"))
os.makedirs(DATA_DIR, exist_ok=True)
DB_PATH = os.path.join(DATA_DIR, "mta.db")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "")
TOL = 3
LOCK = threading.RLock()

app = FastAPI(title="Milestone Trend Analysis")


# ----------------------------------------------------------------------------- db
def db():
    con = sqlite3.connect(DB_PATH, check_same_thread=False)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA foreign_keys=ON")
    return con


SCHEMA = """
CREATE TABLE IF NOT EXISTS projects(id TEXT PRIMARY KEY, name TEXT NOT NULL, freq TEXT);
CREATE TABLE IF NOT EXISTS milestones(id TEXT PRIMARY KEY, project_id TEXT NOT NULL REFERENCES projects(id),
  name TEXT NOT NULL, category TEXT, baseline TEXT NOT NULL, owner TEXT, priority TEXT, description TEXT);
CREATE TABLE IF NOT EXISTS observations(id INTEGER PRIMARY KEY AUTOINCREMENT, milestone_id TEXT NOT NULL REFERENCES milestones(id),
  rep_date TEXT NOT NULL, forecast TEXT NOT NULL, actual TEXT, comment TEXT, source TEXT, created_at TEXT NOT NULL,
  UNIQUE(milestone_id, rep_date));
CREATE TABLE IF NOT EXISTS rules(id INTEGER PRIMARY KEY AUTOINCREMENT, level INTEGER NOT NULL, name TEXT NOT NULL,
  min_delay INTEGER NOT NULL, recipients TEXT NOT NULL, active INTEGER NOT NULL DEFAULT 1);
CREATE TABLE IF NOT EXISTS escalations(id INTEGER PRIMARY KEY AUTOINCREMENT, milestone_id TEXT NOT NULL, rule_id INTEGER,
  level INTEGER, rule_name TEXT, recipients TEXT, delay_days INTEGER, baseline TEXT, forecast TEXT, rep_date TEXT,
  status TEXT NOT NULL, token TEXT UNIQUE, created_at TEXT, decided_at TEXT, decided_by TEXT, decision TEXT, comment TEXT,
  current_delay INTEGER);
CREATE TABLE IF NOT EXISTS outbox(id INTEGER PRIMARY KEY AUTOINCREMENT, created_at TEXT, to_addr TEXT, subject TEXT, body TEXT,
  status TEXT, error TEXT, escalation_id INTEGER);
"""


def now():
    return dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%d %H:%M UTC")


def d(s):
    return dt.date.fromisoformat(s) if s else None


DATA_TABLES = ("baseline_items", "baseline_approvals", "baseline_requests", "baseline_history", "project_approvers",
              "reminders", "outbox", "escalations", "observations", "milestones", "projects")
DATA_VERSION = "2"   # v2: no sample data - projects and milestones are created by the user


def clear_data(con):
    for t in DATA_TABLES:
        con.execute(f"DELETE FROM {t}")


def default_rules(con):
    con.execute("DELETE FROM rules")
    con.executemany("INSERT INTO rules(level,name,min_delay,recipients,active,route_discipline) VALUES(?,?,?,?,1,1)", [
        (1, "Level 1 escalation (1 week late)", 7, ""),
        (2, "Level 2 escalation (1 month late)", 30, ""),
    ])


def seed(con):
    for m in ("ALTER TABLE projects ADD COLUMN start_date TEXT", "ALTER TABLE projects ADD COLUMN end_date TEXT",
              "ALTER TABLE projects ADD COLUMN manager TEXT"):
        try:
            con.execute(m)
        except Exception:
            pass
    con.execute("CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY, value TEXT)")
    ver = con.execute("SELECT value FROM settings WHERE key='data_version'").fetchone()
    if not ver or ver[0] != DATA_VERSION:
        # one-time removal of the earlier sample data, sample responsibility matrix and placeholder rule emails
        clear_data(con)
        con.execute("DELETE FROM disciplines")
        default_rules(con)
        con.execute("INSERT OR REPLACE INTO settings VALUES('data_version', ?)", (DATA_VERSION,))
    if con.execute("SELECT COUNT(*) FROM rules").fetchone()[0] == 0:
        default_rules(con)
    con.commit()


with LOCK:
    _c = db()
    _c.executescript(SCHEMA)
    rm.setup(_c)
    bl.setup(_c)
    seed(_c)
    _c.close()


# ----------------------------------------------------------------------------- auth
def require_admin(x_admin_password: str = Header(default="")):
    if ADMIN_PASSWORD and not secrets.compare_digest(x_admin_password, ADMIN_PASSWORD):
        raise HTTPException(401, "Admin password required")
    return True


# ----------------------------------------------------------------------------- analysis
def milestone_state(con):
    """Latest state of every milestone (all observations, i.e. 'as of latest report')."""
    out = {}
    obs = con.execute("SELECT * FROM observations ORDER BY milestone_id, rep_date").fetchall()
    by = {}
    for o in obs:
        by.setdefault(o["milestone_id"], []).append(o)
    for m in con.execute("SELECT m.*, p.name AS project_name FROM milestones m JOIN projects p ON p.id=m.project_id"):
        os_ = by.get(m["id"], [])
        st = dict(m=dict(m), n=len(os_))
        if os_:
            cur = os_[-1]
            actual = max((o["actual"] for o in os_ if o["actual"]), default=None)
            base = d(m["baseline"])
            st.update(cur=cur["forecast"], rep=cur["rep_date"], actual=actual,
                      delay=((d(actual) if actual else d(cur["forecast"])) - base).days)
            st["status"] = "Completed" if actual else ("Delayed" if st["delay"] > TOL else "Ahead of Schedule" if st["delay"] < -TOL else "On Schedule")
        else:
            st["status"] = "Not Yet Started"
        out[m["id"]] = st
    return out


# ----------------------------------------------------------------------------- email
def email_mode():
    if os.environ.get("RESEND_API_KEY"):
        return "resend"
    if os.environ.get("SMTP_HOST"):
        return "smtp"
    return "dry-run"


def deliver(con, to, subject, body, esc_id=None):
    mode = email_mode()
    status, err = ("dry-run (not sent)", None)
    if not (to or "").strip():
        con.execute("INSERT INTO outbox(created_at,to_addr,subject,body,status,error,escalation_id) VALUES(?,?,?,?,?,?,?)",
                    (now(), "", subject, body, "not sent", "No recipients: add emails to the escalation rule or the responsibility matrix.", esc_id))
        return
    sender = os.environ.get("EMAIL_FROM", "MTA Escalations <onboarding@resend.dev>")
    try:
        if mode == "resend":
            req = urllib.request.Request("https://api.resend.com/emails", method="POST",
                                         data=json.dumps({"from": sender, "to": [a.strip() for a in to.split(",")],
                                                          "subject": subject, "text": body}).encode(),
                                         headers={"Authorization": "Bearer " + os.environ["RESEND_API_KEY"],
                                                  "Content-Type": "application/json"})
            urllib.request.urlopen(req, timeout=20).read()
            status = "sent (Resend)"
        elif mode == "smtp":
            msg = EmailMessage()
            msg["From"], msg["To"], msg["Subject"] = sender, to, subject
            msg.set_content(body)
            with smtplib.SMTP(os.environ["SMTP_HOST"], int(os.environ.get("SMTP_PORT", "587")), timeout=20) as s:
                s.starttls()
                if os.environ.get("SMTP_USER"):
                    s.login(os.environ["SMTP_USER"], os.environ.get("SMTP_PASSWORD", ""))
                s.send_message(msg)
            status = "sent (SMTP)"
    except Exception as e:  # keep the record, show the failure on the Outbox page
        status, err = "failed", str(e)[:500]
    con.execute("INSERT INTO outbox(created_at,to_addr,subject,body,status,error,escalation_id) VALUES(?,?,?,?,?,?,?)",
                (now(), to, subject, body, status, err, esc_id))


def base_url(request: Request = None):
    if os.environ.get("PUBLIC_URL"):
        return os.environ["PUBLIC_URL"].rstrip("/")
    if request is None:
        return LAST_BASE["url"]
    proto = request.headers.get("x-forwarded-proto", request.url.scheme)
    return f"{proto}://{request.headers.get('host', request.url.netloc)}"


LAST_BASE = {"url": os.environ.get("PUBLIC_URL", "").rstrip("/")}


def fmt(s):
    return d(s).strftime("%d-%b-%Y") if s else "-"


def escalation_email(e, st, url):
    m = st["m"]
    link = f"{url}/approve/{e['token']}"
    body = (f"Milestone escalation - Level {e['level']}: {e['rule_name']}\n\n"
            f"Project:            {m['project_id']} {m['project_name']}\n"
            f"Milestone:          {m['id']} {m['name']}\n"
            f"Owner:              {m['owner'] or '-'}\n"
            f"Discipline:         {m['category'] or '-'}" + (f"  (responsible: {e['disc']['person']})" if e.get('disc') else "") + "\n"
            f"Original baseline:  {fmt(m['baseline'])}\n"
            f"Current forecast:   {fmt(st['cur'])}  (reported {fmt(st['rep'])})\n"
            f"Delay vs baseline:  {st['delay']} days  (threshold {e['threshold']} days)\n\n"
            f"Please review and approve the revised date or reject it and request a recovery plan:\n{link}\n\n"
            f"Dashboard: {url}/\n")
    subj = f"[MTA] L{e['level']} escalation: {m['id']} {m['name']} is {st['delay']} days late"
    return subj, body


def evaluate(con, url=None):
    """Create / auto-close escalations from the active rules. Returns a summary."""
    url = url or LAST_BASE["url"] or "(set PUBLIC_URL)"
    rules = [dict(r) for r in con.execute("SELECT * FROM rules WHERE active=1 ORDER BY min_delay DESC")]
    state = milestone_state(con)
    created = closed = 0
    for mid, st in state.items():
        open_esc = con.execute("SELECT * FROM escalations WHERE milestone_id=? AND status='Pending approval' AND COALESCE(kind,'delay')='delay'", (mid,)).fetchall()
        if st["status"] == "Completed" or st["n"] == 0:
            for e in open_esc:
                con.execute("UPDATE escalations SET status='Closed', decision='Auto-closed: milestone completed', decided_at=? WHERE id=?", (now(), e["id"]))
                closed += 1
            continue
        delay = st["delay"]
        for e in open_esc:
            rule = next((r for r in rules if r["id"] == e["rule_id"]), None)
            if rule is None or delay < rule["min_delay"]:
                con.execute("UPDATE escalations SET status='Closed', decision='Auto-closed: delay back under threshold', decided_at=?, current_delay=? WHERE id=?",
                            (now(), delay, e["id"]))
                closed += 1
            else:
                con.execute("UPDATE escalations SET current_delay=? WHERE id=?", (delay, e["id"]))
                if not (e["recipients"] or "").strip():   # contacts added after the escalation was raised
                    recips = [x.strip() for x in rule["recipients"].split(",") if x.strip()]
                    if rule.get("route_discipline", 1):
                        recips += rm.contacts_for(con, st["m"]["category"], rule["level"])[0]
                    recips = ", ".join(dict.fromkeys(recips))
                    if recips:
                        con.execute("UPDATE escalations SET recipients=? WHERE id=?", (recips, e["id"]))
                        subj, body = escalation_email(dict(level=rule["level"], rule_name=rule["name"], token=e["token"],
                                                           threshold=rule["min_delay"]), st, url)
                        deliver(con, recips, subj, body, e["id"])
        hit = next((r for r in rules if delay >= r["min_delay"]), None)  # highest threshold reached
        if not hit:
            continue
        exists = con.execute("SELECT 1 FROM escalations WHERE milestone_id=? AND rule_id=? AND status IN ('Pending approval','Approved','Rejected')",
                             (mid, hit["id"])).fetchone()
        if exists:
            continue
        # a new, higher level supersedes lower pending ones
        for e in con.execute("SELECT * FROM escalations WHERE milestone_id=? AND status='Pending approval'", (mid,)).fetchall():
            con.execute("UPDATE escalations SET status='Superseded', decision=?, decided_at=? WHERE id=?",
                        (f"Escalated to level {hit['level']}", now(), e["id"]))
        token = secrets.token_urlsafe(18)
        recips = [x.strip() for x in hit["recipients"].split(",") if x.strip()]
        disc = None
        if hit.get("route_discipline", 1):
            extra, disc = rm.contacts_for(con, st["m"]["category"], hit["level"])
            recips += extra
        recips = ", ".join(dict.fromkeys(recips))
        cur = con.execute("""INSERT INTO escalations(milestone_id,rule_id,level,rule_name,recipients,delay_days,baseline,forecast,rep_date,
                             status,token,created_at,current_delay,kind,discipline) VALUES(?,?,?,?,?,?,?,?,?,'Pending approval',?,?,?,'delay',?)""",
                          (mid, hit["id"], hit["level"], hit["name"], recips, delay, st["m"]["baseline"], st["cur"], st["rep"],
                           token, now(), delay, disc["discipline"] if disc else st["m"]["category"]))
        subj, body = escalation_email(dict(level=hit["level"], rule_name=hit["name"], token=token, threshold=hit["min_delay"],
                                           disc=disc), st, url)
        deliver(con, recips, subj, body, cur.lastrowid)
        created += 1
    upd = rm.evaluate_updates(con, deliver, now, url, fmt)
    con.commit()
    return dict(created=created + upd["created"], closed=closed + upd["closed"], reminded=upd["reminded"])


def scheduler():
    every = max(5, int(os.environ.get("ESCALATION_CHECK_MINUTES", "60")))
    while True:
        time.sleep(every * 60)
        try:
            with LOCK:
                c = db()
                evaluate(c)
                c.close()
        except Exception as e:  # never kill the thread
            print("scheduler error:", e, flush=True)


@app.on_event("startup")
def startup():
    with LOCK:
        c = db()
        evaluate(c)
        c.close()
    threading.Thread(target=scheduler, daemon=True).start()


@app.middleware("http")
async def remember_base(request: Request, call_next):
    if not LAST_BASE["url"]:
        LAST_BASE["url"] = base_url(request)
    return await call_next(request)


# ----------------------------------------------------------------------------- read API
@app.get("/api/data")
def api_data():
    with LOCK:
        c = db()
        P = [[r["id"], r["name"], r["freq"] or "", r["start_date"], r["end_date"], r["manager"] or ""]
             for r in c.execute("SELECT * FROM projects ORDER BY id")]
        M = [[r["project_id"], r["id"], r["name"], r["category"] or "", r["baseline"], r["owner"] or "", r["priority"] or "",
              r["description"] or ""] for r in c.execute("SELECT * FROM milestones ORDER BY project_id, baseline, id")]
        O = [[r["milestone_id"], r["rep_date"], r["forecast"], r["actual"], r["comment"], r["id"], r["progress"]]
             for r in c.execute("SELECT * FROM observations ORDER BY id")]
        pending = [r[0] for r in c.execute("""SELECT i.milestone_id FROM baseline_items i JOIN baseline_requests q ON q.id=i.request_id
                                              WHERE q.status LIKE 'Pending%'""")]
        appr = {}
        for r in c.execute("SELECT * FROM project_approvers ORDER BY project_id, level"):
            appr.setdefault(r["project_id"], []).append(r["approver"])
        hist = {}
        for r in c.execute("SELECT * FROM baseline_history ORDER BY changed_at"):
            hist.setdefault(r["milestone_id"], []).append([r["old_baseline"], r["new_baseline"], r["changed_at"]])
        c.close()
    return dict(P=P, M=M, O=O, pending_baseline=pending, approvers=appr, baseline_history=hist)


@app.get("/api/config")
def api_config():
    return dict(email_mode=email_mode(), admin_required=bool(ADMIN_PASSWORD), tolerance=TOL,
                check_minutes=int(os.environ.get("ESCALATION_CHECK_MINUTES", "60")))


@app.get("/api/escalations")
def api_escalations():
    with LOCK:
        c = db()
        rows = [dict(r) for r in c.execute("""SELECT e.*, m.name AS milestone_name, m.project_id, m.owner
                                               FROM escalations e LEFT JOIN milestones m ON m.id=e.milestone_id ORDER BY e.id DESC""")]
        c.close()
    for r in rows:
        r.pop("token", None)
    return rows


@app.get("/api/outbox")
def api_outbox():
    with LOCK:
        c = db()
        rows = [dict(r) for r in c.execute("SELECT * FROM outbox ORDER BY id DESC LIMIT 200")]
        c.close()
    return rows


@app.get("/api/rules")
def api_rules():
    with LOCK:
        c = db()
        rows = [dict(r) for r in c.execute("SELECT * FROM rules ORDER BY min_delay")]
        c.close()
    return rows


# ----------------------------------------------------------------------------- write API
class Rule(BaseModel):
    id: int | None = None
    level: int
    name: str
    min_delay: int
    recipients: str
    active: bool = True
    route_discipline: bool = True


def valid_emails(s):
    parts = [p.strip() for p in s.split(",") if p.strip()]
    return parts and all("@" in p and "." in p.split("@")[-1] and " " not in p for p in parts)


@app.put("/api/rules")
def put_rules(rules: list[Rule], request: Request, _=Depends(require_admin)):
    for r in rules:
        if r.min_delay < 1 or not r.name.strip():
            raise HTTPException(400, f"Rule '{r.name}': give it a name and a delay of at least 1 day.")
        if r.recipients.strip() and not valid_emails(r.recipients):
            raise HTTPException(400, f"Rule '{r.name}': recipients must be valid email addresses, separated by commas.")
        if not r.recipients.strip() and not r.route_discipline:
            raise HTTPException(400, f"Rule '{r.name}': add recipients or tick 'Notify discipline contacts', otherwise nobody is told.")
    if len({r.min_delay for r in rules}) != len(rules):
        raise HTTPException(400, "Two rules have the same delay threshold.")
    with LOCK:
        c = db()
        keep = [r.id for r in rules if r.id]
        c.execute(f"UPDATE rules SET active=0 WHERE id NOT IN ({','.join('?' * len(keep)) or 'NULL'})", keep)
        for r in rules:
            if r.id:
                c.execute("UPDATE rules SET level=?,name=?,min_delay=?,recipients=?,active=?,route_discipline=? WHERE id=?",
                          (r.level, r.name.strip(), r.min_delay, r.recipients.strip(), int(r.active), int(r.route_discipline), r.id))
            else:
                c.execute("INSERT INTO rules(level,name,min_delay,recipients,active,route_discipline) VALUES(?,?,?,?,?,?)",
                          (r.level, r.name.strip(), r.min_delay, r.recipients.strip(), int(r.active), int(r.route_discipline)))
        c.commit()
        res = evaluate(c, base_url(request))
        c.close()
    return dict(ok=True, **res)


class Obs(BaseModel):
    milestone_id: str
    forecast: str
    actual: str | None = None
    comment: str | None = None


class Snapshot(BaseModel):
    rep_date: str
    observations: list[Obs]


@app.post("/api/snapshot")
def post_snapshot(s: Snapshot, request: Request, _=Depends(require_admin)):
    try:
        rep = dt.date.fromisoformat(s.rep_date)
    except ValueError:
        raise HTTPException(400, "Reporting date must be a valid date.")
    with LOCK:
        c = db()
        added, errors = 0, []
        for o in s.observations:
            if not c.execute("SELECT 1 FROM milestones WHERE id=?", (o.milestone_id,)).fetchone():
                errors.append(f"{o.milestone_id}: unknown milestone")
                continue
            if c.execute("SELECT 1 FROM observations WHERE milestone_id=? AND rep_date=?", (o.milestone_id, rep.isoformat())).fetchone():
                errors.append(f"{o.milestone_id}: already recorded for {rep:%d-%b-%Y}")
                continue
            try:
                fc = dt.date.fromisoformat(o.forecast)
                ac = dt.date.fromisoformat(o.actual) if o.actual else None
            except ValueError:
                errors.append(f"{o.milestone_id}: invalid date")
                continue
            if ac and ac > rep:
                errors.append(f"{o.milestone_id}: actual date is after the reporting date")
                continue
            c.execute("INSERT INTO observations(milestone_id,rep_date,forecast,actual,comment,source,created_at) VALUES(?,?,?,?,?,?,?)",
                      (o.milestone_id, rep.isoformat(), fc.isoformat(), ac.isoformat() if ac else None, o.comment or None, "web form", now()))
            added += 1
        c.commit()
        res = evaluate(c, base_url(request))
        c.close()
    return dict(added=added, errors=errors, **res)


@app.post("/api/import/msp")
async def import_msp(request: Request, files: list[UploadFile] = File(...), date_culture: str = Form("en-GB"), _=Depends(require_admin)):
    tmp = tempfile.mkdtemp()
    for f in files:
        name = os.path.basename(f.filename or "upload.csv")
        with open(os.path.join(tmp, name), "wb") as fh:
            fh.write(await f.read())
    try:
        msp_mirror.CULTURE = date_culture
        obs, ms = msp_mirror.observations(tmp), msp_mirror.milestones(tmp)
    except Exception as e:
        raise HTTPException(400, f"Could not read the CSV files: {e}")
    with LOCK:
        c = db()
        new_p = new_m = added = skipped = 0
        for (pid, mid, name, desc, base, owner, dept, prio, cat) in ms:
            if not c.execute("SELECT 1 FROM projects WHERE id=?", (pid,)).fetchone():
                c.execute("INSERT INTO projects(id,name,freq) VALUES(?,?,?)", (pid, pid, "MS Project import"))
                new_p += 1
            if c.execute("SELECT 1 FROM milestones WHERE id=?", (mid,)).fetchone():
                c.execute("UPDATE milestones SET name=?,category=?,owner=?,priority=? WHERE id=?", (name, cat, owner, prio, mid))
            else:
                c.execute("INSERT INTO milestones VALUES(?,?,?,?,?,?,?,?)", (mid, pid, name, cat, base.isoformat(), owner, prio, desc))
                new_m += 1
        for (mid, rep, fc, act, cm) in obs:
            try:
                c.execute("INSERT INTO observations(milestone_id,rep_date,forecast,actual,comment,source,created_at) VALUES(?,?,?,?,?,?,?)",
                          (mid, rep.isoformat(), fc.isoformat(), act.isoformat() if act else None, cm, "MS Project CSV", now()))
                added += 1
            except sqlite3.IntegrityError:
                skipped += 1
        c.commit()
        res = evaluate(c, base_url(request))
        c.close()
    return dict(files=len(files), new_projects=new_p, new_milestones=new_m, added=added, skipped_existing=skipped, **res)


class Decision(BaseModel):
    decision: str
    name: str
    comment: str = ""


def decide(c, esc, dec: Decision, via):
    if esc["status"] != "Pending approval":
        raise HTTPException(409, f"This escalation is already {esc['status'].lower()}.")
    if dec.decision not in ("Approved", "Rejected", "Acknowledged") or not dec.name.strip():
        raise HTTPException(400, "Choose a decision and enter your name.")
    if dec.decision == "Rejected" and not dec.comment.strip():
        raise HTTPException(400, "Say what recovery action you need when rejecting.")
    c.execute("UPDATE escalations SET status=?, decision=?, decided_by=?, comment=?, decided_at=? WHERE id=?",
              (dec.decision, {"Approved": "Revised date accepted", "Rejected": "Recovery plan required",
                              "Acknowledged": "Missed update acknowledged"}[dec.decision] + f" ({via})",
               dec.name.strip(), dec.comment.strip(), now(), esc["id"]))
    m = c.execute("SELECT * FROM milestones WHERE id=?", (esc["milestone_id"],)).fetchone()
    what = f"{m['id']} {m['name']}" if m else f"{esc['discipline']} missed update"
    deliver(c, esc["recipients"], f"[MTA] {dec.decision}: {what} (L{esc['level']} escalation)",
            f"{dec.name.strip()} {dec.decision.lower()} the level {esc['level']} escalation for {what}.\n"
            f"Delay at escalation: {esc['delay_days']} days." + (f" Forecast: {fmt(esc['forecast'])}." if esc['forecast'] else "") + "\n"
            f"Comment: {dec.comment.strip() or '-'}\n", esc["id"])
    c.commit()


@app.post("/api/escalations/{eid}/decision")
def post_decision(eid: int, dec: Decision, _=Depends(require_admin)):
    with LOCK:
        c = db()
        esc = c.execute("SELECT * FROM escalations WHERE id=?", (eid,)).fetchone()
        if not esc:
            raise HTTPException(404, "Escalation not found")
        decide(c, esc, dec, "dashboard")
        c.close()
    return dict(ok=True)


@app.post("/api/escalations/run")
def run_now(request: Request, _=Depends(require_admin)):
    with LOCK:
        c = db()
        res = evaluate(c, base_url(request))
        c.close()
    return res


@app.post("/api/reset")
def reset(request: Request, _=Depends(require_admin)):
    with LOCK:
        c = db()
        clear_data(c)
        c.commit()
        res = evaluate(c, base_url(request))
        c.close()
    return dict(ok=True, **res)


@app.post("/api/login")
def login(_=Depends(require_admin)):
    return dict(ok=True)


# ----------------------------------------------------------------------------- approval page (email link)
PAGE = """<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Milestone escalation</title><style>
body{{font:15px/1.55 'Segoe UI',Arial,sans-serif;background:#F2F4F8;color:#16233B;margin:0;padding:24px 16px}}
.c{{max-width:620px;margin:0 auto;background:#fff;border:1px solid #D8DEE8;border-radius:10px;padding:22px}}
h1{{font-size:21px;margin:0 0 4px;color:#1F3864}} .m{{color:#5D6779;font-size:13px}} table{{width:100%;border-collapse:collapse;margin:14px 0}}
td{{padding:6px 4px;border-bottom:1px solid #E6ECF5}} td:first-child{{color:#5D6779;width:42%}}
label{{display:block;font-weight:600;font-size:13px;margin:12px 0 4px}} input,textarea{{width:100%;box-sizing:border-box;font:inherit;padding:8px;border:1px solid #C9D1DE;border-radius:6px}}
.b{{display:flex;gap:10px;margin-top:16px;flex-wrap:wrap}} button{{font:600 14px 'Segoe UI',Arial,sans-serif;padding:10px 16px;border-radius:7px;border:0;cursor:pointer;color:#fff}}
.ok{{background:#1E8449}} .no{{background:#C0392B}} .pill{{display:inline-block;padding:3px 10px;border-radius:99px;font-weight:600;font-size:13px;background:#E6ECF5}}
.late{{color:#C0392B;font-weight:700}}</style></head><body><div class="c">{content}</div></body></html>"""


def esc_view(c, token):
    esc = c.execute("SELECT * FROM escalations WHERE token=?", (token,)).fetchone()
    if not esc:
        raise HTTPException(404, "This approval link is not valid.")
    m = c.execute("SELECT m.*, p.name AS pname FROM milestones m JOIN projects p ON p.id=m.project_id WHERE m.id=?", (esc["milestone_id"],)).fetchone()
    if m is None:
        dr = c.execute("SELECT * FROM disciplines WHERE discipline=? COLLATE NOCASE", (esc["discipline"] or "",)).fetchone()
        m = dict(kind="missed", discipline=esc["discipline"], person=dr["person"] if dr else "-", email=dr["email"] if dr else "-",
                 cadence=rm.cadence_text(dr) if dr else "-", token=dr["token"] if dr else None)
    return esc, m


def approval_html(esc, m, msg=""):
    e = html.escape
    if isinstance(m, dict) and m.get("kind") == "missed":
        rows = [("Discipline", m["discipline"]), ("Responsible", f"{m['person']} <{m['email']}>"), ("Update cadence", m["cadence"]),
                ("Escalation", f"Level {esc['level']}: {esc['rule_name']}"), ("First missed due date", fmt(esc["period"])),
                ("Overdue when escalated", f"<span class='late'>{esc['delay_days']} days</span>"), ("Raised", esc["created_at"])]
        t = "".join(f"<tr><td>{a}</td><td>{b if a.startswith('Overdue') else e(str(b))}</td></tr>" for a, b in rows)
        if esc["status"] == "Pending approval":
            form = """<form method="post"><label for="n">Your name</label><input id="n" name="name" required>
            <label for="cm">Comment (optional)</label><textarea id="cm" name="comment" rows="3"></textarea>
            <div class="b"><button class="ok" name="decision" value="Acknowledged">Acknowledge - I will chase the update</button></div></form>"""
        else:
            form = f"<p><span class='pill'>{e(esc['status'])}</span> {e(esc['decision'] or '')} {('by ' + e(esc['decided_by'])) if esc['decided_by'] else ''}</p>"
        portal = f"<p><a href='/update/{m['token']}'>Open the {e(m['discipline'])} update page</a></p>" if m.get("token") else ""
        return PAGE.format(content=f"<h1>Missed milestone update</h1><div class='m'>Milestone Trend Analysis responsibility matrix</div>"
                                   f"{('<p><b>' + e(msg) + '</b></p>') if msg else ''}<table>{t}</table>{form}{portal}")
    rows = [("Project", f"{m['project_id']} {m['pname']}"), ("Milestone", f"{m['id']} {m['name']}"), ("Owner", m["owner"] or "-"),
            ("Escalation", f"Level {esc['level']}: {esc['rule_name']}"), ("Original baseline", fmt(esc["baseline"])),
            ("Forecast when escalated", f"{fmt(esc['forecast'])} (report {fmt(esc['rep_date'])})"),
            ("Delay vs baseline", f"<span class='late'>{esc['delay_days']} days</span>"), ("Raised", esc["created_at"])]
    t = "".join(f"<tr><td>{a}</td><td>{b if a == 'Delay vs baseline' else e(str(b))}</td></tr>" for a, b in rows)
    if esc["status"] == "Pending approval":
        form = f"""<form method="post"><label for="n">Your name</label><input id="n" name="name" required>
        <label for="cm">Comment (required when rejecting)</label><textarea id="cm" name="comment" rows="3"></textarea>
        <div class="b"><button class="ok" name="decision" value="Approved">Approve revised date</button>
        <button class="no" name="decision" value="Rejected">Reject - recovery plan required</button></div></form>"""
    else:
        form = f"<p><span class='pill'>{e(esc['status'])}</span> {e(esc['decision'] or '')} {('by ' + e(esc['decided_by'])) if esc['decided_by'] else ''} {e(esc['decided_at'] or '')}</p>" + \
               (f"<p class='m'>Comment: {e(esc['comment'])}</p>" if esc["comment"] else "")
    return PAGE.format(content=f"<h1>Milestone escalation</h1><div class='m'>Milestone Trend Analysis approval workflow</div>"
                               f"{('<p><b>' + e(msg) + '</b></p>') if msg else ''}<table>{t}</table>{form}"
                               f"<p class='m'><a href='/#escalations'>Open the dashboard</a></p>")


@app.get("/approve/{token}", response_class=HTMLResponse)
def approve_get(token: str):
    with LOCK:
        c = db()
        esc, m = esc_view(c, token)
        c.close()
    return approval_html(esc, m)


@app.post("/approve/{token}", response_class=HTMLResponse)
def approve_post(token: str, decision: str = Form(...), name: str = Form(""), comment: str = Form("")):
    with LOCK:
        c = db()
        esc, m = esc_view(c, token)
        try:
            decide(c, esc, Decision(decision=decision, name=name, comment=comment), "email link")
            msg = f"Thank you - your decision ({decision.lower()}) has been recorded."
        except HTTPException as ex:
            msg = ex.detail
        esc, m = esc_view(c, token)
        c.close()
    return approval_html(esc, m, msg)


# ----------------------------------------------------------------------------- projects & milestones (user maintained)
class ProjectIn(BaseModel):
    id: str
    name: str
    start_date: str
    end_date: str
    manager: str = ""


def _check_project(p):
    if not p.id.strip() or not p.name.strip():
        raise HTTPException(400, "Enter a Project ID and a Project Name.")
    try:
        sd, ed = dt.date.fromisoformat(p.start_date), dt.date.fromisoformat(p.end_date)
    except ValueError:
        raise HTTPException(400, "Enter valid project start and end dates.")
    if ed < sd:
        raise HTTPException(400, "The project end date is before the start date.")
    return sd, ed


@app.post("/api/projects")
def create_project(p: ProjectIn, _=Depends(require_admin)):
    sd, ed = _check_project(p)
    with LOCK:
        c = db()
        if c.execute("SELECT 1 FROM projects WHERE id=? COLLATE NOCASE", (p.id.strip(),)).fetchone():
            c.close()
            raise HTTPException(400, f"Project ID {p.id.strip()} already exists.")
        c.execute("INSERT INTO projects(id,name,freq,start_date,end_date,manager) VALUES(?,?,?,?,?,?)",
                  (p.id.strip(), p.name.strip(), "Monthly", sd.isoformat(), ed.isoformat(), p.manager.strip()))
        c.commit()
        c.close()
    return dict(ok=True)


@app.put("/api/projects/{pid}")
def update_project(pid: str, p: ProjectIn, request: Request, _=Depends(require_admin)):
    sd, ed = _check_project(p)
    with LOCK:
        c = db()
        if not c.execute("SELECT 1 FROM projects WHERE id=?", (pid,)).fetchone():
            c.close()
            raise HTTPException(404, "Project not found.")
        c.execute("UPDATE projects SET name=?, start_date=?, end_date=?, manager=? WHERE id=?",
                  (p.name.strip(), sd.isoformat(), ed.isoformat(), p.manager.strip(), pid))
        c.commit()
        c.close()
    return dict(ok=True)


@app.delete("/api/projects/{pid}")
def delete_project(pid: str, _=Depends(require_admin)):
    with LOCK:
        c = db()
        n = c.execute("SELECT COUNT(*) FROM milestones WHERE project_id=?", (pid,)).fetchone()[0]
        if n:
            c.close()
            raise HTTPException(400, f"Delete the project's {n} milestone(s) first.")
        c.execute("DELETE FROM projects WHERE id=?", (pid,))
        c.execute("DELETE FROM project_approvers WHERE project_id=?", (pid,))
        c.commit()
        c.close()
    return dict(ok=True)


class MilestoneIn(BaseModel):
    project_id: str
    id: str
    name: str
    category: str = ""
    baseline: str
    owner: str = ""
    priority: str = ""
    description: str = ""


def _check_milestone(c, m):
    if not m.id.strip() or not m.name.strip():
        raise HTTPException(400, "Enter a Milestone ID and a milestone name.")
    pr = c.execute("SELECT * FROM projects WHERE id=?", (m.project_id,)).fetchone()
    if not pr:
        raise HTTPException(400, "Choose a project.")
    try:
        b = dt.date.fromisoformat(m.baseline)
    except ValueError:
        raise HTTPException(400, f"{m.id}: enter a valid baseline date.")
    if pr["start_date"] and pr["end_date"] and not (pr["start_date"] <= b.isoformat() <= pr["end_date"]):
        raise HTTPException(400, f"{m.id}: the baseline date must fall between the project start ({fmt(pr['start_date'])}) and end ({fmt(pr['end_date'])}).")
    return b


@app.post("/api/milestones")
def create_milestone(m: MilestoneIn, _=Depends(require_admin)):
    with LOCK:
        c = db()
        try:
            b = _check_milestone(c, m)
            if c.execute("SELECT 1 FROM milestones WHERE id=? COLLATE NOCASE", (m.id.strip(),)).fetchone():
                raise HTTPException(400, f"Milestone ID {m.id.strip()} already exists. Each milestone needs a unique ID.")
            c.execute("INSERT INTO milestones VALUES(?,?,?,?,?,?,?,?)", (m.id.strip(), m.project_id, m.name.strip(), m.category.strip(),
                                                                         b.isoformat(), m.owner.strip(), m.priority.strip(), m.description.strip()))
            c.commit()
        finally:
            c.close()
    return dict(ok=True)


@app.put("/api/milestones/{mid}")
def update_milestone(mid: str, m: MilestoneIn, request: Request, _=Depends(require_admin)):
    with LOCK:
        c = db()
        try:
            cur = c.execute("SELECT * FROM milestones WHERE id=?", (mid,)).fetchone()
            if not cur:
                raise HTTPException(404, "Milestone not found.")
            m.project_id = cur["project_id"]
            b = _check_milestone(c, m)
            if b.isoformat() != cur["baseline"]:
                if c.execute("SELECT 1 FROM observations WHERE milestone_id=?", (mid,)).fetchone():
                    raise HTTPException(400, f"{mid}: forecasts have already been reported, so the baseline can only change through "
                                             "'If baseline to be changed' on the monthly report (L1/L2/L3 approval).")
            c.execute("UPDATE milestones SET name=?, category=?, baseline=?, owner=?, priority=?, description=? WHERE id=?",
                      (m.name.strip(), m.category.strip(), b.isoformat(), m.owner.strip(), m.priority.strip(), m.description.strip(), mid))
            c.commit()
            res = evaluate(c, base_url(request))
        finally:
            c.close()
    return dict(ok=True, **res)


@app.delete("/api/milestones/{mid}")
def delete_milestone(mid: str, _=Depends(require_admin)):
    with LOCK:
        c = db()
        try:
            n = c.execute("SELECT COUNT(*) FROM observations WHERE milestone_id=?", (mid,)).fetchone()[0]
            if n:
                raise HTTPException(400, f"{mid} already has {n} reported forecast(s); it cannot be deleted without losing its history.")
            c.execute("DELETE FROM escalations WHERE milestone_id=?", (mid,))
            c.execute("DELETE FROM milestones WHERE id=?", (mid,))
            c.commit()
        finally:
            c.close()
    return dict(ok=True)


# ----------------------------------------------------------------------------- monthly report + baseline approvals
@app.post("/api/monthly")
async def post_monthly(request: Request, _=Depends(require_admin)):
    payload = await request.json()
    url = base_url(request)
    with LOCK:
        c = db()
        res = bl.save_monthly(c, payload, now, deliver, url)
        if res.get("errors"):
            c.close()
            raise HTTPException(400, " | ".join(res["errors"]))
        ev = evaluate(c, url)
        c.close()
    return dict(**res, **ev)


@app.get("/api/baseline-requests")
def api_baseline_requests(x_admin_password: str = Header(default="")):
    admin = is_admin(x_admin_password)
    with LOCK:
        c = db()
        ids = [r[0] for r in c.execute("SELECT id FROM baseline_requests ORDER BY id DESC LIMIT 100")]
        out = [bl.request_detail(c, i) for i in ids]
        c.close()
    for d in out:
        for a in d["approvals"]:
            tok = a.pop("token")
            if admin and a["status"] == "Pending":
                a["link"] = f"/baseline/{tok}"
    return out


def baseline_page(d, a, msg=""):
    e = html.escape
    items = "".join(f"<tr><td><b>{e(i['milestone_id'])}</b> {e(i['milestone_name'] or '')}</td><td>{fmt(i['old_baseline'])}</td>"
                    f"<td><b>{fmt(i['new_baseline'])}</b></td><td class='late'>{(d_(i['new_baseline']) - d_(i['old_baseline'])).days:+d} d</td>"
                    f"<td>{e(i['reason'] or '')}</td></tr>" for i in d["items"])
    chain = "".join(f"<tr><td>L{x['level']}</td><td>{e(x['approver'])}</td><td><span class='pill'>{e(x['status'])}</span></td>"
                    f"<td>{e(x['decided_by'] or '')} {e(x['decided_at'] or '')}{(' - ' + e(x['comment'])) if x['comment'] else ''}</td></tr>"
                    for x in d["approvals"])
    if a["status"] == "Pending":
        form = f"""<form method="post"><label for="n">Your name</label><input id="n" name="name" required>
        <label for="cm">Comment (required when rejecting)</label><textarea id="cm" name="comment" rows="3"></textarea>
        <div class="b"><button class="ok" name="decision" value="Approved">Approve baseline change (L{a['level']})</button>
        <button class="no" name="decision" value="Rejected">Reject</button></div></form>"""
    else:
        form = f"<p><span class='pill'>Level {a['level']}: {e(a['status'])}</span></p>"
    content = (f"<h1>Baseline change approval - level {a['level']} of {len(d['approvals'])}</h1>"
               f"<div class='m'>{e(d['project_id'])} {e(d['project_name'])} · requested by {e(d['requested_by'] or '-')} on {e(d['created_at'] or '')}"
               f" · reporting period {e(d['period'] or '-')} · request status: <b>{e(d['status'])}</b></div>"
               f"{('<p><b>' + e(msg) + '</b></p>') if msg else ''}"
               f"<h3 style='margin:16px 0 4px'>Proposed changes</h3><table><tr><td><b>Milestone</b></td><td><b>Current baseline</b></td>"
               f"<td><b>New baseline</b></td><td><b>Change</b></td><td><b>Reason</b></td></tr>{items}</table>"
               f"<h3 style='margin:16px 0 4px'>Approval chain</h3><table>{chain}</table>{form}"
               f"<p class='m'>The new baselines take effect only after every level has approved. A rejection at any level stops the request.</p>"
               f"<p class='m'><a href='/#escalations'>Open the dashboard</a></p>")
    return PAGE.format(content=content).replace("max-width:620px", "max-width:820px")


def d_(s):
    return dt.date.fromisoformat(s)


@app.get("/baseline/{token}", response_class=HTMLResponse)
def baseline_get(token: str):
    with LOCK:
        c = db()
        a = c.execute("SELECT * FROM baseline_approvals WHERE token=?", (token,)).fetchone()
        if not a:
            c.close()
            raise HTTPException(404, "This approval link is not valid.")
        d = bl.request_detail(c, a["request_id"])
        c.close()
    return baseline_page(d, dict(a))


@app.post("/baseline/{token}", response_class=HTMLResponse)
def baseline_post(token: str, request: Request, decision: str = Form(...), name: str = Form(""), comment: str = Form("")):
    url = base_url(request)
    with LOCK:
        c = db()
        a = c.execute("SELECT * FROM baseline_approvals WHERE token=?", (token,)).fetchone()
        if not a:
            c.close()
            raise HTTPException(404, "This approval link is not valid.")
        try:
            bl.decide(c, token, decision, name, comment, now, deliver, url, lambda: evaluate(c, url))
            msg = f"Thank you - your decision ({decision.lower()}) has been recorded."
        except ValueError as ex:
            msg = str(ex)
        a = c.execute("SELECT * FROM baseline_approvals WHERE token=?", (token,)).fetchone()
        d = bl.request_detail(c, a["request_id"])
        c.close()
    return baseline_page(d, dict(a), msg)


# ----------------------------------------------------------------------------- responsibility matrix
def is_admin(pw):
    return not ADMIN_PASSWORD or secrets.compare_digest(pw or "", ADMIN_PASSWORD)


@app.get("/api/matrix")
def api_matrix(request: Request, x_admin_password: str = Header(default="")):
    admin = is_admin(x_admin_password)
    url = base_url(request)
    with LOCK:
        c = db()
        rows = rm.matrix_rows(c, active_only=False)
        comp = {r["id"]: rm.compliance(c, r) for r in rows if r["active"]}
        cats = [x[0] for x in c.execute("SELECT DISTINCT category FROM milestones WHERE category IS NOT NULL AND category<>''")]
        c.close()
    known = {r["discipline"].lower() for r in rows if r["active"]}
    for r in rows:
        r["compliance"] = comp.get(r["id"])
        r["cadence"] = rm.cadence_text(r)
        if admin:
            r["portal"] = f"{url}/update/{r['token']}"
        r.pop("token")
    return dict(rows=rows, unassigned=sorted(x for x in cats if x.lower() not in known), admin=admin)


class Discipline(BaseModel):
    id: int | None = None
    discipline: str
    person: str
    email: str
    frequency: str
    due_day: int
    grace_days: int = 3
    l1_email: str = ""
    l2_email: str = ""
    active: bool = True


@app.put("/api/matrix")
def put_matrix(rows: list[Discipline], request: Request, _=Depends(require_admin)):
    seen = set()
    for r in rows:
        nm = r.discipline.strip()
        if not nm or not r.person.strip():
            raise HTTPException(400, "Every row needs a discipline and a responsible person.")
        if nm.lower() in seen:
            raise HTTPException(400, f"Discipline '{nm}' is listed twice.")
        seen.add(nm.lower())
        if not valid_emails(r.email) or (r.l1_email and not valid_emails(r.l1_email)) or (r.l2_email and not valid_emails(r.l2_email)):
            raise HTTPException(400, f"{nm}: check the email addresses (separate several with commas).")
        if r.frequency not in ("Weekly", "Monthly"):
            raise HTTPException(400, f"{nm}: frequency must be Weekly or Monthly.")
        if (r.frequency == "Weekly" and not 1 <= r.due_day <= 7) or (r.frequency == "Monthly" and not 1 <= r.due_day <= 31):
            raise HTTPException(400, f"{nm}: due day must be a weekday (1 = Monday ... 7 = Sunday) or a day of the month (1-31).")
        if not 0 <= r.grace_days <= 30:
            raise HTTPException(400, f"{nm}: grace days must be between 0 and 30.")
    with LOCK:
        c = db()
        keep = [r.id for r in rows if r.id]
        c.execute(f"UPDATE disciplines SET active=0 WHERE id NOT IN ({','.join('?' * len(keep)) or 'NULL'})", keep)
        for r in rows:
            vals = (r.discipline.strip(), r.person.strip(), r.email.strip(), r.frequency, r.due_day, r.grace_days,
                    r.l1_email.strip(), r.l2_email.strip(), int(r.active))
            if r.id:
                c.execute("""UPDATE disciplines SET discipline=?,person=?,email=?,frequency=?,due_day=?,grace_days=?,l1_email=?,l2_email=?,active=?
                             WHERE id=?""", vals + (r.id,))
            else:
                c.execute("""INSERT INTO disciplines(discipline,person,email,frequency,due_day,grace_days,l1_email,l2_email,active,token)
                             VALUES(?,?,?,?,?,?,?,?,?,?)""", vals + (secrets.token_urlsafe(16),))
        c.commit()
        res = evaluate(c, base_url(request))
        c.close()
    return dict(ok=True, **res)


def portal_row(c, token):
    r = c.execute("SELECT * FROM disciplines WHERE token=? AND active=1", (token,)).fetchone()
    if not r:
        raise HTTPException(404, "This update link is not valid. Ask the PMO for your discipline's current link.")
    return dict(r)


def save_updates(c, r, items, source):
    allowed = {m["id"] for m in rm.discipline_milestones(c, r["discipline"])}
    added, errors = 0, []
    for it in items:
        tag = f"Row {it['row']} " if it.get("row") else ""
        mid = it["milestone_id"]
        if it.get("error"):
            errors.append(f"{tag}{mid}: {it['error']}")
            continue
        if not it.get("forecast"):
            continue  # nothing entered for this milestone
        if mid not in allowed:
            errors.append(f"{tag}{mid}: not an open {r['discipline']} milestone")
            continue
        rep_ = it.get("rep_date")
        if not rep_:
            errors.append(f"{tag}{mid}: reporting date missing")
            continue
        if it.get("actual") and it["actual"] > rep_:
            errors.append(f"{tag}{mid}: actual date is after the reporting date")
            continue
        c.execute("""INSERT INTO observations(milestone_id,rep_date,forecast,actual,comment,source,created_at,progress) VALUES(?,?,?,?,?,?,?,?)
                     ON CONFLICT(milestone_id,rep_date) DO UPDATE SET forecast=excluded.forecast, actual=excluded.actual,
                       comment=excluded.comment, source=excluded.source, created_at=excluded.created_at, progress=excluded.progress""",
                  (mid, rep_.isoformat(), it["forecast"].isoformat(), it["actual"].isoformat() if it.get("actual") else None,
                   it.get("comment") or None, f"{source} ({r['person']})", now(), "Completed" if it.get("actual") else "In progress"))
        added += 1
    c.commit()
    return added, errors


PORTAL_CSS = """<style>body{font:15px/1.5 'Segoe UI',Arial,sans-serif;background:#F2F4F8;color:#16233B;margin:0;padding:24px 16px}
.c{max-width:1100px;margin:0 auto;display:grid;gap:16px}.card{background:#fff;border:1px solid #D8DEE8;border-radius:10px;padding:18px;overflow-x:auto}
h1{font-size:22px;margin:0;color:#1F3864}h2{font-size:16px;margin:0 0 8px}.m{color:#5D6779;font-size:13px}
table{border-collapse:collapse;width:100%;font-size:13.5px}th,td{padding:7px 8px;border-bottom:1px solid #E6ECF5;text-align:left;white-space:nowrap}
th{background:#E6ECF5;font-size:11.5px;text-transform:uppercase;letter-spacing:.05em;color:#5D6779}
input[type=date],input[type=text]{font:inherit;padding:5px 7px;border:1px solid #E2C35A;background:#FFF7D6;border-radius:6px}
button,.btn{font:600 14px 'Segoe UI',Arial,sans-serif;padding:9px 16px;border-radius:7px;border:0;cursor:pointer;color:#fff;background:#1F3864;text-decoration:none;display:inline-block}
.pill{display:inline-block;padding:2px 10px;border-radius:99px;font-weight:600;font-size:12.5px;color:#fff}
.Overdue{background:#C0392B}.Due{background:#C55A11}.Up{background:#1E8449}.No{background:#8A939F}
.msg{padding:10px 14px;border-radius:8px;background:#E2F0D9}.err{background:#F8D7DA}.row{display:flex;gap:12px;flex-wrap:wrap;align-items:center}</style>"""


def period_opts():
    t = rm.today()
    ks = []
    y, m = t.year, t.month
    for _ in range(13):
        ks.append(f"{y}-{m:02d}")
        y, m = (y - 1, 12) if m == 1 else (y, m - 1)
    return "".join(f"<option value='{k}'{' selected' if i == 0 else ''}>{dt.date(int(k[:4]), int(k[5:]), 1):%b-%Y} "
                   f"(report date {bl.period_date(k):%d-%b-%y})</option>" for i, k in enumerate(ks))


def portal_html(c, r, msg="", errors=()):
    e = html.escape
    comp = rm.compliance(c, r)
    ms = rm.discipline_milestones(c, r["discipline"])
    stale = set(comp.get("stale", []))
    status = comp["status"]
    extra = (f" since {fmt(comp['missed'][0])} ({comp['overdue_days']} days)" if status == "Overdue"
             else f" - please update by {fmt(comp['grace_until'])}" if status == "Due" else "")
    rows = "".join(
        f"<tr><td>{e(m['project_id'])}</td><td><b>{e(m['id'])}</b> {e(m['name'])}</td><td>{fmt(m['baseline'])}</td>"
        f"<td>{fmt(m['last_rep'])}{' <b style=color:#C0392B>needs update</b>' if m['id'] in stale else ''}</td><td>{fmt(m['last_fc'])}</td>"
        f"<td><input type='date' name='fc_{e(m['id'])}' value='{m['last_fc'] or ''}' aria-label='New forecast {e(m['id'])}'></td>"
        f"<td><input type='checkbox' name='done_{e(m['id'])}' aria-label='{e(m['id'])} achieved'></td>"
        f"<td><input type='text' name='cm_{e(m['id'])}' placeholder='Reason for any change' size='28'></td></tr>" for m in ms)
    msg_html = (f"<div class='msg'>{e(msg)}</div>" if msg else "") + \
               (f"<div class='msg err'><b>Not saved:</b><br>{'<br>'.join(e(x) for x in errors)}</div>" if errors else "")
    return f"""<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{e(r['discipline'])} milestone update</title>{PORTAL_CSS}</head><body><div class="c">
<div class="card"><h1>{e(r['discipline'])} - milestone date update</h1>
<p class="m">Responsible: <b>{e(r['person'])}</b> &lt;{e(r['email'])}&gt; · Cadence: {e(rm.cadence_text(r))} · Grace: {r['grace_days']} day(s) ·
Next due: {fmt(comp['next_due'])}</p>
<p><span class="pill {status.split()[0]}">{e(status)}</span> {e(extra)} · {comp.get('updated', 0)} of {comp['reporting']} reporting milestones current</p></div>
{msg_html}
<form class="card" method="post"><h2>Update dates online</h2>
<p class="m">Each submitted row adds a new forecast observation; earlier forecasts are kept. Leave a forecast unchanged if the date has not moved, but do submit it, so the update counts.</p>
<div class="row" style="margin:8px 0 12px"><label>Reporting period <select name="period" required>{period_opts()}</select></label>
<span class="m">Report date is the 30th of the month (last day for shorter months). Submitting again in the same month corrects that month.</span></div>
<table><thead><tr><th>Project</th><th>Milestone</th><th>Baseline</th><th>Last reported</th><th>Last forecast</th><th>New forecast</th><th>Achieved</th><th>Comment</th></tr></thead>
<tbody>{rows or '<tr><td colspan=8>No open milestones for this discipline.</td></tr>'}</tbody></table>
<p style="margin-top:14px"><button type="submit">Submit update</button></p></form>
<div class="card"><h2>Or use the input file</h2><p class="m">Download the Excel input file, fill the yellow columns, and upload it here.</p>
<div class="row"><a class="btn" href="/update/{e(r['token'])}/input.xlsx">Download input file (.xlsx)</a>
<form method="post" action="/update/{e(r['token'])}/upload" enctype="multipart/form-data" class="row">
<input type="file" name="file" accept=".xlsx,.csv" required aria-label="Completed input file"><button type="submit">Upload completed file</button></form></div></div>
<p class="m"><a href="/">Open the dashboard</a></p></div></body></html>"""


@app.get("/update/{token}", response_class=HTMLResponse)
def portal(token: str):
    with LOCK:
        c = db()
        r = portal_row(c, token)
        out = portal_html(c, r)
        c.close()
    return out


@app.post("/update/{token}", response_class=HTMLResponse)
async def portal_post(token: str, request: Request):
    form = await request.form()
    with LOCK:
        c = db()
        r = portal_row(c, token)
        items, errors = [], []
        try:
            rep_ = bl.period_date(form.get("period", ""))
        except Exception:
            rep_ = None
        for m in rm.discipline_milestones(c, r["discipline"]):
            v = form.get(f"fc_{m['id']}")
            if not v:
                continue
            fc = dt.date.fromisoformat(v)
            if v == m["last_fc"] and not form.get(f"cm_{m['id']}") and not form.get(f"done_{m['id']}") and m["last_rep"] == (rep_.isoformat() if rep_ else None):
                continue
            items.append(dict(milestone_id=m["id"], rep_date=rep_, forecast=fc, actual=fc if form.get(f"done_{m['id']}") else None,
                              comment=form.get(f"cm_{m['id']}") or None))
        added, errors = save_updates(c, r, items, "update page")
        res = evaluate(c, base_url(request))
        out = portal_html(c, r, f"Saved {added} forecast(s) for {fmt(rep_.isoformat()) if rep_ else '-'}. "
                                f"Escalation check: {res['created']} raised, {res['closed']} closed.", errors)
        c.close()
    return out


@app.get("/update/{token}/input.xlsx")
def portal_file(token: str):
    with LOCK:
        c = db()
        r = portal_row(c, token)
        data = rm.input_workbook(c, r)
        c.close()
    fn = f"MTA_input_{r['discipline'].replace(' ', '_')}_{rm.today():%Y-%m-%d}.xlsx"
    return Response(data, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    headers={"Content-Disposition": f'attachment; filename="{fn}"'})


@app.post("/update/{token}/upload", response_class=HTMLResponse)
async def portal_upload(token: str, request: Request, file: UploadFile = File(...)):
    data = await file.read()
    with LOCK:
        c = db()
        r = portal_row(c, token)
        try:
            items = rm.parse_input_file(file.filename or "upload.xlsx", data)
            added, errors = save_updates(c, r, items, "input file")
            res = evaluate(c, base_url(request))
            msg = f"Imported {file.filename}: {added} forecast(s) saved. Escalation check: {res['created']} raised, {res['closed']} closed."
        except Exception as ex:
            msg, errors = "", [str(ex)]
        out = portal_html(c, r, msg, errors)
        c.close()
    return out


# ----------------------------------------------------------------------------- front end
@app.get("/health")
def health():
    return {"ok": True}


@app.get("/")
def index():
    return FileResponse(os.path.join(HERE, "static", "index.html"))


@app.exception_handler(HTTPException)
async def http_err(request: Request, exc: HTTPException):
    if request.url.path.startswith(("/approve/", "/update/", "/baseline/")) and request.method == "GET":
        return HTMLResponse(PAGE.format(content=f"<h1>Link not valid</h1><p>{html.escape(str(exc.detail))}</p>"), status_code=exc.status_code)
    return JSONResponse({"detail": exc.detail}, status_code=exc.status_code)
