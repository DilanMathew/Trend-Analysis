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
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from pydantic import BaseModel

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "tools"))
sys.path.insert(0, os.path.join(HERE, "tools"))
import msp_mirror  # noqa: E402
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


def seed(con, reset=False):
    if reset:
        for t in ("outbox", "escalations", "observations", "milestones", "projects"):
            con.execute(f"DELETE FROM {t}")
    if con.execute("SELECT COUNT(*) FROM projects").fetchone()[0] == 0:
        projects, ms, obs = sample_data.build()
        con.executemany("INSERT INTO projects VALUES(?,?,?)", [(p["id"], p["name"], p["freq"]) for p in projects])
        con.executemany("INSERT INTO milestones VALUES(?,?,?,?,?,?,?,?)",
                        [(m["id"], m["project"], m["name"], m["cat"], m["base"].isoformat(), m["owner"], m["prio"], m["desc"]) for m in ms])
        con.executemany("INSERT INTO observations(milestone_id,rep_date,forecast,actual,comment,source,created_at) VALUES(?,?,?,?,?,?,?)",
                        [(o["mid"], o["rep"].isoformat(), o["fc"].isoformat(), o["actual"].isoformat() if o["actual"] else None,
                          o["comment"] or None, "sample", now()) for o in obs])
    if con.execute("SELECT COUNT(*) FROM rules").fetchone()[0] == 0:
        con.executemany("INSERT INTO rules(level,name,min_delay,recipients,active) VALUES(?,?,?,?,1)", [
            (1, "Project manager escalation (1 week late)", 7, "pm-escalations@example.com"),
            (2, "Director escalation (1 month late)", 30, "director-escalations@example.com"),
        ])
    con.commit()


with LOCK:
    _c = db()
    _c.executescript(SCHEMA)
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
        open_esc = con.execute("SELECT * FROM escalations WHERE milestone_id=? AND status='Pending approval'", (mid,)).fetchall()
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
        cur = con.execute("""INSERT INTO escalations(milestone_id,rule_id,level,rule_name,recipients,delay_days,baseline,forecast,rep_date,
                             status,token,created_at,current_delay) VALUES(?,?,?,?,?,?,?,?,?,'Pending approval',?,?,?)""",
                          (mid, hit["id"], hit["level"], hit["name"], hit["recipients"], delay, st["m"]["baseline"], st["cur"], st["rep"],
                           token, now(), delay))
        subj, body = escalation_email(dict(level=hit["level"], rule_name=hit["name"], token=token, threshold=hit["min_delay"]), st, url)
        deliver(con, hit["recipients"], subj, body, cur.lastrowid)
        created += 1
    con.commit()
    return dict(created=created, closed=closed)


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
        P = [[r["id"], r["name"], r["freq"] or ""] for r in c.execute("SELECT * FROM projects ORDER BY id")]
        M = [[r["project_id"], r["id"], r["name"], r["category"] or "", r["baseline"], r["owner"] or ""]
             for r in c.execute("SELECT * FROM milestones ORDER BY project_id, baseline, id")]
        O = [[r["milestone_id"], r["rep_date"], r["forecast"], r["actual"], r["comment"], r["id"]]
             for r in c.execute("SELECT * FROM observations ORDER BY id")]
        c.close()
    return dict(P=P, M=M, O=O)


@app.get("/api/config")
def api_config():
    return dict(email_mode=email_mode(), admin_required=bool(ADMIN_PASSWORD), tolerance=TOL,
                check_minutes=int(os.environ.get("ESCALATION_CHECK_MINUTES", "60")))


@app.get("/api/escalations")
def api_escalations():
    with LOCK:
        c = db()
        rows = [dict(r) for r in c.execute("""SELECT e.*, m.name AS milestone_name, m.project_id, m.owner
                                               FROM escalations e JOIN milestones m ON m.id=e.milestone_id ORDER BY e.id DESC""")]
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


def valid_emails(s):
    parts = [p.strip() for p in s.split(",") if p.strip()]
    return parts and all("@" in p and "." in p.split("@")[-1] and " " not in p for p in parts)


@app.put("/api/rules")
def put_rules(rules: list[Rule], request: Request, _=Depends(require_admin)):
    for r in rules:
        if r.min_delay < 1 or not r.name.strip() or not valid_emails(r.recipients):
            raise HTTPException(400, f"Rule '{r.name}': delay must be at least 1 day and recipients must be valid email addresses, separated by commas.")
    if len({r.min_delay for r in rules}) != len(rules):
        raise HTTPException(400, "Two rules have the same delay threshold.")
    with LOCK:
        c = db()
        keep = [r.id for r in rules if r.id]
        c.execute(f"UPDATE rules SET active=0 WHERE id NOT IN ({','.join('?' * len(keep)) or 'NULL'})", keep)
        for r in rules:
            if r.id:
                c.execute("UPDATE rules SET level=?,name=?,min_delay=?,recipients=?,active=? WHERE id=?",
                          (r.level, r.name.strip(), r.min_delay, r.recipients.strip(), int(r.active), r.id))
            else:
                c.execute("INSERT INTO rules(level,name,min_delay,recipients,active) VALUES(?,?,?,?,?)",
                          (r.level, r.name.strip(), r.min_delay, r.recipients.strip(), int(r.active)))
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
                c.execute("INSERT INTO projects VALUES(?,?,?)", (pid, pid, "MS Project import"))
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
    if dec.decision not in ("Approved", "Rejected") or not dec.name.strip():
        raise HTTPException(400, "Choose Approve or Reject and enter your name.")
    if dec.decision == "Rejected" and not dec.comment.strip():
        raise HTTPException(400, "Say what recovery action you need when rejecting.")
    c.execute("UPDATE escalations SET status=?, decision=?, decided_by=?, comment=?, decided_at=? WHERE id=?",
              (dec.decision, ("Revised date accepted" if dec.decision == "Approved" else "Recovery plan required") + f" ({via})",
               dec.name.strip(), dec.comment.strip(), now(), esc["id"]))
    m = c.execute("SELECT * FROM milestones WHERE id=?", (esc["milestone_id"],)).fetchone()
    deliver(c, esc["recipients"], f"[MTA] {dec.decision}: {m['id']} {m['name']} (L{esc['level']} escalation)",
            f"{dec.name.strip()} {dec.decision.lower()} the level {esc['level']} escalation for {m['id']} {m['name']}.\n"
            f"Delay at escalation: {esc['delay_days']} days. Forecast: {fmt(esc['forecast'])}.\n"
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
        seed(c, reset=True)
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
.b{{display:flex;gap:10px;margin-top:16px;flex-wrap:wrap}} button{{font:600 14px inherit;padding:10px 16px;border-radius:7px;border:0;cursor:pointer;color:#fff}}
.ok{{background:#1E8449}} .no{{background:#C0392B}} .pill{{display:inline-block;padding:3px 10px;border-radius:99px;font-weight:600;font-size:13px;background:#E6ECF5}}
.late{{color:#C0392B;font-weight:700}}</style></head><body><div class="c">{content}</div></body></html>"""


def esc_view(c, token):
    esc = c.execute("SELECT * FROM escalations WHERE token=?", (token,)).fetchone()
    if not esc:
        raise HTTPException(404, "This approval link is not valid.")
    m = c.execute("SELECT m.*, p.name AS pname FROM milestones m JOIN projects p ON p.id=m.project_id WHERE m.id=?", (esc["milestone_id"],)).fetchone()
    return esc, m


def approval_html(esc, m, msg=""):
    e = html.escape
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


# ----------------------------------------------------------------------------- front end
@app.get("/health")
def health():
    return {"ok": True}


@app.get("/")
def index():
    return FileResponse(os.path.join(HERE, "static", "index.html"))


@app.exception_handler(HTTPException)
async def http_err(request: Request, exc: HTTPException):
    if request.url.path.startswith("/approve/") and request.method == "GET":
        return HTMLResponse(PAGE.format(content=f"<h1>Link not valid</h1><p>{html.escape(str(exc.detail))}</p>"), status_code=exc.status_code)
    return JSONResponse({"detail": exc.detail}, status_code=exc.status_code)
