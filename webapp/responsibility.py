"""Responsibility matrix: who updates which discipline's milestone dates, how often,
and who is escalated to when updates are missed or milestones slip.

A discipline is matched to milestones by the milestone's Category (case-insensitive).
"""
import calendar
import csv
import datetime as dt
import io
import os
import secrets

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

MATRIX_SCHEMA = """
CREATE TABLE IF NOT EXISTS disciplines(id INTEGER PRIMARY KEY AUTOINCREMENT, discipline TEXT NOT NULL UNIQUE COLLATE NOCASE,
  person TEXT NOT NULL, email TEXT NOT NULL, frequency TEXT NOT NULL DEFAULT 'Monthly', due_day INTEGER NOT NULL DEFAULT 1,
  grace_days INTEGER NOT NULL DEFAULT 3, l1_email TEXT, l2_email TEXT, token TEXT UNIQUE NOT NULL, active INTEGER NOT NULL DEFAULT 1);
CREATE TABLE IF NOT EXISTS reminders(discipline_id INTEGER NOT NULL, due_date TEXT NOT NULL, sent_at TEXT, PRIMARY KEY(discipline_id, due_date));
"""
MIGRATIONS = [
    "ALTER TABLE escalations ADD COLUMN kind TEXT DEFAULT 'delay'",
    "ALTER TABLE escalations ADD COLUMN discipline TEXT",
    "ALTER TABLE escalations ADD COLUMN period TEXT",
    "ALTER TABLE rules ADD COLUMN route_discipline INTEGER DEFAULT 1",
]
SEED = [  # discipline, responsible person, email, frequency, due day, grace, L1 (manager), L2 (director)
    ("Governance", "Sarah Mitchell", "governance.updates@example.com", "Monthly", 1, 5, "pmo.manager@example.com", "portfolio.director@example.com"),
    ("Design", "Lena Fischer", "design.updates@example.com", "Monthly", 1, 5, "design.manager@example.com", "engineering.director@example.com"),
    ("Procurement", "Helen Price", "procurement.updates@example.com", "Monthly", 1, 5, "procurement.manager@example.com", "commercial.director@example.com"),
    ("Build", "Kenji Watanabe", "build.updates@example.com", "Monthly", 1, 5, "build.manager@example.com", "engineering.director@example.com"),
    ("Construction", "Marco Rossi", "construction.updates@example.com", "Monthly", 1, 5, "construction.manager@example.com", "operations.director@example.com"),
    ("Testing", "Anna Kowalski", "testing.updates@example.com", "Weekly", 5, 2, "test.manager@example.com", "engineering.director@example.com"),
    ("Deployment", "Maria Lopez", "deployment.updates@example.com", "Weekly", 5, 2, "release.manager@example.com", "delivery.director@example.com"),
    ("Handover", "David Okafor", "handover.updates@example.com", "Monthly", 1, 5, "operations.manager@example.com", "operations.director@example.com"),
]
WEEKDAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]


def today():
    return dt.date.fromisoformat(os.environ["MTA_TODAY"]) if os.environ.get("MTA_TODAY") else dt.date.today()


def setup(con):
    con.executescript(MATRIX_SCHEMA)
    for m in MIGRATIONS:
        try:
            con.execute(m)
        except Exception:
            pass  # column already exists
    if con.execute("SELECT COUNT(*) FROM disciplines").fetchone()[0] == 0:
        con.executemany("""INSERT INTO disciplines(discipline,person,email,frequency,due_day,grace_days,l1_email,l2_email,token)
                           VALUES(?,?,?,?,?,?,?,?,?)""", [s + (secrets.token_urlsafe(16),) for s in SEED])
    con.commit()


# ----------------------------------------------------------------------------- cadence
def _month_day(y, m, day):
    return dt.date(y, m, min(day, calendar.monthrange(y, m)[1]))


def due_dates_after(freq, due_day, after, until):
    """Due dates d with after < d <= until."""
    out = []
    if freq == "Weekly":
        d = after + dt.timedelta(days=1)
        while d.isoweekday() != due_day:
            d += dt.timedelta(days=1)
        while d <= until:
            out.append(d)
            d += dt.timedelta(days=7)
    else:
        y, m = after.year, after.month
        while True:
            d = _month_day(y, m, due_day)
            if d > until:
                break
            if d > after:
                out.append(d)
            y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    return out


def next_due(freq, due_day, frm):
    return due_dates_after(freq, due_day, frm - dt.timedelta(days=1), frm + dt.timedelta(days=62))[0]


def cadence_text(r):
    return f"Weekly, every {WEEKDAYS[r['due_day'] - 1]}" if r["frequency"] == "Weekly" else f"Monthly, day {r['due_day']}"


# ----------------------------------------------------------------------------- compliance
def discipline_milestones(con, disc):
    """Open milestones of a discipline with their latest observation."""
    rows = con.execute("""SELECT m.*, p.name AS project_name,
                            (SELECT MAX(rep_date) FROM observations o WHERE o.milestone_id=m.id) AS last_rep,
                            (SELECT forecast FROM observations o WHERE o.milestone_id=m.id ORDER BY rep_date DESC LIMIT 1) AS last_fc,
                            (SELECT COUNT(*) FROM observations o WHERE o.milestone_id=m.id AND o.actual IS NOT NULL) AS done
                          FROM milestones m JOIN projects p ON p.id=m.project_id
                          WHERE m.category = ? COLLATE NOCASE ORDER BY m.project_id, m.baseline""", (disc,)).fetchall()
    return [dict(r) for r in rows if not r["done"]]


def compliance(con, r):
    t = today()
    ms = discipline_milestones(con, r["discipline"])
    reporting = [m for m in ms if m["last_rep"]]
    res = dict(id=r["id"], discipline=r["discipline"], person=r["person"], cadence=cadence_text(r), open=len(ms),
               reporting=len(reporting), next_due=next_due(r["frequency"], r["due_day"], t + dt.timedelta(days=1)).isoformat(),
               missed=[], status="No open milestones", stale=[])
    if not reporting:
        return res
    oldest = min(dt.date.fromisoformat(m["last_rep"]) for m in reporting)
    dues = due_dates_after(r["frequency"], r["due_day"], oldest, t)
    latest_due = dues[-1] if dues else None
    res["current_due"] = latest_due.isoformat() if latest_due else None
    res["missed"] = [d.isoformat() for d in dues if d + dt.timedelta(days=r["grace_days"]) < t]
    if latest_due:
        res["stale"] = [m["id"] for m in reporting if dt.date.fromisoformat(m["last_rep"]) < latest_due]
    res["updated"] = len(reporting) - len(res["stale"])
    if res["missed"]:
        res["status"] = "Overdue"
        res["overdue_days"] = (t - dt.date.fromisoformat(res["missed"][0])).days
    elif res["stale"]:
        res["status"] = "Due"
        res["grace_until"] = (latest_due + dt.timedelta(days=r["grace_days"])).isoformat()
    else:
        res["status"] = "Up to date"
    return res


def matrix_rows(con, active_only=True):
    q = "SELECT * FROM disciplines" + (" WHERE active=1" if active_only else "") + " ORDER BY discipline"
    return [dict(r) for r in con.execute(q)]


def contacts_for(con, category, level):
    """Discipline recipients for a delay escalation at a given level."""
    r = con.execute("SELECT * FROM disciplines WHERE discipline=? COLLATE NOCASE AND active=1", (category or "",)).fetchone()
    if not r:
        return [], None
    emails = [r["email"]] + ([r["l1_email"]] if r["l1_email"] else [])
    if level >= 2 and r["l2_email"]:
        emails.append(r["l2_email"])
    return emails, dict(r)


def evaluate_updates(con, deliver, now, url, fmt):
    """Reminders when an update falls due; missed-update escalations (L1 after one missed due date,
    L2 after two); auto-close when the discipline catches up."""
    created = closed = reminded = 0
    for r in matrix_rows(con):
        c = compliance(con, r)
        link = f"{url}/update/{r['token']}"
        # reminder for the current due date (once)
        if c.get("current_due") and c["stale"] and not con.execute(
                "SELECT 1 FROM reminders WHERE discipline_id=? AND due_date=?", (r["id"], c["current_due"])).fetchone():
            con.execute("INSERT INTO reminders VALUES(?,?,?)", (r["id"], c["current_due"], now()))
            deliver(con, r["email"], f"[MTA] {r['discipline']} milestone update due {fmt(c['current_due'])}",
                    f"Hello {r['person']},\n\nThe milestone date update for "
                    f"{r['discipline']} ({cadence_text(r).lower()}) was due on {fmt(c['current_due'])}.\n{len(c['stale'])} of {c['reporting']} milestones still need a "
                    f"current forecast: {', '.join(c['stale'])}.\n\nUpdate the dates online or download your input file here:\n{link}\n\n"
                    f"Please update within {r['grace_days']} day(s) of the due date to avoid escalation.\n")
            reminded += 1
        pend = con.execute("SELECT * FROM escalations WHERE kind='missed_update' AND discipline=? AND status IN ('Pending approval','Acknowledged')",
                           (r["discipline"],)).fetchall()
        if not c["missed"]:
            for e in pend:
                con.execute("UPDATE escalations SET status='Closed', decision='Auto-closed: update received', decided_at=? WHERE id=?", (now(), e["id"]))
                closed += 1
            continue
        level = 2 if len(c["missed"]) >= 2 else 1
        period = c["missed"][0]
        if con.execute("SELECT 1 FROM escalations WHERE kind='missed_update' AND discipline=? AND period=? AND level=?",
                       (r["discipline"], period, level)).fetchone():
            continue
        for e in pend:
            con.execute("UPDATE escalations SET status='Superseded', decision=?, decided_at=? WHERE id=?", (f"Escalated to level {level}", now(), e["id"]))
        recips = [r["email"], r["l1_email"]] + ([r["l2_email"]] if level >= 2 else [])
        recips = ", ".join(dict.fromkeys(x for x in recips if x))
        token = secrets.token_urlsafe(18)
        cur = con.execute("""INSERT INTO escalations(milestone_id,rule_id,level,rule_name,recipients,delay_days,baseline,forecast,rep_date,
                             status,token,created_at,current_delay,kind,discipline,period)
                             VALUES('',NULL,?,?,?,?,NULL,NULL,NULL,'Pending approval',?,?,?,'missed_update',?,?)""",
                          (level, f"Missed {r['frequency'].lower()} update" + (" (2+ periods)" if level >= 2 else ""), recips,
                           c["overdue_days"], token, now(), c["overdue_days"], r["discipline"], period))
        deliver(con, recips, f"[MTA] L{level} escalation: {r['discipline']} milestone update overdue by {c['overdue_days']} days",
                f"Missed milestone update - Level {level}\n\nDiscipline:   {r['discipline']}\nResponsible:  {r['person']} <{r['email']}>\n"
                f"Cadence:      {cadence_text(r)}\nMissed due dates: {', '.join(fmt(x) for x in c['missed'])}\n"
                f"Milestones without a current forecast: {', '.join(c['stale'])}\n\n"
                f"Acknowledge this escalation:\n{url}/approve/{token}\n\nUpdate portal for the discipline:\n{link}\n", cur.lastrowid)
        created += 1
    return dict(created=created, closed=closed, reminded=reminded)


# ----------------------------------------------------------------------------- input file
INPUT_HEADERS = ["Discipline", "Responsible", "Project ID", "Project Name", "Milestone ID", "Milestone Name", "Baseline Date",
                 "Last Reported", "Last Forecast", "Reporting Date", "New Forecast Date", "Actual Completion Date", "Comment"]


def input_workbook(con, r):
    ms = discipline_milestones(con, r["discipline"])
    wb = Workbook()
    ws = wb.active
    ws.title = "Update"
    navy, yellow, thin = "1F3864", "FFF2CC", Side(style="thin", color="D8DEE8")
    ws["A1"] = f"{r['discipline']} - milestone date update"
    ws["A1"].font = Font(name="Arial", size=14, bold=True, color=navy)
    ws["A2"] = (f"Responsible: {r['person']} <{r['email']}>   |   Cadence: {cadence_text(r)}   |   "
                f"Generated {today():%d-%b-%Y}")
    ws["A3"] = ("Fill the yellow columns: Reporting Date (status date), New Forecast Date for every milestone, Actual Completion Date "
                "only if achieved, and a Comment for any change. Do not change the other columns. Upload the file on your update page.")
    for c in ("A2", "A3"):
        ws[c].font = Font(name="Arial", size=9, italic=True, color="5D6779")
    hr = 5
    for j, h in enumerate(INPUT_HEADERS, 1):
        cell = ws.cell(hr, j, h)
        cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", start_color="B8860B" if j >= 10 else navy)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[hr].height = 30
    t = today()
    for i, m in enumerate(ms, hr + 1):
        vals = [r["discipline"], r["person"], m["project_id"], m["project_name"], m["id"], m["name"],
                dt.date.fromisoformat(m["baseline"]), dt.date.fromisoformat(m["last_rep"]) if m["last_rep"] else None,
                dt.date.fromisoformat(m["last_fc"]) if m["last_fc"] else None, t, None, None, None]
        for j, v in enumerate(vals, 1):
            cell = ws.cell(i, j, v)
            cell.font = Font(name="Arial", size=10, color="0000FF" if j >= 10 else "000000")
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            if j >= 10:
                cell.fill = PatternFill("solid", start_color=yellow)
            if j in (7, 8, 9, 10, 11, 12):
                cell.number_format = "dd-mmm-yyyy"
    last = hr + max(1, len(ms))
    dv = DataValidation(type="date", operator="between", formula1="DATE(2020,1,1)", formula2="DATE(2040,12,31)",
                        showErrorMessage=True, errorTitle="Invalid date", error="Enter a real date, e.g. 15-Jan-2027.")
    dv.add(f"J{hr + 1}:L{last}")
    ws.add_data_validation(dv)
    for col, w in zip("ABCDEFGHIJKLM", (14, 16, 10, 26, 12, 34, 13, 13, 13, 14, 16, 16, 34)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = f"G{hr + 1}"
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def _parse_date(v):
    if v in (None, ""):
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = str(v).strip()
    for f in ("%Y-%m-%d", "%d-%b-%Y", "%d/%m/%Y", "%d/%m/%y", "%d %b %Y", "%d-%b-%y"):
        try:
            return dt.datetime.strptime(s, f).date()
        except ValueError:
            pass
    raise ValueError(f"'{s}' is not a date")


def parse_input_file(name, data):
    """Returns list of dicts with milestone_id, rep_date, forecast, actual, comment (row number for messages)."""
    if name.lower().endswith((".xlsx", ".xlsm")):
        ws = load_workbook(io.BytesIO(data), data_only=True).active
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
    else:
        text = data.decode("utf-8-sig", errors="replace")
        rows = list(csv.reader(io.StringIO(text)))
    hi = next((i for i, r in enumerate(rows) if r and "Milestone ID" in [str(x).strip() if x is not None else "" for x in r]), None)
    if hi is None:
        raise ValueError("No header row with 'Milestone ID' found. Use the input file downloaded from your update page.")
    idx = {str(h).strip(): j for j, h in enumerate(rows[hi]) if h is not None}
    need = ["Milestone ID", "Reporting Date", "New Forecast Date"]
    miss = [n for n in need if n not in idx]
    if miss:
        raise ValueError("Missing column(s): " + ", ".join(miss))
    out = []
    g = lambda r, k: r[idx[k]] if k in idx and idx[k] < len(r) else None  # noqa: E731
    for n, r in enumerate(rows[hi + 1:], hi + 2):
        mid = g(r, "Milestone ID")
        if not mid or not str(mid).strip():
            continue
        item = dict(row=n, milestone_id=str(mid).strip(), comment=(str(g(r, "Comment")).strip() if g(r, "Comment") else None))
        try:
            item["forecast"] = _parse_date(g(r, "New Forecast Date"))
            item["rep_date"] = _parse_date(g(r, "Reporting Date"))
            item["actual"] = _parse_date(g(r, "Actual Completion Date"))
        except ValueError as e:
            item["error"] = str(e)
        out.append(item)
    return out
