"""
SQLite persistence layer for frozen attendance registers.
"""
import sqlite3
import shutil
import math as _math
import json
from pathlib import Path
from datetime import datetime
from typing import Optional

DB_PATH    = Path(__file__).parent.parent / "attendance.db"
FROZEN_DIR = Path(__file__).parent.parent / "frozen"
PENDING_DIR = Path(__file__).parent.parent / "pending"
FROZEN_DIR.mkdir(exist_ok=True)
PENDING_DIR.mkdir(exist_ok=True)


def _conn():
    c = sqlite3.connect(str(DB_PATH))
    c.row_factory = sqlite3.Row
    c.execute("PRAGMA foreign_keys = ON")
    return c


def init_db():
    c = _conn()
    try:
        c.executescript("""
            CREATE TABLE IF NOT EXISTS registers (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                month       TEXT    NOT NULL,
                year        INTEGER NOT NULL,
                month_num   INTEGER NOT NULL,
                total_emp   INTEGER,
                frozen_at   TEXT    NOT NULL,
                frozen_file TEXT
            );

            CREATE TABLE IF NOT EXISTS employee_records (
                id                   INTEGER PRIMARY KEY AUTOINCREMENT,
                register_id          INTEGER NOT NULL REFERENCES registers(id) ON DELETE CASCADE,
                name                 TEXT,
                emp_id               TEXT,
                designation          TEXT,
                department           TEXT,
                area                 TEXT,
                location             TEXT,
                type                 TEXT,
                week_off             INTEGER,
                code                 TEXT,
                present_days         INTEGER,
                absent_days          INTEGER,
                half_days            INTEGER,
                total_net_hours      REAL,
                total_lwp            REAL,
                el_opening           REAL,
                cl_opening           REAL,
                sl_opening           REAL,
                el_closing           REAL,
                cl_closing           REAL,
                sl_closing           REAL,
                el_accrual           REAL,
                sl_applied           REAL,
                working_days         INTEGER,
                total_overtime_hours REAL,
                ot_days              REAL
            );

            CREATE TABLE IF NOT EXISTS daily_attendance (
                id                 INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_record_id INTEGER NOT NULL REFERENCES employee_records(id) ON DELETE CASCADE,
                day                INTEGER,
                date               TEXT,
                dow                TEXT,
                day_type           TEXT,
                status             TEXT,
                in_time            TEXT,
                out_time           TEXT,
                short_mins         REAL,
                long_mins          REAL,
                net_mins           REAL,
                late_mark          TEXT,
                half_day           TEXT,
                overtime_mins      REAL
            );

            CREATE TABLE IF NOT EXISTS users (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                email       TEXT NOT NULL UNIQUE,
                name        TEXT,
                picture     TEXT,
                role        TEXT NOT NULL DEFAULT 'maker',
                created_at  TEXT NOT NULL,
                last_login  TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS pending_registers (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                session_id      TEXT NOT NULL,
                month           TEXT NOT NULL,
                year            INTEGER NOT NULL,
                month_num       INTEGER NOT NULL,
                total_emp       INTEGER,
                submitted_at    TEXT NOT NULL,
                submitted_by    TEXT,
                status          TEXT NOT NULL DEFAULT 'pending',
                pending_file    TEXT,
                summary_json    TEXT NOT NULL,
                employees_json  TEXT NOT NULL
            );
        """)
        cols = {
            row["name"]
            for row in c.execute("PRAGMA table_info(daily_attendance)").fetchall()
        }
        if "extra_punches" not in cols:
            c.execute("ALTER TABLE daily_attendance ADD COLUMN extra_punches TEXT")
        c.commit()
    finally:
        c.close()


def upsert_user(email: str, name: str = "", picture: str = "",
                default_role: str = "maker", admin_emails=None,
                maker_emails=None) -> dict:
    email = (email or "").strip().lower()
    if not email:
        raise ValueError("Email is required")
    admin_emails = {e.strip().lower() for e in (admin_emails or []) if e.strip()}
    maker_emails = {e.strip().lower() for e in (maker_emails or []) if e.strip()}
    now = datetime.now().isoformat()
    c = _conn()
    try:
        row = c.execute("SELECT * FROM users WHERE email=?", (email,)).fetchone()
        if row:
            if email in admin_emails:
                role = "admin"
            elif email in maker_emails:
                role = "maker"
            else:
                role = row["role"]
            c.execute(
                """UPDATE users
                   SET name=?, picture=?, role=?, last_login=?
                   WHERE email=?""",
                (name, picture, role, now, email),
            )
        else:
            existing_count = c.execute("SELECT COUNT(*) AS n FROM users").fetchone()["n"]
            if email in admin_emails:
                role = "admin"
            elif email in maker_emails:
                role = "maker"
            else:
                role = "admin" if existing_count == 0 else default_role
            c.execute(
                """INSERT INTO users (email, name, picture, role, created_at, last_login)
                   VALUES (?,?,?,?,?,?)""",
                (email, name, picture, role, now, now),
            )
        c.commit()
        return get_user(email)
    finally:
        c.close()


def get_user(email: str) -> Optional[dict]:
    c = _conn()
    try:
        row = c.execute(
            "SELECT id, email, name, picture, role, created_at, last_login FROM users WHERE email=?",
            ((email or "").strip().lower(),),
        ).fetchone()
        return dict(row) if row else None
    finally:
        c.close()


def list_users() -> list:
    c = _conn()
    try:
        rows = c.execute(
            "SELECT id, email, name, picture, role, created_at, last_login FROM users ORDER BY email"
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        c.close()


def set_user_role(email: str, role: str) -> Optional[dict]:
    if role not in ("maker", "admin"):
        raise ValueError("Role must be maker or admin")
    c = _conn()
    try:
        c.execute("UPDATE users SET role=? WHERE email=?", (role, email.strip().lower()))
        c.commit()
        return get_user(email)
    finally:
        c.close()


def freeze_register(summary: dict, employees: list,
                    src_excel_path: str, session_id: str) -> int:
    frozen_file = f"frozen_{session_id}.xlsx"
    shutil.copy(src_excel_path, str(FROZEN_DIR / frozen_file))

    c = _conn()
    try:
        cur = c.execute(
            """INSERT INTO registers (month, year, month_num, total_emp, frozen_at, frozen_file)
               VALUES (?,?,?,?,?,?)""",
            (summary["month"], summary["year"], summary["month_num"],
             len(employees), datetime.now().isoformat(), frozen_file)
        )
        reg_id = cur.lastrowid

        for emp in employees:
            cur2 = c.execute(
                """INSERT INTO employee_records
                   (register_id, name, emp_id, designation, department, area,
                    location, type, week_off, code,
                    present_days, absent_days, half_days,
                    total_net_hours, total_lwp,
                    el_opening, cl_opening, sl_opening,
                    el_closing, cl_closing, sl_closing,
                    el_accrual, sl_applied,
                    working_days, total_overtime_hours, ot_days)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (reg_id,
                 emp.get("name"), emp.get("emp_id", ""), emp.get("designation", ""),
                 emp.get("department", ""), emp.get("area", ""),
                 emp.get("location"), emp.get("type"), emp.get("week_off"), emp.get("code"),
                 emp.get("present_days"), emp.get("absent_days"), emp.get("half_days", 0),
                 emp.get("total_net_hours"), emp.get("total_lwp"),
                 emp.get("el_opening"), emp.get("cl_opening"), emp.get("sl_opening"),
                 emp.get("el_closing"), emp.get("cl_closing"), emp.get("sl_closing"),
                 emp.get("el_accrual"), emp.get("sl_applied"),
                 emp.get("working_days"), emp.get("total_overtime_hours"), emp.get("ot_days"))
            )
            eid = cur2.lastrowid

            for d in emp.get("daily", []):
                c.execute(
                    """INSERT INTO daily_attendance
                       (employee_record_id, day, date, dow, day_type, status,
                        in_time, out_time, short_mins, long_mins, net_mins,
                        late_mark, half_day, overtime_mins, extra_punches)
                       VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (eid,
                     d.get("day"), d.get("date"), d.get("dow"),
                     d.get("day_type"), d.get("status"),
                     d.get("in_time"), d.get("out_time"),
                     d.get("short_mins", 0), d.get("long_mins", 0), d.get("net_mins", 0),
                     d.get("late_mark", ""), d.get("half_day", ""),
                     d.get("overtime_mins", 0),
                     json.dumps(d.get("extra_punches", [])))
                )

        c.commit()
        return reg_id
    except Exception:
        c.rollback()
        raise
    finally:
        c.close()


def submit_register_for_approval(summary: dict, employees: list,
                                 src_excel_path: str, session_id: str,
                                 submitted_by: str) -> int:
    pending_file = f"pending_{session_id}.xlsx"
    shutil.copy(src_excel_path, str(PENDING_DIR / pending_file))
    c = _conn()
    try:
        cur = c.execute(
            """INSERT INTO pending_registers
               (session_id, month, year, month_num, total_emp, submitted_at,
                submitted_by, status, pending_file, summary_json, employees_json)
               VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
            (session_id, summary["month"], summary["year"], summary["month_num"],
             len(employees), datetime.now().isoformat(), submitted_by, "pending",
             pending_file, json.dumps(summary), json.dumps(employees))
        )
        c.commit()
        return cur.lastrowid
    except Exception:
        c.rollback()
        raise
    finally:
        c.close()


def list_pending_registers(status: str = "pending") -> list:
    c = _conn()
    try:
        rows = c.execute(
            """SELECT id, session_id, month, year, month_num, total_emp,
                      submitted_at, submitted_by, status
               FROM pending_registers
               WHERE status=?
               ORDER BY id DESC""",
            (status,),
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        c.close()


def get_pending_register(pending_id: int) -> Optional[dict]:
    c = _conn()
    try:
        row = c.execute("SELECT * FROM pending_registers WHERE id=?", (pending_id,)).fetchone()
        if not row:
            return None
        data = dict(row)
        data["summary"] = json.loads(data.pop("summary_json") or "{}")
        data["employees"] = json.loads(data.pop("employees_json") or "[]")
        return data
    finally:
        c.close()


def get_pending_file(pending_id: int) -> Optional[str]:
    c = _conn()
    try:
        row = c.execute("SELECT pending_file FROM pending_registers WHERE id=?", (pending_id,)).fetchone()
        if not row or not row["pending_file"]:
            return None
        p = PENDING_DIR / row["pending_file"]
        return str(p) if p.exists() else None
    finally:
        c.close()


def mark_pending_status(pending_id: int, status: str) -> None:
    c = _conn()
    try:
        c.execute("UPDATE pending_registers SET status=? WHERE id=?", (status, pending_id))
        c.commit()
    finally:
        c.close()


def list_registers() -> list:
    c = _conn()
    try:
        rows = c.execute(
            "SELECT id, month, year, month_num, total_emp, frozen_at FROM registers ORDER BY id DESC"
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        c.close()


def _recompute_emp_summary(emp_type: str, daily: list,
                           el_opening, cl_opening, sl_opening, sl_applied) -> dict:
    """
    Recompute all employee summary fields fresh from daily_attendance rows.
    The stored summary can be stale (captured before user edits were applied);
    daily_attendance is the authoritative source of truth.
    """
    el_o = float(el_opening or 0)
    cl_o = float(cl_opening or 0)
    sl_o = float(sl_opening or 0)
    sl_a = float(sl_applied or 0)

    total_net_mins = sum(float(d.get("net_mins") or 0) for d in daily)
    wd_absent      = [d for d in daily
                      if d.get("day_type") == "Working Day" and d.get("status") == "Absent"]

    if emp_type == "Labour":
        all_present   = [d for d in daily if d.get("status") == "Present"]
        total_ot_mins = sum(float(d.get("overtime_mins") or 0) for d in daily)
        ot_hrs        = round(total_ot_mins / 60, 4)
        return {
            "present_days":        len(all_present),
            "absent_days":         len(wd_absent),
            "half_days":           0,
            "total_net_hours":     round(total_net_mins / 60, 4),
            "total_lwp":           0.0,
            "el_opening":          None,
            "cl_opening":          None,
            "sl_opening":          None,
            "el_closing":          None,
            "cl_closing":          None,
            "sl_closing":          None,
            "el_accrual":          None,
            "sl_applied":          None,
            "total_overtime_hours": ot_hrs,
            "ot_days":             round(total_ot_mins / (6 * 60), 4),
        }

    # ── Full-time employee ────────────────────────────────────────────────────
    wd_present    = [d for d in daily
                     if d.get("day_type") == "Working Day" and d.get("status") == "Present"]
    half_days_lst = [d for d in daily if d.get("half_day") == "Half Day"]

    present_days = len(wd_present)
    absent_days  = len(wd_absent)
    num_half     = len(half_days_lst)
    net_hours    = total_net_mins / 60

    # EL accrual: floor(net_hours/9) when positive, ceil when negative
    el_accrual = (_math.ceil(net_hours / 9)
                  if net_hours < 0
                  else _math.floor(net_hours / 9))

    total_absent = absent_days + 0.5 * num_half

    # Adjust closing balances
    cl_adj    = min(1.0, cl_o, max(0.0, total_absent))
    sl_adj    = min(sl_a, sl_o)
    adj_el_o  = el_o + el_accrual              # can be negative (deficit)
    el_adj    = min(max(0.0, adj_el_o),
                    max(0.0, total_absent - cl_adj - sl_adj))

    # LWP from EL deficit + excess absences not covered by any leave
    lwp_neg_el = max(0.0, -adj_el_o)
    lwp_excess = max(0.0, total_absent - cl_adj - sl_adj - el_adj)
    total_lwp  = round(lwp_neg_el + lwp_excess, 4)

    return {
        "present_days":        present_days,
        "absent_days":         absent_days,
        "half_days":           num_half,
        "total_net_hours":     round(net_hours, 4),
        "total_lwp":           total_lwp,
        "el_opening":          el_o,
        "cl_opening":          cl_o,
        "sl_opening":          sl_o,
        "el_closing":          round(max(0.0, adj_el_o - el_adj), 4),
        "cl_closing":          round(max(0.0, cl_o - cl_adj), 4),
        "sl_closing":          round(max(0.0, sl_o - sl_adj), 4),
        "el_accrual":          el_accrual,
        "sl_applied":          sl_a,
        "total_overtime_hours": 0.0,
        "ot_days":             0.0,
    }


def get_register_detail(register_id: int) -> Optional[dict]:
    c = _conn()
    try:
        reg = c.execute("SELECT * FROM registers WHERE id=?", (register_id,)).fetchone()
        if not reg:
            return None
        reg = dict(reg)
        emps = c.execute(
            """SELECT id, name, emp_id, designation, department, area, location, type,
                      el_opening, cl_opening, sl_opening, sl_applied
               FROM employee_records WHERE register_id=? ORDER BY name""",
            (register_id,)
        ).fetchall()
        employees = []
        for emp in emps:
            ed = dict(emp)
            daily = c.execute(
                "SELECT * FROM daily_attendance WHERE employee_record_id=? ORDER BY day",
                (emp["id"],)
            ).fetchall()
            ed["daily"] = []
            for d in daily:
                row = dict(d)
                try:
                    row["extra_punches"] = json.loads(row.get("extra_punches") or "[]")
                except Exception:
                    row["extra_punches"] = []
                ed["daily"].append(row)

            # Recompute summary from daily rows — stored summary may be stale
            computed = _recompute_emp_summary(
                emp_type   = ed["type"],
                daily      = ed["daily"],
                el_opening = ed.get("el_opening"),
                cl_opening = ed.get("cl_opening"),
                sl_opening = ed.get("sl_opening"),
                sl_applied = ed.get("sl_applied"),
            )
            ed.update(computed)
            employees.append(ed)
        reg["employees"] = employees
        return reg
    finally:
        c.close()


def get_frozen_file(register_id: int) -> Optional[str]:
    c = _conn()
    try:
        row = c.execute(
            "SELECT frozen_file FROM registers WHERE id=?", (register_id,)
        ).fetchone()
        if not row or not row["frozen_file"]:
            return None
        p = FROZEN_DIR / row["frozen_file"]
        return str(p) if p.exists() else None
    finally:
        c.close()
