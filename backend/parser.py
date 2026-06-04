"""
Parse the ONtime (Secureye) Employee Performance Register PDF.

Key design
----------
We use pdfplumber's word-level x/y coordinate extraction instead of a
text-stream queue.  Every time token is assigned to a day column by its
x-centre position, with a half-column-width tolerance.

Effect: a "MIS" day (only one punch recorded) leaves the other punch as
None instead of stealing the next day's value — which was the root cause
of the bug reported (Ronak's 04/04/26 out-time was being shown as 17:46,
the value belonging to 05/04/26).

Returns
-------
punch_data   : { emp_code : { day_num : (in_str|None, out_str|None) } }
name_to_code : { lower_name : emp_code }
year         : int
month        : int  (1-12)
"""

import re
import pdfplumber
from datetime import date as _date

TIME_RE   = re.compile(r"^\d{1,2}:\d{2}$")
EMP_RE    = re.compile(
    r"Emp\s*Code\s*:\s*(\S+)\s+Emp\s*Name\s*:\s*(.+?)\s+Present\s*:",
    re.IGNORECASE,
)
PERIOD_RE = re.compile(r"(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})")


# ── Positional helpers ────────────────────────────────────────────────────────

def _group_lines(words, y_tol=5):
    """
    Cluster pdfplumber word-dicts into horizontal lines.
    Words within y_tol points of each other (vertically) are on the same line.
    Words are pre-sorted by top then x0 so page-order is guaranteed.
    Each returned line is sorted left-to-right by x0.
    """
    if not words:
        return []
    # Sort top-to-bottom, then left-to-right within each row
    words = sorted(words, key=lambda w: (round(w["top"] / y_tol), w["x0"]))
    rows, cur = [], [words[0]]
    for w in words[1:]:
        if abs(w["top"] - cur[0]["top"]) <= y_tol:
            cur.append(w)
        else:
            rows.append(sorted(cur, key=lambda x: x["x0"]))
            cur = [w]
    rows.append(sorted(cur, key=lambda x: x["x0"]))
    return rows


def _build_day_map(line_words):
    """
    Given words from the day-number line (1 … 31),
    return {day_num: x_centre_of_that_column}.
    """
    day_map = {}
    for w in line_words:
        t = w["text"]
        if t.isdigit() and 1 <= int(t) <= 31:
            day_map[int(t)] = (w["x0"] + w["x1"]) / 2.0
    return day_map


def _assign_times(time_words, day_map):
    """
    Assign each time token to the nearest day column by x-centre.

    Tolerance = 48 % of the minimum column gap.  A time whose centre
    is more than that distance from every day column is discarded rather
    than being mis-assigned to a neighbouring day.

    Returns {day_num: time_str}.  Days with no time present in the PDF
    simply won't appear in the returned dict (caller treats them as None).
    """
    if not day_map:
        return {}

    xs = sorted(day_map.values())
    if len(xs) >= 2:
        min_gap  = min(xs[i + 1] - xs[i] for i in range(len(xs) - 1))
        # 60% of minimum column gap: tolerant enough for slight PDF rendering
        # offsets but still strict enough not to assign a time to a wrong day.
        half_col = min_gap * 0.60
    else:
        half_col = 30.0                    # sensible fallback for 1-day edge case

    result = {}
    for w in time_words:
        if not TIME_RE.match(w["text"]):
            continue
        wx = (w["x0"] + w["x1"]) / 2.0
        best_day, best_dist = None, float("inf")
        for day, dx in day_map.items():
            d = abs(wx - dx)
            if d < best_dist:
                best_dist, best_day = d, day
        if best_day is not None and best_dist <= half_col:
            result[best_day] = w["text"]

    return result


# ── Period extraction ─────────────────────────────────────────────────────────

def _extract_period(text_lines):
    """
    Scan the first ~40 raw text lines for a DD/MM/YYYY date near a
    'For Period' header.  Returns (year, month); falls back to today.
    """
    for line in text_lines[:40]:
        if re.search(r"For\s+Period", line, re.IGNORECASE) or \
           re.match(r"^\d{2}/\d{2}/\d{4}", line.strip()):
            m = PERIOD_RE.search(line)
            if m:
                _, mo, yr = int(m.group(1)), int(m.group(2)), int(m.group(3))
                if 1 <= mo <= 12:
                    return yr, mo
    t = _date.today()
    return t.year, t.month


# ── Per-page positional parser ────────────────────────────────────────────────

def _parse_page(page):
    """
    Extract all employee sections from one PDF page using word coordinates.

    Returns a list of dicts:
      {
        "code":    str,
        "name":    str,
        "day_map": {day_num: x_centre},
        "in_map":  {day_num: time_str},
        "out_map": {day_num: time_str},
      }
    """
    # ── CRITICAL: do NOT pass extra_attrs ────────────────────────────────
    # extra_attrs overwrites the correctly-computed word-level x0/x1 with
    # single-character values, breaking x-centre calculations for multi-
    # character tokens like "09:15".  x0, x1, top are always present in
    # pdfplumber word dicts without any extra_attrs.
    words = page.extract_words(x_tolerance=3, y_tolerance=3)
    if not words:
        return []

    lines    = _group_lines(words)
    sections = []
    current  = None

    for line in lines:
        texts    = [w["text"] for w in line]
        line_str = " ".join(texts)

        # ── New employee header ───────────────────────────────────────────
        emp_m = EMP_RE.search(line_str)
        if emp_m:
            if current and current["day_map"]:
                sections.append(current)
            current = {
                "code":    emp_m.group(1).strip(),
                "name":    emp_m.group(2).strip(),
                "day_map": {},
                "in_map":  {},
                "out_map": {},
            }
            continue

        if current is None:
            continue

        # ── Waiting for the day-number row ────────────────────────────────
        if not current["day_map"]:
            # Day-number rows have 10+ small integers in [1, 31].
            # Threshold is 10 (not 15) to handle short months and split pages.
            day_nums = [t for t in texts if re.match(r"^\d{1,2}$", t)
                        and 1 <= int(t) <= 31]
            if len(day_nums) >= 10:
                current["day_map"] = _build_day_map(line)
            continue          # keep scanning until day_map is found

        # ── day_map ready — identify In Time / Out Time rows ─────────────
        # Detect by: line starts with "In Time" or "Out Time" (any spacing),
        # AND the line contains at least 1 actual time token.
        # Using re.search (not match) to tolerate any leading tokens.
        time_words = [w for w in line if TIME_RE.match(w["text"])]
        if not time_words:
            continue

        if re.search(r"\bIn\s+Time\b", line_str, re.IGNORECASE):
            current["in_map"].update(_assign_times(time_words, current["day_map"]))

        elif re.search(r"\bOut\s+Time\b", line_str, re.IGNORECASE):
            current["out_map"].update(_assign_times(time_words, current["day_map"]))

    # Flush the last section on this page
    if current and current["day_map"]:
        sections.append(current)

    return sections


# ── Public entry point ────────────────────────────────────────────────────────

def parse_pdf_punch(pdf_path: str):
    """
    Parse every page and merge results for employees that span pages.

    Returns (punch_data, name_to_code, year, month).
    """
    accumulated  = {}   # code → {day_map, in_map, out_map}
    name_to_code = {}
    year = month = None

    with pdfplumber.open(pdf_path) as pdf:
        # Extract period from the raw text of the first page
        first_text = (pdf.pages[0].extract_text() or "").split("\n")
        year, month = _extract_period(first_text)

        for page in pdf.pages:
            for sec in _parse_page(page):
                code = sec["code"]
                name_to_code[sec["name"].lower()] = code

                if code not in accumulated:
                    accumulated[code] = {
                        "day_map": {},
                        "in_map":  {},
                        "out_map": {},
                    }
                # Merge — handles employees whose data spans two pages
                accumulated[code]["day_map"].update(sec["day_map"])
                accumulated[code]["in_map"].update(sec["in_map"])
                accumulated[code]["out_map"].update(sec["out_map"])

    # Build the final punch dict
    punch_data = {}
    for code, acc in accumulated.items():
        emp_dict = {}
        for day in acc["day_map"]:
            emp_dict[day] = (
                acc["in_map"].get(day),    # None  →  no in-punch for this day
                acc["out_map"].get(day),   # None  →  no out-punch for this day
            )
        punch_data[code] = emp_dict

    return punch_data, name_to_code, year, month


# ─────────────────────────────────────────────────────────────────────────────
#  Excel meta-data parser  (unchanged API)
# ─────────────────────────────────────────────────────────────────────────────

_DAY_NAMES = {
    "monday": 0, "mon": 0,
    "tuesday": 1, "tue": 1,
    "wednesday": 2, "wed": 2,
    "thursday": 3, "thu": 3,
    "friday": 4, "fri": 4,
    "saturday": 5, "sat": 5,
    "sunday": 6, "sun": 6,
}


def _resolve_week_off(week_off_str: str, location: str) -> int:
    """Convert day-name string → weekday int (0=Mon … 6=Sun).
    Falls back to location default when the string is unrecognised."""
    if week_off_str:
        key = week_off_str.lower().strip()
        if key in _DAY_NAMES:
            return _DAY_NAMES[key]
    return 1 if "factory" in location.lower() else 6


def parse_xlsx_meta(xlsx_path: str) -> dict:
    import pandas as pd
    from datetime import date as dt_date

    df   = pd.read_excel(xlsx_path, sheet_name=0, header=None)
    rows = df.values.tolist()

    leave_balances = {}
    sl_applied_map = {}
    workforce      = []
    holidays       = []
    section        = None

    for row in rows:
        r0 = str(row[0]).strip() if row[0] is not None and str(row[0]) != "nan" else ""

        if "Leave Balances"      in r0: section = "leave";        continue
        if "Sick Leaves Applied" in r0: section = "sick_applied"; continue
        if "List of Employees"   in r0: section = "workforce";    continue
        if "List of Company"     in r0: section = "holidays";     continue

        # Skip header / label rows
        if r0.lower() in ("name", "date", "", "nan", "emp id", "empid",
                          "employee id", "location", "type",
                          "week off", "week-off"):
            continue

        # ── Leave Balances ────────────────────────────────────────────────
        if section == "leave" and r0:
            try:
                cl_v = float(str(row[1])) if str(row[1]) != "nan" else 0.0
                el_v = float(str(row[2])) if str(row[2]) != "nan" else 0.0
                sl_v = (float(str(row[3]))
                        if len(row) > 3 and str(row[3]) != "nan" else 0.0)
                leave_balances[r0.lower()] = {"cl": cl_v, "el": el_v, "sl": sl_v}
            except Exception:
                pass

        # ── Sick Leaves Applied ───────────────────────────────────────────
        elif section == "sick_applied" and r0:
            try:
                sl_applied_map[r0.lower()] = (
                    float(str(row[1])) if str(row[1]) != "nan" else 0.0
                )
            except Exception:
                pass

        # ── Workforce ─────────────────────────────────────────────────────
        elif section == "workforce" and r0:
            try:
                def _col(idx):
                    return (str(row[idx]).strip()
                            if len(row) > idx and str(row[idx]) != "nan"
                            else "")

                # ── Auto-detect Excel column format ──────────────────────
                # New 8-col: Name | EmpId | Designation | Dept | Area | Location | Type | Week-Off
                # Old 4-col: Name | Location | Type | Week-Off
                loc    = _col(5)
                typ    = _col(6)
                wo_str = _col(7)

                if loc and typ:
                    # New format — all extra columns present
                    emp_id = _col(1)
                    desig  = _col(2)
                    dept   = _col(3)
                    area   = _col(4)
                else:
                    # Old 4-column format fallback
                    loc    = _col(1)
                    typ    = _col(2)
                    wo_str = _col(3)
                    emp_id = ""
                    desig  = ""
                    dept   = ""
                    area   = ""

                if loc and typ:
                    workforce.append({
                        "name":         r0,
                        "emp_id":       emp_id,
                        "designation":  desig,
                        "department":   dept,
                        "area":         area,
                        "location":     loc,
                        "type":         typ,
                        "week_off_str": wo_str,
                    })
            except Exception:
                pass

        # ── Holidays ──────────────────────────────────────────────────────
        elif section == "holidays" and r0:
            try:
                val = row[0]
                if hasattr(val, "date"):
                    holidays.append(val.date())
                elif isinstance(val, dt_date):
                    holidays.append(val)
            except Exception:
                pass

    # Build employee list
    employees = []
    for w in workforce:
        key      = w["name"].lower()
        lb       = leave_balances.get(key, {"el": 0, "cl": 0, "sl": 0})
        sa       = sl_applied_map.get(key, 0)
        loc      = w["location"].strip()
        typ      = w["type"].strip()
        week_off = _resolve_week_off(w.get("week_off_str", ""), loc)
        employees.append({
            "name":        w["name"],
            "emp_id":      w.get("emp_id", ""),
            "designation": w.get("designation", ""),
            "department":  w.get("department", ""),
            "area":        w.get("area", ""),
            "location":    loc,
            "type":        typ,
            "week_off":    week_off,
            "el":          lb["el"] if typ == "Full-time" else None,
            "cl":          lb["cl"] if typ == "Full-time" else None,
            "sl":          lb["sl"] if typ == "Full-time" else None,
            "sl_applied":  sa       if typ == "Full-time" else None,
        })

    return {"employees": employees, "holidays": holidays}
