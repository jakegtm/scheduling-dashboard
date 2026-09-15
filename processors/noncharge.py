from __future__ import annotations
# ============================================================
# processors/noncharge.py
# ============================================================
# Reads non-charge time from the OpenAir time report.
#
# Replaces the old "Utilization" tab, which read utilization %
# out of the schedule workbook's "Utilization by Month" sheet.
# This reads actual logged non-charge time instead, with the
# task, notes and description, so people can be asked about it
# directly.
#
# Report format (same file variance.py parses — the report now
# carries three extra columns, which that parser ignores):
#   Row 1:  Title row (skipped)
#   Row 2:  Headers — "Project - Name", "Date", "Employee",
#           "Time (Hours)", "Task", "Notes", "Description"
#   Row 3+: One time entry per row
#   Footer: a grand-total row and "Generated on:" /
#           "Filter set applied:" rows, all skipped
#
# Non-charge time is any project matching VARIANCE_EXCLUDE_PREFIXES
# ("GTM"), i.e. GTM - NONCHG / TRAINING / PTO / HOLIDAYS. That is
# exactly the set variance.py excludes, so this tab shows the hours
# the variance tab drops.
# ============================================================

from datetime import datetime, timedelta
from config import (
    EMAIL_LOOKUP,
    FIRST_NAMES,
    OPENAIR_EMPLOYEE_MAP,
    VARIANCE_EXCLUDE_PREFIXES,
    NONCHARGE_NO_NOTE_TASKS,
)

AVAILABLE_TIME_TASK = "available time"

# The report packs multi-line note text onto one line with " | ".
_NOTE_SEPARATOR = " | "


def _parse_date(s: str):
    s = s.strip()
    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%m/%d/%y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def _date_to_period(d) -> str:
    """Half-month period label, e.g. 'September 1-15' / 'September 16-30'.
    Must stay identical to variance.py's period labels so the app's period
    selector drives both tabs."""
    month_name = d.strftime("%B")
    if d.day <= 15:
        return f"{month_name} 1-15"
    last_day = (d.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)
    return f"{month_name} 16-{last_day.day}"


def _is_noncharge(project_name: str) -> bool:
    upper = project_name.strip().upper()
    return any(upper.startswith(p.upper()) for p in VARIANCE_EXCLUDE_PREFIXES)


def _clean_note(raw: str) -> str:
    """Turn the report's ' | ' joins into real line breaks, and drop the
    stray leading apostrophe Excel adds to text starting with '+'."""
    text = (raw or "").strip()
    if text.startswith("'"):
        text = text[1:]
    parts = [p.strip() for p in text.split(_NOTE_SEPARATOR)]
    return "\n".join(p for p in parts if p)


def parse_noncharge_report(file_obj) -> dict:
    """
    Parse non-charge time entries out of the OpenAir time report.

    Returns:
        {
            "S. O'Donnell": [
                {
                    "person":         "S. O'Donnell",
                    "date":           date(2026, 9, 8),
                    "date_str":       "09/08/2026",
                    "period":         "September 1-15",
                    "project":        "GTM - NONCHG",
                    "task":           "Available Time",
                    "hours":          2.0,
                    "notes":          "Mainly looked into mass emails...",
                    "description":    "Researched more into ...",
                    "available_time": True,
                    "needs_response": True,
                },
                ...
            ],
        }

    Each person's entries are sorted with Available Time first, then by date.
    """
    import csv, io as _io

    if hasattr(file_obj, "read"):
        content = file_obj.read()
        if isinstance(content, bytes):
            content = content.decode("utf-8-sig", errors="replace")  # strip BOM
        file_obj = _io.StringIO(content)

    reader = list(csv.reader(file_obj))
    result: dict = {}

    # ---- Locate the header row and map columns by name ----
    header_idx = None
    col_project = col_date = col_employee = col_hours = None
    col_task = col_notes = col_description = None

    for i, row in enumerate(reader):
        row_lower = [str(c).strip().lower() for c in row]
        if "date" in row_lower and "employee" in row_lower:
            header_idx = i
            for j, h in enumerate(row_lower):
                if "project" in h:        col_project     = j
                elif h == "date":         col_date        = j
                elif "employee" in h:     col_employee    = j
                elif "hour" in h:         col_hours       = j
                elif h == "task":         col_task        = j
                elif h == "notes":        col_notes       = j
                elif h == "description":  col_description = j
            break

    if header_idx is None or col_employee is None:
        return result  # unrecognised format

    # The Task/Notes/Description columns only exist on the newer report. If an
    # older export is uploaded, fall back gracefully rather than blowing up.
    if col_task is None:
        return result

    for row in reader[header_idx + 1:]:
        needed = [x for x in (col_project, col_date, col_employee, col_hours, col_task)
                  if x is not None]
        if not needed or len(row) <= max(needed):
            continue

        def _cell(idx):
            return str(row[idx]).strip() if idx is not None and idx < len(row) else ""

        project_name = _cell(col_project)
        date_str     = _cell(col_date)
        employee_str = _cell(col_employee)
        hours_str    = _cell(col_hours)
        task         = _cell(col_task)

        # Footer rows ("Generated on: ...") have no employee — skipped here.
        if not employee_str or not date_str or not project_name:
            continue
        if not _is_noncharge(project_name):
            continue

        try:
            hours = float(hours_str.replace(",", ""))
        except ValueError:
            continue
        if hours <= 0:
            continue

        d = _parse_date(date_str)
        if d is None:
            continue

        # Map "LastName, FirstName" to the schedule's name key. The explicit
        # map handles ambiguous last names (two O'Donnells).
        if employee_str in OPENAIR_EMPLOYEE_MAP:
            person = OPENAIR_EMPLOYEE_MAP[employee_str]
        else:
            person = employee_str.split(",")[0].strip() if "," in employee_str else employee_str

        notes       = _clean_note(_cell(col_notes))
        description = _clean_note(_cell(col_description))

        is_available = task.strip().lower() == AVAILABLE_TIME_TASK

        # Ask for a response on Available Time always, and on any entry logged
        # with no explanation at all — except tasks where a note would be
        # meaningless (holidays, PTO, bereavement, jury duty).
        needs_response = is_available or (
            not notes
            and not description
            and task.strip().upper() not in NONCHARGE_NO_NOTE_TASKS
        )

        result.setdefault(person, []).append({
            "person":         person,
            "date":           d,
            "date_str":       d.strftime("%m/%d/%Y"),
            "period":         _date_to_period(d),
            "project":        project_name,
            "task":           task,
            "hours":          hours,
            "notes":          notes,
            "description":    description,
            "available_time": is_available,
            "needs_response": needs_response,
            "person_email":   EMAIL_LOOKUP.get(person),
            "first_name":     FIRST_NAMES.get(person, person),
        })

    # Available Time first, then chronological.
    for person in result:
        result[person].sort(key=lambda e: (not e["available_time"], e["date"]))

    return result


def filter_noncharge(data: dict, periods: list = None, people: list = None) -> dict:
    """Narrow parsed data to the selected periods / roster. Returns the same
    shape, dropping people who have nothing left."""
    out = {}
    for person, entries in data.items():
        if people and person not in people:
            continue
        rows = entries
        if periods:
            wanted = set(periods)
            rows = [e for e in rows if e["period"] in wanted]
        if rows:
            out[person] = rows
    return out


def flatten_noncharge(data: dict) -> list:
    """All entries as one flat list — for the app's table view."""
    return [e for entries in data.values() for e in entries]


def noncharge_totals(data: dict) -> list:
    """
    Per-person rollup for the tab's summary table.

    Returns a list of dicts sorted by most non-charge hours first.
    """
    totals = []
    for person, entries in data.items():
        available = sum(e["hours"] for e in entries if e["available_time"])
        pto       = sum(e["hours"] for e in entries if "PTO" in e["task"].upper())
        holiday   = sum(e["hours"] for e in entries if "HOLIDAY" in e["task"].upper())
        training  = sum(e["hours"] for e in entries if "TRAINING" in e["task"].upper())
        total     = sum(e["hours"] for e in entries)
        totals.append({
            "person":         person,
            "first_name":     FIRST_NAMES.get(person, person),
            "person_email":   EMAIL_LOOKUP.get(person),
            "available_time": available,
            "pto":            pto,
            "holiday":        holiday,
            "training":       training,
            "other":          total - available - pto - holiday - training,
            "total":          total,
            "needs_response": sum(1 for e in entries if e["needs_response"]),
        })
    totals.sort(key=lambda t: t["total"], reverse=True)
    return totals
