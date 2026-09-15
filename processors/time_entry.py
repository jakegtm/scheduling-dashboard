from __future__ import annotations
# ============================================================
# processors/time_entry.py
# ============================================================
# Detects who hasn't entered their time in OpenAir for the
# previous week, so the Current Month Hours tab can carry a
# reminder.
#
# Reads the same time report as variance.py and noncharge.py.
# Counts ALL time — chargeable and non-charge alike — since
# holiday and PTO hours are entered time too.
#
# The week checked is the last full Mon-Sun week before the
# report was generated, NOT before today. A report pulled last
# Friday shouldn't make the whole team look delinquent for a
# week it never covered.
# ============================================================

from datetime import datetime, timedelta, date
from config import (
    EMAIL_LOOKUP,
    FIRST_NAMES,
    OPENAIR_EMPLOYEE_MAP,
    TIME_ENTRY_EXPECTED_WEEKLY_HOURS,
    TIME_ENTRY_MIN_HOURS_TO_COUNT,
)


def _parse_date(s: str):
    s = s.strip()
    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%m/%d/%y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def parse_time_coverage(file_obj) -> tuple[dict, date | None]:
    """
    Read every time entry in the report and total hours per person per day.

    Returns:
        (
            {"S. O'Donnell": {date(2026, 9, 8): 8.0, ...}, ...},
            date(2026, 9, 14),   # the report's "Generated on" date, or the
                                 # latest entry date if that footer is absent
        )
    """
    import csv, io as _io

    if hasattr(file_obj, "read"):
        content = file_obj.read()
        if isinstance(content, bytes):
            content = content.decode("utf-8-sig", errors="replace")
        file_obj = _io.StringIO(content)

    reader = list(csv.reader(file_obj))
    coverage: dict = {}
    generated_on = None
    latest_entry = None

    header_idx = None
    col_date = col_employee = col_hours = None

    for i, row in enumerate(reader):
        row_lower = [str(c).strip().lower() for c in row]
        if "date" in row_lower and "employee" in row_lower:
            header_idx = i
            for j, h in enumerate(row_lower):
                if h == "date":       col_date     = j
                elif "employee" in h: col_employee = j
                elif "hour" in h:     col_hours    = j
            break

    if header_idx is None or col_employee is None or col_date is None:
        return coverage, None

    for row in reader[header_idx + 1:]:
        if not row:
            continue

        # Footer: "Generated on: 09/14/2026 08:06 AM"
        first = str(row[0]).strip() if row else ""
        if first.lower().startswith("generated on"):
            token = first.split(":", 1)[-1].strip().split(" ")[0]
            generated_on = _parse_date(token) or generated_on
            continue

        needed = [x for x in (col_date, col_employee, col_hours) if x is not None]
        if len(row) <= max(needed):
            continue

        employee_str = str(row[col_employee]).strip()
        date_str     = str(row[col_date]).strip()
        hours_str    = str(row[col_hours]).strip() if col_hours is not None else ""

        if not employee_str or not date_str:
            continue

        d = _parse_date(date_str)
        if d is None:
            continue

        try:
            hours = float(hours_str.replace(",", ""))
        except ValueError:
            continue

        if employee_str in OPENAIR_EMPLOYEE_MAP:
            person = OPENAIR_EMPLOYEE_MAP[employee_str]
        else:
            person = employee_str.split(",")[0].strip() if "," in employee_str else employee_str

        coverage.setdefault(person, {})
        coverage[person][d] = coverage[person].get(d, 0.0) + hours

        if latest_entry is None or d > latest_entry:
            latest_entry = d

    return coverage, (generated_on or latest_entry)


def previous_week(reference: date = None) -> tuple[date, date]:
    """The last full Mon-Sun week before the week containing `reference`."""
    if reference is None:
        reference = date.today()
    this_monday = reference - timedelta(days=reference.weekday())
    start = this_monday - timedelta(days=7)
    return start, start + timedelta(days=6)


def format_week(start: date, end: date) -> str:
    return f"{start.strftime('%m/%d/%Y')} - {end.strftime('%m/%d/%Y')}"


def find_missing_time(
    coverage: dict,
    roster,
    week_start: date,
    week_end: date,
    expected_hours: float = None,
) -> dict:
    """
    Work out who is short on time entry for the given week.

    `roster` is the list of people who should have entered time — people with
    no rows at all in the report won't appear in `coverage`, so the roster is
    what surfaces them.

    Returns {person: {"status", "hours", "days", "period", "message"}} for
    anyone who is missing or short. People who are fully entered are omitted.
    """
    if expected_hours is None:
        expected_hours = TIME_ENTRY_EXPECTED_WEEKLY_HOURS

    period = format_week(week_start, week_end)
    results = {}

    for person in roster:
        days = coverage.get(person, {})
        week_days = {
            d: h for d, h in days.items()
            if week_start <= d <= week_end and h >= TIME_ENTRY_MIN_HOURS_TO_COUNT
        }
        total = sum(week_days.values())

        if not week_days:
            status = "none"
            message = (
                f"Your time was not entered in OpenAir for the period of {period}. "
                f"Please enter and submit as soon as possible."
            )
        elif total < expected_hours:
            status = "partial"
            message = (
                f"Only {total:g} of {expected_hours:g} hours were entered in OpenAir "
                f"for the period of {period}. Please review and submit the missing "
                f"time as soon as possible."
            )
        else:
            continue  # fully entered, nothing to say

        results[person] = {
            "person":      person,
            "first_name":  FIRST_NAMES.get(person, person),
            "person_email": EMAIL_LOOKUP.get(person),
            "status":      status,
            "hours":       total,
            "days":        len(week_days),
            "expected":    expected_hours,
            "period":      period,
            "message":     message,
        }

    return results
