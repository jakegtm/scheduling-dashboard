from __future__ import annotations
# ============================================================
# processors/variance.py — OpenAir vs Schedule Variance
# ============================================================

import re
from datetime import date, timedelta
from collections import defaultdict

from config import EMAIL_LOOKUP, FIRST_NAMES, DEFAULT_VARIANCE_MIN, DEFAULT_VARIANCE_MAX


# ================================================================
# HELPERS
# ================================================================

def _normalize(s: str) -> str:
    """Lowercase, strip punctuation and whitespace for fuzzy matching."""
    return re.sub(r"[^a-z0-9]", "", s.lower())


def _normalize_period(s: str) -> str:
    """
    Standardize half-month period labels to 'Month D1-D2' (no spaces around dash).
    Handles en-dash and em-dash too.
    'May 1 - 15' -> 'May 1-15',  'May 16 - 31' -> 'May 16-31'
    """
    return re.sub(r"\s*[-–—]\s*", "-", str(s).strip())


def _is_period_2(period_label: str) -> bool:
    """Return True if this is the second half of the month (16th onwards)."""
    m = re.search(r"(\d+)\s*[-–]\s*(\d+)", str(period_label))
    if m:
        start_day = int(m.group(1))
        return start_day >= 16
    return False


def _match_project(openair_name: str, schedule_codes: list) -> str | None:
    """
    Fuzzy-match an OpenAir project name to a schedule project code.
    Priority: 1) exact  2) normalized  3) containment (longest overlap)
    """
    oa_norm  = _normalize(openair_name)
    best     = None
    best_len = 0

    for sc in schedule_codes:
        sc_norm = _normalize(sc)
        if oa_norm == sc_norm:
            return sc
        if oa_norm in sc_norm or sc_norm in oa_norm:
            overlap = min(len(oa_norm), len(sc_norm))
            if overlap > best_len:
                best, best_len = sc, overlap

    return best


def _lookup_email(name: str) -> str | None:
    if not name:
        return None
    name = name.strip()
    if name in EMAIL_LOOKUP:
        return EMAIL_LOOKUP[name]
    for k, v in EMAIL_LOOKUP.items():
        if k.lower() == name.lower():
            return v
    return None


def _lookup_first(name: str) -> str:
    return FIRST_NAMES.get(name.strip(), name.strip()) if name else "there"


# ================================================================
# OPENAIR PARSER
# ================================================================

def parse_openair_report(file_obj) -> dict:
    """
    Parse an OpenAir time report CSV.

    Actual flat CSV format:
        Row 1: Title row (skipped)
        Row 2: Headers — "Project - Name", "Date", "Employee", "Time (Hours)"
        Row 3+: Data — one entry per row

    Employee format: "LastName, FirstName"  (last name used as key)
    Date format:     MM/DD/YYYY

    Returns:
        {
            "Wojtowicz": {
                "Braze '25 OTP Implementation": {
                    "January 1-15":  8.0,
                    "January 16-31": 16.0,
                },
            },
        }
    """
    import csv, io as _io
    from datetime import datetime as _dt

    if hasattr(file_obj, "read"):
        content = file_obj.read()
        if isinstance(content, bytes):
            content = content.decode("utf-8-sig", errors="replace")  # utf-8-sig strips BOM
        file_obj = _io.StringIO(content)

    reader  = list(csv.reader(file_obj))
    result  = {}

    def _parse_date(s: str):
        s = s.strip()
        for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%m/%d/%y"):
            try:
                return _dt.strptime(s, fmt).date()
            except ValueError:
                pass
        return None

    def _date_to_period(d) -> str:
        month_name = d.strftime("%B")
        if d.day <= 15:
            return f"{month_name} 1-15"
        last_day = (d.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)
        return f"{month_name} 16-{last_day.day}"

    # Find the header row (contains "Date" and "Employee")
    header_idx = None
    col_project = col_date = col_employee = col_hours = None

    for i, row in enumerate(reader):
        row_lower = [str(c).strip().lower() for c in row]
        if "date" in row_lower and "employee" in row_lower:
            header_idx  = i
            # Map columns by header name
            for j, h in enumerate(row_lower):
                if "project" in h:      col_project  = j
                elif h == "date":       col_date     = j
                elif "employee" in h:   col_employee = j
                elif "hour" in h:       col_hours    = j
            break

    if header_idx is None or col_employee is None:
        return result  # unrecognised format

    # Parse data rows
    for row in reader[header_idx + 1:]:
        _valid_cols = [x for x in [col_project, col_date, col_employee, col_hours]
                       if x is not None]  # safe — no filter(None.__ne__) to avoid NotImplemented
        if not _valid_cols or len(row) <= max(_valid_cols):
            continue

        project_name  = str(row[col_project]).strip()  if col_project  is not None else ""
        date_str      = str(row[col_date]).strip()     if col_date     is not None else ""
        employee_str  = str(row[col_employee]).strip() if col_employee is not None else ""
        hours_str     = str(row[col_hours]).strip()    if col_hours    is not None else ""

        if not employee_str or not date_str or not project_name:
            continue

        # Parse hours
        try:
            hours = float(hours_str.replace(",", ""))
        except ValueError:
            continue
        if hours <= 0:
            continue

        # Last name from "LastName, FirstName"
        last_name = employee_str.split(",")[0].strip() if "," in employee_str else employee_str

        # Parse date → period label
        d = _parse_date(date_str)
        if d is None:
            continue
        period = _date_to_period(d)

        # Accumulate
        result.setdefault(last_name, {}).setdefault(project_name, {})
        result[last_name][project_name][period] = (
            result[last_name][project_name].get(period, 0.0) + hours
        )

    return result


def _all_year_periods(year: int = None) -> list:
    """
    Generate all 24 standard half-month periods for the given year.
    e.g. ["January 1-15", "January 16-31", "February 1-15", ...]
    Handles leap-year February automatically.
    """
    import calendar as _cal
    from datetime import date as _date
    if year is None:
        year = _date.today().year
    result = []
    for month_num in range(1, 13):
        month_name = _cal.month_name[month_num]
        last_day   = _cal.monthrange(year, month_num)[1]
        result.append(f"{month_name} 1-15")
        result.append(f"{month_name} 16-{last_day}")
    return result


def _period_sort_key(period: str) -> tuple:
    """Return (month_num, start_day) for chronological sorting."""
    import calendar as _cal
    m = re.match(r"(\w+)\s+(\d+)", str(period))
    if not m:
        return (99, 99)
    month_str = m.group(1)
    start_day = int(m.group(2))
    for i, name in enumerate(_cal.month_name):
        if name.lower().startswith(month_str.lower()[:3]):
            return (i, start_day)
    return (99, 99)


def _classify_periods(all_periods: set, year: int = None) -> tuple:
    """
    Given a set of period strings, return:
        (sorted_periods, future_periods)
    where future_periods are those whose start date is after today.
    """
    import calendar as _cal
    from datetime import date as _date
    today = _date.today()
    if year is None:
        year = today.year

    sorted_periods = sorted(all_periods, key=_period_sort_key)
    future_periods = []

    for period in sorted_periods:
        m = re.match(r"(\w+)\s+(\d+)", str(period))
        if not m:
            continue
        month_str = m.group(1)
        start_day = int(m.group(2))
        month_num = 0
        for i, name in enumerate(_cal.month_name):
            if name.lower().startswith(month_str.lower()[:3]):
                month_num = i
                break
        if month_num == 0:
            continue
        try:
            if _date(year, month_num, start_day) > today:
                future_periods.append(period)
        except ValueError:
            pass

    return sorted_periods, future_periods


def get_available_months(actual_data: dict) -> tuple:
    """
    Return (all_periods_sorted, future_periods).

    Periods = all 24 standard half-month periods for the current year
              UNION any periods found in actual_data (OpenAir).
    Future periods are those whose start date is after today.
    """
    from datetime import date as _date
    year    = _date.today().year
    periods = set(_all_year_periods(year))

    # Add any extra periods from OpenAir data (e.g. previous year carry-over)
    for projects in actual_data.values():
        for periods_dict in projects.values():
            periods.update(periods_dict.keys())

    return _classify_periods(periods, year)


def get_schedule_periods(wb, sheet_name: str) -> tuple:
    """
    Return (all_periods_sorted, future_periods) for use when no OpenAir
    file is uploaded. Always includes all 24 standard periods for the year
    so the user can select any period to see scheduled hours.
    """
    from datetime import date as _date
    year    = _date.today().year
    periods = set(_all_year_periods(year))

    # Also union with any periods found in the actual schedule data
    sched = read_schedule_hours(wb, sheet_name)
    for projects in sched.values():
        for pdata in projects.values():
            periods.update(pdata.keys())

    return _classify_periods(periods, year)


def filter_by_months(actual_data: dict, selected_periods: list) -> dict:
    """Filter actual_data to only include the selected period labels.
    Normalizes both sides so 'May 1 - 15' and 'May 1-15' always match."""
    selected = {_normalize_period(p) for p in selected_periods}
    filtered = {}
    for person, projects in actual_data.items():
        filtered[person] = {}
        for proj, periods in projects.items():
            kept = {_normalize_period(p): h
                    for p, h in periods.items()
                    if _normalize_period(p) in selected}
            if kept:
                filtered[person][proj] = kept
    return filtered


# ================================================================
# SCHEDULE READER
# ================================================================

def read_schedule_hours(wb, sheet_name: str) -> dict:
    """
    Read scheduled hours from a month tab.

    Returns:
    {
        "Hendrickson": {
            "Accelleron '26 Tax Department Management": {
                "May 1-15":  20.0,
                "May 16-31": 30.0,
            }
        }
    }
    """
    ws = wb[sheet_name]

    PERSON_NAME_ROW  = 2
    PERIOD_LABEL_ROW = 6
    DATA_START_ROW   = 8
    PERSON_COL_START = 7
    PERSON_COL_END   = 40

    # Build column map: col → (person_name, period_label)
    col_map = {}
    for col in range(PERSON_COL_START, PERSON_COL_END + 1):
        period_val = ws.cell(row=PERIOD_LABEL_ROW, column=col).value
        period_str = _normalize_period(period_val) if period_val else ""
        pair_start = PERSON_COL_START + ((col - PERSON_COL_START) // 2) * 2
        person_val = ws.cell(row=PERSON_NAME_ROW, column=pair_start).value
        person_str = str(person_val).strip() if person_val else ""
        if person_str and period_str:
            col_map[col] = (person_str, period_str)

    schedule           = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    consecutive_blank  = 0

    for row_num in range(DATA_START_ROW, 5000):
        project_code = ws.cell(row=row_num, column=2).value
        client       = ws.cell(row=row_num, column=1).value
        if not client and not project_code:
            consecutive_blank += 1
            if consecutive_blank >= 10:
                break
            continue
        consecutive_blank = 0

        code_str = str(project_code).strip() if project_code else ""
        for col, (person, period) in col_map.items():
            val = ws.cell(row=row_num, column=col).value
            try:
                hours = float(val) if val is not None else 0.0
            except (ValueError, TypeError):
                hours = 0.0
            if hours != 0:
                schedule[person][code_str][period] += hours

    return {p: dict(proj) for p, proj in schedule.items()}


# ================================================================
# VARIANCE CALCULATOR
# ================================================================

def compute_variances(
    actual_data: dict,
    schedule_data: dict,
    min_diff: float = DEFAULT_VARIANCE_MIN,
    max_diff: float = DEFAULT_VARIANCE_MAX,
    selected_periods: list | None = None,
) -> list:
    """
    Compare per-person / per-project / per-period actuals vs schedule.

    Flags rows where:
        difference <= min_diff   (actual well below schedule)
        OR
        difference >= max_diff   (actual at or above schedule)

    Both bounds are inclusive (equal-to is flagged).
    """
    variances = []

    for person, actual_projects in actual_data.items():
        # Case-insensitive match to schedule person key
        sched_person = None
        for sp in schedule_data:
            if sp.lower() == person.lower():
                sched_person = sp
                break
        if sched_person is None:
            continue

        first_name      = _lookup_first(sched_person)
        sched_projects  = schedule_data[sched_person]
        sched_code_list = list(sched_projects.keys())

        for oa_project, actual_periods in actual_projects.items():
            matched_code  = _match_project(oa_project, sched_code_list)
            sched_periods = sched_projects.get(matched_code, {}) if matched_code else {}

            if selected_periods:
                all_periods = set(selected_periods)
            else:
                all_periods = set(actual_periods) | set(sched_periods)

            for period in all_periods:
                actual_hrs = round(actual_periods.get(period, 0.0), 1)
                sched_hrs  = round(sched_periods.get(period, 0.0), 1)
                diff       = round(actual_hrs - sched_hrs, 1)

                # Include equal-to on both bounds
                if not (diff <= min_diff or diff >= max_diff):
                    continue

                # Context-aware question
                if diff > 0:
                    question = (
                        "Do additional hours need to be added to the "
                        "schedule for this project?"
                    )
                else:
                    if _is_period_2(period):
                        question = "Are these scheduled hours still expected to hit?"
                    else:
                        question = "Do these hours need to be pushed to another period?"

                variances.append({
                    "person":       sched_person,
                    "first_name":   first_name,
                    "project_code": matched_code or oa_project,
                    "oa_name":      oa_project,
                    "period":       period,
                    "actual_hours": actual_hrs,
                    "sched_hours":  sched_hrs,
                    "difference":   diff,
                    "question":     question,
                    "person_email": _lookup_email(sched_person),
                })

    return variances
