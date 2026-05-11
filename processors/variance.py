from __future__ import annotations
# ============================================================
# processors/variance.py — OpenAir CSV vs Schedule Variance
# ============================================================

import csv
import io
import re
import calendar
from collections import defaultdict
from datetime import date

VARIANCE_THRESHOLD = 0


# ---- PERIOD LABEL NORMALIZATION ----

def _normalize_period(label: str) -> str:
    """
    Normalize period label to consistent format with no spaces around dash.
    'May 1 - 15' → 'May 1-15'
    'May 16-31'  → 'May 16-31'  (already clean)
    """
    if not label:
        return ""
    return re.sub(r'\s*[-–]\s*', '-', str(label).strip())


def _is_internal(project_name: str) -> bool:
    name = str(project_name).strip().upper()
    return name.startswith("GTM -") or name in ("", " ")


def _last_name(employee: str) -> str:
    return str(employee).split(",")[0].strip()


def _period_label(entry_date: date) -> str:
    """Return normalized period label for a date."""
    last_day = calendar.monthrange(entry_date.year, entry_date.month)[1]
    mon = entry_date.strftime("%b")
    return f"{mon} 1-15" if entry_date.day <= 15 else f"{mon} 16-{last_day}"


def _is_period_2(period_label: str) -> bool:
    m = re.search(r'(\d+)\s*[-–]', str(period_label))
    return int(m.group(1)) >= 16 if m else False


def _is_future_period(period_label: str, year: int) -> bool:
    """Return True only if the period START date is after today (no actuals possible yet)."""
    today = date.today()
    m = re.search(r'([A-Za-z]+)\s+(\d+)[-–]', str(period_label))
    if not m:
        return False
    _MONTH_MAP = {mn[:3].lower(): i for i, mn in enumerate([
        'January','February','March','April','May','June',
        'July','August','September','October','November','December'], 1)}
    mon_num   = _MONTH_MAP.get(m.group(1).lower()[:3])
    start_day = int(m.group(2))
    if not mon_num:
        return False
    try:
        start_date = date(year, mon_num, start_day)
        return start_date > today
    except ValueError:
        return False


def _normalize(s: str) -> str:
    return re.sub(r'[^a-z0-9]', '', s.lower())


def _match_project(oa_name: str, schedule_codes: list) -> str | None:
    oa_norm = _normalize(oa_name)
    best, best_len = None, 0
    for sc in schedule_codes:
        sc_norm = _normalize(sc)
        if oa_norm == sc_norm:
            return sc
        if oa_norm in sc_norm or sc_norm in oa_norm:
            overlap = min(len(oa_norm), len(sc_norm))
            if overlap > best_len:
                best, best_len = sc, overlap
    return best


# ---- OPENAIR PARSER ----

def parse_openair_report(file_obj, default_year: int = None) -> dict:
    """
    Parse OpenAir CSV. Returns:
        {last_name: {project_name: {period_label: hours}}}
    Period labels are normalized (no spaces around dash).
    """
    if default_year is None:
        default_year = date.today().year

    if hasattr(file_obj, "read"):
        content = file_obj.read()
        if isinstance(content, bytes):
            content = content.decode("utf-8-sig")
    else:
        with open(file_obj, encoding="utf-8-sig") as f:
            content = f.read()

    reader = csv.reader(io.StringIO(content))
    rows   = list(reader)
    if len(rows) < 3:
        return {}

    result = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))

    for row in rows[2:]:
        if len(row) < 4:
            continue
        project   = row[0].strip()
        date_str  = row[1].strip()
        employee  = row[2].strip()
        hours_str = row[3].strip()

        if not project or not employee or not date_str:
            continue
        if _is_internal(project):
            continue

        try:
            parts      = date_str.split("/")
            entry_date = date(int(parts[2]), int(parts[0]), int(parts[1]))
        except (ValueError, IndexError):
            continue

        try:
            hours = float(hours_str)
        except ValueError:
            continue

        if hours == 0:
            continue

        last   = _last_name(employee)
        period = _period_label(entry_date)   # already normalized
        result[last][project][period] += hours

    return {p: {proj: dict(periods) for proj, periods in projs.items()}
            for p, projs in result.items()}


# ---- SCHEDULE READER ----

def read_schedule_hours(wb, sheet_name: str) -> dict:
    """
    Read scheduled hours from a month tab.
    Period labels are normalized to match OpenAir format.
    Returns: {last_name: {project_code: {period_label: hours}}}
    """
    ws = wb[sheet_name]
    PERSON_NAME_ROW  = 2
    PERIOD_LABEL_ROW = 6
    DATA_START_ROW   = 8
    PERSON_COL_START = 7
    PERSON_COL_END   = 32

    col_map = {}
    for col in range(PERSON_COL_START, PERSON_COL_END + 1):
        period_val = ws.cell(row=PERIOD_LABEL_ROW, column=col).value
        # Normalize: 'May 1 - 15' → 'May 1-15'
        period_str = _normalize_period(str(period_val)) if period_val else ""
        pair_start = PERSON_COL_START + ((col - PERSON_COL_START) // 2) * 2
        person_val = ws.cell(row=PERSON_NAME_ROW, column=pair_start).value
        person_str = str(person_val).strip() if person_val else ""
        if person_str and period_str:
            col_map[col] = (person_str, period_str)

    schedule = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    consec   = 0

    for row_num in range(DATA_START_ROW, 5000):
        project_code = ws.cell(row=row_num, column=2).value
        client       = ws.cell(row=row_num, column=1).value
        if not client and not project_code:
            consec += 1
            if consec >= 10:
                break
            continue
        consec = 0
        code_str = str(project_code).strip() if project_code else ""
        for col, (person, period) in col_map.items():
            val = ws.cell(row=row_num, column=col).value
            try:
                hours = float(val) if val is not None else 0.0
            except (ValueError, TypeError):
                hours = 0.0
            if hours:
                schedule[person][code_str][period] += hours

    return {p: {proj: dict(periods) for proj, periods in projs.items()}
            for p, projs in schedule.items()}


# ---- AVAILABLE PERIODS ----

def get_available_months(actual_data: dict, year: int = None) -> tuple[list, list]:
    """
    Return (all_periods, future_periods) for the given year.
    all_periods: sorted list of all period labels that appear in actual data
    future_periods: subset of all_periods where end date > today
    """
    if year is None:
        year = date.today().year

    # Collect periods from actual data
    actual_periods = set()
    for projects in actual_data.values():
        for periods in projects.values():
            for p in periods:
                actual_periods.add(_normalize_period(p))

    _MONTH_ORDER = {mn[:3]: i for i, mn in enumerate([
        'January','February','March','April','May','June',
        'July','August','September','October','November','December'], 1)}

    def _sort_key(p):
        parts = p.split()
        mon = _MONTH_ORDER.get(parts[0][:3], 99) if parts else 99
        m = re.search(r'(\d+)', p)
        day = int(m.group(1)) if m else 0
        return (mon, day)

    sorted_periods = sorted(actual_periods, key=_sort_key)
    future = [p for p in sorted_periods if _is_future_period(p, year)]

    return sorted_periods, future


def filter_by_months(actual_data: dict, selected_periods: list) -> dict:
    """Filter actual_data to only selected (normalized) periods."""
    selected = {_normalize_period(p) for p in selected_periods}
    result = {}
    for person, projects in actual_data.items():
        filtered_projects = {}
        for proj, periods in projects.items():
            filtered = {_normalize_period(p): h
                        for p, h in periods.items()
                        if _normalize_period(p) in selected}
            if filtered:
                filtered_projects[proj] = filtered
        if filtered_projects:
            result[person] = filtered_projects
    return result


# ---- VARIANCE CALCULATOR ----

def compute_variances(actual_data: dict, schedule_data: dict,
                      selected_periods: list = None,
                      threshold: float = VARIANCE_THRESHOLD) -> list:
    """
    Compare actuals to schedule per person/project/period.
    selected_periods: if provided, only include these periods (both sides).
    """
    from processors.lookup import lookup_by_openair, lookup_first_name

    selected = ({_normalize_period(p) for p in selected_periods}
                if selected_periods else None)

    variances = []

    for last_name, actual_projects in actual_data.items():
        sched_key = lookup_by_openair(last_name)
        if sched_key is None:
            for sp in schedule_data:
                if sp.lower() == last_name.lower():
                    sched_key = sp
                    break
        if not sched_key or sched_key not in schedule_data:
            continue

        first_name      = lookup_first_name(sched_key)
        sched_projects  = schedule_data[sched_key]
        sched_code_list = list(sched_projects.keys())

        for oa_project, actual_periods in actual_projects.items():
            matched_code  = _match_project(oa_project, sched_code_list)
            sched_periods = sched_projects.get(matched_code, {}) if matched_code else {}

            # Normalize all period keys
            actual_norm = {_normalize_period(p): h for p, h in actual_periods.items()}
            sched_norm  = {_normalize_period(p): h for p, h in sched_periods.items()}

            # Only include selected periods
            all_periods = set(actual_norm) | set(sched_norm)
            if selected:
                all_periods = all_periods & selected

            for period in all_periods:
                actual_hrs = round(actual_norm.get(period, 0.0), 1)
                sched_hrs  = round(sched_norm.get(period, 0.0), 1)
                diff       = round(actual_hrs - sched_hrs, 1)

                if abs(diff) <= threshold:
                    continue

                if diff > 0:
                    question = ("Do additional hours need to be added "
                                "to the schedule for this project?")
                elif _is_period_2(period):
                    question = "Are these scheduled hours still expected to hit?"
                else:
                    question = "Do these hours need to be pushed to another period?"

                variances.append({
                    "person":       sched_key,
                    "first_name":   first_name,
                    "project_code": matched_code or oa_project,
                    "oa_name":      oa_project,
                    "period":       period,
                    "actual_hours": actual_hrs,
                    "sched_hours":  sched_hrs,
                    "difference":   diff,
                    "question":     question,
                    "is_future":    _is_future_period(period, date.today().year),
                })

    return variances
