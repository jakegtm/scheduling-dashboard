from __future__ import annotations
# ============================================================
# processors/variance.py — OpenAir CSV vs Schedule Variance
# ============================================================
# OpenAir CSV: Project-Name, Date(MM/DD/YYYY), Employee(Last,First), Hours
# Row 1 = title, Row 2 = headers, Row 3+ = data
# Internal GTM rows (GTM - *) are skipped.
# Data is day-by-day so periods are exact — no proration needed.
#
# Variance question logic:
#   actual > scheduled  → "Do additional hours need to be added?"
#   actual < scheduled, period 1 (1-15)  → "Push to another period?"
#   actual < scheduled, period 2 (16-end) → "Still expected to hit?"
# ============================================================

import csv
import io
import re
from collections import defaultdict
from datetime import date

VARIANCE_THRESHOLD = 3


def _is_internal(project_name: str) -> bool:
    name = str(project_name).strip().upper()
    return name.startswith("GTM -") or name in ("", " ")


def _last_name(employee: str) -> str:
    return str(employee).split(",")[0].strip()


def _period_label(entry_date: date) -> str:
    import calendar
    last_day = calendar.monthrange(entry_date.year, entry_date.month)[1]
    mon = entry_date.strftime("%b")
    return f"{mon} 1-15" if entry_date.day <= 15 else f"{mon} 16-{last_day}"


def _is_period_2(period_label: str) -> bool:
    m = re.search(r'(\d+)\s*[-–]', str(period_label))
    return int(m.group(1)) >= 16 if m else False


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


def parse_openair_report(file_obj, default_year: int = None) -> dict:
    """
    Parse OpenAir CSV. Returns:
        {last_name: {project_name: {period_label: hours}}}
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

    # Row 0 = title, Row 1 = headers, Row 2+ = data
    result = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))

    for row in rows[2:]:
        if len(row) < 4:
            continue
        project  = row[0].strip()
        date_str = row[1].strip()
        employee = row[2].strip()
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
        period = _period_label(entry_date)
        result[last][project][period] += hours

    return {p: {proj: dict(periods) for proj, periods in projs.items()}
            for p, projs in result.items()}


def read_schedule_hours(wb, sheet_name: str) -> dict:
    """
    Read scheduled hours from a month tab.
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
        period_str = str(period_val).strip() if period_val else ""
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


def compute_variances(actual_data: dict, schedule_data: dict,
                      threshold: float = VARIANCE_THRESHOLD) -> list:
    """
    Compare actuals to schedule. Uses lookup_by_openair for name matching.
    """
    from processors.lookup import lookup_by_openair, lookup_first_name
    variances = []

    for last_name, actual_projects in actual_data.items():
        # Match OpenAir last name → schedule file key
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
            all_periods   = set(actual_periods) | set(sched_periods)

            for period in all_periods:
                actual_hrs = round(actual_periods.get(period, 0.0), 1)
                sched_hrs  = round(sched_periods.get(period, 0.0), 1)
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
                })

    return variances


def get_available_months(actual_data: dict) -> list:
    """Return sorted unique period labels from actual data."""
    months = set()
    for projects in actual_data.values():
        for periods in projects.values():
            for period in periods:
                months.add(period)

    _MONTH_ORDER = {m[:3]: i for i, m in enumerate([
        'January','February','March','April','May','June',
        'July','August','September','October','November','December'], 1)}

    def _sort_key(p):
        parts = p.split()
        mon = _MONTH_ORDER.get(parts[0][:3], 99) if parts else 99
        m = re.search(r'(\d+)', p)
        day = int(m.group(1)) if m else 0
        return (mon, day)

    return sorted(months, key=_sort_key)


def filter_by_months(actual_data: dict, selected_periods: list) -> dict:
    """Filter actual_data to only the selected period labels."""
    selected = set(selected_periods)
    result = {}
    for person, projects in actual_data.items():
        filtered_projects = {}
        for proj, periods in projects.items():
            filtered = {p: h for p, h in periods.items() if p in selected}
            if filtered:
                filtered_projects[proj] = filtered
        if filtered_projects:
            result[person] = filtered_projects
    return result
