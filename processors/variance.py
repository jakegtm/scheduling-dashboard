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

    Expected hierarchical structure:
        Person row   — e.g. "Hendrickson, Jake" in col A
        Project row  — e.g. "  Accelleron '26 Tax..."  (indented)
        Week row     — date + hours per week column

    Returns:
        {
            "Hendrickson": {
                "Accelleron '26 Tax Department Management": {
                    "May 1-15":  8.0,
                    "May 16-31": 16.0,
                },
            },
        }
    """
    import csv, io

    if hasattr(file_obj, "read"):
        content = file_obj.read()
        if isinstance(content, bytes):
            content = content.decode("utf-8", errors="replace")
        file_obj = io.StringIO(content)

    reader = list(csv.reader(file_obj))

    result       = {}
    current_person  = None
    current_project = None

    MONTH_ABBREVS = {
        "jan": "January", "feb": "February", "mar": "March",
        "apr": "April",   "may": "May",      "jun": "June",
        "jul": "July",    "aug": "August",   "sep": "September",
        "oct": "October", "nov": "November", "dec": "December",
    }

    def _parse_date(s):
        s = s.strip()
        for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%m/%d/%y"):
            try:
                return date(*[int(x) for x in
                               __import__("datetime").datetime.strptime(s, fmt).timetuple()[:3]])
            except (ValueError, AttributeError):
                pass
        return None

    def _week_to_period(d: date) -> str:
        month_name = d.strftime("%B")
        if d.day <= 15:
            return f"{month_name} 1-15"
        else:
            last_day = (d.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)
            return f"{month_name} 16-{last_day.day}"

    for row in reader:
        if not row or not any(row):
            continue
        first_cell = str(row[0]).strip()

        # Blank first cell with data further right → skip header/total rows
        if not first_cell:
            continue

        # Detect person row: "LastName, FirstName" with a comma, not indented
        if "," in first_cell and not first_cell.startswith(" "):
            parts = first_cell.split(",")
            last_name = parts[0].strip()
            current_person  = last_name
            current_project = None
            if current_person not in result:
                result[current_person] = {}
            continue

        # Detect project row: starts with whitespace, followed by project name
        if first_cell.startswith(" ") or (row[0] != first_cell):
            project_name = first_cell.strip()
            if current_person and project_name and not re.match(r"^\d{1,2}[/\-]", project_name):
                current_project = project_name
                if current_project not in result.get(current_person, {}):
                    result.setdefault(current_person, {})[current_project] = {}
            continue

        # Detect week row: first cell is a date
        d = _parse_date(first_cell)
        if d and current_person and current_project:
            period = _week_to_period(d)
            # Sum all numeric values in the row (hours columns)
            total = 0.0
            for cell in row[1:]:
                try:
                    total += float(str(cell).replace(",", "").strip())
                except (ValueError, TypeError):
                    pass
            if total > 0:
                result[current_person][current_project][period] = (
                    result[current_person][current_project].get(period, 0.0) + total
                )

    return result


def get_available_months(actual_data: dict) -> list:
    """Return sorted list of unique period labels found in actual_data."""
    periods = set()
    for projects in actual_data.values():
        for periods_dict in projects.values():
            periods.update(periods_dict.keys())
    return sorted(periods)


def filter_by_months(actual_data: dict, selected_periods: list) -> dict:
    """Filter actual_data to only include the selected period labels."""
    selected = set(selected_periods)
    filtered = {}
    for person, projects in actual_data.items():
        filtered[person] = {}
        for proj, periods in projects.items():
            kept = {p: h for p, h in periods.items() if p in selected}
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
        period_str = str(period_val).strip() if period_val else ""
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
