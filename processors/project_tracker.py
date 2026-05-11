from __future__ import annotations
# ============================================================
# processors/project_tracker.py
# ============================================================
# A(1): Client  B(2): Project Code (TBD check)  C(3): Status
# H(8): Project Owner  I(9): Budget  J:Q(10:17): Rates
# ============================================================

from collections import defaultdict
from processors.lookup import lookup_email, lookup_first_name

COL_CLIENT      = 1
COL_CODE        = 2
COL_STATUS      = 3
COL_OWNER       = 8
COL_BUDGET      = 9
COL_RATES_START = 10
COL_RATES_END   = 17

RATE_LABELS = [
    "Intern", "Analyst", "Senior Analyst", "Supervisor",
    "Manager", "Senior Manager", "Director", "Managing Director"
]


def _to_float(val):
    try:
        return float(val) if val is not None else None
    except (ValueError, TypeError):
        return None


def process_project_tracker(ws) -> tuple:
    issues = []
    tbd_projects = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        row = list(row) + [None] * 10

        client       = row[COL_CLIENT - 1]
        project_code = row[COL_CODE - 1]
        status       = row[COL_STATUS - 1]
        owner        = row[COL_OWNER - 1]
        budget       = _to_float(row[COL_BUDGET - 1])

        if not client and not project_code:
            continue

        project_code_str = str(project_code).strip() if project_code else ""
        status_str       = str(status).strip() if status else ""
        owner_str        = str(owner).strip() if owner else ""

        if "TBD" in project_code_str.upper():
            tbd_projects.append({
                "client": client, "project_code": project_code_str,
                "owner": owner_str, "budget": budget or 0.0,
            })
            continue

        if status_str.lower() == "unknown":
            continue

        problems = []
        if budget is None or budget == 0:
            problems.append("Missing budget")

        rate_values = row[COL_RATES_START - 1 : COL_RATES_END]
        for label, val in zip(RATE_LABELS, rate_values):
            if val is None or str(val).strip() == "":
                problems.append(f"Missing {label} Rate")

        if problems:
            issues.append({
                "client":      client,
                "project_code": project_code_str,
                "owner":       owner_str,
                "owner_email": lookup_email(owner_str),
                "owner_first": lookup_first_name(owner_str),
                "budget":      budget or 0.0,
                "problems":    problems,
            })

    return issues, tbd_projects
