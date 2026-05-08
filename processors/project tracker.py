# ============================================================
# processors/project_tracker.py
# ============================================================
# Project Tracker — verified column layout:
#   A(1):  Client
#   B(2):  Project Code  <- "TBD" check here
#   C(3):  Status (Known/Unknown) <- filter to Known only
#   H(8):  Project Owner  <- email target
#   I(9):  Budget         <- flag if blank/zero
#   J(10): Intern rate
#   K(11): Analyst rate
#   L(12): Senior Analyst rate
#   M(13): Supervisor rate
#   N(14): Manager rate
#   O(15): Senior Manager rate
#   P(16): Director rate
#   Q(17): Managing Director rate
#
# Rules (Known non-TBD projects only):
#   - Missing budget: col I is blank or 0
#   - Missing rates:  ANY of J:Q is blank
# ============================================================

from collections import defaultdict
from config import EMAIL_LOOKUP

COL_CLIENT      = 1
COL_CODE        = 2
COL_STATUS      = 3
COL_OWNER       = 8
COL_BUDGET      = 9
COL_RATES_START = 10   # J — Intern
COL_RATES_END   = 17   # Q — Managing Director

RATE_LABELS = [
    "Intern", "Analyst", "Senior Analyst", "Supervisor",
    "Manager", "Senior Manager", "Director", "Managing Director"
]


def _to_float(val):
    try:
        return float(val) if val is not None else None
    except (ValueError, TypeError):
        return None


def _lookup_email(name):
    if not name:
        return None
    name = str(name).strip()
    if name in EMAIL_LOOKUP:
        return EMAIL_LOOKUP[name]
    for key, email in EMAIL_LOOKUP.items():
        if key.lower() == name.lower():
            return email
    return None


def process_project_tracker(ws) -> tuple:
    """
    Returns:
        issues       — list of Known non-TBD projects with missing budget or any missing rate
        tbd_projects — list of TBD projects (for reference)
    """
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

        # TBD: project code contains "TBD"
        if "TBD" in project_code_str.upper():
            tbd_projects.append({
                "client": client,
                "project_code": project_code_str,
                "owner": owner_str,
                "budget": budget or 0.0,
            })
            continue

        # Filter to Known only (treat blank as Known for backward compat)
        if status_str.lower() == "unknown":
            continue

        # Collect all problems for this row
        problems = []

        # Missing budget
        if budget is None or budget == 0:
            problems.append("Missing budget")

        # Any missing rate in J:Q
        rate_values = row[COL_RATES_START - 1 : COL_RATES_END]
        for label, val in zip(RATE_LABELS, rate_values):
            if val is None or str(val).strip() == "":
                problems.append(f"Missing {label} Rate")

        if problems:
            issues.append({
                "client":       client,
                "project_code": project_code_str,
                "owner":        owner_str,
                "owner_email":  _lookup_email(owner_str),
                "budget":       budget or 0.0,
                "problems":     problems,
            })

    return issues, tbd_projects


def build_tracker_emails(issues: list, cc_email: str) -> list:
    grouped = defaultdict(list)
    for issue in issues:
        grouped[issue["owner"]].append(issue)

    emails = []
    for owner, owner_issues in grouped.items():
        owner_email = owner_issues[0].get("owner_email")
        if not owner_email:
            continue
        emails.append({
            "to":      owner_email,
            "subject": "Scheduling Review — Action Required",
            "owner":   owner,
            "issues":  owner_issues,
            "section": "tracker",
        })
    return emails
