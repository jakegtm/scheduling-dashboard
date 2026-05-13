from __future__ import annotations
# ============================================================
# processors/project_tracker.py
# ============================================================
# Column layout (verified against real file):
#   A(1):  Client
#   B(2):  Project Code   <- "TBD" check here
#   C(3):  Status         <- filter to Known only
#   H(8):  Project Owner  <- email target
#   I(9):  Budget
#   J-Q (10-17): Rates (Intern → Managing Director)
#
# Rules:
#   - TBD (project code contains "TBD") → collect separately
#   - Pending SOW (status is "Pending SOW") → treat same as TBD
#   - Known non-TBD → check for missing rates; flag if any are blank
#   - All other statuses (Unknown, Closed, blank) → skip
# ============================================================

from collections import defaultdict
from config import EMAIL_LOOKUP, FIRST_NAMES

COL_CLIENT      = 1   # A
COL_CODE        = 2   # B
COL_STATUS      = 3   # C
COL_OWNER       = 8   # H
COL_BUDGET      = 9   # I
COL_RATES_START = 10  # J — Intern
COL_RATES_END   = 17  # Q — Managing Director

RATE_LABELS = [
    "Intern", "Analyst", "Senior Analyst", "Supervisor",
    "Manager", "Senior Manager", "Director", "Managing Director",
]

# Statuses that are treated the same as TBD (excluded from rate checks,
# included in the TBD/Pending SOW email section)
TBD_STATUSES = {"tbd", "pending sow"}


def _to_float(val):
    try:
        return float(val) if val is not None else None
    except (ValueError, TypeError):
        return None


def _lookup_email(name: str):
    if not name:
        return None
    name = str(name).strip()
    if name in EMAIL_LOOKUP:
        return EMAIL_LOOKUP[name]
    for key, email in EMAIL_LOOKUP.items():
        if key.lower() == name.lower():
            return email
    return None


def _lookup_first(name: str) -> str:
    if not name:
        return "there"
    return FIRST_NAMES.get(name.strip(), name.strip())


def process_project_tracker(ws) -> tuple:
    """
    Returns:
        issues       — Known projects with any blank rate in J:Q
        tbd_projects — TBD and Pending SOW projects (for reference / email section)
    """
    issues       = []
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

        code_str   = str(project_code).strip() if project_code else ""
        status_str = str(status).strip()        if status       else ""
        owner_str  = str(owner).strip()         if owner        else ""

        # --- TBD / Pending SOW ---
        # Either the project code contains "TBD" OR the status is a TBD-like value
        is_tbd = (
            "TBD" in code_str.upper()
            or status_str.lower() in TBD_STATUSES
        )
        if is_tbd:
            tbd_projects.append({
                "client":       client,
                "project_code": code_str,
                "status":       status_str,
                "owner":        owner_str,
                "owner_email":  _lookup_email(owner_str),
                "owner_first":  _lookup_first(owner_str),
                "budget":       budget or 0.0,
            })
            continue

        # --- Known only ---
        if status_str.lower() != "known":
            continue

        # Check for any blank rate in J:Q
        rate_values = row[COL_RATES_START - 1: COL_RATES_END]
        missing_labels = [
            label for label, val in zip(RATE_LABELS, rate_values)
            if val is None or str(val).strip() == ""
        ]

        if missing_labels:
            issues.append({
                "client":        client,
                "project_code":  code_str,
                "owner":         owner_str,
                "owner_email":   _lookup_email(owner_str),
                "owner_first":   _lookup_first(owner_str),
                "budget":        budget or 0.0,
                "missing_rates": missing_labels,
                "problems":      [f"Missing rate(s): {', '.join(missing_labels)}"],
            })

    return issues, tbd_projects


def build_tracker_emails(issues: list, tbd_projects: list,
                         sender_name: str = "Jake") -> list:
    """
    Build one email per project owner covering:
      - Missing rates section (from issues)
      - TBD / Pending SOW section (from tbd_projects)
    """
    # Group by owner
    owner_issues = defaultdict(list)
    for issue in issues:
        owner_issues[issue["owner"]].append(issue)

    owner_tbd = defaultdict(list)
    for proj in tbd_projects:
        owner_tbd[proj["owner"]].append(proj)

    all_owners = set(owner_issues) | set(owner_tbd)
    emails = []

    for owner in all_owners:
        # Find email
        email = None
        if owner_issues[owner]:
            email = owner_issues[owner][0].get("owner_email")
        if not email and owner_tbd[owner]:
            email = owner_tbd[owner][0].get("owner_email")
        if not email:
            continue

        first = _lookup_first(owner)
        sections = []

        # Missing rates
        if owner_issues[owner]:
            lines = [
                "The following projects assigned to you are missing billing rates. "
                "Please review and update as soon as possible.\n"
            ]
            for issue in owner_issues[owner]:
                lines.append(
                    f"  • {issue['client']} — {issue['project_code']}\n"
                    f"    Missing: {', '.join(issue['missing_rates'])}"
                )
            sections.append("\n".join(lines))

        # TBD / Pending SOW
        if owner_tbd[owner]:
            lines = [
                "The following projects currently have TBD or Pending SOW budgets. "
                "If you have any updates on these, please reply with the latest — "
                "otherwise, no action is needed.\n"
            ]
            for proj in owner_tbd[owner]:
                status_label = f" [{proj['status']}]" if proj["status"] else " [TBD]"
                lines.append(
                    f"  • {proj['client']} — {proj['project_code']}{status_label}"
                )
            sections.append("\n".join(lines))

        body = (
            f"Hi {first},\n\n"
            + "\n\n".join(sections)
            + f"\n\nBest,\n{sender_name}"
        )

        emails.append({
            "to":      email,
            "subject": "Project Tracker — Review Required",
            "owner":   owner,
            "body":    body,
            "section": "tracker",
        })

    return emails
