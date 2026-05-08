# ============================================================
# processors/budget_actual.py
# ============================================================
# Budget to Actual — verified column layout:
#   A(1):  Client
#   B(2):  Project Code
#   H(8):  Project Owner  <- email target
#   I(9):  Budget Amount
#   J(10): Status (Known/Unknown)  <- filter to Known only
#   L(12): Remaining               <- flag if negative OR > threshold
#
# Rules (Known projects only):
#   "negative"      : remaining < 0
#   "not_projected" : remaining > threshold (default $20,000)
# ============================================================

from collections import defaultdict
from config import EMAIL_LOOKUP

COL_CLIENT    = 1
COL_CODE      = 2
COL_OWNER     = 8
COL_BUDGET    = 9
COL_STATUS    = 10   # J
COL_REMAINING = 12   # L


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


def process_budget_actual(ws, budget_threshold: float = 20000) -> list:
    """
    Return list of flagged issue dicts for Known projects only.
    Each issue has: client, project_code, owner, owner_email,
                    budget, remaining, type, description
    """
    issues = []
    consecutive_blank = 0

    for row in ws.iter_rows(min_row=3, max_row=5000, values_only=True):
        row = list(row) + [None] * 20

        project_code = row[COL_CODE - 1]
        if not project_code or str(project_code).strip() == "":
            consecutive_blank += 1
            if consecutive_blank >= 10:
                break
            continue
        consecutive_blank = 0

        status    = str(row[COL_STATUS - 1]).strip() if row[COL_STATUS - 1] else ""
        # Filter: only process Known projects (skip Unknown; treat blank as Known)
        if status.lower() == "unknown":
            continue

        client    = row[COL_CLIENT - 1]
        owner     = str(row[COL_OWNER - 1]).strip() if row[COL_OWNER - 1] else ""
        budget    = _to_float(row[COL_BUDGET - 1]) or 0.0
        remaining = _to_float(row[COL_REMAINING - 1])

        if remaining is None:
            continue

        base = dict(
            client=client,
            project_code=str(project_code).strip(),
            owner=owner,
            owner_email=_lookup_email(owner),
            budget=budget,
            remaining=remaining,
        )

        if remaining < 0:
            issues.append({**base,
                "type": "negative",
                "description": f"Negative budget of ({abs(remaining):,.0f})",
            })
        elif remaining > budget_threshold:
            issues.append({**base,
                "type": "not_projected",
                "description": f"Remaining unscheduled budget of {remaining:,.0f}",
            })

    return issues


def build_budget_emails(issues: list, cc_email: str) -> list:
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
            "section": "budget",
        })
    return emails
