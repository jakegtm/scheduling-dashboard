from __future__ import annotations
# ============================================================
# processors/budget_actual.py
# ============================================================
# A(1): Client  B(2): Project Code  C(3): Status (Known only)
# H(8): Project Owner  I(9): Budget  L(12): Remaining
#
# Flags:
#   "negative"      : remaining < -negative_threshold  (default -$100)
#   "not_projected" : remaining > unscheduled_threshold (default $20k)
# ============================================================

from collections import defaultdict
from processors.lookup import lookup_email, lookup_first_name

COL_CLIENT    = 1
COL_CODE      = 2
COL_STATUS    = 3
COL_OWNER     = 8
COL_BUDGET    = 9
COL_REMAINING = 12


def _to_float(val):
    try:
        return float(val) if val is not None else None
    except (ValueError, TypeError):
        return None


def process_budget_actual(ws,
                          unscheduled_threshold: float = 20000,
                          negative_threshold: float = 100) -> list:
    """
    negative_threshold: flag if remaining < -negative_threshold
                        (positive number, e.g. 100 → flags anything below -$100)
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

        status = str(row[COL_STATUS - 1]).strip() if row[COL_STATUS - 1] else ""
        if status.lower() != "known":
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
            owner_email=lookup_email(owner),
            owner_first=lookup_first_name(owner),
            budget=budget,
            remaining=remaining,
        )

        if remaining < -negative_threshold:
            issues.append({**base,
                "type": "negative",
                "description": f"Over budget by ${abs(remaining):,.0f}",
            })
        elif remaining > unscheduled_threshold:
            issues.append({**base,
                "type": "not_projected",
                "description": f"${remaining:,.0f} unscheduled",
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
            "to": owner_email, "owner": owner, "issues": owner_issues,
        })
    return emails
