from __future__ import annotations
# ============================================================
# processors/budget_actual.py
# ============================================================
# Column layout (verified against real file):
#   A(1):  Client
#   B(2):  Project Code
#   C(3):  Status  <- filter to Known only
#   H(8):  Project Owner  <- email target
#   I(9):  Budget Amount
#   L(12): Remaining  <- flag if negative OR too high
#
# Rules (Known projects only):
#   "negative"      : remaining < -neg_thresh
#   "not_projected" : remaining > budget_thresh
# ============================================================

from config import EMAIL_LOOKUP, FIRST_NAMES

COL_CLIENT    = 1   # A
COL_CODE      = 2   # B
COL_STATUS    = 3   # C
COL_OWNER     = 8   # H
COL_BUDGET    = 9   # I
COL_REMAINING = 12  # L


def _to_float(val):
    try:
        return float(val) if val is not None else None
    except (ValueError, TypeError):
        return None


def _lookup_email(name: str) -> str | None:
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
    name = str(name).strip()
    if name in FIRST_NAMES:
        return FIRST_NAMES[name]
    for key, first in FIRST_NAMES.items():
        if key.lower() == name.lower():
            return first
    return name.split()[0] if name else "there"


def process_budget_actual(
    ws,
    budget_thresh: float = 20000,
    proj_pct: float = 0.80,
    neg_thresh: float = 100,
) -> list:
    """
    Scan the Budget to Actual sheet and return flagged issues.

    Each issue dict contains:
        client, project_code, owner, owner_email, owner_first,
        budget, remaining, type ("negative" | "not_projected"),
        description
    """
    issues = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        row = list(row) + [None] * 20

        client    = row[COL_CLIENT - 1]
        code      = row[COL_CODE - 1]
        status    = row[COL_STATUS - 1]
        owner     = row[COL_OWNER - 1]
        budget    = _to_float(row[COL_BUDGET - 1])
        remaining = _to_float(row[COL_REMAINING - 1])

        # Skip blank rows
        if not client and not code:
            continue

        status_str = str(status).strip().lower() if status else ""

        # Only process Known projects
        if status_str != "known":
            continue

        if budget is None or remaining is None:
            continue

        owner_str = str(owner).strip() if owner else ""

        # Flag over-budget (remaining is very negative)
        if remaining < -neg_thresh:
            issues.append({
                "client":       client,
                "project_code": str(code).strip() if code else "",
                "owner":        owner_str,
                "owner_email":  _lookup_email(owner_str),
                "owner_first":  _lookup_first(owner_str),
                "budget":       budget,
                "remaining":    remaining,
                "type":         "negative",
                "description":  f"Over budget by ${abs(remaining):,.0f}",
            })

        # Flag under-scheduled (remaining is large relative to budget)
        elif budget > 0 and remaining >= budget_thresh:
            issues.append({
                "client":       client,
                "project_code": str(code).strip() if code else "",
                "owner":        owner_str,
                "owner_email":  _lookup_email(owner_str),
                "owner_first":  _lookup_first(owner_str),
                "budget":       budget,
                "remaining":    remaining,
                "type":         "not_projected",
                "description":  f"${remaining:,.0f} unscheduled remaining",
            })

    return issues
