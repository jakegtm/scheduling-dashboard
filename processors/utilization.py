from __future__ import annotations
# ============================================================
# processors/utilization.py
# ============================================================
# Reads the "Utilization by Month" tab.
#
# Actual column layout (verified against real file):
#   A (0): Date marker (datetime object, e.g. 2026-05-01)
#   B (1): Month name OR Role code OR "# of Days"
#   C (2): Person last name OR "Total"
#   D (3): blank
#   E (4): Chargeable hours
#   F (5): Holiday hours
#   G (6): PTO hours
#   H (7): Month Total hours
#   I (8): Remaining hours
#   J (9): Utilization % (stored as decimal, e.g. 0.607)
#   K (10): Goal % (stored as decimal)
#   L (11): Difference % (stored as decimal)
#
# Section structure:
#   Row 1: [datetime, "MonthName", ...]  <- section start
#   Row 2: [None, "# of Days", N, ...]   <- header / day count
#   Rows: data rows (one per person)
#   Last: [None, None, "Total", ...]     <- end of section
# ============================================================

from datetime import date
from config import EMAIL_LOOKUP, FIRST_NAMES

UTIL_SHEET_KEYWORDS = ["utilization by month", "util by month", "utilization"]

COL_ROLE        = 1   # B
COL_PERSON      = 2   # C
COL_CHARGEABLE  = 4   # E
COL_HOLIDAY     = 5   # F
COL_PTO         = 6   # G
COL_MONTH_TOTAL = 7   # H
COL_REMAINING   = 8   # I
COL_UTILIZATION = 9   # J  (stored as decimal 0-1)
COL_GOAL        = 10  # K  (stored as decimal 0-1)
COL_DIFFERENCE  = 11  # L  (stored as decimal)


def _find_util_sheet(wb) -> str | None:
    for name in wb.sheetnames:
        if any(kw in name.lower() for kw in UTIL_SHEET_KEYWORDS):
            return name
    return None


def _to_float(v) -> float | None:
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace("%", "").strip())
    except (ValueError, TypeError):
        return None


def _to_pct(v) -> float | None:
    """Convert a 0-1 decimal to a percentage float (e.g. 0.607 -> 60.7)."""
    f = _to_float(v)
    if f is None:
        return None
    # Values are stored as decimals (0.607 means 60.7%)
    return round(f * 100, 1)


def _month_matches(cell_val, target_month: str) -> bool:
    """Check if a cell value (string or datetime) matches the target month."""
    if cell_val is None:
        return False
    if hasattr(cell_val, "strftime"):           # datetime object
        return cell_val.strftime("%B").lower() == target_month.lower()
    s = str(cell_val).strip().lower()
    return s == target_month.lower() or s == target_month[:3].lower()


def process_utilization(wb, target_month: str = None) -> list:
    """
    Return per-person utilization data for the given month.
    """
    if target_month is None:
        target_month = date.today().strftime("%B")

    sheet_name = _find_util_sheet(wb)
    if sheet_name is None:
        return []

    ws      = wb[sheet_name]
    results = []

    in_section = False

    for row in ws.iter_rows(values_only=True):
        # Pad row to ensure enough columns
        row = list(row) + [None] * 15

        # ---- Section start: col B (index 1) matches target month ----
        if not in_section:
            if _month_matches(row[1], target_month):
                in_section = True
            continue

        # ---- Skip the "# of Days" header row ----
        if row[1] is not None and str(row[1]).strip().lower() == "# of days":
            continue

        # ---- End of section: "Total" in col C ----
        person_val = row[COL_PERSON]
        if person_val is not None and str(person_val).strip().lower() == "total":
            break

        # ---- Start of next month section: a new datetime in col A ----
        if row[0] is not None and hasattr(row[0], "strftime"):
            break

        # ---- Skip blank rows ----
        if not any(row[COL_CHARGEABLE:COL_DIFFERENCE + 1]):
            continue

        role        = str(row[COL_ROLE]).strip()   if row[COL_ROLE]   else ""
        person_name = str(person_val).strip()       if person_val      else ""

        if not person_name or person_name.lower() in ("", "total", "# of days"):
            continue

        util_pct = _to_pct(row[COL_UTILIZATION])
        goal_pct = _to_pct(row[COL_GOAL])
        diff_pct = _to_pct(row[COL_DIFFERENCE])

        if util_pct is None and goal_pct is None:
            continue

        results.append({
            "person":          person_name,
            "role":            role,
            "chargeable":      _to_float(row[COL_CHARGEABLE]),
            "holiday":         _to_float(row[COL_HOLIDAY]),
            "pto":             _to_float(row[COL_PTO]),
            "month_total":     _to_float(row[COL_MONTH_TOTAL]),
            "remaining":       _to_float(row[COL_REMAINING]),
            "utilization_pct": util_pct,
            "goal_pct":        goal_pct,
            "difference_pct":  diff_pct,
            "person_email":    EMAIL_LOOKUP.get(person_name),
            "first_name":      FIRST_NAMES.get(person_name, person_name),
        })

    return results


def build_utilization_emails(
    util_data: list,
    month: str = None,
    sender_name: str = "Jake",
) -> list:
    """
    Build one email per person based on their utilization vs goal.

    Thresholds:
      diff > +10%  → over-utilized, ask if hours need reallocation
      diff < -10%  → under-utilized, ask about non-charge time / unscheduled projects
      otherwise    → informational only
    """
    if not month:
        month = date.today().strftime("%B")

    emails = []
    for u in util_data:
        person      = u.get("person", "")
        first_name  = u.get("first_name", person)
        email_addr  = u.get("person_email")
        util_pct    = u.get("utilization_pct")
        goal_pct    = u.get("goal_pct")
        diff_pct    = u.get("difference_pct")

        if not email_addr or util_pct is None:
            continue

        util_str = f"{util_pct:.1f}%"
        goal_str = f"{goal_pct:.0f}%" if goal_pct is not None else "N/A"
        diff_str = f"{diff_pct:+.1f}%" if diff_pct is not None else "N/A"

        if diff_pct is not None and diff_pct > 10:
            message = (
                f"Your current utilization for {month} is {util_str} against a goal of "
                f"{goal_str} ({diff_str}). You are currently over your utilization goal. "
                f"Do you need any help reallocating or shifting hours to another period?"
            )
        elif diff_pct is not None and diff_pct < -10:
            message = (
                f"Your current utilization for {month} is {util_str} against a goal of "
                f"{goal_str} ({diff_str}). You are currently below your utilization goal. "
                f"Do you have any non-chargeable time planned, or are there any unscheduled "
                f"projects that should be added to the schedule?"
            )
        else:
            message = (
                f"Your current utilization for {month} is {util_str} against a goal of "
                f"{goal_str} ({diff_str}). You are currently on track — no action needed."
            )

        body = (
            f"Hi {first_name},\n\n"
            f"{message}\n\n"
            f"Best,\n{sender_name}"
        )

        emails.append({
            "person":  person,
            "to":      email_addr,
            "subject": f"Utilization Update — {month}",
            "body":    body,
        })

    return emails


def get_pto_schedule(wb, target_months: list) -> dict:
    """
    Extract scheduled PTO hours for each person across the given months.

    Returns:
        {
            "Wojtowicz": {"May": 8, "June": 72, "July": 0},
            "Brooks":    {"May": 56, "June": 0,  "July": 0},
            ...
        }
    Only includes people who have at least one non-zero PTO entry.
    """
    sheet_name = _find_util_sheet(wb)
    if not sheet_name:
        return {}

    ws   = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))

    # Normalise target months to title-case for matching
    # (handles "May", "may", "JUNE", etc.)
    target_set = {m.strip().title() for m in target_months}

    # Map: {person: {month: pto}}
    result: dict = {}

    current_month = None
    in_section    = False

    for row in rows:
        row = list(row) + [None] * 15

        col_a = row[0]
        col_b = row[1]
        col_c = row[2]
        pto   = row[COL_PTO]   # index 6 = column G

        # ── Detect new month section ─────────────────────────
        if col_a is not None and hasattr(col_a, "strftime"):
            # Month header row — col_b is the month name (e.g. "May", "June")
            month_label = str(col_b).strip().title() if col_b else col_a.strftime("%B")
            current_month = month_label
            in_section    = current_month in target_set
            continue

        if not in_section or current_month is None:
            continue

        # ── Skip header / total rows ─────────────────────────
        if col_b is not None and str(col_b).strip().lower() == "# of days":
            continue
        if col_c is not None and str(col_c).strip().lower() == "total":
            in_section = False
            continue

        # ── Data row ─────────────────────────────────────────
        person = str(col_c).strip() if col_c else ""
        if not person or person.lower() in ("", "total", "# of days"):
            continue

        pto_val = _to_float(pto) or 0.0

        result.setdefault(person, {})
        result[person][current_month] = pto_val

    return result
