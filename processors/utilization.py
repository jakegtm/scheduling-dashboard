from __future__ import annotations
# ============================================================
# processors/utilization.py
# ============================================================
# Reads the "Utilization by Month" tab, finds the current
# month's section, and returns per-person utilization data.
#
# Expected tab structure (each month section):
#   Row N:   Date marker (1-May-26) in col A + month name header
#   Row N+1: Header row — col B "# of Days", col C = # of days,
#             then Chargeable | Holiday | PTO | Month Total |
#             Remaining | Utilization | Goal | Difference
#   Row N+2+: One data row per person
#             col B: Role (MD, DR, SM, MR, SR, AN)
#             col C: Last Name
#             col D: Chargeable hours
#             col E: Holiday hours
#             col F: PTO hours
#             col G: Month Total
#             col H: Remaining
#             col I: Utilization %
#             col J: Goal %
#             col K: Difference %
#   Last row: "Total" row — signals end of section
# ============================================================

from datetime import date
from config import EMAIL_LOOKUP, FIRST_NAMES

UTIL_SHEET_KEYWORDS = ["utilization by month", "util by month", "utilization"]


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
    """Convert a cell value to a percentage (0–100 scale)."""
    f = _to_float(v)
    if f is None:
        return None
    # Excel may store as decimal (e.g. 0.0476) or already as % (e.g. 4.76)
    if -1.0 <= f <= 1.0:
        return round(f * 100, 2)
    return round(f, 2)


def process_utilization(wb, target_month: str = None) -> list:
    """
    Read the Utilization by Month tab and return per-person data
    for the given month (defaults to current calendar month).

    Returns list of dicts:
        person, role, chargeable, holiday, pto, month_total, remaining,
        utilization_pct, goal_pct, difference_pct, person_email, first_name
    """
    if target_month is None:
        target_month = date.today().strftime("%B")   # e.g. "May"

    sheet_name = _find_util_sheet(wb)
    if sheet_name is None:
        return []

    ws = wb[sheet_name]
    results = []

    in_section    = False
    header_found  = False
    blank_count   = 0

    # Column indices (0-based) — determined from header row
    col_role = col_person = None
    col_chargeable = col_holiday = col_pto = col_month_total = None
    col_remaining = col_utilization = col_goal = col_difference = None

    for row in ws.iter_rows(values_only=True):
        # ---- Look for section start ----
        if not in_section:
            # Any cell in first 6 cols mentions the target month?
            # openpyxl returns Excel dates as Python datetime objects,
            # so we must use strftime — not str() — to get the month name.
            row_text = ""
            for c in row[:6]:
                if c is None:
                    continue
                if hasattr(c, "strftime"):          # datetime / date object
                    row_text += " " + c.strftime("%B %b").lower()
                else:
                    row_text += " " + str(c).lower()
            if target_month.lower() in row_text:
                in_section   = True
                header_found  = False
                blank_count   = 0
            continue

        # ---- Find header row inside section ----
        if not header_found:
            row_lower = [str(c).strip().lower() if c is not None else "" for c in row]
            if "chargeable" in row_lower:
                col_chargeable  = row_lower.index("chargeable")
                col_role        = max(col_chargeable - 2, 0)   # B, 2 cols before
                col_person      = col_chargeable - 1            # C
                # Remaining headers: scan for keywords
                for i, v in enumerate(row_lower):
                    if v == "holiday":                            col_holiday     = i
                    elif v == "pto":                              col_pto         = i
                    elif "month total" in v or v == "total":      col_month_total = i
                    elif v == "remaining":                        col_remaining   = i
                    elif "utilization" in v:                      col_utilization = i
                    elif v == "goal":                             col_goal        = i
                    elif "difference" in v or "diff" in v:       col_difference  = i
                header_found = True
            continue

        # ---- Data rows ----
        row_vals = list(row)

        # Detect end of section: "Total" in first few cells
        first_vals = [str(v).strip().lower() for v in row_vals[:5] if v is not None]
        if "total" in first_vals:
            break

        # Detect next month section: a date-like value in col A containing a month name
        if row_vals[0] is not None:
            cell_str = str(row_vals[0]).strip().lower()
            if any(m in cell_str for m in [
                "jan", "feb", "mar", "apr", "may", "jun",
                "jul", "aug", "sep", "oct", "nov", "dec"
            ]) and cell_str != target_month.lower():
                break

        # Skip blank rows
        if not any(v for v in row_vals[:8]):
            blank_count += 1
            if blank_count >= 3:
                break
            continue
        blank_count = 0

        # Parse person
        def _get(idx):
            if idx is None or idx >= len(row_vals):
                return None
            return row_vals[idx]

        role        = str(_get(col_role)).strip()   if _get(col_role)   else ""
        person_name = str(_get(col_person)).strip() if _get(col_person) else ""

        if not person_name or person_name.lower() in ("total", "# of days", ""):
            continue

        util_pct = _to_pct(_get(col_utilization))
        goal_pct = _to_pct(_get(col_goal))
        diff_pct = _to_pct(_get(col_difference))

        # Skip rows with no meaningful data
        if util_pct is None and goal_pct is None:
            continue

        person_email = EMAIL_LOOKUP.get(person_name)
        first_name   = FIRST_NAMES.get(person_name, person_name)

        results.append({
            "person":          person_name,
            "role":            role,
            "chargeable":      _to_float(_get(col_chargeable)),
            "holiday":         _to_float(_get(col_holiday)),
            "pto":             _to_float(_get(col_pto)),
            "month_total":     _to_float(_get(col_month_total)),
            "remaining":       _to_float(_get(col_remaining)),
            "utilization_pct": util_pct,
            "goal_pct":        goal_pct,
            "difference_pct":  diff_pct,
            "person_email":    person_email,
            "first_name":      first_name,
        })

    return results


def build_utilization_emails(util_data: list, month: str = None,
                             sender_name: str = "Jake") -> list:
    """
    Build one email per person based on their utilization difference.

    Thresholds:
        Within ±10%  → informational only, no question
        Below -10%   → ask about non-charge time plan + unscheduled projects
        Above +10%   → ask about reallocation / shifting hours
    """
    if month is None:
        month = date.today().strftime("%B")

    emails = []

    for person in util_data:
        email = person.get("person_email")
        if not email:
            continue

        first = person.get("first_name", person["person"])
        diff  = person.get("difference_pct")
        util  = person.get("utilization_pct")
        goal  = person.get("goal_pct")

        def _fmt_hrs(v):
            if v is None: return "-"
            return f"{v:,.1f}" if v != int(v) else f"{int(v)}"

        def _fmt_pct(v):
            if v is None: return "-"
            return f"{v:.1f}%"

        table_lines = [
            f"  {'Chargeable Hours':<22} {_fmt_hrs(person.get('chargeable'))}",
            f"  {'Holiday Hours':<22} {_fmt_hrs(person.get('holiday'))}",
            f"  {'PTO Hours':<22} {_fmt_hrs(person.get('pto'))}",
            f"  {'Month Total':<22} {_fmt_hrs(person.get('month_total'))}",
            f"  {'Remaining Hours':<22} {_fmt_hrs(person.get('remaining'))}",
            f"  {'Utilization':<22} {_fmt_pct(util)}",
            f"  {'Goal':<22} {_fmt_pct(goal)}",
            f"  {'Difference':<22} {_fmt_pct(diff)}",
        ]
        table = "\n".join(table_lines)

        # Follow-up question based on difference
        question = ""
        if diff is not None:
            if diff < -10:
                question = (
                    "\nCould you also provide a brief note on how you plan to use your "
                    "non-charge time this month? Additionally, are there any projects "
                    "you're aware of that aren't currently reflected in the schedule?"
                )
            elif diff > 10:
                question = (
                    "\nIt looks like you're projected above your utilization goal this "
                    "month — do you need any assistance reallocating or shifting hours "
                    "across the team?"
                )

        body = (
            f"Hi {first},\n\n"
            f"Please see below your projected utilization for {month}:\n\n"
            f"{table}\n"
            f"{question}\n\n"
            f"Best,\n"
            f"{sender_name}"
        )

        emails.append({
            "to":      email,
            "subject": f"Utilization Update — {month}",
            "person":  person["person"],
            "body":    body,
            "section": "utilization",
            "diff_pct": diff,
        })

    return emails
