from __future__ import annotations
# ============================================================
# processors/month_tab.py
# ============================================================
# Actual May tab structure (verified against TAS_Schedule file):
#
#   Row 1: Position titles (Managing Director, Director, etc.) -- span pairs
#   Row 2: Person last names at ODD columns starting from G(7)
#          e.g. G=Sorrentino, I=Colonna, K=Browne, M=Jean...
#          Each person has 2 columns (col N = May 1-15, col N+1 = May 16-31)
#   Row 3: Total hours sums per person
#   Row 4: Holiday hours
#   Row 5: PTO hours
#   Row 6: Period labels "May 1 - 15" / "May 16-31" across all person columns
#   Row 7: Column headers (Client, Project Code, etc.) + period totals
#   Row 8+: Data rows
#
#   A(1): Client | B(2): Project Code | C(3): Project Type
#   D(4): Client Owner | E(5): Project Owner
#   F(6): Total hours for this row
#   G(7) onwards: person hour cells (pairs per person)
#
# "Done" indicator: theme color index 8 (accent5 = #A02B93, purple)
#   in this workbook's theme. Regular scheduled (unconfirmed) hours
#   have an FFFFFFCC (pale yellow) fill or no fill.
# ============================================================

import re
from collections import defaultdict
from datetime import date, datetime
from config import EMAIL_LOOKUP

PERSON_NAME_ROW  = 2   # person last names
PERIOD_LABEL_ROW = 6   # "May 1 - 15" / "May 16-31"
DATA_HEADER_ROW  = 7   # Client, Project Code, etc.
DATA_START_ROW   = 8   # first actual data row
PERSON_COL_START = 7   # column G (1-based)
PERSON_COL_END   = 32  # column AF (1-based)

# Persons appear at odd columns within the range; each occupies 2 columns
# G(7)=person1_p1, H(8)=person1_p2, I(9)=person2_p1, J(10)=person2_p2, ...

_PERIOD_RE = re.compile(r"([A-Za-z]+)\s+(\d+)\s*[-\u2013]\s*(\d+)", re.IGNORECASE)
_MONTH_MAP = {m[:3].lower(): i + 1 for i, m in enumerate([
    "january","february","march","april","may","june",
    "july","august","september","october","november","december"
])}
_MONTH_MAP.update({m.lower(): v for m, v in zip([
    "january","february","march","april","may","june",
    "july","august","september","october","november","december"
], range(1, 13))})


def _parse_deadline(label: str, year: int):
    if not label:
        return None
    m = _PERIOD_RE.search(str(label))
    if not m:
        return None
    month_str = m.group(1).lower()
    end_day   = int(m.group(3))
    month_num = _MONTH_MAP.get(month_str[:3]) or _MONTH_MAP.get(month_str)
    if not month_num:
        return None
    try:
        return date(year, month_num, end_day)
    except ValueError:
        return None


def _is_done(cell) -> bool:
    """Return True if cell has the 'done' purple fill (theme index 8 = accent5)."""
    fill = cell.fill
    if not fill or not fill.fgColor:
        return False
    color = fill.fgColor
    # Primary check: theme color 8 (accent5 = #A02B93 in this workbook's theme)
    if color.type == "theme" and color.theme == 8:
        return True
    # Fallback: check known purple hex codes (with or without alpha prefix)
    if color.type == "rgb":
        hex6 = color.rgb[-6:].upper()
        PURPLE_HEXES = {"A02B93","7030A0","8064A2","9B59B6","B1A0C7","702FA0","800080"}
        if hex6 in PURPLE_HEXES:
            # Also heuristic check
            try:
                r, g, b = int(hex6[0:2],16), int(hex6[2:4],16), int(hex6[4:6],16)
                if r > 80 and b > 80 and g < (r + b) // 3:
                    return True
            except ValueError:
                pass
    return False


def _lookup_email(name: str):
    if not name:
        return None
    name = str(name).strip()
    if name in EMAIL_LOOKUP:
        return EMAIL_LOOKUP[name]
    for key, email in EMAIL_LOOKUP.items():
        if key.lower() == name.lower():
            return email
    # Partial match (handles "S. O'Donnell" vs "O'Donnell")
    for key, email in EMAIL_LOOKUP.items():
        if key.lower() in name.lower() or name.lower() in key.lower():
            return email
    return None


def _find_month_sheet(wb):
    today  = date.today()
    abbr   = today.strftime("%b").lower()
    full   = today.strftime("%B").lower()
    for name in wb.sheetnames:
        low = name.lower().strip()
        if low.startswith(abbr) or low.startswith(full):
            return wb[name]
    return None


def _build_column_map(ws, year: int) -> dict:
    """
    Build a map of col_index -> {person, period, deadline} for
    all person columns (G through AF).

    Person names are in row 2 at the first column of each pair.
    Period labels are in row 6 for every column.
    """
    col_map = {}

    for col in range(PERSON_COL_START, PERSON_COL_END + 1):
        period_val  = ws.cell(row=PERIOD_LABEL_ROW, column=col).value
        period_str  = str(period_val).strip() if period_val else ""

        # Person name is in row 2 at the odd (first) column of each pair
        # Pairs: (7,8), (9,10), (11,12), ...
        # First column of pair = PERSON_COL_START + 2*n for n=0,1,2,...
        pair_start  = PERSON_COL_START + ((col - PERSON_COL_START) // 2) * 2
        person_val  = ws.cell(row=PERSON_NAME_ROW, column=pair_start).value
        person_str  = str(person_val).strip() if person_val else ""

        deadline    = _parse_deadline(period_str, year)

        col_map[col] = {
            "person":   person_str,
            "period":   period_str,
            "deadline": deadline,
        }

    return col_map


def process_month_tab(wb, deadline_warning_days: int = 2, sheet_name: str = None):
    """
    Find current month tab and return unconfirmed hours with an approaching deadline.

    Returns:
        issues     -- list of issue dicts
        sheet_name -- name of sheet processed (or None)
    """
    today = date.today()
    ws    = wb[sheet_name] if sheet_name else _find_month_sheet(wb)
    if ws is None:
        return [], None

    col_map = _build_column_map(ws, today.year)

    # Find columns whose deadline is within the warning window
    upcoming = {
        col: info
        for col, info in col_map.items()
        if info["deadline"] is not None
        and info["deadline"] >= today
        and (info["deadline"] - today).days <= deadline_warning_days
    }

    if not upcoming:
        return [], ws.title

    issues = []
    consecutive_blank = 0

    for row_num in range(DATA_START_ROW, 5000):  # hard cap - file won't exceed this
        client       = ws.cell(row=row_num, column=1).value
        project_code = ws.cell(row=row_num, column=2).value

        # Stop after 10 consecutive blank rows — end of data
        if not client and not project_code:
            consecutive_blank += 1
            if consecutive_blank >= 10:
                break
            continue
        consecutive_blank = 0

        # Skip header-like rows
        if str(client).strip().lower() == "client":
            continue

        project_owner = ws.cell(row=row_num, column=5).value

        for col, info in upcoming.items():
            cell      = ws.cell(row=row_num, column=col)
            hours_val = cell.value

            # Skip blank or zero
            if hours_val is None:
                continue
            try:
                if float(hours_val) == 0:
                    continue
            except (ValueError, TypeError):
                pass

            # Skip confirmed (purple / theme=8) cells
            if _is_done(cell):
                continue

            person       = info["person"]
            days_left    = (info["deadline"] - today).days

            issues.append({
                "client":        client,
                "project_code":  project_code,
                "project_owner": str(project_owner).strip() if project_owner else "",
                "person":        person,
                "person_email":  _lookup_email(person),
                "period":        info["period"],
                "deadline":      info["deadline"],
                "days_left":     days_left,
                "hours":         hours_val,
            })

    return issues, ws.title


def build_month_emails(issues: list, cc_email: str) -> list:
    grouped = defaultdict(list)
    for issue in issues:
        grouped[issue["person"]].append(issue)

    emails = []
    for person, person_issues in grouped.items():
        person_email = person_issues[0].get("person_email")
        if not person_email:
            continue

        deadline  = person_issues[0]["deadline"]
        period    = person_issues[0]["period"]
        days_left = person_issues[0]["days_left"]

        urgency = "today" if days_left == 0 else (
            "tomorrow" if days_left == 1 else f"in {days_left} days ({deadline.strftime('%B %d')})"
        )

        lines = [
            f"Hi {person},\n\n"
            f"Just a reminder -- the deadline for confirming your hours for {period} is {urgency}.\n\n"
            "The following projects still have unconfirmed hours assigned to you:\n"
        ]
        for i in person_issues:
            lines.append(f"    * {i['client']} | {i['project_code']} | {i['hours']} hr(s)")

        lines.append(
            "\n\nPlease confirm, move, or delete these hours so Laren can "
            "finalize the schedule before the deadline.\n\nThanks,\nLaren"
        )

        emails.append({
            "to":      person_email,
            "subject": f"Action Required: Confirm Hours by {deadline.strftime('%B %d')}",
            "body":    "\n".join(lines),
            "person":  person,
            "issues":  person_issues,
        })

    return emails
