from __future__ import annotations
# ============================================================
# processors/month_tab.py
# ============================================================

import re
from collections import defaultdict
from datetime import date, timedelta
from config import PURPLE_HEX_CODES
from processors.lookup import lookup_email, lookup_first_name

PERSON_NAME_ROW  = 2
PERIOD_LABEL_ROW = 6
DATA_START_ROW   = 8
PERSON_COL_START = 7
PERSON_COL_END   = 32

_PERIOD_RE = re.compile(r'([A-Za-z]+)\s+(\d{1,2})\s*[-\u2013]\s*(\d{1,2})', re.IGNORECASE)
_MONTH_MAP = {m[:3].lower(): i + 1 for i, m in enumerate([
    'January','February','March','April','May','June',
    'July','August','September','October','November','December'])}
_MONTH_MAP.update({m.lower(): v for m, v in zip([
    'january','february','march','april','may','june',
    'july','august','september','october','november','december'],
    range(1, 13))})


def _parse_deadline(label: str, year: int):
    if not label:
        return None
    m = _PERIOD_RE.search(str(label))
    if not m:
        return None
    month_num = _MONTH_MAP.get(m.group(1).lower()[:3])
    if not month_num:
        return None
    try:
        return date(year, month_num, int(m.group(3)))
    except ValueError:
        return None


def _is_done(cell) -> bool:
    fill = cell.fill
    if not fill or not fill.fgColor:
        return False
    color = fill.fgColor
    if color.type == "theme" and color.theme == 8:
        return True
    if color.type == "rgb":
        hex6 = color.rgb[-6:].upper()
        if hex6 in PURPLE_HEX_CODES:
            try:
                r, g, b = int(hex6[0:2],16), int(hex6[2:4],16), int(hex6[4:6],16)
                return r > 80 and b > 80 and g < (r + b) // 3
            except ValueError:
                pass
    return False


def _find_month_sheet(wb):
    today = date.today()
    abbr  = today.strftime("%b").lower()
    full  = today.strftime("%B").lower()
    for name in wb.sheetnames:
        low = name.lower().strip()
        if low.startswith(abbr) or low.startswith(full):
            return wb[name]
    return None


def _build_column_map(ws, year: int) -> dict:
    col_map = {}
    for col in range(PERSON_COL_START, PERSON_COL_END + 1):
        period_val = ws.cell(row=PERIOD_LABEL_ROW, column=col).value
        period_str = str(period_val).strip() if period_val else ""
        pair_start = PERSON_COL_START + ((col - PERSON_COL_START) // 2) * 2
        person_val = ws.cell(row=PERSON_NAME_ROW, column=pair_start).value
        person_str = str(person_val).strip() if person_val else ""
        deadline   = _parse_deadline(period_str, year)
        col_map[col] = {"person": person_str, "period": period_str, "deadline": deadline}
    return col_map


def process_month_tab(wb, deadline_warning_days: int = 2, sheet_name: str = None):
    today = date.today()
    ws    = wb[sheet_name] if sheet_name else _find_month_sheet(wb)
    if ws is None:
        return [], None

    col_map  = _build_column_map(ws, today.year)
    upcoming = {
        col: info for col, info in col_map.items()
        if info["deadline"] is not None
        and info["deadline"] >= today
        and (info["deadline"] - today).days <= deadline_warning_days
    }

    if not upcoming:
        return [], ws.title

    issues = []
    consecutive_blank = 0

    for row_num in range(DATA_START_ROW, 5000):
        client       = ws.cell(row=row_num, column=1).value
        project_code = ws.cell(row=row_num, column=2).value
        if not client and not project_code:
            consecutive_blank += 1
            if consecutive_blank >= 10:
                break
            continue
        consecutive_blank = 0
        if str(client).strip().lower() == "client":
            continue

        project_owner = ws.cell(row=row_num, column=5).value

        for col, info in upcoming.items():
            cell      = ws.cell(row=row_num, column=col)
            hours_val = cell.value
            if hours_val is None:
                continue
            try:
                if float(hours_val) == 0:
                    continue
            except (ValueError, TypeError):
                pass
            if _is_done(cell):
                continue

            person    = info["person"]
            days_left = (info["deadline"] - today).days

            issues.append({
                "client":        client,
                "project_code":  project_code,
                "project_owner": str(project_owner).strip() if project_owner else "",
                "person":        person,
                "person_email":  lookup_email(person),
                "person_first":  lookup_first_name(person),
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
        emails.append({"to": person_email, "person": person, "issues": person_issues})
    return emails
