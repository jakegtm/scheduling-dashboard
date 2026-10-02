from __future__ import annotations
# ============================================================
# processors/write_ups.py
# ============================================================
# Collects non-zero "Write Up / (Down)" amounts from every month tab
# that has that column, so each project owner sees the write ups/downs
# on their projects.
#
# Month tab layout:
#   Row 1:  role labels; the Write Up / (Down) header sits here (AG in Sept–Dec)
#   Row 7:  table headers (Client, Project Code, ..., Project Owner)
#   Row 8+: project rows
#
# The column is found by header text, so tabs without it (Jan–Aug, Jan27)
# are skipped automatically and picked up as soon as the column is added.
# ============================================================

from processors.project_tracker import _lookup_email as lookup_email, _lookup_first as lookup_first_name

HEADER_SEARCH_ROWS = (1, 6, 7)
TABLE_HEADER_ROW   = 7
DATA_START_ROW     = 8

_MONTH_PREFIXES = ("jan", "feb", "mar", "apr", "may", "jun",
                   "jul", "aug", "sep", "oct", "nov", "dec")


def _norm(v) -> str:
    return " ".join(str(v).lower().split()) if isinstance(v, str) else ""


def _is_month_tab(name: str) -> bool:
    low = name.lower().strip()
    return low.startswith(_MONTH_PREFIXES) and "actual" not in low


def _find_writeup_col(ws):
    for r in HEADER_SEARCH_ROWS:
        for cell in next(ws.iter_rows(min_row=r, max_row=r, max_col=80)):
            h = _norm(cell.value)
            if "write up" in h or "write-up" in h:
                return cell.column
    return None


def _find_table_cols(ws) -> dict:
    cols = {"client": 1, "code": 2, "owner": 5}
    for cell in next(ws.iter_rows(min_row=TABLE_HEADER_ROW,
                                  max_row=TABLE_HEADER_ROW, max_col=10)):
        h = _norm(cell.value)
        if h == "client":
            cols["client"] = cell.column
        elif h == "project code":
            cols["code"] = cell.column
        elif h == "project owner":
            cols["owner"] = cell.column
    return cols


def get_write_ups(wb) -> list:
    """
    Returns a list of dicts, in workbook tab order:
        month, client, project_code, owner, owner_email, owner_first, amount
    Only non-zero amounts are returned.
    """
    out = []
    for name in wb.sheetnames:
        if not _is_month_tab(name):
            continue
        ws = wb[name]
        wu_col = _find_writeup_col(ws)
        if not wu_col:
            continue
        cols  = _find_table_cols(ws)
        max_c = max(wu_col, *cols.values())

        blank = 0
        for row in ws.iter_rows(min_row=DATA_START_ROW, max_col=max_c,
                                values_only=True):
            client = row[cols["client"] - 1]
            code   = row[cols["code"] - 1]
            if not client and not code:
                blank += 1
                if blank >= 10:
                    break
                continue
            blank = 0

            val = row[wu_col - 1]
            try:
                amount = float(val) if val is not None else 0.0
            except (ValueError, TypeError):
                continue
            if abs(amount) < 0.005:
                continue

            owner = str(row[cols["owner"] - 1] or "").strip()
            out.append({
                "month":        name.strip(),
                "client":       client,
                "project_code": str(code).strip() if code else "",
                "owner":        owner,
                "owner_email":  lookup_email(owner),
                "owner_first":  lookup_first_name(owner),
                "amount":       amount,
            })
    return out
