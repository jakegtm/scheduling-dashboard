from __future__ import annotations
# ============================================================
# report_export.py — per-person Excel workbook builder + ZIP packager
# Replaces email_utils.py now that reports are exported, not emailed.
# ============================================================

import io
import zipfile
from datetime import date, timedelta

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from config import _rank

_HEADER_FILL = PatternFill(start_color="0E2841", end_color="0E2841", fill_type="solid")
_HEADER_FONT = Font(color="FFFFFF", bold=True)
_NEG_FONT    = Font(color="C0392B", bold=True)   # over-budget / worked-less-than-scheduled
_OVER_FONT   = Font(color="E67E22", bold=True)   # worked-more-than-scheduled
_POS_FONT    = Font(color="1A6B2F")              # on-track / positive


def _next_monday() -> str:
    """Return the coming Monday as a readable string, e.g. 'Monday, June 1'."""
    today      = date.today()
    days_ahead = (7 - today.weekday()) % 7
    monday     = today + timedelta(days=days_ahead)
    day = monday.strftime("%B %d").replace(" 0", " ")
    return f"Monday, {day}"


def _write_sheet(wb, title, headers, rows, response_col=True, col_widths=None):
    """Create a sheet, write a styled header row, then the data rows.
    Returns the worksheet so callers can apply extra per-cell styling."""
    ws = wb.create_sheet(title=title[:31])  # Excel sheet-name length limit
    all_headers = list(headers) + (["Response"] if response_col else [])
    ws.append(all_headers)
    for col_idx in range(1, len(all_headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT

    for row in rows:
        row_values = list(row)
        if response_col:
            row_values.append("")  # blank cell for the person to fill in
        ws.append(row_values)

    widths = col_widths or [18] * len(all_headers)
    for i, w in enumerate(widths[:len(all_headers)], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    if response_col:
        # widen + tint the response column so it's obvious where to type
        resp_col_idx = len(all_headers)
        ws.column_dimensions[get_column_letter(resp_col_idx)].width = max(
            widths[resp_col_idx - 1] if resp_col_idx <= len(widths) else 18, 25
        )
        resp_fill = PatternFill(start_color="FFFBE6", end_color="FFFBE6", fill_type="solid")
        for row_idx in range(2, ws.max_row + 1):
            ws.cell(row=row_idx, column=resp_col_idx).fill = resp_fill

    ws.freeze_panes = "A2"
    return ws


def build_person_workbook(
    owner: str,
    first_name: str,
    tracker_issues: list,
    budget_issues: list,
    tbd_projects: list,
    variance_issues: list,
    util_data: list       = None,
    pto_schedule: dict    = None,
    pto_months: list      = None,
    has_openair: bool     = False,
    no_openair_note: bool = False,
    selected_months: list = None,
    is_staff: bool        = False,
) -> bytes | None:
    """
    Build one Excel workbook (in-memory) for a single person, one sheet per
    applicable report section. Mirrors the old build_html_email's inputs and
    "only include a section if it has content" behavior.

    Returns None if nothing applies for this person (mirrors build_html_email
    returning "" when there's nothing to send) — caller should skip them.
    """
    wb = Workbook()
    wb.remove(wb.active)  # drop the default blank sheet
    any_section = False

    # ── Summary / cover sheet ────────────────────────────────
    cover = wb.create_sheet("Summary")
    cover["A1"] = f"Scheduling Review — {first_name}"
    cover["A1"].font = Font(bold=True, size=14, color="0E2841")
    deadline = _next_monday()
    cover["A3"] = f"Please review the applicable tabs and reply/update by {deadline} at 12:00 PM."
    next_row = 5
    if selected_months and len(selected_months) > 1:
        cover.cell(row=next_row, column=1,
                   value=f"Note: this report covers multiple periods: {', '.join(selected_months)}")
        next_row += 2
    if no_openair_note:
        cover.cell(row=next_row, column=1,
                   value="Note: No OpenAir report was uploaded, so actual hours are shown as 0 "
                         "on the Variance tab. Scheduled hours reflect what is planned.")
    cover.column_dimensions["A"].width = 100

    # ── Project Tracker ──────────────────────────────────────
    if tracker_issues:
        rows = []
        for i in tracker_issues:
            missing = i.get("missing_rates", [])
            for m in missing:
                rows.append([i.get("project_code", ""), f"Missing {m} Rate"])
            if not missing:
                for prob in i.get("problems", []):
                    rows.append([i.get("project_code", ""), prob])
        if rows:
            _write_sheet(wb, "Project Tracker", ["Project Code", "To be reviewed"], rows,
                         col_widths=[20, 55])
            any_section = True

    # ── Budget to Actual ─────────────────────────────────────
    if budget_issues:
        rows, row_types = [], []
        for i in budget_issues:
            rows.append([i.get("project_code", ""), i.get("description", "")])
            row_types.append(i.get("type"))
        ws = _write_sheet(wb, "Budget to Actual", ["Project Code", "To be reviewed"], rows,
                          col_widths=[20, 55])
        for idx, t in enumerate(row_types, start=2):
            ws.cell(row=idx, column=2).font = _NEG_FONT if t == "negative" else _POS_FONT
        any_section = True

    # ── TBD / Pending SOW ─────────────────────────────────────
    owner_tbd = [p for p in (tbd_projects or []) if p.get("owner") == owner]
    if owner_tbd:
        rows = [
            [p.get("project_code", ""), p.get("status", "TBD"),
             f"${p.get('budget', 0):,.0f}" if p.get("budget") else "TBD",
             p.get("notes", "")]
            for p in owner_tbd
        ]
        _write_sheet(wb, "TBD Projects", ["Project Code", "Status", "Budget", "Notes"], rows,
                    col_widths=[20, 15, 15, 45])
        any_section = True

    # ── Variance ──────────────────────────────────────────────
    if variance_issues:
        _sorted_var = sorted(
            variance_issues,
            key=lambda v: (
                0 if v.get("person", "") == owner else 1,
                _rank(v.get("person", "")),
                v.get("project_code", ""),
            )
        )
        rows, diffs = [], []
        for v in _sorted_var:
            diff = v.get("difference", 0)
            diffs.append(diff)
            if is_staff:
                rows.append([
                    v.get("project_code", ""), v.get("period", ""),
                    v.get("actual_hours", ""), v.get("sched_hours", ""),
                    diff, v.get("question", ""),
                ])
            else:
                rows.append([
                    v.get("person", ""), v.get("project_code", ""), v.get("period", ""),
                    v.get("actual_hours", ""), v.get("sched_hours", ""),
                    diff, v.get("question", ""),
                ])
        headers = (
            ["Project Code", "Period", "Actual Hrs", "Scheduled Hrs", "Difference", "To be reviewed"]
            if is_staff else
            ["Person", "Project Code", "Period", "Actual Hrs", "Scheduled Hrs", "Difference", "To be reviewed"]
        )
        widths = ([18, 12, 12, 14, 12, 45] if is_staff
                  else [18, 18, 12, 12, 14, 12, 45])
        diff_col = 5 if is_staff else 6
        ws = _write_sheet(wb, "Variance", headers, rows, col_widths=widths)
        for idx, diff in enumerate(diffs, start=2):
            ws.cell(row=idx, column=diff_col).font = _OVER_FONT if diff < 0 else _NEG_FONT
        any_section = True

    # ── Utilization ───────────────────────────────────────────
    if util_data:
        person_util = [u for u in util_data if u.get("person") == owner]
        if person_util:
            u = person_util[0]
            util_pct = u.get("utilization_pct")
            goal_pct = u.get("goal_pct")
            diff_pct = u.get("difference_pct")

            if diff_pct is not None and diff_pct < -10:
                question = ("What do you plan to do with your non-charge time? "
                            "Are there any projects you know of that aren't in the schedule yet?")
            elif diff_pct is not None and diff_pct > 10:
                question = ("Is there any project work you could use assistance with, "
                            "or places where we can shift hours?")
            else:
                question = ""

            rows = [[
                f"{util_pct:.1f}%" if util_pct is not None else "-",
                f"{goal_pct:.0f}%" if goal_pct is not None else "-",
                f"{diff_pct:+.1f}%" if diff_pct is not None else "-",
                u.get("chargeable", "-"), u.get("remaining", "-"),
                question,
            ]]
            ws = _write_sheet(wb, "Utilization",
                              ["Utilization", "Goal", "Difference", "Chargeable Hrs",
                               "Remaining Hrs", "To be reviewed"], rows,
                              col_widths=[14, 10, 14, 16, 16, 45])
            if diff_pct is not None:
                ws.cell(row=2, column=3).font = (
                    _NEG_FONT if diff_pct > 10 else (_OVER_FONT if diff_pct < -10 else _POS_FONT)
                )
            any_section = True

    # ── PTO Schedule ─────────────────────────────────────────
    if pto_schedule and pto_months:
        person_pto = pto_schedule.get(owner, {})
        has_any_pto = any(person_pto.get(m, 0) for m in pto_months)
        if person_pto and has_any_pto:
            months_to_show = [m for m in pto_months if m in person_pto]
            if months_to_show:
                rows = [[m, f"{int(person_pto[m])} hrs" if person_pto.get(m) else "—"]
                       for m in months_to_show]
                _write_sheet(wb, "PTO", ["Month", "PTO Hours"], rows,
                            response_col=False, col_widths=[20, 15])
                any_section = True

    if not any_section:
        return None

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_reports_zip(people_payload: list[dict]) -> bytes:
    """
    people_payload: list of dicts, each containing build_person_workbook's
    kwargs PLUS a 'filename' key (e.g. 'Jake_Smith.xlsx') used inside the zip.
    Anyone whose workbook comes back empty (no applicable sections) is skipped.
    Returns raw ZIP bytes ready for a Streamlit download_button.
    """
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for person in people_payload:
            person = dict(person)  # don't mutate caller's dict
            filename = person.pop("filename")
            wb_bytes = build_person_workbook(**person)
            if wb_bytes:
                zf.writestr(filename, wb_bytes)
    zip_buf.seek(0)
    return zip_buf.getvalue()
