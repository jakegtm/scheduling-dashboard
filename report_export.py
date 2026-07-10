from __future__ import annotations
# ============================================================
# report_export.py — per-person Excel workbook builder + ZIP packager
# Replaces email_utils.py now that reports are exported, not emailed.
# ============================================================

import io
import zipfile
from datetime import date, timedelta

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

from config import _rank

_BRAND       = "0E2841"
_HEADER_FILL = PatternFill(start_color=_BRAND, end_color=_BRAND, fill_type="solid")
_ZEBRA_FILL  = PatternFill(start_color="F5F7FA", end_color="F5F7FA", fill_type="solid")
_RESP_FILL   = PatternFill(start_color="FFFBE6", end_color="FFFBE6", fill_type="solid")

_HEADER_FONT = Font(name="Arial", color="FFFFFF", bold=True, size=11)
_BODY_FONT   = Font(name="Arial", size=10.5)
_NEG_FONT    = Font(name="Arial", color="C0392B", bold=True, size=10.5)  # over-budget / worked-less-than-scheduled
_OVER_FONT   = Font(name="Arial", color="E67E22", bold=True, size=10.5)  # worked-more-than-scheduled
_POS_FONT    = Font(name="Arial", color="1A6B2F", size=10.5)             # on-track / positive

_THIN   = Side(style="thin", color="D9D9D9")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

_LEFT_TOP    = Alignment(horizontal="left",   vertical="top",    wrap_text=True)
_LEFT_MID    = Alignment(horizontal="left",   vertical="center", wrap_text=False)
_CENTER_MID  = Alignment(horizontal="center", vertical="center", wrap_text=False)

# Columns that hold long free-text and should wrap instead of stretching forever
_WRAP_HEADERS = {"To be reviewed", "Notes", "Response"}
# Columns that read better centered (short numbers/labels)
_CENTER_HEADERS = {
    "Actual Hrs", "Scheduled Hrs", "Difference", "Chargeable Hrs",
    "Remaining Hrs", "Utilization", "Goal", "PTO Hours", "Status", "Budget",
}
_NUMERIC_HEADERS = {"Actual Hrs", "Scheduled Hrs", "Difference"}

_MIN_WIDTH      = 10
_MAX_WIDTH      = 32
_MAX_WRAP_WIDTH = 48
_CHARS_PER_LINE = 46  # used to estimate wrapped row height


def _next_monday() -> str:
    """Return the coming Monday as a readable string, e.g. 'Monday, June 1'."""
    today      = date.today()
    days_ahead = (7 - today.weekday()) % 7
    monday     = today + timedelta(days=days_ahead)
    day = monday.strftime("%B %d").replace(" 0", " ")
    return f"Monday, {day}"


def _autofit_widths(all_headers: list, display_rows: list) -> list:
    """Compute a sensible column width per header based on actual content,
    capped so a single long note doesn't blow out the whole sheet."""
    widths = []
    for col_idx, header in enumerate(all_headers):
        max_len = len(str(header))
        for row in display_rows:
            val = row[col_idx] if col_idx < len(row) else ""
            text = "" if val is None else str(val)
            max_len = max(max_len, len(text))
        cap = _MAX_WRAP_WIDTH if header in _WRAP_HEADERS else _MAX_WIDTH
        widths.append(min(cap, max(_MIN_WIDTH, max_len + 2)))
    return widths


def _write_sheet(wb, title, headers, rows, response_col=True):
    """Create a sheet, write a styled + autofit + zebra-striped table.
    Returns the worksheet so callers can layer extra per-cell styling
    (e.g. coloring a specific "Difference" cell red/orange) on top."""
    ws = wb.create_sheet(title=title[:31])  # Excel sheet-name length limit
    ws.sheet_properties.tabColor = _BRAND
    all_headers = list(headers) + (["Response"] if response_col else [])
    n_cols = len(all_headers)

    # ── header row ──
    ws.append(all_headers)
    ws.row_dimensions[1].height = 20
    for col_idx in range(1, n_cols + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill      = _HEADER_FILL
        cell.font      = _HEADER_FONT
        cell.border    = _BORDER
        cell.alignment = _LEFT_MID

    # ── data rows ──
    display_rows = []
    for row in rows:
        row_values = list(row)
        if response_col:
            row_values.append("")  # blank cell for the person to fill in
        display_rows.append(row_values)
        ws.append(row_values)

    widths = _autofit_widths(all_headers, display_rows)

    for r_offset, row_values in enumerate(display_rows):
        row_idx  = r_offset + 2
        is_even  = (r_offset % 2 == 1)
        max_lines = 1
        for col_idx in range(1, n_cols + 1):
            header = all_headers[col_idx - 1]
            cell   = ws.cell(row=row_idx, column=col_idx)
            cell.font   = _BODY_FONT
            cell.border = _BORDER

            if header in _WRAP_HEADERS:
                cell.alignment = _LEFT_TOP
                text = "" if cell.value is None else str(cell.value)
                if text:
                    max_lines = max(max_lines, -(-len(text) // _CHARS_PER_LINE))
            elif header in _CENTER_HEADERS:
                cell.alignment = _CENTER_MID
            else:
                cell.alignment = _LEFT_MID

            if header in _NUMERIC_HEADERS and isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0.0;-#,##0.0;0"

            if header == "Response":
                cell.fill = _RESP_FILL
            elif is_even:
                cell.fill = _ZEBRA_FILL

        ws.row_dimensions[row_idx].height = max(15, max_lines * 15)

    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A2"

    # Print-friendly: landscape + fit-to-width so a wide table doesn't spill
    # across multiple pages if someone prints or exports this to PDF.
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 0

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
    cover.sheet_properties.tabColor = _BRAND
    cover["A1"] = f"Scheduling Review — {first_name}"
    cover["A1"].font = Font(name="Arial", bold=True, size=16, color=_BRAND)
    cover.row_dimensions[1].height = 26

    deadline = _next_monday()
    cover["A3"] = f"Please review the applicable tabs and reply/update by {deadline} at 12:00 PM."
    cover["A3"].font = Font(name="Arial", size=11, italic=True, color="444444")

    next_row = 5
    _note_font = Font(name="Arial", size=10.5, color="555555")
    if selected_months and len(selected_months) > 1:
        cover.cell(row=next_row, column=1,
                   value=f"Note: this report covers multiple periods: {', '.join(selected_months)}"
                   ).font = _note_font
        next_row += 2
    if no_openair_note:
        cover.cell(row=next_row, column=1,
                   value="Note: No OpenAir report was uploaded, so actual hours are shown as 0 "
                         "on the Current Month Hours tab. Scheduled hours reflect what is planned."
                   ).font = _note_font
    cover.column_dimensions["A"].width = 95
    cover.page_setup.orientation = "landscape"
    cover.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    cover.page_setup.fitToWidth  = 1
    cover.page_setup.fitToHeight = 0

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
            _write_sheet(wb, "Project Tracker", ["Project Code", "To be reviewed"], rows)
            any_section = True

    # ── Budget to Actual ─────────────────────────────────────
    if budget_issues:
        rows, row_types = [], []
        for i in budget_issues:
            rows.append([i.get("project_code", ""), i.get("description", "")])
            row_types.append(i.get("type"))
        ws = _write_sheet(wb, "Budget to Actual", ["Project Code", "To be reviewed"], rows)
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
        _write_sheet(wb, "TBD Projects", ["Project Code", "Status", "Budget", "Notes"], rows)
        any_section = True

    # ── Current Month Hours (all hours, flagged + matches) ──────
    if variance_issues:
        _sorted_var = sorted(
            variance_issues,
            key=lambda v: (
                0 if v.get("person", "") == owner else 1,
                _rank(v.get("person", "")),
                v.get("project_code", ""),
            )
        )
        rows, diffs, is_variance_flags = [], [], []
        for v in _sorted_var:
            diff = v.get("difference", 0)
            is_var = v.get("is_variance", True)  # default True = old callers/back-compat
            diffs.append(diff)
            is_variance_flags.append(is_var)
            review_text = v.get("question", "") if is_var else "✓ Hours match — no action needed"
            if is_staff:
                rows.append([
                    v.get("project_code", ""), v.get("period", ""),
                    v.get("actual_hours", ""), v.get("sched_hours", ""),
                    diff, review_text,
                ])
            else:
                rows.append([
                    v.get("person", ""), v.get("project_code", ""), v.get("period", ""),
                    v.get("actual_hours", ""), v.get("sched_hours", ""),
                    diff, review_text,
                ])
        headers = (
            ["Project Code", "Period", "Actual Hrs", "Scheduled Hrs", "Difference", "To be reviewed"]
            if is_staff else
            ["Person", "Project Code", "Period", "Actual Hrs", "Scheduled Hrs", "Difference", "To be reviewed"]
        )
        diff_col     = 5 if is_staff else 6
        response_col = len(headers) + 1  # _write_sheet appends "Response" as the last column
        ws = _write_sheet(wb, "Current Month Hours", headers, rows)

        _optional_font = Font(name="Arial", size=10, italic=True, color="999999")
        for r_offset, (diff, is_var) in enumerate(zip(diffs, is_variance_flags)):
            row_idx = r_offset + 2
            ws.cell(row=row_idx, column=diff_col).font = (
                (_OVER_FONT if diff < 0 else _NEG_FONT) if is_var else _POS_FONT
            )
            if not is_var:
                # Non-flagged row: make the Response cell read as optional —
                # pre-filled hint text, no yellow tint, matching the row's
                # normal zebra/white background instead.
                resp_cell = ws.cell(row=row_idx, column=response_col)
                resp_cell.value = "Optional"
                resp_cell.font  = _optional_font
                resp_cell.fill  = _ZEBRA_FILL if (r_offset % 2 == 1) else PatternFill(fill_type=None)
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
                               "Remaining Hrs", "To be reviewed"], rows)
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
                            response_col=False)
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
