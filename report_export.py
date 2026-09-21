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

# Missing-time-entry reminder banner (amber, to read as urgent without
# competing with the red used for flagged variances)
_BANNER_FILL = PatternFill(start_color="FDEBD0", end_color="FDEBD0", fill_type="solid")
_BANNER_FONT = Font(name="Arial", color="9C4500", bold=True, size=11)

_THIN   = Side(style="thin", color="D9D9D9")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

_LEFT_TOP    = Alignment(horizontal="left",   vertical="top",    wrap_text=True)
_LEFT_MID    = Alignment(horizontal="left",   vertical="center", wrap_text=False)
_CENTER_MID  = Alignment(horizontal="center", vertical="center", wrap_text=False)

# Columns that hold long free-text and should wrap instead of stretching forever
_WRAP_HEADERS = {"To be reviewed", "Notes", "Description", "Response"}
# Columns that read better centered (short numbers/labels)
_CENTER_HEADERS = {
    "Actual Hrs", "Scheduled Hrs", "Difference", "Chargeable Hrs",
    "Remaining Hrs", "Utilization", "Goal", "PTO Hours", "Status", "Budget",
    "Budget Amount", "Budget = Actual", "Hrs", "Date",
}
_NUMERIC_HEADERS = {"Actual Hrs", "Scheduled Hrs", "Difference", "Hrs"}

_MIN_WIDTH      = 10
_MAX_WIDTH      = 32
_MAX_WRAP_WIDTH = 48
_CHARS_PER_LINE = 46  # fallback line-width estimate (row height now uses actual column width)


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


def _data_start_row(banner=None) -> int:
    """First data row in a sheet written by _write_sheet (3 with a banner, else 2)."""
    return 3 if banner else 2


def _write_sheet(wb, title, headers, rows, response_col=True, banner=None):
    """Create a sheet, write a styled + autofit + zebra-striped table.
    Returns the worksheet so callers can layer extra per-cell styling
    (e.g. coloring a specific "Difference" cell red/orange) on top.

    `banner` writes an attention row above the header — used for the
    missing-time-entry reminder. When a banner is present the header lands
    on row 2 and data starts on row 3, so callers doing per-row styling
    must offset by _data_start_row(banner)."""
    ws = wb.create_sheet(title=title[:31])  # Excel sheet-name length limit
    ws.sheet_properties.tabColor = _BRAND
    all_headers = list(headers) + (["Response"] if response_col else [])
    n_cols = len(all_headers)

    # ── banner row (optional) ──
    header_row = 1
    if banner:
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
        cell = ws.cell(row=1, column=1, value=banner)
        cell.fill      = _BANNER_FILL
        cell.font      = _BANNER_FONT
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        cell.border    = _BORDER
        ws.row_dimensions[1].height = max(22, -(-len(banner) // max(n_cols * 12, 1)) * 15)
        header_row = 2

    # ── header row ──
    ws.append(all_headers)
    ws.row_dimensions[header_row].height = 20
    for col_idx in range(1, n_cols + 1):
        cell = ws.cell(row=header_row, column=col_idx)
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
        row_idx  = r_offset + header_row + 1
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
                    # Count wrapped lines against this column's actual width,
                    # and respect explicit line breaks — a multi-line note
                    # needs a row per segment, not per total character count.
                    per_line = max(10, int(widths[col_idx - 1]) - 2)
                    lines = sum(
                        max(1, -(-len(seg) // per_line))
                        for seg in text.split("\n")
                    )
                    max_lines = max(max_lines, lines)
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

    ws.freeze_panes = f"A{header_row + 1}"

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
    noncharge_data: list  = None,
    missing_time: dict    = None,
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

    # Missing-time-entry reminder. Shown as a banner on the Current Month Hours
    # tab; if that tab isn't generated for this person, it falls back to the
    # Summary sheet so the reminder still reaches them.
    time_banner = (missing_time or {}).get("message")

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
            budget_val = i.get("budget", 0) or 0
            rows.append([
                i.get("project_code", ""),
                f"${budget_val:,.0f}",
                i.get("description", ""),
                i.get("budget_equals_actual", "No"),
            ])
            row_types.append(i.get("type"))
        ws = _write_sheet(
            wb, "Budget to Actual",
            ["Project Code", "Budget Amount", "To be reviewed", "Budget = Actual"], rows,
        )
        for idx, t in enumerate(row_types, start=2):
            ws.cell(row=idx, column=3).font = _NEG_FONT if t == "negative" else _POS_FONT
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

    if time_banner and not variance_issues:
        cover.cell(row=next_row, column=1, value=time_banner).font = _BANNER_FONT
        cover.cell(row=next_row, column=1).fill = _BANNER_FILL
        next_row += 2
        any_section = True  # a missing-time reminder alone is worth sending

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
            if is_var:
                review_text = v.get("question", "")
            elif diff == 0:
                review_text = "✓ Hours match — no action needed"
            else:
                # Not flagged, but not equal either — don't claim they match.
                review_text = "Within threshold — no action needed"
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
        ws = _write_sheet(wb, "Current Month Hours", headers, rows, banner=time_banner)
        base_row = _data_start_row(time_banner)

        _optional_font = Font(name="Arial", size=10, italic=True, color="999999")
        for r_offset, (diff, is_var) in enumerate(zip(diffs, is_variance_flags)):
            row_idx = r_offset + base_row
            # diff = actual - scheduled, so diff < 0 == worked LESS than scheduled.
            # _NEG_FONT is the worked-less color; _OVER_FONT is worked-more.
            ws.cell(row=row_idx, column=diff_col).font = (
                (_NEG_FONT if diff < 0 else _OVER_FONT) if is_var else _POS_FONT
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

    # ── Non-Charge Time ───────────────────────────────────────
    if noncharge_data:
        rows, flags = [], []
        for e in noncharge_data:
            if e.get("available_time"):
                prompt = ("This was logged as Available Time. Is there project work "
                          "that should be scheduled for you?")
            elif e.get("needs_response"):
                prompt = "No note was logged for this entry. What was this time spent on?"
            else:
                prompt = ""
            rows.append([
                e.get("date_str", ""), e.get("task", ""), e.get("hours", 0),
                e.get("notes", "") or "—", e.get("description", "") or "—", prompt,
            ])
            flags.append(bool(e.get("needs_response")))

        headers = ["Date", "Task", "Hrs", "Notes", "Description", "To be reviewed"]
        response_col = len(headers) + 1  # _write_sheet appends "Response" last
        ws = _write_sheet(wb, "Non-Charge Time", headers, rows)

        _optional_font = Font(name="Arial", size=10, italic=True, color="999999")
        for r_offset, (e, needs) in enumerate(zip(noncharge_data, flags)):
            row_idx = r_offset + 2
            if e.get("available_time"):
                # Available Time is the line that most needs an answer — make the
                # task cell stand out the way a flagged variance does.
                ws.cell(row=row_idx, column=2).font = _NEG_FONT
            if not needs:
                # Entry already has a note: mark the Response cell optional,
                # matching the Current Month Hours tab's treatment.
                resp_cell = ws.cell(row=row_idx, column=response_col)
                resp_cell.value = "Optional"
                resp_cell.font  = _optional_font
                resp_cell.fill  = _ZEBRA_FILL if (r_offset % 2 == 1) else PatternFill(fill_type=None)
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


# Rows holding a block title or banner, per sheet — excluded from autofit.
_TITLE_ROWS: dict = {}


def _autofit(ws, skip_rows=(), min_w=9, max_w=46):
    """Size every column to its widest cell, once the sheet is fully written.

    Title/banner rows are skipped: they are long strings sitting in column A
    and would stretch it across the screen. Formula cells are measured by
    their result's likely width, not the formula text.
    """
    skip = set(skip_rows)
    widths = {}
    for row in ws.iter_rows():
        if row and row[0].row in skip:
            continue
        for cell in row:
            v = cell.value
            if v is None:
                continue
            if isinstance(v, str) and v.startswith("="):
                length = 8              # a summed hours figure, not the formula
            elif isinstance(v, float):
                length = len(f"{v:g}")
            else:
                length = max((len(line) for line in str(v).split("\n")), default=0)
            widths[cell.column] = max(widths.get(cell.column, 0), length)
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = min(
            max(w + 2, min_w), max_w)


def _block(ws, r0, title, headers, rows, widths, total_row_idx=None):
    """Write one titled table starting at row r0. Returns the next free row.

    Unlike _write_sheet this stacks several tables on one sheet, so the
    Summary can carry a combined total plus per-month blocks, and By Week can
    carry a block per week.
    """
    ws.cell(row=r0, column=1, value=title).font = Font(
        name="Arial", bold=True, size=12, color=_BRAND)
    ws.row_dimensions[r0].height = 20
    _TITLE_ROWS.setdefault(ws.title, set()).add(r0)

    hr = r0 + 1
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=hr, column=c, value=h)
        cell.fill, cell.font, cell.border = _HEADER_FILL, _HEADER_FONT, _BORDER
        cell.alignment = _LEFT_MID
        widths[c] = max(widths.get(c, 0), len(str(h)) + 2)
    ws.row_dimensions[hr].height = 20

    for i, row in enumerate(rows):
        r = hr + 1 + i
        is_total = (total_row_idx is not None and i == total_row_idx)
        for c, v in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = _BORDER
            cell.alignment = _CENTER_MID if c > 1 else _LEFT_MID
            if is_total:
                cell.font = Font(name="Arial", bold=True, size=10.5)
            elif i % 2 == 1:
                cell.fill = _ZEBRA_FILL
            if isinstance(v, (int, float)):
                # "#,##0.##" renders 20 as "20." because the dot is literal.
                cell.number_format = "General"
            widths[c] = max(widths.get(c, 0), len(str(v)) + 2)
    return hr + 1 + len(rows) + 2  # blank row between blocks


def _summary_rows(people, entries_for, chargeable_for, display_names):
    """Build the data rows (plus TOTAL) for one summary block."""
    from config import NONCHARGE_ACTIVITY_COLUMNS, NONCHARGE_TIMEOFF_COLUMNS
    from processors.noncharge import task_group

    act, off = NONCHARGE_ACTIVITY_COLUMNS, NONCHARGE_TIMEOFF_COLUMNS
    n_num = len(act) + 1 + len(off)          # activity + chargeable + time off
    rows, grand = [], [0.0] * n_num
    for person in people:
        buckets = {c: 0.0 for c in act + off}
        for e in entries_for(person):
            buckets[task_group(e["task"])] += e["hours"]
        nums = ([round(buckets[c], 2) for c in act]
                + [round(chargeable_for(person), 2)]
                + [round(buckets[c], 2) for c in off])
        for i, v in enumerate(nums):
            grand[i] += v
        # Written as values, not formulas: openpyxl stores no cached result,
        # and Excel was opening these columns blank.
        tnc = round(sum(nums[:len(act)]), 2)
        all_hrs = round(tnc + sum(nums[len(act):]), 2)
        rows.append([display_names.get(person, person)]
                    + nums[:len(act)] + [tnc] + nums[len(act):] + [all_hrs])
    g = [round(v, 2) for v in grand]
    g_tnc = round(sum(g[:len(act)]), 2)
    g_all = round(g_tnc + sum(g[len(act):]), 2)
    rows.append(["TOTAL"] + g[:len(act)] + [g_tnc] + g[len(act):] + [g_all])
    return rows, act, off


def build_consolidated_noncharge(
    noncharge_by_person: dict,
    month_label: str = "",
    periods: list = None,
    display_names: dict = None,
    rank_fn=None,
    chargeable: dict = None,
    weeks: list = None,
    month_blocks: list = None,
    scope_days: set = None,
) -> bytes:
    """
    One workbook covering the whole team's non-charge time.

    Summary   combined total, plus a block per month when more than one is
              selected. Activity columns sum to Total Non-Charge; Chargeable,
              PTO, Holiday and Leave sit beside it and roll into Total ALL.
    By Week   a matrix of Total Non-Charge per person per week, then a full
              block per week. Weeks are Mon-Sun clipped to the selected
              periods and labelled with their date span and day count.
    Detail    every entry with its notes and description.

    A review document, not a questionnaire: no Response column.
    Returns b"" when there is nothing to report.
    """
    if not noncharge_by_person:
        return b""

    display_names = display_names or {}
    rank    = rank_fn or (lambda p: p)
    charge  = chargeable or {}
    people  = sorted(noncharge_by_person, key=rank)
    scope   = month_label or ", ".join(periods or [])

    def _charge_in(person, days=None):
        per_day = charge.get(person, {})
        if days is None:
            return sum(per_day.values())
        return sum(h for d, h in per_day.items() if d in days)

    # Chargeable must be clipped to the reported scope, or the column shows
    # the person's whole-year total next to one month of non-charge time.
    if not scope_days:
        if weeks:
            scope_days = {d for w in weeks for d in w["days"]}
        else:
            scope_days = {e["date"] for v in noncharge_by_person.values() for e in v}

    _TITLE_ROWS.clear()

    wb = Workbook()
    wb.remove(wb.active)

    # ══ Summary ═══════════════════════════════════════════════
    ws = wb.create_sheet("Summary")
    ws.sheet_properties.tabColor = _BRAND
    ws.cell(row=1, column=1, value=f"Non-Charge Time \u2014 {scope}").font = Font(
        name="Arial", bold=True, size=14, color=_BRAND)
    ws.row_dimensions[1].height = 24

    widths = {}
    next_row = 3

    def _write_summary_block(title, subset, days=None):
        nonlocal next_row
        rows, act, off = _summary_rows(
            people,
            lambda p: [e for e in subset.get(p, [])],
            lambda p: _charge_in(p, days),
            display_names,
        )
        headers = (["Person"] + act + ["Total Non-Charge", "Total Chargeable"]
                   + off + ["Total ALL"])
        start = next_row
        next_row = _block(ws, start, title, headers, rows,
                          widths, total_row_idx=len(rows) - 1)

        return start

    _write_summary_block(f"Total \u2014 {scope}", noncharge_by_person, scope_days)

    # Sub-blocks under the combined total, ordered broad to narrow:
    # per-month first (when more than one month is selected), then per-period.
    for label, subset, days in (month_blocks or []):
        _write_summary_block(label, subset, days)

    _autofit(ws, skip_rows={1, 2} | _TITLE_ROWS.get(ws.title, set()))
    ws.freeze_panes = "B4"

    # ══ By Week ═══════════════════════════════════════════════
    if weeks:
        ws2 = wb.create_sheet("By Week")
        ws2.sheet_properties.tabColor = _BRAND
        ws2.cell(row=1, column=1,
                 value=f"Non-Charge Time by Week \u2014 {scope}").font = Font(
            name="Arial", bold=True, size=14, color=_BRAND)
        ws2.cell(row=2, column=1, value=(
            "Weeks run Monday-Sunday, trimmed to the selected period(s); the "
            "day count in each heading shows how many days that week "
            "contributes, so the weekly figures add back to the Summary."
        )).font = Font(name="Arial", italic=True, size=9, color="666666")
        ws2.row_dimensions[1].height = 24

        w2, row2 = {}, 4

        # -- one full block per week --
        for w in weeks:
            dayset = set(w["days"])
            subset = {p: [e for e in noncharge_by_person.get(p, [])
                          if e["date"] in dayset] for p in people}
            rows, act, off = _summary_rows(
                people, lambda p: subset.get(p, []),
                lambda p: _charge_in(p, dayset), display_names)
            headers = (["Person"] + act + ["Total Non-Charge", "Total Chargeable"]
                       + off + ["Total ALL"])
            start = row2
            row2 = _block(ws2, start, f"Week of {w['label']}", headers, rows,
                          w2, total_row_idx=len(rows) - 1)
        _autofit(ws2, skip_rows={1, 2} | _TITLE_ROWS.get(ws2.title, set()))
        ws2.freeze_panes = "A4"

    # ══ Detail ════════════════════════════════════════════════
    detail_rows = []
    for person in people:
        for e in sorted(noncharge_by_person[person],
                        key=lambda x: (not x["available_time"], x["date"])):
            detail_rows.append([
                display_names.get(person, person), e["date_str"], e["period"],
                e["task"], e["hours"], e["notes"] or "\u2014",
                e["description"] or "\u2014",
            ])

    _write_sheet(
        wb, "Detail",
        ["Person", "Date", "Period", "Task", "Hrs", "Notes", "Description"],
        detail_rows, response_col=False)

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
