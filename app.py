from __future__ import annotations
# ============================================================
# app.py — GTM Scheduling Analyzer  |  streamlit run app.py
# ============================================================

import gc
import hashlib
import io
import warnings
import zipfile
from collections import defaultdict
from datetime import datetime

import openpyxl
import streamlit as st


from config import (
    EMAIL_LOOKUP, STAFF_NAMES, NAME_ALIASES,
    POSITION_ORDER, PERSON_ROLE, _rank, DISPLAY_NAMES,
    DEFAULT_BUDGET_THRESHOLD, DEFAULT_NEGATIVE_THRESHOLD,
    DEFAULT_PROJECTION_THRESHOLD_PCT,
    DEFAULT_VARIANCE_MIN, DEFAULT_VARIANCE_MAX,
)
from report_export import (build_person_workbook, build_reports_zip,
                           build_consolidated_noncharge)
from processors.budget_actual   import process_budget_actual
from processors.project_tracker import process_project_tracker
from processors.variance        import (
    parse_openair_report, read_schedule_hours,
    compute_variances, get_available_months, filter_by_months,
    get_schedule_periods,
)
from processors.utilization import get_pto_schedule
from processors.noncharge import (
    parse_noncharge_report, filter_noncharge, noncharge_totals,
    months_from_periods, periods_in_months,
)
from processors.time_entry import (
    parse_time_coverage, previous_week, find_missing_time, format_week,
)

warnings.filterwarnings("ignore", category=UserWarning)

st.set_page_config(page_title="GTM Scheduling Analyzer",
                   layout="wide", page_icon="📊")

st.markdown("""<style>
[data-testid="metric-container"] {
    background:#f8fafc; border:1px solid #e2e8f0;
    border-radius:10px; padding:16px 20px;
    box-shadow:0 1px 3px rgba(0,0,0,.06);
}
[data-testid="metric-container"] [data-testid="stMetricLabel"] {
    font-size:13px; color:#64748b; font-weight:500;
}
[data-testid="metric-container"] [data-testid="stMetricValue"] {
    font-size:28px; font-weight:700; color:#0E2841;
}
.stTabs [data-baseweb="tab-list"] {
    gap:4px; background:#f1f5f9; border-radius:10px; padding:4px;
}
.stTabs [data-baseweb="tab"] {
    border-radius:8px; padding:8px 20px; font-weight:500; color:#64748b;
}
.stTabs [aria-selected="true"] {
    background:white !important; color:#0E2841 !important;
    box-shadow:0 1px 3px rgba(0,0,0,.1);
}
[data-testid="stSidebar"] { background:#f8fafc; border-right:1px solid #e2e8f0; }
[data-testid="stExpander"] { border:1px solid #e2e8f0 !important; border-radius:8px !important; }
</style>""", unsafe_allow_html=True)

import os as _os
_logo_path = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "assets", "logo.png")
if _os.path.exists(_logo_path):
    try:
        st.logo(_logo_path)
    except Exception:
        pass

# Header with logo embedded via HTML to prevent column clipping
_logo_b64 = ""
if _os.path.exists(_logo_path):
    import base64 as _b64
    with open(_logo_path, "rb") as _f:
        _logo_b64 = _b64.b64encode(_f.read()).decode()

if _logo_b64:
    st.markdown(
        f'''<div style="display:flex;align-items:center;gap:16px;margin-bottom:4px;">
        <img src="data:image/png;base64,{_logo_b64}"
             style="height:72px;width:auto;object-fit:contain;flex-shrink:0;">
        <div>
          <div style="font-size:1.8rem;font-weight:700;color:#0E2841;line-height:1.2;">GTM Scheduling Analyzer</div>
          <div style=\"font-size:0.85rem;color:#64748b;\">Today: {datetime.now().strftime('%A, %B %d, %Y')}</div>
        </div></div>''',
        unsafe_allow_html=True)
else:
    st.markdown("## GTM Scheduling Analyzer")
    st.caption(f"Today: {datetime.now().strftime('%A, %B %d, %Y')}")

# ============================================================
# AUTHENTICATION
# Credentials stored in Streamlit Secrets:
#   [auth]
#   username = "gtmtas"
#   password_hash = "77c0b600dc99b2c0b5dc5db009c929f16927148a53cd11b4be986237599f69ee"
# Falls back to hardcoded hash if secrets not configured.
# ============================================================
import hashlib as _hl

_FALLBACK_USER = "gtmtas"
_FALLBACK_HASH = "77c0b600dc99b2c0b5dc5db009c929f16927148a53cd11b4be986237599f69ee"

def _check_credentials(username: str, password: str) -> bool:
    pw_hash = _hl.sha256(password.encode()).hexdigest()
    try:
        stored_user = st.secrets["auth"]["username"]
        stored_hash = st.secrets["auth"]["password_hash"]
    except (KeyError, FileNotFoundError):
        stored_user = _FALLBACK_USER
        stored_hash = _FALLBACK_HASH
    return username.strip() == stored_user and pw_hash == stored_hash

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown("<div style='height:40px'></div>", unsafe_allow_html=True)
    _, login_col, _ = st.columns([1, 1.2, 1])
    with login_col:
        if _logo_b64:
            st.markdown(
                f'<div style="text-align:center;margin-bottom:8px;">'
                f'<img src="data:image/png;base64,{_logo_b64}" style="height:80px;width:auto;"></div>',
                unsafe_allow_html=True)
        st.markdown(
            '<div style="text-align:center;font-size:1.4rem;font-weight:700;'
            'color:#0E2841;margin-bottom:4px;">GTM Scheduling Analyzer</div>',
            unsafe_allow_html=True)
        st.markdown(
            '<div style="text-align:center;color:#64748b;margin-bottom:24px;'
            'font-size:0.9rem;">Sign in to continue</div>',
            unsafe_allow_html=True)
        with st.form("login_form"):
            username = st.text_input("Username", placeholder="Enter username")
            password = st.text_input("Password", type="password", placeholder="Enter password")
            submitted = st.form_submit_button("Sign In", use_container_width=True, type="primary")
            if submitted:
                if _check_credentials(username, password):
                    st.session_state.authenticated = True
                    st.rerun()
                else:
                    st.error("Incorrect username or password.")
    st.stop()

# ── Logout button in sidebar ──────────────────────────────────
# (rendered after authentication check so sidebar only shows when logged in)

# ============================================================
# HELPERS
# ============================================================

def _hash(b: bytes) -> str:
    return hashlib.md5(b).hexdigest()

def _find_sheet(sheetnames, keywords):
    for name in sheetnames:
        if any(kw in name.lower() for kw in keywords):
            return name
    return None

def _lookup_email(name: str):
    if not name:
        return None
    name = name.strip()
    if name in EMAIL_LOOKUP:
        return EMAIL_LOOKUP[name]
    for k, v in EMAIL_LOOKUP.items():
        if k.lower() == name.lower():
            return v
    return None

def _normalize_name(name: str) -> str:
    """Resolve name aliases so bare 'O\'Donnell' → 'J. O\'Donnell'."""
    if not name:
        return name
    return NAME_ALIASES.get(name.strip(), name.strip())

# ============================================================
# SESSION STATE
# ============================================================
_DEFAULTS = {
    "run_triggered":   False,
    "sched_bytes":     None,
    "oa_bytes":        None,
    "selected_owners": set(),
    "settings": {
        "budget_threshold":   float(DEFAULT_BUDGET_THRESHOLD),
        "negative_threshold": float(DEFAULT_NEGATIVE_THRESHOLD),
        "variance_min":       float(DEFAULT_VARIANCE_MIN),
        "variance_max":       float(DEFAULT_VARIANCE_MAX),
    },
}
for _k, _v in _DEFAULTS.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v

# ============================================================
# CACHING
# cache_resource — for complex objects (workbooks, nested dicts)
#                  stored in memory, NOT pickled → no crash
# cache_data     — for simple serialisable returns (lists, tuples of strings)
# ============================================================

@st.cache_resource(max_entries=2, show_spinner=False)
def _load_wb(file_hash: str, _file_bytes: bytes):
    return openpyxl.load_workbook(io.BytesIO(_file_bytes), data_only=True)

@st.cache_resource(max_entries=2, show_spinner=False)
def _parse_openair(oa_hash: str, _oa_bytes: bytes) -> dict:
    """Parse OpenAir CSV and cache the nested dict as a resource (no pickling)."""
    return parse_openair_report(io.BytesIO(_oa_bytes))

@st.cache_data(show_spinner=False)
def run_budget(file_hash, _b, budget_thr, proj_pct, neg_thr, _v="v2"):  # bump to bust cache after logic changes
    wb    = _load_wb(file_hash, _b)
    sheet = _find_sheet(wb.sheetnames, ["budget to actual", "budget"])
    if not sheet:
        return [], None
    result = process_budget_actual(wb[sheet], budget_thr, proj_pct, neg_thr)
    gc.collect()
    return result, sheet

@st.cache_data(show_spinner=False)
def run_tracker(file_hash, _b):
    wb    = _load_wb(file_hash, _b)
    sheet = _find_sheet(wb.sheetnames, ["project tracker", "tracker"])
    if not sheet:
        return [], [], None, {}
    issues, tbd, owner_map = process_project_tracker(wb[sheet])
    gc.collect()
    return issues, tbd, sheet, owner_map

@st.cache_data(show_spinner=False)
def run_time_coverage(oa_hash, _oa_bytes):
    """Per-person daily hours from the time report, plus the report's own
    generated-on date (used to pick which week to check)."""
    import io
    try:
        cov, report_date = parse_time_coverage(io.BytesIO(_oa_bytes))
    except Exception:
        cov, report_date = {}, None
    gc.collect()
    return cov, report_date


@st.cache_data(show_spinner=False)
def run_noncharge(oa_hash, _oa_bytes):
    """Parse non-charge time from the OpenAir report. Keyed on the OpenAir
    file hash — unlike the old utilization data, this comes from the time
    report, not the schedule workbook."""
    import io
    try:
        data = parse_noncharge_report(io.BytesIO(_oa_bytes))
    except Exception:
        data = {}
    gc.collect()
    return data

def _months_from_periods(periods) -> list:
    """['July 1-15', 'August 1-15'] -> ['July', 'August'] (order preserved, deduped).

    Single source of truth for "which month tabs do the selected periods
    imply" — used for reading schedule hours AND for building the roster."""
    import re as _re
    months = []
    for p in periods:
        m = _re.match(r"([A-Za-z]+)", str(p))
        if m:
            name = m.group(1).capitalize()
            if name not in months:
                months.append(name)
    return months


def _find_month_sheet(wb, month: str):
    """First sheet whose name starts with the month's 3-letter prefix.
    Handles tabs named 'Aug'/'August', 'Sept'/'September', 'March', etc."""
    return next(
        (s for s in wb.sheetnames if s.lower().startswith(month[:3].lower())),
        None,
    )


@st.cache_data(show_spinner=False)
def get_valid_people(file_hash, _b, months_tuple):
    """Roster = everyone staffed on ANY month tab implied by the selected
    periods. Previously this read only the current calendar month, which
    silently dropped anyone not staffed in that one month."""
    wb = _load_wb(file_hash, _b)
    people = set()
    for month in months_tuple:
        sheet = _find_month_sheet(wb, month)
        if not sheet:
            continue
        ws = wb[sheet]
        for col in range(7, 45):
            val = ws.cell(row=2, column=col).value
            if val:
                people.add(_normalize_name(str(val).strip()))
    return people

def get_oa_periods(oa_hash, _oa_bytes):
    """Return (available, future) period lists from OpenAir data.
    Not cached — period logic is fast; underlying parse IS cached in _parse_openair."""
    actual = _parse_openair(oa_hash, _oa_bytes)
    return get_available_months(actual)

def get_sched_periods(file_hash, _b, active_month):
    """Return (available, future) period lists from the schedule sheet.
    Not cached — period logic is fast; underlying workbook IS cached in _load_wb."""
    wb = _load_wb(file_hash, _b)
    sheet = next(
        (s for s in wb.sheetnames if s.lower().startswith(active_month[:3].lower())),
        None,
    )
    if not sheet:
        return [], []
    return get_schedule_periods(wb, sheet)

@st.cache_data(show_spinner=False, max_entries=10)
def run_variance(file_hash, _b, oa_hash, _oa,
                 selected_months_tuple, var_min, var_max, active_month,
                 include_all=False,
                 _v="v16"):  # bump _v to bust stale cache after code changes
    wb = _load_wb(file_hash, _b)

    # Determine which month tabs to read based on selected periods.
    # e.g. selecting "May 1-15" and "June 1-15" requires both May and June tabs.
    # Read exactly the tabs the SELECTED periods point at. active_month is a
    # fallback for when nothing parses — not an unconditional addition, which
    # used to drag in the current month's tab on every single run.
    needed_months = _months_from_periods(selected_months_tuple) or [active_month]

    sheets_to_read = []
    for month in needed_months:
        sheet = _find_month_sheet(wb, month)
        if sheet and sheet not in sheets_to_read:
            sheets_to_read.append(sheet)

    if not sheets_to_read:
        return [], f"No sheet found for {active_month}"

    try:
        # Merge schedule data across all relevant month tabs
        sched = {}
        for sheet in sheets_to_read:
            sheet_sched = read_schedule_hours(wb, sheet)
            for person, projects in sheet_sched.items():
                if person not in sched:
                    sched[person] = {}
                for proj, periods in projects.items():
                    if proj not in sched[person]:
                        sched[person][proj] = {}
                    sched[person][proj].update(periods)

        if oa_hash and _oa:
            # Use real OpenAir actuals
            actual   = _parse_openair(oa_hash, _oa)
            filtered = filter_by_months(actual, list(selected_months_tuple))
        else:
            # No OpenAir — build fake actual_data with 0s from the schedule
            # so that any scheduled hours will produce a variance (actual=0).
            actual = {}
            for person, projects in sched.items():
                actual[person] = {}
                for proj, periods in projects.items():
                    actual[person][proj] = {period: 0.0 for period in periods}
            filtered = filter_by_months(actual, list(selected_months_tuple))
        variances = compute_variances(filtered, sched,
                                      min_diff=var_min, max_diff=var_max,
                                      selected_periods=list(selected_months_tuple),
                                      include_all=include_all)
        gc.collect()
        return variances, None
    except Exception as e:
        gc.collect()
        return [], str(e)

# ============================================================
# SIDEBAR — settings (st.form prevents reruns on +/- clicks)
# ============================================================
active_month = datetime.now().strftime("%B")

# Compute current month + next 2 for PTO schedule
def _pto_months(current: str) -> list:
    import calendar
    month_names = list(calendar.month_name)[1:]  # Jan..Dec
    try:
        idx = month_names.index(current)
    except ValueError:
        return [current]
    return [month_names[(idx + i) % 12] for i in range(3)]

_pto_month_list = _pto_months(active_month)

with st.sidebar:
    _lcol1, _lcol2 = st.columns([0.65, 0.35])
    with _lcol1:
        st.header("⚙️ Settings")
    with _lcol2:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Sign Out", use_container_width=True):
            st.session_state.authenticated = False
            st.rerun()
    with st.form("settings_form"):
        st.subheader("💰 Budget to Actual")
        f_budget = st.number_input(
            "Flag unscheduled remaining over ($)",
            value=int(st.session_state.settings["budget_threshold"]),
            step=1000, min_value=0,
        )
        f_negative = st.number_input(
            "Flag negative budgets below -($)",
            value=int(st.session_state.settings["negative_threshold"]),
            step=50, min_value=0,
        )
        st.divider()
        st.subheader("📊 Variance Thresholds")
        applied = st.form_submit_button(
            "✔ Apply Budget Settings", type="primary", use_container_width=True)

    if applied:
        st.session_state.settings["budget_threshold"]   = float(f_budget)
        st.session_state.settings["negative_threshold"] = float(f_negative)
        st.rerun()

    # ── Variance thresholds — outside form for real-time slider↔input sync
    st.divider()
    st.subheader("📊 Variance Thresholds")

    if "_vmin" not in st.session_state:
        st.session_state._vmin = float(st.session_state.settings["variance_min"])
    if "_vmax" not in st.session_state:
        st.session_state._vmax = float(st.session_state.settings["variance_max"])

    # on_change callbacks keep slider ↔ number input in sync
    def _sl_min_changed():
        st.session_state._vmin = st.session_state._sl_min
        st.session_state._ni_min = st.session_state._sl_min
    def _ni_min_changed():
        v = float(st.session_state._ni_min or 0.0)
        st.session_state._vmin = v
        st.session_state._sl_min = min(v, 50.0)
    def _sl_max_changed():
        st.session_state._vmax = st.session_state._sl_max
        st.session_state._ni_max = st.session_state._sl_max
    def _ni_max_changed():
        v = float(st.session_state._ni_max or 0.0)
        st.session_state._vmax = v
        st.session_state._sl_max = min(v, 50.0)

    st.markdown("**📉 Scheduled but not actual**")
    st.caption("Flag when scheduled hrs exceed actual hrs by more than:")
    st.slider("Scheduled but not actual threshold", min_value=0.0, max_value=50.0, step=0.5, format="%.1f hrs",
              value=min(st.session_state._vmin, 50.0), key="_sl_min",
              label_visibility="collapsed", on_change=_sl_min_changed)
    st.number_input("Exact (hrs):", min_value=0.0, step=0.5,
                    value=float(st.session_state._vmin or 0.0),
                    key="_ni_min", on_change=_ni_min_changed)

    st.markdown("**📈 Actual but not scheduled**")
    st.caption("Flag when actual hrs exceed scheduled hrs by more than:")
    st.slider("Actual but not scheduled threshold", min_value=0.0, max_value=50.0, step=0.5, format="%.1f hrs",
              value=min(st.session_state._vmax, 50.0), key="_sl_max",
              label_visibility="collapsed", on_change=_sl_max_changed)
    st.number_input("Exact (hrs):", min_value=0.0, step=0.5,
                    value=float(st.session_state._vmax or 0.0),
                    key="_ni_max", on_change=_ni_max_changed)

    if st.button("✔ Apply Variance Settings", type="primary", use_container_width=True):
        st.session_state.settings["variance_min"] = st.session_state._vmin
        st.session_state.settings["variance_max"] = st.session_state._vmax
        st.rerun()

    st.divider()
    st.subheader("📋 Email Lookup")
    st.caption(f"{len(EMAIL_LOOKUP)} people configured")
    if st.checkbox("Show lookup table"):
        st.dataframe(
            [{"Name": k, "Email": v} for k, v in EMAIL_LOOKUP.items()],
            use_container_width=True, hide_index=True,
        )

budget_threshold   = st.session_state.settings["budget_threshold"]
negative_threshold = st.session_state.settings["negative_threshold"]
variance_min       = st.session_state.settings["variance_min"]
variance_max       = st.session_state.settings["variance_max"]

# ============================================================
# FILE UPLOAD + RUN BUTTON
#
# Before run:   show uploaders + Run button
# After run:    hide uploaders, lock bytes in session state
#               show "Change Files" button to reset
# ============================================================
st.subheader("📁 Upload Files")

if not st.session_state.run_triggered:
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        schedule_file = st.file_uploader(
            "Schedule File (.xlsx) — required",
            type=["xlsx", "csv"], key="schedule")
    with col_f2:
        openair_file = st.file_uploader(
            "OpenAir Report (.csv or .xlsx) — optional",
            type=["csv", "xlsx"], key="openair")

    run_col, _ = st.columns([0.25, 0.75])
    with run_col:
        run_clicked = st.button(
            "▶ Run Analysis",
            type="primary",
            disabled=not schedule_file,
            use_container_width=True,
        )

    if run_clicked and schedule_file:
        with st.spinner("Reading files…"):
            try:
                st.session_state.sched_bytes = bytes(schedule_file.read())
                st.session_state.oa_bytes = (
                    bytes(openair_file.read()) if openair_file else None
                )
                st.session_state.run_triggered = True
                st.session_state._analysis_done = False  # show spinner on fresh run
            except Exception as e:
                st.error(f"Could not read files: {e}")
        st.rerun()
    else:
        st.info("Upload your schedule file then click **▶ Run Analysis** to begin.")
        st.stop()

else:
    # Show change-files button but NOT while analysis is running
    chg_col, _ = st.columns([0.4, 0.6])
    with chg_col:
        if st.button("🔄 Change Files / Run Again", use_container_width=True):
            st.session_state.run_triggered = False
            st.session_state.sched_bytes   = None
            st.session_state.oa_bytes      = None
            st.rerun()

sched_bytes = st.session_state.sched_bytes
oa_bytes    = st.session_state.oa_bytes
file_hash   = _hash(sched_bytes)
has_openair = bool(oa_bytes)
oa_hash     = _hash(oa_bytes) if oa_bytes else ""

# ============================================================
# ANALYSIS — all runs under a single spinner to block UI
# ============================================================
_show_spinner = not st.session_state.get("_analysis_done", False)
with st.spinner("🔄 Running analysis — please wait…") if _show_spinner else st.empty():

    try:
        _wb_check = _load_wb(file_hash, sched_bytes)
        sheets    = list(_wb_check.sheetnames)
    except Exception as e:
        st.error(f"Could not open schedule file: {e}")
        st.stop()

    budget_issues, budget_sheet_name = [], None
    try:
        budget_issues, budget_sheet_name = run_budget(
            file_hash, sched_bytes,
            budget_threshold, DEFAULT_PROJECTION_THRESHOLD_PCT, negative_threshold,
        )
    except Exception as e:
        st.warning(f"Budget error: {e}")

    tracker_issues, tbd_projects, tracker_sheet_name, tracker_owner_map = [], [], None, {}
    try:
        tracker_issues, tbd_projects, tracker_sheet_name, tracker_owner_map = run_tracker(
            file_hash, sched_bytes)
    except Exception as e:
        st.warning(f"Tracker error: {e}")

    noncharge_all = {}
    time_coverage, time_report_date = {}, None
    try:
        if has_openair:
            noncharge_all = run_noncharge(oa_hash, oa_bytes)
            time_coverage, time_report_date = run_time_coverage(oa_hash, oa_bytes)
    except Exception as e:
        st.warning(f"Non-charge error: {e}")

    pto_schedule_data = {}
    try:
        _wb_pto = _load_wb(file_hash, sched_bytes)
        pto_schedule_data = get_pto_schedule(_wb_pto, _pto_month_list)
    except Exception:
        pass

    # NOTE: valid_people is computed AFTER the period selector below, since the
    # roster now depends on which month tabs the selected periods point at.
    st.session_state._analysis_done = True  # suppress spinner on settings reruns

    # OpenAir or schedule-derived periods
    available_months, future_months = [], []
    openair_error = None
    try:
        if has_openair:
            available_months, future_months = get_oa_periods(oa_hash, oa_bytes)
        else:
            available_months, future_months = get_sched_periods(
                file_hash, sched_bytes, active_month)
    except Exception as e:
        openair_error = str(e)

st.success(f"✅ Loaded **{len(sheets)}** sheet(s)")
c1, c2, c3 = st.columns(3)
c1.info(f"💰 Budget: **{budget_sheet_name or 'Not found'}**")
c2.info(f"📋 Tracker: **{tracker_sheet_name or 'Not found'}**")
c3.info(f"📅 Month: **{active_month}**")

if not has_openair:
    st.info("ℹ️ No OpenAir report uploaded — variance will show scheduled hours "
            "with actual hours as 0. Upload an OpenAir report for real actuals.")

st.divider()

# ============================================================
# PERIOD SELECTOR + VARIANCE
# ============================================================
selected_months = []
variance_issues = []
all_hours_issues = []
var_error       = None

if available_months:
    current_abbr   = datetime.now().strftime("%b")
    default_months = [m for m in available_months
                      if m.startswith(current_abbr) and m not in future_months]
    if not default_months:
        default_months = [m for m in available_months if m not in future_months][-1:]

    st.subheader("📅 Variance Period Selection")

    def _period_label(p):
        if p in future_months:
            return f"🔮 {p} (future — no actuals yet)"
        return p

    selected_months = st.multiselect(
        "Select period(s) for variance analysis:",
        options=available_months,
        default=default_months,
        format_func=_period_label,
        help=(
            "All 24 half-month periods for the year are shown. "
            "🔮 = future (scheduled hours only, actual = 0). "
            "Past periods use OpenAir actuals if uploaded."
        ),
    )

    if selected_months:
        with st.spinner("Computing variances…"):
            try:
                variance_issues, var_error = run_variance(
                    file_hash, sched_bytes,
                    oa_hash, oa_bytes,
                    tuple(selected_months),
                    -variance_min, variance_max,  # min is stored positive, negated here
                    active_month,
                )
                # Full "current month hours" list (matches + variances) used
                # only for the per-person Excel reports — the on-screen tab
                # above still shows flagged-only variances for quick review.
                all_hours_issues, _all_hours_err = run_variance(
                    file_hash, sched_bytes,
                    oa_hash, oa_bytes,
                    tuple(selected_months),
                    -variance_min, variance_max,
                    active_month,
                    include_all=True,
                )
            except Exception as e:
                var_error = str(e)

# ── Roster, derived from the SELECTED periods' month tabs ───────────
_roster_months = _months_from_periods(selected_months) or [active_month]
valid_people = set()
try:
    valid_people = get_valid_people(file_hash, sched_bytes, tuple(_roster_months))
except Exception:
    pass

# Non-charge time honours the same period selector as the variance tab.
noncharge_data = filter_noncharge(
    noncharge_all,
    periods=selected_months or None,
    people=sorted(valid_people) if valid_people else None,
)
noncharge_summary = noncharge_totals(noncharge_data)

# Who didn't enter their time in OpenAir last week. The roster comes from the
# schedule (or EMAIL_LOOKUP as a fallback), not the report — people who entered
# nothing don't appear in the report at all.
missing_time, missing_week = {}, None
if has_openair and time_report_date:
    _mw_start, _mw_end = previous_week(time_report_date)
    missing_week = format_week(_mw_start, _mw_end)
    _roster = sorted(valid_people) if valid_people else sorted(EMAIL_LOOKUP)
    missing_time = find_missing_time(time_coverage, _roster, _mw_start, _mw_end)

# Empty selection is the single most common cause of a blank report — e.g.
# early in a new month, before that month's tab has been scheduled. Say so
# plainly instead of silently producing workbooks with no Hours tab.
if selected_months and not all_hours_issues and not var_error:
    st.warning(
        f"No scheduled or actual hours found in **{', '.join(_roster_months)}** "
        f"for the selected period(s): {', '.join(selected_months)}. "
        "The Current Month Hours tab will be omitted from the reports. "
        "If that month isn't scheduled yet, pick a different period above."
    )

# ============================================================
# OWNER MAP
# Build owners_data for everyone in valid_people who has an email.
# This ensures ALL schedule staff appear in the report list.
# ============================================================
owners_data = defaultdict(lambda: {
    "email": None, "first_name": "there",
    "tracker": [], "budget": [], "variance": [], "noncharge": [],
})

# Seed ALL valid_people who have emails so nobody is missed
for person in valid_people:
    email = _lookup_email(person)
    if email and not owners_data[person]["email"]:
        from config import FIRST_NAMES
        owners_data[person]["email"]      = email
        owners_data[person]["first_name"] = FIRST_NAMES.get(person, FIRST_NAMES.get(NAME_ALIASES.get(person, ""), person))

for issue in tracker_issues:
    o = _normalize_name(issue.get("owner", ""))
    if not o or (valid_people and o not in valid_people):
        continue
    if not owners_data[o]["email"]:
        owners_data[o]["email"]      = issue.get("owner_email")
        owners_data[o]["first_name"] = issue.get("owner_first", o)
    owners_data[o]["tracker"].append(issue)

for issue in budget_issues:
    o = _normalize_name(issue.get("owner", ""))
    if not o or (valid_people and o not in valid_people):
        continue
    if not owners_data[o]["email"]:
        owners_data[o]["email"]      = issue.get("owner_email")
        owners_data[o]["first_name"] = issue.get("owner_first", o)
    owners_data[o]["budget"].append(issue)

# Build project → owner map from the full tracker (all projects, not just flagged ones)
# This ensures variance rows for staff on any project get routed to the correct project owner.
_project_owner_map = {
    code: _normalize_name(owner)
    for code, owner in tracker_owner_map.items()
    if code and owner
}
# Also fold in budget issues as a fallback for any codes not in tracker
for _issue in budget_issues:
    _code  = str(_issue.get("project_code", "")).strip()
    _owner = _normalize_name(_issue.get("owner", ""))
    if _code and _owner:
        _project_owner_map.setdefault(_code, _owner)

for v in all_hours_issues:
    person       = _normalize_name(v.get("person", ""))
    project_code = v.get("project_code", "")
    proj_owner   = _normalize_name(_project_owner_map.get(project_code, ""))

    # Everyone gets their OWN variance rows in their personal report
    if person and (not valid_people or person in valid_people):
        if not owners_data[person]["email"]:
            owners_data[person]["email"] = _lookup_email(person)
        owners_data[person]["variance"].append(v)

    # Project owner also gets ALL rows for their projects (including other staff rows)
    if proj_owner and proj_owner not in STAFF_NAMES and proj_owner != person and (not valid_people or proj_owner in valid_people):
        if not owners_data[proj_owner]["email"]:
            owners_data[proj_owner]["email"] = _lookup_email(proj_owner)
        owners_data[proj_owner]["variance"].append(v)

for _person, _entries in noncharge_data.items():
    p = _normalize_name(_person)
    if not p or (valid_people and p not in valid_people):
        continue
    if not owners_data[p]["email"]:
        owners_data[p]["email"]      = _entries[0].get("person_email") or _lookup_email(p)
        owners_data[p]["first_name"] = _entries[0].get("first_name", p)
    owners_data[p]["noncharge"].extend(_entries)

# Someone who entered no time at all has no tracker/budget/variance/non-charge
# rows either, so nothing above would have added them — and they're precisely
# the people who need the reminder. Seed them here.
for _person, _info in missing_time.items():
    p = _normalize_name(_person)
    if not p or (valid_people and p not in valid_people):
        continue
    if not owners_data[p]["email"]:
        owners_data[p]["email"]      = _info.get("person_email") or _lookup_email(p)
        owners_data[p]["first_name"] = _info.get("first_name", p)

# active_owners = everyone in valid_people who has an email
active_owners = {
    owner: data for owner, data in owners_data.items()
    if data.get("email")
    and (not valid_people or owner in valid_people)
}

# ============================================================
# ANALYSIS TABS
# ============================================================
tab1, tab2, tab3, tab4 = st.tabs([
    "📋 Project Tracker", "💰 Budget to Actual",
    "🕗 Non-Charge Time", "📊 Variance (OpenAir)"])

with tab1:
    st.header("Project Tracker — Known Projects")
    if not tracker_sheet_name:
        st.error("No Project Tracker tab found.")
    else:
        c1, c2 = st.columns(2)
        c1.metric("⚠️ Issues",           len(tracker_issues))
        c2.metric("📌 TBD / Pending SOW", len(tbd_projects))
        if tbd_projects:
            with st.expander(f"📌 {len(tbd_projects)} TBD / Pending SOW projects"):
                st.dataframe(
                    [{"Client": p.get("client",""), "Project Code": p.get("project_code",""),
                      "Status": p.get("status","TBD"), "Owner": p.get("owner",""),
                      "Budget": f"${p.get('budget',0):,.0f}",
                      "Notes": p.get("notes", "")}
                     for p in tbd_projects],
                    use_container_width=True, hide_index=True)
        if not tracker_issues:
            st.success("✅ No issues found!")
        else:
            rows = []
            for i in tracker_issues:
                missing = i.get("missing_rates", [])
                for m in missing:
                    rows.append({"Client": i.get("client",""),
                                 "Project Code": i.get("project_code",""),
                                 "Owner": i.get("owner",""),
                                 "To Be Reviewed": f"Missing {m} Rate",
                                 "Has Email": "✅" if i.get("owner_email") else "❌"})
                if not missing:
                    for prob in i.get("problems", []):
                        rows.append({"Client": i.get("client",""),
                                     "Project Code": i.get("project_code",""),
                                     "Owner": i.get("owner",""),
                                     "To Be Reviewed": prob,
                                     "Has Email": "✅" if i.get("owner_email") else "❌"})
            st.dataframe(rows, use_container_width=True, hide_index=True)

with tab2:
    st.header("Budget to Actual — Known Projects")
    st.caption(f"Flagging: negative ≤ -${negative_threshold:,.0f} | "
               f"unscheduled > ${budget_threshold:,.0f}")
    if not budget_sheet_name:
        st.error("No Budget to Actual tab found.")
    else:
        neg = [i for i in budget_issues if i.get("type") == "negative"]
        np_ = [i for i in budget_issues if i.get("type") == "not_projected"]
        c1, c2, c3 = st.columns(3)
        c1.metric("🔴 Over Budget", len(neg))
        c2.metric("🟡 Unscheduled", len(np_))
        c3.metric("⚠️ Total",       len(budget_issues))
        if budget_issues:
            st.dataframe(
                [{"Client": i.get("client",""), "Project Code": i.get("project_code",""),
                  "Owner": i.get("owner",""),
                  "Budget": f"${i.get('budget',0):,.0f}",
                  "Remaining": f"${i.get('remaining',0):,.0f}",
                  "Flag": i.get("description",""),
                  "Has Email": "✅" if i.get("owner_email") else "❌"}
                 for i in budget_issues],
                use_container_width=True, hide_index=True)

with tab3:
    st.header("Non-Charge Time")
    if not has_openair:
        st.info("ℹ️ Upload the time report to see non-charge time.")
    elif not selected_months:
        st.warning("Select at least one period above.")
    elif not noncharge_data:
        st.success("✅ No non-charge time logged in the selected period(s).")
    else:
        _avail_total = sum(t["available_time"] for t in noncharge_summary)
        _resp_total  = sum(t["needs_response"] for t in noncharge_summary)
        m1, m2, m3 = st.columns(3)
        m1.metric("Non-Charge Hours", f"{sum(t['total'] for t in noncharge_summary):,.1f}")
        m2.metric("🟡 Available Time", f"{_avail_total:,.1f}")
        m3.metric("Entries Needing a Response", _resp_total)

        if len(selected_months) > 1:
            st.info(f"Showing across {len(selected_months)} periods: "
                    f"{', '.join(selected_months)}")

        st.subheader("By Person")
        st.dataframe(
            [{"Person": DISPLAY_NAMES.get(t["person"], t["person"]),
              "Available Time": t["available_time"],
              "Training": t["training"], "PTO": t["pto"], "Holiday": t["holiday"],
              "Other": round(t["other"], 1), "Total": t["total"],
              "Needs Response": t["needs_response"],
              "Has Email": "✅" if t["person_email"] else "❌"}
             for t in sorted(noncharge_summary, key=lambda t: _rank(t["person"]))],
            use_container_width=True, hide_index=True)

        st.subheader("Detail")
        _only_flagged = st.checkbox(
            "Only show entries needing a response", value=True, key="nc_flagged")
        _detail = [e for entries in noncharge_data.values() for e in entries
                   if e["needs_response"] or not _only_flagged]
        _detail.sort(key=lambda e: (_rank(e["person"]), not e["available_time"], e["date"]))
        if not _detail:
            st.success("✅ Every non-charge entry has a note.")
        else:
            st.dataframe(
                [{"Person": DISPLAY_NAMES.get(e["person"], e["person"]),
                  "Date": e["date_str"], "Task": e["task"],
                  "Hrs": e["hours"],
                  "Notes": e["notes"] or "—",
                  "Description": e["description"] or "—",
                  "Flag": "🟡 Available Time" if e["available_time"]
                          else ("⚠️ No note" if e["needs_response"] else "")}
                 for e in _detail],
                use_container_width=True, hide_index=True)

with tab4:
    st.header("Actual vs Schedule Variance (OpenAir)")
    if missing_time:
        _none  = [p for p, m in missing_time.items() if m["status"] == "none"]
        _short = [p for p, m in missing_time.items() if m["status"] == "partial"]
        _parts = []
        if _none:
            _parts.append(f"**No time entered:** {', '.join(sorted(_none))}")
        if _short:
            _parts.append("**Partial:** " + ", ".join(
                f"{p} ({missing_time[p]['hours']:g} hrs)" for p in sorted(_short)))
        st.warning(f"⏰ Time entry for {missing_week} — " + " · ".join(_parts)
                   + "  \nA reminder is added to these people's reports.")
    if not has_openair:
        st.info("ℹ️ No OpenAir report uploaded — showing scheduled hours with actual = 0.")
    if openair_error:
        st.error(f"Error loading periods: {openair_error}")
    if var_error:
        st.error(f"Variance error: {var_error}")
    if not selected_months:
        st.warning("Select at least one period above.")
    elif not variance_issues:
        st.success(f"✅ No variances outside "
                   f"[{variance_min:+.0f}, {variance_max:+.0f}] hrs.")
    else:
        if len(selected_months) > 1:
            st.info(f"Showing across {len(selected_months)} periods: "
                    f"{', '.join(selected_months)}")
        st.metric("⚠️ Variances Found", len(variance_issues))
        _sorted_v = sorted(variance_issues, key=lambda v: (_rank(v.get("person","")), v.get("project_code","")))
        st.dataframe(
            [{"Person": v.get("person",""), "Project": v.get("project_code",""),
              "Period": v.get("period",""),
              "Actual Hrs": v.get("actual_hours",0), "Sched Hrs": v.get("sched_hours",0),
              "Diff": v.get("difference",0), "To Review": v.get("question",""),
              "Future": "🔮" if v.get("is_future") else ""}
             for v in _sorted_v],
            use_container_width=True, hide_index=True)

# ============================================================
# COMBINED REPORTS (EXCEL — replaces the old emailed reports)
# ============================================================
st.divider()
st.header("📊 Combined Reports")
st.caption("One Excel workbook per person · Project Tracker · Budget · "
           "TBD/Pending SOW · Variance · Non-Charge Time · PTO — "
           "all packaged into a single ZIP for you to save and distribute.")

if not active_owners:
    st.info("No staff found with email addresses — check config.py EMAIL_LOOKUP.")
    st.stop()

st.metric("TAS Members", len(active_owners))
all_owner_keys = sorted(active_owners.keys(), key=_rank)

if not st.session_state.selected_owners.issubset(set(all_owner_keys)):
    st.session_state.selected_owners = set(all_owner_keys)

st.markdown("**Select people to include:**")

def _select_all():
    for owner in all_owner_keys:
        st.session_state[f"chk_{owner}"] = True
    st.session_state.selected_owners = set(all_owner_keys)

def _deselect_all():
    for owner in all_owner_keys:
        st.session_state[f"chk_{owner}"] = False
    st.session_state.selected_owners = set()

col_sa, col_da, _ = st.columns([0.15, 0.18, 0.67])
with col_sa:
    st.button("✅ Select All", on_click=_select_all)
with col_da:
    st.button("⬜ Deselect All", on_click=_deselect_all)

for owner in all_owner_keys:
    data       = active_owners[owner]
    first_name = data.get("first_name", owner)
    _display   = DISPLAY_NAMES.get(owner, owner)
    label = (f"**{first_name} ({_display})** — "
             f"Tracker: {len(data.get('tracker', []))} · "
             f"Budget: {len(data.get('budget', []))} · "
             f"Non-Charge: {len(data.get('noncharge', []))} · "
             f"Hours: {len(data.get('variance', []))}"
             + (" · ⏰ time entry" if owner in missing_time else ""))
    chk_key = f"chk_{owner}"
    if chk_key not in st.session_state:
        st.session_state[chk_key] = owner in st.session_state.selected_owners
    checked = st.checkbox(label, key=chk_key)
    if checked:
        st.session_state.selected_owners.add(owner)
    else:
        st.session_state.selected_owners.discard(owner)


def _safe_filename(name: str) -> str:
    keep = "".join(c if c.isalnum() or c in (" ", "_", "-") else "" for c in name)
    return keep.strip().replace(" ", "_") + ".xlsx"


def _build_payload(owner_list):
    """Assemble the build_person_workbook() kwargs for each owner, same shape
    the old build_html_email() calls used — just renamed 'owner' → included as a key."""
    payload = []
    for owner in sorted(owner_list, key=_rank):
        if owner not in active_owners:
            continue
        data          = active_owners[owner]
        variance_list = data.get("variance", [])
        payload.append(dict(
            owner            = owner,
            first_name       = data.get("first_name", owner),
            tracker_issues   = data.get("tracker", []),
            budget_issues    = data.get("budget", []),
            tbd_projects     = tbd_projects,
            variance_issues  = variance_list,
            noncharge_data   = data.get("noncharge", []),
            missing_time     = missing_time.get(owner),
            pto_schedule     = pto_schedule_data,
            pto_months       = _pto_month_list,
            has_openair      = has_openair,
            no_openair_note  = not has_openair and bool(variance_list),
            selected_months  = selected_months if len(selected_months) > 1 else None,
            is_staff         = owner in STAFF_NAMES,
            filename         = _safe_filename(DISPLAY_NAMES.get(owner, owner)),
        ))
    return payload


def _tabs_included(person: dict) -> list[str]:
    tabs = []
    if person["tracker_issues"]:
        tabs.append("Tracker")
    if person["budget_issues"]:
        tabs.append("Budget")
    if [p for p in tbd_projects if p.get("owner") == person["owner"]]:
        tabs.append("TBD")
    if person["variance_issues"]:
        tabs.append("Current Month Hours")
    if person["noncharge_data"]:
        tabs.append("Non-Charge Time")
    if person.get("missing_time"):
        tabs.append("⏰ Time entry reminder")
    person_pto = pto_schedule_data.get(person["owner"], {})
    if person_pto and any(person_pto.get(m, 0) for m in _pto_month_list):
        tabs.append("PTO")
    return tabs


# ---- Preview: what each selected person's workbook will contain ----
st.markdown("**Preview (tabs included per person):**")
selected_payload = _build_payload(st.session_state.selected_owners)
for person in selected_payload:
    tabs_included = _tabs_included(person)
    _display = DISPLAY_NAMES.get(person["owner"], person["owner"])
    if tabs_included:
        st.write(f"👤 **{person['first_name']}** ({_display}) — {', '.join(tabs_included)}")
    else:
        st.write(f"👤 **{person['first_name']}** ({_display}) — "
                 f"_no applicable sections, will be skipped in the ZIP_")

# ---- Generate & download (single click — no separate "Generate" step) ----
st.divider()
all_payload = _build_payload(all_owner_keys)

with st.spinner("Building workbooks…"):
    selected_zip_bytes = build_reports_zip(selected_payload) if selected_payload else b""
    all_zip_bytes      = build_reports_zip(all_payload) if all_payload else b""


def _included_count(zip_bytes: bytes) -> int:
    """How many workbooks actually ended up in the zip (people with no
    applicable data are silently skipped by build_reports_zip)."""
    if not zip_bytes:
        return 0
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        return len(zf.namelist())


# ---- Consolidated non-charge report (whole team, whole month) ----
# Scoped to the full month behind the selected period(s), not the half-month
# selector — "all the non-charge time for the month".
_nc_months        = months_from_periods(selected_months)
_nc_month_periods = periods_in_months(noncharge_all, _nc_months)
if _nc_months and not _nc_month_periods:
    # Month selected but nothing logged in it. Must stay empty — passing
    # periods=None here would mean "no filter" and dump the whole year.
    _nc_consolidated = {}
else:
    _nc_consolidated = filter_noncharge(
        noncharge_all,
        periods=_nc_month_periods or None,
        people=sorted(valid_people) if valid_people else None,
    )
_nc_year  = next((e["date"].year for v in _nc_consolidated.values() for e in v), None)
_nc_label = (" / ".join(_nc_months) + (f" {_nc_year}" if _nc_year else "")).strip() \
            or "All periods"

_nc_bytes = b""
if _nc_consolidated:
    try:
        _nc_bytes = build_consolidated_noncharge(
            _nc_consolidated, month_label=_nc_label, periods=_nc_month_periods,
            display_names=DISPLAY_NAMES, rank_fn=_rank)
    except Exception as e:
        st.warning(f"Consolidated non-charge report error: {e}")

_nc_entries = sum(len(v) for v in _nc_consolidated.values())
st.download_button(
    f"⬇️ Download Consolidated Non-Charge Report — {_nc_label or 'month'} "
    f"({len(_nc_consolidated)} people, {_nc_entries} entries)",
    data=_nc_bytes,
    file_name=f"noncharge_{(_nc_label or 'month').replace(' ', '_').replace('/', '-')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    disabled=not _nc_bytes,
    use_container_width=True,
)
st.caption("One workbook covering the whole team for the month — every "
           "non-charge entry with its notes and description. This is a review "
           "copy for managers and is **not** included in the team ZIP below.")
st.divider()

col_b1, col_b2 = st.columns(2)
with col_b1:
    n_included = _included_count(selected_zip_bytes)
    st.download_button(
        f"⬇️ Download Selected Reports ({n_included}/{len(selected_payload)})",
        data=selected_zip_bytes,
        file_name=f"scheduling_reports_selected_{datetime.now().strftime('%Y%m%d')}.zip",
        mime="application/zip",
        disabled=not selected_payload,
        type="primary",
        use_container_width=True,
    )
    if selected_payload and n_included < len(selected_payload):
        st.caption(f"{len(selected_payload) - n_included} selected "
                   f"person(s) had no applicable data and were skipped.")
with col_b2:
    n_included_all = _included_count(all_zip_bytes)
    st.download_button(
        f"⬇️ Download All Reports ({n_included_all}/{len(all_payload)})",
        data=all_zip_bytes,
        file_name=f"scheduling_reports_all_{datetime.now().strftime('%Y%m%d')}.zip",
        mime="application/zip",
        disabled=not all_payload,
        use_container_width=True,
    )
