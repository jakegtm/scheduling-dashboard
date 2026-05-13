from __future__ import annotations
# ============================================================
# app.py — GTM Scheduling Analyzer  |  streamlit run app.py
# ============================================================

import gc
import hashlib
import io
import warnings
from collections import defaultdict
from datetime import datetime

import openpyxl
import streamlit as st

from config import (
    SENDER_EMAIL, SENDER_NAMES, EMAIL_LOOKUP, INTERN_NAMES,
    DEFAULT_BUDGET_THRESHOLD, DEFAULT_NEGATIVE_THRESHOLD,
    DEFAULT_PROJECTION_THRESHOLD_PCT,
    DEFAULT_VARIANCE_MIN, DEFAULT_VARIANCE_MAX,
)
from email_utils import email_configured, EMAIL_OK, send_emails_batch
from processors.budget_actual   import process_budget_actual
from processors.project_tracker import process_project_tracker
from processors.variance        import (
    parse_openair_report, read_schedule_hours,
    compute_variances, get_available_months, filter_by_months,
)
from processors.utilization import process_utilization, build_utilization_emails

warnings.filterwarnings("ignore", category=UserWarning)

st.set_page_config(page_title="GTM Scheduling Analyzer",
                   layout="wide", page_icon="📊")
st.title("📊 GTM Scheduling Analyzer")
st.caption(f"Today: {datetime.now().strftime('%A, %B %d, %Y')}")

if not email_configured():
    st.warning("⚠️ **Email credentials not configured.** "
               "Add SendGrid keys in Streamlit Secrets to enable sending.")

# ============================================================
# HELPERS
# ============================================================

def _hash(b: bytes) -> str:
    return hashlib.md5(b).hexdigest()

def _sender_name() -> str:
    try:
        from_email = st.secrets["email"]["from_email"]
    except (KeyError, FileNotFoundError):
        from_email = SENDER_EMAIL
    return SENDER_NAMES.get(from_email, "Jake")

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

# ============================================================
# SESSION STATE — must init before anything renders
# ============================================================
def _init_state():
    defaults = {
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
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

_init_state()

# ============================================================
# CACHING
# ============================================================
@st.cache_resource(max_entries=2, show_spinner=False)
def _load_wb(file_hash: str, _file_bytes: bytes):
    return openpyxl.load_workbook(io.BytesIO(_file_bytes), data_only=True)

@st.cache_data(show_spinner=False)
def run_budget(file_hash, _b, budget_thr, proj_pct, neg_thr):
    wb = _load_wb(file_hash, _b)
    sheet = _find_sheet(wb.sheetnames, ["budget to actual", "budget"])
    if not sheet: return [], None
    return process_budget_actual(wb[sheet], budget_thr, proj_pct, neg_thr), sheet

@st.cache_data(show_spinner=False)
def run_tracker(file_hash, _b):
    wb = _load_wb(file_hash, _b)
    sheet = _find_sheet(wb.sheetnames, ["project tracker", "tracker"])
    if not sheet: return [], [], None
    issues, tbd = process_project_tracker(wb[sheet])
    return issues, tbd, sheet

@st.cache_data(show_spinner=False)
def run_utilization(file_hash, _b, month):
    wb = _load_wb(file_hash, _b)
    try:
        data = process_utilization(wb, target_month=month)
    except Exception:
        data = []
    gc.collect()
    return data

@st.cache_data(show_spinner=False)
def run_openair(oa_hash, _oa):
    actual = parse_openair_report(io.BytesIO(_oa))
    available, future = get_available_months(actual)
    gc.collect()
    return actual, available, future

@st.cache_data(show_spinner=False)
def run_variance(file_hash, _b, oa_hash, _oa, selected_months_tuple,
                 var_min, var_max, active_month):
    wb = _load_wb(file_hash, _b)
    sched_sheet = next(
        (s for s in wb.sheetnames if s.lower().startswith(active_month[:3].lower())),
        None,
    )
    if not sched_sheet:
        return [], f"No sheet found for {active_month}"
    try:
        actual, _, _ = run_openair(oa_hash, _oa)
        filtered     = filter_by_months(actual, list(selected_months_tuple))
        sched        = read_schedule_hours(wb, sched_sheet)
        variances    = compute_variances(filtered, sched,
                                         min_diff=var_min, max_diff=var_max)
        gc.collect()
        return variances, None
    except Exception as e:
        gc.collect()
        return [], str(e)

@st.cache_data(show_spinner=False)
def get_valid_people(file_hash, _b, active_month):
    wb = _load_wb(file_hash, _b)
    for name in wb.sheetnames:
        if name.lower().startswith(active_month[:3].lower()):
            ws = wb[name]
            people = set()
            for col in range(7, 45):
                val = ws.cell(row=2, column=col).value
                if val:
                    people.add(str(val).strip())
            return people
    return set()

# ============================================================
# SIDEBAR — Settings (st.form prevents reruns on +/- clicks)
# ============================================================
active_month = datetime.now().strftime("%B")
sender_name  = _sender_name()

with st.sidebar:
    st.header("⚙️ Settings")

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
        vc1, vc2 = st.columns(2)
        with vc1:
            f_var_min = st.number_input(
                "Min (flag if <=)",
                value=float(st.session_state.settings["variance_min"]),
                step=1.0,
            )
        with vc2:
            f_var_max = st.number_input(
                "Max (flag if >=)",
                value=float(st.session_state.settings["variance_max"]),
                step=1.0,
            )
        applied = st.form_submit_button("✔ Apply Settings", type="primary",
                                        use_container_width=True)

    if applied:
        st.session_state.settings = {
            "budget_threshold":   float(f_budget),
            "negative_threshold": float(f_negative),
            "variance_min":       float(f_var_min),
            "variance_max":       float(f_var_max),
        }
        st.rerun()

    st.divider()
    st.subheader("📋 Email Lookup")
    st.caption(f"{len(EMAIL_LOOKUP)} people configured")
    if st.checkbox("Show lookup table"):
        st.dataframe(
            [{"Name": k, "Email": v} for k, v in EMAIL_LOOKUP.items()],
            use_container_width=True, hide_index=True,
        )

if not EMAIL_OK:
    st.sidebar.warning("SendGrid keys not found — previews work but sending is disabled.")

# Pull applied settings
budget_threshold   = st.session_state.settings["budget_threshold"]
negative_threshold = st.session_state.settings["negative_threshold"]
variance_min       = st.session_state.settings["variance_min"]
variance_max       = st.session_state.settings["variance_max"]

# ============================================================
# FILE UPLOAD + RUN BUTTON
#
# Flow:
#   1. Show file uploaders (not yet run)
#   2. User clicks "▶ Run Analysis" → bytes locked in session state
#   3. App processes using locked bytes (no re-reading on reruns)
#   4. "🔄 Change Files" resets the lock so user can re-upload
# ============================================================
st.subheader("📁 Upload Files")

if not st.session_state.run_triggered:
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        schedule_file = st.file_uploader(
            "Schedule File (.xlsx)", type=["xlsx", "csv"], key="schedule")
    with col_f2:
        openair_file = st.file_uploader(
            "OpenAir Report (.csv or .xlsx) — optional",
            type=["csv", "xlsx"], key="openair")

    st.divider()
    run_col, _ = st.columns([0.25, 0.75])
    with run_col:
        run_clicked = st.button(
            "▶ Run Analysis",
            type="primary",
            disabled=not schedule_file,
            use_container_width=True,
        )

    if run_clicked:
        try:
            st.session_state.sched_bytes = bytes(schedule_file.read())
            st.session_state.oa_bytes = (
                bytes(openair_file.read()) if openair_file else None
            )
            st.session_state.run_triggered = True
        except Exception as e:
            st.error(f"Could not read files: {e}")
        st.rerun()
    else:
        st.info("Upload your schedule file then click **▶ Run Analysis** to begin.")
        st.stop()

else:
    # Locked — show change-files option
    chg_col, _ = st.columns([0.35, 0.65])
    with chg_col:
        if st.button("🔄 Change Files / Run Again", use_container_width=True):
            st.session_state.run_triggered = False
            st.session_state.sched_bytes   = None
            st.session_state.oa_bytes      = None
            st.rerun()
    st.success("Files locked for analysis.")

sched_bytes = st.session_state.sched_bytes
oa_bytes    = st.session_state.oa_bytes
file_hash   = _hash(sched_bytes)

# ============================================================
# LOAD + PROCESS (all cached)
# ============================================================
with st.status("🔄 Analyzing schedule file...", expanded=True) as status:
    st.write("📂 Loading workbook...")

    try:
        _wb_check = _load_wb(file_hash, sched_bytes)
        sheets    = list(_wb_check.sheetnames)
    except Exception as e:
        st.error(f"Could not open schedule file: {e}")
        st.stop()

    st.write("💰 Budget to Actual...")
    budget_issues, budget_sheet_name = [], None
    try:
        budget_issues, budget_sheet_name = run_budget(
            file_hash, sched_bytes,
            budget_threshold, DEFAULT_PROJECTION_THRESHOLD_PCT, negative_threshold,
        )
    except Exception as e:
        st.warning(f"Budget error: {e}")

    st.write("📋 Project Tracker...")
    tracker_issues, tbd_projects, tracker_sheet_name = [], [], None
    try:
        tracker_issues, tbd_projects, tracker_sheet_name = run_tracker(
            file_hash, sched_bytes)
    except Exception as e:
        st.warning(f"Tracker error: {e}")

    st.write("📈 Utilization...")
    util_data = []
    try:
        util_data = run_utilization(file_hash, sched_bytes, active_month)
    except Exception as e:
        st.warning(f"Utilization error: {e}")

    st.write("👥 Valid staff...")
    valid_people = set()
    try:
        valid_people = get_valid_people(file_hash, sched_bytes, active_month)
    except Exception:
        pass

    status.update(label="✅ Analysis complete!", state="complete", expanded=False)

st.success(f"Loaded **{len(sheets)}** sheet(s)")
c1, c2, c3 = st.columns(3)
c1.info(f"💰 Budget: **{budget_sheet_name or 'Not found'}**")
c2.info(f"📋 Tracker: **{tracker_sheet_name or 'Not found'}**")
c3.info(f"📅 Month: **{active_month}**")
st.divider()

# ============================================================
# OPENAIR + PERIOD SELECTOR
# ============================================================
has_openair      = False
available_months = []
future_months    = []
openair_error    = None
actual_data_full = {}
oa_hash          = ""

if oa_bytes:
    with st.status("🔄 Processing OpenAir report...", expanded=True) as oa_status:
        try:
            st.write("📊 Parsing time entries...")
            oa_hash = _hash(oa_bytes)
            actual_data_full, available_months, future_months = run_openair(
                oa_hash, oa_bytes)
            st.write("📅 Identifying available periods...")
            has_openair = True
            oa_status.update(label="✅ OpenAir loaded!", state="complete",
                             expanded=False)
        except Exception as e:
            openair_error = str(e)
            oa_status.update(label="❌ OpenAir error", state="error",
                             expanded=True)

# Period selector
selected_months = []
variance_issues = []

if has_openair and available_months:
    current_abbr   = datetime.now().strftime("%b")
    default_months = [m for m in available_months
                      if m.startswith(current_abbr) and m not in future_months]
    if not default_months:
        default_months = [m for m in available_months if m not in future_months][-1:]

    st.subheader("📅 Variance Period Selection")

    def _period_label(p):
        return f"🔮 {p} (future — no actuals yet)" if p in future_months else p

    selected_months = st.multiselect(
        "Select period(s) for variance analysis:",
        options=available_months,
        default=default_months,
        format_func=_period_label,
        help="🔮 = future period. Default is the current month.",
    )

    if selected_months:
        with st.status("🔄 Computing variances...", expanded=True) as var_status:
            st.write(f"📊 {', '.join(selected_months)}...")
            variance_issues, var_error = run_variance(
                file_hash, sched_bytes,
                oa_hash, oa_bytes,
                tuple(selected_months),
                variance_min, variance_max,
                active_month,
            )
            if var_error:
                st.warning(f"Variance error: {var_error}")
            st.write(f"✅ {len(variance_issues)} variance(s) found")
            var_status.update(
                label=f"✅ Variance — {len(variance_issues)} flagged",
                state="complete", expanded=False)

# ============================================================
# OWNER MAP — merge tracker + budget + variance + util
# ============================================================
owners_data = defaultdict(lambda: {
    "email": None, "first_name": "there",
    "tracker": [], "budget": [], "variance": [], "util": [],
})

for issue in tracker_issues:
    o = issue.get("owner", "")
    if not o or (valid_people and o not in valid_people):
        continue
    if not owners_data[o]["email"]:
        owners_data[o]["email"]      = issue.get("owner_email")
        owners_data[o]["first_name"] = issue.get("owner_first", o)
    owners_data[o]["tracker"].append(issue)

for issue in budget_issues:
    o = issue.get("owner", "")
    if not o or (valid_people and o not in valid_people):
        continue
    if not owners_data[o]["email"]:
        owners_data[o]["email"]      = issue.get("owner_email")
        owners_data[o]["first_name"] = issue.get("owner_first", o)
    owners_data[o]["budget"].append(issue)

for v in variance_issues:
    p = v.get("person", "")
    if not p or (valid_people and p not in valid_people):
        continue
    if not owners_data[p]["email"]:
        owners_data[p]["email"] = _lookup_email(p)
    owners_data[p]["variance"].append(v)

# ALL staff with utilization data are always included
for u in util_data:
    p = u.get("person", "")
    if not p or (valid_people and p not in valid_people):
        continue
    if not owners_data[p]["email"]:
        owners_data[p]["email"]      = u.get("person_email") or _lookup_email(p)
        owners_data[p]["first_name"] = u.get("first_name", p)
    owners_data[p]["util"].append(u)

# active_owners = anyone with at least one item across all four categories
active_owners = {
    owner: data for owner, data in owners_data.items()
    if (data["tracker"] or data["budget"] or data["variance"] or data["util"])
    and (not valid_people or owner in valid_people)
}

# ============================================================
# ANALYSIS TABS
# ============================================================
tab1, tab2, tab3, tab4 = st.tabs([
    "📋 Project Tracker", "💰 Budget to Actual",
    "📈 Utilization",     "📊 Variance (OpenAir)"])

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
                      "Budget": f"${p.get('budget',0):,.0f}"}
                     for p in tbd_projects],
                    use_container_width=True, hide_index=True)
        if not tracker_issues:
            st.success("✅ No issues found!")
        else:
            st.dataframe(
                [{"Client": i.get("client",""), "Project Code": i.get("project_code",""),
                  "Owner": i.get("owner",""),
                  "Missing Rates": ", ".join(i.get("missing_rates",[])),
                  "Has Email": "✅" if i.get("owner_email") else "❌"}
                 for i in tracker_issues],
                use_container_width=True, hide_index=True)

with tab2:
    st.header("Budget to Actual — Known Projects")
    st.caption(f"Flagging: negative < -${negative_threshold:,.0f} | "
               f"unscheduled > ${budget_threshold:,.0f}")
    if not budget_sheet_name:
        st.error("No Budget to Actual tab found.")
    else:
        neg = [i for i in budget_issues if i.get("type") == "negative"]
        np_ = [i for i in budget_issues if i.get("type") == "not_projected"]
        c1, c2, c3 = st.columns(3)
        c1.metric("🔴 Over Budget",  len(neg))
        c2.metric("🟡 Unscheduled",  len(np_))
        c3.metric("⚠️ Total",        len(budget_issues))
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
    st.header(f"Utilization — {active_month}")
    if not util_data:
        st.warning("No utilization data found. Make sure the workbook has a "
                   "'Utilization by Month' tab.")
    else:
        st.dataframe(
            [{"Role": u.get("role",""), "Person": u.get("person",""),
              "Chargeable": u.get("chargeable","-"), "Holiday": u.get("holiday","-"),
              "PTO": u.get("pto","-"), "Month Total": u.get("month_total","-"),
              "Remaining": u.get("remaining","-"),
              "Utilization": f"{u['utilization_pct']:.1f}%" if u.get("utilization_pct") is not None else "-",
              "Goal":        f"{u['goal_pct']:.0f}%"        if u.get("goal_pct")        is not None else "-",
              "Difference":  f"{u['difference_pct']:+.1f}%" if u.get("difference_pct") is not None else "-",
              "Has Email":   "✅" if u.get("person_email") else "❌"}
             for u in util_data],
            use_container_width=True, hide_index=True)

with tab4:
    st.header("Actual vs Schedule Variance (OpenAir)")
    if not oa_bytes:
        st.info("Upload an OpenAir file above to enable variance analysis.")
    elif openair_error:
        st.error(f"Error parsing OpenAir file: {openair_error}")
    elif not selected_months:
        st.warning("Select at least one period above.")
    elif not variance_issues:
        st.success(f"✅ No variances outside "
                   f"[{variance_min:+.0f}, {variance_max:+.0f}] hrs.")
    else:
        if len(selected_months) > 1:
            st.info(f"Showing variances across {len(selected_months)} periods: "
                    f"{', '.join(selected_months)}")
        st.metric("⚠️ Variances Found", len(variance_issues))
        st.dataframe(
            [{"Person": v.get("person",""), "Project": v.get("project_code",""),
              "Period": v.get("period",""),
              "Actual Hrs": v.get("actual_hours",0),
              "Sched Hrs":  v.get("sched_hours",0),
              "Diff":       v.get("difference",0),
              "To Review":  v.get("question",""),
              "Future":     "🔮" if v.get("is_future") else ""}
             for v in variance_issues],
            use_container_width=True, hide_index=True)

# ============================================================
# COMBINED EMAILS
# ============================================================
st.divider()
st.header("📧 Combined Emails")
st.caption("One email per person · Project Tracker · Budget · "
           "TBD/Pending SOW · Variance · Utilization")

if not active_owners:
    st.info("No flagged items found — no emails to send.")
    st.stop()

st.metric("People with items to review", len(active_owners))
all_owner_keys = sorted(active_owners.keys())

# Keep selection in sync
if not st.session_state.selected_owners.issubset(set(all_owner_keys)):
    st.session_state.selected_owners = set(all_owner_keys)

col_sa, col_da, _ = st.columns([0.15, 0.18, 0.67])
with col_sa:
    if st.button("✅ Select All"):
        st.session_state.selected_owners = set(all_owner_keys)
        st.rerun()
with col_da:
    if st.button("⬜ Deselect All"):
        st.session_state.selected_owners = set()
        st.rerun()

st.markdown("**Select recipients:**")
for owner in all_owner_keys:
    data       = active_owners[owner]
    first_name = data.get("first_name", owner)
    email_str  = data.get("email") or "⚠️ no email"
    label = (f"**{first_name} ({owner})** · {email_str} — "
             f"Tracker: {len(data['tracker'])} · "
             f"Budget: {len(data['budget'])} · "
             f"Util: {len(data['util'])} · "
             f"Variance: {len(data['variance'])}")
    checked = owner in st.session_state.selected_owners
    if st.checkbox(label, value=checked, key=f"chk_{owner}"):
        st.session_state.selected_owners.add(owner)
    else:
        st.session_state.selected_owners.discard(owner)

# ---- Build emails ----
util_emails_by_person = {
    e["person"]: e
    for e in build_utilization_emails(
        util_data, month=active_month, sender_name=sender_name)
}

st.markdown("**Email previews:**")
combined_emails = []

for owner in sorted(st.session_state.selected_owners):
    if owner not in active_owners:
        continue
    data         = active_owners[owner]
    person_email = data.get("email") or _lookup_email(owner)
    first_name   = data.get("first_name", owner)
    if not person_email:
        continue

    is_intern     = owner in INTERN_NAMES
    tracker_list  = data.get("tracker", [])
    budget_list   = data.get("budget", [])
    variance_list = data.get("variance", [])
    if is_intern:
        variance_list = [v for v in variance_list if v.get("person") == owner]

    owner_tbd = [p for p in tbd_projects if p.get("owner") == owner]
    sections  = []

    # --- Missing rates ---
    if tracker_list:
        lines = [
            "The following projects assigned to you are missing billing rates. "
            "Please review and update as soon as possible.\n"
        ]
        for issue in tracker_list:
            # Use missing_rates (just the role names, e.g. ["Intern"])
            missing = issue.get("missing_rates", [])
            lines.append(
                f"  - {issue.get('client','')} | {issue.get('project_code','')}\n"
                f"    Missing: {', '.join(missing)}"
            )
        sections.append("\n".join(lines))

    # --- Budget ---
    if budget_list:
        lines = ["Please review the following budget items:\n"]
        for issue in budget_list:
            lines.append(
                f"  - {issue.get('client','')} | {issue.get('project_code','')}: "
                f"${issue.get('remaining',0):,.0f} remaining "
                f"({issue.get('description','')})"
            )
        sections.append("\n".join(lines))

    # --- TBD / Pending SOW ---
    if owner_tbd:
        lines = [
            "The following projects have TBD or Pending SOW budgets. "
            "If you have any updates, please reply — otherwise no action needed.\n"
        ]
        for proj in owner_tbd:
            label = f" [{proj.get('status','TBD')}]"
            lines.append(f"  - {proj.get('client','')} | {proj.get('project_code','')}{label}")
        sections.append("\n".join(lines))

    # --- Variance ---
    if variance_list:
        if is_intern:
            lines = ["Please see your variance for the current period:\n"]
            for v in variance_list:
                lines.append(
                    f"  - {v.get('project_code','')} | {v.get('period','')} | "
                    f"Actual: {v.get('actual_hours',0)}h  "
                    f"Scheduled: {v.get('sched_hours',0)}h  "
                    f"Diff: {v.get('difference',0):+.1f}h"
                )
                if v.get("question"):
                    lines.append(f"    → {v['question']}")
        else:
            lines = [
                "Please see the variance summary for projects assigned to you. "
                "Review any flagged items and reply with updates as needed.\n"
            ]
            for v in variance_list:
                lines.append(
                    f"  - {v.get('person','')} | {v.get('project_code','')} | "
                    f"{v.get('period','')} | "
                    f"Actual: {v.get('actual_hours',0)}h  "
                    f"Scheduled: {v.get('sched_hours',0)}h  "
                    f"Diff: {v.get('difference',0):+.1f}h"
                )
                if v.get("question"):
                    lines.append(f"    → {v['question']}")
        sections.append("\n".join(lines))

    # --- Utilization ---
    util_email = util_emails_by_person.get(owner)
    if util_email:
        try:
            body_lines   = util_email["body"].split("\n")
            util_section = "\n".join(body_lines[2:-3]).strip()
            if util_section:
                sections.append(util_section)
        except Exception:
            pass

    if not sections:
        continue

    body = (
        f"Hi {first_name},\n\n"
        + "\n\n".join(sections)
        + f"\n\nBest,\n{sender_name}"
    )

    combined_emails.append({
        "to":      person_email,
        "subject": f"Scheduling Review — {active_month}",
        "person":  owner,
        "body":    body,
    })

    with st.expander(f"👁 {first_name} ({owner}) · {person_email}"):
        st.text(body)

# ---- Send ----
st.divider()

no_email = [e["person"] for e in combined_emails
            if not active_owners.get(e["person"],{}).get("email")]
if no_email:
    names = [active_owners[o].get("first_name",o) for o in no_email]
    st.warning(f"⚠️ No email for: {', '.join(names)} — will be skipped.")

sendable = [e for e in combined_emails if e.get("to")]
col_b1, col_b2 = st.columns(2)
with col_b1:
    send_sel = st.button(
        f"📤 Send to Selected ({len(sendable)})",
        type="primary", key="send_selected",
        disabled=not EMAIL_OK or not sendable)
with col_b2:
    sendable_all = [
        e for owner in all_owner_keys
        if active_owners.get(owner,{}).get("email")
        for e in [next((x for x in combined_emails if x["person"]==owner), None)]
        if e
    ]
    send_all_btn = st.button(
        f"📤 Send All ({len(sendable_all)})",
        key="send_all",
        disabled=not EMAIL_OK or not sendable_all)

if not EMAIL_OK:
    st.info("Configure SendGrid keys in Streamlit Secrets to enable sending.")

if send_sel or send_all_btn:
    targets = sendable if send_sel else sendable_all
    with st.spinner(f"Sending {len(targets)} email(s)…"):
        try:
            results = send_emails_batch(targets)
            sent   = [r for r in results if r.get("status") == "sent"]
            failed = [r for r in results if r.get("status") != "sent"]
            if sent:
                st.success(f"✅ {len(sent)} email(s) sent!")
            if failed:
                st.error(f"❌ {len(failed)} failed:")
                for r in failed:
                    st.write(f"  • {r.get('to','?')}: {r.get('status','?')}")
        except Exception as e:
            st.error(f"Error sending: {e}")
