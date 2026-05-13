from __future__ import annotations
# ============================================================
# app.py — GTM Scheduling Analyzer  |  streamlit run app.py
# ============================================================

import gc
import hashlib
import io
import warnings

import openpyxl
import streamlit as st
from collections import defaultdict
from datetime import datetime

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

# ============================================================
# PAGE CONFIG
# ============================================================
st.set_page_config(page_title="GTM Scheduling Analyzer",
                   layout="wide", page_icon="📊")
st.title("📊 GTM Scheduling Analyzer")
st.caption(f"Today: {datetime.now().strftime('%A, %B %d, %Y')}")

if not email_configured():
    st.warning("⚠️ **Email credentials not configured.** "
               "Add SendGrid keys in Streamlit Secrets to enable sending.")

# ============================================================
# SESSION STATE — initialize before sidebar
# ============================================================
if "settings" not in st.session_state:
    st.session_state.settings = {
        "budget_threshold":   float(DEFAULT_BUDGET_THRESHOLD),
        "negative_threshold": float(DEFAULT_NEGATIVE_THRESHOLD),
        "variance_min":       float(DEFAULT_VARIANCE_MIN),
        "variance_max":       float(DEFAULT_VARIANCE_MAX),
    }

if "selected_owners" not in st.session_state:
    st.session_state.selected_owners = set()

# ============================================================
# SIDEBAR
# Using st.form so that +/- clicks on number inputs do NOT
# trigger a rerun — only clicking "✔ Apply Settings" does.
# ============================================================
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

        applied = st.form_submit_button(
            "✔ Apply Settings", type="primary", use_container_width=True
        )

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

# Pull the applied values — these never change mid-run
budget_threshold   = st.session_state.settings["budget_threshold"]
negative_threshold = st.session_state.settings["negative_threshold"]
variance_min       = st.session_state.settings["variance_min"]
variance_max       = st.session_state.settings["variance_max"]

try:
    from_email  = st.secrets["email"]["from_email"]
except (KeyError, FileNotFoundError):
    from_email  = SENDER_EMAIL
sender_name = SENDER_NAMES.get(from_email, "Jake")
active_month = datetime.now().strftime("%B")

# ============================================================
# FAST HASHING + CACHING
#
# file_hash (MD5 string, 32 chars) is the cache key — fast to
# compare.  _file_bytes has _ prefix so Streamlit skips hashing
# the large bytes argument.
# cache_resource stores the workbook object without pickling it,
# avoiding the serialization crashes from cache_data.
# ============================================================
def _hash(b: bytes) -> str:
    return hashlib.md5(b).hexdigest()


@st.cache_resource(max_entries=2, show_spinner=False)
def _load_wb(file_hash: str, _file_bytes: bytes):
    """Load workbook once per unique file, keep in memory."""
    return openpyxl.load_workbook(io.BytesIO(_file_bytes), data_only=True)


@st.cache_data(show_spinner=False)
def run_analysis(file_hash: str, _file_bytes: bytes,
                 budget_thr: float, proj_pct: float, neg_thr: float):
    wb     = _load_wb(file_hash, _file_bytes)
    sheets = list(wb.sheetnames)

    def find_sheet(keywords):
        for name in sheets:
            if any(kw in name.lower() for kw in keywords):
                return name
        return None

    budget_sheet  = find_sheet(["budget to actual", "budget"])
    tracker_sheet = find_sheet(["project tracker", "tracker"])

    budget_issues = []
    if budget_sheet:
        try:
            budget_issues = process_budget_actual(
                wb[budget_sheet], budget_thr, proj_pct, neg_thr)
        except Exception:
            pass

    tracker_issues, tbd = [], []
    if tracker_sheet:
        try:
            tracker_issues, tbd = process_project_tracker(wb[tracker_sheet])
        except Exception:
            pass

    util_data = []
    try:
        util_data = process_utilization(wb, target_month=active_month)
    except Exception:
        pass

    valid_people = set()
    try:
        month_prefix = active_month[:3].lower()
        for name in sheets:
            if name.lower().startswith(month_prefix):
                ws = wb[name]
                for col in range(7, 45):
                    val = ws.cell(row=2, column=col).value
                    if val:
                        valid_people.add(str(val).strip())
                break
    except Exception:
        pass

    gc.collect()
    return (budget_issues, tracker_issues, tbd,
            util_data, valid_people,
            budget_sheet, tracker_sheet, sheets)


@st.cache_data(show_spinner=False)
def run_openair(oa_hash: str, _oa_bytes: bytes):
    result = parse_openair_report(io.BytesIO(_oa_bytes))
    gc.collect()
    return result


@st.cache_data(show_spinner=False)
def run_variance(file_hash: str, _file_bytes: bytes,
                 oa_hash: str, _oa_bytes: bytes,
                 selected_months: tuple,
                 var_min: float, var_max: float):
    wb = _load_wb(file_hash, _file_bytes)
    month_prefix = active_month[:3].lower()
    sched_sheet  = next(
        (s for s in wb.sheetnames if s.lower().startswith(month_prefix)), None
    )
    if not sched_sheet:
        return []
    try:
        actual_data = parse_openair_report(io.BytesIO(_oa_bytes))
        filtered    = filter_by_months(actual_data, list(selected_months))
        sched       = read_schedule_hours(wb, sched_sheet)
        result      = compute_variances(filtered, sched,
                                        min_diff=var_min, max_diff=var_max)
        gc.collect()
        return result
    except Exception:
        gc.collect()
        return []


# ============================================================
# FILE UPLOADS
# ============================================================
st.subheader("📁 Upload Files")
col_f1, col_f2 = st.columns(2)
with col_f1:
    schedule_file = st.file_uploader(
        "Schedule File (.xlsx or .csv)", type=["xlsx", "csv"], key="schedule")
with col_f2:
    openair_file = st.file_uploader(
        "OpenAir Report (.xlsx or .csv) — optional",
        type=["xlsx", "csv"], key="openair")

if not schedule_file:
    st.info("Upload the scheduling file above to begin.")
    st.stop()

file_bytes = bytes(schedule_file.read())
if not file_bytes:
    st.error("Schedule file appears to be empty.")
    st.stop()
file_hash = _hash(file_bytes)

# ============================================================
# LOAD + PROCESS
# ============================================================
with st.status("🔄 Analyzing schedule file...", expanded=True) as status:
    st.write("📂 Loading workbook...")
    st.write("💰 Processing Budget to Actual...")
    st.write("📋 Processing Project Tracker...")
    st.write("📈 Processing Utilization...")
    (budget_issues, tracker_issues, tbd_projects,
     util_data, valid_people,
     budget_sheet_name, tracker_sheet_name, sheets) = run_analysis(
        file_hash, file_bytes,
        budget_threshold, DEFAULT_PROJECTION_THRESHOLD_PCT, negative_threshold)
    status.update(label="✅ Analysis complete!", state="complete", expanded=False)

st.success(f"✅ Loaded **{len(sheets)}** sheet(s)")
c1, c2, c3 = st.columns(3)
c1.info(f"💰 Budget: **{budget_sheet_name or 'Not found'}**")
c2.info(f"📋 Tracker: **{tracker_sheet_name or 'Not found'}**")
c3.info(f"📅 Month: **{active_month}**")
st.divider()

# ============================================================
# OPENAIR
# ============================================================
has_openair      = False
available_months = []
future_months    = []
openair_error    = None
oa_bytes         = b""
oa_hash          = ""

if openair_file:
    with st.status("🔄 Processing OpenAir report...", expanded=True) as oa_status:
        try:
            st.write("📊 Parsing time entries...")
            oa_bytes         = bytes(openair_file.read())
            oa_hash          = _hash(oa_bytes)
            actual_data_full = run_openair(oa_hash, oa_bytes)
            st.write("📅 Identifying available periods...")
            available_months, future_months = get_available_months(actual_data_full)
            has_openair = True
            oa_status.update(label="✅ OpenAir loaded!", state="complete",
                             expanded=False)
        except Exception as e:
            openair_error = str(e)
            oa_status.update(label="❌ OpenAir error", state="error",
                             expanded=True)

# ============================================================
# PERIOD SELECTOR + VARIANCE
# ============================================================
selected_months = []
variance_issues = []

if has_openair and available_months:
    current_abbr   = datetime.now().strftime("%b")
    default_months = [m for m in available_months
                      if m.startswith(current_abbr) and m not in future_months]
    if not default_months:
        default_months = [m for m in available_months
                          if m not in future_months][-1:]

    st.subheader("📅 Variance Period Selection")

    def _label(p):
        return f"🔮 {p} (future)" if p in future_months else p

    selected_months = st.multiselect(
        "Select period(s) for variance analysis:",
        options=available_months,
        default=default_months,
        format_func=_label,
        help="🔮 = future period. Default is current month.")

    if selected_months:
        with st.status("🔄 Computing variances...", expanded=True) as var_status:
            st.write(f"📊 Comparing actuals vs schedule for "
                     f"{', '.join(selected_months)}...")
            variance_issues = run_variance(
                file_hash, file_bytes,
                oa_hash, oa_bytes,
                tuple(selected_months),
                variance_min, variance_max)
            st.write(f"✅ Found {len(variance_issues)} variance(s)")
            var_status.update(
                label=f"✅ Variance complete — {len(variance_issues)} flagged",
                state="complete", expanded=False)

# ============================================================
# BUILD OWNER MAP
# ============================================================
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

# Utilization entries go into util list per person
try:
    util_by_person = {u["person"]: u for u in util_data if u.get("person")}
except Exception:
    util_by_person = {}

for person, u in util_by_person.items():
    if valid_people and person not in valid_people:
        continue
    if not owners_data[person]["email"]:
        owners_data[person]["email"] = u.get("person_email") or _lookup_email(person)
    owners_data[person]["util"].append(u)

active_owners = {
    owner: data for owner, data in owners_data.items()
    if (data["tracker"] or data["budget"] or data["variance"] or data["util"])
    and (not valid_people or owner in valid_people)
}

# ============================================================
# TABS
# ============================================================
tab1, tab2, tab3, tab4 = st.tabs([
    "📋 Project Tracker",
    "💰 Budget to Actual",
    "📈 Utilization",
    "📊 Variance (OpenAir)",
])

with tab1:
    st.header("Project Tracker — Known Projects")
    if not tracker_sheet_name:
        st.error("No Project Tracker tab found.")
    else:
        c1, c2 = st.columns(2)
        c1.metric("⚠️ Issues",          len(tracker_issues))
        c2.metric("📌 TBD / Pending SOW", len(tbd_projects))
        if tbd_projects:
            with st.expander(f"📌 {len(tbd_projects)} TBD / Pending SOW projects"):
                st.dataframe(
                    [{
                        "Client":       p.get("client", ""),
                        "Project Code": p.get("project_code", ""),
                        "Status":       p.get("status", "TBD"),
                        "Owner":        p.get("owner", ""),
                        "Budget":       f"${p.get('budget', 0):,.0f}",
                    } for p in tbd_projects],
                    use_container_width=True, hide_index=True)
        if not tracker_issues:
            st.success("✅ No issues found!")
        else:
            rows = []
            for i in tracker_issues:
                problems = i.get("problems", i.get("missing_rates", []))
                for prob in (problems if problems else [""]):
                    rows.append({
                        "Client":       i.get("client", ""),
                        "Project Code": i.get("project_code", ""),
                        "Owner":        i.get("owner", ""),
                        "Issue":        prob,
                        "Has Email":    "✅" if i.get("owner_email") else "❌",
                    })
            st.dataframe(rows, use_container_width=True, hide_index=True)

with tab2:
    st.header("Budget to Actual — Known Projects")
    st.caption(f"Flagging: negative < -${negative_threshold:,.0f}  "
               f"| unscheduled > ${budget_threshold:,.0f}")
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
                [{
                    "Client":       i.get("client", ""),
                    "Project Code": i.get("project_code", ""),
                    "Owner":        i.get("owner", ""),
                    "Budget":       f"${i.get('budget', 0):,.0f}",
                    "Remaining":    f"${i.get('remaining', 0):,.0f}",
                    "Flag":         i.get("description", ""),
                    "Has Email":    "✅" if i.get("owner_email") else "❌",
                } for i in budget_issues],
                use_container_width=True, hide_index=True)

with tab3:
    st.header(f"Utilization — {active_month}")
    if not util_data:
        st.warning("No utilization data found. Make sure the workbook has a "
                   "'Utilization by Month' tab.")
    else:
        st.dataframe(
            [{
                "Role":        u.get("role", ""),
                "Person":      u.get("person", ""),
                "Chargeable":  u.get("chargeable", "-"),
                "Holiday":     u.get("holiday", "-"),
                "PTO":         u.get("pto", "-"),
                "Month Total": u.get("month_total", "-"),
                "Remaining":   u.get("remaining", "-"),
                "Utilization": f"{u['utilization_pct']:.1f}%" if u.get("utilization_pct") is not None else "-",
                "Goal":        f"{u['goal_pct']:.0f}%"        if u.get("goal_pct")        is not None else "-",
                "Difference":  f"{u['difference_pct']:+.1f}%" if u.get("difference_pct") is not None else "-",
                "Has Email":   "✅" if u.get("person_email") else "❌",
            } for u in util_data],
            use_container_width=True, hide_index=True)

with tab4:
    st.header("Actual vs Schedule Variance (OpenAir)")
    if not openair_file:
        st.info("Upload an OpenAir file above to enable variance analysis.")
    elif openair_error:
        st.error(f"Error parsing OpenAir file: {openair_error}")
    elif not selected_months:
        st.warning("Select at least one period above.")
    elif not variance_issues:
        st.success(f"✅ No variances outside "
                   f"[{variance_min:+.0f}, {variance_max:+.0f}] hrs "
                   f"for the selected period(s).")
    else:
        if len(selected_months) > 1:
            st.info(f"📢 Showing variances across "
                    f"**{len(selected_months)} periods**: "
                    f"{', '.join(selected_months)}")
        st.metric("⚠️ Variances Found", len(variance_issues))
        st.dataframe(
            [{
                "Person":       v.get("person", ""),
                "Project":      v.get("project_code", ""),
                "Period":       v.get("period", ""),
                "Actual Hrs":   v.get("actual_hours", 0),
                "Sched Hrs":    v.get("sched_hours", 0),
                "Diff":         v.get("difference", 0),
                "To Review":    v.get("question", ""),
            } for v in variance_issues],
            use_container_width=True, hide_index=True)

# ============================================================
# COMBINED EMAILS
# ============================================================
st.divider()
st.header("📧 Combined Emails")
st.caption("One email per person · Project Tracker · Budget · "
           "Utilization · Variance")

if not active_owners:
    st.info("No flagged items found — no emails to send.")
    st.stop()

st.metric("People with flagged items", len(active_owners))
all_owner_keys = list(active_owners.keys())

# Keep selection in sync if owners changed
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

# Recipient checkboxes (outside a form so they toggle instantly)
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

selected_owners = list(st.session_state.selected_owners)

# ---- Build emails ----
try:
    util_emails_by_person = {
        e["person"]: e
        for e in build_utilization_emails(
            util_data, month=active_month, sender_name=sender_name)
    }
except Exception:
    util_emails_by_person = {}

combined_emails = []
for owner in sorted(selected_owners):
    if owner not in active_owners:
        continue
    data         = active_owners[owner]
    person_email = data.get("email") or _lookup_email(owner)
    first_name   = data.get("first_name") or owner
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

    if tracker_list:
        lines = ["The following projects assigned to you are missing billing "
                 "rates. Please review and update as soon as possible.\n"]
        for issue in tracker_list:
            problems = issue.get("problems", issue.get("missing_rates", []))
            lines.append(
                f"  - {issue.get('client', '')} | "
                f"{issue.get('project_code', '')}\n"
                f"    Missing: {', '.join(problems)}"
            )
        sections.append("\n".join(lines))

    if budget_list:
        lines = ["Please review the following budget items:\n"]
        for issue in budget_list:
            lines.append(
                f"  - {issue.get('client', '')} | "
                f"{issue.get('project_code', '')}: "
                f"${issue.get('remaining', 0):,.0f} remaining "
                f"({issue.get('description', '')})"
            )
        sections.append("\n".join(lines))

    if owner_tbd:
        lines = ["The following projects have TBD or Pending SOW budgets. "
                 "If you have any updates please reply — otherwise no action needed.\n"]
        for proj in owner_tbd:
            lines.append(
                f"  - {proj.get('client', '')} | "
                f"{proj.get('project_code', '')} "
                f"[{proj.get('status', 'TBD')}]"
            )
        sections.append("\n".join(lines))

    if variance_list:
        lines = [
            "Please see your variance for the current period:\n"
            if is_intern else
            "Please see the variance summary for your projects:\n"
        ]
        for v in variance_list:
            prefix = "" if is_intern else f"{v.get('person', '')} | "
            lines.append(
                f"  - {prefix}{v.get('project_code', '')} | "
                f"{v.get('period', '')} | "
                f"Actual: {v.get('actual_hours', 0)}h  "
                f"Scheduled: {v.get('sched_hours', 0)}h  "
                f"Diff: {v.get('difference', 0):+.1f}h"
            )
            if v.get("question"):
                lines.append(f"    {v['question']}")
        sections.append("\n".join(lines))

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

# ---- Previews ----
st.markdown("**Email previews:**")
for email in combined_emails:
    with st.expander(f"👁 Preview — {email['person']} · {email['to']}"):
        st.text(email["body"])

st.divider()

no_email = [e["person"] for e in combined_emails
            if not active_owners.get(e["person"], {}).get("email")]
if no_email:
    st.warning(f"⚠️ No email configured for: {', '.join(no_email)} — will be skipped.")

sendable = [e for e in combined_emails if e.get("to")]

col_b1, col_b2 = st.columns(2)
with col_b1:
    send_sel = st.button(
        f"📤 Send to Selected ({len(sendable)})",
        type="primary", key="send_selected",
        disabled=not EMAIL_OK or not sendable)
with col_b2:
    send_all_btn = st.button(
        f"📤 Send All ({len(sendable)})",
        key="send_all",
        disabled=not EMAIL_OK or not sendable)

if not EMAIL_OK:
    st.info("Configure SendGrid credentials in Streamlit Secrets to enable sending.")

if send_sel or send_all_btn:
    with st.spinner(f"Sending {len(sendable)} email(s)…"):
        try:
            results = send_emails_batch(sendable)
            sent    = [r for r in results if r.get("status") == "sent"]
            failed  = [r for r in results if r.get("status") != "sent"]
            if sent:
                st.success(f"✅ {len(sent)} email(s) sent!")
            if failed:
                st.error(f"❌ {len(failed)} failed:")
                for r in failed:
                    st.write(f"  • **{r.get('to', '?')}**: {r.get('status', '?')}")
        except Exception as e:
            st.error(f"Error sending emails: {e}")
