from __future__ import annotations
# ============================================================
# app.py — GTM Scheduling Analyzer  |  streamlit run app.py
# ============================================================

import openpyxl
import streamlit as st
from collections import defaultdict
from datetime import datetime, date

from config import (SENDER_EMAIL, EMAIL_LOOKUP,
                    DEFAULT_BUDGET_THRESHOLD, DEFAULT_NEGATIVE_THRESHOLD,
                    DEADLINE_WARNING_DAYS)
from email_utils import (build_html_email, send_email, email_configured,
                         build_and_send_combined_emails)
from processors.budget_actual   import process_budget_actual
from processors.project_tracker import process_project_tracker
from processors.month_tab       import process_month_tab
from processors.lookup          import lookup_first_name, lookup_email
from processors.variance        import (parse_openair_report, read_schedule_hours,
                                        compute_variances, get_available_months,
                                        filter_by_months, _normalize_period)

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
# SIDEBAR
# ============================================================
with st.sidebar:
    st.header("⚙️ Settings")
    with st.form("settings_form"):
        st.subheader("📧 Email")
        sender_email = st.text_input(
            "Admin / CC Email",
            value=SENDER_EMAIL,
            help="CC'd on every outgoing email so the admin gets a copy of all messages.")
        st.form_submit_button("✔ Apply", use_container_width=True, key="apply_email")

        st.divider()
        st.subheader("💰 Budget to Actual")
        budget_threshold = st.number_input(
            "Flag unscheduled remaining over ($)",
            value=int(DEFAULT_BUDGET_THRESHOLD), step=1000, min_value=0)
        negative_threshold = st.number_input(
            "Flag negative budgets below -($)",
            value=int(DEFAULT_NEGATIVE_THRESHOLD), step=50, min_value=0,
            help="Flags anything more negative than this. Default $100 → below -$100.")
        st.form_submit_button("✔ Apply", use_container_width=True, key="apply_budget")

        st.divider()
        st.subheader("⏰ Hour Reminders")
        warn_days = st.number_input(
            "Warn X days before period deadline",
            value=DEADLINE_WARNING_DAYS, min_value=1, max_value=14)
        st.form_submit_button("✔ Apply", use_container_width=True, key="apply_hours")

        st.divider()
        st.subheader("📋 Email Lookup")
        st.caption(f"{len(EMAIL_LOOKUP)} people configured")
        show_lookup = st.checkbox("Show lookup table")

    if show_lookup:
        st.dataframe([{"Name": k, "First Name": v["first_name"], "Email": v["email"]}
                      for k, v in EMAIL_LOOKUP.items()],
                     use_container_width=True, hide_index=True)

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

# ============================================================
# LOAD + CACHE HEAVY PROCESSING
# ============================================================
@st.cache_data(show_spinner=False)
def load_workbook_cached(file_bytes: bytes):
    import io
    return openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)


@st.cache_data(show_spinner=False)
def get_valid_people(file_bytes: bytes) -> set:
    """Read valid staff names from row 2 of the current month tab."""
    import io
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    today    = date.today()
    abbr     = today.strftime("%b").lower()
    full     = today.strftime("%B").lower()
    ws       = None
    for name in wb.sheetnames:
        low = name.lower().strip()
        if low.startswith(abbr) or low.startswith(full):
            ws = wb[name]
            break
    if ws is None:
        return set()
    people = set()
    for col in range(7, 33):
        val = ws.cell(row=2, column=col).value
        if val:
            people.add(str(val).strip())
    return people


@st.cache_data(show_spinner=False)
def run_analysis(file_bytes: bytes, budget_thr: float, neg_thr: float,
                 warn_d: int):
    import io
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    sheets = wb.sheetnames

    def find_sheet(keywords):
        for name in sheets:
            if any(kw in name.lower() for kw in keywords):
                return name
        return None

    budget_sheet  = find_sheet(["budget to actual", "budget"])
    tracker_sheet = find_sheet(["project tracker", "tracker"])

    budget_issues = process_budget_actual(
        wb[budget_sheet], budget_thr, neg_thr) if budget_sheet else []
    tracker_issues, tbd = process_project_tracker(
        wb[tracker_sheet]) if tracker_sheet else ([], [])
    month_issues, active_month = process_month_tab(wb, warn_d)

    return (budget_issues, tracker_issues, tbd,
            month_issues, active_month, budget_sheet, tracker_sheet, sheets)


@st.cache_data(show_spinner=False)
def run_openair(oa_bytes: bytes, oa_name: str):
    import io
    return parse_openair_report(io.BytesIO(oa_bytes))


@st.cache_data(show_spinner=False)
def run_variance(sched_bytes: bytes, actual_data: dict,
                 selected_months: tuple, active_month: str):
    import io
    wb = openpyxl.load_workbook(io.BytesIO(sched_bytes), data_only=True)
    if not active_month or not selected_months:
        return []
    filtered  = filter_by_months(actual_data, list(selected_months))
    sched     = read_schedule_hours(wb, active_month)
    return compute_variances(filtered, sched, selected_periods=list(selected_months))


with st.status("🔄 Analyzing schedule file...", expanded=True) as status:
    st.write("📂 Loading workbook...")
    file_bytes = schedule_file.read()

    st.write("💰 Processing Budget to Actual...")
    st.write("📋 Processing Project Tracker...")
    st.write("📅 Processing Month Hours...")
    (budget_issues, tracker_issues, tbd_projects,
     month_issues, active_month,
     budget_sheet_name, tracker_sheet_name, sheets) = run_analysis(
        file_bytes,
        float(budget_threshold),
        float(negative_threshold),
        int(warn_days))

    st.write("👥 Loading valid staff list...")
    valid_people = get_valid_people(file_bytes)

    status.update(label="✅ Analysis complete!", state="complete", expanded=False)

st.success(f"✅ Loaded **{len(sheets)}** sheet(s)")
c1, c2, c3 = st.columns(3)
c1.info(f"💰 Budget: **{budget_sheet_name or 'Not found'}**")
c2.info(f"📋 Tracker: **{tracker_sheet_name or 'Not found'}**")
c3.info(f"📅 Month: **{active_month or datetime.now().strftime('%B')}**")
st.divider()

# ============================================================
# OPENAIR
# ============================================================
actual_data_full = {}
has_openair      = False
available_months = []
openair_error    = None

if openair_file:
    with st.status("🔄 Processing OpenAir report...", expanded=True) as oa_status:
        try:
            st.write("📊 Parsing time entries...")
            oa_bytes         = openair_file.read()
            actual_data_full = run_openair(oa_bytes, openair_file.name)
            st.write("📅 Identifying available periods...")
            available_months, future_months = get_available_months(actual_data_full)
            has_openair      = True
            oa_status.update(label="✅ OpenAir loaded!", state="complete", expanded=False)
        except Exception as e:
            openair_error = str(e)
            oa_status.update(label="❌ OpenAir error", state="error", expanded=True)

# ============================================================
# MONTH SELECTOR
# ============================================================
selected_months = []
variance_issues = []
future_months   = []

if has_openair and available_months:
    current_abbr   = datetime.now().strftime("%b")
    default_months = [m for m in available_months
                      if m.startswith(current_abbr) and m not in future_months]
    if not default_months:
        default_months = [m for m in available_months if m not in future_months][-1:]

    st.subheader("📅 Variance Period Selection")

    def _period_option_label(p):
        return f"🔮 {p} (future — no actuals yet)" if p in future_months else p

    selected_months = st.multiselect(
        "Select period(s) for variance analysis:",
        options=available_months,
        default=default_months,
        format_func=_period_option_label,
        help="🔮 = future period (no actuals yet). Default is current month.")

    selected_future = [m for m in selected_months if m in future_months]
    selected_past   = [m for m in selected_months if m not in future_months]

    if selected_future and selected_past:
        st.info(f"📢 **Mixed selection:** {', '.join(selected_past)} have actuals · "
                f"{', '.join(selected_future)} are future (scheduled hours only)")
    elif selected_future:
        st.warning(f"🔮 **Future periods selected:** {', '.join(selected_future)} — "
                   f"no actual hours exist yet, only scheduled hours will show.")
    elif len(selected_months) > 1:
        st.info(f"📢 **Multi-period mode:** {', '.join(selected_months)}")

    if selected_months and active_month:
        with st.status("🔄 Computing variances...", expanded=True) as var_status:
            st.write(f"📊 Comparing actuals vs schedule for "
                     f"{', '.join(selected_months)}...")
            variance_issues = run_variance(
                file_bytes, actual_data_full,
                tuple(selected_months), active_month)
            st.write(f"✅ Found {len(variance_issues)} variance(s) over 3 hours")
            var_status.update(
                label=f"✅ Variance complete — {len(variance_issues)} item(s) flagged",
                state="complete", expanded=False)

    st.divider()

# ============================================================
# BUILD OWNER DATA — filtered to valid schedule people only
# ============================================================
owners_data = defaultdict(lambda: {
    "email": None, "first_name": "there",
    "tracker": [], "budget": [], "month": [], "variance": []
})

for issue in tracker_issues:
    owner = issue["owner"]
    if owner not in valid_people:
        continue
    if not owners_data[owner]["email"]:
        owners_data[owner]["email"]      = issue.get("owner_email")
        owners_data[owner]["first_name"] = issue.get("owner_first", owner)
    owners_data[owner]["tracker"].append(issue)

for issue in budget_issues:
    owner = issue["owner"]
    if owner not in valid_people:
        continue
    if not owners_data[owner]["email"]:
        owners_data[owner]["email"]      = issue.get("owner_email")
        owners_data[owner]["first_name"] = issue.get("owner_first", owner)
    owners_data[owner]["budget"].append(issue)

for issue in month_issues:
    person = issue["person"]
    if person not in valid_people:
        continue
    if not owners_data[person]["email"]:
        owners_data[person]["email"]      = issue.get("person_email")
        owners_data[person]["first_name"] = issue.get("person_first", person)
    owners_data[person]["month"].append(issue)

for v in variance_issues:
    person = v["person"]
    if person not in valid_people:
        continue
    if owners_data[person]["first_name"] in ("there", ""):
        owners_data[person]["first_name"] = v.get("first_name", person)
    if not owners_data[person]["email"]:
        owners_data[person]["email"] = lookup_email(person)
    owners_data[person]["variance"].append(v)

active_owners = {
    owner: data for owner, data in owners_data.items()
    if (data["tracker"] or data["budget"] or
        data["month"]   or data["variance"])
    and owner in valid_people
}

# ============================================================
# ANALYSIS TABS
# ============================================================
tab1, tab2, tab3, tab4 = st.tabs([
    "📋 Project Tracker", "💰 Budget to Actual",
    "📅 Month Hours",     "📊 Variance (OpenAir)"])

with tab1:
    st.header("Project Tracker — Known Projects")
    if not tracker_sheet_name:
        st.error("No Project Tracker tab found.")
    else:
        c1, c2 = st.columns(2)
        c1.metric("⚠️ Issues", len(tracker_issues))
        c2.metric("📌 TBD (excluded)", len(tbd_projects))
        if tbd_projects:
            with st.expander(f"📌 {len(tbd_projects)} TBD projects"):
                st.dataframe([{"Client": p["client"],
                               "Project Code": p["project_code"],
                               "Owner": p["owner"],
                               "Budget": f"${p['budget']:,.0f}"}
                              for p in tbd_projects],
                             use_container_width=True, hide_index=True)
        if not tracker_issues:
            st.success("✅ No issues found!")
        else:
            st.dataframe([{"Client": i["client"],
                           "Project Code": i["project_code"],
                           "Owner": i["owner"], "Issue": prob,
                           "Has Email": "✅" if i.get("owner_email") else "❌"}
                          for i in tracker_issues for prob in i["problems"]],
                         use_container_width=True, hide_index=True)

with tab2:
    st.header("Budget to Actual — Known Projects")
    st.caption(f"Flagging: negative < -${negative_threshold:,.0f}  "
               f"| unscheduled > ${budget_threshold:,.0f}")
    if not budget_sheet_name:
        st.error("No Budget to Actual tab found.")
    else:
        neg = [i for i in budget_issues if i["type"] == "negative"]
        np  = [i for i in budget_issues if i["type"] == "not_projected"]
        c1, c2, c3 = st.columns(3)
        c1.metric("🔴 Over Budget", len(neg))
        c2.metric("🟡 Unscheduled", len(np))
        c3.metric("⚠️ Total", len(budget_issues))
        if budget_issues:
            st.dataframe([{"Client": i["client"],
                           "Project Code": i["project_code"],
                           "Owner": i["owner"],
                           "Budget": f"${i['budget']:,.0f}",
                           "Remaining": f"${i['remaining']:,.0f}",
                           "Flag": i["description"],
                           "Has Email": "✅" if i.get("owner_email") else "❌"}
                          for i in budget_issues],
                         use_container_width=True, hide_index=True)

with tab3:
    st.header("Month Hours — Deadline Reminders")
    if not active_month:
        st.error(f"No sheet found for {datetime.now().strftime('%B')}.")
    elif not month_issues:
        st.success(f"✅ No unconfirmed hours within {warn_days} day(s) of deadline.")
    else:
        st.info(f"Sheet: **{active_month}** · {warn_days}-day warning")
        st.metric("⏰ Unconfirmed near deadline", len(month_issues))
        st.dataframe([{"Client": i["client"],
                       "Project Code": i["project_code"],
                       "Person": i["person_first"],
                       "Period": i["period"],
                       "Deadline": i["deadline"].strftime("%b %d"),
                       "Days Left": i["days_left"],
                       "Hours": i["hours"],
                       "Has Email": "✅" if i.get("person_email") else "❌"}
                      for i in month_issues],
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
        st.success("✅ No variances over 3 hours for the selected period(s).")
    else:
        if len(selected_months) > 1:
            st.warning(f"📢 Showing variances across "
                       f"**{len(selected_months)} periods**: "
                       f"{', '.join(selected_months)}")
        st.metric("⚠️ Variances > 3 hrs", len(variance_issues))
        st.dataframe([{"Person": v["first_name"],
                       "Project": v["project_code"],
                       "Period": v["period"],
                       "Actual Hrs": v["actual_hours"],
                       "Sched Hrs": v["sched_hours"],
                       "Diff": v["difference"],
                       "To Review": v["question"]}
                      for v in variance_issues],
                     use_container_width=True, hide_index=True)

# ============================================================
# COMBINED EMAIL SECTION
# ============================================================
st.divider()
st.header("📧 Combined Emails by Owner")
st.caption("One email per person · Project Tracker · Budget · "
           "Scheduled Hours · Variance")

if not active_owners:
    st.info("No flagged items found — no emails to send.")
    st.stop()

st.metric("People with flagged items", len(active_owners))

# ---- Session state for checkbox selections ----
all_owner_keys = list(active_owners.keys())

if "selected_owners" not in st.session_state:
    st.session_state.selected_owners = set(all_owner_keys)

# Keep selected_owners in sync with current active_owners
# (removes any stale keys from previous runs)
st.session_state.selected_owners &= set(all_owner_keys)

col_sa, col_da, _ = st.columns([0.15, 0.18, 0.67])
with col_sa:
    if st.button("✅ Select All"):
        st.session_state.selected_owners = set(all_owner_keys)
        st.rerun()
with col_da:
    if st.button("⬜ Deselect All"):
        st.session_state.selected_owners = set()
        st.rerun()

# ---- Recipient list with checkboxes inside a form ----
# Using a form prevents individual checkbox clicks from rerunning the app
st.markdown("**Select recipients:**")
with st.form("recipient_form"):
    form_checks = {}
    for owner, data in active_owners.items():
        owner_email = data.get("email")
        first_name  = data.get("first_name", owner)
        email_str   = owner_email or "⚠️ no email"
        label = (f"**{first_name} ({owner})** · {email_str} — "
                 f"Tracker: {len(data['tracker'])} · "
                 f"Budget: {len(data['budget'])} · "
                 f"Hours: {len(data['month'])} · "
                 f"Variance: {len(data['variance'])}")
        form_checks[owner] = st.checkbox(
            label,
            value=owner in st.session_state.selected_owners,
            key=f"chk_{owner}")

    confirmed = st.form_submit_button("✔ Confirm Selection")
    if confirmed:
        st.session_state.selected_owners = {
            o for o, v in form_checks.items() if v}
        st.success(f"Selection updated: "
                   f"{len(st.session_state.selected_owners)} people selected.")

selected_owners = list(st.session_state.selected_owners)

# ---- Email Previews ----
st.markdown("**Email previews:**")
for owner in selected_owners:
    if owner not in active_owners:
        continue
    data       = active_owners[owner]
    first_name = data.get("first_name", owner)
    with st.expander(f"👁 Preview — {first_name} ({owner})"):
        owner_email = data.get("email")
        if not owner_email:
            st.error("No email address — add to config.py")
        else:
            html = build_html_email(
                owner, first_name,
                data["tracker"], data["budget"],
                data["month"],   data["variance"],
                has_openair=has_openair,
                selected_months=(selected_months
                                 if len(selected_months) > 1 else None),
            )
            if html:
                st.components.v1.html(html, height=500, scrolling=True)

st.divider()

no_email_selected = [o for o in selected_owners
                     if not active_owners.get(o, {}).get("email")]
if no_email_selected:
    names = [active_owners[o]["first_name"] for o in no_email_selected]
    st.warning(f"⚠️ No email configured for: {', '.join(names)} — will be skipped.")

sendable_selected = [o for o in selected_owners
                     if active_owners.get(o, {}).get("email")]
sendable_all      = [o for o in all_owner_keys
                     if active_owners[o].get("email")]

col_b1, col_b2 = st.columns(2)
with col_b1:
    send_sel = st.button(
        f"📤 Send to Selected ({len(sendable_selected)})",
        type="primary", key="send_selected",
        disabled=not email_configured() or not sendable_selected)
with col_b2:
    send_all = st.button(
        f"📤 Send All ({len(sendable_all)})",
        key="send_all",
        disabled=not email_configured() or not sendable_all)

if not email_configured():
    st.info("Configure SendGrid credentials in Streamlit Secrets to enable sending.")


def _do_send(owners_to_send: list):
    with st.spinner(f"Sending {len(owners_to_send)} email(s)…"):
        results = build_and_send_combined_emails(
            active_owners,
            cc_email=sender_email,
            has_openair=has_openair,
            selected_months=selected_months if selected_months else None,
            selected_owners=owners_to_send,
        )
    sent = [r for r in results if r.get("status") == "sent"]
    fail = [r for r in results if r.get("status") != "sent"]
    if sent:
        st.success(f"✅ {len(sent)} email(s) sent!")
    if fail:
        st.error(f"❌ {len(fail)} failed:")
        for r in fail:
            st.write(f"  • **{r['to']}**: {r['status']}")


if send_sel:
    _do_send(sendable_selected)

if send_all:
    _do_send(sendable_all)
