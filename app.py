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
from processors.lookup          import lookup_first_name
from processors.variance        import (parse_openair_report, read_schedule_hours,
                                        compute_variances, get_available_months,
                                        filter_by_months)

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
    sender_email = st.text_input("From / CC Email", value=SENDER_EMAIL)

    st.divider()
    st.subheader("💰 Budget to Actual")
    budget_threshold = st.number_input(
        "Flag unscheduled remaining over ($)",
        value=int(DEFAULT_BUDGET_THRESHOLD), step=1000, min_value=0,
        help="Flag projects where remaining budget exceeds this amount (unscheduled).")

    negative_threshold = st.number_input(
        "Flag negative budgets below -($)",
        value=int(DEFAULT_NEGATIVE_THRESHOLD), step=50, min_value=0,
        help="Flag projects where remaining is more negative than this. "
             "Default $100 means flag anything below -$100.")

    st.divider()
    st.subheader("⏰ Hour Reminders")
    warn_days = st.number_input(
        "Warn X days before period deadline",
        value=DEADLINE_WARNING_DAYS, min_value=1, max_value=14)

    st.divider()
    st.subheader("📋 Email Lookup")
    st.caption(f"{len(EMAIL_LOOKUP)} people configured")
    if st.checkbox("Show lookup table"):
        st.dataframe([{"Name": k,
                       "First Name": v["first_name"],
                       "Email": v["email"]}
                      for k, v in EMAIL_LOOKUP.items()],
                     use_container_width=True, hide_index=True)

# ============================================================
# FILE UPLOADS
# ============================================================
st.subheader("📁 Upload Files")
col_f1, col_f2 = st.columns(2)
with col_f1:
    schedule_file = st.file_uploader(
        "Schedule File (.xlsx)", type=["xlsx"], key="schedule")
with col_f2:
    openair_file = st.file_uploader(
        "OpenAir Report (.csv) — optional, enables variance analysis",
        type=["csv"], key="openair")

if not schedule_file:
    st.info("Upload the scheduling file above to begin.")
    st.stop()

with st.spinner("Loading workbook…"):
    wb = openpyxl.load_workbook(schedule_file, data_only=True)

sheets = wb.sheetnames

def find_sheet(keywords):
    for name in sheets:
        if any(kw in name.lower() for kw in keywords):
            return name
    return None

budget_sheet_name  = find_sheet(["budget to actual", "budget"])
tracker_sheet_name = find_sheet(["project tracker", "tracker"])

st.success(f"✅ Loaded **{len(sheets)}** sheet(s)")
c1, c2, c3 = st.columns(3)
c1.info(f"💰 Budget: **{budget_sheet_name or 'Not found'}**")
c2.info(f"📋 Tracker: **{tracker_sheet_name or 'Not found'}**")
c3.info(f"📅 Month: **{datetime.now().strftime('%B')}** (auto)")
st.divider()

# ============================================================
# PROCESS DATA
# ============================================================
with st.spinner("Analyzing…"):
    budget_issues = []
    if budget_sheet_name:
        budget_issues = process_budget_actual(
            wb[budget_sheet_name],
            unscheduled_threshold=float(budget_threshold),
            negative_threshold=float(negative_threshold))

    tracker_issues, tbd_projects = [], []
    if tracker_sheet_name:
        tracker_issues, tbd_projects = process_project_tracker(wb[tracker_sheet_name])

    month_issues, active_month = process_month_tab(wb, int(warn_days))

    # OpenAir
    actual_data_full = {}
    variance_issues  = []
    has_openair      = False
    openair_error    = None
    available_months = []

    if openair_file:
        try:
            actual_data_full = parse_openair_report(openair_file)
            available_months = get_available_months(actual_data_full)
            has_openair      = True
        except Exception as e:
            openair_error = str(e)

# ============================================================
# MONTH SELECTOR (shown only when OpenAir loaded)
# ============================================================
selected_months = []
if has_openair and available_months:
    current_month_abbr = datetime.now().strftime("%b")
    default_months     = [m for m in available_months
                          if m.startswith(current_month_abbr)]
    if not default_months and available_months:
        default_months = [available_months[-1]]

    st.subheader("📅 Variance Month Selection")
    selected_months = st.multiselect(
        "Select period(s) to include in variance analysis and emails:",
        options=available_months,
        default=default_months,
        help="Default is current month only. Select multiple to include more periods — "
             "this will be called out clearly in all reports and emails."
    )

    if len(selected_months) > 1:
        st.info(f"📢 **Multi-period mode:** Reports and emails will reference "
                f"**{', '.join(selected_months)}** and recipients will be "
                f"notified that multiple periods are included.")

    if selected_months:
        filtered_actual = filter_by_months(actual_data_full, selected_months)
        if active_month:
            sched_data      = read_schedule_hours(wb, active_month)
            variance_issues = compute_variances(filtered_actual, sched_data)
    st.divider()

# ============================================================
# BUILD OWNER-KEYED DATA
# ============================================================
owners_data = defaultdict(lambda: {
    "email": None, "first_name": "there",
    "tracker": [], "budget": [], "month": [], "variance": []
})

for issue in tracker_issues:
    owner = issue["owner"]
    if not owners_data[owner]["email"]:
        owners_data[owner]["email"]      = issue.get("owner_email")
        owners_data[owner]["first_name"] = issue.get("owner_first", owner)
    owners_data[owner]["tracker"].append(issue)

for issue in budget_issues:
    owner = issue["owner"]
    if not owners_data[owner]["email"]:
        owners_data[owner]["email"]      = issue.get("owner_email")
        owners_data[owner]["first_name"] = issue.get("owner_first", owner)
    owners_data[owner]["budget"].append(issue)

for issue in month_issues:
    person = issue["person"]
    if not owners_data[person]["email"]:
        owners_data[person]["email"]      = issue.get("person_email")
        owners_data[person]["first_name"] = issue.get("person_first", person)
    owners_data[person]["month"].append(issue)

for v in variance_issues:
    person = v["person"]
    if not owners_data[person]["first_name"] or owners_data[person]["first_name"] == "there":
        owners_data[person]["first_name"] = v.get("first_name", person)
    owners_data[person]["variance"].append(v)

# ============================================================
# ANALYSIS TABS
# ============================================================
tab1, tab2, tab3, tab4 = st.tabs([
    "📋 Project Tracker", "💰 Budget to Actual",
    "📅 Month Hours",     "📊 Variance (OpenAir)"])

# ---- TAB 1: PROJECT TRACKER ----
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
                st.dataframe([{"Client": p["client"], "Project Code": p["project_code"],
                               "Owner": p["owner"], "Budget": f"${p['budget']:,.0f}"}
                              for p in tbd_projects],
                             use_container_width=True, hide_index=True)
        if not tracker_issues:
            st.success("✅ No issues found!")
        else:
            st.dataframe([{"Client": i["client"], "Project Code": i["project_code"],
                           "Owner": i["owner"], "Issue": prob,
                           "Has Email": "✅" if i.get("owner_email") else "❌"}
                          for i in tracker_issues for prob in i["problems"]],
                         use_container_width=True, hide_index=True)

# ---- TAB 2: BUDGET TO ACTUAL ----
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
            st.dataframe([{"Client": i["client"], "Project Code": i["project_code"],
                           "Owner": i["owner"], "Budget": f"${i['budget']:,.0f}",
                           "Remaining": f"${i['remaining']:,.0f}",
                           "Flag": i["description"],
                           "Has Email": "✅" if i.get("owner_email") else "❌"}
                          for i in budget_issues],
                         use_container_width=True, hide_index=True)

# ---- TAB 3: MONTH HOURS ----
with tab3:
    st.header("Month Hours — Deadline Reminders")
    if not active_month:
        st.error(f"No sheet found for {datetime.now().strftime('%B')}.")
    elif not month_issues:
        st.success(f"✅ No unconfirmed hours within {warn_days} day(s) of deadline.")
    else:
        st.info(f"Sheet: **{active_month}** · {warn_days} day warning")
        st.metric("⏰ Unconfirmed near deadline", len(month_issues))
        st.dataframe([{"Client": i["client"], "Project Code": i["project_code"],
                       "Person": i["person_first"],
                       "Period": i["period"],
                       "Deadline": i["deadline"].strftime("%b %d"),
                       "Days Left": i["days_left"], "Hours": i["hours"],
                       "Has Email": "✅" if i.get("person_email") else "❌"}
                      for i in month_issues],
                     use_container_width=True, hide_index=True)

# ---- TAB 4: VARIANCE ----
with tab4:
    st.header("Actual vs Schedule Variance (OpenAir)")
    if not openair_file:
        st.info("Upload an OpenAir CSV above to enable variance analysis.")
    elif openair_error:
        st.error(f"Error parsing OpenAir file: {openair_error}")
    elif not selected_months:
        st.warning("Select at least one period above to compute variances.")
    elif not variance_issues:
        st.success("✅ No variances over 3 hours for the selected period(s).")
    else:
        if len(selected_months) > 1:
            st.warning(f"📢 Showing variances across **{len(selected_months)} periods**: "
                       f"{', '.join(selected_months)}")
        st.metric("⚠️ Variances > 3 hrs", len(variance_issues))
        st.dataframe([{"Person": v["first_name"], "Project": v["project_code"],
                       "Period": v["period"],
                       "Actual Hrs": v["actual_hours"], "Sched Hrs": v["sched_hours"],
                       "Diff": v["difference"], "To Review": v["question"]}
                      for v in variance_issues],
                     use_container_width=True, hide_index=True)

# ============================================================
# COMBINED EMAIL SECTION
# ============================================================
st.divider()
st.header("📧 Combined Emails by Owner")
st.caption("One email per person · Project Tracker · Budget · Scheduled Hours · Variance")

active_owners = {
    owner: data for owner, data in owners_data.items()
    if data["tracker"] or data["budget"] or data["month"] or data["variance"]
}

if not active_owners:
    st.info("No flagged items found — no emails to send.")
else:
    st.metric("People with flagged items", len(active_owners))

    # ---- Per-owner preview + individual checkbox ----
    st.subheader("Select recipients")
    selected_owners = []

    for owner, data in active_owners.items():
        owner_email  = data.get("email")
        first_name   = data.get("first_name", owner)
        col_chk, col_info = st.columns([0.05, 0.95])
        with col_chk:
            checked = st.checkbox("", value=True, key=f"chk_{owner}",
                                  label_visibility="collapsed")
        with col_info:
            label = (f"**{first_name} ({owner})**  ·  {owner_email or '⚠️ no email'}  —  "
                     f"Tracker: {len(data['tracker'])} · "
                     f"Budget: {len(data['budget'])} · "
                     f"Hours: {len(data['month'])} · "
                     f"Variance: {len(data['variance'])}")
            st.markdown(label)

        if checked:
            selected_owners.append(owner)

        # Email preview
        with st.expander(f"👁 Preview email for {first_name}"):
            if not owner_email:
                st.error("No email address — add to config.py")
            else:
                html = build_html_email(
                    owner, first_name,
                    data["tracker"], data["budget"],
                    data["month"],   data["variance"],
                    has_openair=has_openair,
                    selected_months=selected_months if len(selected_months) > 1 else None,
                )
                if html:
                    st.components.v1.html(html, height=500, scrolling=True)
                else:
                    st.info("No content to preview.")

    st.divider()

    no_email = [o for o in selected_owners
                if not active_owners[o].get("email")]
    if no_email:
        st.warning(f"⚠️ No email for: {', '.join(no_email)} — will be skipped.")

    sendable = [o for o in selected_owners if active_owners[o].get("email")]

    col_b1, col_b2 = st.columns([0.3, 0.7])
    with col_b1:
        send_selected = st.button(
            f"📤 Send to Selected ({len(sendable)})",
            type="primary", key="send_selected",
            disabled=not email_configured() or not sendable)
    with col_b2:
        send_all = st.button(
            f"📤 Send All ({len([o for o in active_owners if active_owners[o].get('email')])})",
            key="send_all",
            disabled=not email_configured())

    if not email_configured():
        st.info("Configure SendGrid credentials in Streamlit Secrets to enable sending.")

    def _do_send(owners_to_send):
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

    if send_selected and sendable:
        _do_send(sendable)

    if send_all:
        all_sendable = [o for o in active_owners if active_owners[o].get("email")]
        _do_send(all_sendable)
