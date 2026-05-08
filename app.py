# ============================================================
# app.py — GTM Scheduling Analyzer  |  streamlit run app.py
# ============================================================

import openpyxl
import streamlit as st
from collections import defaultdict
from datetime import datetime

from config import (LAREN_EMAIL, EMAIL_LOOKUP,
                    DEFAULT_BUDGET_THRESHOLD, DEADLINE_WARNING_DAYS)
from email_utils import build_html_email, send_outlook_email, build_and_send_combined_emails
from processors.budget_actual   import process_budget_actual
from processors.project_tracker import process_project_tracker
from processors.month_tab       import process_month_tab, build_month_emails
from processors.variance        import compute_variances, read_schedule_hours

# ---- Email availability ----

# ============================================================
# PAGE CONFIG
# ============================================================
st.set_page_config(page_title="GTM Scheduling Analyzer",
                   layout="wide", page_icon="📊")

st.title("📊 GTM Scheduling Analyzer")
st.caption(f"Today: {datetime.now().strftime('%A, %B %d, %Y')}")

if not email_configured():
    st.warning("⚠️ **Email credentials not configured.** Add them in Streamlit Secrets to enable sending. "
               "All previews still work.")

# ============================================================
# SIDEBAR
# ============================================================
with st.sidebar:
    st.header("⚙️ Settings")
    laren_email = st.text_input("From / CC (Laren)", value=LAREN_EMAIL)

    st.divider()
    st.subheader("💰 Budget to Actual")
    budget_threshold = st.number_input(
        "Flag remaining budget over ($)",
        value=int(DEFAULT_BUDGET_THRESHOLD), step=1000, min_value=0)

    st.divider()
    st.subheader("⏰ Hour Reminders")
    warn_days = st.number_input(
        "Warn X days before period deadline",
        value=DEADLINE_WARNING_DAYS, min_value=1, max_value=14)

    st.divider()
    st.subheader("📋 Email Lookup")
    st.caption(f"{len(EMAIL_LOOKUP)} people configured")
    if st.checkbox("Show lookup table"):
        st.dataframe([{"Name": k, "Email": v}
                      for k, v in EMAIL_LOOKUP.items()],
                     use_container_width=True, hide_index=True)
    st.caption("Edit `config.py` to add / update names.")

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
        "OpenAir Report (optional — enables variance tab)",
        type=["xlsx", "csv"], key="openair")

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
c1.info(f"💰 Budget tab: **{budget_sheet_name or 'Not found'}**")
c2.info(f"📋 Tracker tab: **{tracker_sheet_name or 'Not found'}**")
c3.info(f"📅 Month tab: auto-detect current month")

st.divider()

# ============================================================
# PROCESS ALL DATA
# ============================================================
with st.spinner("Analyzing…"):
    # Budget to Actual
    budget_issues = []
    if budget_sheet_name:
        budget_issues = process_budget_actual(
            wb[budget_sheet_name], float(budget_threshold))

    # Project Tracker
    tracker_issues, tbd_projects = [], []
    if tracker_sheet_name:
        tracker_issues, tbd_projects = process_project_tracker(
            wb[tracker_sheet_name])

    # Month hours
    month_issues, active_month = process_month_tab(wb, int(warn_days))

    # Variance (only if OpenAir file uploaded)
    variance_issues = []
    openair_available = False
    openair_error = None
    if openair_file:
        try:
            from processors.variance import (parse_openair_report,
                                             read_schedule_hours, compute_variances)
            actual_data = parse_openair_report(openair_file)
            if active_month:
                sched_data = read_schedule_hours(wb, active_month)
                variance_issues = compute_variances(actual_data, sched_data)
            openair_available = True
        except Exception as e:
            openair_error = str(e)
            openair_available = False

    # ---- Build owner-keyed combined data ----
    owners_data = defaultdict(lambda: {
        "email": None, "tracker": [], "budget": [], "variance": []})

    for issue in tracker_issues:
        owner = issue["owner"]
        owners_data[owner]["email"] = issue.get("owner_email")
        owners_data[owner]["tracker"].append(issue)

    for issue in budget_issues:
        owner = issue["owner"]
        if not owners_data[owner]["email"]:
            owners_data[owner]["email"] = issue.get("owner_email")
        owners_data[owner]["budget"].append(issue)

    for v in variance_issues:
        owner = v["person"]
        owners_data[owner]["variance"].append(v)

# ============================================================
# TABS
# ============================================================
tab1, tab2, tab3, tab4 = st.tabs(
    ["📋 Project Tracker", "💰 Budget to Actual",
     "📅 Month Hours", "📊 Variance (OpenAir)"])


# ============================================================
# TAB 1 — PROJECT TRACKER
# ============================================================
with tab1:
    st.header("Project Tracker — Known Projects")
    if not tracker_sheet_name:
        st.error("No Project Tracker tab found.")
    else:
        c1, c2 = st.columns(2)
        c1.metric("⚠️ Issues Found",  len(tracker_issues))
        c2.metric("📌 TBD Projects",  len(tbd_projects))

        if tbd_projects:
            with st.expander(f"📌 {len(tbd_projects)} TBD projects (excluded)"):
                st.dataframe([{"Client": p["client"],
                               "Project Code": p["project_code"],
                               "Owner": p["owner"],
                               "Budget": f"${p['budget']:,.0f}"}
                              for p in tbd_projects],
                             use_container_width=True, hide_index=True)

        if not tracker_issues:
            st.success("✅ No issues found in Known projects!")
        else:
            # Flat table: one row per problem
            flat = []
            for issue in tracker_issues:
                for prob in issue["problems"]:
                    flat.append({
                        "Client":       issue["client"],
                        "Project Code": issue["project_code"],
                        "Owner":        issue["owner"],
                        "Issue":        prob,
                        "Has Email":    "✅" if issue.get("owner_email") else "❌",
                    })
            st.dataframe(flat, use_container_width=True, hide_index=True)

            _no_email_t = [i for i in tracker_issues if not i.get("owner_email")]
            if _no_email_t:
                names = sorted(set(i["owner"] for i in _no_email_t))
                st.warning(f"⚠️ No email found for: {', '.join(names)} — edit config.py")


# ============================================================
# TAB 2 — BUDGET TO ACTUAL
# ============================================================
with tab2:
    st.header("Budget to Actual — Known Projects")
    if not budget_sheet_name:
        st.error("No Budget to Actual tab found.")
    else:
        neg     = [i for i in budget_issues if i["type"] == "negative"]
        not_proj= [i for i in budget_issues if i["type"] == "not_projected"]
        c1, c2, c3 = st.columns(3)
        c1.metric("🔴 Over Budget",    len(neg))
        c2.metric("🟡 Under-Scheduled",len(not_proj))
        c3.metric("⚠️ Total",          len(budget_issues))

        if not budget_issues:
            st.success("✅ No issues found!")
        else:
            st.dataframe([{
                "Client":       i["client"],
                "Project Code": i["project_code"],
                "Owner":        i["owner"],
                "Budget":       f"${i['budget']:,.0f}",
                "Remaining":    f"${i['remaining']:,.0f}",
                "Flag":         i["description"],
                "Has Email":    "✅" if i.get("owner_email") else "❌",
            } for i in budget_issues],
            use_container_width=True, hide_index=True)

            _no_email_b = [i for i in budget_issues if not i.get("owner_email")]
            if _no_email_b:
                names = sorted(set(i["owner"] for i in _no_email_b))
                st.warning(f"⚠️ No email found for: {', '.join(names)} — edit config.py")


# ============================================================
# TAB 3 — MONTH HOURS
# ============================================================
with tab3:
    st.header("Month Hours — Deadline Reminders")
    st.caption(f"Checking **{datetime.now().strftime('%B')}** tab · "
               f"warning {warn_days} day(s) before period end")

    if active_month:
        st.info(f"Active sheet: **{active_month}**")
    else:
        st.error(f"No sheet found for {datetime.now().strftime('%B')}.")
        st.stop()

    if not month_issues:
        st.success(f"✅ No unconfirmed hours within {warn_days} day(s) of deadline.")
    else:
        st.metric("⏰ Unconfirmed near deadline", len(month_issues))
        st.dataframe([{
            "Client":       i["client"],
            "Project Code": i["project_code"],
            "Assigned To":  i["person"],
            "Period":       i["period"],
            "Deadline":     i["deadline"].strftime("%b %d"),
            "Days Left":    i["days_left"],
            "Hours":        i["hours"],
            "Has Email":    "✅" if i.get("person_email") else "❌",
        } for i in month_issues],
        use_container_width=True, hide_index=True)

        _no_email_m = [i for i in month_issues if not i.get("person_email")]
        if _no_email_m:
            names = sorted(set(i["person"] for i in _no_email_m))
            st.warning(f"⚠️ No email found for: {', '.join(names)}")

        if email_configured() and month_issues:
            hour_emails = build_month_emails(month_issues, laren_email)
            with st.expander(f"📧 Preview {len(hour_emails)} reminder email(s)"):
                for e in hour_emails:
                    st.markdown(f"**To:** {e['to']}  |  **Subject:** {e['subject']}")
                    st.text(e["body"])
                    st.divider()
            if st.button("📤 Send Hour Reminders", type="primary", key="send_hours"):
                with st.spinner("Sending…"):
                    from email_utils import send_outlook_email as _send
                    results = [_send(e["to"], e["subject"],
                                     f"<pre>{e['body']}</pre>", laren_email)
                               for e in hour_emails]
                sent = [r for r in results if r["status"] == "sent"]
                fail = [r for r in results if r["status"] != "sent"]
                if sent:  st.success(f"✅ {len(sent)} sent!")
                if fail:  st.error(f"❌ {len(fail)} failed: " +
                                   ", ".join(r["to"] for r in fail))


# ============================================================
# TAB 4 — VARIANCE
# ============================================================
with tab4:
    st.header("Actual vs Schedule Variance (OpenAir)")
    if not openair_file:
        st.info("Upload an OpenAir report in the file section above to enable this tab.")
    elif openair_error:
        st.error(f"Error parsing OpenAir file: {openair_error}")
    elif not openair_available:
        st.warning("Could not process OpenAir file.")
    else:
        flagged = [v for v in variance_issues if abs(v["difference"]) > 0]
        st.metric("⚠️ Variances > 3 hours", len(variance_issues))

        if not variance_issues:
            st.success("✅ No variances over 3 hours found!")
        else:
            st.dataframe([{
                "Person":          v["person"],
                "Project":         v["project_code"],
                "Period":          v["period"],
                "Actual Hrs":      v["actual_hours"],
                "Scheduled Hrs":   v["sched_hours"],
                "Difference":      v["difference"],
                "Flag":            v["question"][:60] + "…",
            } for v in variance_issues],
            use_container_width=True, hide_index=True)

            st.caption("💡 Variances are included in each owner's combined email below.")


# ============================================================
# COMBINED EMAIL SECTION
# ============================================================
st.divider()
st.header("📧 Combined Emails by Owner")
st.caption("One email per project owner covering all three sections: "
           "Project Tracker + Budget to Actual + Variances")

active_owners = {
    owner: data for owner, data in owners_data.items()
    if data["tracker"] or data["budget"] or data["variance"]
}

if not active_owners:
    st.info("No flagged items found — no emails to send.")
else:
    st.metric("Owners with flagged items", len(active_owners))

    for owner, data in active_owners.items():
        owner_email = data.get("email")
        t_count = len(data["tracker"])
        b_count = len(data["budget"])
        v_count = len(data["variance"])

        label = (f"📨 **{owner}** ({owner_email or '⚠️ no email'})  —  "
                 f"Tracker: {t_count} issue(s) · "
                 f"Budget: {b_count} issue(s) · "
                 f"Variance: {v_count} item(s)")

        with st.expander(label):
            if not owner_email:
                st.error("No email address found — add to config.py")
            else:
                # Live HTML preview
                html = build_html_email(
                    owner,
                    data["tracker"],
                    data["budget"],
                    data["variance"],
                )
                st.markdown("**Email preview:**")
                st.components.v1.html(html, height=500, scrolling=True)

    st.divider()
    if email_configured():
        no_email_owners = [o for o, d in active_owners.items() if not d.get("email")]
        if no_email_owners:
            st.warning(f"⚠️ Will skip (no email): {', '.join(no_email_owners)}")

        if st.button("📤 Send All Combined Emails", type="primary",
                     key="send_combined"):
            with st.spinner("Sending via Outlook…"):
                results = build_and_send_combined_emails(
                    active_owners, cc_email=laren_email)
            sent = [r for r in results if r.get("status") == "sent"]
            fail = [r for r in results if r.get("status") != "sent"]
            if sent: st.success(f"✅ {len(sent)} email(s) sent!")
            if fail: st.error(f"❌ {len(fail)} failed:")
            for r in fail:
                st.write(f"  • {r['to']}: {r['status']}")
    else:
        st.info("Configure email credentials in Streamlit Secrets to enable sending.")
