from __future__ import annotations
"""
GTM Scheduling Analyzer — Streamlit App
"""

import io
import warnings
from collections import defaultdict
from datetime import datetime, date

import openpyxl
import streamlit as st

from config import (
    SENDER_EMAIL, SENDER_NAMES,
    DEFAULT_BUDGET_THRESHOLD, DEFAULT_NEGATIVE_THRESHOLD,
    DEFAULT_PROJECTION_THRESHOLD_PCT,
    DEFAULT_VARIANCE_MIN, DEFAULT_VARIANCE_MAX,
    EMAIL_LOOKUP, INTERN_NAMES,
)
from processors.budget_actual import process_budget_actual
from processors.project_tracker import process_project_tracker, build_tracker_emails
from processors.variance import (
    parse_openair_report, read_schedule_hours,
    compute_variances, get_available_months, filter_by_months,
)
from processors.utilization import process_utilization, build_utilization_emails
from email_utils import send_emails_batch, EMAIL_OK

warnings.filterwarnings("ignore", category=UserWarning)

st.set_page_config(
    page_title="GTM Scheduling Analyzer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# HELPERS
# ============================================================

def _get_sender_name() -> str:
    try:
        from_email = st.secrets["email"]["from_email"]
    except (KeyError, FileNotFoundError):
        from_email = SENDER_EMAIL
    return SENDER_NAMES.get(from_email, "Jake")


def lookup_email(name: str) -> str | None:
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
# SIDEBAR — SETTINGS
# ============================================================
with st.sidebar:
    st.header("⚙️ Settings")

    st.subheader("💰 Budget to Actual")
    budget_threshold = st.number_input(
        "Flag unscheduled remaining ≥ ($)",
        value=int(DEFAULT_BUDGET_THRESHOLD), step=1000, min_value=0,
        help="Projects with remaining unscheduled budget at or above this amount will be flagged.",
    )
    negative_threshold = st.number_input(
        "Flag negative remaining ≤ -($)",
        value=int(DEFAULT_NEGATIVE_THRESHOLD), step=50, min_value=0,
        help="Flag if remaining is this far negative or worse (e.g. 100 → flags anything ≤ -$100).",
    )

    st.divider()
    st.subheader("📊 Variance (Actual vs Schedule)")
    variance_col1, variance_col2 = st.columns(2)
    with variance_col1:
        variance_min = st.number_input(
            "Min (flag if ≤)",
            value=DEFAULT_VARIANCE_MIN, step=1,
            help="Flag rows where actual minus schedule is at or below this value.",
        )
    with variance_col2:
        variance_max = st.number_input(
            "Max (flag if ≥)",
            value=DEFAULT_VARIANCE_MAX, step=1,
            help="Flag rows where actual minus schedule is at or above this value.",
        )

    st.divider()
    st.subheader("📋 Email Lookup")
    st.caption(f"{len(EMAIL_LOOKUP)} people configured")
    if st.checkbox("Show lookup table"):
        st.dataframe(
            [{"Name": k, "Email": v} for k, v in EMAIL_LOOKUP.items()],
            use_container_width=True, hide_index=True,
        )


# ============================================================
# EMAIL CREDENTIALS CHECK
# ============================================================
if not EMAIL_OK:
    st.sidebar.warning(
        "⚠️ SendGrid keys not found in Streamlit Secrets. "
        "Add them to enable sending. All previews still work."
    )

sender_name = _get_sender_name()

# ============================================================
# FILE UPLOADS
# ============================================================
st.subheader("📁 Upload Files")
col_f1, col_f2 = st.columns(2)
with col_f1:
    schedule_file = st.file_uploader(
        "Schedule File (.xlsx)", type=["xlsx", "csv"], key="schedule"
    )
with col_f2:
    openair_file = st.file_uploader(
        "OpenAir Report (.csv or .xlsx) — optional, enables actuals in Variance tab",
        type=["csv", "xlsx"], key="openair",
    )

if not schedule_file:
    st.info("Upload the scheduling file above to begin.")
    st.stop()

# ============================================================
# LOAD + CACHE
# ============================================================
@st.cache_data(show_spinner=False)
def load_workbook_cached(file_bytes: bytes):
    return openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)


@st.cache_data(show_spinner=False)
def get_valid_people(file_bytes: bytes) -> set:
    """Read valid staff names from row 2 of the current month tab."""
    wb = load_workbook_cached(file_bytes)
    month_name = datetime.now().strftime("%B")
    for name in wb.sheetnames:
        if name.lower().startswith(month_name[:3].lower()):
            ws = wb[name]
            people = set()
            for col in range(7, 40):
                val = ws.cell(row=2, column=col).value
                if val:
                    people.add(str(val).strip())
            return people
    return set()


file_bytes = schedule_file.read()
schedule_file.seek(0)

with st.spinner("Loading workbook…"):
    wb = load_workbook_cached(file_bytes)

sheets = wb.sheetnames
valid_people = get_valid_people(file_bytes)

def find_sheet(keywords):
    for name in sheets:
        if any(kw in name.lower() for kw in keywords):
            return name
    return None

budget_sheet_name  = find_sheet(["budget to actual", "budget"])
tracker_sheet_name = find_sheet(["project tracker", "tracker"])

st.success(f"✅ Loaded **{len(sheets)}** sheet(s)")
c1, c2 = st.columns(2)
c1.info(f"💰 Budget tab: **{budget_sheet_name or 'Not found'}**")
c2.info(f"📋 Tracker tab: **{tracker_sheet_name or 'Not found'}**")

st.divider()

# ============================================================
# PROCESS ALL DATA
# ============================================================
with st.spinner("Analyzing…"):

    # Budget to Actual
    budget_issues = []
    if budget_sheet_name:
        budget_issues = process_budget_actual(
            wb[budget_sheet_name],
            float(budget_threshold),
            DEFAULT_PROJECTION_THRESHOLD_PCT,
            float(negative_threshold),
        )

    # Project Tracker
    tracker_issues, tbd_projects = [], []
    if tracker_sheet_name:
        tracker_issues, tbd_projects = process_project_tracker(wb[tracker_sheet_name])

    # OpenAir Variance
    variance_issues = []
    has_openair     = False
    openair_error   = None
    actual_data     = {}
    active_month    = datetime.now().strftime("%B")   # e.g. "May"

    if openair_file:
        try:
            actual_data = parse_openair_report(openair_file)
            if active_month in sheets or any(
                s.lower().startswith(active_month[:3].lower()) for s in sheets
            ):
                sched_sheet = next(
                    (s for s in sheets if s.lower().startswith(active_month[:3].lower())),
                    None,
                )
                if sched_sheet:
                    sched_data     = read_schedule_hours(wb, sched_sheet)
                    variance_issues = compute_variances(
                        actual_data, sched_data,
                        min_diff=float(variance_min),
                        max_diff=float(variance_max),
                    )
            has_openair = True
        except Exception as e:
            openair_error = str(e)
    else:
        # No OpenAir — build variance with actuals = 0 so the tab still renders
        sched_sheet = next(
            (s for s in sheets if s.lower().startswith(active_month[:3].lower())),
            None,
        )
        if sched_sheet:
            sched_data = read_schedule_hours(wb, sched_sheet)
            # Build placeholder rows with 0 actuals for every scheduled entry
            for person, projects in sched_data.items():
                for proj_code, periods in projects.items():
                    for period, sched_hrs in periods.items():
                        diff = 0.0 - sched_hrs
                        if diff <= float(variance_min) or diff >= float(variance_max):
                            variance_issues.append({
                                "person":       person,
                                "first_name":   person,
                                "project_code": proj_code,
                                "period":       period,
                                "actual_hours": 0.0,
                                "sched_hours":  round(sched_hrs, 1),
                                "difference":   round(diff, 1),
                                "question":     "",
                                "person_email": lookup_email(person),
                            })

    # Utilization
    util_data = process_utilization(wb, target_month=active_month)

    # ---- Build owner-keyed combined data ----
    owners_data = defaultdict(lambda: {
        "email": None, "first_name": "there",
        "tracker": [], "budget": [], "variance": [],
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

    for v in variance_issues:
        person = v["person"]
        if person not in valid_people:
            continue
        if not owners_data[person]["email"]:
            owners_data[person]["email"] = lookup_email(person)
        owners_data[person]["variance"].append(v)

    active_owners = {
        owner: data for owner, data in owners_data.items()
        if (data["tracker"] or data["budget"] or data["variance"])
        and owner in valid_people
    }


# ============================================================
# TABS
# ============================================================
tab1, tab2, tab3, tab4 = st.tabs([
    "📋 Project Tracker",
    "💰 Budget to Actual",
    "📊 Variance (OpenAir)",
    "📈 Utilization",
])


# ============================================================
# TAB 1 — PROJECT TRACKER
# ============================================================
with tab1:
    st.header("Project Tracker — Known Projects")
    if not tracker_sheet_name:
        st.error("No Project Tracker tab found.")
    else:
        c1, c2 = st.columns(2)
        c1.metric("⚠️ Issues Found",      len(tracker_issues))
        c2.metric("📌 TBD / Pending SOW", len(tbd_projects))

        if tbd_projects:
            with st.expander(f"📌 {len(tbd_projects)} TBD / Pending SOW projects"):
                st.dataframe(
                    [{
                        "Client":       p["client"],
                        "Project Code": p["project_code"],
                        "Status":       p["status"],
                        "Owner":        p["owner"],
                        "Budget":       f"${p['budget']:,.0f}",
                    } for p in tbd_projects],
                    use_container_width=True, hide_index=True,
                )

        if not tracker_issues:
            st.success("✅ No issues found in Known projects!")
        else:
            st.dataframe(
                [{
                    "Client":        i["client"],
                    "Project Code":  i["project_code"],
                    "Owner":         i["owner"],
                    "Missing Rates": ", ".join(i.get("missing_rates", [])),
                    "Has Email":     "✅" if i.get("owner_email") else "❌",
                } for i in tracker_issues],
                use_container_width=True, hide_index=True,
            )


# ============================================================
# TAB 2 — BUDGET TO ACTUAL
# ============================================================
with tab2:
    st.header("Budget to Actual — Known Projects")
    if not budget_sheet_name:
        st.error("No Budget to Actual tab found.")
    else:
        neg      = [i for i in budget_issues if i["type"] == "negative"]
        not_proj = [i for i in budget_issues if i["type"] == "not_projected"]
        c1, c2, c3 = st.columns(3)
        c1.metric("🔴 Over Budget",     len(neg))
        c2.metric("🟡 Under-Scheduled", len(not_proj))
        c3.metric("⚠️ Total",           len(budget_issues))

        if not budget_issues:
            st.success("✅ No issues found!")
        else:
            st.dataframe(
                [{
                    "Client":       i["client"],
                    "Project Code": i["project_code"],
                    "Owner":        i["owner"],
                    "Budget":       f"${i['budget']:,.0f}",
                    "Remaining":    f"${i['remaining']:,.0f}",
                    "Flag":         i["description"],
                    "Has Email":    "✅" if i.get("owner_email") else "❌",
                } for i in budget_issues],
                use_container_width=True, hide_index=True,
            )


# ============================================================
# TAB 3 — VARIANCE (OpenAir)
# ============================================================
with tab3:
    st.header("Actual vs Schedule Variance")

    if not has_openair:
        st.info(
            "No OpenAir report uploaded — actuals are shown as 0. "
            "Upload an OpenAir file above to populate real actuals."
        )
    elif openair_error:
        st.error(f"Error parsing OpenAir file: {openair_error}")

    if not variance_issues:
        st.success(
            f"✅ No variances outside the range "
            f"[{variance_min:+.0f}, {variance_max:+.0f}] hours."
        )
    else:
        st.metric("Flagged Variances", len(variance_issues))
        st.dataframe(
            [{
                "Person":          v["person"],
                "Project Code":    v["project_code"],
                "Period":          v["period"],
                "Actual (hrs)":    v["actual_hours"],
                "Scheduled (hrs)": v["sched_hours"],
                "Difference":      v["difference"],
                "Question":        v.get("question", ""),
            } for v in variance_issues],
            use_container_width=True, hide_index=True,
        )

        no_email = [v for v in variance_issues if not v.get("person_email")]
        if no_email:
            missing = sorted(set(v["person"] for v in no_email))
            st.warning(f"⚠️ No email found for: {', '.join(missing)}")


# ============================================================
# TAB 4 — UTILIZATION
# ============================================================
with tab4:
    st.header(f"Utilization — {active_month}")

    if not util_data:
        st.warning(
            "No utilization data found. Make sure the workbook has a "
            "'Utilization by Month' tab with the current month's section."
        )
    else:
        st.dataframe(
            [{
                "Role":        u["role"],
                "Person":      u["person"],
                "Chargeable":  u.get("chargeable"),
                "Holiday":     u.get("holiday"),
                "PTO":         u.get("pto"),
                "Month Total": u.get("month_total"),
                "Remaining":   u.get("remaining"),
                "Utilization": f"{u['utilization_pct']:.1f}%" if u.get("utilization_pct") is not None else "-",
                "Goal":        f"{u['goal_pct']:.0f}%"        if u.get("goal_pct")        is not None else "-",
                "Difference":  f"{u['difference_pct']:+.1f}%" if u.get("difference_pct") is not None else "-",
                "Has Email":   "✅" if u.get("person_email") else "❌",
            } for u in util_data],
            use_container_width=True, hide_index=True,
        )

        no_email = [u for u in util_data if not u.get("person_email")]
        if no_email:
            st.warning(
                f"⚠️ No email found for: "
                f"{', '.join(u['person'] for u in no_email)}"
            )


# ============================================================
# COMBINED EMAIL SECTION
# ============================================================
st.divider()
st.header("📧 Combined Emails")
st.caption(
    "One email per person covering: Project Tracker issues, "
    "Budget to Actual, TBD/Pending SOW projects, Variance, and Utilization."
)

# Build utilization emails (one per person)
util_emails = build_utilization_emails(util_data, month=active_month, sender_name=sender_name)

# Merge all email data by person
# Start from active_owners (tracker + budget + variance)
# Then layer in utilization
all_email_people = set(active_owners.keys()) | {e["person"] for e in util_emails}

if not all_email_people:
    st.info("No flagged items found — no emails to send.")
else:
    # Build person-selector
    if "selected_people" not in st.session_state:
        st.session_state.selected_people = set(all_email_people)

    col_sel1, col_sel2 = st.columns(2)
    with col_sel1:
        if st.button("Select All"):
            st.session_state.selected_people = set(all_email_people)
            st.rerun()
    with col_sel2:
        if st.button("Deselect All"):
            st.session_state.selected_people = set()
            st.rerun()

    for person in sorted(all_email_people):
        checked = person in st.session_state.selected_people
        if st.checkbox(person, value=checked, key=f"chk_{person}"):
            st.session_state.selected_people.add(person)
        else:
            st.session_state.selected_people.discard(person)

    if st.button("✔ Confirm Selection"):
        st.rerun()

    selected = st.session_state.selected_people

    # ---- Build combined emails ----
    combined_emails = []
    util_by_person  = {e["person"]: e for e in util_emails}

    for person in sorted(selected):
        owner_data   = active_owners.get(person, {})
        person_email = owner_data.get("email") or lookup_email(person)
        first_name   = owner_data.get("first_name", person)
        if not person_email:
            continue

        is_intern = person in INTERN_NAMES

        sections = []

        # --- Tracker issues ---
        if owner_data.get("tracker"):
            lines = [
                "The following projects assigned to you are missing billing rates. "
                "Please review and update as soon as possible.\n"
            ]
            for issue in owner_data["tracker"]:
                lines.append(
                    f"  • {issue['client']} — {issue['project_code']}\n"
                    f"    Missing: {', '.join(issue.get('missing_rates', []))}"
                )
            sections.append("\n".join(lines))

        # --- Budget to Actual issues ---
        if owner_data.get("budget"):
            lines = ["Please review the following budget items:\n"]
            for issue in owner_data["budget"]:
                lines.append(
                    f"  • {issue['client']} — {issue['project_code']}: "
                    f"${issue['remaining']:,.0f} remaining ({issue['description']})"
                )
            sections.append("\n".join(lines))

        # --- TBD / Pending SOW for this owner ---
        owner_tbd = [p for p in tbd_projects if p["owner"] == person]
        if owner_tbd:
            lines = [
                "The following projects currently have TBD or Pending SOW budgets. "
                "If you have any updates on these, please reply with the latest — "
                "otherwise, no action is needed.\n"
            ]
            for proj in owner_tbd:
                label = f" [{proj['status']}]" if proj["status"] else " [TBD]"
                lines.append(f"  • {proj['client']} — {proj['project_code']}{label}")
            sections.append("\n".join(lines))

        # --- Variance ---
        if owner_data.get("variance"):
            variances = owner_data["variance"]
            # Interns only see their own rows; project owners see all with Person column
            if is_intern:
                rows = [v for v in variances if v["person"] == person]
                lines = ["Please see below your variance for the current period:\n"]
                for v in rows:
                    lines.append(
                        f"  • {v['project_code']} | {v['period']} | "
                        f"Actual: {v['actual_hours']}h  Scheduled: {v['sched_hours']}h  "
                        f"Diff: {v['difference']:+.1f}h"
                    )
                    if v.get("question"):
                        lines.append(f"    → {v['question']}")
            else:
                lines = [
                    "Please see below the variance summary for projects assigned to you. "
                    "Review any flagged items and reply with updates as needed.\n"
                ]
                for v in variances:
                    lines.append(
                        f"  • {v['person']} | {v['project_code']} | {v['period']} | "
                        f"Actual: {v['actual_hours']}h  Scheduled: {v['sched_hours']}h  "
                        f"Diff: {v['difference']:+.1f}h"
                    )
                    if v.get("question"):
                        lines.append(f"    → {v['question']}")
            sections.append("\n".join(lines))

        # --- Utilization ---
        util_email_data = util_by_person.get(person)
        if util_email_data:
            # Extract just the body content (strip greeting/sign-off for embedding)
            util_body = util_email_data["body"]
            # Pull the middle section (between greeting and sign-off)
            lines = util_body.split("\n")
            util_section = "\n".join(lines[2:-3]).strip()   # skip "Hi X," and "Best, Jake"
            sections.append(util_section)

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
            "person":  person,
            "body":    body,
        })

    # ---- Preview + Send ----
    st.subheader(f"📨 {len(combined_emails)} email(s) ready")
    for email in combined_emails:
        with st.expander(f"To: {email['to']}  |  {email['subject']}"):
            st.text(email["body"])

    if combined_emails:
        if EMAIL_OK:
            if st.button("📤 Send All Emails", type="primary"):
                with st.spinner("Sending…"):
                    results = send_emails_batch(combined_emails)
                sent   = [r for r in results if r["status"] == "sent"]
                failed = [r for r in results if r["status"] != "sent"]
                if sent:
                    st.success(f"✅ {len(sent)} email(s) sent!")
                if failed:
                    st.error("❌ Some emails failed:")
                    for r in failed:
                        st.write(f"  • {r['to']}: {r['status']}")
        else:
            st.info("Configure SendGrid keys in Streamlit Secrets to enable sending.")
