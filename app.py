from __future__ import annotations
"""
GTM Scheduling Analyzer — Streamlit App
"""

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
from processors.budget_actual import process_budget_actual
from processors.project_tracker import process_project_tracker
from processors.variance import (
    parse_openair_report, read_schedule_hours, compute_variances,
)
from processors.utilization import process_utilization, build_utilization_emails
from email_utils import email_configured, send_email, EMAIL_OK, send_emails_batch

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

def _sender_name() -> str:
    try:
        from_email = st.secrets["email"]["from_email"]
    except (KeyError, FileNotFoundError):
        from_email = SENDER_EMAIL
    return SENDER_NAMES.get(from_email, "Jake")


def _find_sheet(sheetnames: list, keywords: list) -> str | None:
    for name in sheetnames:
        if any(kw in name.lower() for kw in keywords):
            return name
    return None


def _lookup_email(name: str) -> str | None:
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
# CACHED PROCESSORS
# Each function takes bytes, opens its own workbook internally,
# and returns only plain dicts/lists — easily serializable.
# The workbook object is NEVER cached or returned.
# ============================================================

@st.cache_data(show_spinner=False)
def _get_sheet_names(file_bytes: bytes) -> list:
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    names = wb.sheetnames
    wb.close()
    return names


@st.cache_data(show_spinner=False)
def _get_valid_people(file_bytes: bytes, month: str) -> list:
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    people = []
    for name in wb.sheetnames:
        if name.lower().startswith(month[:3].lower()):
            ws = wb[name]
            for col in range(7, 45):
                val = ws.cell(row=2, column=col).value
                if val:
                    people.append(str(val).strip())
            break
    wb.close()
    return people


@st.cache_data(show_spinner=False)
def _run_budget(file_bytes: bytes, budget_thresh: float,
                proj_pct: float, neg_thresh: float) -> list:
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    sheet = _find_sheet(wb.sheetnames, ["budget to actual", "budget"])
    result = []
    if sheet:
        result = process_budget_actual(wb[sheet], budget_thresh, proj_pct, neg_thresh)
    wb.close()
    return result


@st.cache_data(show_spinner=False)
def _run_tracker(file_bytes: bytes) -> tuple:
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    sheet = _find_sheet(wb.sheetnames, ["project tracker", "tracker"])
    result = ([], [])
    if sheet:
        result = process_project_tracker(wb[sheet])
    wb.close()
    return result


@st.cache_data(show_spinner=False)
def _run_variance(sched_bytes: bytes, openair_bytes: bytes,
                  month: str, var_min: float, var_max: float) -> tuple:
    """Returns (variance_issues: list, has_openair: bool, error: str|None)"""
    if not openair_bytes:
        return [], False, None

    wb = openpyxl.load_workbook(io.BytesIO(sched_bytes), data_only=True)
    sched_sheet = next(
        (s for s in wb.sheetnames if s.lower().startswith(month[:3].lower())),
        None,
    )
    if not sched_sheet:
        wb.close()
        return [], False, f"No sheet found for {month}"

    try:
        sched_data  = read_schedule_hours(wb, sched_sheet)
        actual_data = parse_openair_report(io.BytesIO(openair_bytes))
        issues      = compute_variances(actual_data, sched_data,
                                        min_diff=var_min, max_diff=var_max)
        wb.close()
        return issues, True, None
    except Exception as e:
        wb.close()
        return [], False, str(e)


@st.cache_data(show_spinner=False)
def _run_utilization(file_bytes: bytes, month: str) -> list:
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    result = process_utilization(wb, target_month=month)
    wb.close()
    return result


# ============================================================
# SIDEBAR
# ============================================================
with st.sidebar:
    st.header("⚙️ Settings")

    st.subheader("💰 Budget to Actual")
    budget_threshold = st.number_input(
        "Flag unscheduled remaining >= ($)",
        value=int(DEFAULT_BUDGET_THRESHOLD), step=1000, min_value=0,
    )
    negative_threshold = st.number_input(
        "Flag negative remaining <= -($)",
        value=int(DEFAULT_NEGATIVE_THRESHOLD), step=50, min_value=0,
    )

    st.divider()
    st.subheader("📊 Variance (Actual vs Schedule)")
    vc1, vc2 = st.columns(2)
    with vc1:
        variance_min = st.number_input(
            "Min (flag if <=)",
            value=float(DEFAULT_VARIANCE_MIN), step=1.0,
        )
    with vc2:
        variance_max = st.number_input(
            "Max (flag if >=)",
            value=float(DEFAULT_VARIANCE_MAX), step=1.0,
        )

    st.divider()
    st.subheader("📋 Email Lookup")
    st.caption(f"{len(EMAIL_LOOKUP)} people configured")
    if st.checkbox("Show lookup table"):
        st.dataframe(
            [{"Name": k, "Email": v} for k, v in EMAIL_LOOKUP.items()],
            use_container_width=True, hide_index=True,
        )

if not EMAIL_OK:
    st.sidebar.warning(
        "SendGrid keys not found in Streamlit Secrets. "
        "Add them to enable sending. Previews still work."
    )

sender_name  = _sender_name()
active_month = datetime.now().strftime("%B")

# ============================================================
# FILE UPLOADS
# ============================================================
st.subheader("📁 Upload Files")
col_f1, col_f2 = st.columns(2)
with col_f1:
    schedule_file = st.file_uploader(
        "Schedule File (.xlsx)", type=["xlsx", "csv"], key="schedule",
    )
with col_f2:
    openair_file = st.file_uploader(
        "OpenAir Report (.csv or .xlsx) — optional",
        type=["csv", "xlsx"], key="openair",
    )

if not schedule_file:
    st.info("Upload the scheduling file above to begin.")
    st.stop()

# Safely read bytes — order-independent
try:
    sched_bytes = bytes(schedule_file.read())
except Exception as e:
    st.error(f"Could not read schedule file: {e}")
    st.stop()

openair_bytes = None
if openair_file:
    try:
        openair_bytes = bytes(openair_file.read())
    except Exception:
        st.warning("Could not read OpenAir file — variance will be skipped.")

# Quick validation
try:
    sheets = _get_sheet_names(sched_bytes)
except Exception as e:
    st.error(f"Could not open schedule file: {e}")
    st.stop()

budget_sheet  = _find_sheet(sheets, ["budget to actual", "budget"])
tracker_sheet = _find_sheet(sheets, ["project tracker", "tracker"])

st.success(f"Loaded {len(sheets)} sheet(s)")
c1, c2 = st.columns(2)
c1.info(f"💰 Budget tab: **{budget_sheet or 'Not found'}**")
c2.info(f"📋 Tracker tab: **{tracker_sheet or 'Not found'}**")
st.divider()

# ============================================================
# PROCESS (all cached, each in its own try/except)
# ============================================================
with st.spinner("Analyzing…"):

    budget_issues = []
    if budget_sheet:
        try:
            budget_issues = _run_budget(
                sched_bytes, float(budget_threshold),
                DEFAULT_PROJECTION_THRESHOLD_PCT, float(negative_threshold),
            )
        except Exception as e:
            st.warning(f"Budget to Actual error: {e}")

    tracker_issues, tbd_projects = [], []
    try:
        tracker_issues, tbd_projects = _run_tracker(sched_bytes)
    except Exception as e:
        st.warning(f"Project Tracker error: {e}")

    variance_issues, has_openair, openair_error = [], False, None
    try:
        variance_issues, has_openair, openair_error = _run_variance(
            sched_bytes,
            openair_bytes or b"",   # never pass None to cached fn
            active_month,
            float(variance_min), float(variance_max),
        )
    except Exception as e:
        openair_error = str(e)

    util_data = []
    try:
        util_data = _run_utilization(sched_bytes, active_month)
    except Exception as e:
        st.warning(f"Utilization error: {e}")

    try:
        valid_people = set(_get_valid_people(sched_bytes, active_month))
    except Exception:
        valid_people = set()

    # Build owner map
    owners_data = defaultdict(lambda: {
        "email": None, "first_name": "there",
        "tracker": [], "budget": [], "variance": [],
    })

    for issue in tracker_issues:
        o = issue.get("owner", "")
        if not owners_data[o]["email"]:
            owners_data[o]["email"]      = issue.get("owner_email")
            owners_data[o]["first_name"] = issue.get("owner_first", o)
        owners_data[o]["tracker"].append(issue)

    for issue in budget_issues:
        o = issue.get("owner", "")
        if not owners_data[o]["email"]:
            owners_data[o]["email"]      = issue.get("owner_email")
            owners_data[o]["first_name"] = issue.get("owner_first", o)
        owners_data[o]["budget"].append(issue)

    for v in variance_issues:
        p = v.get("person", "")
        if not owners_data[p]["email"]:
            owners_data[p]["email"] = _lookup_email(p)
        owners_data[p]["variance"].append(v)

    active_owners = {
        owner: data for owner, data in owners_data.items()
        if (data["tracker"] or data["budget"] or data["variance"])
        and (not valid_people or owner in valid_people)
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

# ---- TAB 1 ----
with tab1:
    st.header("Project Tracker — Known Projects")
    if not tracker_sheet:
        st.warning("No Project Tracker tab found in the uploaded file.")
    else:
        c1, c2 = st.columns(2)
        c1.metric("Issues Found",      len(tracker_issues))
        c2.metric("TBD / Pending SOW", len(tbd_projects))

        if tbd_projects:
            with st.expander(f"{len(tbd_projects)} TBD / Pending SOW projects"):
                st.dataframe(
                    [{
                        "Client":       p["client"],
                        "Project Code": p["project_code"],
                        "Status":       p.get("status", "TBD"),
                        "Owner":        p["owner"],
                        "Budget":       f"${p['budget']:,.0f}",
                    } for p in tbd_projects],
                    use_container_width=True, hide_index=True,
                )

        if not tracker_issues:
            st.success("No issues found in Known projects!")
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

# ---- TAB 2 ----
with tab2:
    st.header("Budget to Actual — Known Projects")
    if not budget_sheet:
        st.warning("No Budget to Actual tab found in the uploaded file.")
    else:
        neg      = [i for i in budget_issues if i.get("type") == "negative"]
        not_proj = [i for i in budget_issues if i.get("type") == "not_projected"]
        c1, c2, c3 = st.columns(3)
        c1.metric("Over Budget",     len(neg))
        c2.metric("Under-Scheduled", len(not_proj))
        c3.metric("Total",           len(budget_issues))

        if not budget_issues:
            st.success("No issues found!")
        else:
            st.dataframe(
                [{
                    "Client":       i["client"],
                    "Project Code": i["project_code"],
                    "Owner":        i["owner"],
                    "Budget":       f"${i['budget']:,.0f}",
                    "Remaining":    f"${i['remaining']:,.0f}",
                    "Flag":         i.get("description", ""),
                    "Has Email":    "✅" if i.get("owner_email") else "❌",
                } for i in budget_issues],
                use_container_width=True, hide_index=True,
            )

# ---- TAB 3 ----
with tab3:
    st.header("Actual vs Schedule Variance")

    if not has_openair and not openair_error:
        st.info("No OpenAir report uploaded. Upload one above to see actuals.")
    elif openair_error:
        st.error(f"Error parsing OpenAir file: {openair_error}")

    if not variance_issues:
        if has_openair:
            st.success(f"No variances outside [{variance_min:+.0f}, {variance_max:+.0f}] hours.")
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

# ---- TAB 4 ----
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
            use_container_width=True, hide_index=True,
        )


# ============================================================
# COMBINED EMAILS
# ============================================================
st.divider()
st.header("📧 Combined Emails")
st.caption(
    "One email per person: Project Tracker + Budget to Actual + "
    "TBD/Pending SOW + Variance + Utilization."
)

util_emails_by_person = {
    e["person"]: e
    for e in build_utilization_emails(
        util_data, month=active_month, sender_name=sender_name
    )
}

all_people = set(active_owners.keys()) | set(util_emails_by_person.keys())

if not all_people:
    st.info("No flagged items — no emails to send.")
else:
    if "selected_people" not in st.session_state:
        st.session_state.selected_people = set(all_people)

    bc1, bc2 = st.columns(2)
    with bc1:
        if st.button("Select All"):
            st.session_state.selected_people = set(all_people)
            st.rerun()
    with bc2:
        if st.button("Deselect All"):
            st.session_state.selected_people = set()
            st.rerun()

    for person in sorted(all_people):
        checked = person in st.session_state.selected_people
        if st.checkbox(person, value=checked, key=f"chk_{person}"):
            st.session_state.selected_people.add(person)
        else:
            st.session_state.selected_people.discard(person)

    selected       = st.session_state.selected_people
    combined_emails = []

    for person in sorted(selected):
        owner_data    = active_owners.get(person, {})
        person_email  = owner_data.get("email") or _lookup_email(person)
        first_name    = owner_data.get("first_name", person)
        if not person_email:
            continue

        is_intern     = person in INTERN_NAMES
        tracker_list  = owner_data.get("tracker", [])
        budget_list   = owner_data.get("budget", [])
        variance_list = owner_data.get("variance", [])

        if is_intern:
            variance_list = [v for v in variance_list if v.get("person") == person]

        owner_tbd = [p for p in tbd_projects if p.get("owner") == person]
        sections  = []

        if tracker_list:
            lines = [
                "The following projects assigned to you are missing billing rates. "
                "Please review and update as soon as possible.\n"
            ]
            for issue in tracker_list:
                lines.append(
                    f"  - {issue['client']} | {issue['project_code']}\n"
                    f"    Missing: {', '.join(issue.get('missing_rates', []))}"
                )
            sections.append("\n".join(lines))

        if budget_list:
            lines = ["Please review the following budget items:\n"]
            for issue in budget_list:
                lines.append(
                    f"  - {issue['client']} | {issue['project_code']}: "
                    f"${issue['remaining']:,.0f} remaining ({issue.get('description', '')})"
                )
            sections.append("\n".join(lines))

        if owner_tbd:
            lines = [
                "The following projects have TBD or Pending SOW budgets. "
                "If you have any updates, please reply — otherwise no action needed.\n"
            ]
            for proj in owner_tbd:
                label = f" [{proj.get('status', 'TBD')}]"
                lines.append(f"  - {proj['client']} | {proj['project_code']}{label}")
            sections.append("\n".join(lines))

        if variance_list:
            if is_intern:
                lines = ["Please see your variance for the current period:\n"]
            else:
                lines = ["Please see the variance summary for your projects:\n"]
            for v in variance_list:
                prefix = "" if is_intern else f"{v.get('person', '')} | "
                lines.append(
                    f"  - {prefix}{v['project_code']} | {v['period']} | "
                    f"Actual: {v['actual_hours']}h  Scheduled: {v['sched_hours']}h  "
                    f"Diff: {v['difference']:+.1f}h"
                )
                if v.get("question"):
                    lines.append(f"    {v['question']}")
            sections.append("\n".join(lines))

        util_email = util_emails_by_person.get(person)
        if util_email:
            body_lines   = util_email["body"].split("\n")
            util_section = "\n".join(body_lines[2:-3]).strip()
            if util_section:
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

    st.subheader(f"{len(combined_emails)} email(s) ready")

    for email in combined_emails:
        with st.expander(f"To: {email['to']}  |  {email['subject']}"):
            st.text(email["body"])

    if combined_emails:
        if EMAIL_OK:
            if st.button("📤 Send All Emails", type="primary"):
                with st.spinner("Sending…"):
                    results = send_emails_batch(combined_emails)
                sent   = [r for r in results if r.get("status") == "sent"]
                failed = [r for r in results if r.get("status") != "sent"]
                if sent:
                    st.success(f"{len(sent)} email(s) sent!")
                if failed:
                    st.error("Some emails failed:")
                    for r in failed:
                        st.write(f"  • {r.get('to', '?')}: {r.get('status', '?')}")
        else:
            st.info("Configure SendGrid keys in Streamlit Secrets to enable sending.")
