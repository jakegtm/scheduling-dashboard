from __future__ import annotations
"""
GTM Scheduling Analyzer — Streamlit App
"""

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
from processors.budget_actual import process_budget_actual
from processors.project_tracker import process_project_tracker
from processors.variance import (
    parse_openair_report, read_schedule_hours, compute_variances,
)
from processors.utilization import process_utilization, build_utilization_emails
from email_utils import EMAIL_OK, send_emails_batch

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


def _md5(b: bytes) -> str:
    return hashlib.md5(b).hexdigest()


# ============================================================
# SINGLE CACHED PROCESSOR
#
# Performance fixes applied:
# 1. hash_funcs uses MD5 instead of Streamlit's slow default hasher —
#    cuts per-interaction overhead from several seconds to ~1ms.
# 2. max_entries=2 caps how much memory the cache holds.
# 3. gc.collect() after workbook close frees memory immediately.
# 4. Workbook is opened once, all processors run inside the same call.
# ============================================================

@st.cache_data(
    hash_funcs={bytes: _md5},
    max_entries=2,
    show_spinner=False,
)
def _run_all(
    sched_bytes: bytes,
    openair_bytes: bytes,
    active_month: str,
    budget_thresh: float,
    proj_pct: float,
    neg_thresh: float,
    var_min: float,
    var_max: float,
) -> dict:
    out = dict(
        sheets=[],
        budget_sheet=None,
        tracker_sheet=None,
        budget_issues=[],
        tracker_issues=[],
        tbd_projects=[],
        variance_issues=[],
        has_openair=False,
        openair_error=None,
        util_data=[],
        valid_people=[],
        load_error=None,
    )

    # ── Open workbook ────────────────────────────────────────
    wb = None
    try:
        wb = openpyxl.load_workbook(
            io.BytesIO(sched_bytes), data_only=True
        )
        out["sheets"] = list(wb.sheetnames)

        budget_sheet  = _find_sheet(out["sheets"], ["budget to actual", "budget"])
        tracker_sheet = _find_sheet(out["sheets"], ["project tracker", "tracker"])
        out["budget_sheet"]  = budget_sheet
        out["tracker_sheet"] = tracker_sheet

        # ── Budget to Actual ─────────────────────────────────
        if budget_sheet:
            try:
                out["budget_issues"] = process_budget_actual(
                    wb[budget_sheet], budget_thresh, proj_pct, neg_thresh
                )
            except Exception:
                pass

        # ── Project Tracker ──────────────────────────────────
        if tracker_sheet:
            try:
                out["tracker_issues"], out["tbd_projects"] = process_project_tracker(
                    wb[tracker_sheet]
                )
            except Exception:
                pass

        # ── Utilization ──────────────────────────────────────
        try:
            out["util_data"] = process_utilization(wb, target_month=active_month)
        except Exception:
            pass

        # ── Valid people ─────────────────────────────────────
        try:
            month_prefix = active_month[:3].lower()
            for name in wb.sheetnames:
                if name.lower().startswith(month_prefix):
                    ws = wb[name]
                    people = []
                    for col in range(7, 45):
                        val = ws.cell(row=2, column=col).value
                        if val:
                            people.append(str(val).strip())
                    out["valid_people"] = people
                    break
        except Exception:
            pass

    except Exception as e:
        out["load_error"] = str(e)
        return out
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass
        gc.collect()

    # ── Variance (OpenAir) ───────────────────────────────────
    if openair_bytes:
        wb2 = None
        try:
            wb2 = openpyxl.load_workbook(
                io.BytesIO(sched_bytes), data_only=True
            )
            month_prefix = active_month[:3].lower()
            sched_sheet  = next(
                (s for s in wb2.sheetnames if s.lower().startswith(month_prefix)),
                None,
            )
            if not sched_sheet:
                out["openair_error"] = f"No schedule sheet found for {active_month}"
            else:
                sched_data  = read_schedule_hours(wb2, sched_sheet)
                wb2.close()
                wb2 = None
                actual_data = parse_openair_report(io.BytesIO(openair_bytes))
                out["variance_issues"] = compute_variances(
                    actual_data, sched_data, min_diff=var_min, max_diff=var_max
                )
                out["has_openair"] = True
        except Exception as e:
            out["openair_error"] = str(e)
        finally:
            if wb2 is not None:
                try:
                    wb2.close()
                except Exception:
                    pass
            gc.collect()

    return out


# ============================================================
# SESSION STATE — initialize before anything else
# ============================================================
if "applied_settings" not in st.session_state:
    st.session_state.applied_settings = {
        "budget_threshold":   float(DEFAULT_BUDGET_THRESHOLD),
        "negative_threshold": float(DEFAULT_NEGATIVE_THRESHOLD),
        "variance_min":       float(DEFAULT_VARIANCE_MIN),
        "variance_max":       float(DEFAULT_VARIANCE_MAX),
    }

if "selected_people" not in st.session_state:
    st.session_state.selected_people = set()


# ============================================================
# SIDEBAR
# ============================================================
with st.sidebar:
    st.header("⚙️ Settings")

    st.subheader("💰 Budget to Actual")
    draft_budget = st.number_input(
        "Flag unscheduled remaining >= ($)",
        value=int(st.session_state.applied_settings["budget_threshold"]),
        step=1000, min_value=0,
    )
    draft_negative = st.number_input(
        "Flag negative remaining <= -($)",
        value=int(st.session_state.applied_settings["negative_threshold"]),
        step=50, min_value=0,
    )

    st.divider()
    st.subheader("📊 Variance (Actual vs Schedule)")
    vc1, vc2 = st.columns(2)
    with vc1:
        draft_var_min = st.number_input(
            "Min (flag if <=)",
            value=float(st.session_state.applied_settings["variance_min"]),
            step=1.0,
        )
    with vc2:
        draft_var_max = st.number_input(
            "Max (flag if >=)",
            value=float(st.session_state.applied_settings["variance_max"]),
            step=1.0,
        )

    if st.button("✔ Apply Settings", type="primary", use_container_width=True):
        st.session_state.applied_settings = {
            "budget_threshold":   float(draft_budget),
            "negative_threshold": float(draft_negative),
            "variance_min":       float(draft_var_min),
            "variance_max":       float(draft_var_max),
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
    st.sidebar.warning("SendGrid keys not found in Streamlit Secrets. Previews still work.")

budget_threshold   = st.session_state.applied_settings["budget_threshold"]
negative_threshold = st.session_state.applied_settings["negative_threshold"]
variance_min       = st.session_state.applied_settings["variance_min"]
variance_max       = st.session_state.applied_settings["variance_max"]

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

try:
    sched_bytes = bytes(schedule_file.read())
    if not sched_bytes:
        st.error("Schedule file appears to be empty.")
        st.stop()
except Exception as e:
    st.error(f"Could not read schedule file: {e}")
    st.stop()

openair_bytes = b""
if openair_file:
    try:
        openair_bytes = bytes(openair_file.read())
    except Exception:
        st.warning("Could not read OpenAir file — variance will be skipped.")


# ============================================================
# PROCESS
# ============================================================
with st.spinner("Analyzing…"):
    data = _run_all(
        sched_bytes, openair_bytes, active_month,
        budget_threshold, DEFAULT_PROJECTION_THRESHOLD_PCT, negative_threshold,
        variance_min, variance_max,
    )

if data["load_error"]:
    st.error(f"Could not open schedule file: {data['load_error']}")
    st.stop()

sheets          = data["sheets"]
budget_sheet    = data["budget_sheet"]
tracker_sheet   = data["tracker_sheet"]
budget_issues   = data["budget_issues"]
tracker_issues  = data["tracker_issues"]
tbd_projects    = data["tbd_projects"]
variance_issues = data["variance_issues"]
has_openair     = data["has_openair"]
openair_error   = data["openair_error"]
util_data       = data["util_data"]
valid_people    = set(data["valid_people"])

st.success(f"Loaded {len(sheets)} sheet(s)")
c1, c2 = st.columns(2)
c1.info(f"💰 Budget tab: **{budget_sheet or 'Not found'}**")
c2.info(f"📋 Tracker tab: **{tracker_sheet or 'Not found'}**")
st.divider()


# ============================================================
# BUILD PER-OWNER MAP
# ============================================================
owners_data = defaultdict(lambda: {
    "email": None, "first_name": "there",
    "tracker": [], "budget": [], "variance": [],
})

for issue in tracker_issues:
    o = issue.get("owner", "")
    if not o:
        continue
    if not owners_data[o]["email"]:
        owners_data[o]["email"]      = issue.get("owner_email")
        owners_data[o]["first_name"] = issue.get("owner_first", o)
    owners_data[o]["tracker"].append(issue)

for issue in budget_issues:
    o = issue.get("owner", "")
    if not o:
        continue
    if not owners_data[o]["email"]:
        owners_data[o]["email"]      = issue.get("owner_email")
        owners_data[o]["first_name"] = issue.get("owner_first", o)
    owners_data[o]["budget"].append(issue)

for v in variance_issues:
    p = v.get("person", "")
    if not p:
        continue
    if not owners_data[p]["email"]:
        owners_data[p]["email"] = _lookup_email(p)
    owners_data[p]["variance"].append(v)

active_owners = {
    owner: d for owner, d in owners_data.items()
    if (d["tracker"] or d["budget"] or d["variance"])
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
                        "Client":       p.get("client", ""),
                        "Project Code": p.get("project_code", ""),
                        "Status":       p.get("status", "TBD"),
                        "Owner":        p.get("owner", ""),
                        "Budget":       f"${p.get('budget', 0):,.0f}",
                    } for p in tbd_projects],
                    use_container_width=True, hide_index=True,
                )

        if not tracker_issues:
            st.success("No issues found in Known projects!")
        else:
            st.dataframe(
                [{
                    "Client":        i.get("client", ""),
                    "Project Code":  i.get("project_code", ""),
                    "Owner":         i.get("owner", ""),
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
                    "Client":       i.get("client", ""),
                    "Project Code": i.get("project_code", ""),
                    "Owner":        i.get("owner", ""),
                    "Budget":       f"${i.get('budget', 0):,.0f}",
                    "Remaining":    f"${i.get('remaining', 0):,.0f}",
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
        st.error(f"Could not parse OpenAir file: {openair_error}")

    if variance_issues:
        st.metric("Flagged Variances", len(variance_issues))
        st.dataframe(
            [{
                "Person":          v.get("person", ""),
                "Project Code":    v.get("project_code", ""),
                "Period":          v.get("period", ""),
                "Actual (hrs)":    v.get("actual_hours", 0),
                "Scheduled (hrs)": v.get("sched_hours", 0),
                "Difference":      v.get("difference", 0),
                "Question":        v.get("question", ""),
            } for v in variance_issues],
            use_container_width=True, hide_index=True,
        )
    elif has_openair:
        st.success(f"No variances outside [{variance_min:+.0f}, {variance_max:+.0f}] hours.")

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

try:
    util_emails_by_person = {
        e["person"]: e
        for e in build_utilization_emails(
            util_data, month=active_month, sender_name=sender_name
        )
    }
except Exception:
    util_emails_by_person = {}

all_people = set(active_owners.keys()) | set(util_emails_by_person.keys())

if not all_people:
    st.info("No flagged items — no emails to send.")
else:
    if not st.session_state.selected_people.issubset(all_people):
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

    combined_emails = []
    for person in sorted(st.session_state.selected_people):
        owner_data   = active_owners.get(person, {})
        person_email = owner_data.get("email") or _lookup_email(person)
        first_name   = owner_data.get("first_name") or person
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
                    f"  - {issue.get('client', '')} | {issue.get('project_code', '')}\n"
                    f"    Missing: {', '.join(issue.get('missing_rates', []))}"
                )
            sections.append("\n".join(lines))

        if budget_list:
            lines = ["Please review the following budget items:\n"]
            for issue in budget_list:
                lines.append(
                    f"  - {issue.get('client', '')} | {issue.get('project_code', '')}: "
                    f"${issue.get('remaining', 0):,.0f} remaining "
                    f"({issue.get('description', '')})"
                )
            sections.append("\n".join(lines))

        if owner_tbd:
            lines = [
                "The following projects have TBD or Pending SOW budgets. "
                "If you have any updates, please reply — otherwise no action needed.\n"
            ]
            for proj in owner_tbd:
                lines.append(
                    f"  - {proj.get('client', '')} | {proj.get('project_code', '')} "
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
                    f"  - {prefix}{v.get('project_code', '')} | {v.get('period', '')} | "
                    f"Actual: {v.get('actual_hours', 0)}h  "
                    f"Scheduled: {v.get('sched_hours', 0)}h  "
                    f"Diff: {v.get('difference', 0):+.1f}h"
                )
                if v.get("question"):
                    lines.append(f"    {v['question']}")
            sections.append("\n".join(lines))

        util_email = util_emails_by_person.get(person)
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
                    try:
                        results = send_emails_batch(combined_emails)
                        sent    = [r for r in results if r.get("status") == "sent"]
                        failed  = [r for r in results if r.get("status") != "sent"]
                        if sent:
                            st.success(f"{len(sent)} email(s) sent!")
                        if failed:
                            st.error("Some emails failed:")
                            for r in failed:
                                st.write(f"  • {r.get('to', '?')}: {r.get('status', '?')}")
                    except Exception as e:
                        st.error(f"Error sending emails: {e}")
        else:
            st.info("Configure SendGrid keys in Streamlit Secrets to enable sending.")
