from __future__ import annotations
# ============================================================
# email_utils.py — SendGrid HTML email builder + sender
# ============================================================

import urllib.request
import urllib.error
import json
from datetime import date, timedelta
import streamlit as st
from config import SENDER_EMAIL, SENDER_NAMES, _rank


def _get_credentials() -> tuple[str, str]:
    try:
        return st.secrets["email"]["sendgrid_api_key"], st.secrets["email"]["from_email"]
    except (KeyError, FileNotFoundError):
        return "", ""


def _get_sender_name() -> str:
    try:
        from_email = st.secrets["email"]["from_email"]
    except (KeyError, FileNotFoundError):
        from_email = SENDER_EMAIL
    return SENDER_NAMES.get(from_email,
           from_email.split("@")[0].split(".")[0].capitalize())


def _next_monday() -> str:
    """Return the coming Monday as a readable string, e.g. 'Monday, May 19'."""
    today      = date.today()
    days_ahead = (7 - today.weekday()) % 7
    if days_ahead == 0:
        days_ahead = 7
    monday = today + timedelta(days=days_ahead)
    # %-d is Linux-only; use lstrip for portability
    day = monday.strftime("%B %d").replace(" 0", " ")
    return f"Monday, {day}"


def email_configured() -> bool:
    api_key, _ = _get_credentials()
    return bool(api_key)


def send_email(to_email: str, subject: str, body: str,
               cc_email: str = None) -> dict:
    api_key, from_email = _get_credentials()
    if not api_key:
        return {"status": "error: SendGrid credentials not configured", "to": to_email}

    content_type = "text/html" if body.strip().startswith("<") else "text/plain"

    personalization = {"to": [{"email": to_email}]}
    if cc_email and cc_email != to_email:
        personalization["cc"] = [{"email": cc_email}]

    payload = {
        "personalizations": [personalization],
        "from":    {"email": from_email},
        "subject": subject,
        "content": [{"type": content_type, "value": body}],
    }

    data = json.dumps(payload).encode("utf-8")
    req  = urllib.request.Request(
        "https://api.sendgrid.com/v3/mail/send",
        data=data,
        headers={"Authorization": f"Bearer {api_key}",
                 "Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req) as resp:
            return {"status": "sent" if resp.status == 202
                    else f"unexpected {resp.status}", "to": to_email}
    except urllib.error.HTTPError as e:
        body_err = e.read().decode("utf-8", errors="ignore")
        return {"status": f"error {e.code}: {body_err[:200]}", "to": to_email}
    except Exception as e:
        return {"status": f"error: {str(e)}", "to": to_email}


# ---- HTML STYLES ----
_CSS = """
<style>
  body { font-family: Calibri, Arial, sans-serif; font-size: 14px; color: #1a1a1a; }
  h3   { color: #0E2841; margin-top: 24px; margin-bottom: 6px; font-size: 15px; }
  p    { margin: 4px 0 10px 0; }
  .notice { background:#fff3cd; border-left:4px solid #e6a817;
            padding:8px 12px; margin:10px 0; font-size:13px; border-radius:3px; }
  .info   { background:#e8f4fd; border-left:4px solid #2980b9;
            padding:8px 12px; margin:10px 0; font-size:13px; border-radius:3px; }
  table { border-collapse: collapse; width: 100%; margin-bottom: 18px; font-size: 13px; }
  th    { background-color: #0E2841; color: #fff; padding: 8px 12px;
          text-align: left; font-weight: 600; }
  td    { padding: 7px 12px; border-bottom: 1px solid #dce3ea; vertical-align: top; }
  tr:nth-child(even) td { background-color: #f5f7fa; }
  .response-col { background-color: #fffbe6 !important; min-width: 180px; }
  .neg  { color: #c0392b; font-weight: 600; }
  .pos  { color: #1a6b2f; }
  .over { color: #e67e22; font-weight: 600; }
  .sig  { font-size: 13px; color: #444; margin-top: 28px; }
</style>
"""


def _table(headers: list, rows: list, response_col: bool = True) -> str:
    all_headers = headers + (["Response"] if response_col else [])
    th_html = "".join(f"<th>{h}</th>" for h in all_headers)
    html = f"<table><thead><tr>{th_html}</tr></thead><tbody>"
    for row in rows:
        cells = "".join(f"<td>{c}</td>" for c in row)
        if response_col:
            cells += '<td class="response-col">&nbsp;</td>'
        html += f"<tr>{cells}</tr>"
    html += "</tbody></table>"
    return html


def build_html_email(
    owner: str,
    first_name: str,
    tracker_issues: list,
    budget_issues: list,
    tbd_projects: list,
    variance_issues: list,
    util_data: list        = None,
    pto_schedule: dict     = None,
    pto_months: list       = None,
    has_openair: bool      = False,
    no_openair_note: bool  = False,
    selected_months: list  = None,
    is_staff: bool         = False,
) -> str:
    """Build one combined HTML email for an owner."""
    sender_name = _get_sender_name()
    sections    = []

    # Multi-period notice
    if selected_months and len(selected_months) > 1:
        month_list = ", ".join(selected_months)
        sections.append(
            f'<div class="notice">⚠️ <strong>Note:</strong> This report covers '
            f'multiple periods: <strong>{month_list}</strong>.</div>'
        )

    # ── Project Tracker ──────────────────────────────────────
    if tracker_issues:
        rows = []
        for i in tracker_issues:
            missing = i.get("missing_rates", [])
            for m in missing:
                rows.append([i["project_code"], f"Missing {m} Rate"])
            # If missing_rates is empty but problems exists
            if not missing:
                for prob in i.get("problems", []):
                    rows.append([i["project_code"], prob])
        if rows:
            sections.append(
                "<h3>Project Tracker</h3>"
                "<p>The below projects have missing rates and/or budget. "
                "Please review and update.</p>"
                + _table(["Project Code", "To be reviewed"], rows)
            )

    # ── Budget to Actual ─────────────────────────────────────
    if budget_issues:
        rows = []
        for i in budget_issues:
            css  = "neg" if i.get("type") == "negative" else "pos"
            rows.append([i["project_code"],
                         f'<span class="{css}">{i.get("description", "")}</span>'])
        sections.append(
            "<h3>Budget to Actual</h3>"
            "<p>The below projects have either a negative remaining budget "
            "or an unscheduled budget over the threshold. Please review and advise.</p>"
            + _table(["Project Code", "To be reviewed"], rows)
        )

    # ── TBD / Pending SOW ────────────────────────────────────
    owner_tbd = [p for p in (tbd_projects or []) if p.get("owner") == owner]
    if owner_tbd:
        rows = [[p.get("project_code", ""), p.get("status", "TBD")]
                for p in owner_tbd]
        sections.append(
            "<h3>TBD / Pending SOW Projects</h3>"
            "<p>Please review the below projects with TBD budgets and provide any applicable updates.</p>"
            + _table(["Project Code", "Status"], rows, response_col=True)
        )

    # ── Variance ─────────────────────────────────────────────
    if variance_issues:
        if no_openair_note:
            sections.append(
                '<div class="info">ℹ️ <strong>Note:</strong> No OpenAir report was uploaded, '
                'so actual hours are shown as 0. Schedule hours reflect what is planned.</div>'
            )
        # Sort: owner's own rows first, then others by seniority rank then alpha
        _sorted_var = sorted(
            variance_issues,
            key=lambda v: (
                0 if v.get("person", "") == owner else 1,
                _rank(v.get("person", "")),
                v.get("project_code", ""),
            )
        )
        rows = []
        for v in _sorted_var:
            diff     = v.get("difference", 0)
            css      = "over" if diff < 0 else "neg"
            diff_fmt = f'<span class="{css}">{diff:+.1f}</span>'
            if is_staff:
                rows.append([
                    v.get("project_code", ""), v.get("period", ""),
                    str(v.get("actual_hours", "")), str(v.get("sched_hours", "")),
                    diff_fmt, v.get("question", ""),
                ])
            else:
                rows.append([
                    v.get("person", ""), v.get("project_code", ""), v.get("period", ""),
                    str(v.get("actual_hours", "")), str(v.get("sched_hours", "")),
                    diff_fmt, v.get("question", ""),
                ])
        var_headers = (
            ["Project Code", "Period", "Actual Hrs", "Scheduled Hrs", "Difference", "To be reviewed"]
            if is_staff else
            ["Person", "Project Code", "Period", "Actual Hrs", "Scheduled Hrs", "Difference", "To be reviewed"]
        )
        sections.append(
            "<h3>Actual vs Schedule Variances</h3>"
            "<p>Please review the below variances and provide updates as needed.</p>"
            + _table(var_headers, rows)
        )

    # ── Utilization ──────────────────────────────────────────
    if util_data:
        person_util = [u for u in util_data if u.get("person") == owner]
        if person_util:
            u = person_util[0]
            util_pct = u.get("utilization_pct")
            goal_pct = u.get("goal_pct")
            diff_pct = u.get("difference_pct")
            util_str = f"{util_pct:.1f}%" if util_pct is not None else "-"
            goal_str = f"{goal_pct:.0f}%" if goal_pct is not None else "-"
            if diff_pct is not None:
                css      = "neg" if diff_pct > 10 else ("over" if diff_pct < -10 else "pos")
                diff_str = f'<span class="{css}">{diff_pct:+.1f}%</span>'
            else:
                diff_str = "-"

            # Question based on difference
            if diff_pct is not None and diff_pct < -10:
                util_question = (
                    "What do you plan to do with your non-charge time? "
                    "Are there any projects you know of that aren't in the schedule yet?"
                )
            elif diff_pct is not None and diff_pct > 10:
                util_question = (
                    "Is there any project work you could use assistance with, "
                    "or places where we can shift hours?"
                )
            else:
                util_question = ""  # Within 10% — informational only, no action needed

            rows = [[util_str, goal_str, diff_str,
                     str(u.get("chargeable", "-")), str(u.get("remaining", "-")),
                     util_question]]
            sections.append(
                "<h3>Utilization</h3>"
"<p>Please see below projected utilization for the month.</p>"
                + _table(["Utilization", "Goal", "Difference", "Chargeable Hrs", "Remaining Hrs",
                           "To be reviewed"],
                          rows, response_col=True)
            )

    # ── PTO Schedule ─────────────────────────────────────────
    if pto_schedule and pto_months:
        person_pto = pto_schedule.get(owner, {})
        # Only include PTO table if person has at least one non-zero month
        has_any_pto = any(person_pto.get(m, 0) for m in pto_months)
        if person_pto and has_any_pto:
            months_to_show = [m for m in pto_months if m in person_pto]
            if months_to_show:
                rows = [
                    [m, f"{int(person_pto[m])} hrs" if person_pto.get(m) else "—"]
                    for m in months_to_show
                ]
                sections.append(
                    "<h3>Scheduled PTO</h3>"
                    "<p>Your scheduled PTO for the current and next two months is shown below. "
                    "Please reply if any updates are needed.</p>"
                    + _table(["Month", "PTO Hours"], rows, response_col=False)
                )

    if not sections:
        return ""

    deadline = _next_monday()
    return f"""<html><head>{_CSS}</head><body>
<p>Hi {first_name},</p>
<p>Please review the items below and reply to <strong>Laren</strong> by <strong>{deadline} at 10:00 AM</strong>.</p>
{"".join(sections)}
<p class="sig">Best,<br>{sender_name}</p>
</body></html>"""


# ── Aliases ───────────────────────────────────────────────────
EMAIL_OK = email_configured()


def send_emails_batch(emails: list) -> list:
    """
    Send a list of email dicts with keys: to, subject, body.
    Returns list of result dicts with keys: to, status.
    """
    results = []
    for email in emails:
        result = send_email(
            to_email=email["to"],
            subject=email["subject"],
            body=email["body"],
        )
        results.append(result)
    return results
