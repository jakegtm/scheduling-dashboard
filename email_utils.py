from __future__ import annotations
# ============================================================
# email_utils.py — SendGrid HTML email builder + sender
# ============================================================

import urllib.request
import urllib.error
import json
import streamlit as st
from config import SENDER_EMAIL, SENDER_NAMES


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


def email_configured() -> bool:
    api_key, _ = _get_credentials()
    return bool(api_key)


def send_email(to_email: str, subject: str, html_body: str,
               cc_email: str = SENDER_EMAIL) -> dict:
    api_key, from_email = _get_credentials()
    if not api_key:
        return {"status": "error: SendGrid credentials not configured", "to": to_email}

    payload = {
        "personalizations": [{
            "to": [{"email": to_email}],
            "cc": [{"email": cc_email}] if cc_email and cc_email != to_email else [],
        }],
        "from":    {"email": from_email},
        "subject": subject,
        "content": [{"type": "text/html", "value": html_body}],
    }
    if not payload["personalizations"][0]["cc"]:
        del payload["personalizations"][0]["cc"]

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
        body = e.read().decode("utf-8", errors="ignore")
        return {"status": f"error {e.code}: {body[:200]}", "to": to_email}
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


def build_html_email(owner: str,
                     first_name: str,
                     tracker_issues: list,
                     budget_issues: list,
                     month_issues: list,
                     variance_issues: list,
                     has_openair: bool = False,
                     selected_months: list = None) -> str:
    """Build one combined HTML email for an owner."""
    sender_name = _get_sender_name()
    sections    = []

    # ---- Multi-month notice ----
    month_notice = ""
    if selected_months and len(selected_months) > 1:
        month_list = ", ".join(selected_months)
        month_notice = (f'<div class="notice">⚠️ <strong>Note:</strong> This report '
                        f'covers multiple periods: <strong>{month_list}</strong>. '
                        f'Please review all periods listed below.</div>')

    # ---- Project Tracker ----
    if tracker_issues:
        rows = [[i["project_code"], prob]
                for i in tracker_issues for prob in i["problems"]]
        sections.append(
            "<h3>Project Tracker</h3>"
            "<p>The below projects have missing rates and/or budget. "
            "Please review and update.</p>"
            + _table(["Project Code", "To be reviewed"], rows)
        )

    # ---- Budget to Actual ----
    if budget_issues:
        rows = []
        for i in budget_issues:
            css  = "neg" if i["type"] == "negative" else "pos"
            rows.append([i["project_code"],
                         f'<span class="{css}">{i["description"]}</span>'])
        sections.append(
            "<h3>Budget to Actual</h3>"
            "<p>The below projects have either a negative remaining budget "
            "or an unscheduled budget over $20,000. Please review and advise.</p>"
            + _table(["Project Code", "To be reviewed"], rows)
        )

    # ---- Scheduled Hours (no OpenAir needed) ----
    if month_issues:
        rows = [[str(i["project_code"]), str(i["period"]), str(i["hours"])]
                for i in month_issues]
        sections.append(
            "<h3>Scheduled Hours — Action Required</h3>"
            "<p>The following hours are scheduled for you but have not yet been "
            "confirmed. Please confirm, update, or advise on any changes needed "
            "before the period deadline.</p>"
            + _table(["Project Code", "Period", "Scheduled Hours"], rows)
        )

    # ---- Variance (OpenAir required) ----
    if has_openair and variance_issues:
        period_note = ""
        if selected_months and len(selected_months) > 1:
            period_note = (f" for the periods: <strong>"
                           f"{', '.join(selected_months)}</strong>")
        rows = []
        for v in variance_issues:
            diff     = v["difference"]
            css      = "over" if diff < 0 else "neg"
            diff_fmt = f'<span class="{css}">{diff:+.1f}</span>'
            rows.append([v["project_code"], v["period"],
                         str(v["actual_hours"]), str(v["sched_hours"]),
                         diff_fmt, v["question"]])
        sections.append(
            f"<h3>Actual vs Schedule Variances</h3>"
            f"<p>Please review the below variances{period_note} and "
            f"provide schedule updates as needed.</p>"
            + _table(["Project Code", "Period", "Actual Hrs",
                       "Scheduled Hrs", "Difference", "To be reviewed"], rows)
        )

    if not sections:
        return ""

    return f"""<html><head>{_CSS}</head><body>
<p>Hi {first_name},</p>
{month_notice}
<p>Please review the items below and provide your input directly in the yellow
response fields. Once complete, simply reply to this email and your responses
will be on their way. We appreciate your time and prompt attention to these items.</p>
{"".join(sections)}
<p class="sig">Best,<br>{sender_name}</p>
</body></html>"""


def build_and_send_combined_emails(owners_data: dict,
                                   cc_email: str = SENDER_EMAIL,
                                   has_openair: bool = False,
                                   selected_months: list = None,
                                   selected_owners: list = None) -> list:
    """
    Send one combined email per owner.
    selected_owners: if provided, only send to those owners.
    """
    results = []
    for owner, data in owners_data.items():
        if selected_owners and owner not in selected_owners:
            continue
        owner_email = data.get("email")
        if not owner_email:
            results.append({"to": owner, "status": "skipped: no email", "subject": ""})
            continue
        if not any([data.get("tracker"), data.get("budget"),
                    data.get("month"), data.get("variance")]):
            continue

        first_name = data.get("first_name", owner.split()[0] if owner else "there")
        html = build_html_email(
            owner, first_name,
            data.get("tracker", []), data.get("budget", []),
            data.get("month", []), data.get("variance", []),
            has_openair=has_openair, selected_months=selected_months,
        )
        if not html:
            continue

        subject = "Scheduling Review — Action Required"
        result  = send_email(owner_email, subject, html, cc_email)
        result["subject"] = subject
        results.append(result)

    return results
