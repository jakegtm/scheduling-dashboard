from __future__ import annotations
# ============================================================
# email_utils.py — Email sending via SendGrid
# Works on Streamlit Cloud and locally.
#
# Required Streamlit secrets:
#   [email]
#   sendgrid_api_key = "SG.xxxxxxxxxx"
#   from_email       = "jodonnell@gtmtax.com"
# ============================================================

import urllib.request
import urllib.error
import json
import streamlit as st
from config import SENDER_EMAIL


def _get_credentials() -> tuple[str, str]:
    """Pull SendGrid credentials from Streamlit secrets."""
    try:
        api_key    = st.secrets["email"]["sendgrid_api_key"]
        from_email = st.secrets["email"]["from_email"]
        return api_key, from_email
    except (KeyError, FileNotFoundError):
        return "", ""


def send_email(to_email: str, subject: str, html_body: str,
               cc_email: str = SENDER_EMAIL) -> dict:
    """Send a single HTML email via SendGrid API."""
    api_key, from_email = _get_credentials()
    if not api_key:
        return {"status": "error: SendGrid credentials not configured in Streamlit secrets",
                "to": to_email}

    payload = {
        "personalizations": [{
            "to":  [{"email": to_email}],
            "cc":  [{"email": cc_email}] if cc_email and cc_email != to_email else [],
        }],
        "from":    {"email": from_email},
        "subject": subject,
        "content": [{"type": "text/html", "value": html_body}],
    }

    # Remove empty cc list
    if not payload["personalizations"][0]["cc"]:
        del payload["personalizations"][0]["cc"]

    data = json.dumps(payload).encode("utf-8")
    req  = urllib.request.Request(
        "https://api.sendgrid.com/v3/mail/send",
        data=data,
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type":  "application/json",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req) as resp:
            if resp.status == 202:
                return {"status": "sent", "to": to_email}
            return {"status": f"unexpected status {resp.status}", "to": to_email}
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="ignore")
        return {"status": f"error {e.code}: {body[:200]}", "to": to_email}
    except Exception as e:
        return {"status": f"error: {str(e)}", "to": to_email}


def email_configured() -> bool:
    """Return True if SendGrid credentials are present."""
    api_key, _ = _get_credentials()
    return bool(api_key)


# ---- HTML STYLES ----
_CSS = """
<style>
  body { font-family: Calibri, Arial, sans-serif; font-size: 14px; color: #1a1a1a; }
  h3   { color: #0E2841; margin-top: 24px; margin-bottom: 6px; font-size: 15px; }
  p    { margin: 4px 0 10px 0; }
  table {
    border-collapse: collapse;
    width: 100%;
    margin-bottom: 18px;
    font-size: 13px;
  }
  th {
    background-color: #0E2841;
    color: #ffffff;
    padding: 8px 12px;
    text-align: left;
    font-weight: 600;
  }
  td {
    padding: 7px 12px;
    border-bottom: 1px solid #dce3ea;
    vertical-align: top;
  }
  tr:nth-child(even) td { background-color: #f5f7fa; }
  .response-col {
    background-color: #fffbe6 !important;
    min-width: 200px;
    color: #888;
    font-style: italic;
  }
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
            cells += '<td class="response-col">Type your response here…</td>'
        html += f"<tr>{cells}</tr>"
    html += "</tbody></table>"
    return html


def build_html_email(owner: str, tracker_issues: list,
                     budget_issues: list, variance_issues: list) -> str:
    first_name = owner.split()[0] if owner else "there"
    sections   = []

    if tracker_issues:
        rows = []
        for issue in tracker_issues:
            for prob in issue["problems"]:
                rows.append([issue["project_code"], prob])
        sections.append(
            "<h3>Project Tracker</h3>"
            "<p>The below table shows projects in which you are the owner "
            "that have missing rates and/or budget. Please review and update.</p>"
            + _table(["Project Code", "To be reviewed"], rows)
        )

    if budget_issues:
        rows = []
        for issue in budget_issues:
            css  = "neg" if issue["type"] == "negative" else "pos"
            desc = f'<span class="{css}">{issue["description"]}</span>'
            rows.append([issue["project_code"], desc])
        sections.append(
            "<h3>Budget to Actual</h3>"
            "<p>The below table shows projects in which you are the owner "
            "that have either a negative remaining budget or an unscheduled "
            "budget over $20,000. Please review and advise.</p>"
            + _table(["Project Code", "To be reviewed"], rows)
        )

    if variance_issues:
        rows = []
        for v in variance_issues:
            diff     = v["difference"]
            css      = "over" if diff < 0 else "neg"
            diff_fmt = f'<span class="{css}">{diff:+.1f}</span>'
            rows.append([
                v["project_code"],
                str(v["actual_hours"]),
                str(v["sched_hours"]),
                diff_fmt,
                v["question"],
            ])
        sections.append(
            "<h3>Actual to Schedule Variances</h3>"
            "<p>Please review the below actual to schedule variances "
            "and provide schedule updates as needed.</p>"
            + _table(["Project Code", "Actual Hours", "Schedule Hours",
                      "Difference", "To be reviewed"], rows)
        )

    return f"""<html><head>{_CSS}</head><body>
<p>Hi {first_name},</p>
<p>Please review the following items and provide input for the schedule.</p>
{"".join(sections)}
<p class="sig">Best,<br>Scheduling Team</p>
</body></html>"""


def build_and_send_combined_emails(owners_data: dict,
                                   cc_email: str = SENDER_EMAIL) -> list:
    """One combined HTML email per owner. Returns list of result dicts."""
    results = []
    for owner, data in owners_data.items():
        owner_email = data.get("email")
        if not owner_email:
            results.append({"to": owner, "status": "skipped: no email", "subject": ""})
            continue
        if not data.get("tracker") and not data.get("budget") and not data.get("variance"):
            continue
        html    = build_html_email(owner, data.get("tracker", []),
                                   data.get("budget", []), data.get("variance", []))
        subject = "Scheduling Review — Action Required"
        result  = send_email(owner_email, subject, html, cc_email)
        result["subject"] = subject
        results.append(result)
    return results
