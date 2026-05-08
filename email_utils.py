from __future__ import annotations
# ============================================================
# email_utils.py — Email sending via SMTP (Microsoft 365)
# Works on Streamlit Cloud and locally.
#
# Credentials are stored in Streamlit secrets (never in code):
#   .streamlit/secrets.toml  (local)
#   Streamlit Cloud > App Settings > Secrets  (cloud)
#
# Required secrets format:
#   [email]
#   smtp_user     = "laren@gtmtax.com"
#   smtp_password = "your-password-here"
# ============================================================

import smtplib
import streamlit as st
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from config import LAREN_EMAIL

SMTP_SERVER = "smtp.office365.com"
SMTP_PORT   = 587


def _get_smtp_credentials() -> tuple[str, str] | tuple[None, None]:
    """Pull SMTP credentials from Streamlit secrets."""
    try:
        user     = st.secrets["email"]["smtp_user"]
        password = st.secrets["email"]["smtp_password"]
        return user, password
    except (KeyError, FileNotFoundError):
        return None, None


def send_email(to_email: str, subject: str, html_body: str,
               cc_email: str = LAREN_EMAIL) -> dict:
    """Send a single HTML email via Microsoft 365 SMTP."""
    smtp_user, smtp_password = _get_smtp_credentials()
    if not smtp_user or not smtp_password:
        return {"status": "error: email credentials not configured in secrets",
                "to": to_email}
    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"]    = smtp_user
        msg["To"]      = to_email
        msg["CC"]      = cc_email
        msg.attach(MIMEText(html_body, "html"))

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.ehlo()
            server.starttls()
            server.login(smtp_user, smtp_password)
            recipients = [to_email]
            if cc_email and cc_email != to_email:
                recipients.append(cc_email)
            server.sendmail(smtp_user, recipients, msg.as_string())

        return {"status": "sent", "to": to_email}
    except smtplib.SMTPAuthenticationError:
        return {"status": "error: authentication failed — check credentials in secrets",
                "to": to_email}
    except Exception as e:
        return {"status": f"error: {str(e)}", "to": to_email}


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
<p class="sig">Best,<br>Laren</p>
</body></html>"""


def build_and_send_combined_emails(owners_data: dict,
                                   cc_email: str = LAREN_EMAIL) -> list:
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
