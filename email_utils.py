# ============================================================
# email_utils.py — HTML table emails via Outlook win32com
# ============================================================

from config import LAREN_EMAIL


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
    """Build an HTML table. If response_col=True, appends an empty Response column."""
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


def build_html_email(owner: str,
                     tracker_issues: list,
                     budget_issues: list,
                     variance_issues: list) -> str:
    """
    Build a complete HTML email body for one project owner.
    Any of the three issue lists may be empty — that section is omitted.
    """
    first_name = owner.split()[0] if owner else "there"
    has_tracker  = bool(tracker_issues)
    has_budget   = bool(budget_issues)
    has_variance = bool(variance_issues)

    sections = []

    # ---- Project Tracker section ----
    if has_tracker:
        rows = []
        for issue in tracker_issues:
            for problem in issue["problems"]:
                rows.append([issue["project_code"], problem])
        sections.append(
            "<h3>Project Tracker</h3>"
            "<p>The below table shows projects in which you are the owner "
            "that have missing rates and/or budget. Please review and update.</p>"
            + _table(["Project Code", "To be reviewed"], rows)
        )

    # ---- Budget to Actual section ----
    if has_budget:
        rows = []
        for issue in budget_issues:
            css_class = "neg" if issue["type"] == "negative" else "pos"
            desc = f'<span class="{css_class}">{issue["description"]}</span>'
            rows.append([issue["project_code"], desc])
        sections.append(
            "<h3>Budget to Actual</h3>"
            "<p>The below table shows projects in which you are the owner "
            "that have either a negative remaining budget or an unscheduled "
            "budget over $20,000. Please review and advise.</p>"
            + _table(["Project Code", "To be reviewed"], rows)
        )

    # ---- Variance section ----
    if has_variance:
        rows = []
        for v in variance_issues:
            diff = v["difference"]
            diff_fmt = f'<span class="{"over" if diff < 0 else "neg"}">{diff:+.1f}</span>'
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
            + _table(
                ["Project Code", "Actual Hours", "Schedule Hours",
                 "Difference", "To be reviewed"],
                rows,
            )
        )

    body = f"""<html><head>{_CSS}</head><body>
<p>Hi {first_name},</p>
<p>Please review the following items and provide input for the schedule.</p>
{"".join(sections)}
<p class="sig">Best,<br>Laren</p>
</body></html>"""

    return body


def send_outlook_email(to_email: str, subject: str, html_body: str,
                       cc_email: str = LAREN_EMAIL) -> dict:
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To       = to_email
        mail.CC       = cc_email
        mail.Subject  = subject
        mail.HTMLBody = html_body   # rich HTML, not plain text
        mail.Send()
        return {"status": "sent", "to": to_email}
    except ImportError:
        return {"status": "error: pywin32 not installed", "to": to_email}
    except Exception as e:
        return {"status": f"error: {str(e)}", "to": to_email}


def build_and_send_combined_emails(owners_data: dict, cc_email: str = LAREN_EMAIL) -> list:
    """
    Build one combined HTML email per owner covering all three sections.

    owners_data: {
        "Hendrickson": {
            "email":     "hendrickson@gtmtax.com",
            "tracker":   [...],   # from process_project_tracker
            "budget":    [...],   # from process_budget_actual
            "variance":  [...],   # from compute_variances
        }
    }

    Returns list of result dicts {to, status, subject}.
    """
    results = []
    for owner, data in owners_data.items():
        owner_email = data.get("email")
        if not owner_email:
            results.append({"to": owner, "status": "skipped: no email", "subject": ""})
            continue

        tracker_issues  = data.get("tracker", [])
        budget_issues   = data.get("budget", [])
        variance_issues = data.get("variance", [])

        if not tracker_issues and not budget_issues and not variance_issues:
            continue

        html_body = build_html_email(
            owner, tracker_issues, budget_issues, variance_issues
        )
        subject = "Scheduling Review — Action Required"
        result  = send_outlook_email(owner_email, subject, html_body, cc_email)
        result["subject"] = subject
        results.append(result)

    return results
