from __future__ import annotations
# ============================================================
# processors/project_tracker.py
# ============================================================
# Column layout (verified against real file):
#   A(1):  Client
#   B(2):  Project Code
#   C(3):  Status         <- filter to Known / TBD / Pending SOW
#   E(5):  Notes
#   I(9):  Project Owner  <- email target (H is Client Owner, ignored)
#   J(10): 2026 Budget
#   K(11): Reclass to 2027
#   L-S (12-19): Rates (Intern → Managing Director)
#   T(20): Write Up / (Down)
#
# Columns are located by their row-1 header text, so inserting new
# columns won't break the reader. The numbers above are only fallbacks.
#
# Rules:
#   - TBD or Pending SOW (status only) → collect in TBD list
#   - Known non-TBD → check for missing rates; flag if any are blank
#   - All other statuses (Unknown, Closed, blank) → skip
# ============================================================

from collections import defaultdict
from config import EMAIL_LOOKUP, FIRST_NAMES

COL_CLIENT      = 1   # A
COL_CODE        = 2   # B
COL_STATUS      = 3   # C
COL_NOTES       = 5   # E
COL_OWNER       = 9   # I — Project Owner
COL_BUDGET      = 10  # J
COL_RECLASS     = 11  # K — Reclass to 2027
COL_RATES_START = 12  # L — Intern
COL_RATES_END   = 19  # S — Managing Director

RATE_LABELS = [
    "Intern", "Analyst", "Senior Analyst", "Supervisor",
    "Manager", "Senior Manager", "Director", "Managing Director",
]

# Statuses that are treated as TBD (excluded from rate checks,
# included in the TBD/Pending SOW email section)
TBD_STATUSES = {"tbd", "pending sow"}


def _to_float(val):
    try:
        return float(val) if val is not None else None
    except (ValueError, TypeError):
        return None


def _lookup_email(name: str):
    if not name:
        return None
    name = str(name).strip()
    if name in EMAIL_LOOKUP:
        return EMAIL_LOOKUP[name]
    for key, email in EMAIL_LOOKUP.items():
        if key.lower() == name.lower():
            return email
    return None


def _lookup_first(name: str) -> str:
    if not name:
        return "there"
    return FIRST_NAMES.get(name.strip(), name.strip())


def _norm(h) -> str:
    return " ".join(str(h).lower().split()) if h is not None else ""


def _resolve_columns(ws) -> dict:
    """Find column positions from the row-1 headers, falling back to the
    fixed defaults for anything not found."""
    headers = {}
    for cell in next(ws.iter_rows(min_row=1, max_row=1)):
        h = _norm(cell.value)
        if h and h not in headers:
            headers[h] = cell.column

    def find(*names, default):
        for n in names:
            if n in headers:
                return headers[n]
        for n in names:                      # loose "contains" match
            for h, c in headers.items():
                if n in h:
                    return c
        return default

    cols = {
        "client":  find("client", default=COL_CLIENT),
        "code":    find("project code", default=COL_CODE),
        "status":  find("status", default=COL_STATUS),
        "notes":   find("notes", default=COL_NOTES),
        "owner":   find("project owner", default=COL_OWNER),
        "budget":  find("2026 budget", "budget", default=COL_BUDGET),
        "reclass": find("reclass to 2027", "reclass", default=COL_RECLASS),
    }
    rate_cols = []
    for i, label in enumerate(RATE_LABELS):
        c = headers.get(label.lower())
        rate_cols.append(c if c else COL_RATES_START + i)
    cols["rates"] = rate_cols
    return cols


def _cell(row, col):
    return row[col - 1] if col and col - 1 < len(row) else None


def get_reclass_projects(ws) -> list:
    """
    Projects with a non-zero 'Reclass to 2027' amount, for the project
    owner's report. Closed projects are skipped.
    """
    cols = _resolve_columns(ws)
    out  = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        row = list(row)
        client = _cell(row, cols["client"])
        code   = _cell(row, cols["code"])
        if not client and not code:
            continue
        status = str(_cell(row, cols["status"]) or "").strip()
        if status.lower() == "closed":
            continue
        reclass = _to_float(_cell(row, cols["reclass"]))
        if not reclass:
            continue
        owner = str(_cell(row, cols["owner"]) or "").strip()
        out.append({
            "client":       client,
            "project_code": str(code).strip() if code else "",
            "status":       status,
            "owner":        owner,
            "owner_email":  _lookup_email(owner),
            "owner_first":  _lookup_first(owner),
            "budget":       _to_float(_cell(row, cols["budget"])) or 0.0,
            "reclass":      reclass,
        })
    return out


def process_project_tracker(ws) -> tuple:
    """
    Returns:
        issues            — Known projects with any blank rate in K:R
        tbd_projects      — TBD and Pending SOW projects (for reference / email section)
        project_owner_map — dict of {project_code: owner} for ALL projects (used for variance routing)
    """
    issues            = []
    tbd_projects      = []
    project_owner_map = {}

    cols = _resolve_columns(ws)

    for row in ws.iter_rows(min_row=2, values_only=True):
        row = list(row)

        client       = _cell(row, cols["client"])
        project_code = _cell(row, cols["code"])
        status       = _cell(row, cols["status"])
        owner        = _cell(row, cols["owner"])
        budget       = _to_float(_cell(row, cols["budget"]))

        if not client and not project_code:
            continue

        code_str   = str(project_code).strip() if project_code else ""
        status_str = str(status).strip()        if status       else ""
        owner_str  = str(owner).strip()         if owner        else ""

        # Always populate the owner map regardless of status
        if code_str and owner_str:
            project_owner_map[code_str] = owner_str

        # --- TBD / Pending SOW --- filter by status only
        if status_str.lower() in TBD_STATUSES:
            notes_val = _cell(row, cols["notes"])
            tbd_projects.append({
                "client":       client,
                "project_code": code_str,
                "status":       status_str,
                "owner":        owner_str,
                "owner_email":  _lookup_email(owner_str),
                "owner_first":  _lookup_first(owner_str),
                "budget":       budget or 0.0,
                "notes":        str(notes_val).strip() if notes_val else "",
            })
            continue

        # --- Known only ---
        if status_str.lower() != "known":
            continue

        # Check for any blank rate (Intern → Managing Director)
        rate_values = [_cell(row, c) for c in cols["rates"]]
        missing_labels = [
            label for label, val in zip(RATE_LABELS, rate_values)
            if val is None or str(val).strip() == ""
        ]

        if missing_labels:
            issues.append({
                "client":        client,
                "project_code":  code_str,
                "owner":         owner_str,
                "owner_email":   _lookup_email(owner_str),
                "owner_first":   _lookup_first(owner_str),
                "budget":        budget or 0.0,
                "missing_rates": missing_labels,
                "problems":      [f"Missing rate(s): {', '.join(missing_labels)}"],
            })

    return issues, tbd_projects, project_owner_map


def build_tracker_emails(issues: list, tbd_projects: list,
                         sender_name: str = "Jake") -> list:
    """
    Build one email per project owner covering:
      - Missing rates section (from issues)
      - TBD / Pending SOW section (from tbd_projects)
    """
    owner_issues = defaultdict(list)
    for issue in issues:
        owner_issues[issue["owner"]].append(issue)

    owner_tbd = defaultdict(list)
    for proj in tbd_projects:
        owner_tbd[proj["owner"]].append(proj)

    all_owners = set(owner_issues) | set(owner_tbd)
    emails = []

    for owner in all_owners:
        email = None
        if owner_issues[owner]:
            email = owner_issues[owner][0].get("owner_email")
        if not email and owner_tbd[owner]:
            email = owner_tbd[owner][0].get("owner_email")
        if not email:
            continue

        first = _lookup_first(owner)
        sections = []

        if owner_issues[owner]:
            lines = [
                "The following projects assigned to you are missing billing rates. "
                "Please review and update as soon as possible.\n"
            ]
            for issue in owner_issues[owner]:
                lines.append(
                    f"  • {issue['client']} — {issue['project_code']}\n"
                    f"    Missing: {', '.join(issue['missing_rates'])}"
                )
            sections.append("\n".join(lines))

        if owner_tbd[owner]:
            lines = [
                "The following projects currently have TBD or Pending SOW budgets. "
                "If you have any updates on these, please reply with the latest — "
                "otherwise, no action is needed.\n"
            ]
            for proj in owner_tbd[owner]:
                status_label = f" [{proj['status']}]" if proj["status"] else " [TBD]"
                lines.append(
                    f"  • {proj['client']} — {proj['project_code']}{status_label}"
                )
            sections.append("\n".join(lines))

        body = (
            f"Hi {first},\n\n"
            + "\n\n".join(sections)
            + f"\n\nBest,\n{sender_name}"
        )

        emails.append({
            "to":      email,
            "subject": "Project Tracker — Review Required",
            "owner":   owner,
            "body":    body,
            "section": "tracker",
        })

    return emails
