# ============================================================
# config.py — GTM Scheduling Analyzer Configuration
# ============================================================

# --- EMAIL SENDER ---
SENDER_EMAIL = "jodonnell@gtmtax.com"   # fallback if not in secrets

SENDER_NAMES = {
    "jodonnell@gtmtax.com":    "Jake",
    "lhendrickson@gtmtax.com": "Laren",
    # Add other senders here as needed
}

# --- THRESHOLDS (sidebar defaults) ---
DEFAULT_BUDGET_THRESHOLD         = 20000   # Flag unscheduled remaining >= this
DEFAULT_PROJECTION_THRESHOLD_PCT = 0.80    # Flag if remaining/budget >= this
DEFAULT_NEGATIVE_THRESHOLD       = 100     # Flag remaining <= -this

# Variance (Actual vs Schedule) — min/max difference to flag
DEFAULT_VARIANCE_MIN = 5    # Hours under-worked threshold (displayed as positive, negated in code)
DEFAULT_VARIANCE_MAX = 0    # Hours over-worked threshold (flag if actual exceeds scheduled by more than this)

# --- ROLE RULES ---
# No interns on the team — everyone is staff or a project owner.
# Retained as an empty set so the import in app.py keeps working; the
# separate intern filtering it drove was redundant with STAFF_NAMES
# (staff never accumulate other people's rows) and has been removed.
INTERN_NAMES: set = set()

# Staff who only receive their own variance rows (not all project rows).
# Project owners (managers/directors) are everyone NOT in this set.
STAFF_NAMES = {"Avington", "S. O'Donnell", "J. O'Donnell", "McGrogan"}

# Project code prefixes to exclude from variance analysis.
# GTM internal/non-chargeable codes (NONCHG, TRAINING, HOLIDAYS, etc.) should not appear.
# The Non-Charge tab uses this same set in reverse — it shows exactly these projects.
VARIANCE_EXCLUDE_PREFIXES = {"GTM"}

# --- NON-CHARGE TAB ---
# Tasks where a blank note is expected and no response should be requested.
# Nobody needs to explain why they took a holiday. Every other non-charge task
# logged with no note and no description gets flagged for a response.
# Match on the task name exactly as it appears in the report, upper-cased.
NONCHARGE_NO_NOTE_TASKS = {
    "NONCHARGE HRS HOLIDAY",
    "NONCHARGE HRS FLOATING HOLIDAY",
    "NONCHARGE HRS PTO",
    "NCH MATERNITY/PATERNITY/ADOPT",
    "NCH BEREAVEMENT",
    "NCH JURY DUTY",
    "NCH PPL",
}

# --- NON-CHARGE COLUMN GROUPING ---
# Which tasks get their own column in the non-charge summaries, and in what
# order. Derived from a full-year analysis of the time report: Admin/Scheduling
# and Business Development alone are 63% of all non-charge time, so leaving
# them in "Other" made the breakout useless.
#
# Ordering runs activity first (what people chose to do), then time off.
# Match is on the exact task name, upper-cased — substring matching quietly
# misfiles anything new. Tasks not listed here fall into "Other"; the app
# surfaces those so a new task doesn't disappear silently.
NONCHARGE_TASK_GROUPS = {
    "Available Time":    {"AVAILABLE TIME"},
    "Admin/Scheduling":  {"NONCHARGE HRS ADMIN/SCHEDULING"},
    "Business Dev":      {"NCH BUSINESS DEVELOPMENT"},
    "Training":          {"NONCHARGE HRS TRAINING"},
    "Marketing":         {"NONCHARGE HRS MARKETING"},
    "Client Assistance": {"TTG - GTM CLIENT ASSISTANCE", "TTG - TECHNICAL RESEARCH"},
    "PTO":               {"NONCHARGE HRS PTO"},
    "Holiday":           {"NONCHARGE HRS HOLIDAY", "NONCHARGE HRS FLOATING HOLIDAY"},
    # Statutory / extended leave, grouped because each is individually rare.
    # FMLA and Military Leave are on the OpenAir task list but unused so far.
    "Leave":             {"NCH MATERNITY/PATERNITY/ADOPT", "NCH BEREAVEMENT",
                          "NCH PPL", "NCH JURY DUTY", "FMLA", "MILITARY LEAVE"},
}
NONCHARGE_OTHER_LABEL = "Other"
# Column order for the summary tables.
NONCHARGE_COLUMN_ORDER = list(NONCHARGE_TASK_GROUPS) + [NONCHARGE_OTHER_LABEL]

# --- TIME ENTRY CHECK ---
# Flags people who didn't enter (or under-entered) their time in OpenAir for
# the previous full Mon-Sun week. The reminder appears on the Current Month
# Hours tab of their workbook.
# Set TIME_ENTRY_EXPECTED_WEEKLY_HOURS to 0 to only flag people who entered
# nothing at all, and never flag partial weeks.
TIME_ENTRY_EXPECTED_WEEKLY_HOURS = 40
# Daily totals below this are treated as no entry (guards against a stray 0).
TIME_ENTRY_MIN_HOURS_TO_COUNT = 0.01

# --- PURPLE FILL DETECTION (Month tab "done" cells) ---
PURPLE_HEX_CODES = {
    "A02B93",   # Confirmed accent5 from this workbook's theme
    "7030A0",   # Standard Excel Purple
    "8064A2",
    "9B59B6",
    "800080",
}

# --- POSITION ORDER ---
# Role codes from the Utilization tab (col B). Used for sorting names everywhere.
POSITION_ORDER = {"MD": 1, "DR": 2, "SM": 3, "MR": 4, "SR": 5, "AN": 6, "IN": 7}

# Maps each person's name (as it appears in the file) to their role code.
PERSON_ROLE = {
    "Sorrentino":   "MD",
    "Colonna":      "DR",
    "Browne":       "SM",
    "Jean":         "SM",
    "Hendrickson":  "SM",
    "Lowry":        "SM",
    "Wojtowicz":    "MR",
    "Brooks":       "MR",
    "Lighthall":    "MR",
    "McGrogan":     "SR",
    "S. O'Donnell": "AN",
    "Avington":     "AN",
    "J. O'Donnell": "AN",
}

def _rank(name: str) -> tuple:
    """Sort key: (position_rank, last_name). Lower rank = higher seniority."""
    role = PERSON_ROLE.get(name, "ZZ")
    return (POSITION_ORDER.get(role, 99), name)

# --- NAME ALIASES ---
# Maps ambiguous names (no initials) to their canonical form.
# "O'Donnell" in older sheets has no initials — treat it as "J. O'Donnell".
# Update when the file adds initials for everyone.
NAME_ALIASES = {
    "O'Donnell": "J. O'Donnell",
}

# Maps full OpenAir employee names ('LastName, FirstName') to the exact
# name used in the schedule file. Only needed when last names are ambiguous.
OPENAIR_EMPLOYEE_MAP = {
    "O'Donnell, Jake":  "J. O'Donnell",
    "O'Donnell, Scott": "S. O'Donnell",
}

# --- FIRST NAMES (for email greetings) ---
# Keys match exactly how names appear in the file (last name / "S. O'Donnell" style)
FIRST_NAMES = {
    "J. O'Donnell": "Jake",
    "S. O'Donnell": "Scott",
    "Sorrentino":   "Anthony",
    "Colonna":      "Dante",
    "Browne":       "Nicole",
    "Jean":         "Ricot",
    "Hendrickson":  "Laren",
    "Lowry":        "Blake",
    "Wojtowicz":    "Agnes",
    "Brooks":       "Valerie",
    "Lighthall":    "Haley",
    "McGrogan":     "Amanda",
    "Avington":     "Alyssa",
}

# --- DISPLAY NAMES (optional override for recipient labels in the app) ---
# Maps schedule key → display label shown in parentheses in the UI.
# Use when someone's name in the schedule differs from their current name.
DISPLAY_NAMES = {
    "McGrogan": "Holmes (McGrogan)",   # schedule uses initial name; show new name in UI
}

# ============================================================
# EMAIL LOOKUP — TESTING MODE
# All emails route to jodonnell@gtmtax.com.
# When ready for production, replace each value with the real email.
#
# Keys must match EXACTLY how names appear in the file:
#   - Budget to Actual  col H (Project Owner)
#   - Project Tracker   col H (Project Owner)
#   - Month tabs        row 2 (last name / "S. O'Donnell" style)
#   - Utilization tab   person name column
# ============================================================
EMAIL_LOOKUP = {
    "J. O'Donnell": "jodonnell@gtmtax.com",
    "S. O'Donnell": "sodonnell@gtmtax.com",
    "Sorrentino":   "asorrentino@gtmtax.com",
    "Colonna":      "dcolonna@gtmtax.com",
    "Browne":       "nbrowne@gtmtax.com",
    "Jean":         "rjean@gtmtax.com",
    "Hendrickson":  "lhendrickson@gtmtax.com",
    "Lowry":        "blowry@gtmtax.com",
    "Wojtowicz":    "awojtowicz@gtmtax.com",
    "Brooks":       "vbrooks@gtmtax.com",
    "Lighthall":    "hlighthall@gtmtax.com",
    "McGrogan":     "aholmes@gtmtax.com",
    "Avington":     "aavington@gtmtax.com",
}
