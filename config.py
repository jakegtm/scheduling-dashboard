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
DEFAULT_VARIANCE_MIN = -5   # Flag if difference <= this (actual < schedule)
DEFAULT_VARIANCE_MAX = 0    # Flag if difference >= this (actual > schedule)

# --- ROLE RULES ---
# Names of staff who only receive their own variance rows (not all team rows).
# Add more names here as needed (e.g. when more analysts are hired).
INTERN_NAMES = {"Avington"}  # kept for backwards compatibility

# Staff who only receive their own variance rows (not all project rows).
# Project owners (managers/directors) are everyone NOT in this set.
STAFF_NAMES = {"Avington", "S. O'Donnell", "J. O'Donnell", "McGrogan"}

# Project code prefixes to exclude from variance analysis.
# GTM internal/non-chargeable codes (NONCHG, TRAINING, HOLIDAYS, etc.) should not appear.
VARIANCE_EXCLUDE_PREFIXES = {"GTM"}

# --- PURPLE FILL DETECTION (Month tab "done" cells) ---
PURPLE_HEX_CODES = {
    "A02B93",   # Confirmed accent5 from this workbook's theme
    "7030A0",   # Standard Excel Purple
    "8064A2",
    "9B59B6",
    "800080",
}

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
    "O'Donnell":    "Jake",
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
    "O'Donnell":    "jodonnell@gtmtax.com",
    "J. O'Donnell": "jodonnell@gtmtax.com",
    "S. O'Donnell": "jodonnell@gtmtax.com",
    "Sorrentino":   "jodonnell@gtmtax.com",
    "Colonna":      "jodonnell@gtmtax.com",
    "Browne":       "jodonnell@gtmtax.com",
    "Jean":         "jodonnell@gtmtax.com",
    "Hendrickson":  "jodonnell@gtmtax.com",
    "Lowry":        "jodonnell@gtmtax.com",
    "Wojtowicz":    "jodonnell@gtmtax.com",
    "Brooks":       "jodonnell@gtmtax.com",
    "Lighthall":    "jodonnell@gtmtax.com",
    "McGrogan":     "jodonnell@gtmtax.com",
    "Avington":     "jodonnell@gtmtax.com",
}
