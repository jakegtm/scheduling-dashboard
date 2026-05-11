from __future__ import annotations
# ============================================================
# config.py — GTM Scheduling Analyzer Configuration
# ============================================================

SENDER_EMAIL = "jodonnell@gtmtax.com"

# --- SENDER NAME DIRECTORY ---
SENDER_NAMES = {
    "jodonnell@gtmtax.com":    "Jake",
    "lhendrickson@gtmtax.com": "Laren",
    "laren@gtmtax.com":        "Laren",
}

# --- THRESHOLDS (UI defaults) ---
DEFAULT_BUDGET_THRESHOLD      = 20000  # Flag remaining > this (unscheduled)
DEFAULT_NEGATIVE_THRESHOLD    = 100    # Flag if remaining < -$X (e.g. -$100)
DEFAULT_PROJECTION_THRESHOLD_PCT = 0.80
DEADLINE_WARNING_DAYS         = 2

# --- PURPLE FILL DETECTION ---
PURPLE_HEX_CODES = {
    "A02B93", "7030A0", "8064A2", "9B59B6", "800080",
}

# ============================================================
# EMAIL LOOKUP
# Key  = name as it appears in the schedule file (last name,
#         or "S. O'Donnell" style for duplicates)
# Each entry has:
#   email       : recipient address
#   first_name  : used in email greeting ("Hi Laren,")
# ============================================================
EMAIL_LOOKUP = {
    "O'Donnell":    {"email": "jodonnell@gtmtax.com",  "first_name": "Jake"},
    "J. O'Donnell": {"email": "jodonnell@gtmtax.com",  "first_name": "Jake"},
    "S. O'Donnell": {"email": "jodonnell@gtmtax.com",  "first_name": "Scott"},
    "Sorrentino":   {"email": "jodonnell@gtmtax.com",  "first_name": "Anthony"},
    "Colonna":      {"email": "jodonnell@gtmtax.com",  "first_name": "Dante"},
    "Browne":       {"email": "jodonnell@gtmtax.com",  "first_name": "Nicole"},
    "Jean":         {"email": "jodonnell@gtmtax.com",  "first_name": "Ricot"},
    "Hendrickson":  {"email": "jodonnell@gtmtax.com",  "first_name": "Laren"},
    "Lowry":        {"email": "jodonnell@gtmtax.com",  "first_name": "Blake"},
    "Wojtowicz":    {"email": "jodonnell@gtmtax.com",  "first_name": "Agnieszka"},
    "Brooks":       {"email": "jodonnell@gtmtax.com",  "first_name": "Valerie"},
    "Lighthall":    {"email": "jodonnell@gtmtax.com",  "first_name": "Haley"},
    "McGrogan":     {"email": "jodonnell@gtmtax.com",  "first_name": "Amanda"},
    "Avington":     {"email": "jodonnell@gtmtax.com",  "first_name": "Alyssa"},
}

# ============================================================
# OpenAir last name → schedule file key
# Handles cases where OpenAir "Last, First" last name doesn't
# exactly match the schedule column header (e.g. duplicates)
# ============================================================
OPENAIR_NAME_MAP = {
    # "OpenAir last name": "schedule file key"
    # Add entries only when they differ
    # e.g. "ODonnell": "J. O'Donnell"
}
