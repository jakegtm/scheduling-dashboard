# ============================================================
# config.py — GTM Scheduling Analyzer Configuration
# ============================================================

# --- EMAIL ---
# Testing: all emails route to J. O'Donnell
# Production: swap LAREN_EMAIL to laren@gtmtax.com
LAREN_EMAIL = "jodonnell@gtmtax.com"

# --- THRESHOLDS (used as UI defaults, all overridable in sidebar) ---
DEFAULT_BUDGET_THRESHOLD     = 20000  # Flag remaining unscheduled budget over this amount
DEFAULT_PROJECTION_THRESHOLD_PCT = 0.80
DEADLINE_WARNING_DAYS        = 2      # Days before period end to send reminder

# --- PURPLE FILL DETECTION (May tab) ---
# "Done" cells use theme color index 8 = accent5 = #A02B93 in this workbook's theme.
PURPLE_HEX_CODES = {
    "A02B93",  # Confirmed accent5 from this workbook's theme
    "7030A0",  # Standard Excel Purple
    "8064A2",
    "9B59B6",
    "800080",
}

# ============================================================
# EMAIL LOOKUP — TESTING MODE
# Every name routes to jodonnell@gtmtax.com for now.
# When ready for production, replace each value with the real email.
#
# Keys must match EXACTLY how names appear in the file:
#   - Budget to Actual  col H (Project Owner)
#   - Project Tracker   col H (Project Owner)
#   - Month tabs        row 2 (person last name / "S. O'Donnell" style)
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
