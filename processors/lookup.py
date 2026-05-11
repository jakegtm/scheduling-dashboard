from __future__ import annotations
# ============================================================
# processors/lookup.py — shared name/email lookup helpers
# ============================================================
from config import EMAIL_LOOKUP


def lookup_email(name: str) -> str | None:
    """Return email for a schedule file name (last name / S. O'Donnell style)."""
    name = str(name).strip() if name else ""
    if name in EMAIL_LOOKUP:
        return EMAIL_LOOKUP[name]["email"]
    for key, val in EMAIL_LOOKUP.items():
        if key.lower() == name.lower():
            return val["email"]
    return None


def lookup_first_name(name: str) -> str:
    """Return first name for a schedule file name. Falls back to the key itself."""
    name = str(name).strip() if name else ""
    if name in EMAIL_LOOKUP:
        return EMAIL_LOOKUP[name]["first_name"]
    for key, val in EMAIL_LOOKUP.items():
        if key.lower() == name.lower():
            return val["first_name"]
    # Fallback: use whatever was passed in
    return name.split()[0] if name else "there"


def lookup_by_openair(last_name: str) -> str | None:
    """
    Match an OpenAir last name to a schedule file key.
    Handles O'Donnell ambiguity via OPENAIR_NAME_MAP in config.
    Returns the matching EMAIL_LOOKUP key or None.
    """
    from config import OPENAIR_NAME_MAP
    ln = str(last_name).strip()

    # Explicit override map first
    if ln in OPENAIR_NAME_MAP:
        return OPENAIR_NAME_MAP[ln]

    # Direct match
    if ln in EMAIL_LOOKUP:
        return ln

    # Case-insensitive
    for key in EMAIL_LOOKUP:
        if key.lower() == ln.lower():
            return key

    # Partial (handles apostrophes, punctuation)
    clean = ln.replace("'", "").replace("-", "").lower()
    for key in EMAIL_LOOKUP:
        if key.replace("'", "").replace("-", "").lower() == clean:
            return key

    return None
