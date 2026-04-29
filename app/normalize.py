"""Apartment number + last-name normalization.

Ported faithfully from the permit-audit skill (SKILL.md, Data Parsing section).
Apartment numbers come in many formats across rent rolls and parking systems;
this module's job is to converge them onto a canonical key for cross-referencing.
"""
from __future__ import annotations

import re
from typing import Optional

# Tokens that are NOT apartment units — skip these in both data sources
_SKIP_TOKENS = {
    "OFFICE STAFF", "STAFF", "ADMIN", "DOWN", "VACANT", "MODEL",
    "NOT ON RENT ROLL", "OFFICE", "MGMT", "MANAGEMENT",
}

# Date pattern that sometimes appears in name fields (NTV move-out date, etc.)
_DATE_RE = re.compile(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b")

# Apartment number extraction patterns, ordered most-specific first.
# Each must capture the digit portion in group 1.
_APT_PATTERNS = [
    re.compile(r"^[A-Z]{2,}-(\d+)$"),       # "PAPC-0304" → 304
    re.compile(r"^[A-Z]{2,}(\d+)$"),        # "PAPC0304"  → 304
    re.compile(r"^#\s*(\d+)$"),             # "#207"      → 207
    re.compile(r"^[A-Z](\d+)$"),            # "A0111"     → 111
    re.compile(r"^(\d+)\s*-+$"),            # "111-"      → 111
    re.compile(r"^(\d+)$"),                 # "0111"      → 111 (after lstrip)
]


def extract_apt_number(raw) -> Optional[str]:
    """Normalize a raw apartment number to canonical form (digits only, no leading zeros).

    Returns None if the value is not an apartment unit (office, staff, blank, etc.).

    Examples (from SKILL.md):
        "0111"      → "111"
        "#207"      → "207"
        "PAPC-0304" → "304"
        "A0111"     → "111"
        "111-"      → "111"
        "Office Staff" → None
    """
    if raw is None:
        return None
    s = str(raw).strip().lstrip("'")  # Excel sometimes prefixes "text-formatted" cells with '
    if not s:
        return None
    upper = s.upper()
    if upper in _SKIP_TOKENS:
        return None
    # Quick reject: must contain at least one digit
    if not any(c.isdigit() for c in s):
        return None
    for pat in _APT_PATTERNS:
        m = pat.match(upper)
        if m:
            digits = m.group(1).lstrip("0")
            return digits or "0"  # e.g. "000" → "0"
    return None


def display_apt_number(raw_unit: str, canonical: str) -> str:
    """If the source data uses leading zeros (e.g., '0111'), preserve that for display.

    `raw_unit` is the original string from the rent roll; `canonical` is the
    output of extract_apt_number(raw_unit). Returns whichever the property uses.
    """
    if raw_unit is None:
        return canonical
    s = str(raw_unit).strip()
    # Strip non-digit prefixes/suffixes but preserve internal zeros
    digits_only = re.sub(r"^[A-Z#-]+", "", s.upper())
    digits_only = re.sub(r"[-]+$", "", digits_only)
    if digits_only and digits_only.isdigit() and digits_only != canonical:
        return digits_only
    return canonical


def extract_last_name(name) -> str:
    """Extract a normalized last name for discrepancy comparison.

    Per SKILL.md:
      "Smith, John" → "smith"   (comma format: take before comma)
      "John Smith"  → "smith"   (space format: take last word)
      Strip "05/23/2025"-style date noise.
      Case-insensitive comparison.
    """
    if name is None:
        return ""
    s = str(name).strip()
    if not s:
        return ""
    # Strip date noise that sometimes contaminates name fields
    s = _DATE_RE.sub("", s).strip()
    if not s:
        return ""
    if s.upper() in _SKIP_TOKENS:
        return ""
    if "," in s:
        last = s.split(",", 1)[0].strip()
    else:
        parts = [p for p in s.split() if p]
        if not parts:
            return ""
        # Strip generational suffixes: "John Smith Jr" → "Smith", "Mary Lee III" → "Lee"
        suffixes = {"jr", "jr.", "sr", "sr.", "ii", "iii", "iv", "v"}
        while len(parts) > 1 and parts[-1].lower().rstrip(".,") in suffixes:
            parts.pop()
        last = parts[-1]
    return last.lower().strip(".,;:'\"")


def is_skip_name(name) -> bool:
    """Return True if this 'name' should be excluded from compliance + discrepancy logic.

    SKILL.md: skip VACANT, MODEL, ADMIN/DOWN, NOT ON RENT ROLL entries.
    """
    if name is None:
        return True
    s = str(name).strip().upper()
    if not s:
        return True
    return any(tok in s for tok in _SKIP_TOKENS)
