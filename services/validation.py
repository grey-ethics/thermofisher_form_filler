"""
services/validation.py
----------------------
Input validation and normalization helpers.

Functions:
- normalize_project_level(value): return valid string or None
- parse_bool(s): robust bool parser for CSV cells
"""

def normalize_project_level(value: str | None) -> str | None:
    if not value:
        return None
    v = str(value).strip()
    allowed = {"L1", "L2L", "L2", "L3L"}
    return v if v in allowed else None

def parse_bool(s) -> bool:
    if s is None:
        return False
    v = str(s).strip().lower()
    return v in ("1", "true", "yes", "y", "t")
