"""Cell type detection, format parsing, and semantic type classification.

Implements the 9 semantic types from SpreadsheetLLM paper Section 3.4:
Year, Integer, Float, Percentage, Scientific, Date, Time, Currency, Email.
"""

from __future__ import annotations

import re
from typing import Tuple

from openpyxl.cell.cell import Cell
from openpyxl.utils import get_column_letter

# Pre-compiled regex patterns (module-level for performance)
EMAIL_REGEX = re.compile(r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$")
CURRENCY_REGEX = re.compile(r"[$€£¥]")
SCIENTIFIC_REGEX = re.compile(r"E[+-]", re.IGNORECASE)
DATE_KEYWORDS = re.compile(r"(yyyy|yy|mmmm|mmm|dd|ddd|dddd)", re.IGNORECASE)
TIME_KEYWORDS = re.compile(r"(hh?|ss?|am/pm|a/p)", re.IGNORECASE)
YEAR_ONLY_REGEX = re.compile(r"^y{2,4}$", re.IGNORECASE)
NUMERIC_NFS_REGEX = re.compile(r"^[#0,]+(\.[#0]+)?$")

# Style fingerprint type
StyleFingerprint = Tuple[int, int, int, int]


def infer_cell_data_type(cell: Cell) -> str:
    """Infer the data type of a cell from openpyxl metadata and value.

    Returns one of: empty, text, numeric, boolean, datetime, email, error.
    """
    if cell.value is None:
        return "empty"

    # Check email before falling through to text
    if isinstance(cell.value, str) and EMAIL_REGEX.match(cell.value):
        return "email"

    dt = cell.data_type
    if dt == "s":
        return "text"
    if dt == "n":
        return "numeric"
    if dt == "b":
        return "boolean"
    if dt == "d":
        return "datetime"
    if dt == "e":
        return "error"
    if dt == "f":
        # Formula — infer from cached value
        v = cell.value
        if isinstance(v, str):
            return "text"
        if isinstance(v, (int, float)):
            return "numeric"
        if isinstance(v, bool):
            return "boolean"
        return "formula"
    return "text"  # fallback


def get_number_format_string(cell: Cell) -> str:
    """Return the raw Number Format String for a cell."""
    try:
        nfs = cell.number_format
        if nfs is None or nfs == "":
            return "General"
        return str(nfs)
    except Exception:
        return "General"


def categorize_number_format(nfs: str, cell: Cell) -> str:
    """Categorize a number format string into a broad category.

    Returns one of: general, currency, percentage, scientific, fraction,
    date_custom, time_custom, datetime_custom, integer, float,
    other_numeric, not_applicable.
    """
    cell_type = infer_cell_data_type(cell)
    if cell_type not in ("numeric", "datetime"):
        return "not_applicable"

    if nfs is None or nfs.lower() == "general":
        return "datetime_general" if cell_type == "datetime" else "general"

    if nfs == "@" or nfs.lower() == "text":
        return "text_format"

    if CURRENCY_REGEX.search(nfs):
        return "currency"
    if "%" in nfs:
        return "percentage"
    if SCIENTIFIC_REGEX.search(nfs):
        return "scientific"
    if "#" in nfs and "/" in nfs and "?" in nfs:
        return "fraction"

    nfs_lower = nfs.lower()
    is_date = bool(DATE_KEYWORDS.search(nfs_lower))
    is_time = bool(TIME_KEYWORDS.search(nfs_lower))

    # Bare "m" could be month or minute — disambiguate
    if "m" in nfs_lower and not is_date and not is_time:
        if any(k in nfs_lower for k in ("h", "s")):
            is_time = True
        elif re.fullmatch(r"m{1,5}", nfs_lower):
            is_date = True

    if is_date and is_time:
        return "datetime_custom"
    if is_date:
        return "date_custom"
    if is_time:
        return "time_custom"

    if cell_type == "numeric":
        if nfs in ("0", "#,##0"):
            return "integer"
        if nfs in ("0.00", "#,##0.00", "0.0", "#,##0.0"):
            return "float"
        if NUMERIC_NFS_REGEX.match(nfs):
            return "other_numeric"

    if cell_type == "datetime":
        return "other_date"

    return "not_applicable"


def detect_semantic_type(cell: Cell) -> str:
    """Detect the semantic type of a cell per the paper's 9 categories.

    Returns one of: year, integer, float, percentage, scientific,
    date, time, currency, email, text, empty, boolean, error.
    """
    data_type = infer_cell_data_type(cell)

    if data_type == "empty":
        return "empty"
    if data_type == "email":
        return "email"
    if data_type == "boolean":
        return "boolean"
    if data_type == "error":
        return "error"

    nfs = get_number_format_string(cell)
    category = categorize_number_format(nfs, cell)
    nfs_lower = nfs.lower()

    if category == "percentage":
        return "percentage"
    if category == "currency":
        return "currency"
    if category == "scientific":
        return "scientific"
    if category in ("date_custom", "datetime_custom", "datetime_general", "other_date"):
        # Check if it's year-only format
        if YEAR_ONLY_REGEX.match(nfs_lower):
            return "year"
        return "date"
    if category == "time_custom":
        return "time"

    if data_type == "numeric":
        if isinstance(cell.value, int) or category == "integer":
            return "integer"
        if isinstance(cell.value, float) or category in ("float", "other_numeric"):
            return "float"
        return "integer"  # numeric fallback

    return "text"


def get_cell_style_fingerprint(cell: Cell) -> StyleFingerprint:
    """Return a hashable style fingerprint for boundary comparison.

    Uses attribute hashing (works in both in-memory and loaded workbooks).
    """
    try:
        font = cell.font
        f_key = (font.bold, font.italic, font.underline, font.sz,
                 str(font.color.rgb) if font.color and hasattr(font.color, 'rgb') and font.color.rgb else None)
    except Exception:
        f_key = None

    try:
        border = cell.border
        b_key = (border.left.style, border.right.style,
                 border.top.style, border.bottom.style)
    except Exception:
        b_key = None

    try:
        fill = cell.fill
        fill_key = (fill.patternType,
                    str(fill.fgColor.rgb) if fill.fgColor and hasattr(fill.fgColor, 'rgb') and fill.fgColor.rgb else None)
    except Exception:
        fill_key = None

    try:
        align = cell.alignment
        a_key = (align.horizontal, align.vertical)
    except Exception:
        a_key = None

    return (hash(f_key), hash(b_key), hash(fill_key), hash(a_key))


def cell_coord(row: int, col: int) -> str:
    """Convert (row, col) to Excel coordinate like 'A1'."""
    return f"{get_column_letter(col)}{row}"


def split_cell_ref(ref: str) -> tuple[str, int]:
    """Split 'AB12' into ('AB', 12)."""
    col_str = ""
    row_str = ""
    for ch in ref:
        if ch.isalpha():
            col_str += ch
        else:
            row_str += ch
    return col_str, int(row_str)
