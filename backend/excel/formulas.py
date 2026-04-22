"""
RightCut — Formula helpers.
Pure functions — no openpyxl imports here.
"""

from __future__ import annotations

import re
from openpyxl.utils import get_column_letter


# ── Cell reference utilities ──────────────────────────────────────────────────

def col_to_letter(n: int) -> str:
    """1-indexed column number → Excel letter(s). e.g. 1→'A', 27→'AA'."""
    return get_column_letter(n)


def cell_ref(row: int, col: int) -> str:
    """1-indexed (row, col) → Excel cell reference e.g. (1,1)→'A1'."""
    return f"{col_to_letter(col)}{row}"


def range_ref(start_row: int, start_col: int, end_row: int, end_col: int) -> str:
    """Build a range like 'A1:C10'."""
    return f"{cell_ref(start_row, start_col)}:{cell_ref(end_row, end_col)}"


# ── Formula validation ────────────────────────────────────────────────────────

def validate_formula(formula: str) -> bool:
    """Basic sanity check: starts with '=', balanced parentheses."""
    if not formula.startswith("="):
        return False
    depth = 0
    for ch in formula:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        if depth < 0:
            return False
    return depth == 0


# ── Private markets formula builders ─────────────────────────────────────────

def cagr_formula(start_cell: str, end_cell: str, years: int) -> str:
    """CAGR = ((End/Start)^(1/N)) - 1"""
    return f"=(({end_cell}/{start_cell})^(1/{years}))-1"


def moic_formula(invested_cell: str, current_cell: str) -> str:
    """MOIC = Current Value / Invested Capital"""
    return f"=({current_cell}/{invested_cell})"


def irr_formula(range_ref: str) -> str:
    """Standard IRR formula."""
    return f"=IRR({range_ref})"


def xirr_formula(values_range: str, dates_range: str) -> str:
    """XIRR for irregular cash flows."""
    return f"=XIRR({values_range},{dates_range})"


def xnpv_formula(rate_cell: str, values_range: str, dates_range: str) -> str:
    """XNPV for irregular cash flows."""
    return f"=XNPV({rate_cell},{values_range},{dates_range})"


def ebitda_margin_formula(ebitda_cell: str, revenue_cell: str) -> str:
    """EBITDA Margin = EBITDA / Revenue"""
    return f"=IF({revenue_cell}<>0,{ebitda_cell}/{revenue_cell},\"N/A\")"


def ev_multiple_formula(ev_cell: str, metric_cell: str) -> str:
    """EV / Metric multiple (e.g. EV/EBITDA)."""
    return f"=IF({metric_cell}<>0,{ev_cell}/{metric_cell},\"N/A\")"


# ── Number format strings ─────────────────────────────────────────────────────

def currency_format() -> str:
    return '"$"#,##0'


def currency_decimal_format() -> str:
    return '"$"#,##0.0'


def percent_format() -> str:
    return "0.0%"


def multiple_format() -> str:
    return '0.0"x"'


def integer_format() -> str:
    return "#,##0"


def decimal_format(places: int = 2) -> str:
    return f"#,##0.{'0' * places}"


# ── Format string inference ───────────────────────────────────────────────────

def infer_format(header: str) -> str | None:
    """Guess an appropriate number format from a column header."""
    h = header.lower()
    if any(k in h for k in ("revenue", "ebitda", "ev ", "enterprise", "price", "value", "cost", "usd", "$", "arr", "mrr")):
        return currency_format()
    if any(k in h for k in ("margin", "cagr", "irr", "growth", "rate", "%", "percent")):
        return percent_format()
    if any(k in h for k in ("moic", "multiple", "ev/", "p/e", "ebitda/", "x")):
        return multiple_format()
    return None
