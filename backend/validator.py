"""
Validator: the last line of defense before an AI action reaches the add-in.

Every response from the AI engine passes through `validate_action()`. If the
action's shape is wrong, or the formula references cells that don't exist, or
the user's "all of X" intent doesn't match the formula's range, we swap the
action for a friendly `show_insight` instead of letting bad output reach Excel.

Three layers:
  1. STRUCTURAL: action_type-specific shape checks (formula starts with =,
     range looks like A1:D10, chart_type is on the whitelist, etc).
  2. REFERENCE check: formula references stay inside the sheet's actual
     dimensions (no SUM(B2:B50) when there are only 10 rows).
  3. SEMANTIC: type check (SUM on text columns is rejected) and range
     completeness (user said "all" but the formula misses rows).

Layers (2) and (3) are skipped if `sheet_data` / `headers` / `user_message`
aren't passed in, so old call sites without that context still work.
"""

from __future__ import annotations

import re
from typing import Any

from models import ProcessResponse


# ---------------------------------------------------------------------------
# Future work for the validation pipeline:
#   - Formula syntax check (parse the formula through openpyxl.formula or
#     tree-sitter to catch unbalanced parens, unknown function names).
#   - Circular reference detection (flag formulas that reference their own
#     target cell directly or transitively).
#   - Test execution (write to a hidden temp cell, check for #ERROR before
#     returning the action).
#   - Reasonableness check (a percentage shouldn't be 50,000%).
# ---------------------------------------------------------------------------


_RANGE_PATTERN = re.compile(r"^[A-Z]+\d+(:[A-Z]+\d+)?$")
_FORMULA_REF_PATTERN = re.compile(r"\b([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?\b")

_AGGREGATION_FUNCTIONS = {"SUM", "AVERAGE", "AVG", "MIN", "MAX", "MEDIAN", "STDEV"}

_VALID_CHART_TYPES = {
    "ColumnClustered",
    "ColumnStacked",
    "BarClustered",
    "Bar",
    "Line",
    "Pie",
    "Area",
    "Scatter",
    "XYScatter",
    "Doughnut",
}


# ---------- structural gates (existing) ----------


def _reject(reason: str) -> ProcessResponse:
    return ProcessResponse(
        action_type="show_insight",
        params={"text": f"I couldn't do that: {reason}"},
        preview_text=f"I couldn't do that: {reason}",
        confidence=0.0,
    )


def _validate_insert_formula(params: dict) -> ProcessResponse | None:
    formula = params.get("formula", "")
    cell = params.get("cell", "")
    if not isinstance(formula, str) or not formula.startswith("="):
        return _reject("the formula is missing or doesn't start with '='.")
    if not isinstance(cell, str) or not _RANGE_PATTERN.match(cell.upper()):
        return _reject(f"the target cell '{cell}' isn't a valid reference.")
    return None


def _validate_format_cells(params: dict) -> ProcessResponse | None:
    target = params.get("range", "")
    if not isinstance(target, str) or not _RANGE_PATTERN.match(target.upper()):
        return _reject(f"the range '{target}' isn't a valid cell reference.")
    return None


def _validate_write_values(params: dict) -> ProcessResponse | None:
    values = params.get("values")
    if not isinstance(values, list) or not values:
        return _reject("no values were provided to write.")
    if not all(isinstance(row, list) for row in values):
        return _reject("values must be a 2D array (list of rows).")
    return None


def _validate_create_chart(params: dict) -> ProcessResponse | None:
    chart_type = params.get("chart_type", "")
    if chart_type not in _VALID_CHART_TYPES:
        return _reject(f"chart type '{chart_type}' isn't supported.")
    data_range = params.get("data_range", "")
    if not isinstance(data_range, str) or not _RANGE_PATTERN.match(data_range.upper()):
        return _reject(f"the chart data range '{data_range}' isn't valid.")
    return None


def _validate_sort_range(params: dict) -> ProcessResponse | None:
    target = params.get("range", "")
    if not isinstance(target, str) or not _RANGE_PATTERN.match(target.upper()):
        return _reject(f"the sort range '{target}' isn't valid.")
    col = params.get("sort_column")
    if not isinstance(col, int) or col < 0:
        return _reject("the sort column index is missing or negative.")
    return None


_STRUCTURAL_VALIDATORS = {
    "insert_formula": _validate_insert_formula,
    "format_cells": _validate_format_cells,
    "write_values": _validate_write_values,
    "create_chart": _validate_create_chart,
    "sort_range": _validate_sort_range,
}


# ---------- reference / type / completeness checks ----------


def _col_letter_to_index(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def _extract_formula_refs(formula: str) -> list[tuple[int, int, int, int]]:
    """
    Parse a formula like "=SUM(B2:B10)" and return a list of
    (start_col, start_row, end_col, end_row) tuples (0-based) for each
    cell or range reference found.
    """
    refs: list[tuple[int, int, int, int]] = []
    for match in _FORMULA_REF_PATTERN.finditer(formula.upper()):
        c1, r1, c2, r2 = match.groups()
        start_col = _col_letter_to_index(c1)
        start_row = int(r1) - 1
        end_col = _col_letter_to_index(c2) if c2 else start_col
        end_row = (int(r2) - 1) if r2 else start_row
        refs.append((start_col, start_row, end_col, end_row))
    return refs


def _extract_aggregation_function(formula: str) -> str | None:
    """If the formula starts with =FUNC(, return FUNC if it's an aggregation."""
    match = re.match(r"^=([A-Z]+)\(", formula.upper())
    if not match:
        return None
    func = match.group(1)
    return func if func in _AGGREGATION_FUNCTIONS else None


def _check_references_exist(
    formula: str, sheet_data: list[list[Any]], headers: list[Any]
) -> ProcessResponse | None:
    """All cell refs in the formula must point at existing rows/cols."""
    if not sheet_data:
        return None
    n_rows = len(sheet_data)
    n_cols = max((len(r) for r in sheet_data), default=len(headers) or 0)
    for start_col, start_row, end_col, end_row in _extract_formula_refs(formula):
        if start_row >= n_rows or end_row >= n_rows:
            return _reject(
                f"the formula references row {max(end_row, start_row) + 1} "
                f"but the sheet only has {n_rows} rows."
            )
        if start_col >= n_cols or end_col >= n_cols:
            return _reject(
                f"the formula references column {end_col + 1} but the sheet "
                f"only has {n_cols} columns."
            )
    return None


def _column_is_numeric(
    sheet_data: list[list[Any]], col_index: int, start_row: int, end_row: int
) -> bool:
    """At least 70% of values in the range must be numeric for a 'yes'."""
    values: list[Any] = []
    for r in range(start_row, end_row + 1):
        if r >= len(sheet_data):
            continue
        row = sheet_data[r]
        if col_index >= len(row):
            continue
        cell = row[col_index]
        if cell is None or cell == "":
            continue
        values.append(cell)
    if not values:
        return False
    numeric_count = 0
    for v in values:
        try:
            float(v)
            numeric_count += 1
        except (TypeError, ValueError):
            pass
    return (numeric_count / len(values)) >= 0.7


def _check_aggregation_types(
    formula: str, sheet_data: list[list[Any]]
) -> ProcessResponse | None:
    """SUM / AVERAGE / etc. must reference numeric columns."""
    if not sheet_data:
        return None
    func = _extract_aggregation_function(formula)
    if func is None:
        return None
    refs = _extract_formula_refs(formula)
    if not refs:
        return None
    # First range argument (after the SUM/AVERAGE token, the parser picked it up)
    start_col, start_row, end_col, end_row = refs[0]
    if start_col != end_col:
        # Multi-column ranges: skip this check; too noisy for the demo.
        return None
    if not _column_is_numeric(sheet_data, start_col, start_row, end_row):
        return _reject(
            f"can't {func.lower()} that column, the values look like text "
            f"rather than numbers."
        )
    return None


_ALL_INTENT = re.compile(r"\b(all|every|entire|whole|total)\b", re.IGNORECASE)


def _check_range_completeness(
    formula: str,
    sheet_data: list[list[Any]],
    user_message: str,
) -> ProcessResponse | None:
    """
    If the user said 'sum ALL sales' but the range only covers half the data,
    swap in a confirmation insight so they can correct it.
    """
    if not sheet_data or not user_message:
        return None
    if not _ALL_INTENT.search(user_message):
        return None
    refs = _extract_formula_refs(formula)
    if not refs:
        return None
    n_rows = len(sheet_data)
    last_data_row = n_rows - 1  # 0-based; assumes header is row 0
    for _, _, _, end_row in refs:
        if end_row < last_data_row - 0:  # tolerate exact match
            return _reject(
                f"you said 'all' but the formula only covers rows up to "
                f"{end_row + 1} of {n_rows}. Want me to extend it to row {n_rows}?"
            )
    return None


# ---------- public entry ----------


def validate_action(
    response: ProcessResponse,
    sheet_data: list[list[Any]] | None = None,
    headers: list[Any] | None = None,
    user_message: str = "",
) -> ProcessResponse:
    """
    Run all applicable validation gates. Returns either the original response
    or a `show_insight` rejection.

    `sheet_data`, `headers`, and `user_message` are optional. When omitted, the
    semantic gates (reference / type / completeness) are skipped and only the
    structural shape is checked.
    """
    structural = _STRUCTURAL_VALIDATORS.get(response.action_type)
    if structural is not None:
        rejection = structural(response.params)
        if rejection is not None:
            return rejection

    if response.action_type == "insert_formula":
        formula = response.params.get("formula", "")
        if sheet_data is not None and headers is not None:
            ref_rejection = _check_references_exist(formula, sheet_data, headers)
            if ref_rejection is not None:
                return ref_rejection
            type_rejection = _check_aggregation_types(formula, sheet_data)
            if type_rejection is not None:
                return type_rejection
        if sheet_data is not None:
            completeness_rejection = _check_range_completeness(
                formula, sheet_data, user_message
            )
            if completeness_rejection is not None:
                return completeness_rejection

    return response
