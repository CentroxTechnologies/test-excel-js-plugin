"""
Validator — the last line of defense before an AI action reaches the add-in.

Every response from the AI engine passes through `validate_action()`. If something
looks off (malformed formula, bad range, unknown chart type), we swap it for a
friendly show_insight error instead of letting the add-in crash Excel.
"""

import re
from models import ProcessResponse


# ---------------------------------------------------------------------------
# PLANNED v2 VALIDATION PIPELINE
# ---------------------------------------------------------------------------
# Today's checks are structural — they confirm the action's shape is reasonable.
# Once the real LLM is wired up, we'll layer on:
#
#   1. Formula syntax check — parse the formula through a lightweight AST (e.g.
#      openpyxl.formula or a tree-sitter grammar) to catch unbalanced parens,
#      unknown function names, bad argument counts.
#
#   2. Reference bounds check — ensure cell references in formulas point to
#      cells that actually exist in the sheet (compare against sheet_data
#      dimensions).
#
#   3. Type check — if the formula expects numerics (SUM, AVERAGE), verify the
#      referenced columns actually contain numeric data.
#
#   4. Circular reference detection — flag formulas that reference their own
#      target cell (directly or transitively through other pending actions).
#
# Each layer returns the same ProcessResponse shape; on failure, we replace the
# action with a show_insight that explains what went wrong in plain English.
# ---------------------------------------------------------------------------


# Matches "A1", "B10", "A1:D10", "AA1:ZZ100" etc. Single cell or range.
_RANGE_PATTERN = re.compile(r"^[A-Z]+\d+(:[A-Z]+\d+)?$")

_VALID_CHART_TYPES = {
    "ColumnClustered",
    "Line",
    "Pie",
    "Bar",
    "Area",
    "Scatter",
    "ColumnStacked",
    "BarClustered",
    "Doughnut",
}


def _reject(reason: str) -> ProcessResponse:
    """Turn a validation failure into a user-facing insight."""
    return ProcessResponse(
        action_type="show_insight",
        params={"text": f"I couldn't do that: {reason}"},
        preview_text=f"I couldn't do that: {reason}",
        confidence=0.0,
    )


def _validate_insert_formula(params: dict) -> ProcessResponse | None:
    """Return None if valid, otherwise a rejection response."""
    formula = params.get("formula", "")
    cell = params.get("cell", "")
    if not isinstance(formula, str) or not formula.startswith("="):
        return _reject("the formula is missing or doesn't start with '='.")
    if not isinstance(cell, str) or not _RANGE_PATTERN.match(cell.upper()):
        return _reject(f"the target cell '{cell}' isn't a valid reference.")
    return None


def _validate_format_cells(params: dict) -> ProcessResponse | None:
    """Range must look like A1 or A1:D10."""
    target = params.get("range", "")
    if not isinstance(target, str) or not _RANGE_PATTERN.match(target.upper()):
        return _reject(f"the range '{target}' isn't a valid cell reference.")
    return None


def _validate_write_values(params: dict) -> ProcessResponse | None:
    """Values must be a non-empty 2D list."""
    values = params.get("values")
    if not isinstance(values, list) or not values:
        return _reject("no values were provided to write.")
    if not all(isinstance(row, list) for row in values):
        return _reject("values must be a 2D array (list of rows).")
    return None


def _validate_create_chart(params: dict) -> ProcessResponse | None:
    """Chart type must be one we know how to build."""
    chart_type = params.get("chart_type", "")
    if chart_type not in _VALID_CHART_TYPES:
        return _reject(f"chart type '{chart_type}' isn't supported.")
    data_range = params.get("data_range", "")
    if not isinstance(data_range, str) or not _RANGE_PATTERN.match(data_range.upper()):
        return _reject(f"the chart data range '{data_range}' isn't valid.")
    return None


def _validate_sort_range(params: dict) -> ProcessResponse | None:
    """Range must be valid and sort_column must be a non-negative integer."""
    target = params.get("range", "")
    if not isinstance(target, str) or not _RANGE_PATTERN.match(target.upper()):
        return _reject(f"the sort range '{target}' isn't valid.")
    col = params.get("sort_column")
    if not isinstance(col, int) or col < 0:
        return _reject("the sort column index is missing or negative.")
    return None


# Map of action_type -> validator function. Anything not listed is passed through.
_VALIDATORS = {
    "insert_formula": _validate_insert_formula,
    "format_cells": _validate_format_cells,
    "write_values": _validate_write_values,
    "create_chart": _validate_create_chart,
    "sort_range": _validate_sort_range,
}


def validate_action(response: ProcessResponse) -> ProcessResponse:
    """Run the registered validator for this action_type; swap in an error on failure."""
    validator_fn = _VALIDATORS.get(response.action_type)
    if validator_fn is None:
        # show_insight and anything else we don't have structural rules for: pass through.
        return response
    rejection = validator_fn(response.params)
    return rejection if rejection is not None else response
