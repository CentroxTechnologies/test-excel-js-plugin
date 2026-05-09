"""Tests for the AI-output validator."""

from __future__ import annotations

import pytest

from models import ProcessResponse
from validator import validate_action


def _resp(action_type: str, params: dict, preview: str = "ok") -> ProcessResponse:
    return ProcessResponse(
        action_type=action_type,
        params=params,
        preview_text=preview,
        confidence=1.0,
    )


# ---------- existing structural gates ----------


def test_insert_formula_passes_when_well_formed():
    result = validate_action(
        _resp("insert_formula", {"cell": "C7", "formula": "=SUM(B2:B10)"})
    )
    assert result.action_type == "insert_formula"


def test_insert_formula_rejects_missing_equals():
    result = validate_action(
        _resp("insert_formula", {"cell": "C7", "formula": "SUM(B2:B10)"})
    )
    assert result.action_type == "show_insight"


def test_insert_formula_rejects_bad_cell_reference():
    result = validate_action(
        _resp("insert_formula", {"cell": "Z!!", "formula": "=SUM(B2:B10)"})
    )
    assert result.action_type == "show_insight"


def test_format_cells_rejects_invalid_range():
    result = validate_action(
        _resp("format_cells", {"range": "not_a_range", "bold": True})
    )
    assert result.action_type == "show_insight"


def test_create_chart_rejects_unknown_chart_type():
    result = validate_action(
        _resp(
            "create_chart",
            {"data_range": "A1:B10", "chart_type": "WafflePlot", "title": "x"},
        )
    )
    assert result.action_type == "show_insight"


def test_sort_range_rejects_negative_column():
    result = validate_action(
        _resp("sort_range", {"range": "A1:D8", "sort_column": -1})
    )
    assert result.action_type == "show_insight"


def test_show_insight_passes_through():
    result = validate_action(
        _resp("show_insight", {"text": "Highest is 42 in row 5"})
    )
    assert result.action_type == "show_insight"


# ---------- new gates: reference check, type check, range completeness ----------


def test_reference_check_rejects_formula_pointing_past_data(monkeypatch):
    """If formula references row 50 but sheet only has 10 rows, reject."""
    sheet_data = [["Name", "Sales"], *([["A", 100]] * 9)]  # 10 rows
    result = validate_action(
        _resp("insert_formula", {"cell": "C1", "formula": "=SUM(B2:B50)"}),
        sheet_data=sheet_data,
        headers=["Name", "Sales"],
    )
    assert result.action_type == "show_insight"
    assert "row" in result.preview_text.lower() or "exist" in result.preview_text.lower()


def test_reference_check_rejects_formula_pointing_past_columns():
    sheet_data = [["A", "B"], ["1", "2"]]
    result = validate_action(
        _resp("insert_formula", {"cell": "C1", "formula": "=SUM(D1:D2)"}),
        sheet_data=sheet_data,
        headers=["A", "B"],
    )
    assert result.action_type == "show_insight"


def test_reference_check_passes_when_refs_inside_data():
    sheet_data = [["Name", "Sales"], *([["A", 100]] * 9)]
    result = validate_action(
        _resp("insert_formula", {"cell": "C1", "formula": "=SUM(B2:B10)"}),
        sheet_data=sheet_data,
        headers=["Name", "Sales"],
    )
    assert result.action_type == "insert_formula"


def test_type_check_rejects_sum_on_text_column():
    sheet_data = [
        ["Name", "Region"],
        ["Ali", "East"],
        ["Sara", "West"],
        ["Hassan", "East"],
    ]
    result = validate_action(
        _resp("insert_formula", {"cell": "C1", "formula": "=SUM(B2:B4)"}),
        sheet_data=sheet_data,
        headers=["Name", "Region"],
    )
    assert result.action_type == "show_insight"
    assert "numeric" in result.preview_text.lower() or "text" in result.preview_text.lower()


def test_type_check_passes_sum_on_numeric_column():
    sheet_data = [
        ["Name", "Sales"],
        ["Ali", 100],
        ["Sara", 200],
        ["Hassan", 300],
    ]
    result = validate_action(
        _resp("insert_formula", {"cell": "C1", "formula": "=SUM(B2:B4)"}),
        sheet_data=sheet_data,
        headers=["Name", "Sales"],
    )
    assert result.action_type == "insert_formula"


def test_type_check_only_runs_for_aggregation_functions():
    """COUNTA on a text column is fine; only SUM/AVERAGE/etc need numeric data."""
    sheet_data = [["Name"], ["Ali"], ["Sara"]]
    result = validate_action(
        _resp("insert_formula", {"cell": "B1", "formula": "=COUNTA(A2:A3)"}),
        sheet_data=sheet_data,
        headers=["Name"],
    )
    assert result.action_type == "insert_formula"


def test_range_completeness_warns_when_user_says_all_but_range_is_partial():
    """User said 'sum ALL sales' but formula only covers rows 2-5 of 10."""
    sheet_data = [["Sales"], *([[100]] * 10)]  # 11 rows total (header + 10)
    user_message = "sum all sales"
    result = validate_action(
        _resp("insert_formula", {"cell": "B1", "formula": "=SUM(A2:A5)"}),
        sheet_data=sheet_data,
        headers=["Sales"],
        user_message=user_message,
    )
    # We turn this into an insight asking the user to confirm
    assert result.action_type == "show_insight"
    assert (
        "all" in result.preview_text.lower()
        or "complete" in result.preview_text.lower()
        or "row" in result.preview_text.lower()
    )


def test_range_completeness_passes_when_range_covers_all_rows():
    sheet_data = [["Sales"], *([[100]] * 10)]
    result = validate_action(
        _resp("insert_formula", {"cell": "B1", "formula": "=SUM(A2:A11)"}),
        sheet_data=sheet_data,
        headers=["Sales"],
        user_message="sum all sales",
    )
    assert result.action_type == "insert_formula"


def test_validator_handles_missing_optional_args():
    """Old call sites without sheet_data / headers / user_message must still work."""
    result = validate_action(
        _resp("insert_formula", {"cell": "C7", "formula": "=SUM(B2:B10)"})
    )
    assert result.action_type == "insert_formula"


def test_unknown_action_type_is_passed_through():
    result = validate_action(
        _resp("frobnicate", {"foo": "bar"})
    )
    assert result.action_type == "frobnicate"
