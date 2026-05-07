"""
Pydantic models for request/response between the Excel add-in and backend.

Every payload flowing through /api/process is validated against these models.
Keep them simple, this is the shape juniors will see first when reading the code.
"""

from typing import Any
from pydantic import BaseModel, Field


class ProcessRequest(BaseModel):
    """What the add-in sends us when the user types a command."""

    # The user's plain English instruction, e.g. "sum column revenue"
    message: str

    # Full used-range of the active sheet as a 2D array (rows of cells).
    # First row is usually headers, but we don't assume, we also pass `headers` separately.
    sheet_data: list[list[Any]] = Field(default_factory=list)

    # Extracted first-row values, passed separately so the AI engine has easy access to column names.
    headers: list[Any] = Field(default_factory=list)

    # Currently selected cell address, e.g. "B5". Used to decide where to place results.
    active_cell: str = "A1"

    # Name of the active worksheet, e.g. "Sheet1". Informational for now.
    sheet_name: str = "Sheet1"


class ProcessResponse(BaseModel):
    """What we send back, one action for the add-in to preview and execute."""

    # One of: insert_formula, write_values, format_cells, create_chart, show_insight, sort_range
    action_type: str

    # Action-specific parameters. Shape depends on action_type (see ai_engine.py for details).
    params: dict[str, Any] = Field(default_factory=dict)

    # Human-readable summary shown in the sidebar before the user clicks Apply.
    preview_text: str

    # 0.0 to 1.0 confidence score. Placeholder today, will reflect LLM confidence later.
    confidence: float = 1.0
