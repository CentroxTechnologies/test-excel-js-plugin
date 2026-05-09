"""
AI engine, turns the user's plain English Excel command into a structured
ProcessResponse via a real LLM call (OpenAI or Anthropic, switchable via .env).

The validator (validator.py) is the next layer; we deliberately don't sanitize
LLM output here beyond stripping markdown fences. If the model returns
something structurally wrong, validator catches it and swaps in a friendly
show_insight error.

If no API key is configured, generate_action returns a show_insight pointing
the user at backend/.env.
"""

from __future__ import annotations

import json
import os
import re
from abc import ABC, abstractmethod
from time import perf_counter
from typing import Any

from models import ProcessRequest, ProcessResponse


SYSTEM_PROMPT = """\
You are an AI assistant embedded inside Microsoft Excel as a sidebar plugin.
The user will give you a command in plain English along with their current spreadsheet context.

Your job is to figure out what action to perform and return a JSON response.

CRITICAL RULES:
- Return ONLY valid JSON. No markdown, no explanation, no backticks, no text before or after the JSON.
- NEVER compute values yourself. If the user asks for a sum, average, count, or any calculation, generate an Excel FORMULA and let Excel compute it.
- Every formula must be a valid Excel formula starting with "=".
- Cell references must match the actual data context provided.
- If unsure what the user wants, return a show_insight action asking for clarification.

CRITICAL ON `preview_text`:
- It MUST be conversational and human-friendly. Talk to the user, not at them.
- DO NOT start with "Insert =SUM(...)" or other formula-dump style. Start with a verb that describes the user-facing outcome ("Total up the Sales column...", "Lay down a budget template...", "Sort the table by...").
- End with "Apply?" so the user knows the next step, OR with a period if the answer is informational only.
- Mention column names, row counts, or other concrete numbers from the sheet context where it makes the line clearer.

Return JSON in exactly this format:
{
    "action_type": "one of: insert_formula, write_values, format_cells, create_chart, show_insight, sort_range",
    "params": { ... },
    "preview_text": "Conversational, friendly sentence describing what will happen, ending with Apply? or a period."
}

Action types and params:

insert_formula: { "cell": "B10", "formula": "=SUM(B2:B9)" }

write_values: { "start_cell": "A1", "values": [["a","b"],["c","d"]] }

format_cells: { "range": "A1:D1", "bold": true, "background": "#4472C4", "font_color": "#FFFFFF" }

create_chart: { "data_range": "A1:B10", "chart_type": "ColumnClustered", "title": "Revenue" }
Valid chart types: ColumnClustered, ColumnStacked, BarClustered, Line, Pie, Area, Doughnut, XYScatter

show_insight: { "text": "Your answer here" }

sort_range: { "range": "A1:D20", "sort_column": 1, "ascending": false }

WORKED EXAMPLES: when the user asks to "build", "create", or "make me a template", prefer write_values with realistic populated rows and formulas. Empty templates feel hollow; seed at least 5 rows of plausible data plus the formulas that would normally compute over them.

EXAMPLE 1, User: "build me a quarterly budget template starting at A1"
{
  "action_type": "write_values",
  "params": {
    "start_cell": "A1",
    "values": [
      ["Category", "Q1", "Q2", "Q3", "Q4", "Total"],
      ["Revenue", 50000, 65000, 72000, 81000, "=SUM(B2:E2)"],
      ["Cost of goods sold", 20000, 26000, 28800, 32400, "=SUM(B3:E3)"],
      ["Gross profit", "=B2-B3", "=C2-C3", "=D2-D3", "=E2-E3", "=SUM(B4:E4)"],
      ["Salaries", 12000, 12500, 13000, 13500, "=SUM(B5:E5)"],
      ["Marketing", 4000, 5500, 6000, 7000, "=SUM(B6:E6)"],
      ["Rent", 6000, 6000, 6000, 6000, "=SUM(B7:E7)"],
      ["Utilities", 1500, 1500, 1700, 1700, "=SUM(B8:E8)"],
      ["Other expenses", 2000, 2200, 2400, 2600, "=SUM(B9:E9)"],
      ["Total expenses", "=SUM(B5:B9)", "=SUM(C5:C9)", "=SUM(D5:D9)", "=SUM(E5:E9)", "=SUM(B10:E10)"],
      ["Net profit", "=B4-B10", "=C4-C10", "=D4-D10", "=E4-E10", "=SUM(B11:E11)"]
    ]
  },
  "preview_text": "Lay down a quarterly budget template starting at A1: 11 rows of revenue, cost, and expense lines plus computed totals and net profit. Apply?",
  "confidence": 0.95
}

EXAMPLE 2, User: "add a tax column at 8.5% next to sales" (assume Sales is column B with 5 data rows)
The tax column goes in the next empty column (e.g. column D if A=Name, B=Sales, C=Region). Header in row 1, formula =B2*0.085 in each populated row.
{
  "action_type": "write_values",
  "params": {
    "start_cell": "D1",
    "values": [
      ["Tax (8.5%)"],
      ["=B2*0.085"],
      ["=B3*0.085"],
      ["=B4*0.085"],
      ["=B5*0.085"]
    ]
  },
  "preview_text": "Add a Tax (8.5%) column in D, with one formula per Sales row referencing column B. Apply?",
  "confidence": 0.92
}

EXAMPLE 3, User: "make a sales tracker template"
{
  "action_type": "write_values",
  "params": {
    "start_cell": "A1",
    "values": [
      ["Date", "Customer", "Product", "Quantity", "Unit Price", "Total"],
      ["2026-01-15", "Acme Corp", "Widget Pro", 10, 49.99, "=D2*E2"],
      ["2026-01-18", "Globex Inc", "Widget Lite", 25, 19.99, "=D3*E3"],
      ["2026-01-22", "Initech", "Widget Pro", 5, 49.99, "=D4*E4"],
      ["2026-01-25", "Umbrella Co", "Widget Max", 50, 89.99, "=D5*E5"],
      ["2026-01-29", "Acme Corp", "Widget Lite", 100, 19.99, "=D6*E6"],
      ["", "", "", "", "Grand Total", "=SUM(F2:F6)"]
    ]
  },
  "preview_text": "Drop in a 6-row sales tracker with date, customer, product, qty, price, and an auto-computed total per row plus a grand total. Apply?",
  "confidence": 0.93
}

When in doubt about a "build / create / make" request, default to write_values with at least 5 rows of realistic seed data plus formulas, empty templates feel hollow.
"""


_PLACEHOLDER_VALUES = {
    "",
    "your-openai-key-here",
    "your-anthropic-key-here",
}


def _is_real_key(value: str | None) -> bool:
    return bool(value) and value.strip() not in _PLACEHOLDER_VALUES


class AIProvider(ABC):
    name: str
    model: str

    @abstractmethod
    def call(self, system_prompt: str, user_prompt: str) -> tuple[str, dict[str, Any]]:
        """Return (raw_text, usage_dict). usage may be empty."""


class OpenAIProvider(AIProvider):
    name = "openai"

    def __init__(self, api_key: str, model: str) -> None:
        from openai import OpenAI
        self._client = OpenAI(api_key=api_key)
        self.model = model

    def call(self, system_prompt: str, user_prompt: str) -> tuple[str, dict[str, Any]]:
        response = self._client.chat.completions.create(
            model=self.model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            response_format={"type": "json_object"},
            temperature=0,
        )
        text = response.choices[0].message.content or ""
        usage_obj = getattr(response, "usage", None)
        usage = {
            "tokens_in": getattr(usage_obj, "prompt_tokens", None),
            "tokens_out": getattr(usage_obj, "completion_tokens", None),
        }
        return text, usage


class AnthropicProvider(AIProvider):
    name = "anthropic"

    def __init__(self, api_key: str, model: str) -> None:
        import anthropic
        self._client = anthropic.Anthropic(api_key=api_key)
        self.model = model

    def call(self, system_prompt: str, user_prompt: str) -> tuple[str, dict[str, Any]]:
        response = self._client.messages.create(
            model=self.model,
            max_tokens=1024,
            system=system_prompt,
            messages=[{"role": "user", "content": user_prompt}],
        )
        text = "".join(
            block.text for block in response.content if getattr(block, "type", None) == "text"
        )
        usage_obj = getattr(response, "usage", None)
        usage = {
            "tokens_in": getattr(usage_obj, "input_tokens", None),
            "tokens_out": getattr(usage_obj, "output_tokens", None),
        }
        return text, usage


def get_provider() -> AIProvider | None:
    """Pick the provider based on AI_PROVIDER env. Return None if its key is missing."""
    choice = (os.getenv("AI_PROVIDER") or "openai").strip().lower()

    if choice == "openai":
        key = os.getenv("OPENAI_API_KEY")
        if not _is_real_key(key):
            return None
        model = os.getenv("OPENAI_MODEL") or "gpt-4o"
        return OpenAIProvider(api_key=key, model=model)

    if choice == "anthropic":
        key = os.getenv("ANTHROPIC_API_KEY")
        if not _is_real_key(key):
            return None
        model = os.getenv("ANTHROPIC_MODEL") or "claude-sonnet-4-6"
        return AnthropicProvider(api_key=key, model=model)

    return None


_FENCE_PATTERN = re.compile(r"^\s*```(?:json)?\s*|\s*```\s*$", re.IGNORECASE)


def _strip_code_fences(text: str) -> str:
    """Strip leading/trailing markdown fences. Some models add them despite instructions."""
    cleaned = text.strip()
    cleaned = _FENCE_PATTERN.sub("", cleaned)
    return cleaned.strip()


def _show_insight(text: str) -> ProcessResponse:
    return ProcessResponse(
        action_type="show_insight",
        params={"text": text},
        preview_text=text,
        confidence=0.0,
    )


_MAX_ROWS = 50


def _format_table(sheet_data: list[list[Any]]) -> str:
    if not sheet_data:
        return "(empty sheet)"
    rows = sheet_data[:_MAX_ROWS]
    lines = ["\t".join("" if c is None else str(c) for c in row) for row in rows]
    if len(sheet_data) > _MAX_ROWS:
        lines.append(f"... and {len(sheet_data) - _MAX_ROWS} more rows not shown")
    return "\n".join(lines)


def _build_user_prompt(req: ProcessRequest) -> str:
    headers_line = ", ".join(str(h) for h in req.headers) if req.headers else "(none)"
    total_cols = len(req.headers) if req.headers else (
        max((len(r) for r in req.sheet_data), default=0)
    )
    return (
        "SPREADSHEET CONTEXT:\n"
        f"Sheet name: {req.sheet_name}\n"
        f"Active cell: {req.active_cell}\n"
        f"Column headers: {headers_line}\n"
        "Data (first 50 rows max):\n"
        f"{_format_table(req.sheet_data)}\n"
        f"Total rows: {len(req.sheet_data)}\n"
        f"Total columns: {total_cols}\n"
        "\n"
        f"USER COMMAND: {req.message}"
    )


def _log_request(message: str, provider: AIProvider, elapsed_ms: int, usage: dict[str, Any]) -> None:
    snippet = message.replace("\n", " ")[:120]
    tin = usage.get("tokens_in")
    tout = usage.get("tokens_out")
    print(
        f'[ai] msg="{snippet}" provider={provider.name} model={provider.model} '
        f"elapsed_ms={elapsed_ms} tokens_in={tin} tokens_out={tout}",
        flush=True,
    )


def generate_action(req: ProcessRequest) -> ProcessResponse:
    """Entry point. main.py calls this."""
    provider = get_provider()
    if provider is None:
        return _show_insight(
            "API key not configured. Add your key to backend/.env "
            "(set AI_PROVIDER and the matching API key), then restart the server."
        )

    user_prompt = _build_user_prompt(req)

    t0 = perf_counter()
    try:
        raw_text, usage = provider.call(SYSTEM_PROMPT, user_prompt)
    except Exception as exc:
        return _show_insight(f"AI request failed: {exc}")
    elapsed_ms = int((perf_counter() - t0) * 1000)
    _log_request(req.message, provider, elapsed_ms, usage)

    json_text = _strip_code_fences(raw_text)
    try:
        data = json.loads(json_text)
    except json.JSONDecodeError:
        preview = raw_text[:200].replace("\n", " ")
        return _show_insight(f"Couldn't parse model response: {preview}")

    try:
        return ProcessResponse(**data)
    except Exception as exc:
        return _show_insight(f"Model returned an unexpected shape: {exc}")
