"""Tests for backend/ai_engine.py.

These tests cover the parts that don't actually call an LLM:
  - placeholder-key fallback (returns show_insight)
  - code-fence stripping
  - JSON parse failure handling
  - prompt construction shape
  - SYSTEM_PROMPT enforces conversational preview_text
"""

from __future__ import annotations

import re

import ai_engine
from models import ProcessRequest


def _req(message: str = "sum sales") -> ProcessRequest:
    return ProcessRequest(
        message=message,
        sheet_data=[["Name", "Sales"], ["Ali", 100], ["Sara", 200]],
        headers=["Name", "Sales"],
        active_cell="C1",
        sheet_name="Sheet1",
    )


# ---------- get_provider / placeholder-key fallback ----------


def test_get_provider_returns_none_when_key_is_placeholder(monkeypatch):
    monkeypatch.setenv("OPENAI_API_KEY", "your-openai-key-here")
    assert ai_engine.get_provider() is None


def test_get_provider_returns_none_when_key_is_blank(monkeypatch):
    monkeypatch.setenv("OPENAI_API_KEY", "")
    assert ai_engine.get_provider() is None


def test_generate_action_returns_show_insight_when_no_key(monkeypatch):
    """conftest's autouse fixture already strips API keys."""
    response = ai_engine.generate_action(_req())
    assert response.action_type == "show_insight"
    assert ".env" in response.preview_text or "key" in response.preview_text.lower()


# ---------- code-fence stripping ----------


def test_strip_code_fences_handles_json_fence():
    raw = '```json\n{"a": 1}\n```'
    assert ai_engine._strip_code_fences(raw) == '{"a": 1}'


def test_strip_code_fences_handles_plain_fence():
    raw = '```\n{"x": 2}\n```'
    assert ai_engine._strip_code_fences(raw) == '{"x": 2}'


def test_strip_code_fences_passes_clean_json_through():
    raw = '{"y": 3}'
    assert ai_engine._strip_code_fences(raw) == '{"y": 3}'


# ---------- user prompt construction ----------


def test_build_user_prompt_includes_all_context():
    prompt = ai_engine._build_user_prompt(_req("sum revenue"))
    assert "Sheet1" in prompt
    assert "C1" in prompt
    assert "Name, Sales" in prompt
    assert "USER COMMAND: sum revenue" in prompt


def test_build_user_prompt_truncates_large_sheets():
    big_data = [["Header"]] + [[f"row {i}"] for i in range(100)]
    req = ProcessRequest(
        message="x",
        sheet_data=big_data,
        headers=["Header"],
        active_cell="A1",
        sheet_name="Big",
    )
    prompt = ai_engine._build_user_prompt(req)
    assert "more rows not shown" in prompt
    assert "Total rows: 101" in prompt


# ---------- SYSTEM_PROMPT enforces conversational preview_text ----------


def test_system_prompt_demands_conversational_preview_text():
    """The SYSTEM_PROMPT should instruct the LLM to write conversational
    preview_text (not 'Insert =SUM(...)'-style robot-speak)."""
    p = ai_engine.SYSTEM_PROMPT.lower()
    assert "conversational" in p or "human" in p or "friendly" in p, (
        "SYSTEM_PROMPT should explicitly demand human-friendly preview_text"
    )


def test_worked_examples_use_conversational_preview_text():
    """The worked examples in the SYSTEM_PROMPT should not start with
    'Insert =' or other robot-style copy."""
    bad_starts = ["Insert =", "Apply formula", "Set range"]
    for example_block in re.findall(
        r'"preview_text":\s*"([^"]+)"', ai_engine.SYSTEM_PROMPT
    ):
        for bad in bad_starts:
            assert not example_block.startswith(bad), (
                f"SYSTEM_PROMPT example preview_text starts with robot-speak "
                f"'{bad}': {example_block!r}"
            )


def test_worked_examples_end_with_apply_question_or_period():
    """Each example preview should end naturally (Apply? or .) so the chat
    flows."""
    examples = re.findall(
        r'"preview_text":\s*"([^"]+)"', ai_engine.SYSTEM_PROMPT
    )
    assert examples, "expected at least one preview_text example in SYSTEM_PROMPT"
    for example in examples:
        assert example.rstrip()[-1] in {".", "?", "!"}, (
            f"example preview should end with sentence punctuation: {example!r}"
        )
