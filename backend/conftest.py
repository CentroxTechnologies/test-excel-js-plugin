"""Shared pytest fixtures for the backend test suite."""

from __future__ import annotations

import os
import sys
from pathlib import Path

# Make backend/ importable as the test root.
sys.path.insert(0, str(Path(__file__).parent))

import pytest


@pytest.fixture
def isolated_workflows_dir(tmp_path, monkeypatch):
    """Point WORKFLOWS_DIR at a tmp dir so tests never touch the real data file."""
    data_dir = tmp_path / "data"
    data_dir.mkdir()
    monkeypatch.setenv("WORKFLOWS_DATA_DIR", str(data_dir))
    return data_dir


@pytest.fixture(autouse=True)
def _clear_provider_env(monkeypatch):
    """Strip API keys from env unless a test sets them. Keeps ai_engine in
    its 'no key configured' branch by default so we never hit a real LLM."""
    for var in ("OPENAI_API_KEY", "ANTHROPIC_API_KEY"):
        monkeypatch.delenv(var, raising=False)
    monkeypatch.setenv("AI_PROVIDER", "openai")
