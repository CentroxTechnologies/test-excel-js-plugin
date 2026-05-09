"""
Workflow storage for PowerPair.

Saved workflows are stored as a JSON array in `data/workflows.json`. Each
workflow has a list of steps, where each step is a captured AI action
(action_type + params) that can be replayed deterministically without
re-calling the LLM.

The data directory location is controlled by the `WORKFLOWS_DATA_DIR` env
var (defaults to `<this file's dir>/data`). Tests use the env var to point
at a temporary directory.

This module is intentionally simple. When we add scheduling and audit
trail in a later phase, the storage layer should swap to SQLite while
keeping this same surface API.
"""

from __future__ import annotations

import json
import os
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


_FILE_NAME = "workflows.json"


def _data_dir() -> Path:
    override = os.getenv("WORKFLOWS_DATA_DIR")
    if override:
        return Path(override)
    return Path(__file__).parent / "data"


def _data_file() -> Path:
    return _data_dir() / _FILE_NAME


def _now_iso() -> str:
    return datetime.now(tz=timezone.utc).isoformat(timespec="seconds")


def _new_id(prefix: str) -> str:
    return f"{prefix}_{uuid.uuid4().hex[:12]}"


def _read_all() -> list[dict[str, Any]]:
    path = _data_file()
    if not path.exists():
        return []
    try:
        text = path.read_text(encoding="utf-8")
        if not text.strip():
            return []
        return json.loads(text)
    except json.JSONDecodeError:
        # Corrupt file is treated as empty so the app keeps working; admin
        # can inspect the file by hand.
        return []


def _write_all(workflows: list[dict[str, Any]]) -> None:
    path = _data_file()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(workflows, indent=2), encoding="utf-8")


def list_workflows() -> list[dict[str, Any]]:
    """Return all saved workflows in insertion order."""
    return _read_all()


def get_workflow(workflow_id: str) -> dict[str, Any] | None:
    for wf in _read_all():
        if wf["id"] == workflow_id:
            return wf
    return None


def save_workflow(
    name: str,
    description: str,
    steps: list[dict[str, Any]],
) -> dict[str, Any]:
    """Create a new workflow and persist it."""
    if not name or not name.strip():
        raise ValueError("workflow name is required")
    if not steps:
        raise ValueError("workflow needs at least one step")

    now = _now_iso()
    wf = {
        "id": _new_id("wf"),
        "name": name.strip(),
        "description": description or "",
        "created_at": now,
        "updated_at": now,
        "steps": [_normalize_step(s, idx + 1) for idx, s in enumerate(steps)],
        "schedule": None,
    }
    all_workflows = _read_all()
    all_workflows.append(wf)
    _write_all(all_workflows)
    return wf


def update_workflow(workflow_id: str, updates: dict[str, Any]) -> dict[str, Any] | None:
    """Update name / description / steps. Returns the updated workflow or None."""
    all_workflows = _read_all()
    for wf in all_workflows:
        if wf["id"] != workflow_id:
            continue
        if "name" in updates:
            new_name = (updates["name"] or "").strip()
            if not new_name:
                raise ValueError("name cannot be blank")
            wf["name"] = new_name
        if "description" in updates:
            wf["description"] = updates["description"] or ""
        if "steps" in updates:
            steps = updates["steps"]
            if not steps:
                raise ValueError("workflow needs at least one step")
            wf["steps"] = [
                _normalize_step(s, idx + 1) for idx, s in enumerate(steps)
            ]
        wf["updated_at"] = _now_iso()
        _write_all(all_workflows)
        return wf
    return None


def delete_workflow(workflow_id: str) -> bool:
    all_workflows = _read_all()
    remaining = [wf for wf in all_workflows if wf["id"] != workflow_id]
    if len(remaining) == len(all_workflows):
        return False
    _write_all(remaining)
    return True


def run_workflow(workflow_id: str) -> list[dict[str, Any]] | None:
    """
    Replay a workflow.

    For now: returns each step's stored action as-is, in order. The frontend
    is responsible for executing them against the active sheet via Office.js.
    No LLM is called during replay (steps are deterministic by construction).
    """
    wf = get_workflow(workflow_id)
    if wf is None:
        return None
    results: list[dict[str, Any]] = []
    for step in sorted(wf["steps"], key=lambda s: s.get("order", 0)):
        results.append(
            {
                "step_id": step.get("step_id"),
                "action_type": step.get("action_type"),
                "params": step.get("params", {}),
                "preview_text": step.get("message", ""),
            }
        )
    return results


def _normalize_step(step: dict[str, Any], order: int) -> dict[str, Any]:
    return {
        "step_id": step.get("step_id") or _new_id("step"),
        "order": step.get("order") or order,
        "message": step.get("message") or "",
        "action_type": step["action_type"],
        "params": step.get("params") or {},
    }
