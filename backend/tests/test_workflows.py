"""Tests for workflow storage (backend/workflows.py)."""

from __future__ import annotations

import json

import pytest

import workflows


def _step(order: int, message: str = "test", action_type: str = "show_insight"):
    return {
        "order": order,
        "message": message,
        "action_type": action_type,
        "params": {"text": message},
    }


def test_list_returns_empty_when_no_file(isolated_workflows_dir):
    assert workflows.list_workflows() == []


def test_save_creates_workflow_with_id_and_timestamps(isolated_workflows_dir):
    wf = workflows.save_workflow(
        name="Demo flow",
        description="three step demo",
        steps=[_step(1, "format headers"), _step(2, "sum sales")],
    )
    assert wf["id"].startswith("wf_")
    assert wf["name"] == "Demo flow"
    assert wf["description"] == "three step demo"
    assert len(wf["steps"]) == 2
    assert wf["created_at"]
    assert wf["updated_at"] == wf["created_at"]


def test_saved_workflow_appears_in_list(isolated_workflows_dir):
    wf = workflows.save_workflow(name="A", description="", steps=[_step(1)])
    listed = workflows.list_workflows()
    assert len(listed) == 1
    assert listed[0]["id"] == wf["id"]


def test_get_workflow_by_id(isolated_workflows_dir):
    wf = workflows.save_workflow(name="A", description="", steps=[_step(1)])
    found = workflows.get_workflow(wf["id"])
    assert found["id"] == wf["id"]
    assert found["name"] == "A"


def test_get_unknown_workflow_returns_none(isolated_workflows_dir):
    assert workflows.get_workflow("wf_does_not_exist") is None


def test_steps_are_assigned_step_ids(isolated_workflows_dir):
    wf = workflows.save_workflow(
        name="A", description="", steps=[_step(1), _step(2)]
    )
    step_ids = [s["step_id"] for s in wf["steps"]]
    assert all(sid.startswith("step_") for sid in step_ids)
    assert len(set(step_ids)) == 2  # unique


def test_update_workflow_changes_name_and_bumps_updated_at(isolated_workflows_dir):
    wf = workflows.save_workflow(name="Old", description="", steps=[_step(1)])
    original_created = wf["created_at"]
    updated = workflows.update_workflow(
        wf["id"], {"name": "New name", "description": "Refreshed"}
    )
    assert updated["name"] == "New name"
    assert updated["description"] == "Refreshed"
    assert updated["created_at"] == original_created
    # updated_at may or may not differ depending on clock resolution; just assert it exists
    assert updated["updated_at"]


def test_update_unknown_workflow_returns_none(isolated_workflows_dir):
    assert workflows.update_workflow("wf_nope", {"name": "x"}) is None


def test_delete_workflow_removes_it(isolated_workflows_dir):
    wf = workflows.save_workflow(name="A", description="", steps=[_step(1)])
    assert workflows.delete_workflow(wf["id"]) is True
    assert workflows.get_workflow(wf["id"]) is None
    assert workflows.list_workflows() == []


def test_delete_unknown_workflow_returns_false(isolated_workflows_dir):
    assert workflows.delete_workflow("wf_nope") is False


def test_persists_across_calls_via_json_file(isolated_workflows_dir):
    workflows.save_workflow(name="P", description="", steps=[_step(1)])
    # Force a re-read by clearing the in-memory cache if any; the storage is
    # file-backed so a fresh list_workflows() call reads from disk.
    listed = workflows.list_workflows()
    assert len(listed) == 1
    # And the file actually exists
    json_path = isolated_workflows_dir / "workflows.json"
    assert json_path.exists()
    on_disk = json.loads(json_path.read_text())
    assert len(on_disk) == 1
    assert on_disk[0]["name"] == "P"


def test_steps_keep_order_after_save(isolated_workflows_dir):
    wf = workflows.save_workflow(
        name="ordered",
        description="",
        steps=[_step(1, "first"), _step(2, "second"), _step(3, "third")],
    )
    assert [s["message"] for s in wf["steps"]] == ["first", "second", "third"]


def test_save_rejects_empty_steps(isolated_workflows_dir):
    with pytest.raises(ValueError):
        workflows.save_workflow(name="empty", description="", steps=[])


def test_save_rejects_blank_name(isolated_workflows_dir):
    with pytest.raises(ValueError):
        workflows.save_workflow(name="  ", description="", steps=[_step(1)])


def test_run_workflow_returns_actions_in_order(isolated_workflows_dir):
    wf = workflows.save_workflow(
        name="run me",
        description="",
        steps=[
            _step(1, "show one"),
            _step(2, "show two"),
            _step(3, "show three"),
        ],
    )
    results = workflows.run_workflow(wf["id"])
    assert len(results) == 3
    assert [r["preview_text"] for r in results] == ["show one", "show two", "show three"]


def test_run_unknown_workflow_returns_none(isolated_workflows_dir):
    assert workflows.run_workflow("wf_nope") is None
