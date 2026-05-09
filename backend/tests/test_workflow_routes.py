"""
Integration tests for the workflow API routes.

We use FastAPI's TestClient (httpx-based). The `isolated_workflows_dir`
fixture from conftest points WORKFLOWS_DATA_DIR at a tmp dir so each test
starts from a clean state.
"""

from __future__ import annotations

import pytest
from fastapi.testclient import TestClient

from main import app


@pytest.fixture
def client(isolated_workflows_dir):
    return TestClient(app)


def _payload(name: str = "demo") -> dict:
    return {
        "name": name,
        "description": "test",
        "steps": [
            {
                "message": "format headers",
                "action_type": "format_cells",
                "params": {"range": "A1:D1", "bold": True},
            },
            {
                "message": "sum sales",
                "action_type": "insert_formula",
                "params": {"cell": "B9", "formula": "=SUM(B2:B8)"},
            },
        ],
    }


def test_get_workflows_empty(client):
    r = client.get("/api/workflows")
    assert r.status_code == 200
    assert r.json() == []


def test_post_workflow_returns_201_with_id(client):
    r = client.post("/api/workflows", json=_payload())
    assert r.status_code == 201
    body = r.json()
    assert body["id"].startswith("wf_")
    assert body["name"] == "demo"
    assert len(body["steps"]) == 2


def test_post_workflow_rejects_blank_name(client):
    bad = _payload()
    bad["name"] = "  "
    r = client.post("/api/workflows", json=bad)
    assert r.status_code == 400


def test_post_workflow_rejects_empty_steps(client):
    bad = _payload()
    bad["steps"] = []
    r = client.post("/api/workflows", json=bad)
    assert r.status_code == 400


def test_get_workflow_by_id(client):
    created = client.post("/api/workflows", json=_payload()).json()
    r = client.get(f"/api/workflows/{created['id']}")
    assert r.status_code == 200
    assert r.json()["id"] == created["id"]


def test_get_workflow_404(client):
    r = client.get("/api/workflows/wf_nope")
    assert r.status_code == 404


def test_put_workflow_updates_name(client):
    created = client.post("/api/workflows", json=_payload()).json()
    r = client.put(
        f"/api/workflows/{created['id']}",
        json={"name": "renamed"},
    )
    assert r.status_code == 200
    assert r.json()["name"] == "renamed"


def test_put_workflow_404(client):
    r = client.put("/api/workflows/wf_nope", json={"name": "x"})
    assert r.status_code == 404


def test_delete_workflow_204_then_404(client):
    created = client.post("/api/workflows", json=_payload()).json()
    r = client.delete(f"/api/workflows/{created['id']}")
    assert r.status_code == 204
    r2 = client.get(f"/api/workflows/{created['id']}")
    assert r2.status_code == 404


def test_delete_unknown_workflow_404(client):
    r = client.delete("/api/workflows/wf_nope")
    assert r.status_code == 404


def test_run_workflow_returns_steps_in_order(client):
    created = client.post("/api/workflows", json=_payload()).json()
    r = client.post(f"/api/workflows/{created['id']}/run")
    assert r.status_code == 200
    results = r.json()
    assert len(results) == 2
    assert results[0]["action_type"] == "format_cells"
    assert results[1]["action_type"] == "insert_formula"


def test_run_unknown_workflow_404(client):
    r = client.post("/api/workflows/wf_nope/run")
    assert r.status_code == 404


def test_full_round_trip_create_list_get_update_delete(client):
    """End-to-end: create, list, get, update, delete."""
    # Create
    created = client.post("/api/workflows", json=_payload()).json()
    wf_id = created["id"]

    # List
    listed = client.get("/api/workflows").json()
    assert any(wf["id"] == wf_id for wf in listed)

    # Get
    one = client.get(f"/api/workflows/{wf_id}").json()
    assert one["name"] == "demo"

    # Update steps
    new_steps = [
        {
            "message": "highest in sales",
            "action_type": "show_insight",
            "params": {"text": "biggest is 9500 in row 5"},
        }
    ]
    updated = client.put(
        f"/api/workflows/{wf_id}",
        json={"steps": new_steps},
    ).json()
    assert len(updated["steps"]) == 1
    assert updated["steps"][0]["action_type"] == "show_insight"

    # Delete
    assert client.delete(f"/api/workflows/{wf_id}").status_code == 204
    assert client.get("/api/workflows").json() == []
