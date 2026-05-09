"""
FastAPI entrypoint.

Routes:
  GET  /api/health                       smoke test
  POST /api/process                      single AI command execution
  GET  /api/workflows                    list saved workflows
  POST /api/workflows                    create a new workflow
  GET  /api/workflows/{id}               fetch one workflow
  PUT  /api/workflows/{id}               update name / description / steps
  DELETE /api/workflows/{id}             remove a workflow
  POST /api/workflows/{id}/run           replay a workflow

Run locally: `uvicorn main:app --reload --port 8001`
"""

from dotenv import load_dotenv
load_dotenv()

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware

import workflows
from ai_engine import generate_action
from models import (
    ProcessRequest,
    ProcessResponse,
    Workflow,
    WorkflowCreate,
    WorkflowRunResult,
    WorkflowUpdate,
)
from validator import validate_action


app = FastAPI(title="PowerPair", version="0.1.0")


# Wide-open CORS for development. Tighten this for production (specific origins + methods).
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/api/health")
def health() -> dict[str, str]:
    """Smoke test endpoint. Returns 200 with a small JSON payload."""
    return {"status": "ok"}


@app.post("/api/process", response_model=ProcessResponse)
def process(req: ProcessRequest) -> ProcessResponse:
    """
    Main entrypoint for single-command processing.

    1. Run the AI engine to decide what action to perform.
    2. Validate the chosen action (structural + reference + type + completeness).
    3. The add-in previews the action and only executes it on user click.
    """
    action = generate_action(req)
    return validate_action(
        action,
        sheet_data=req.sheet_data,
        headers=req.headers,
        user_message=req.message,
    )


# ---------- Workflows ----------


@app.get("/api/workflows", response_model=list[Workflow])
def list_workflows_route() -> list[dict]:
    return workflows.list_workflows()


@app.post("/api/workflows", response_model=Workflow, status_code=201)
def create_workflow(payload: WorkflowCreate) -> dict:
    try:
        return workflows.save_workflow(
            name=payload.name,
            description=payload.description,
            steps=[s.model_dump() for s in payload.steps],
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.get("/api/workflows/{workflow_id}", response_model=Workflow)
def get_workflow_route(workflow_id: str) -> dict:
    wf = workflows.get_workflow(workflow_id)
    if wf is None:
        raise HTTPException(status_code=404, detail="workflow not found")
    return wf


@app.put("/api/workflows/{workflow_id}", response_model=Workflow)
def update_workflow_route(workflow_id: str, payload: WorkflowUpdate) -> dict:
    updates: dict = {}
    if payload.name is not None:
        updates["name"] = payload.name
    if payload.description is not None:
        updates["description"] = payload.description
    if payload.steps is not None:
        updates["steps"] = [s.model_dump() for s in payload.steps]
    try:
        wf = workflows.update_workflow(workflow_id, updates)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    if wf is None:
        raise HTTPException(status_code=404, detail="workflow not found")
    return wf


@app.delete("/api/workflows/{workflow_id}", status_code=204)
def delete_workflow_route(workflow_id: str) -> None:
    if not workflows.delete_workflow(workflow_id):
        raise HTTPException(status_code=404, detail="workflow not found")


@app.post(
    "/api/workflows/{workflow_id}/run",
    response_model=list[WorkflowRunResult],
)
def run_workflow_route(workflow_id: str) -> list[dict]:
    results = workflows.run_workflow(workflow_id)
    if results is None:
        raise HTTPException(status_code=404, detail="workflow not found")
    return results
