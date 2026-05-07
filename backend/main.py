"""
FastAPI entrypoint. Two routes: health check and command processing.

Run locally: `uvicorn main:app --reload --port 8001`
"""

from dotenv import load_dotenv
load_dotenv()

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from models import ProcessRequest, ProcessResponse
from ai_engine import generate_action
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
    Main endpoint.

    1. Run the mock AI engine to decide what the user wants.
    2. Validate the chosen action before sending it back.
    3. The add-in previews the action and only executes it on user click.
    """
    action = generate_action(req)
    return validate_action(action)
