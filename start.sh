#!/usr/bin/env bash
# One-command dev bootstrap: start the FastAPI backend and the Office.js dev server.
#
# Both run in the foreground so Ctrl+C stops everything. Logs are interleaved;
# prefix each line with a tag so you can tell them apart.

set -e

ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"

# --- Sanity checks ----------------------------------------------------------

if [ ! -d "$ROOT_DIR/backend/.venv" ]; then
  echo "[setup] creating Python venv..."
  python3 -m venv "$ROOT_DIR/backend/.venv"
  "$ROOT_DIR/backend/.venv/bin/pip" install --upgrade pip >/dev/null
  "$ROOT_DIR/backend/.venv/bin/pip" install -r "$ROOT_DIR/backend/requirements.txt"
fi

if [ ! -d "$ROOT_DIR/addin/node_modules" ]; then
  echo "[setup] installing add-in dependencies..."
  (cd "$ROOT_DIR/addin" && npm install)
fi

if [ ! -f "$HOME/.office-addin-dev-certs/localhost.crt" ]; then
  echo "[setup] installing HTTPS dev certificates..."
  (cd "$ROOT_DIR/addin" && npm run cert)
fi

# --- Launch -----------------------------------------------------------------

echo ""
echo "Backend : http://localhost:8001"
echo "Frontend: https://localhost:3000"
echo ""
echo "Press Ctrl+C to stop both servers."
echo ""

PIDS=()
trap 'echo ""; echo "stopping..."; kill ${PIDS[@]} 2>/dev/null; exit 0' INT TERM

(
  cd "$ROOT_DIR/backend"
  .venv/bin/uvicorn main:app --port 8001 --reload 2>&1 | sed 's/^/[backend] /'
) &
PIDS+=($!)

(
  cd "$ROOT_DIR/addin"
  npm start 2>&1 | sed 's/^/[addin]   /'
) &
PIDS+=($!)

wait
