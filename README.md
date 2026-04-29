# Excel AI Assistant — Foundation

An Office.js Excel add-in paired with a Python FastAPI backend. The user types a plain-English command in the sidebar ("sum column revenue"), the backend decides what to do, and the add-in executes it in Excel.

This repo is the **foundation skeleton**. It ships with:

- A FastAPI backend that pattern-matches commands to actions (placeholder for a real LLM).
- An Office.js sidebar with a chat UI, preview cards, and Apply / Cancel flow.
- The end-to-end wiring: sheet context → backend → validated action → Excel.

Three things are **not yet built** but have clear placeholders:

- OpenAI / Claude integration (`backend/ai_engine.py` has the planned prompt structure in comments).
- Saved workflows / macros.
- Scheduled runs.

---

## The golden rule

**The AI does not compute values.** It emits Excel formulas (`=SUM(B2:B100)`) and lets Excel compute them. The only exception is `show_insight` answers ("highest value is 95000 in row 7"), where the backend reads `sheet_data` directly because there's no formula that returns a text answer into a chat.

Keep this rule in mind when extending the engine.

---

## Prerequisites

- **Node.js 18+** (`node --version`)
- **Python 3.10+** (`python3 --version`)
- **pip** (comes with Python)
- **A Microsoft 365 account** — free dev tenant works fine. See next section.

### Getting a Microsoft 365 developer account

1. Go to <https://developer.microsoft.com/en-us/microsoft-365/dev-program>.
2. Click **Join now** and sign in with any Microsoft account (or create one).
3. Pick **Instant sandbox** — takes about 60 seconds and gives you an `@your-tenant.onmicrosoft.com` account with Excel Online.
4. Save the admin email + password somewhere safe. You'll use this account to sideload the add-in.

---

## Setup (first run)

Easiest path: run `./start.sh` from the repo root. It does venv + npm install + cert generation + launches both servers.

If you prefer doing it manually:

### 1. Backend

```bash
cd backend
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn main:app --reload --port 8001
```

Verify: `curl http://localhost:8001/api/health` → `{"status":"ok"}`

### 2. Frontend

```bash
cd addin
npm install
npm run cert         # installs localhost HTTPS certs (one time)
npm start            # serves src/ on https://localhost:3000
```

Visit `https://localhost:3000/taskpane/taskpane.html` once in your browser and accept the cert warning — Excel Online won't load the add-in otherwise.

---

## Sideloading the add-in into Excel Online

1. Sign into <https://www.office.com> with your Microsoft 365 account.
2. Open Excel Online (**Excel** from the app launcher) and create a new blank workbook.
3. Put some sample data in so you have something to play with. Example:

   | Name    | Revenue |
   |---------|---------|
   | Alpha   | 12000   |
   | Beta    | 8500    |
   | Gamma   | 23000   |
   | Delta   | 4200    |

4. On the ribbon, click the **Insert** tab → **Office Add-ins** → **Upload My Add-in**.
5. Browse to `addin/manifest.xml` in this repo and upload it.
6. Excel will add an **AI Assistant** group on the **Home** tab with an **Open Assistant** button. Click it. The sidebar opens.

---

## Testing walkthrough

With the sample data above and the sidebar open:

1. Click cell **C1** (so the AI knows where to put the result).
2. In the sidebar, type `sum column revenue` and press Enter.
3. You'll see a preview card: *"Insert =SUM(B2:B5) into C1 to sum column 'Revenue'."*
4. Click **Apply**. Cell C1 now contains the formula; it displays 47700.

Try a few more:

- `average column revenue` → AVERAGE formula
- `highest value in revenue` → chat-only insight ("23000 in row 4")
- `sort by revenue` → sorts the used range descending
- `format headers` → bolds row 1 with a blue background
- `create a chart` → column chart from the used range
- `how many rows` → COUNTA formula

If you type something the mock engine doesn't recognize, it replies with a help message listing what it can do.

---

## Project layout

```
officejs/
├── backend/
│   ├── main.py           # FastAPI app + CORS + routes
│   ├── ai_engine.py      # Mock pattern-matcher + LLM placeholder comments
│   ├── validator.py      # Structural validation before shipping actions to the add-in
│   ├── models.py         # Pydantic request/response classes
│   └── requirements.txt
├── addin/
│   ├── manifest.xml      # Office Add-in manifest (points to localhost:3000)
│   ├── package.json      # Dev deps: http-server, office-addin-dev-certs
│   └── src/
│       ├── taskpane/     # Sidebar HTML/CSS/JS
│       ├── commands/     # Ribbon command handler
│       └── assets/       # Icon PNGs (served at /assets/...)
├── start.sh              # One-command dev bootstrap
└── README.md
```

---

## Troubleshooting

### The sidebar shows a blank page or "can't reach server"

- Did you visit `https://localhost:3000/taskpane/taskpane.html` in your browser and accept the cert? Excel Online won't load self-signed content until the browser trusts it.
- Is `npm start` actually running? `curl -k https://localhost:3000/taskpane/taskpane.html` should return HTML.

### CORS errors in the browser console

- The backend allows `*` by default. If you changed it, put your add-in origin back on the allowlist.
- If you're running the backend somewhere other than `localhost:8001`, update `BACKEND_URL` at the top of `addin/src/taskpane/taskpane.js`.

### Certificate errors

- Re-run `npm run cert` from the `addin/` folder. The certs land in `~/.office-addin-dev-certs/`.
- On some Linux distros you may also need to trust the root CA manually:
  `sudo cp ~/.office-addin-dev-certs/ca.crt /usr/local/share/ca-certificates/office-addin-dev.crt && sudo update-ca-certificates`

### Manifest validation fails when uploading

- Run `cd addin && npm run validate` to see specific errors.
- The GUID in `<Id>` must be unique per add-in. If you fork and deploy another copy, generate a new one with `uuidgen`.

### Office.js "Office is not defined" errors

- The taskpane must be loaded through HTTPS, not `file://` or plain `http`. Always go through `https://localhost:3000/...`.
- Check the Network tab — `office.js` should load from `appsforoffice.microsoft.com` with a 200 response.

### Actions fail with "InvalidArgument" or "ItemNotFound"

- The active cell might be on a different sheet than the used range, or the range references a cell that doesn't exist yet. Select a cell inside or near your data before sending commands.

---

## What to build next

When you're ready to move past the skeleton:

1. **Real LLM call** — replace the body of `generate_action()` in `ai_engine.py` with the provider call described in the big comment block.
2. **Saved workflows** — store `{name, command_sequence, bindings}` records; add a `/workflows` route group.
3. **Scheduled runs** — a lightweight cron or APScheduler setup that replays a saved workflow against a specified workbook on a fixed cadence.

Each feature is its own phase. Don't bolt them onto this foundation in one go — keep phases isolated so the skeleton stays junior-readable.
