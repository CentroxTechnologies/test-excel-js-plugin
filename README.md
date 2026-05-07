# PowerPair

[![CI](https://github.com/CentroxTechnologies/test-excel-js-plugin/actions/workflows/ci.yml/badge.svg?branch=main)](https://github.com/CentroxTechnologies/test-excel-js-plugin/actions/workflows/ci.yml)
[![License: Proprietary](https://img.shields.io/badge/license-Proprietary-red.svg)](./LICENSE)
[![Office Add-in](https://img.shields.io/badge/Office.js-Add--in-217346?logo=microsoftexcel&logoColor=white)](https://learn.microsoft.com/office/dev/add-ins/)
[![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue.svg?logo=python&logoColor=white)](https://www.python.org/)
[![FastAPI](https://img.shields.io/badge/FastAPI-009688?logo=fastapi&logoColor=white)](https://fastapi.tiangolo.com/)
[![Status](https://img.shields.io/badge/status-POC-orange.svg)](#whats-next-tayyab--start-here)

> Simple like ChatGPT, accurate like a formula, safe like an audit trail.

PowerPair is an Office.js Excel add-in. Type plain English in the sidebar — the AI returns a structured action (a formula, a chart, a sort, a bulk write) which Excel previews and then applies. The AI never computes numbers itself; it generates Excel formulas and lets Excel compute. Workflows, scheduling, and a recording feature are on the roadmap (see "What's next" below).

This repo is the working scaffold for the Phase-1 MVP.

---

## Install in Excel Online (anyone with a Microsoft account)

The add-in is hosted on GitHub Pages. To use it:

1. Save the manifest from this URL (right-click → Save link as `manifest.xml`):
   <https://raw.githubusercontent.com/CentroxTechnologies/test-excel-js-plugin/main/addin/manifest.xml>
2. Go to <https://office.com>, sign in, open Excel for the web, create a blank workbook.
3. Use the search bar at the top (`Alt + Q`) → type "add-in" → click **Get Add-ins**.
4. In the dialog, click **More Add-ins** → **Upload My Add-in** → pick the `manifest.xml` you saved.
5. **Home** tab → click **Open PowerPair**. The sidebar opens on the right.

If the upload option is missing, your Microsoft account is a personal tier — Microsoft removed sideload from personal accounts on Excel for Web. Workaround: get a free Microsoft 365 Developer Program tenant at <https://developer.microsoft.com/microsoft-365/dev-program> and sign in with the sandbox admin account.

---

## Try the demo (60 seconds)

Add some sample data first — `Name`, `Sales`, `Region`, `Quarter` columns work great. Or click any **suggestion chip** above the input to skip the setup.

Demo commands worth running first:

| Command | What you'll see |
|---|---|
| `Build me a quarterly budget template starting at A1` | 11 rows × 6 cols of categories + formulas appear on screen |
| `Make a sales tracker template` | 6 rows of seed data with line totals + grand total |
| `Add a "Q4 vs Q1 growth %" column` | New column with growth-percentage formulas |
| `Highlight the Net profit row in green` | Single row repaints |
| `Make a column chart of all four quarters` | Real Excel chart pops in |
| `Sort by Total descending` | Rows reorder live |
| Click the **💾 Save as Workflow** button | Roadmap message: workflows + scheduling are Phase 2 |

Each AI action shows a preview card — review it, click **Apply**, watch the sheet update.

---

## How it works (the 30-second tour)

```
You type a command in the sidebar
        │
        ▼
taskpane.js reads sheet context (headers, data, active cell)
        │
        ▼
POST /api/process to FastAPI backend (or local pattern-matcher if USE_LOCAL_ENGINE=true)
        │
        ▼
ai_engine.py builds a prompt, calls GPT-4o (or Claude), parses JSON response
        │
        ▼
validator.py checks the action is structurally sane
        │
        ▼
Sidebar renders a preview card with Apply / Cancel
        │
        ▼  on Apply
Office.js writes the formula / values / format / chart back to Excel
```

Six action types the AI can emit: `insert_formula`, `write_values`, `format_cells`, `create_chart`, `sort_range`, `show_insight`. Adding a seventh is a 4-step recipe documented in [DEV-GUIDE.md](./DEV-GUIDE.md).

---

## Run the backend locally (for real LLM responses)

The plugin works **without a backend** — it has a frontend pattern-matcher (the "local engine") that handles ~7 canned commands (sum, average, count, max, min, sort, format, chart). For natural-language commands like "build me a quarterly budget" you need the real backend.

### One-time setup

```bash
cd backend
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
```

Open `backend/.env` and paste your OpenAI key in place of `your-openai-key-here`. Set `AI_PROVIDER=anthropic` if you'd rather use Claude.

### Run

```bash
uvicorn main:app --reload --port 8001
```

`{"status":"ok"}` from `curl http://localhost:8001/api/health` confirms it's up.

### Switch the plugin to use the backend

Two paths depending on whether you want the deployed plugin or your local copy to talk to the backend:

**Option A — local dev loop (fastest):**
1. `cd addin && npm install && npm run cert && npm start` — serves the add-in over HTTPS on `localhost:3000`
2. Edit `addin/manifest.xml`, swap the `https://centroxtechnologies.github.io/test-excel-js-plugin` URLs back to `https://localhost:3000` (or use a separate local-dev manifest)
3. Sideload that manifest in Excel
4. In `addin/src/taskpane/taskpane.js`, set `USE_LOCAL_ENGINE = false`
5. Refresh the sidebar — every command now hits your backend

**Option B — keep using the deployed plugin:**
1. Set `USE_LOCAL_ENGINE = false` in `addin/src/taskpane/taskpane.js`
2. Push the change to the `gh-pages` branch (the live site updates in ~30 sec)
3. Backend must be reachable from the deployed plugin. Localhost works from HTTPS pages on most browsers (special exception); if it breaks, run `ngrok http 8001` and put the ngrok URL in `BACKEND_URL` at the top of `taskpane.js`

---

## Local engine vs real backend

| Mode | Where it runs | Good for | Limits |
|---|---|---|---|
| `USE_LOCAL_ENGINE = true` | Browser only | Quick demos, offline testing, no API costs | Only 7 hardcoded patterns; doesn't understand "build me a budget" |
| `USE_LOCAL_ENGINE = false` | FastAPI backend → OpenAI/Anthropic | Full natural-language commands, bulk write_values, chained reasoning | Needs backend running + API key + key has credits |

Flag lives at the top of `addin/src/taskpane/taskpane.js`.

---

## Project layout

```
officejs/
├── README.md                          # ← you are here
├── DEV-GUIDE.md                       # how to extend the codebase
├── start.sh                           # boots backend + addin server in one command
├── addin/                             # the Office.js add-in
│   ├── manifest.xml                   # Office Add-in manifest, points to GitHub Pages URL
│   ├── package.json                   # http-server + dev cert tooling
│   └── src/
│       ├── taskpane/
│       │   ├── taskpane.html          # sidebar shell
│       │   ├── taskpane.js            # UI + Office.js + local engine
│       │   └── taskpane.css
│       ├── commands/                  # ribbon command handler
│       └── assets/                    # icon PNGs
└── backend/                           # FastAPI service
    ├── main.py                        # /api/health + /api/process
    ├── ai_engine.py                   # Provider abstraction (OpenAI, Anthropic), system prompt
    ├── validator.py                   # validates AI output before sheet write
    ├── models.py                      # Pydantic request/response models
    ├── requirements.txt
    └── .env.example                   # copy to .env and add your API key
```

---

## What's next (Tayyab — start here)

This scaffold is Phase-1 / Week-1-2 of the spec. Still to build, in priority order:

1. **Saved workflows** — record a sequence of commands as a named macro, replay on new data, edit later. The killer feature vs Microsoft Copilot.
2. **Scheduling** — cron triggers + email notification flow (Office.js can't run when Excel is closed)
3. **Real validation pipeline** — current `validator.py` does ~30% of the spec. Missing reference checks, type checks, range completeness, test execution, reasonableness checks.
4. **Multi-LLM routing** — task classifier picks GPT-mini for simple, Claude for complex, Gemini for large data, Azure OpenAI for sensitive
5. **Audit trail** — DB-persisted log of every action for finance/compliance buyers
6. **Confidence-driven UX** — auto-apply for high-confidence, sidebar-only for low-confidence
7. **Edge cases** — merged cells, multiple header rows, hidden rows, pivot tables, mixed-type columns

How to add a new action type, switch providers, etc.: see [DEV-GUIDE.md](./DEV-GUIDE.md).

---

## Troubleshooting

**Sidebar shows blank page or "can't reach server".** First time only: open `https://centroxtechnologies.github.io/test-excel-js-plugin/taskpane/taskpane.html` in a browser to confirm the page loads. If it doesn't, GitHub Pages may still be deploying — wait 30 sec and retry.

**Sidebar shows "Couldn't process that".** Backend isn't reachable. Either flip `USE_LOCAL_ENGINE = true` in `taskpane.js` (and redeploy gh-pages) or start the backend with `uvicorn main:app --port 8001`.

**API key not configured.** Open `backend/.env`, replace `your-openai-key-here` with a real key from <https://platform.openai.com/api-keys>, then restart the backend.

**Manifest validation fails when uploading.** `cd addin && npm run validate` to see specifics. The `<Id>` GUID in `manifest.xml` must be unique per add-in — if you fork and deploy a separate copy, run `uuidgen` and replace it.

**CORS errors in the browser console.** The backend allows `*` by default; if you've changed it, add the GitHub Pages origin (`https://centroxtechnologies.github.io`) back to `allow_origins` in `backend/main.py`.

**Local engine doesn't recognize a command.** Expected — it only handles ~7 patterns. To get full natural-language support, switch to the real backend (`USE_LOCAL_ENGINE = false`) per the steps above.

**Personal Microsoft account can't see "Upload My Add-in".** Microsoft removed sideload for personal accounts on Excel for Web. Use a Microsoft 365 Developer Program sandbox tenant (free), or use Excel Desktop instead.
