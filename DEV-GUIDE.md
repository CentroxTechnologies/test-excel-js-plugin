# Developer Guide, PowerPair scaffold

You're picking up a working Office.js add-in scaffold. This doc explains what's here, how to run it, and where to plug in real logic.

> Out of scope for this guide: the product vision, pricing, client requirements. See `/home/adeel/centrox/projects-overview/excel-project/SUMMARY/` for those.

---

## TL;DR, run it in 3 commands

```bash
cd addin
npm install            # one-time
npm run cert           # one-time, installs HTTPS dev certs
npm start              # serves the add-in on https://localhost:3000
```

The plugin works **without a backend**. There's a local pattern-matcher in `taskpane.js` that fakes the AI for sideload demos.

To sideload into Excel Online:

1. Open `https://localhost:3000/taskpane/taskpane.html` in your browser, accept the cert warning
2. Sign into [office.com](https://office.com) → open Excel → new blank workbook
3. Ribbon → **Insert** → **Office Add-ins** → **Upload My Add-in** → pick `addin/manifest.xml`
4. **Home** tab → **Open Assistant**
5. Type `sum column sales` → preview card → **Apply**

You should see the formula land in the active cell. Done.

---

## Project layout

```
officejs/
├── addin/                             # The add-in (sideloads into Excel)
│   ├── manifest.xml                   # Office Add-in manifest, points to localhost:3000
│   ├── package.json                   # http-server + dev cert tooling
│   └── src/
│       ├── taskpane/
│       │   ├── taskpane.html          # Sidebar shell
│       │   ├── taskpane.js            # ALL the UI + Excel + local engine logic
│       │   └── taskpane.css
│       ├── commands/                  # Ribbon button handler
│       └── assets/                    # Icons
├── backend/                           # FastAPI service, OPTIONAL for now
│   ├── main.py                        # Two routes: /api/health and /api/process
│   ├── ai_engine.py                   # Real LLM provider abstraction
│   ├── validator.py                   # Validates AI output before it reaches the add-in
│   ├── models.py                      # Pydantic request/response models
│   ├── requirements.txt
│   └── .env.example                   # Copy to .env and add a real API key
├── start.sh                           # Boots backend + addin server together
├── DEV-GUIDE.md                       # ← you are here
└── README.md
```

---

## How a request flows through the code

```
User types "sum column sales" in the sidebar
        │
        ▼
taskpane.js / onSendClicked()
        │
        ▼  reads sheet context (headers, data, active cell)
readSheetContext()  ── Excel.run if sideloaded, DOM table if browser preview
        │
        ▼
postToBackend(message, context)
        │
        ▼  if USE_LOCAL_ENGINE === true:
localEngine() pattern-matches the message and returns a fake response
        │
        ▼  otherwise:
fetch http://localhost:8001/api/process
        │
        ▼  backend:
ai_engine.generate_action(req)        →  validator.validate_action(action)
        │
        ▼  response shape:
{
  "action_type": "insert_formula",
  "params": { "cell": "C1", "formula": "=SUM(B2:B9)" },
  "preview_text": "Insert =SUM(B2:B9) into C1 to sum column 'Sales'.",
  "confidence": 0.9
}
        │
        ▼
handleBackendResponse() → renderActionCard() shows Apply / Cancel
        │
        ▼  user clicks Apply
applyAction(response) → applyInsertFormula() → Excel.run() writes formula
```

---

## Two modes you should know about

### Excel mode (sideloaded)
Office.js detects the host and loads. `Excel.run` is used to read/write cells.

### Mock browser mode
If you open the taskpane URL in a plain browser (no Excel), Office.js never announces itself. The 2s timeout fires, `mockMode = true`, and a DOM table renders with editable cells. Every Apply path has both an `applyXxx()` (Office.js) and an `applyMockXxx()` (DOM) implementation. Useful for fast iteration without sideloading every change.

---

## How to add a new action type

The system supports six actions today: `insert_formula`, `write_values`, `format_cells`, `create_chart`, `show_insight`, `sort_range`. To add a seventh (e.g. `apply_filter`), touch these four spots:

### 1. Backend system prompt
`backend/ai_engine.py`, append the new action to the `Action types and params:` block in `SYSTEM_PROMPT`. Give it a one-line params example. The LLM only knows what's in this prompt.

### 2. Backend validator
`backend/validator.py`, add `_validate_apply_filter(params)`, register it in the `_VALIDATORS` map. Return `_reject("...")` for anything malformed; return `None` if it's fine.

### 3. Frontend Excel handler
`addin/src/taskpane/taskpane.js`, add `applyApplyFilter(params)` using `Excel.run`, and add a `case "apply_filter":` in `applyAction()`.

### 4. Frontend mock handler
Same file, add `applyMockApplyFilter(params)` and a `case "apply_filter":` in `applyMockAction()` so it works in browser preview too.

### 5. Local engine (optional)
If you want the new action to work without the backend, add a branch in `localEngine()` in `taskpane.js`.

---

## Switching from local engine to real backend

When you're ready to use real LLMs:

1. `cd backend && python3 -m venv .venv && source .venv/bin/activate && pip install -r requirements.txt`
2. `cp .env.example .env` then paste a real `OPENAI_API_KEY` (or `ANTHROPIC_API_KEY` and flip `AI_PROVIDER=anthropic`)
3. `uvicorn main:app --reload --port 8001` (or just `./start.sh` from the project root)
4. In `addin/src/taskpane/taskpane.js`, set `USE_LOCAL_ENGINE = false`
5. Refresh the add-in (close + reopen the sidebar in Excel)

The sidebar will now route every command through the backend → LLM → validator → back to Excel.

---

## How to add a new LLM provider

`backend/ai_engine.py` has an `AIProvider` ABC with two implementations: `OpenAIProvider`, `AnthropicProvider`. To add Gemini:

1. Create `GeminiProvider(AIProvider)` with `name`, `model`, `call(system_prompt, user_prompt) -> tuple[str, dict]`. Use `google-generativeai` SDK.
2. Add a branch in `get_provider()` for `AI_PROVIDER=gemini`.
3. Add `GEMINI_API_KEY=...` and `GEMINI_MODEL=...` to `.env.example` and `.env`.
4. Add `google-generativeai` to `requirements.txt`.

The system prompt and user prompt are provider-agnostic, no other changes needed.

---

## Testing

There are no automated tests yet. Manual checks:

| Check | How |
|---|---|
| Local engine works | `npm start`, open browser preview, type each canned command |
| Backend health | `curl http://localhost:8001/api/health` returns `{"status":"ok"}` |
| Backend with placeholder key | `curl -X POST http://localhost:8001/api/process -H 'Content-Type: application/json' -d '{"message":"x","sheet_data":[],"headers":[],"active_cell":"A1","sheet_name":"S"}'` returns a `show_insight` about missing key |
| Sideload | Upload `manifest.xml` in Excel Online; sidebar opens; commands work |

---

## What's NOT built yet (the real work)

This scaffold maps to Phase 1 / Week 1-2 of the spec. Still to build:

1. **Saved workflows**, record sequences of actions, name them, replay on new data, edit. The killer feature vs Copilot.
2. **Scheduling**, cron + email notifications (Office.js can't run when Excel is closed)
3. **Multi-LLM routing**, task classifier picks GPT-mini for simple, Claude for complex, Gemini for large data, Azure OpenAI for sensitive
4. **Real validation pipeline**, current validator does ~30% of the spec. Missing reference checks, type checks, range completeness, test execution, reasonableness
5. **Audit trail**, DB-persisted log of every action with user, model, prompt, formula, result
6. **Confidence-driven UX**, auto-apply for high-confidence, sidebar-only for low-confidence
7. **Edge cases**, merged cells, multiple header rows, hidden rows, pivot tables, mixed types

Full spec: `/home/adeel/centrox/projects-overview/excel-project/CLIENT_REQUIREMENTS.md`.
Sequenced backlog: `/home/adeel/centrox/projects-overview/excel-project/SUMMARY/03-NEXT-STEPS.md`.

---

## Troubleshooting

**Sidebar shows blank page.** Open `https://localhost:3000/taskpane/taskpane.html` in your browser and accept the cert. Excel won't load self-signed otherwise.

**`Office is not defined` errors.** The taskpane must load over HTTPS, not file://. Always go through `https://localhost:3000/...`.

**Manifest validation fails.** `cd addin && npm run validate` to see specifics. The `<Id>` GUID must be unique per add-in instance.

**`Excel.run` throws InvalidArgument.** Active cell is on a different sheet than the data. Click a cell inside your data before sending.

**Local engine doesn't recognize a command.** That's expected, it only handles ~7 patterns. To add more, edit `localEngine()` in `taskpane.js`. To get full natural-language support, switch to the real backend (`USE_LOCAL_ENGINE = false`).
