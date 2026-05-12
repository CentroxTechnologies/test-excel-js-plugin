/**
 * Taskpane logic for the PowerPair sidebar.
 *
 * Two modes:
 *   1. Excel mode, running inside the Excel host. Uses real Office.js.
 *   2. Mock mode, running in a plain browser. Uses an editable mock
 *      spreadsheet rendered in the page for demos and local testing.
 *
 * Detection: Office.onReady is raced against a 2-second timeout. If Excel
 * never announces itself, we fall back to mock mode.
 *
 * The send/preview/apply flow is identical in both modes, only the
 * sheet-reading and action-application layers differ.
 */

import { computeChartPlacement } from "./chart-placement.js";

// Backend base URL. In production this moves to an env-driven config.
const BACKEND_URL = "http://localhost:8001";

// When true, the sidebar handles requests with a tiny local pattern-matcher and
// the backend is not contacted. Flip this to false to route every command
// through the FastAPI backend + real LLM. See DEV-GUIDE.md.
const USE_LOCAL_ENGINE = false;

// Cached DOM refs.
let chatEl = null;
let inputEl = null;
let sendBtn = null;

// Runtime mode flag. Flipped to true when Office.onReady doesn't fire in time.
let mockMode = false;

// Seed data used to populate the mock spreadsheet on startup.
const MOCK_SEED = [
  ["Name", "Sales", "Region", "Quarter"],
  ["Ali", 5000, "East", "Q1"],
  ["Sara", 8000, "West", "Q2"],
  ["Hassan", 3000, "East", "Q1"],
  ["Fatima", 9500, "North", "Q3"],
  ["Omar", 7200, "West", "Q2"],
  ["Ayesha", 4100, "South", "Q1"],
  ["Bilal", 6800, "East", "Q3"],
];

// Mock sheet state, the currently "selected" cell and the rendered <table> ref.
let mockActiveCell = "A1";
let mockSheetEl = null;

// Boot as soon as the DOM is ready. We start detection in parallel so a slow
// Office.onReady never blocks the UI from appearing.
if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", boot);
} else {
  boot();
}

async function boot() {
  chatEl = document.getElementById("chat");
  inputEl = document.getElementById("input");
  sendBtn = document.getElementById("send");

  sendBtn.addEventListener("click", onSendClicked);
  inputEl.addEventListener("keydown", onKeyDown);

  wireTabs();
  wireRecording();
  wireSaveModal();
  wireScheduleStub();
  renderSuggestionChips();

  const info = await detectHost(2000);
  const insideExcel = info && info.host === Office.HostType.Excel;

  if (insideExcel) {
    setBanner(true);
    renderSystemMessage(
      "Hi. Tap a chip below or type something like 'build a quarterly budget' or 'sum column sales'."
    );
  } else {
    mockMode = true;
    setBanner(false);
    initMockSheet();
    renderSystemMessage(
      "Browser preview. Tap a chip below or try 'sum column sales', 'sort by sales', or 'highlight headers in blue'."
    );
  }
}

// Resolve when Office.onReady fires with a real host; null after timeoutMs.
function detectHost(timeoutMs) {
  return new Promise((resolve) => {
    let settled = false;
    const finish = (v) => {
      if (settled) return;
      settled = true;
      resolve(v);
    };
    setTimeout(() => finish(null), timeoutMs);
    if (typeof Office !== "undefined" && Office.onReady) {
      try {
        Office.onReady((info) => {
          if (info && info.host) finish(info);
        });
      } catch (_err) {
        // Office.js failed to initialize, fall through to the timeout.
      }
    }
  });
}

function setBanner(isExcel) {
  const banner = document.getElementById("mode-banner");
  if (!banner) return;
  banner.hidden = false;
  if (isExcel) {
    banner.textContent = "Connected to Excel";
    banner.className = "mode-banner banner-excel";
  } else {
    banner.textContent =
      "Browser Preview Mode, connect to Excel for full functionality";
    banner.className = "mode-banner banner-mock";
  }
}

// ---------- Tab navigation ----------

function wireTabs() {
  document.querySelectorAll(".tab").forEach((tab) => {
    tab.addEventListener("click", () => switchView(tab.dataset.view));
  });
}

function switchView(viewName) {
  document.querySelectorAll(".tab").forEach((t) => {
    const active = t.dataset.view === viewName;
    t.classList.toggle("tab-active", active);
    t.setAttribute("aria-selected", active ? "true" : "false");
  });
  document.querySelectorAll(".view").forEach((v) => {
    const active = v.id === `view-${viewName}`;
    v.classList.toggle("view-active", active);
    v.hidden = !active;
  });
  if (viewName === "workflows") {
    refreshWorkflowsList();
  }
}

// ---------- Recording state ----------

let recording = false;
let recordedSteps = [];

function wireRecording() {
  const recordBtn = document.getElementById("record-btn");
  const stopBtn = document.getElementById("stop-recording");
  recordBtn.addEventListener("click", () => {
    if (recording) {
      handleStopClick();
    } else {
      startRecording();
    }
  });
  stopBtn.addEventListener("click", handleStopClick);
}

function handleStopClick() {
  if (recordedSteps.length === 0) {
    renderSystemMessage("Stopped recording. No actions were applied yet, so nothing to save.");
    stopRecordingDiscard();
    return;
  }
  openSaveModal();
}

function startRecording() {
  recording = true;
  recordedSteps = [];
  document.getElementById("record-btn").setAttribute("aria-pressed", "true");
  document.getElementById("recording-banner").hidden = false;
  updateRecCounter();
  renderSystemMessage("Recording started. Every action you Apply becomes a workflow step.");
}

function stopRecordingDiscard() {
  recording = false;
  recordedSteps = [];
  document.getElementById("record-btn").setAttribute("aria-pressed", "false");
  document.getElementById("recording-banner").hidden = true;
  updateRecCounter();
}

function recordAppliedAction(message, response) {
  if (!recording) return;
  if (response.action_type === "show_insight") return; // skip chat-only answers
  recordedSteps.push({
    message: message || "",
    action_type: response.action_type,
    params: response.params || {},
  });
  updateRecCounter();
}

function updateRecCounter() {
  const n = recordedSteps.length;
  const counter = document.getElementById("rec-counter");
  if (counter) counter.textContent = `${n} step${n === 1 ? "" : "s"}`;
  const stopBtn = document.getElementById("stop-recording");
  if (stopBtn) stopBtn.textContent = n === 0 ? "Stop" : "Stop & Save";
}

// ---------- Save workflow modal ----------

function wireSaveModal() {
  document.getElementById("save-cancel").addEventListener("click", closeSaveModal);
  document.getElementById("save-confirm").addEventListener("click", submitSaveModal);
  document.getElementById("save-modal").addEventListener("click", (e) => {
    if (e.target.id === "save-modal") closeSaveModal();
  });
}

function openSaveModal() {
  if (recordedSteps.length === 0) {
    renderSystemMessage("No steps recorded yet. Run a few commands and click Apply, then try again.");
    stopRecordingDiscard();
    return;
  }
  document.getElementById("save-modal-stepcount").textContent =
    `${recordedSteps.length} step${recordedSteps.length === 1 ? "" : "s"} captured`;
  document.getElementById("save-name").value = "";
  document.getElementById("save-desc").value = "";
  document.getElementById("save-modal").hidden = false;
  document.getElementById("save-name").focus();
}

function closeSaveModal() {
  document.getElementById("save-modal").hidden = true;
}

async function submitSaveModal() {
  const name = document.getElementById("save-name").value.trim();
  const description = document.getElementById("save-desc").value.trim();
  if (!name) {
    document.getElementById("save-name").focus();
    return;
  }
  if (recordedSteps.length === 0) {
    closeSaveModal();
    renderSystemMessage("Nothing to save: record at least one applied action first.");
    return;
  }
  try {
    const res = await fetch(`${BACKEND_URL}/api/workflows`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ name, description, steps: recordedSteps }),
    });
    if (!res.ok) {
      const detail = await res.text();
      throw new Error(`backend ${res.status}: ${detail}`);
    }
    const wf = await res.json();
    closeSaveModal();
    stopRecordingDiscard();
    renderSystemMessage(`Workflow saved: "${wf.name}" (${wf.steps.length} steps).`);
  } catch (err) {
    renderErrorMessage(`Couldn't save workflow: ${err.message || err}`);
  }
}

function wireScheduleStub() {
  document.querySelectorAll(".coming-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      renderAiMessage(
        "Schedule: not yet implemented. Will let you pick a saved workflow and set a daily / weekly / monthly cadence with email notifications."
      );
    });
  });
}

// ---------- Workflows view ----------

async function refreshWorkflowsList() {
  const list = document.getElementById("workflows-list");
  const empty = document.getElementById("workflows-empty");
  list.innerHTML = "";
  empty.hidden = true;

  let workflows;
  try {
    const res = await fetch(`${BACKEND_URL}/api/workflows`);
    if (!res.ok) throw new Error(`backend ${res.status}`);
    workflows = await res.json();
  } catch (err) {
    list.innerHTML = "";
    empty.hidden = false;
    empty.innerHTML = `Couldn't reach backend: <code>${err.message || err}</code>. Start it with <code>uvicorn main:app --port 8001</code> in <code>backend/</code>.`;
    return;
  }

  if (!workflows.length) {
    empty.hidden = false;
    return;
  }

  for (const wf of workflows) {
    list.appendChild(renderWorkflowCard(wf));
  }
}

function renderWorkflowCard(wf) {
  const card = document.createElement("div");
  card.className = "workflow-card";

  const name = document.createElement("div");
  name.className = "wf-name";
  name.textContent = wf.name;
  card.appendChild(name);

  if (wf.description) {
    const desc = document.createElement("div");
    desc.className = "wf-desc";
    desc.textContent = wf.description;
    card.appendChild(desc);
  }

  const meta = document.createElement("div");
  meta.className = "wf-meta";
  const created = new Date(wf.created_at).toLocaleDateString();
  meta.textContent = `${wf.steps.length} step${wf.steps.length === 1 ? "" : "s"} · saved ${created}`;
  card.appendChild(meta);

  const stepsBox = document.createElement("div");
  stepsBox.className = "workflow-steps";
  const ol = document.createElement("ol");
  for (const step of wf.steps) {
    const li = document.createElement("li");
    li.textContent = step.message || step.action_type;
    ol.appendChild(li);
  }
  stepsBox.appendChild(ol);
  card.appendChild(stepsBox);

  const actions = document.createElement("div");
  actions.className = "wf-actions";

  const runBtn = document.createElement("button");
  runBtn.type = "button";
  runBtn.className = "wf-run";
  runBtn.textContent = "Run";
  runBtn.addEventListener("click", () => runWorkflow(wf));
  actions.appendChild(runBtn);

  const deleteBtn = document.createElement("button");
  deleteBtn.type = "button";
  deleteBtn.className = "wf-delete";
  deleteBtn.textContent = "Delete";
  deleteBtn.addEventListener("click", () => deleteWorkflow(wf));
  actions.appendChild(deleteBtn);

  card.appendChild(actions);
  return card;
}

async function runWorkflow(wf) {
  switchView("chat");
  renderSystemMessage(`Running workflow "${wf.name}" (${wf.steps.length} steps)...`);
  let results;
  try {
    const res = await fetch(`${BACKEND_URL}/api/workflows/${wf.id}/run`, { method: "POST" });
    if (!res.ok) throw new Error(`backend ${res.status}`);
    results = await res.json();
  } catch (err) {
    renderErrorMessage(`Couldn't run workflow: ${err.message || err}`);
    return;
  }
  let success = 0;
  for (let i = 0; i < results.length; i++) {
    const step = results[i];
    renderSystemMessage(`Step ${i + 1}/${results.length}: ${step.preview_text || step.action_type}`);
    try {
      await applyAction(step);
      success += 1;
    } catch (err) {
      renderErrorMessage(`Step ${i + 1} failed: ${err.message || err}`);
    }
    // Small pause so the eye can follow
    await sleep(300);
  }
  renderAiMessage(`Workflow "${wf.name}" complete. ${success}/${results.length} steps applied.`);
}

async function deleteWorkflow(wf) {
  if (!confirm(`Delete workflow "${wf.name}"?`)) return;
  try {
    const res = await fetch(`${BACKEND_URL}/api/workflows/${wf.id}`, { method: "DELETE" });
    if (!res.ok && res.status !== 204) throw new Error(`backend ${res.status}`);
    refreshWorkflowsList();
  } catch (err) {
    renderErrorMessage(`Couldn't delete: ${err.message || err}`);
  }
}

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

const SUGGESTION_CHIPS = [
  "Build a quarterly budget",
  "Make a sales tracker",
  "Add a tax column at 8.5%",
  "Sum column sales",
  "Sort by revenue descending",
  "Highlight headers in blue",
  "Create a column chart",
];

function renderSuggestionChips() {
  const host = document.getElementById("suggestion-chips");
  if (!host) return;
  host.innerHTML = "";
  for (const text of SUGGESTION_CHIPS) {
    const chip = document.createElement("button");
    chip.type = "button";
    chip.className = "chip";
    chip.textContent = text;
    chip.addEventListener("click", () => {
      inputEl.value = text;
      onSendClicked();
    });
    host.appendChild(chip);
  }
}

// Enter sends, Shift+Enter makes a new line.
function onKeyDown(event) {
  if (event.key === "Enter" && !event.shiftKey) {
    event.preventDefault();
    onSendClicked();
  }
}

// Send button / Enter-key handler.
async function onSendClicked() {
  const text = inputEl.value.trim();
  if (!text) return;

  inputEl.value = "";
  renderUserMessage(text);

  const thinkingEl = renderThinking();
  sendBtn.disabled = true;

  try {
    const context = await readSheetContext();
    const response = await postToBackend(text, context);
    response._userMessage = text;
    thinkingEl.remove();
    handleBackendResponse(response);
  } catch (err) {
    thinkingEl.remove();
    renderErrorMessage(`Couldn't process that: ${err.message || err}`);
  } finally {
    sendBtn.disabled = false;
    inputEl.focus();
  }
}

// Read everything the backend needs to understand the sheet right now.
// Branches on mode, mock reads from the DOM table, Excel reads via Office.js.
async function readSheetContext() {
  if (mockMode) return readMockSheetContext();
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRangeOrNullObject(true);
    const activeCell = ctx.workbook.getActiveCell();

    usedRange.load(["values", "address"]);
    activeCell.load(["address"]);
    sheet.load(["name"]);

    await ctx.sync();

    let sheetData = [];
    let headers = [];

    if (!usedRange.isNullObject) {
      sheetData = usedRange.values || [];
      headers = sheetData.length > 0 ? sheetData[0] : [];
    }

    // activeCell.address looks like "Sheet1!B5", strip the sheet prefix.
    const rawAddress = activeCell.address || "A1";
    const bareCell = rawAddress.includes("!") ? rawAddress.split("!")[1] : rawAddress;

    return {
      sheet_data: sheetData,
      headers,
      active_cell: bareCell,
      sheet_name: sheet.name || "Sheet1",
    };
  });
}

// POST the user's command + sheet context to the backend.
// If USE_LOCAL_ENGINE is true, skip the network and run a local pattern-matcher
// instead, the plugin works without any backend for sideload demos.
async function postToBackend(message, context) {
  if (USE_LOCAL_ENGINE) {
    return localEngine(message, context);
  }

  const payload = {
    message,
    sheet_data: context.sheet_data,
    headers: context.headers,
    active_cell: context.active_cell,
    sheet_name: context.sheet_name,
  };

  const res = await fetch(`${BACKEND_URL}/api/process`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });

  if (!res.ok) {
    throw new Error(`backend returned ${res.status}`);
  }
  return res.json();
}

// ---------- Local engine (frontend-only fallback) ----------
//
// Junior dev: this is a stub. It mimics what the FastAPI backend returns so the
// plugin runs end-to-end without a server. Each branch shows the response shape
// the real backend must produce. To replace with real LLM calls, set
// USE_LOCAL_ENGINE = false at the top of this file and start the FastAPI server.

function localEngine(message, context) {
  const msg = (message || "").toLowerCase();
  const headers = context.headers || [];
  const lastRow = Math.max(context.sheet_data?.length || 0, 1);

  if (/\b(highest|largest|max|maximum)\b/.test(msg)) {
    return localMaxMin(msg, headers, context.sheet_data, true);
  }
  if (/\b(lowest|smallest|min|minimum)\b/.test(msg)) {
    return localMaxMin(msg, headers, context.sheet_data, false);
  }
  if (/\b(how many|count)\b/.test(msg)) {
    return localFormula("COUNTA", msg, headers, lastRow, context.active_cell);
  }
  if (/\b(average|mean|avg)\b/.test(msg)) {
    return localFormula("AVERAGE", msg, headers, lastRow, context.active_cell);
  }
  if (/\bsum\b/.test(msg)) {
    return localFormula("SUM", msg, headers, lastRow, context.active_cell);
  }
  if (/\b(chart|graph|plot)\b/.test(msg)) {
    return localChart(headers, lastRow);
  }
  if (/\bsort\b/.test(msg)) {
    return localSort(msg, headers, lastRow);
  }
  if (/\b(bold|format|highlight|header)\b/.test(msg)) {
    return localFormat(headers);
  }

  return localInsight(
    "Local engine doesn't understand that yet. Try: 'sum column sales', " +
      "'average column B', 'sort by name', 'format headers', 'highest in sales', or 'create a chart'."
  );
}

function localFormula(func, msg, headers, lastRow, activeCell) {
  const colIndex = guessColumn(msg, headers);
  if (colIndex === null) {
    return localInsight(
      `Tell me which column to ${func.toLowerCase()}. Example: '${func.toLowerCase()} column sales'.`
    );
  }
  const letter = columnLetter(colIndex);
  const formula = `=${func}(${letter}2:${letter}${lastRow})`;
  const headerLabel = headers[colIndex] != null ? String(headers[colIndex]) : letter;
  return {
    action_type: "insert_formula",
    params: { cell: activeCell || "A1", formula },
    preview_text: friendlyFormulaPreview(func, headerLabel, activeCell || "A1", lastRow - 1),
    confidence: 0.9,
  };
}

function friendlyFormulaPreview(func, columnLabel, targetCell, rowCount) {
  const verbMap = {
    SUM: `Total up the ${columnLabel} column (${rowCount} rows)`,
    AVERAGE: `Average the ${columnLabel} column (${rowCount} rows)`,
    COUNTA: `Count how many rows have a value in ${columnLabel}`,
  };
  const lead = verbMap[func] || `Run ${func} on ${columnLabel}`;
  return `${lead} and drop the answer into ${targetCell}. Apply?`;
}

function localMaxMin(msg, headers, sheetData, wantMax) {
  const colIndex = guessColumn(msg, headers);
  if (colIndex === null) {
    return localInsight("Tell me which column you mean. Example: 'highest in sales'.");
  }
  const numericRows = [];
  (sheetData || []).slice(1).forEach((row, i) => {
    const v = Number(row[colIndex]);
    if (Number.isFinite(v)) numericRows.push({ row: i + 2, value: v });
  });
  if (numericRows.length === 0) {
    return localInsight("Couldn't find numeric values in that column.");
  }
  const chosen = numericRows.reduce((best, cur) =>
    wantMax ? (cur.value > best.value ? cur : best) : cur.value < best.value ? cur : best
  );
  const word = wantMax ? "highest" : "lowest";
  const label = headers[colIndex] != null ? String(headers[colIndex]) : columnLetter(colIndex);
  return localInsight(`The ${word} value in '${label}' is ${chosen.value} in row ${chosen.row}.`);
}

function localChart(headers, lastRow) {
  if (!headers.length) {
    return localInsight("Add some data to the sheet before creating a chart.");
  }
  const dataRange = `A1:${columnLetter(headers.length - 1)}${lastRow}`;
  return {
    action_type: "create_chart",
    params: { data_range: dataRange, chart_type: "ColumnClustered", title: "Chart" },
    preview_text: `Build a column chart from your data and drop it next to the table. Apply?`,
    confidence: 0.8,
  };
}

function localSort(msg, headers, lastRow) {
  if (!headers.length) {
    return localInsight("Nothing to sort yet.");
  }
  const colIndex = guessColumn(msg, headers) ?? 0;
  const ascending = /ascending|asc\b/.test(msg);
  const range = `A1:${columnLetter(headers.length - 1)}${lastRow}`;
  const label = headers[colIndex] != null ? String(headers[colIndex]) : columnLetter(colIndex);
  return {
    action_type: "sort_range",
    params: { range, sort_column: colIndex, ascending },
    preview_text: `Sort the table by ${label}, ${ascending ? "smallest first" : "biggest first"}. Apply?`,
    confidence: 0.85,
  };
}

function localFormat(headers) {
  const end = columnLetter(Math.max(headers.length - 1, 0));
  const range = `A1:${end}1`;
  return {
    action_type: "format_cells",
    params: { range, bold: true, background: "#4472C4", font_color: "#FFFFFF" },
    preview_text: `Style the header row with a blue background, white bold text. Apply?`,
    confidence: 0.85,
  };
}

function localInsight(text) {
  return {
    action_type: "show_insight",
    params: { text },
    preview_text: text,
    confidence: 0.5,
  };
}

function guessColumn(msg, headers) {
  const letterMatch = msg.match(/column\s+([a-z])\b/i);
  if (letterMatch) return letterMatch[1].toUpperCase().charCodeAt(0) - 65;
  for (const token of msg.match(/[a-z_]+/gi) || []) {
    const idx = headers.findIndex(
      (h) => h != null && String(h).toLowerCase().trim() === token.toLowerCase().trim()
    );
    if (idx >= 0) return idx;
  }
  for (const token of msg.match(/[a-z_]+/gi) || []) {
    const idx = headers.findIndex(
      (h) => h != null && String(h).toLowerCase().includes(token.toLowerCase())
    );
    if (idx >= 0) return idx;
  }
  return null;
}

// Decide how to render a backend response.
function handleBackendResponse(response) {
  if (response.action_type === "show_insight") {
    renderAiMessage(response.preview_text || response.params?.text || "Done.");
    return;
  }
  renderActionCard(response);
}

// ---------- DOM builders ----------

function renderUserMessage(text) {
  const el = document.createElement("div");
  el.className = "msg msg-user";
  el.textContent = text;
  chatEl.appendChild(el);
  scrollToBottom();
}

function renderAiMessage(text) {
  const el = document.createElement("div");
  el.className = "msg msg-ai";
  el.textContent = text;
  chatEl.appendChild(el);
  scrollToBottom();
}

function renderErrorMessage(text) {
  const el = document.createElement("div");
  el.className = "msg msg-error";
  el.textContent = text;
  chatEl.appendChild(el);
  scrollToBottom();
}

function renderSystemMessage(text) {
  const el = document.createElement("div");
  el.className = "msg msg-system";
  el.textContent = text;
  chatEl.appendChild(el);
  scrollToBottom();
}

function renderThinking() {
  const el = document.createElement("div");
  el.className = "thinking";
  el.textContent = "thinking";
  chatEl.appendChild(el);
  scrollToBottom();
  return el;
}

// Preview card with Apply / Cancel buttons, the core "preview before execute" pattern.
function renderActionCard(response) {
  const card = document.createElement("div");
  card.className = "action-card";

  const label = document.createElement("div");
  label.className = "label";
  label.textContent = response.action_type.replace("_", " ");
  card.appendChild(label);

  const preview = document.createElement("div");
  preview.className = "preview";
  preview.textContent = response.preview_text || "Ready to apply.";
  card.appendChild(preview);

  const buttonRow = document.createElement("div");
  buttonRow.className = "button-row";

  const applyBtn = document.createElement("button");
  applyBtn.className = "btn btn-apply";
  applyBtn.textContent = "Apply";

  const cancelBtn = document.createElement("button");
  cancelBtn.className = "btn btn-cancel";
  cancelBtn.textContent = "Cancel";

  applyBtn.addEventListener("click", async () => {
    applyBtn.disabled = true;
    cancelBtn.disabled = true;
    try {
      await applyAction(response);
      recordAppliedAction(response._userMessage, response);
      preview.textContent = `Done. ${response.preview_text}`;
      label.textContent = "applied";
      label.style.color = "#1a7f37";
      buttonRow.remove();
    } catch (err) {
      renderErrorMessage(`Excel rejected that action: ${err.message || err}`);
      applyBtn.disabled = false;
      cancelBtn.disabled = false;
    }
  });

  cancelBtn.addEventListener("click", () => {
    renderSystemMessage("Action cancelled.");
    buttonRow.remove();
    label.textContent = "cancelled";
    label.style.color = "#6a737d";
  });

  buttonRow.appendChild(applyBtn);
  buttonRow.appendChild(cancelBtn);
  card.appendChild(buttonRow);

  chatEl.appendChild(card);
  scrollToBottom();
}

function scrollToBottom() {
  chatEl.scrollTop = chatEl.scrollHeight;
}

// ---------- Action dispatcher ----------

// Translate a validated backend action into actual Excel.run operations.
// In mock mode the same action is rendered visually into the DOM table.
async function applyAction(response) {
  if (mockMode) return applyMockAction(response);
  switch (response.action_type) {
    case "insert_formula":
      return applyInsertFormula(response.params);
    case "write_values":
      return applyWriteValues(response.params);
    case "format_cells":
      return applyFormatCells(response.params);
    case "create_chart":
      return applyCreateChart(response.params);
    case "sort_range":
      return applySortRange(response.params);
    default:
      throw new Error(`unknown action_type: ${response.action_type}`);
  }
}

async function applyInsertFormula({ cell, formula }) {
  await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(cell);
    range.formulas = [[formula]];
    await ctx.sync();
  });
}

async function applyWriteValues({ start_cell, values }) {
  await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    // Size the target range to match the 2D values array.
    const rows = values.length;
    const cols = values[0]?.length || 0;
    if (rows === 0 || cols === 0) {
      throw new Error("no values to write");
    }
    const start = sheet.getRange(start_cell);
    const target = start.getResizedRange(rows - 1, cols - 1);
    target.values = values;
    await ctx.sync();
  });
}

async function applyFormatCells({ range, bold, background, font_color }) {
  await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const target = sheet.getRange(range);
    if (typeof bold === "boolean") target.format.font.bold = bold;
    if (background) target.format.fill.color = background;
    if (font_color) target.format.font.color = font_color;
    await ctx.sync();
  });
}

async function applyCreateChart({ data_range, chart_type, title }) {
  await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRangeOrNullObject(true);
    usedRange.load(["address"]);
    await ctx.sync();

    const source = sheet.getRange(data_range);
    const chart = sheet.charts.add(chart_type, source, Excel.ChartSeriesBy.auto);
    if (title) chart.title.text = title;

    // Place the chart away from the data so it doesn't sit on top of the table.
    const usedAddress = usedRange.isNullObject ? data_range : usedRange.address;
    const bareUsed = String(usedAddress).includes("!")
      ? usedAddress.split("!")[1]
      : usedAddress;
    const placement = computeChartPlacement(bareUsed);
    const tl = sheet.getRange(placement.topLeft);
    const br = sheet.getRange(placement.bottomRight);
    chart.setPosition(tl, br);

    await ctx.sync();
  });
}

async function applySortRange({ range, sort_column, ascending }) {
  await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const target = sheet.getRange(range);
    target.sort.apply(
      [{ key: sort_column, ascending: Boolean(ascending) }],
      false, // matchCase
      true,  // hasHeaders, we assume row 1 is headers
      Excel.SortOrientation.rows
    );
    await ctx.sync();
  });
}

// ---------- Mock mode: sheet rendering, reading, and action execution ----------

// Build the DOM table from seed data and show the mock-sheet container.
function initMockSheet() {
  const container = document.getElementById("mock-sheet-container");
  container.hidden = false;
  mockSheetEl = document.getElementById("mock-sheet");
  renderMockSheet(MOCK_SEED);
}

// Render a 2D array into the mock table. Called on init and after sorts.
function renderMockSheet(data) {
  mockSheetEl.innerHTML = "";
  const colCount = Math.max(...data.map((r) => r.length));

  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  headRow.appendChild(document.createElement("th")); // top-left corner
  for (let c = 0; c < colCount; c++) {
    const th = document.createElement("th");
    th.textContent = columnLetter(c);
    headRow.appendChild(th);
  }
  thead.appendChild(headRow);
  mockSheetEl.appendChild(thead);

  const tbody = document.createElement("tbody");
  data.forEach((row, rIdx) => {
    const tr = document.createElement("tr");
    const rowNum = document.createElement("td");
    rowNum.className = "row-header";
    rowNum.textContent = String(rIdx + 1);
    tr.appendChild(rowNum);
    for (let c = 0; c < colCount; c++) {
      const td = document.createElement("td");
      td.contentEditable = "true";
      td.spellcheck = false;
      const value = row[c];
      td.textContent = value === undefined || value === null ? "" : String(value);
      td.dataset.row = String(rIdx);
      td.dataset.col = String(c);
      td.dataset.ref = columnLetter(c) + (rIdx + 1);
      td.addEventListener("focus", () => setMockActiveCell(td.dataset.ref));
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  });
  mockSheetEl.appendChild(tbody);

  applyActiveCellHighlight();
}

// Mark a single cell as "active" (the target for insert_formula and the
// cursor-position equivalent that Office.js would report).
function setMockActiveCell(ref) {
  mockActiveCell = ref;
  applyActiveCellHighlight();
}

function applyActiveCellHighlight() {
  if (!mockSheetEl) return;
  mockSheetEl.querySelectorAll("td[data-ref]").forEach((td) => {
    td.classList.toggle("active-cell", td.dataset.ref === mockActiveCell);
  });
}

// Read the current state of the mock table and return the same shape the
// backend receives from the Excel-mode reader.
function readMockSheetContext() {
  const rows = [];
  mockSheetEl.querySelectorAll("tbody tr").forEach((tr) => {
    const row = [];
    tr.querySelectorAll("td[data-col]").forEach((td) => {
      row.push(coerceCellValue(td.textContent));
    });
    rows.push(row);
  });
  const headers = rows.length > 0 ? rows[0] : [];
  return {
    sheet_data: rows,
    headers,
    active_cell: mockActiveCell,
    sheet_name: "MockSheet",
  };
}

// Parse a cell's text into a number when possible, otherwise keep the string.
function coerceCellValue(raw) {
  const trimmed = (raw || "").trim();
  if (trimmed === "") return "";
  const num = Number(trimmed);
  return Number.isFinite(num) && trimmed !== "" && /^-?\d/.test(trimmed)
    ? num
    : trimmed;
}

// Dispatch a backend action to the matching mock handler.
async function applyMockAction(response) {
  switch (response.action_type) {
    case "insert_formula":
      return applyMockInsertFormula(response.params);
    case "write_values":
      return applyMockWriteValues(response.params);
    case "format_cells":
      return applyMockFormatCells(response.params);
    case "sort_range":
      return applyMockSortRange(response.params);
    case "create_chart":
      renderAiMessage(
        "Chart would be created here. Real Excel charts can't render in browser preview."
      );
      return;
    default:
      throw new Error(`unknown action_type: ${response.action_type}`);
  }
}

// Show the literal formula text in the target cell and flash it green.
function applyMockInsertFormula({ cell, formula }) {
  const td = findMockCell(cell);
  if (!td) throw new Error(`cell ${cell} is outside the mock sheet`);
  td.textContent = formula;
  td.title = "Formula, would compute inside real Excel.";
  flashCell(td);
}

// Overwrite a contiguous range starting at start_cell with the provided 2D values.
function applyMockWriteValues({ start_cell, values }) {
  const start = refToIndex(start_cell);
  if (!start) throw new Error(`invalid start_cell ${start_cell}`);
  values.forEach((row, dr) => {
    row.forEach((val, dc) => {
      const ref = columnLetter(start.col + dc) + (start.row + dr + 1);
      const td = findMockCell(ref);
      if (!td) return;
      td.textContent = val === null || val === undefined ? "" : String(val);
      flashCell(td);
    });
  });
}

// Apply inline CSS to every cell in the expanded range.
function applyMockFormatCells({ range, bold, background, font_color }) {
  expandRange(range).forEach((ref) => {
    const td = findMockCell(ref);
    if (!td) return;
    if (bold) td.style.fontWeight = "bold";
    if (background) td.style.background = background;
    if (font_color) td.style.color = font_color;
    flashCell(td);
  });
}

// Sort the data rows (row 1 is treated as headers) by the given column.
function applyMockSortRange({ sort_column, ascending }) {
  const { sheet_data } = readMockSheetContext();
  if (sheet_data.length < 2) return;
  const header = sheet_data[0];
  const body = sheet_data.slice(1);
  body.sort((a, b) => compareCells(a[sort_column], b[sort_column], ascending));
  renderMockSheet([header, ...body]);
}

function compareCells(a, b, ascending) {
  const bothNumeric = typeof a === "number" && typeof b === "number";
  const cmp = bothNumeric
    ? a - b
    : String(a ?? "").localeCompare(String(b ?? ""));
  return ascending ? cmp : -cmp;
}

// ---------- Small spreadsheet utilities ----------

function columnLetter(index) {
  let n = index;
  let letters = "";
  while (true) {
    letters = String.fromCharCode(65 + (n % 26)) + letters;
    n = Math.floor(n / 26) - 1;
    if (n < 0) break;
  }
  return letters;
}

function refToIndex(ref) {
  const match = /^([A-Z]+)(\d+)$/.exec(ref);
  if (!match) return null;
  let col = 0;
  for (let i = 0; i < match[1].length; i++) {
    col = col * 26 + (match[1].charCodeAt(i) - 64);
  }
  return { row: parseInt(match[2], 10) - 1, col: col - 1 };
}

function expandRange(range) {
  const parts = range.split(":");
  if (parts.length === 1) return [parts[0]];
  const start = refToIndex(parts[0]);
  const end = refToIndex(parts[1]);
  if (!start || !end) return [];
  const refs = [];
  for (let r = start.row; r <= end.row; r++) {
    for (let c = start.col; c <= end.col; c++) {
      refs.push(columnLetter(c) + (r + 1));
    }
  }
  return refs;
}

function findMockCell(ref) {
  if (!mockSheetEl) return null;
  return mockSheetEl.querySelector(`td[data-ref="${ref}"]`);
}

function flashCell(td) {
  td.classList.remove("flash-green");
  void td.offsetWidth; // force reflow so the animation replays
  td.classList.add("flash-green");
}
