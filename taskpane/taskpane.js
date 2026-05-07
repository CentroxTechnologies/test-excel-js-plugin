/**
 * Taskpane logic for the Excel AI Assistant sidebar.
 *
 * Two modes:
 *   1. Excel mode — running inside the Excel host. Uses real Office.js.
 *   2. Mock mode — running in a plain browser. Uses an editable mock
 *      spreadsheet rendered in the page for demos and local testing.
 *
 * Detection: Office.onReady is raced against a 2-second timeout. If Excel
 * never announces itself, we fall back to mock mode.
 *
 * The send/preview/apply flow is identical in both modes — only the
 * sheet-reading and action-application layers differ.
 */

// Backend base URL. In production this moves to an env-driven config.
const BACKEND_URL = "http://localhost:8001";

// When true, the sidebar handles requests with a tiny local pattern-matcher and
// the backend is not contacted. Junior dev: flip this to false once the FastAPI
// backend is running and configured with an API key. See DEV-GUIDE.md.
const USE_LOCAL_ENGINE = true;

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

// Mock sheet state — the currently "selected" cell and the rendered <table> ref.
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

  const info = await detectHost(2000);
  const insideExcel = info && info.host === Office.HostType.Excel;

  if (insideExcel) {
    setBanner(true);
    renderSystemMessage(
      "Hi. Try 'sum column revenue', 'sort by name', or 'highest value in revenue'."
    );
  } else {
    mockMode = true;
    setBanner(false);
    initMockSheet();
    renderSystemMessage(
      "Browser preview. Try 'sum column sales', 'sort by sales', 'format headers', or 'highest in sales'."
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
        // Office.js failed to initialize — fall through to the timeout.
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
      "Browser Preview Mode — connect to Excel for full functionality";
    banner.className = "mode-banner banner-mock";
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
// Branches on mode — mock reads from the DOM table, Excel reads via Office.js.
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

    // activeCell.address looks like "Sheet1!B5" — strip the sheet prefix.
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
// instead — the plugin works without any backend for sideload demos.
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
    preview_text: `Insert ${formula} into ${activeCell || "A1"} to ${func.toLowerCase()} column '${headerLabel}'.`,
    confidence: 0.9,
  };
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
    preview_text: `Create a column chart from ${dataRange}.`,
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
    preview_text: `Sort ${range} by '${label}' (${ascending ? "ascending" : "descending"}).`,
    confidence: 0.85,
  };
}

function localFormat(headers) {
  const end = columnLetter(Math.max(headers.length - 1, 0));
  const range = `A1:${end}1`;
  return {
    action_type: "format_cells",
    params: { range, bold: true, background: "#4472C4", font_color: "#FFFFFF" },
    preview_text: `Make ${range} bold, blue background, white text.`,
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

// Preview card with Apply / Cancel buttons — the core "preview before execute" pattern.
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
    const source = sheet.getRange(data_range);
    const chart = sheet.charts.add(chart_type, source, Excel.ChartSeriesBy.auto);
    if (title) chart.title.text = title;
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
      true,  // hasHeaders — we assume row 1 is headers
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
  td.title = "Formula — would compute inside real Excel.";
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
