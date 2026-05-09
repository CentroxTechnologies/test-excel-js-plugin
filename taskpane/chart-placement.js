/**
 * Smart chart placement.
 *
 * Given the used-range address of the active sheet, return where to drop the
 * new chart so it doesn't sit on top of existing data.
 *
 * Strategy:
 *   1. Parse the used range into col / row indices (0-based).
 *   2. First choice: 2 columns right of the last data column, at the start row.
 *   3. If the right-side placement would push past column ~20 (still on screen
 *      for typical viewports), fall back to placing below the data with a 1-row
 *      gap, anchored at the start column.
 *   4. Default chart size: 8 columns wide, 15 rows tall.
 *
 * Pure function. Imports cleanly under both the browser (taskpane.js) and Node
 * (tests). No DOM, no Office.js deps.
 */

const CHART_WIDTH_COLS = 8;
const CHART_HEIGHT_ROWS = 15;
const HORIZONTAL_GAP = 2;
const VERTICAL_GAP = 2;
const RIGHT_FALLBACK_THRESHOLD = 20;

const RANGE_RE = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/;

export function parseRange(address) {
  const match = RANGE_RE.exec(String(address || "").toUpperCase());
  if (!match) {
    throw new Error(`invalid range: ${address}`);
  }
  const [, c1, r1, c2, r2] = match;
  const startCol = colLetterToIndex(c1);
  const startRow = parseInt(r1, 10) - 1;
  const endCol = c2 != null ? colLetterToIndex(c2) : startCol;
  const endRow = r2 != null ? parseInt(r2, 10) - 1 : startRow;
  return { startCol, startRow, endCol, endRow };
}

export function refOf({ col, row }) {
  return colIndexToLetter(col) + (row + 1);
}

function colLetterToIndex(letters) {
  let n = 0;
  for (const ch of letters) {
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n - 1;
}

function colIndexToLetter(index) {
  let n = index;
  let letters = "";
  while (true) {
    letters = String.fromCharCode(65 + (n % 26)) + letters;
    n = Math.floor(n / 26) - 1;
    if (n < 0) break;
  }
  return letters;
}

export function computeChartPlacement(usedRangeAddress) {
  if (!usedRangeAddress) {
    return placementAt({ col: 2, row: 0 });
  }
  let parsed;
  try {
    parsed = parseRange(usedRangeAddress);
  } catch (_err) {
    return placementAt({ col: 2, row: 0 });
  }
  const { startCol, startRow, endCol, endRow } = parsed;

  // Right of the data?
  const rightTopLeft = { col: endCol + HORIZONTAL_GAP, row: startRow };
  if (rightTopLeft.col <= RIGHT_FALLBACK_THRESHOLD) {
    return placementAt(rightTopLeft);
  }

  // Otherwise below the data
  const belowTopLeft = { col: startCol, row: endRow + VERTICAL_GAP };
  return placementAt(belowTopLeft);
}

function placementAt(topLeft) {
  const bottomRight = {
    col: topLeft.col + CHART_WIDTH_COLS - 1,
    row: topLeft.row + CHART_HEIGHT_ROWS - 1,
  };
  return {
    topLeft: refOf(topLeft),
    bottomRight: refOf(bottomRight),
  };
}
