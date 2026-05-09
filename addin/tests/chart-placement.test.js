/**
 * Tests for chart-placement helper.
 *
 * Run: node --test addin/tests/chart-placement.test.js
 *      (from project root)
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import {
  computeChartPlacement,
  parseRange,
  refOf,
} from "../src/taskpane/chart-placement.js";

test("parseRange handles single cell address", () => {
  assert.deepEqual(parseRange("A1"), {
    startCol: 0,
    startRow: 0,
    endCol: 0,
    endRow: 0,
  });
});

test("parseRange handles A1:F12", () => {
  assert.deepEqual(parseRange("A1:F12"), {
    startCol: 0,
    startRow: 0,
    endCol: 5,
    endRow: 11,
  });
});

test("parseRange handles AA1:AB10 (multi-letter columns)", () => {
  assert.deepEqual(parseRange("AA1:AB10"), {
    startCol: 26,
    startRow: 0,
    endCol: 27,
    endRow: 9,
  });
});

test("refOf converts col/row indices back to A1 notation", () => {
  assert.equal(refOf({ col: 0, row: 0 }), "A1");
  assert.equal(refOf({ col: 7, row: 0 }), "H1");
  assert.equal(refOf({ col: 25, row: 51 }), "Z52");
  assert.equal(refOf({ col: 26, row: 0 }), "AA1");
});

test("empty/null used range falls back to placement at C1", () => {
  const placement = computeChartPlacement(null);
  assert.equal(placement.topLeft, "C1");
});

test("typical small range (A1:F12) places chart 2 cols right at H1", () => {
  const placement = computeChartPlacement("A1:F12");
  assert.equal(placement.topLeft, "H1");
  // Chart should be at least 8 cols wide and 15 rows tall (default size)
  const tl = parseRange(placement.topLeft);
  const br = parseRange(placement.bottomRight);
  assert.ok(br.endCol - tl.startCol >= 7, "chart should be ~8 cols wide");
  assert.ok(br.endRow - tl.startRow >= 14, "chart should be ~15 rows tall");
});

test("tall thin data (A1:B100) places chart at D1 to the right", () => {
  const placement = computeChartPlacement("A1:B100");
  assert.equal(placement.topLeft, "D1");
});

test("wide short data (A1:Z5) falls back to below-data placement", () => {
  // endCol=25, +2 = 27 which is past column 20, so should fall back vertical
  const placement = computeChartPlacement("A1:Z5");
  // Chart should be BELOW the data, starting at column A
  const tl = parseRange(placement.topLeft);
  assert.equal(tl.startCol, 0, "should start at column A when falling back vertically");
  assert.ok(tl.startRow > 5, "should start below the data row");
});

test("single cell active (A1) places chart at C1", () => {
  const placement = computeChartPlacement("A1");
  assert.equal(placement.topLeft, "C1");
});

test("placement is always a valid two-cell range string", () => {
  for (const range of ["A1:F12", "A1:B100", "A1:Z5", "A1", null, "AA1:AB10"]) {
    const { topLeft, bottomRight } = computeChartPlacement(range);
    assert.match(topLeft, /^[A-Z]+\d+$/);
    assert.match(bottomRight, /^[A-Z]+\d+$/);
  }
});
