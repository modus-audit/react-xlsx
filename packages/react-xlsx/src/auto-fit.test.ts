import assert from "node:assert/strict";
import { test } from "node:test";
import { autoFitSizes, type FitSheet, fittedColumnWidth, fittedRowHeight } from "./auto-fit.ts";

// 1px per character keeps the arithmetic readable.
const measure = (text: string) => text.length;

function sheet(
  cells: Record<string, { text: string; wrap?: boolean; size?: number }>,
  merged: string[] = [],
  spans: Record<string, { rowSpan: number; colSpan: number }> = {}
): FitSheet {
  return {
    usedRange: () => [0, 0, 3, 3],
    getFormattedValueAt: (row, col) => cells[`${row}:${col}`]?.text ?? "",
    getCellStyleAt: (row, col) => {
      const cell = cells[`${row}:${col}`];
      return { font: { size: cell?.size ?? 12 }, alignment: { wrapText: cell?.wrap === true } };
    },
    getMergeSpan: (row, col) => spans[`${row}:${col}`] ?? null,
    isMergedSecondary: (row, col) => merged.includes(`${row}:${col}`)
  };
}

test("a column fits its widest unwrapped text, ignoring merged cells", () => {
  const cells = { "0:1": { text: "x".repeat(40) }, "1:1": { text: "x".repeat(90) }, "2:1": { text: "x".repeat(300) } };
  assert.equal(fittedColumnWidth(sheet(cells, ["2:1"]), 1, measure), 110);
});

test("a column keeps the padding spaces accounting formats draw", () => {
  assert.equal(fittedColumnWidth(sheet({ "0:0": { text: " $(150,078)" } }), 0, measure), 11 + 20);
});

test("a row counts wrapped lines at the column's width; an empty row has no fit", () => {
  // 12pt = 16px; three 30px words in a 50px-wide column (42px after padding) wrap to 3 lines.
  const words = ["a".repeat(30), "b".repeat(30), "c".repeat(30)].join(" ");
  assert.equal(fittedRowHeight(sheet({ "1:0": { text: words, wrap: true } }), 1, () => 50, measure), Math.ceil(3 * 16 * 1.2 + 2));
  assert.equal(fittedRowHeight(sheet({ "0:2": { text: "Total" } }), 0, () => 80, measure), Math.ceil(16 * 1.2 + 2));
  assert.equal(fittedRowHeight(sheet({}), 0, () => 80, measure), null);
});

test("a row wraps a merge across its columns at the merge's width, and skips taller merges", () => {
  const words = Array.from({ length: 30 }, () => "word").join(" ");
  const cells = { "0:0": { text: words, wrap: true } };
  const across = sheet(cells, ["0:1", "0:2"], { "0:0": { rowSpan: 1, colSpan: 3 } });
  // 149 chars in a 3 x 50px merge (142px of text) wrap to 2 lines; one 42px column would take 4.
  assert.equal(fittedRowHeight(across, 0, () => 50, measure), Math.ceil(2 * 16 * 1.2 + 2));
  const tall = sheet(cells, [], { "0:0": { rowSpan: 2, colSpan: 1 } });
  assert.equal(fittedRowHeight(tall, 0, () => 50, measure), null);
});

test("autofit skips hidden and already-fitting indices, resets empty rows and keeps empty columns", () => {
  const fitSheet = sheet({ "0:0": { text: "x".repeat(40) }, "0:1": { text: "x".repeat(40) } });
  const base = { sheet: fitSheet, columnWidthPx: () => 64, rowHeightPx: () => 40, defaultRowHeightPx: 20, gridlinePx: 1, measure };
  assert.deepEqual(autoFitSizes({ ...base, axis: "column", indices: [0, 1, 2, 0], hidden: new Set([1]) }), [{ index: 0, sizePx: 60 }]);
  // Row 0 fits its text; row 2 is empty and returns to the default height.
  assert.deepEqual(autoFitSizes({ ...base, axis: "row", indices: [0, 2], hidden: new Set() }), [
    { index: 0, sizePx: Math.ceil(16 * 1.2 + 2) },
    { index: 2, sizePx: 21 }
  ]);
});
