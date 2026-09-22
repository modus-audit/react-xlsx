import assert from "node:assert/strict";
import { test } from "node:test";
import { findDataBoundary } from "./data-navigation.ts";

const sheet = { maxUsedRow: 20, maxUsedCol: 20 };
const cells = new Set([1, 2, 3, 7, 8, 12]);
for (const vertical of [true, false]) {
  const cell = (position: number) => vertical ? { row: position, col: 0 } : { row: 0, col: position };
  const forward = vertical ? "ArrowDown" : "ArrowRight";
  const backward = vertical ? "ArrowUp" : "ArrowLeft";
  const run = (start: number, reverse = false, hidden: number[] = []) => findDataBoundary({ cell: cell(start), direction: reverse ? backward : forward, maxRow: 20, maxCol: 20 }, { ...sheet, [vertical ? "hiddenRows" : "hiddenCols"]: hidden }, (row, col) => cells.has(vertical ? row : col));
  test(`data regions and blank gaps on ${vertical ? "rows" : "columns"}`, () => {
    assert.deepEqual(run(1), cell(3));
    assert.deepEqual(run(3), cell(7));
    assert.deepEqual(run(5), cell(7));
    assert.deepEqual(run(7), cell(8));
    assert.deepEqual(run(8), cell(12));
    assert.deepEqual(run(12), cell(20));
    assert.deepEqual(run(20), cell(20));
    assert.deepEqual(run(8, true), cell(7));
    assert.deepEqual(run(7, true), cell(3));
    assert.deepEqual(run(1, true), cell(0));
    assert.deepEqual(run(0, true), cell(0));
  });
  test(`skips hidden ${vertical ? "rows" : "columns"}`, () => {
    assert.deepEqual(run(3, false, [7]), cell(8));
    assert.deepEqual(run(1, false, [2]), cell(3));
  });
}
