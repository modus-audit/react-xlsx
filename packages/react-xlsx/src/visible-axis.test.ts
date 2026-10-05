import assert from "node:assert/strict";
import { test } from "node:test";
import { lastSheetIndex, visibleSpan } from "./visible-axis.ts";

// Rows 0-2 and 6 hidden, as when a sheet's top rows are hidden.
const visible = [3, 4, 5, 7, 8];

test("a range with hidden edges draws over its shown rows", () => {
  assert.deepEqual(visibleSpan(visible, 0, 8), [0, 4]);
  assert.deepEqual(visibleSpan(visible, 0, 6), [0, 2]);
  assert.deepEqual(visibleSpan(visible, 4, 7), [1, 3]);
});

test("a range of only hidden rows has nothing to draw", () => {
  assert.equal(visibleSpan(visible, 0, 2), undefined);
  assert.equal(visibleSpan(visible, 6, 6), undefined);
});

test("select-all reaches hidden rows past the last shown one", () => {
  assert.equal(lastSheetIndex(8, [0, 1, 2, 6]), 8);
  assert.equal(lastSheetIndex(969, [0, 1, 2860]), 2860);
  assert.equal(lastSheetIndex(5, undefined), 5);
});
