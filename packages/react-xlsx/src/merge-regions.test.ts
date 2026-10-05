import assert from "node:assert/strict";
import { test } from "node:test";
import { mergesTouching, mergeTarget } from "./merge-regions.ts";

const regions = [
  { startRow: 11, startCol: 0, endRow: 12, endCol: 1, range: "A12:B13" },
  { startRow: 0, startCol: 3, endRow: 0, endCol: 4, range: "D1:E1" }
];

test("selecting a merged block's top-left cell finds the whole block", () => {
  assert.deepEqual(mergesTouching(regions, { start: { row: 11, col: 0 }, end: { row: 11, col: 0 } }), ["A12:B13"]);
});

test("a selection over several merges finds each of them and nothing else", () => {
  assert.deepEqual(mergesTouching(regions, { start: { row: 12, col: 4 }, end: { row: 0, col: 1 } }), ["A12:B13", "D1:E1"]);
  assert.deepEqual(mergesTouching(regions, { start: { row: 5, col: 5 }, end: { row: 6, col: 6 } }), []);
  assert.deepEqual(mergesTouching(undefined, { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } }), []);
});

test("merging over existing merges covers them all, as Excel does", () => {
  const chained = [
    { startRow: 8, startCol: 2, endRow: 8, endCol: 7, range: "C9:H9" },
    { startRow: 8, startCol: 8, endRow: 9, endCol: 9, range: "I9:J10" },
    { startRow: 9, startCol: 10, endRow: 9, endCol: 11, range: "K10:L10" }
  ];
  // D9:I9 reaches C9:H9 and I9:J10; the grown box then reaches K10:L10 too.
  assert.deepEqual(mergeTarget(chained, { start: { row: 8, col: 3 }, end: { row: 8, col: 10 } }), {
    range: { start: { row: 8, col: 2 }, end: { row: 9, col: 11 } },
    merges: ["C9:H9", "I9:J10", "K10:L10"]
  });
  assert.deepEqual(mergeTarget(chained, { start: { row: 19, col: 1 }, end: { row: 19, col: 3 } }), {
    range: { start: { row: 19, col: 1 }, end: { row: 19, col: 3 } },
    merges: []
  });
});
