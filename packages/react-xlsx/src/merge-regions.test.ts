import assert from "node:assert/strict";
import { test } from "node:test";
import { mergesTouching } from "./merge-regions.ts";

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
