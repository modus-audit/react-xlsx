import assert from "node:assert/strict";
import { afterEach, test } from "node:test";
import { cutPayload, cutStillApplies, forgetCopy, rememberCopy, rememberedCells, valuesOnly, withoutCopyOrigins } from "./clipboard-memory.ts";

const copied = JSON.stringify({
  rows: 1,
  cols: 2,
  merges: [{ rowOffset: 0, colOffset: 0, rowSpan: 1, colSpan: 2 }],
  styles: [{ font: { bold: true } }],
  cells: [
    { rowOffset: 0, colOffset: 0, formula: "=A1*2", value: "29,081.00", raw: 29081, styleIndex: 0, source: { row: 4, col: 0 } },
    { rowOffset: 0, colOffset: 1, formula: null, value: "", styleOnly: true, styleIndex: 0 }
  ]
});

afterEach(forgetCopy);

test("paste values keeps each cell's value and drops formulas, formatting, merges and merged cells", () => {
  assert.deepEqual(JSON.parse(valuesOnly(copied) ?? "null"), {
    rows: 1,
    cols: 2,
    merges: [],
    styles: [],
    cells: [{ rowOffset: 0, colOffset: 0, value: "29,081.00", raw: 29081, formula: null }]
  });
  assert.equal(valuesOnly("not json"), null);
  assert.equal(valuesOnly(JSON.stringify({ cells: "nope" })), null);
});

test("a cut keeps formulas unrelocated by dropping copy origins", () => {
  const cut = JSON.parse(withoutCopyOrigins(copied) ?? "null");
  assert.equal(cut.cells[0].source, undefined);
  assert.equal(cut.cells[0].formula, "=A1*2");
});

test("the remembered copy is used only while the clipboard text is that copy", () => {
  rememberCopy("29,081.00", copied);
  assert.equal(rememberedCells("29,081.00\r\n"), copied);
  assert.equal(rememberedCells("something else"), null);
  forgetCopy();
  assert.equal(rememberedCells("29,081.00"), null);
});

test("a cut whose clear is refused copies instead, so pasting it still relocates formulas", () => {
  assert.equal(cutPayload(copied, false), copied);
  assert.equal(JSON.parse(cutPayload(copied, true)).cells[0].source, undefined);
});

test("an async cut leaves the cells when the sheet or workbook changed during the clipboard write", () => {
  const source = { sheet: 0, generation: 3 };
  assert.equal(cutStillApplies(source, { sheet: 0, generation: 3 }), true);
  assert.equal(cutStillApplies(source, { sheet: 1, generation: 3 }), false);
  assert.equal(cutStillApplies(source, { sheet: 0, generation: 4 }), false);
});
