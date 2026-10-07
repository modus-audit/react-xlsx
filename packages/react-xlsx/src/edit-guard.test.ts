import assert from "node:assert/strict";
import { test } from "node:test";
import { clipboardTextGrid, fillWrites, formulaDelta, formulaProblem, formulaText, payloadWrites, rangeFrom, textWrites } from "./edit-guard.ts";

test("a formula entry gets one leading =", () => {
  assert.equal(formulaText("A1*2"), "=A1*2");
  assert.equal(formulaText("=A1"), "=A1");
});

test("malformed formulas are a syntax problem", () => {
  for (const formula of ["=RoundD24:D26)", "=ROUND(A1", '="open', "=A1+", "=SUM(A1,"]) {
    assert.deepEqual(formulaProblem(formula, []), { kind: "syntax" }, formula);
  }
  for (const formula of [
    "=ROUND(D24,0)",
    '=IF(A1>1,"(",B1)',
    '=ROUND([1]!TBLink("Interim 3/31 Combined - TB Database","FINAL[7]","4600","1","1"),0)'
  ]) {
    assert.equal(formulaProblem(formula, []), undefined, formula);
  }
});

test("a formula naming a sheet the workbook lacks is a problem; other workbooks are not checked", () => {
  const sheets = ["Lead", "Rev (2)"];
  assert.deepEqual(formulaProblem("=Leads!A1+1", sheets), { kind: "missingSheet", sheetName: "Leads" });
  assert.deepEqual(formulaProblem("='Gone Sheet'!A1", sheets), { kind: "missingSheet", sheetName: "Gone Sheet" });
  assert.equal(formulaProblem("=Lead!A1+'Rev (2)'!B2", sheets), undefined);
  assert.equal(formulaProblem("=[1]Summary!A1+'[Book.xlsx]Other'!A1", sheets), undefined);
  assert.equal(formulaProblem('="Nope!"&Lead!A1', sheets), undefined);
});

test("formula delta counts each cell once and replacements as no growth", () => {
  const existing = new Set(["0:0"]);
  const hasFormula = (row: number, col: number) => existing.has(`${row}:${col}`);
  const start = { row: 0, col: 0 };
  // Replacing A1's formula and adding B1's grows by one; a typed value in A2 adds nothing.
  assert.equal(formulaDelta(textWrites(start, clipboardTextGrid("=2\t=3\n7\n")), hasFormula), 1);
  // Clearing A1 by pasting a value makes room.
  assert.equal(formulaDelta(textWrites(start, clipboardTextGrid("7")), hasFormula), -1);
  const cells = [
    { rowOffset: 0, colOffset: 1, formula: "=A1" },
    { rowOffset: 0, colOffset: 2, formula: null, styleOnly: true }
  ];
  assert.equal(formulaDelta(payloadWrites(start, cells), hasFormula), 1);
});

test("a fill repeats the source's formulas across the target, outside the source", () => {
  const hasFormula = (row: number, col: number) => row === 0 && col === 0;
  const source = { start: { row: 0, col: 0 }, end: { row: 0, col: 1 } };
  const target = { start: { row: 0, col: 0 }, end: { row: 2, col: 1 } };
  const writes = [...fillWrites(source, target, hasFormula)];
  assert.equal(writes.length, 4);
  assert.equal(writes.filter((write) => write.formula).length, 2);
});

test("a paste's range runs from the start over the copied extent", () => {
  assert.deepEqual(rangeFrom({ row: 2, col: 1 }, 2, 3), { start: { row: 2, col: 1 }, end: { row: 3, col: 3 } });
  assert.deepEqual(clipboardTextGrid("a\tb\r\nc\n"), [["a", "b"], ["c"]]);
});
