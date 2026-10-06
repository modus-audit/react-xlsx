import assert from "node:assert/strict";
import { test } from "node:test";
import { copyFormula, fillSourceIndex } from "./formula-copy.ts";

test("copy relocates cells, ranges and mixed references on both axes", () => {
  assert.equal(copyFormula("=SUM(A1:B3)+$C4+D$5+$E$6", 2, 3), "=SUM(D3:E5)+$C6+G$5+$E$6");
  assert.equal(copyFormula("=a1+$b2+c$3", 1, 1), "=b2+$b3+d$3");
  assert.equal(copyFormula("=SUM(A : B,$C:$D,1:3,$4:$5)", 2, 1), "=SUM(B : C,$C:$D,3:5,$4:$5)");
});

test("copy leaves strings, escaped sheet names, functions and structured names intact", () => {
  const formula = `=IF(A1="A1""B2",LOG10(B2),SUM('A1 and O''Brien'!C3,Sheet1!D4,A1:B2!E5,Table1[A1],A1_column,XFE1,[Book1.xlsx]Sheet1!F6))`;
  assert.equal(copyFormula(formula, 1, 1), `=IF(B2="A1""B2",LOG10(C3),SUM('A1 and O''Brien'!D4,Sheet1!E5,A1:B2!F6,Table1[A1],A1_column,XFE1,[Book1.xlsx]Sheet1!G7))`);
  assert.equal(copyFormula("=SUM(Table1[[#Headers],[A1]],A1)", 1, 1), "=SUM(Table1[[#Headers],[A1]],B2)");
  assert.equal(copyFormula('=CCH(A1,"[A1]",$B$2)', 2, 2), '=CCH(C3,"[A1]",$B$2)');
});

test("copy produces reference errors at grid edges and preserves locked axes", () => {
  assert.equal(copyFormula("=A1+$A1+A$1+$A$1", -1, -1), "=#REF!+#REF!+#REF!+$A$1");
  assert.equal(copyFormula("=XFD1048576+$XFD$1048576", 1, 1), "=#REF!+$XFD$1048576");
  assert.equal(copyFormula("=A1# + @B2 + 1.2E10 + tax.A1", 1, 1), "=B2# + @C3 + 1.2E10 + tax.A1");
});

test("upward and leftward fills repeat from the source anchor", () => {
  assert.deepEqual([2, 3, 4, 5, 6, 7].map((row) => fillSourceIndex(row, 4, 3)), [5, 6, 4, 5, 6, 4]);
});
