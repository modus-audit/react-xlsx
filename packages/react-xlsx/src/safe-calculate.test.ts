import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { before, test } from "node:test";
import initSheetsWasm, { Workbook } from "@dukelib/sheets-wasm";
import { externalCalcOptions, externalCallKey } from "./external-fn.ts";
import { safeCalculate } from "./safe-calculate.ts";

before(async () => {
  const wasmFile = readFileSync(new URL(import.meta.resolve("@dukelib/sheets-wasm/duke_sheets_wasm_bg.wasm")));
  await initSheetsWasm({ module_or_path: wasmFile });
});

test("recalculation passes external add-in values to the engine", () => {
  const key = externalCallKey("TBLink", ["Ledger"]);
  const workbook = {
    sheetCount: 0,
    sheetNames: [],
    calculate(options?: { externalFnFn?: (name: string, args: string[]) => string | number | null }) {
      assert.equal(options?.externalFnFn?.("TBLink", ["Ledger"]), 123.45);
      assert.equal(options?.externalFnFn?.("TBLink", ["Missing"]), null);
    },
  } as unknown as Workbook;

  const result = safeCalculate(workbook, { calcOptions: externalCalcOptions({ [key]: 123.45 }) });
  assert.equal(result.calculated, true);
  assert.equal(result.skipReason, null);
  assert.strictEqual(result.workbook, workbook);
});

function workbookWithFormulaCount(formulaCount: number) {
  const source = new Workbook();
  source.addSheet("Main");
  const sheet = source.getSheet(0);
  for (let index = 1; index <= formulaCount; index += 1) {
    sheet.setFormula(`B${index}`, "1+2");
  }
  const bytes = source.saveXlsxBytes();
  source.free();
  return Workbook.fromBytes(bytes);
}

test("recalculates past the 5,000 formulas where the threaded engine used to trap", () => {
  // The browser engine used to switch to a rayon thread pool at 5,000 formulas, which panics
  // without threads and poisoned the workbook. It now always evaluates serially.
  for (const formulaCount of [4_999, 5_000, 20_000]) {
    const workbook = workbookWithFormulaCount(formulaCount);
    const result = safeCalculate(workbook);

    assert.equal(result.calculated, true);
    assert.equal(result.skipReason, null);
    assert.equal(workbook.getSheet(0).getFormattedValue(`B${formulaCount}`), "3");
    workbook.free();
  }
});
