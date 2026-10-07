import { formulaErrorTooltip } from "./formula-error.ts";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { before, test } from "node:test";
import initSheetsWasm, { Workbook } from "@dukelib/sheets-wasm";
import { externalCalcOptions, externalCallKey } from "./external-fn.ts";
import { safeCalculate } from "./safe-calculate.ts";
import { calculationReport, cellCalculationDiagnostic, countSourceWorkbookFormulas } from "./calculation-diagnostics.ts";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

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
  assert.equal(result.calculation.status, "partial");
  assert.equal(result.calculation.errorCount, 0);
  assert.equal(result.calculation.engineErrorCount, null);
  assert.equal(result.calculation.reason, "engine-stats-unavailable");
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

test("reports typed formula errors even when the engine errors statistic is zero", () => {
  const workbook = new Workbook();
  workbook.addSheet("Main");
  workbook.getSheet(0).setFormula("B2", "1/0");
  const result = safeCalculate(workbook, { sourceFormulaCount: 1 });
  assert.equal(result.calculated, true);
  assert.equal(result.calculation.status, "partial");
  assert.equal(result.calculation.reason, "formula-errors");
  assert.equal(result.calculation.engineErrorCount, 0);
  assert.equal(result.calculation.errorCount, 1);
  assert.deepEqual(result.calculation.issues, [{ sheet: workbook.sheetNames[0], cell: "B2", error: "#DIV/0!" }]);
  assert.deepEqual(cellCalculationDiagnostic(workbook.getSheet(0), 1, 1, "971,835", result.calculation), {
    source: "calculated", error: "#DIV/0!"
  });
  workbook.free();
});

test("successful calculation reports real coverage without claiming Excel equivalence", () => {
  const workbook = new Workbook();
  workbook.addSheet("Main");
  const sheet = workbook.getSheet(0);
  sheet.setFormula("B2", "1+2");
  const sourceCount = countSourceWorkbookFormulas(workbook.saveXlsxBytes());
  assert.equal(sourceCount, 1);
  const result = safeCalculate(workbook, { sourceFormulaCount: sourceCount });
  assert.equal(result.calculation.status, "complete");
  assert.equal(result.calculation.formulaCount, 1);
  assert.equal(result.calculation.evaluatedFormulaCount, 1);
  assert.equal(result.calculation.errorCount, 0);
  assert.deepEqual(cellCalculationDiagnostic(sheet, 1, 1, "2", result.calculation), { source: "calculated", error: null });
  workbook.free();
});

test("an imported invalid formula cannot report complete when calculation excludes it", () => {
  const source = new Workbook();
  source.addSheet("Main");
  source.getSheet(0).setFormula("A1", "1+2");
  const archive = unzipSync(source.saveXlsxBytes());
  const xml = strFromU8(archive["xl/worksheets/sheet1.xml"]!);
  archive["xl/worksheets/sheet1.xml"] = strToU8(xml.replace("<f>1+2</f>", "<f>1+</f>"));
  const bytes = zipSync(archive);
  const sourceCount = countSourceWorkbookFormulas(bytes);
  assert.equal(sourceCount, 1);
  const loaded = Workbook.fromBytes(bytes);
  assert.equal(loaded.getSheet(0).formulaCount, 1);
  const result = safeCalculate(loaded, { sourceFormulaCount: sourceCount });
  assert.equal(result.calculation.status, "partial");
  assert.equal(result.calculation.reason, "formula-errors");
  assert.equal(result.calculation.formulaCount, 1);
  assert.equal(result.calculation.parsedFormulaCount, 1);
  assert.equal(result.calculation.evaluatedFormulaCount, 0);
  loaded.free();
  source.free();
});

test("missing sheets skip calculation while preserving the cached workbook", () => {
  const workbook = new Workbook();
  workbook.addSheet("Main");
  workbook.getSheet(0).setFormula("B2", "'Missing Sheet'!A1");
  const result = safeCalculate(workbook, { sourceFormulaCount: 1 });
  assert.equal(result.calculation.status, "skipped");
  assert.equal(result.calculation.reason, "unresolved-sheet-refs");
  assert.equal(result.calculation.errorCount, null);
  assert.equal(result.calculation.evaluatedFormulaCount, null);
  assert.doesNotThrow(() => workbook.getSheet(0));
  workbook.free();
});

test("source inventory mismatches and missing inventories prevent a complete report", () => {
  const workbook = new Workbook();
  workbook.getSheet(0).setFormula("A1", "1+2");
  const mismatch = safeCalculate(workbook, { sourceFormulaCount: 2 });
  assert.equal(mismatch.calculation.status, "partial");
  assert.equal(mismatch.calculation.reason, "formula-import-mismatch");
  assert.equal(mismatch.calculation.formulaCount, 2);
  assert.equal(mismatch.calculation.parsedFormulaCount, 1);
  const unavailable = safeCalculate(workbook);
  assert.equal(unavailable.calculation.status, "partial");
  assert.equal(unavailable.calculation.reason, "source-formula-inventory-unavailable");
  assert.equal(countSourceWorkbookFormulas(new Uint8Array([0xd0, 0xcf])), null);
  workbook.free();
});

test("partial coverage never labels an ignored cached formula as calculated", () => {
  const source = new Workbook();
  source.getSheet(0).setFormula("A1", "1+2");
  const archive = unzipSync(source.saveXlsxBytes());
  const xml = strFromU8(archive["xl/worksheets/sheet1.xml"]!);
  archive["xl/worksheets/sheet1.xml"] = strToU8(xml.replace("<f>1+2</f>", "<f>SUM(A2:)</f><v>77</v>"));
  const bytes = zipSync(archive);
  const loaded = Workbook.fromBytes(bytes);
  const result = safeCalculate(loaded, { sourceFormulaCount: countSourceWorkbookFormulas(bytes) });
  assert.equal(result.calculation.status, "partial");
  assert.equal(result.calculation.evaluatedFormulaCount, 0);
  assert.equal(loaded.getSheet(0).getCalculatedValueAt(0, 0).asNumber(), 77);
  assert.deepEqual(cellCalculationDiagnostic(loaded.getSheet(0), 0, 0, "77", result.calculation), {
    source: "unknown", error: null
  });
  loaded.free();
  source.free();
});

test("traps return a failed report and the fresh cache instance", () => {
  const cached = { sheetCount: 0, sheetNames: [] } as unknown as Workbook;
  const workbook = { ...cached, calculate() { throw new Error("test calculate trap"); } } as unknown as Workbook;
  const result = safeCalculate(workbook, { reparse: () => cached, sourceFormulaCount: 0 });
  assert.strictEqual(result.workbook, cached);
  assert.equal(result.calculation.status, "failed");
  assert.equal(result.calculation.reason, "calculate-trapped");
  assert.equal(result.calculation.errorCount, null);
});

test("cell inventory excludes sparkline, validation, and conditional-format extension formulas", () => {
  const source = new Workbook();
  source.getSheet(0).setFormula("A1", "1+2");
  const archive = unzipSync(source.saveXlsxBytes());
  const xml = strFromU8(archive["xl/worksheets/sheet1.xml"]!);
  const extension = `<extLst><ext uri="test" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:sparklineGroups><x14:sparklineGroup><x14:sparklines><x14:sparkline><xm:f>Sheet1!A1:A2</xm:f><xm:sqref>B1</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups><x14:dataValidations><x14:dataValidation><x14:formula1><xm:f>Sheet1!A1:A2</xm:f></x14:formula1></x14:dataValidation></x14:dataValidations><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule><xm:f>A1&gt;0</xm:f></x14:cfRule></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>`;
  archive["xl/worksheets/sheet1.xml"] = strToU8(xml.replace("</worksheet>", extension + "</worksheet>"));
  const bytes = zipSync(archive);
  const sourceCount = countSourceWorkbookFormulas(bytes);
  assert.equal(sourceCount, 1);
  const loaded = Workbook.fromBytes(bytes);
  const result = safeCalculate(loaded, { sourceFormulaCount: sourceCount });
  assert.equal(result.calculation.status, "complete");
  assert.equal(result.calculation.parsedFormulaCount, 1);
  assert.equal(result.calculation.evaluatedFormulaCount, 1);
  loaded.free();
  source.free();
});

test("source inventory handles namespaced shared formulas and ignores XML text", () => {
  const xml = `<s:worksheet xmlns:s = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"><s:sheetData><s:row r="1"><s:c r="A1"><s:f t="shared" si="0">1+2</s:f></s:c><s:c r="B1"><s:f t="shared" si="0"/></s:c><s:c r="C1"><s:is><s:t><![CDATA[<s:f>fake</s:f>]]></s:t></s:is></s:c><!-- <s:c><s:f>fake</s:f></s:c> --></s:row></s:sheetData></s:worksheet>`;
  const source = new Workbook();
  const archive = unzipSync(source.saveXlsxBytes());
  archive["xl/worksheets/sheet1.xml"] = strToU8(xml);
  assert.equal(countSourceWorkbookFormulas(zipSync(archive)), 2);
  source.free();
});

test("a later skipped attempt cannot label an earlier engine value as file-saved", () => {
  const workbook = new Workbook();
  const sheet = workbook.getSheet(0);
  sheet.setFormula("A1", "1+2");
  const first = safeCalculate(workbook, { sourceFormulaCount: 1 });
  assert.equal(first.calculation.status, "complete");
  sheet.setFormula("B1", "'Missing Sheet'!A1");
  const skipped = safeCalculate(workbook);
  assert.equal(skipped.calculation.status, "skipped");
  const value = sheet.getCalculatedValueAt(0, 0);
  assert.equal(value.asNumber(), 3);
  value.free();
  assert.deepEqual(cellCalculationDiagnostic(sheet, 0, 0, "77", skipped.calculation, true), { source: "unknown", error: null });
  assert.deepEqual(cellCalculationDiagnostic(sheet, 0, 0, "77", calculationReport("skipped", "auto-formula-limit"), false), { source: "saved", error: null });
  workbook.free();
});

test("intentional edits invalidate coverage without inventing an import loss", () => {
  const workbook = workbookWithFormulaCount(1001);
  const first = safeCalculate(workbook, { sourceFormulaCount: 1001 });
  assert.equal(first.calculation.status, "complete");
  workbook.getSheet(0).setFormula("C1", "4+5");
  const edited = calculationReport("partial", "workbook-edited", 1002);
  assert.deepEqual(cellCalculationDiagnostic(workbook.getSheet(0), 0, 2, undefined, edited), { source: "unknown", error: null });
  const second = safeCalculate(workbook, { sourceFormulaCount: null });
  assert.equal(second.calculation.reason, "source-formula-inventory-unavailable");
  assert.equal(second.calculation.parsedFormulaCount, 1002);
  assert.equal(second.calculation.evaluatedFormulaCount, 1002);
  workbook.free();
});

test("real engine repeated calculation uses changed external add-in inputs", () => {
  const workbook = new Workbook();
  const sheet = workbook.getSheet(0);
  sheet.setFormula("A1", '[1]!TBLink("Ledger")');
  sheet.setFormula("B1", "A1+2");
  const key = externalCallKey("TBLink", ["Ledger"]);
  for (const externalValue of [100, 250]) {
    const result = safeCalculate(workbook, { sourceFormulaCount: 2, calcOptions: externalCalcOptions({ [key]: externalValue }) });
    assert.equal(result.calculation.status, "complete");
    assert.equal(result.calculation.evaluatedFormulaCount, 2);
    const value = sheet.getCalculatedValueAt(0, 1);
    assert.equal(value.asNumber(), externalValue + 2);
    value.free();
  }
  workbook.free();
});

test("keeps every calculator error and supports unfamiliar error codes", () => {
  const workbook = new Workbook();
  workbook.addSheet("Errors");
  for (let row = 1; row <= 25; row += 1) workbook.getSheet(0).setFormula(`B${row}`, "1/0");
  const report = safeCalculate(workbook, { sourceFormulaCount: 25 }).calculation;
  assert.equal(report.errorCount, 25);
  assert.equal(report.issues.length, 25);
  assert.deepEqual(report.issues.map(issue => issue.cell).sort(), Array.from({ length: 25 }, (_, index) => `B${index + 1}`).sort());
  assert.equal(formulaErrorTooltip(report.issues[0].error), "Divides by zero or an empty cell.");
  assert.equal(formulaErrorTooltip("#FUTURE!"), "This formula could not be calculated.");
  workbook.free();
});

test("source inventory ignores orphan worksheets and fails closed on missing linked parts", () => {
  const source = new Workbook();
  source.getSheet(0).setFormula("A1", "1+2");
  const archive = unzipSync(source.saveXlsxBytes());
  archive["xl/worksheets/orphan.xml"] = archive["xl/worksheets/sheet1.xml"]!;
  assert.equal(countSourceWorkbookFormulas(zipSync(archive)), 1);
  delete archive["xl/worksheets/sheet1.xml"];
  assert.equal(countSourceWorkbookFormulas(zipSync(archive)), null);
  source.free();
});
