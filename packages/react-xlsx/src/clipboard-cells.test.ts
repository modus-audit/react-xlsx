import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { before, test } from "node:test";
import initSheetsWasm, { Workbook } from "@dukelib/sheets-wasm";
import {
  type ClipboardMatrixCell,
  clipboardStyleTable,
  copiedCell,
  plainCellStyle,
  styleToRestore,
  writePastedCell
} from "./clipboard-cells.ts";

before(async () => {
  const wasmFile = readFileSync(new URL(import.meta.resolve("@dukelib/sheets-wasm/duke_sheets_wasm_bg.wasm")));
  await initSheetsWasm({ module_or_path: wasmFile });
});

const origin = { rowOffset: 0, colOffset: 0 };
const yellow = { fillType: "solid", color: { colorType: "rgb", hex: "FFFF00" } } as const;
/** A new workbook's default font, as its first cell format names it. */
const plain = plainCellStyle({ name: "Calibri", size: 11 });

function newWorkbook() {
  const workbook = new Workbook();
  if (workbook.sheetCount === 0) workbook.addSheet("Sheet1");
  return workbook;
}

/** Copies one cell and pastes it at `a1`, through JSON as the clipboard carries it. */
function copyPaste(worksheet: ReturnType<Workbook["getSheet"]>, from: [number, number], to: [number, number], a1: string) {
  const table = clipboardStyleTable();
  const style = worksheet.getCellStyleAt(from[0], from[1]);
  const cell = copiedCell(worksheet, from[0], from[1], origin, worksheet.getFormattedValueAt(from[0], from[1]), style, table);
  const payload = JSON.parse(JSON.stringify({ cells: [cell], styles: table.styles }));
  writePastedCell(worksheet, a1, to[0], to[1], payload.cells[0], payload.styles, plain);
}

test("a pasted cell keeps its number and its formatting", () => {
  const worksheet = newWorkbook().getSheet(0);
  worksheet.setCell("A1", 29081);
  worksheet.setCellStyleAt(0, 0, { font: { bold: true }, fill: yellow, numberFormat: { formatString: "#,##0.00" } });

  copyPaste(worksheet, [0, 0], [1, 1], "B2");

  assert.equal(worksheet.getCellAt(1, 1).toJs(), 29081);
  assert.equal(worksheet.getFormattedValueAt(1, 1), "29,081.00");
  assert.equal(worksheet.getCellStyleAt(1, 1).font.bold, true);
  assert.deepEqual(worksheet.getCellStyleAt(1, 1).fill, worksheet.getCellStyleAt(0, 0).fill);
});

test("pasting a plain cell resets the cell it lands on to the default format, as in Excel", () => {
  const worksheet = newWorkbook().getSheet(0);
  worksheet.setCell("A1", "clean");
  worksheet.setCell("B2", 5);
  worksheet.setCellStyleAt(1, 1, {
    font: { name: "Arial", size: 8, bold: true },
    fill: yellow,
    numberFormat: { formatString: "0.00" },
    protection: { locked: false }
  });

  copyPaste(worksheet, [0, 0], [1, 1], "B2");

  assert.equal(worksheet.getCellAt(1, 1).toJs(), "clean");
  assert.equal(worksheet.getCellStyleAt(1, 1), null);
});

test("restoring a cell that had no style of its own leaves it with none, font and lock included", () => {
  const worksheet = newWorkbook().getSheet(0);
  worksheet.setCell("A1", 1);
  worksheet.setCellStyleAt(0, 0, { font: { name: "Arial", size: 8, bold: true }, fill: yellow, protection: { locked: false } });
  worksheet.setCellStyleAt(0, 0, styleToRestore(null, plain));
  assert.equal(worksheet.getCellStyleAt(0, 0), null);
});

test("a merged block's covered cells carry their style only", () => {
  const worksheet = newWorkbook().getSheet(0);
  worksheet.setCell("C3", "keep");
  const covered: ClipboardMatrixCell = { ...origin, formula: null, value: "", styleOnly: true, styleIndex: 0 };
  writePastedCell(worksheet, "C3", 2, 2, covered, [{ border: { right: { style: "thin" } } }], plain);
  assert.equal(worksheet.getCellAt(2, 2).toJs(), "keep");
  assert.equal(worksheet.getCellStyleAt(2, 2).border.right.style, "thin");
});

test("copied styles are stored once per distinct style", () => {
  const worksheet = newWorkbook().getSheet(0);
  const table = clipboardStyleTable();
  for (let row = 0; row < 3; row += 1) {
    worksheet.setCell(`A${row + 1}`, row);
    worksheet.setCellStyleAt(row, 0, { font: { bold: true } });
    copiedCell(worksheet, row, 0, { rowOffset: row, colOffset: 0 }, String(row), worksheet.getCellStyleAt(row, 0), table);
  }
  assert.equal(table.styles.length, 1);
});

test("a copied formula carries its calculated value for pasting values only", () => {
  const workbook = newWorkbook();
  const worksheet = workbook.getSheet(0);
  worksheet.setCell("A1", 2);
  worksheet.setFormula("A2", "A1*3");
  workbook.calculate();

  const copied = copiedCell(worksheet, 1, 0, origin, "6", null, clipboardStyleTable());
  assert.match(copied.formula ?? "", /A1\*3$/);
  assert.equal(copied.raw, 6);
});

test("copying a formula relocates relative references and preserves absolute references", () => {
  const workbook = newWorkbook();
  const sheet = workbook.getSheet(0);
  sheet.setCell("A1", 2);
  sheet.setCell("B2", 5);
  sheet.setFormula("A2", "=A1+$A$1");
  workbook.calculate();
  copyPaste(sheet, [1, 0], [2, 1], "B3");
  workbook.calculate();
  assert.equal(sheet.getFormulaAt(2, 1), "=B2+$A$1");
  assert.equal(sheet.getCalculatedValueAt(2, 1).toJs(), 7);
  assert.equal(sheet.getFormulaAt(1, 0), "=A1+$A$1");
});

test("a payload copied before values and styles travelled still pastes its text", () => {
  const worksheet = newWorkbook().getSheet(0);
  const legacy: ClipboardMatrixCell = { ...origin, formula: null, value: "Cash" };
  writePastedCell(worksheet, "C3", 2, 2, legacy, undefined, plain);
  assert.equal(worksheet.getCellAt(2, 2).toJs(), "Cash");
  assert.equal(worksheet.getCellStyleAt(2, 2), null);
});
