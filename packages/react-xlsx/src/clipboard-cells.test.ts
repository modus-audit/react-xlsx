import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { before, test } from "node:test";
import initSheetsWasm, { Workbook } from "@dukelib/sheets-wasm";
import { type ClipboardMatrixCell, clipboardStyleTable, copiedCell, styleToRestore, writePastedCell } from "./clipboard-cells.ts";

before(async () => {
  const wasmFile = readFileSync(new URL(import.meta.resolve("@dukelib/sheets-wasm/duke_sheets_wasm_bg.wasm")));
  await initSheetsWasm({ module_or_path: wasmFile });
});

const origin = { rowOffset: 0, colOffset: 0 };
const yellow = { fillType: "solid", color: { colorType: "rgb", hex: "FFFF00" } } as const;

function newWorkbook() {
  const workbook = new Workbook();
  if (workbook.sheetCount === 0) workbook.addSheet("Sheet1");
  return workbook;
}

/** Copies one cell and pastes it at `a1`, through JSON as the clipboard carries it. */
function copyPaste(worksheet: ReturnType<Workbook["getSheet"]>, from: [number, number], to: [number, number], a1: string) {
  const table = clipboardStyleTable();
  const cell = copiedCell(worksheet, from[0], from[1], origin, worksheet.getFormattedValueAt(from[0], from[1]), table);
  const payload = JSON.parse(JSON.stringify({ cells: [cell], styles: table.styles }));
  writePastedCell(worksheet, a1, to[0], to[1], payload.cells[0], payload.styles);
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

test("pasting a plain cell clears the formatting it lands on, as in Excel", () => {
  const worksheet = newWorkbook().getSheet(0);
  worksheet.setCell("A1", "clean");
  worksheet.setCell("B2", 5);
  worksheet.setCellStyleAt(1, 1, { font: { bold: true }, fill: yellow, numberFormat: { formatString: "0.00" } });

  copyPaste(worksheet, [0, 0], [1, 1], "B2");

  const style = worksheet.getCellStyleAt(1, 1);
  assert.equal(worksheet.getCellAt(1, 1).toJs(), "clean");
  assert.equal(style?.font?.bold ?? false, false);
  assert.equal(style?.fill?.fillType ?? "none", "none");
  assert.equal(style?.numberFormat?.formatType ?? "general", "general");
});

test("copied styles are stored once per distinct style", () => {
  const worksheet = newWorkbook().getSheet(0);
  const table = clipboardStyleTable();
  for (let row = 0; row < 3; row += 1) {
    worksheet.setCell(`A${row + 1}`, row);
    worksheet.setCellStyleAt(row, 0, { font: { bold: true } });
    copiedCell(worksheet, row, 0, { rowOffset: row, colOffset: 0 }, String(row), table);
  }
  assert.equal(table.styles.length, 1);
});

test("a copied formula carries its calculated value for pasting values only", () => {
  const workbook = newWorkbook();
  const worksheet = workbook.getSheet(0);
  worksheet.setCell("A1", 2);
  worksheet.setFormula("A2", "A1*3");
  workbook.calculate();

  const copied = copiedCell(worksheet, 1, 0, origin, "6", clipboardStyleTable());
  assert.match(copied.formula ?? "", /A1\*3$/);
  assert.equal(copied.raw, 6);
});

test("a payload copied before values and styles travelled still pastes its text", () => {
  const worksheet = newWorkbook().getSheet(0);
  const legacy: ClipboardMatrixCell = { ...origin, formula: null, value: "Cash" };
  writePastedCell(worksheet, "C3", 2, 2, legacy, undefined);
  assert.equal(worksheet.getCellAt(2, 2).toJs(), "Cash");
  assert.equal(worksheet.getCellStyleAt(2, 2), null);
});

test("restoring a cell that had no style of its own resets it to plain", () => {
  const worksheet = newWorkbook().getSheet(0);
  worksheet.setCell("A1", 1);
  worksheet.setCellStyleAt(0, 0, { font: { bold: true } });
  worksheet.setCellStyleAt(0, 0, styleToRestore(null));
  assert.equal(worksheet.getCellStyleAt(0, 0)?.font?.bold ?? false, false);
});
