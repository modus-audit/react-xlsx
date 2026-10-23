import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { before, test } from "node:test";
import initSheetsWasm, { Workbook } from "@dukelib/sheets-wasm";
import { normalizeCellValue } from "./clipboard-cells.ts";
import { borderPlan } from "./border-plan.ts";
import { applyStyleBatch } from "./style-batch.ts";
import { numericSequence, fillAxis } from "./numeric-fill.ts";
import type { XlsxCellRange } from "./types";

const range = (top:number,left:number,bottom=top,right=left):XlsxCellRange => ({start:{row:top,col:left},end:{row:bottom,col:right}});
before(async () => {
  await initSheetsWasm({module_or_path:readFileSync(new URL(import.meta.resolve("@dukelib/sheets-wasm/duke_sheets_wasm_bg.wasm")))});
});

test("border batch preserves merge interiors and original theme/diagonal styles; refusal and failure are atomic", () => {
  const wb = new Workbook(); if (!wb.sheetCount) wb.addSheet("Sheet1"); const ws = wb.getSheet(0);
  const original = {border:{left:{style:"thick",color:{colorType:"theme",themeIndex:4,tint:0.2}},diagonal:{style:"thin"},diagonalDirection:"up"}};
  ws.setRangeStyle("B2",original);
  ws.mergeCells("B2:C2");
  const plan = borderPlan([range(1,1,1,2)],ws.mergedRegions,"all");
  let writes = 0;
  const a1 = (cell:{row:number;col:number}) => `${String.fromCharCode(65+cell.col)}${cell.row+1}`;
  const operations = {
    allow: (_r:XlsxCellRange) => true,
    capture: (cell:{row:number;col:number}) => structuredClone(ws.getCellStyleAt(cell.row,cell.col)),
    apply: (edit:typeof plan[number]) => {
      writes++;
      ws.setRangeStyle(`${a1(edit.range.start)}:${a1(edit.range.end)}`,edit.style);
    },
    restore: (cell:{row:number;col:number},state:unknown) => {
      if (state) ws.setCellStyleAt(cell.row,cell.col,state);
      else ws.clearCellStyleAt(cell.row,cell.col);
    }
  };
  const before=operations.capture({row:1,col:1});
  assert.equal(applyStyleBatch(plan,{...operations,allow:r => r.start.row !== 1}),null);
  assert.equal(writes,0);
  assert.deepEqual(operations.capture({row:1,col:1}),before);
  assert.throws(() => applyStyleBatch(plan,{...operations,apply:edit => {
    operations.apply(edit); if (writes===2) throw new Error("mutation failure");
  }}),/mutation failure/);
  assert.deepEqual(operations.capture({row:1,col:1}),before);
  const changes = applyStyleBatch(plan,operations)!;
  assert.equal(new Set(changes.map(c => `${c.cell.row}:${c.cell.col}`)).size,changes.length);
  assert.equal(ws.getCellStyleAt(1,1).border.right,undefined);
  assert.equal(ws.getCellStyleAt(1,2).border.left,undefined);
  assert.equal(ws.getCellStyleAt(1,1).border.diagonalDirection,"up");
  assert.equal(ws.getCellStyleAt(1,1).border.diagonal.style,"thin");
  assert.equal(ws.getCellStyleAt(0,1).border.bottom.style,"thin");
  // One returned collection restores every affected cell, including outside neighbors.
  for (const change of changes) operations.restore(change.cell,change.before);
  assert.deepEqual(operations.capture({row:1,col:1}),before);
  assert.throws(() => borderPlan([range(1,1)],ws.mergedRegions,"outside"),/entire merged/);
  assert.throws(() => borderPlan([range(0,0,1000,1000)],[],"all"),/250,000/);
  ws.free(); wb.free();
});

test("numeric fill uses signed coordinates and refuses ambiguous seed types", () => {
  for (const [values,first,step] of [[[1,2],1,1],[[5,3,1],5,-2],[[0.1,0.2,0.3],0.1,0.1],[[7,7],7,0]] as const) {
    const series=numericSequence(values.map(value => ({value})))!;
    assert.equal(series.first,first); assert.ok(Math.abs(series.step-step)<1e-12);
    assert.ok(Math.abs((series.first+(-2)*series.step)-(first-2*step))<1e-12);
    assert.ok(Math.abs((series.first+5*series.step)-(first+5*step))<1e-12);
  }
  for (const seeds of [[{value:1}],[{value:"001"},{value:"002"}],[{value:1},{value:null}],
    [{value:1},{value:2,formula:"=1+1"}],[{value:1,date:true},{value:2,date:true}],
    [{value:1},{value:2},{value:4}],[{value:Infinity},{value:2}],[{value:1e308},{value:-1e308}]]) {
    assert.equal(numericSequence(seeds),null);
  }
  assert.equal(fillAxis(range(2,2,3,2),range(0,2,3,2)),"row");
  assert.equal(fillAxis(range(2,2,2,3),range(2,0,2,3)),"col");
  assert.equal(fillAxis(range(2,2),range(0,0,3,3)),null);
});


test("clearing a style restores exact null through the WASM boundary and retains cell contents", () => {
  const wb = new Workbook(); if (!wb.sheetCount) wb.addSheet("Sheet1"); const ws = wb.getSheet(0);
  ws.setCell("A1", "001");
  ws.setFormula("A2", "=1+1");
  ws.setRangeStyle("A1:A2", { font: { name: "Arial", family: 2, bold: true }, numberFormat: { formatString: "0%" } });
  ws.clearCellStyleAt(0, 0);
  ws.clearCellStyleAt(1, 0);
  assert.equal(ws.getCellStyleAt(0, 0), null);
  assert.equal(ws.getCellStyleAt(1, 0), null);
  const value = ws.getCellAt(0, 0);
  assert.equal(value.toJs(), "001");
  value.free();
  assert.equal(ws.getFormulaAt(1, 0), "=1+1");
  assert.throws(() => ws.clearCellStyleAt(0, 65536), /outside/);
  ws.free(); wb.free();
});


test("format and fill undo preserve absent blanks and literal empty strings as distinct WASM values", () => {
  const wb = new Workbook(); if (!wb.sheetCount) wb.addSheet("Sheet1"); const ws = wb.getSheet(0);
  ws.setCell("B1", "");
  ws.setFormula("D1", "=COUNTA(A1)");
  const a1 = (cell:{row:number;col:number}) => `${String.fromCharCode(65+cell.col)}${cell.row+1}`;
  const capture = (cell:{row:number;col:number}) => {
    const value = ws.getCellAt(cell.row, cell.col);
    try { return {value:value.toJs(),style:ws.getCellStyleAt(cell.row,cell.col)}; }
    finally { value.free(); }
  };
  const beforeBlank = capture({row:0,col:0});
  const beforeLiteral = capture({row:0,col:1});
  assert.equal(beforeBlank.value, null);
  assert.equal(beforeLiteral.value, "");
  const undo = (cell:{row:number;col:number}, state:typeof beforeBlank) => {
    ws.setCell(a1(cell), normalizeCellValue(state.value));
    if (state.style) ws.setCellStyleAt(cell.row,cell.col,state.style);
    else ws.clearCellStyleAt(cell.row,cell.col);
  };
  ws.setRangeStyle("A1:B1",{font:{bold:true}});
  undo({row:0,col:0},beforeBlank); undo({row:0,col:1},beforeLiteral);
  wb.calculate();
  let total = ws.getCalculatedValueAt(0,3);
  assert.equal(total.toJs(),0); total.free();
  assert.equal(capture({row:0,col:0}).value,null);
  assert.equal(capture({row:0,col:1}).value,"");
  ws.setCell("A1",3); ws.setCell("B1",4);
  undo({row:0,col:0},beforeBlank); undo({row:0,col:1},beforeLiteral);
  wb.calculate();
  total = ws.getCalculatedValueAt(0,3);
  assert.equal(total.toJs(),0); total.free();
  assert.equal(capture({row:0,col:0}).value,null);
  assert.equal(capture({row:0,col:1}).value,"");
  ws.free(); wb.free();
});
