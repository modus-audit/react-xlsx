import assert from "node:assert/strict";
import { test } from "node:test";
import { projectReviewRows, scanReviewFilter, reviewColumnValues } from "./review-filter.ts";
import { findDataBoundary } from "./data-navigation.ts";
import type { XlsxSheetData } from "./types";

test("local review filtering shares raw-coordinate projection, hidden rows and navigation without altering saved metadata", async () => {
  const input = { range: {start:{row:0,col:0},end:{row:4,col:1}}, column:0, mode:"equals" as const, query:"Match" };
  const saved = { name:"Lead", hiddenRows:[2], visibleRows:[0,1,3,4,5], rowHeights:[20,21,23,24,25], defaultRowHeightPx:20 } as XlsxSheetData;
  const before = structuredClone(saved);
  const filter = await scanReviewFilter(input,new Set(saved.hiddenRows),async () => reviewColumnValues([
    {index:1,cells:[{col:0,value:"MATCH"}]}, {index:2,cells:[{col:0,value:"Match"}]},
    {index:3,cells:[{col:0,value:""}]}, {index:4,cells:[{col:0,value:"Other"}]},
  ],0),()=>true);
  assert.ok(filter);
  assert.deepEqual(filter.excludedRows,[3,4]);
  assert.equal(filter.matchingRows,1); assert.equal(filter.totalRows,3);
  const projected=projectReviewRows(saved,filter.excludedRows);
  assert.deepEqual(projected.visibleRows,[0,1,5]);
  assert.deepEqual(projected.rowHeights,[20,21,25]);
  assert.deepEqual(projected.hiddenRows,[2]);
  assert.deepEqual(saved,before);
  assert.deepEqual(findDataBoundary({cell:{row:1,col:0},direction:"ArrowDown",maxRow:5,maxCol:1,excludedRows:filter.excludedRows},
    {...saved,maxUsedRow:5,maxUsedCol:1},()=>true),{row:5,col:0});
  const blank = await scanReviewFilter({...input,mode:"blank"},new Set(),async()=>new Map(),()=>true);
  assert.equal(blank?.matchingRows,4);
  let current=true;
  const canceled=await scanReviewFilter(input,new Set(),async()=>{current=false;return new Map();},()=>current);
  assert.equal(canceled,null);
  assert.throws(()=>reviewColumnValues([{index:1,cells:[{col:0,value:12}]}],0),/could not be read/);
});
