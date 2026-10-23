import { checkedStyleRanges, normalizedStyleRange } from "./border-plan.ts";
import type { XlsxCellAddress, XlsxRangeStyleEdit } from "./types";

/** Preflight the entire batch before writing; capture each cell once, including overlapping ranges. */
export function applyStyleBatch<State>(edits: XlsxRangeStyleEdit[], operations: {
  allow: (range: XlsxRangeStyleEdit["range"]) => boolean;
  capture: (cell: XlsxCellAddress) => State;
  apply: (edit: XlsxRangeStyleEdit) => void;
  restore: (cell: XlsxCellAddress, state: State) => void;
}): Array<{cell: XlsxCellAddress; before: State; after: State}> | null {
  const ranges = edits.map(edit => normalizedStyleRange(edit.range));
  checkedStyleRanges(ranges);
  const before = new Map<string, {cell:XlsxCellAddress; state:State}>();
  for (const range of ranges) for (let row=range.start.row; row<=range.end.row; row++) {
    for (let col=range.start.col; col<=range.end.col; col++) {
      const key = `${row}:${col}`;
      if (!before.has(key)) before.set(key,{cell:{row,col},state:operations.capture({row,col})});
    }
  }
  if (!ranges.every(operations.allow)) return null;
  try {
    edits.forEach((edit,index) => operations.apply({...edit, range:ranges[index]}));
    return Array.from(before.values(), ({cell,state}) => ({cell,before:state,after:operations.capture(cell)}));
  } catch (error) {
    for (const {cell,state} of before.values()) operations.restore(cell,state);
    throw error;
  }
}


export class FormattingRefusal extends Error {}

/** Inherited formats need an effective-style editing model; refuse partial edits until it exists. */
export function requireEditableStyle(own: unknown, inherited: unknown): void {
  if (own == null && inherited != null) {
    throw new FormattingRefusal("These cells inherit row or column formatting. Edit their formatting in Excel.");
  }
}
