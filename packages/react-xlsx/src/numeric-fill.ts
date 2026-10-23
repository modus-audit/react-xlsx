import type { XlsxCellRange } from "./types";
export type FillMode = "auto" | "copy" | "series";
export interface FillSeed { value: unknown; formula?: string | null; date?: boolean }

/** A lane must already be an arithmetic sequence of numeric constants. Never infer from display text. */
export function numericSequence(seeds: FillSeed[]): {first:number; step:number} | null {
  if (seeds.length < 2 || seeds.some(s => typeof s.value !== "number" || !Number.isFinite(s.value) || s.formula || s.date)) return null;
  const numbers = seeds.map(s => s.value as number);
  const first = numbers[0];
  const step = numbers[1]-first;
  if (!Number.isFinite(step) || numbers.some((value,index) => {
    const expected = first+index*step;
    return !Number.isFinite(expected) || Math.abs(value-expected) > 1e-12*Math.max(1,Math.abs(value),Math.abs(expected));
  })) return null;
  return {first,step};
}

export function fillAxis(source: XlsxCellRange, target: XlsxCellRange): "row" | "col" | null {
  const rows = target.start.row !== source.start.row || target.end.row !== source.end.row;
  const cols = target.start.col !== source.start.col || target.end.col !== source.end.col;
  return rows === cols ? null : rows ? "row" : "col";
}
