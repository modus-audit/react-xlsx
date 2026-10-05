import type { XlsxCellRange } from "./types";

type MergedRegion = { startRow: number; startCol: number; endRow: number; endCol: number; range: string };

function isMergedRegion(value: unknown): value is MergedRegion {
  if (typeof value !== "object" || value === null) return false;
  const region = value as Record<string, unknown>;
  return (
    typeof region.range === "string" &&
    ["startRow", "startCol", "endRow", "endCol"].every((key) => typeof region[key] === "number")
  );
}

/** The A1 ranges of the merged blocks `range` touches. The engine unmerges only an exact merged
 *  range, while Excel's Unmerge clears every merge in the selection, including one selected by its
 *  top-left cell. */
export function mergesTouching(mergedRegions: unknown, range: XlsxCellRange): string[] {
  return regionsTouching(mergedRegions, range).map((region) => region.range);
}

function regionsTouching(mergedRegions: unknown, range: XlsxCellRange): MergedRegion[] {
  if (!Array.isArray(mergedRegions)) return [];
  const top = Math.min(range.start.row, range.end.row);
  const bottom = Math.max(range.start.row, range.end.row);
  const left = Math.min(range.start.col, range.end.col);
  const right = Math.max(range.start.col, range.end.col);
  return mergedRegions
    .filter(isMergedRegion)
    .filter((region) => region.startRow <= bottom && region.endRow >= top && region.startCol <= right && region.endCol >= left);
}

/** What Excel merges when the selection touches existing merges: the smallest box holding the
 *  selection and every merge it reaches, plus those merges (to clear first). The engine refuses to
 *  merge over a merged cell. */
export function mergeTarget(mergedRegions: unknown, range: XlsxCellRange): { range: XlsxCellRange; merges: string[] } {
  let top = Math.min(range.start.row, range.end.row);
  let bottom = Math.max(range.start.row, range.end.row);
  let left = Math.min(range.start.col, range.end.col);
  let right = Math.max(range.start.col, range.end.col);
  let merges: MergedRegion[] = [];
  // Growing the box can reach more merges, so repeat until it stops growing.
  for (;;) {
    merges = regionsTouching(mergedRegions, { start: { row: top, col: left }, end: { row: bottom, col: right } });
    const nextTop = Math.min(top, ...merges.map((region) => region.startRow));
    const nextBottom = Math.max(bottom, ...merges.map((region) => region.endRow));
    const nextLeft = Math.min(left, ...merges.map((region) => region.startCol));
    const nextRight = Math.max(right, ...merges.map((region) => region.endCol));
    if (nextTop === top && nextBottom === bottom && nextLeft === left && nextRight === right) break;
    [top, bottom, left, right] = [nextTop, nextBottom, nextLeft, nextRight];
  }
  return {
    range: { start: { row: top, col: left }, end: { row: bottom, col: right } },
    merges: merges.map((region) => region.range)
  };
}
