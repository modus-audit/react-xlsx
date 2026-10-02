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
  if (!Array.isArray(mergedRegions)) return [];
  const top = Math.min(range.start.row, range.end.row);
  const bottom = Math.max(range.start.row, range.end.row);
  const left = Math.min(range.start.col, range.end.col);
  const right = Math.max(range.start.col, range.end.col);
  return mergedRegions
    .filter(isMergedRegion)
    .filter((region) => region.startRow <= bottom && region.endRow >= top && region.startCol <= right && region.endCol >= left)
    .map((region) => region.range);
}
