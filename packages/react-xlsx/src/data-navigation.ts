import type { Worksheet } from "@dukelib/sheets-wasm";
import type { XlsxCellAddress, XlsxSheetData } from "./types";

export type DataDirection = "ArrowUp" | "ArrowDown" | "ArrowLeft" | "ArrowRight";
export interface DataNavigationRequest {
  cell: XlsxCellAddress;
  direction: DataDirection;
  maxRow: number;
  maxCol: number;
}

export function findDataBoundary(
  request: DataNavigationRequest,
  sheet: Pick<XlsxSheetData, "hiddenRows" | "hiddenCols" | "maxUsedRow" | "maxUsedCol">,
  hasContent: (row: number, col: number) => boolean
): XlsxCellAddress {
  const vertical = request.direction === "ArrowUp" || request.direction === "ArrowDown";
  const step = request.direction === "ArrowUp" || request.direction === "ArrowLeft" ? -1 : 1;
  const hidden = new Set((vertical ? sheet.hiddenRows : sheet.hiddenCols) ?? []);
  const max = vertical ? request.maxRow : request.maxCol;
  const usedMax = vertical ? sheet.maxUsedRow : sheet.maxUsedCol;
  const start = vertical ? request.cell.row : request.cell.col;
  const nextVisible = (position: number) => {
    let next = position + step;
    while (next >= 0 && next <= max && hidden.has(next)) next += step;
    return next >= 0 && next <= max ? next : null;
  };
  const occupied = (position: number) => position <= usedMax && hasContent(
    vertical ? position : request.cell.row,
    vertical ? request.cell.col : position
  );
  let destination = start;
  let next = nextVisible(start);
  if (next !== null) {
    const contiguous = occupied(start) && occupied(next);
    while (next !== null) {
      const filled = occupied(next);
      if (contiguous && !filled) break;
      destination = next;
      if (!contiguous && filled) break;
      next = nextVisible(next);
    }
  }
  return vertical ? { row: destination, col: request.cell.col } : { row: request.cell.row, col: destination };
}

export function worksheetHasContent(worksheet: Worksheet, row: number, col: number): boolean {
  if (worksheet.getFormulaAt(row, col)) return true;
  const value = worksheet.getCalculatedValueAt(row, col);
  try {
    return !value.is_empty;
  } finally {
    value.free();
  }
}
