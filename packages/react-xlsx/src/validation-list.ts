import type { XlsxCellValue } from "./types";
import type { Workbook } from "@dukelib/sheets-wasm";

/** Resolve ordinary Excel validation lists without evaluating arbitrary formulas. */
export function validationListItems(source: string | undefined, workbook: Workbook | null, sheetIndex: number): XlsxCellValue[] | null {
  let reference = source?.trim().replace(/^=/, "") ?? "";
  if (!reference) return null;
  if (reference.startsWith('"') && reference.endsWith('"')) {
    return reference.slice(1, -1).replace(/""/g, '"').split(",");
  }
  if (!workbook) return null;
  try {
    // Names can point at other names. Bound cycles and leave dynamic formulas to Excel.
    for (let depth = 0; depth < 8; depth += 1) {
      const named = workbook.getNamedRange(reference);
      if (!named) break;
      reference = named.trim().replace(/^=/, "");
    }
    const match = /^(?:(?:'((?:[^']|'')+)'|([^'!]+))!)?\$?([A-Z]+)\$?([1-9]\d*)(?::\$?([A-Z]+)\$?([1-9]\d*))?$/i.exec(reference);
    if (!match) return null;
    const [, quotedSheet, plainSheet, startCol, startRow, endCol = startCol, endRow = startRow] = match;
    const sheetName = quotedSheet?.replace(/''/g, "'") ?? plainSheet;
    const index = sheetName === undefined ? sheetIndex : workbook.sheetIndex(sheetName);
    if (index === undefined || index < 0) return null;
    const column = (label: string) => [...label.toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
    const left = column(startCol), right = column(endCol);
    const top = Number(startRow) - 1, bottom = Number(endRow) - 1;
    if (left > right || top > bottom || right >= 16384 || bottom >= 1048576 ||
        (right - left + 1) * (bottom - top + 1) > 10000) return null;
    const sheet = workbook.getSheet(index);
    try {
      const items: XlsxCellValue[] = [];
      for (let row = top; row <= bottom; row += 1) {
        for (let col = left; col <= right; col += 1) {
          const value = sheet.getCalculatedValueAt(row, col);
          try {
            const raw: unknown = value.toJs();
            if (typeof raw === "string" || typeof raw === "number" || typeof raw === "boolean") items.push(raw);
          } finally { value.free(); }
        }
      }
      return [...new Set(items)];
    } finally { sheet.free(); }
  } catch { return null; }
}
