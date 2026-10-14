import type { XlsxCellRange, XlsxSheetData } from "./types";

export interface XlsxReviewFilterInput {
  /** First row is a retained header; keep totals outside this range. */
  range: XlsxCellRange;
  column: number;
  mode: "contains" | "equals" | "blank" | "nonblank";
  query: string;
}
export interface XlsxReviewFilter extends XlsxReviewFilterInput {
  excludedRows: number[];
  matchingRows: number;
  totalRows: number;
}

export function reviewFilterProblem(input: XlsxReviewFilterInput, sheet: XlsxSheetData): string | null {
  const { start, end } = input.range;
  if (![start.row, start.col, end.row, end.col, input.column].every(Number.isSafeInteger)
    || start.row < 0 || start.col < 0 || end.row <= start.row || end.col < start.col
    || end.row > sheet.maxUsedRow || end.col > sheet.maxUsedCol
    || input.column < start.col || input.column > end.col) {
    return "Choose a data range with a header and at least one data row, and a column inside it.";
  }
  if (end.row - start.row > 50_000) return "Review filters support up to 50,000 data rows at a time.";
  if (sheet.hasVerticalMerges) return "Review filters are unavailable on sheets with vertically merged records.";
  return null;
}

/** Missing sparse records are blank; malformed batch data is an unavailable read, not a blank. */
export function reviewColumnValues(rows: unknown[], column: number): Map<number, string> {
  const values = new Map<number, string>();
  for (const row of rows) {
    if (!row || typeof row !== "object" || !("index" in row) || !Number.isSafeInteger(row.index)
      || !("cells" in row) || !Array.isArray(row.cells)) throw new Error("The worksheet values could not be read.");
    for (const cell of row.cells) {
      if (!cell || typeof cell !== "object" || !("col" in cell) || !("value" in cell)) {
        throw new Error("The worksheet values could not be read.");
      }
      if (cell.col === column) {
        if (typeof cell.value !== "string") throw new Error("The worksheet values could not be read.");
        values.set(Number(row.index), cell.value);
      }
    }
  }
  return values;
}

export async function scanReviewFilter(
  input: XlsxReviewFilterInput,
  hiddenRows: ReadonlySet<number>,
  read: (start: number, count: number) => Promise<Map<number, string>>,
  isCurrent: () => boolean,
): Promise<XlsxReviewFilter | null> {
  const excludedRows: number[] = [];
  let matchingRows = 0;
  let totalRows = 0;
  const query = input.query.toLocaleLowerCase();
  for (let start = input.range.start.row + 1; start <= input.range.end.row; start += 250) {
    if (!isCurrent()) return null;
    const count = Math.min(250, input.range.end.row - start + 1);
    const values = await read(start, count);
    if (!isCurrent()) return null;
    for (let row = start; row < start + count; row++) {
      const value = (values.get(row) ?? "").toLocaleLowerCase();
      const matches = input.mode === "blank" ? value === ""
        : input.mode === "nonblank" ? value !== ""
        : input.mode === "equals" ? value === query : value.includes(query);
      if (!matches) excludedRows.push(row);
      if (!hiddenRows.has(row)) {
        totalRows++;
        if (matches) matchingRows++;
      }
    }
    await new Promise<void>(resolve => setTimeout(resolve, 0));
  }
  return { ...input, excludedRows, matchingRows, totalRows };
}

/** Project display geometry without changing saved hidden rows, sizes or worksheet coordinates. */
export function projectReviewRows(sheet: XlsxSheetData, excludedRows: number[]): XlsxSheetData {
  const excluded = new Set(excludedRows);
  const visibleRows: number[] = [];
  const rowHeights: number[] = [];
  sheet.visibleRows.forEach((row, index) => {
    if (excluded.has(row)) return;
    visibleRows.push(row);
    rowHeights.push(sheet.rowHeights[index] ?? sheet.defaultRowHeightPx);
  });
  return { ...sheet, excludedRows, visibleRows, rowHeights };
}
