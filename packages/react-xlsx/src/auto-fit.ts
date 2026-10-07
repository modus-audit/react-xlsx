/** Column widths and row heights that fit the displayed text, like Excel's AutoFit. */

/** The grid's 4px cell padding on each side. */
const CELL_PADDING_PX = 8;
/** Room Excel leaves after the widest text when it fits a column. */
const FIT_GAP_PX = 12;
const MIN_WIDTH_PX = 24;
const MAX_WIDTH_PX = 640;
const MIN_HEIGHT_PX = 15;
const MAX_HEIGHT_PX = 409;
const LINE_HEIGHT = 1.2;
const ROW_PADDING_PX = 2;

export interface FitSheet {
  usedRange(): unknown;
  getFormattedValueAt(row: number, col: number): string;
  getCellStyleAt(row: number, col: number): unknown;
  getMergeSpan(row: number, col: number): unknown;
  isMergedSecondary(row: number, col: number): boolean;
}

export type MeasureText = (text: string, font: string) => number;

function record(value: unknown, key: string): Record<string, unknown> | null {
  if (typeof value !== "object" || value === null || !(key in value)) return null;
  const inner: unknown = (value as Record<string, unknown>)[key];
  return typeof inner === "object" && inner !== null ? (inner as Record<string, unknown>) : null;
}

function fontPx(style: unknown): number {
  const size = record(style, "font")?.size;
  return ((typeof size === "number" ? size : 11) * 96) / 72;
}

function cssFont(style: unknown): string {
  const font = record(style, "font");
  const name = typeof font?.name === "string" ? font.name : "Calibri";
  return `${font?.italic === true ? "italic " : ""}${font?.bold === true ? "bold " : ""}${fontPx(style)}px "${name}", sans-serif`;
}

function usedEnd(sheet: FitSheet, index: 2 | 3): number {
  const range = sheet.usedRange();
  return Array.isArray(range) && typeof range[index] === "number" ? range[index] : -1;
}

/** A cell that covers more than one row or column; AutoFit skips those, as Excel does. */
export function isMerged(sheet: Pick<FitSheet, "getMergeSpan" | "isMergedSecondary">, row: number, col: number): boolean {
  if (sheet.isMergedSecondary(row, col)) return true;
  const span = sheet.getMergeSpan(row, col);
  return (
    typeof span === "object" &&
    span !== null &&
    (("colSpan" in span && Number(span.colSpan) > 1) || ("rowSpan" in span && Number(span.rowSpan) > 1))
  );
}

function wraps(style: unknown): boolean {
  return record(style, "alignment")?.wrapText === true;
}

/** Lines `text` takes when wrapped at word breaks within `widthPx`. */
function wrappedLines(text: string, font: string, widthPx: number, measure: MeasureText): number {
  let lines = 0;
  for (const paragraph of text.split("\n")) {
    let line = "";
    lines += 1;
    for (const word of paragraph.split(" ")) {
      const next = line ? `${line} ${word}` : word;
      if (line && measure(next, font) > widthPx) {
        lines += 1;
        line = word;
      } else {
        line = next;
      }
    }
  }
  return lines;
}

/** Pixel width for `col`, or null when the column holds no fittable text. Skips merged and wrapped cells. */
export function fittedColumnWidth(sheet: FitSheet, col: number, measure: MeasureText): number | null {
  let widest = 0;
  const lastRow = usedEnd(sheet, 2);
  for (let row = 0; row <= lastRow; row += 1) {
    // Untrimmed: accounting formats pad with spaces (`_(` and `_)`) that the grid draws too.
    const text = sheet.getFormattedValueAt(row, col);
    if (!text.trim() || isMerged(sheet, row, col)) continue;
    const style = sheet.getCellStyleAt(row, col);
    if (!style || wraps(style)) continue;
    widest = Math.max(widest, measure(text, cssFont(style)));
  }
  return widest > 0
    ? Math.min(MAX_WIDTH_PX, Math.max(MIN_WIDTH_PX, Math.ceil(widest + CELL_PADDING_PX + FIT_GAP_PX)))
    : null;
}

/** Columns a one-row merge anchored at (`row`, `col`) spans, 1 for a plain cell, or 0 for a cell
 *  that can't size this row: inside a merge, or anchoring one taller than a row. */
function rowFitSpan(sheet: FitSheet, row: number, col: number): number {
  if (sheet.isMergedSecondary(row, col)) return 0;
  const span = sheet.getMergeSpan(row, col);
  if (typeof span !== "object" || span === null) return 1;
  if ("rowSpan" in span && Number(span.rowSpan) > 1) return 0;
  return "colSpan" in span ? Math.max(1, Number(span.colSpan) || 1) : 1;
}

/** Pixel height for `row`: the tallest cell, counting wrapped lines at each column's width (a
 *  merge across columns of this one row wraps at the merge's width, which Excel's AutoFit skips).
 *  Null when the row holds no text. Skips merges taller than the row. */
export function fittedRowHeight(
  sheet: FitSheet,
  row: number,
  columnWidthPx: (col: number) => number,
  measure: MeasureText
): number | null {
  let tallest = 0;
  const lastCol = usedEnd(sheet, 3);
  for (let col = 0; col <= lastCol; col += 1) {
    const text = sheet.getFormattedValueAt(row, col).trim();
    if (!text) continue;
    const span = rowFitSpan(sheet, row, col);
    if (span === 0) continue;
    const style = sheet.getCellStyleAt(row, col);
    let widthPx = 0;
    for (let spanned = col; spanned < col + span; spanned += 1) widthPx += columnWidthPx(spanned);
    const lines = wraps(style)
      ? wrappedLines(text, cssFont(style), widthPx - CELL_PADDING_PX, measure)
      : text.split("\n").length;
    tallest = Math.max(tallest, lines * fontPx(style) * LINE_HEIGHT);
  }
  return tallest > 0 ? Math.min(MAX_HEIGHT_PX, Math.max(MIN_HEIGHT_PX, Math.ceil(tallest + ROW_PADDING_PX))) : null;
}

let context: CanvasRenderingContext2D | null | undefined;
let contextFont = "";

export function measureText(text: string, font: string): number {
  context ??= typeof document === "undefined" ? null : document.createElement("canvas").getContext("2d");
  if (!context) return text.length * 7;
  // Setting the canvas font parses it each time; cells in a column mostly share one.
  if (font !== contextFont) {
    context.font = font;
    contextFont = font;
  }
  return context.measureText(text).width;
}

/** The sizes AutoFit would set for `indices` along `axis`, in grid pixels (gridline included),
 *  skipping hidden ones and those already within a pixel. An empty row returns to the default
 *  height; an empty column keeps its width. */
export function autoFitSizes(input: {
  axis: "column" | "row";
  indices: Iterable<number>;
  sheet: FitSheet;
  hidden: ReadonlySet<number>;
  /** Current content sizes (no gridline) in pixels. */
  columnWidthPx: (col: number) => number;
  rowHeightPx: (row: number) => number;
  defaultRowHeightPx: number;
  gridlinePx: number;
  measure?: MeasureText;
}): Array<{ index: number; sizePx: number }> {
  const { axis, sheet, hidden, gridlinePx, measure = measureText } = input;
  const sizes: Array<{ index: number; sizePx: number }> = [];
  const seen = new Set<number>();
  for (const index of input.indices) {
    if (seen.has(index) || hidden.has(index)) continue;
    seen.add(index);
    const fitted =
      axis === "column"
        ? fittedColumnWidth(sheet, index, measure)
        : (fittedRowHeight(sheet, index, input.columnWidthPx, measure) ?? input.defaultRowHeightPx + gridlinePx);
    const current = (axis === "column" ? input.columnWidthPx(index) : input.rowHeightPx(index)) + gridlinePx;
    if (fitted !== null && Math.abs(fitted - current) >= 1) sizes.push({ index, sizePx: fitted });
  }
  return sizes;
}
