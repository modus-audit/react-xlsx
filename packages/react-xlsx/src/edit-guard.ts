/** What an edit is about to change, for `onBeforeEdit`, and the formula checks it carries. Pure:
 *  the controller supplies the worksheet reads. */

export interface EditAddress {
  row: number;
  col: number;
}

export interface EditRange {
  start: EditAddress;
  end: EditAddress;
}

/** A formula the engine would mishandle silently: it leaves a malformed formula blank and stops
 *  recalculating on one that names a missing sheet. */
export type XlsxFormulaProblem = { kind: "syntax" } | { kind: "missingSheet"; sheetName: string };

export type XlsxEditKind =
  /** Cell values or formulas: typing, clearing, pasting, filling. */
  | "content"
  | "style"
  | "merge"
  | "unmerge"
  | "resize"
  /** Undo or redo. */
  | "history"
  /** Sheets, defined names, table sorts and form controls. */
  | "structure"
  /** Charts and images. */
  | "drawing";

/** An edit the viewer is about to make. `onBeforeEdit` returning false refuses it, before any change. */
export interface XlsxEdit {
  kind: XlsxEditKind;
  /** The workbook sheet the edit changes. */
  sheetIndex: number;
  /** The cells the edit writes (for a paste, from the active cell to the copied extent). */
  range?: EditRange;
  /** The net change in the workbook's formula count, counting replaced formulas once. Computed on
   *  call, so a host that refuses on `range` first never pays for it. Absent when the edit can't
   *  add formulas (a clear, a typed value). */
  formulaDelta?: () => number;
  /** A single formula entry, normalized to one leading `=`. */
  formula?: string;
  formulaProblem?: XlsxFormulaProblem;
  axis?: "row" | "column";
  /** The rows or columns a resize changes. */
  indices?: number[];
}

/** A formula as the user sees it, with one leading `=`. */
export function formulaText(formula: string): string {
  return `=${formula.replace(/^=/, "")}`;
}

function syntaxProblem(formula: string): boolean {
  let depth = 0;
  let quote: '"' | "'" | null = null;
  for (const char of formula) {
    if (quote) {
      if (char === quote) quote = null;
    } else if (char === '"' || char === "'") quote = char;
    else if (char === "(") depth += 1;
    else if (char === ")") depth -= 1;
    if (depth < 0) return true;
  }
  return quote !== null || depth !== 0 || /[+\-*/^&=<>,(]\s*$/.test(formula.slice(1));
}

const SHEET_REFERENCE = /'((?:[^']|'')+)'!|(?<![#\w.\u0080-\uFFFF])([A-Za-z_\u0080-￿][\w.\u0080-￿]*)!/g;

/** A sheet the formula names that the workbook lacks. References into other workbooks
 *  (`[1]Sheet!A1`, `'[Book.xlsx]Sheet'!A1`) are not checked. */
function missingSheet(formula: string, known: readonly string[]): string | null {
  const names = new Set(known);
  const text = formula.replace(/"[^"]*"/g, "");
  for (const match of text.matchAll(SHEET_REFERENCE)) {
    const name = (match[1] ?? match[2] ?? "").replace(/''/g, "'");
    if (name.startsWith("[") || text[(match.index ?? 0) - 1] === "]") continue;
    if (!names.has(name)) return name;
  }
  return null;
}

/** Unbalanced parentheses or quotes, a trailing operator, or a sheet the workbook lacks. */
export function formulaProblem(formula: string, sheetNames: readonly string[]): XlsxFormulaProblem | undefined {
  if (syntaxProblem(formula)) return { kind: "syntax" };
  const sheetName = missingSheet(formula, sheetNames);
  return sheetName === null ? undefined : { kind: "missingSheet", sheetName };
}

/** Same tab/newline reading as the viewer's plain-text paste. */
export function clipboardTextGrid(text: string): string[][] {
  const rows = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
  if (rows.length > 1 && rows[rows.length - 1] === "") rows.pop();
  return rows.map((row) => row.split("\t"));
}

export function rangeFrom(start: EditAddress, rows: number, cols: number): EditRange {
  return { start, end: { row: start.row + Math.max(rows, 1) - 1, col: start.col + Math.max(cols, 1) - 1 } };
}

/** Whether a cell currently holds a formula. */
export type HasFormula = (row: number, col: number) => boolean;

/** Net formula-count change from writing `writes`, counting each cell once (the last write wins). */
export function formulaDelta(writes: Iterable<EditAddress & { formula: boolean }>, hasFormula: HasFormula): number {
  const last = new Map<string, EditAddress & { formula: boolean }>();
  for (const write of writes) last.set(`${write.row}:${write.col}`, write);
  let delta = 0;
  for (const { row, col, formula } of last.values()) delta += Number(formula) - Number(hasFormula(row, col));
  return delta;
}

/** Whether a pasted text field becomes a formula: `=…`, unless pasting values only. */
export function isPastedFormula(value: string, literal = false): boolean {
  return !literal && value.startsWith("=") && value.length > 1;
}

export function* textWrites(start: EditAddress, grid: string[][], literal = false) {
  for (let row = 0; row < grid.length; row += 1) {
    const line = grid[row] ?? [];
    for (let col = 0; col < line.length; col += 1) {
      const value = line[col] ?? "";
      yield { row: start.row + row, col: start.col + col, formula: isPastedFormula(value, literal) };
    }
  }
}

export function* payloadWrites(
  start: EditAddress,
  cells: ReadonlyArray<{ rowOffset: number; colOffset: number; formula: string | null; styleOnly?: boolean }>
) {
  for (const cell of cells) {
    if (cell.styleOnly) continue;
    yield { row: start.row + cell.rowOffset, col: start.col + cell.colOffset, formula: Boolean(cell.formula) };
  }
}

/** The cells a fill writes from `source` across `target` (outside the source), and whether each
 *  takes a formula. */
export function* fillWrites(source: EditRange, target: EditRange, hasFormula: HasFormula) {
  const height = source.end.row - source.start.row + 1;
  const width = source.end.col - source.start.col + 1;
  const wrap = (index: number, start: number, size: number) => start + ((((index - start) % size) + size) % size);
  for (let row = target.start.row; row <= target.end.row; row += 1) {
    for (let col = target.start.col; col <= target.end.col; col += 1) {
      const inSource =
        row >= source.start.row && row <= source.end.row && col >= source.start.col && col <= source.end.col;
      if (inSource) continue;
      yield { row, col, formula: hasFormula(wrap(row, source.start.row, height), wrap(col, source.start.col, width)) };
    }
  }
}
