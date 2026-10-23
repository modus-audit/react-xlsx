import type { XlsxCellBorderStyleInput, XlsxCellRange, XlsxRangeStyleEdit, XlsxBorderPreset } from "./types";

const MAX_CELLS = 250_000;
const MAX_ROW = 1_048_575;
const MAX_COL = 16_383;
type Region = { startRow: number; startCol: number; endRow: number; endCol: number };
type Edge = "top" | "bottom" | "left" | "right";

export function normalizedStyleRange(range: XlsxCellRange): XlsxCellRange {
  const { start, end } = range;
  if (![start.row, start.col, end.row, end.col].every(Number.isInteger) ||
      Math.min(start.row, end.row) < 0 || Math.min(start.col, end.col) < 0 ||
      Math.max(start.row, end.row) > MAX_ROW || Math.max(start.col, end.col) > MAX_COL) {
    throw new Error("Select cells within the worksheet.");
  }
  return { start: { row: Math.min(start.row, end.row), col: Math.min(start.col, end.col) },
    end: { row: Math.max(start.row, end.row), col: Math.max(start.col, end.col) } };
}

/** Bounds synchronous formatting work even for hosts without their own edit policy. */
export function checkedStyleRanges(ranges: XlsxCellRange[]): XlsxCellRange[] {
  const normalized = ranges.map(normalizedStyleRange);
  const count = normalized.reduce((total, r) => total + (r.end.row-r.start.row+1)*(r.end.col-r.start.col+1), 0);
  if (count > MAX_CELLS) throw new Error("Select 250,000 cells or fewer to format at once.");
  return normalized;
}

function regions(raw: unknown): Region[] {
  if (!Array.isArray(raw)) return [];
  return raw.filter((r): r is Region => r !== null && typeof r === "object" &&
    ["startRow", "startCol", "endRow", "endCol"].every(k => Number.isInteger(r[k])));
}

/** Edges are written on both incident cells so a stronger adjacent border cannot hide the action.
 * Covered merge cells keep perimeter styles, but never acquire new interior borders. */
export function borderPlan(ranges: XlsxCellRange[], rawMerges: unknown, preset: XlsxBorderPreset): XlsxRangeStyleEdit[] {
  const selected = checkedStyleRanges(ranges);
  const merges = regions(rawMerges);
  const patches = new Map<string, { row: number; col: number; border: XlsxCellBorderStyleInput }>();
  const line = preset === "clear" ? { style: "none" as const } : {
    style: preset === "doubleBottom" ? "double" as const : "thin" as const,
    color: { colorType: "rgb" as const, hex: "000000" }
  };
  const put = (row: number, col: number, edge: Edge) => {
    if (row < 0 || col < 0 || row > MAX_ROW || col > MAX_COL) return;
    const key = `${row}:${col}`;
    const patch = patches.get(key) ?? { row, col, border: {} };
    patch.border[edge] = line;
    patches.set(key, patch);
  };
  const boundary = (row: number, col: number, edge: Edge) => {
    put(row, col, edge);
    if (edge === "top") put(row-1, col, "bottom");
    if (edge === "bottom") put(row+1, col, "top");
    if (edge === "left") put(row, col-1, "right");
    if (edge === "right") put(row, col+1, "left");
  };
  for (const r of selected) {
    const touched = merges.filter(m => m.startRow <= r.end.row && m.endRow >= r.start.row &&
      m.startCol <= r.end.col && m.endCol >= r.start.col);
    if (touched.some(m => m.startRow < r.start.row || m.endRow > r.end.row ||
      m.startCol < r.start.col || m.endCol > r.end.col)) {
      throw new Error("Select the entire merged cell before changing its borders.");
    }
    const merged = new Map<string, Region>();
    if (preset === "all") for (const m of touched) {
      for (let row=m.startRow; row<=m.endRow; row++) for (let col=m.startCol; col<=m.endCol; col++) {
        merged.set(`${row}:${col}`, m);
      }
    }
    if (preset === "bottom" || preset === "doubleBottom") {
      for (let col=r.start.col; col<=r.end.col; col++) boundary(r.end.row,col,"bottom");
      continue;
    }
    if (preset === "outside") {
      for (let col=r.start.col; col<=r.end.col; col++) { boundary(r.start.row,col,"top"); boundary(r.end.row,col,"bottom"); }
      for (let row=r.start.row; row<=r.end.row; row++) { boundary(row,r.start.col,"left"); boundary(row,r.end.col,"right"); }
      continue;
    }
    for (let row=r.start.row; row<=r.end.row; row++) for (let col=r.start.col; col<=r.end.col; col++) {
      const m = merged.get(`${row}:${col}`);
      if (!m || row === m.startRow) boundary(row,col,"top");
      if (!m || row === m.endRow) boundary(row,col,"bottom");
      if (!m || col === m.startCol) boundary(row,col,"left");
      if (!m || col === m.endCol) boundary(row,col,"right");
    }
  }
  // Coalesce equal horizontal runs so both host preflight and WASM calls stay bounded by rows,
  // rather than one bridge call per cell in a large All Borders selection.
  const result: XlsxRangeStyleEdit[] = [];
  let previousKey = "";
  for (const p of Array.from(patches.values()).sort((a,b) => a.row-b.row || a.col-b.col)) {
    const border: XlsxCellBorderStyleInput = {top:p.border.top,bottom:p.border.bottom,left:p.border.left,right:p.border.right};
    const key = JSON.stringify(border);
    const previous = result[result.length-1];
    if (previous && previousKey === key && previous.range.end.row === p.row && previous.range.end.col+1 === p.col) {
      previous.range.end.col = p.col;
    } else {
      result.push({range:{start:{row:p.row,col:p.col},end:{row:p.row,col:p.col}},style:{border}});
    }
    previousKey = key;
  }
  return result;
}
