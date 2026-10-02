import type { StyleInput, Workbook } from "@dukelib/sheets-wasm";

type Worksheet = ReturnType<Workbook["getSheet"]>;
type CellPrimitive = string | number | boolean;

/** One copied cell in the internal clipboard payload. `raw` and `styleIndex` are optional so
 *  payloads copied before they existed still paste. */
export type ClipboardMatrixCell = {
  colOffset: number;
  formula: string | null;
  rowOffset: number;
  /** The text the grid shows; pasted only when `raw` is absent. */
  value: string;
  /** The value as stored, or as calculated for a formula, so a pasted number stays a number. */
  raw?: CellPrimitive;
  /** The cell's style in the payload's `styles`: Excel's paste carries formatting by default. */
  styleIndex?: number;
};

/** A cell with no style of its own reads back as `null`, and the engine has no call that clears a
 *  style, so writing one back (a paste, an undo) resets it to these plain settings instead. Font
 *  name and size stay as they are. */
export const PLAIN_CELL_STYLE: StyleInput = {
  font: {
    bold: false,
    italic: false,
    underline: "none",
    strikethrough: false,
    color: { colorType: "auto" },
    verticalAlign: "baseline"
  },
  fill: { fillType: "none" },
  border: {
    left: { style: "none" },
    right: { style: "none" },
    top: { style: "none" },
    bottom: { style: "none" },
    diagonal: { style: "none" },
    diagonalDirection: "none"
  },
  alignment: { horizontal: "general", vertical: "bottom", wrapText: false, shrinkToFit: false, indent: 0, rotation: 0 },
  numberFormat: { formatType: "general" }
};

/** The style to write back for a cell whose style read as `style`. */
export function styleToRestore(style: unknown): StyleInput {
  return style && typeof style === "object" ? (style as StyleInput) : PLAIN_CELL_STYLE;
}

/** Collects copied styles once each, since a range usually repeats a handful of them. */
export function clipboardStyleTable() {
  const styles: unknown[] = [];
  const indexByKey = new Map<string, number>();
  return {
    styles,
    add(style: unknown): number {
      const key = JSON.stringify(style ?? null);
      let index = indexByKey.get(key);
      if (index === undefined) {
        index = styles.push(style ?? null) - 1;
        indexByKey.set(key, index);
      }
      return index;
    }
  };
}

function primitive(value: unknown): CellPrimitive | undefined {
  return typeof value === "string" || typeof value === "number" || typeof value === "boolean" ? value : undefined;
}

/** The clipboard entry for the cell at `row`, `col`; `offset` places it within the copied range. */
export function copiedCell(
  worksheet: Worksheet,
  row: number,
  col: number,
  offset: { rowOffset: number; colOffset: number },
  displayValue: string,
  styles: ReturnType<typeof clipboardStyleTable>
): ClipboardMatrixCell {
  const formula = worksheet.getFormulaAt(row, col) ?? null;
  const value = formula ? worksheet.getCalculatedValueAt(row, col) : worksheet.getCellAt(row, col);
  const raw = primitive(value.toJs());
  return {
    ...offset,
    formula,
    value: displayValue,
    ...(raw === undefined ? {} : { raw }),
    styleIndex: styles.add(worksheet.getCellStyleAt(row, col))
  };
}

/** Writes a copied cell at `row`, `col`: its formula or value, then its style when the payload
 *  carries one. */
export function writePastedCell(
  worksheet: Worksheet,
  a1: string,
  row: number,
  col: number,
  cell: ClipboardMatrixCell,
  styles: unknown[] | undefined
) {
  if (cell.formula) {
    worksheet.setFormula(a1, cell.formula);
  } else {
    worksheet.setCell(a1, cell.raw ?? cell.value);
  }
  if (cell.styleIndex !== undefined && styles && cell.styleIndex < styles.length) {
    worksheet.setCellStyleAt(row, col, styleToRestore(styles[cell.styleIndex]));
  }
}
