import type { Workbook } from "@dukelib/sheets-wasm";
import type { ExternalCalcOptions } from "./external-fn.ts";
import { calculationReport, inspectCalculation, type XlsxCalculationReport } from "./calculation-diagnostics.ts";

/** Automatic load-time calculation stays conservative to keep large workbooks responsive. */
export const AUTO_CALCULATE_FORMULA_THRESHOLD = 1_000;

const SHEET_REF_REGEX = /'((?:[^']|'')+)'!|(?<![#\w.\u0080-\uFFFF])([A-Za-z_\u0080-\uFFFF][\w.\u0080-\uFFFF]*)!/g;

type FormulaCell = { formula?: string | null };

function collectReferencedSheetNames(workbook: Workbook): Set<string> {
  const referenced = new Set<string>();
  for (let sheetIdx = 0; sheetIdx < workbook.sheetCount; sheetIdx += 1) {
    let sheet;
    try {
      sheet = workbook.getSheet(sheetIdx);
    } catch {
      continue;
    }
    const cells = sheet.formulaCells as FormulaCell[] | null | undefined;
    if (!Array.isArray(cells)) {
      continue;
    }
    for (const cell of cells) {
      const formula = cell?.formula;
      if (!formula) {
        continue;
      }
      SHEET_REF_REGEX.lastIndex = 0;
      let match: RegExpExecArray | null;
      while ((match = SHEET_REF_REGEX.exec(formula)) !== null) {
        const raw = match[1] ?? match[2];
        if (!raw) {
          continue;
        }
        referenced.add(raw.replace(/''/g, "'"));
      }
    }
  }
  return referenced;
}

function hasUnresolvedSheetReferences(workbook: Workbook): boolean {
  let names: string[];
  try {
    names = workbook.sheetNames;
  } catch {
    return false;
  }
  const known = new Set(names);
  const referenced = collectReferencedSheetNames(workbook);
  for (const name of referenced) {
    if (!known.has(name)) {
      return true;
    }
  }
  return false;
}

export type SafeCalculateSkipReason = "unresolved-sheet-refs" | "calculate-trapped";

export type SafeCalculateResult = {
  workbook: Workbook;
  calculated: boolean;
  skipReason: SafeCalculateSkipReason | null;
  calculation: XlsxCalculationReport;
};

export type SafeCalculateOptions = {
  reparse?: () => Workbook;
  calcOptions?: ExternalCalcOptions;
  sourceFormulaCount?: number | null;
};

export function countWorkbookFormulas(workbook: Workbook): number {
  let total = 0;
  for (let index = 0; index < workbook.sheetCount; index += 1) {
    total += workbook.getSheet(index).formulaCount;
  }
  return total;
}

// Pre-scans for formulas referencing missing sheets (which cause the Rust
// engine to panic into a wasm `unreachable` trap that poisons the Workbook
// instance). On trap, `reparse` is used to return a fresh usable instance.
export function safeCalculate(workbook: Workbook, options: SafeCalculateOptions = {}): SafeCalculateResult {
  const formulaCount = countWorkbookFormulas(workbook);
  const sourceFormulaCount = options.sourceFormulaCount ?? null;
  const skipped = (reason: SafeCalculateSkipReason, nextWorkbook = workbook): SafeCalculateResult => ({
    workbook: nextWorkbook, calculated: false, skipReason: reason,
    calculation: calculationReport(reason === "calculate-trapped" ? "failed" : "skipped", reason, formulaCount, sourceFormulaCount)
  });
  if (hasUnresolvedSheetReferences(workbook)) {
    return skipped("unresolved-sheet-refs");
  }
  try {
    const start = performance.now();
    const stats: unknown = workbook.calculate(options.calcOptions);
    return { workbook, calculated: true, skipReason: null,
      calculation: inspectCalculation(workbook, stats, formulaCount, sourceFormulaCount, performance.now() - start) };
  } catch (err) {
    console.warn("[react-xlsx] workbook.calculate() trapped; falling back to cached formula values", err);
    if (options.reparse) {
      try {
        return skipped("calculate-trapped", options.reparse());
      } catch (reparseErr) {
        console.warn("[react-xlsx] workbook reparse after calculate trap failed", reparseErr);
      }
    }
    return skipped("calculate-trapped");
  }
}
