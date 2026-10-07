import type { Workbook } from "@dukelib/sheets-wasm";
export { countSourceWorkbookFormulas } from "./source-formula-inventory.ts";

export interface XlsxCalculationIssue {
  sheet: string;
  cell: string;
  error: string;
}

/** Engine execution coverage, not verification against Excel or financial correctness. */
export interface XlsxCalculationReport {
  status: "idle" | "calculating" | "complete" | "partial" | "skipped" | "failed";
  reason: string | null;
  formulaCount: number;
  parsedFormulaCount: number;
  sourceFormulaCount: number | null;
  evaluatedFormulaCount: number | null;
  errorCount: number | null;
  engineErrorCount: number | null;
  durationMs: number | null;
  revision: number;
  issues: XlsxCalculationIssue[];
}

export interface XlsxCellCalculationDiagnostic {
  source: "literal" | "calculated" | "saved" | "unknown";
  error: string | null;
}

export function calculationReport(
  status: XlsxCalculationReport["status"],
  reason: string | null = null,
  parsedFormulaCount = 0,
  sourceFormulaCount: number | null = null,
): XlsxCalculationReport {
  return {
    status, reason, parsedFormulaCount, sourceFormulaCount,
    formulaCount: Math.max(parsedFormulaCount, sourceFormulaCount ?? 0),
    evaluatedFormulaCount: null, errorCount: null, engineErrorCount: null,
    durationMs: null, revision: 0, issues: []
  };
}

function nonnegativeInteger(value: unknown): number | null {
  return typeof value === "number" && Number.isSafeInteger(value) && value >= 0 ? value : null;
}

function cellAddress(row: number, col: number): string {
  let column = col + 1;
  let label = "";
  while (column > 0) {
    label = String.fromCharCode(65 + (column - 1) % 26) + label;
    column = Math.floor((column - 1) / 26);
  }
  return `${label}${row + 1}`;
}

/** Scan typed results; the pinned engine's stats.errors is zero for some error-valued cells. */
export function inspectCalculation(
  workbook: Workbook,
  rawStats: unknown,
  parsedFormulaCount: number,
  sourceFormulaCount: number | null,
  durationMs: number,
): XlsxCalculationReport {
  const stats = rawStats && typeof rawStats === "object" ? rawStats as Record<string, unknown> : {};
  const evaluatedFormulaCount = nonnegativeInteger(stats.cellsCalculated);
  const engineErrorCount = nonnegativeInteger(stats.errors);
  const issues: XlsxCalculationIssue[] = [];
  let resultErrors = 0;
  let inspected = 0;
  let inspectionComplete = true;
  try {
    for (let index = 0; index < workbook.sheetCount; index += 1) {
      const sheet = workbook.getSheet(index);
      const cells: unknown = sheet.formulaCells;
      if (!Array.isArray(cells)) { inspectionComplete = false; continue; }
      for (const cell of cells) {
        if (!cell || !Number.isInteger(cell.row) || !Number.isInteger(cell.col)) {
          inspectionComplete = false;
          continue;
        }
        const value = sheet.getCalculatedValueAt(cell.row, cell.col);
        inspected += 1;
        if (value.is_error) {
          resultErrors += 1;
          issues.push({
            sheet: workbook.sheetNames[index] ?? String(index),
            cell: cellAddress(cell.row, cell.col),
            error: value.asError() ?? "Unknown formula error"
          });
        }
        value.free();
      }
    }
  } catch {
    inspectionComplete = false;
  }
  inspectionComplete &&= inspected === parsedFormulaCount;
  const errorCount = inspectionComplete ? resultErrors : null;
  let reason: string | null = null;
  if (sourceFormulaCount !== null && parsedFormulaCount !== sourceFormulaCount) reason = "formula-import-mismatch";
  else if (resultErrors > 0 || (engineErrorCount ?? 0) > 0) reason = "formula-errors";
  else if ((nonnegativeInteger(stats.circularReferences) ?? 0) > 0) reason = "circular-references";
  else if (stats.converged === false) reason = "not-converged";
  else if (!inspectionComplete) reason = "result-inspection-incomplete";
  else if (evaluatedFormulaCount === null || engineErrorCount === null || nonnegativeInteger(stats.formulaCount) === null
    || nonnegativeInteger(stats.circularReferences) === null || typeof stats.converged !== "boolean") reason = "engine-stats-unavailable";
  else if (stats.formulaCount !== parsedFormulaCount || evaluatedFormulaCount < parsedFormulaCount) reason = "evaluation-incomplete";
  else if (sourceFormulaCount === null) reason = "source-formula-inventory-unavailable";
  return {
    ...calculationReport(reason ? "partial" : "complete", reason, parsedFormulaCount, sourceFormulaCount),
    evaluatedFormulaCount, engineErrorCount, errorCount, durationMs, issues
  };
}

export function cellCalculationDiagnostic(
  worksheet: ReturnType<Workbook["getSheet"]>,
  row: number,
  col: number,
  cachedValue: string | undefined,
  report: XlsxCalculationReport,
  hasCalculatedValues = true,
): XlsxCellCalculationDiagnostic {
  if (!worksheet.getFormulaAt(row, col)) return { source: "literal", error: null };
  const value = worksheet.getCalculatedValueAt(row, col);
  const error = value.is_error ? value.asError() ?? "Unknown formula error" : null;
  value.free();
  if (error && hasCalculatedValues) return { source: "calculated", error };
  // Aggregate partial coverage cannot establish whether this particular cell was evaluated.
  const source = report.status === "complete" ? "calculated"
    : report.status === "partial" || report.status === "calculating" ? "unknown"
    : !hasCalculatedValues && cachedValue !== undefined ? "saved" : "unknown";
  return { source, error };
}
