import type { Workbook } from "@dukelib/sheets-wasm";
import { strFromU8, unzipSync } from "fflate";

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
  source: "literal" | "calculated" | "saved" | "saved-fallback" | "unknown";
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

/** Read the input inventory separately so missing imported formulas cannot appear complete. */
export function countSourceWorkbookFormulas(bytes: Uint8Array): number | null {
  if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) return null;
  try {
    const archive = unzipSync(bytes, { filter: (entry) => /^xl\/worksheets\/[^/]+\.xml$/i.test(entry.name) });
    let count = 0;
    for (const data of Object.values(archive)) {
      const xml = strFromU8(data);
      count += countWorksheetCellFormulas(xml);
    }
    return count;
  } catch {
    return null;
  }
}

// A DOM-free scan for workers: extension range references (xm:f) are not formula cells.
function countWorksheetCellFormulas(xml: string): number {
  const spreadsheetNamespaces = new Set([
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "http://purl.oclc.org/ooxml/spreadsheetml/main"
  ]);
  type Element = { name: string; namespace: string; namespaces: Record<string, string> };
  const stack: Element[] = [];
  const formulaPath = ["worksheet", "sheetData", "row", "c"];
  let count = 0;
  const tags = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<[^>"']*(?:"[^"]*"[^>"']*|'[^']*'[^>"']*)*>/g;
  for (const match of xml.matchAll(tags)) {
    const tag = match[0];
    if (tag.startsWith("<!") || tag.startsWith("<?")) continue;
    if (tag.startsWith("</")) { stack.pop(); continue; }
    const qualifiedName = /^<([^\s/>]+)/.exec(tag)?.[1];
    if (!qualifiedName) continue;
    let namespaces = stack[stack.length - 1]?.namespaces ?? Object.create(null) as Record<string, string>;
    if (/\sxmlns(?::|\s*=)/.test(tag)) {
      namespaces = Object.create(namespaces) as Record<string, string>;
      for (const declaration of tag.matchAll(/\sxmlns(?::([\w.-]+))?\s*=\s*(["'])(.*?)\2/g)) {
        namespaces[declaration[1] ?? ""] = declaration[3] ?? "";
      }
    }
    const separator = qualifiedName.indexOf(":");
    const name = separator < 0 ? qualifiedName : qualifiedName.slice(separator + 1);
    const prefix = separator < 0 ? "" : qualifiedName.slice(0, separator);
    const namespace = namespaces[prefix] ?? "";
    if (name === "f" && spreadsheetNamespaces.has(namespace) && stack.length === 4
      && stack.every((element, index) => element.name === formulaPath[index]
        && spreadsheetNamespaces.has(element.namespace))) count += 1;
    if (!/\/\s*>$/.test(tag)) stack.push({ name, namespace, namespaces });
  }
  return count;
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
          if (issues.length < 20) issues.push({
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
  if (error && cachedValue !== undefined) return { source: "saved-fallback", error };
  // Aggregate partial coverage cannot establish whether this particular cell was evaluated.
  const source = report.status === "complete" ? "calculated"
    : report.status === "partial" || report.status === "calculating" ? "unknown"
    : !hasCalculatedValues && cachedValue !== undefined ? "saved" : "unknown";
  return { source, error };
}
