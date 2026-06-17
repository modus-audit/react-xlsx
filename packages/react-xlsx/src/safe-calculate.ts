import type { Workbook } from "@dukelib/sheets-wasm";

const SHEET_REF_REGEX = /'((?:[^']|'')+)'!|([A-Za-z_\u0080-\uFFFF][\w.\u0080-\uFFFF]*)!/g;

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
};

export type SafeCalculateOptions = {
  reparse?: () => Workbook;
  // Passed straight through to `workbook.calculate(...)` — e.g. `{ externalFnFn }` so the engine can
  // resolve `[N]!FN(args)` add-in calls (CCH TBLink) against host-supplied values.
  calcOptions?: unknown;
};

// Pre-scans for formulas referencing missing sheets (which cause the Rust
// engine to panic into a wasm `unreachable` trap that poisons the Workbook
// instance). On trap, `reparse` is used to return a fresh usable instance.
export function safeCalculate(workbook: Workbook, options: SafeCalculateOptions = {}): SafeCalculateResult {
  if (hasUnresolvedSheetReferences(workbook)) {
    return { workbook, calculated: false, skipReason: "unresolved-sheet-refs" };
  }
  try {
    workbook.calculate(options.calcOptions);
    return { workbook, calculated: true, skipReason: null };
  } catch (err) {
    console.warn("[react-xlsx] workbook.calculate() trapped; falling back to cached formula values", err);
    if (options.reparse) {
      try {
        return { workbook: options.reparse(), calculated: false, skipReason: "calculate-trapped" };
      } catch (reparseErr) {
        console.warn("[react-xlsx] workbook reparse after calculate trap failed", reparseErr);
      }
    }
    return { workbook, calculated: false, skipReason: "calculate-trapped" };
  }
}

export function tryRecalculate(
  workbook: Workbook,
  calcOptions?: unknown,
): { calculated: boolean; error: unknown } {
  try {
    workbook.calculate(calcOptions);
    return { calculated: true, error: null };
  } catch (err) {
    console.warn("[react-xlsx] workbook.calculate() trapped during recalculation", err);
    return { calculated: false, error: err };
  }
}
