const MAX_ROW = 1_048_576;
const MAX_COL = 16_384;
const identifier = /[\p{L}\p{N}_.$\\]/u;
const sheetRange = /:\s*[\p{L}\p{N}_.$\\]+!/uy;
const cell = /^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)$/;
const axisRange = /(\$?[A-Za-z]{1,3}|\$?[1-9]\d*)(\s*:\s*)(\$?[A-Za-z]{1,3}|\$?[1-9]\d*)/y;

function columnNumber(letters: string): number {
  return [...letters.toUpperCase()].reduce((number, letter) => number * 26 + letter.charCodeAt(0) - 64, 0);
}

function columnLetters(number: number): string {
  let result = "";
  while (number > 0) {
    number -= 1;
    result = String.fromCharCode(65 + number % 26) + result;
    number = Math.floor(number / 26);
  }
  return result;
}

function shiftedAxis(token: string, delta: number, column: boolean): string {
  const absolute = token.startsWith("$");
  const value = token.replace(/^\$/, "");
  const original = column ? columnNumber(value) : Number(value);
  const limit = column ? MAX_COL : MAX_ROW;
  if (original > limit) return token;
  const next = original + (absolute ? 0 : delta);
  if (next < 1 || next > limit) return "#REF!";
  const text = column ? columnLetters(next) : String(next);
  return `${absolute ? "$" : ""}${column && value === value.toLowerCase() ? text.toLowerCase() : text}`;
}

function shiftedCell(token: string, rows: number, cols: number): string {
  const match = cell.exec(token);
  if (!match || columnNumber(match[2]) > MAX_COL || Number(match[4]) > MAX_ROW) return token;
  const col = shiftedAxis(match[1] + match[2], cols, true);
  const row = shiftedAxis(match[3] + match[4], rows, false);
  return col === "#REF!" || row === "#REF!" ? "#REF!" : col + row;
}

/** Copy semantics for A1 formulas. Scan lexical spans so strings, sheet names, function names,
 *  named operands and structured-reference brackets never get mistaken for cell references.
 *  Keep the original formula text; this does not parse or calculate the expression. */
export function copyFormula(formula: string, rows: number, cols: number): string {
  if (rows === 0 && cols === 0) return formula;
  let result = "";
  let index = 0;
  while (index < formula.length) {
    const start = index;
    const char = formula[index];
    if (char === '"' || char === "'") {
      index += 1;
      while (index < formula.length) {
        if (formula[index++] !== char) continue;
        if (formula[index] !== char) break;
        index += 1;
      }
      result += formula.slice(start, index);
      continue;
    }
    if (char === "[") {
      let depth = 1;
      index += 1;
      while (index < formula.length && depth > 0) {
        // In structured references an apostrophe escapes the following special character.
        if (formula[index] === "'") { index += 2; continue; }
        if (formula[index] === "[") depth += 1;
        if (formula[index] === "]") depth -= 1;
        index += 1;
      }
      result += formula.slice(start, index);
      continue;
    }
    if (!identifier.test(char)) { result += char; index += 1; continue; }

    axisRange.lastIndex = index;
    const range = axisRange.exec(formula);
    if (range && !identifier.test(formula[index + range[0].length] ?? "") &&
        formula[index + range[0].length] !== "!") {
      const columns = /[A-Za-z]/.test(range[1]);
      if (columns === /[A-Za-z]/.test(range[3])) {
        result += shiftedAxis(range[1], columns ? cols : rows, columns) + range[2] +
          shiftedAxis(range[3], columns ? cols : rows, columns);
        index += range[0].length;
        continue;
      }
    }
    while (index < formula.length && identifier.test(formula[index])) index += 1;
    const token = formula.slice(start, index);
    let following = index;
    while (following < formula.length && /\s/.test(formula[following])) following += 1;
    // A1! is a sheet name; LOG10(...) is a function; A1[...] is a table identifier.
    sheetRange.lastIndex = following;
    const qualifier = sheetRange.test(formula);
    result += formula[following] === "!" || formula[following] === "(" || formula[following] === "[" || qualifier
      ? token : shiftedCell(token, rows, cols);
  }
  return result;
}

/** Repeating fills are anchored at the source, including when dragging above or to its left. */
export function fillSourceIndex(target: number, start: number, size: number): number {
  return start + ((target - start) % size + size) % size;
}
