/** The viewer's last copy or cut, page-wide like the system clipboard it shadows. Browsers keep the
 *  viewer's own cell format out of the async clipboard (it takes plain text and HTML only), so a
 *  menu paste, or a paste event without that format, uses the remembered cells when the clipboard
 *  text still matches them. */

export const INTERNAL_CLIPBOARD_MIME = "application/x-react-xlsx-range+json";
/** Marks a grid, so a copy anywhere else on the page forgets the remembered cells. */
export const GRID_SURFACE_ATTRIBUTE = "data-react-xlsx-grid";

let lastCopy: { text: string; structured: string } | null = null;
let watching = false;

export function forgetCopy() {
  lastCopy = null;
}

function watchOtherCopies() {
  if (watching || typeof document === "undefined") return;
  watching = true;
  const forget = (event: Event) => {
    const target = event.target;
    const inGrid =
      typeof Element !== "undefined" && target instanceof Element && target.closest(`[${GRID_SURFACE_ATTRIBUTE}]`);
    if (!inGrid) forgetCopy();
  };
  document.addEventListener("copy", forget, true);
  document.addEventListener("cut", forget, true);
}

export function rememberCopy(text: string, structured: string) {
  watchOtherCopies();
  lastCopy = { text, structured };
}

const sameText = (a: string, b: string) => a.replace(/\r\n/g, "\n").trimEnd() === b.replace(/\r\n/g, "\n").trimEnd();

/** The remembered cells, when `text` is what was copied with them. */
export function rememberedCells(text: string): string | null {
  return lastCopy && sameText(text, lastCopy.text) ? lastCopy.structured : null;
}

type PayloadCell = Record<string, unknown>;

function payloadCells(structured: string): { payload: Record<string, unknown>; cells: PayloadCell[] } | null {
  try {
    const payload: unknown = JSON.parse(structured);
    if (typeof payload !== "object" || payload === null || !("cells" in payload) || !Array.isArray(payload.cells)) {
      return null;
    }
    const cells = payload.cells.filter((cell): cell is PayloadCell => typeof cell === "object" && cell !== null);
    return { payload: payload as Record<string, unknown>, cells };
  } catch {
    return null;
  }
}

/** A cut moves formulas unchanged: without source coordinates, paste does not relocate them. */
export function withoutCopyOrigins(structured: string): string | null {
  const parsed = payloadCells(structured);
  if (!parsed) return null;
  const cells = parsed.cells.map(({ source: _source, ...cell }) => cell);
  return JSON.stringify({ ...parsed.payload, cells });
}

/** What a cut puts on the clipboard: the move payload when its clear goes ahead, else the copy. */
export function cutPayload(structured: string, clearAllowed: boolean): string {
  return (clearAllowed && withoutCopyOrigins(structured)) || structured;
}

/** Excel's Paste Values: values only, no formulas, formatting or merges. */
export function valuesOnly(structured: string): string | null {
  const parsed = payloadCells(structured);
  if (!parsed) return null;
  const cells = parsed.cells.flatMap((cell) => {
    const { rowOffset, colOffset } = cell;
    if (cell.styleOnly === true || typeof rowOffset !== "number" || typeof colOffset !== "number") return [];
    const value = typeof cell.value === "string" ? cell.value : "";
    return [{ rowOffset, colOffset, value, raw: cell.raw, formula: null }];
  });
  return JSON.stringify({ ...parsed.payload, cells, merges: [], styles: [] });
}
