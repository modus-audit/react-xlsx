/** Positions of the first and last shown rows (or columns) inside `start`..`end`, where `visible`
 *  is the shown indices in order. A selection can start or end on a hidden row (select-all over
 *  hidden rows, as in Excel), so it draws over the shown part; undefined when none of it shows. */
export function visibleSpan(visible: readonly number[], start: number, end: number): [number, number] | undefined {
  const first = firstAtLeast(visible, Math.min(start, end));
  const last = firstAtLeast(visible, Math.max(start, end) + 1) - 1;
  return first <= last ? [first, last] : undefined;
}

function firstAtLeast(sorted: readonly number[], value: number): number {
  let low = 0;
  let high = sorted.length;
  while (low < high) {
    const middle = (low + high) >> 1;
    if (sorted[middle] < value) low = middle + 1;
    else high = middle;
  }
  return low;
}

