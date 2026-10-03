const MAX_SLOP_PX = 8;
const MIN_SLOP_PX = 2;

/** How far from a header border the pointer still grabs it to resize, for a header `sizePx` long.
 *  A fixed 8px took 16 of a default 20px row header, so clicking or dragging row numbers resized
 *  rows instead of selecting them. Excel's grab zone is a few pixels. */
export function resizeHitSlopPx(sizePx: number): number {
  return Math.max(MIN_SLOP_PX, Math.min(MAX_SLOP_PX, Math.floor(sizePx / 6)));
}
