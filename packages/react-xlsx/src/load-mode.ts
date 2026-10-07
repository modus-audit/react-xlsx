/** The worker renders read-only; a workbook it loaded must reload on the main thread to edit. A
 *  workbook forced read-only by its size stays on the worker. */
export function needsMainThreadReload(state: {
  isWorkerBacked: boolean;
  isLoading: boolean;
  requestedReadOnly: boolean;
  forcedReadOnly: boolean;
}): boolean {
  return state.isWorkerBacked && !state.isLoading && !state.requestedReadOnly && !state.forcedReadOnly;
}
