import type { ExternalFnValues } from "./external-fn";

/** Which external add-in values the workbook calculates with: the `externalFnValues` prop, or the
 *  values of the last explicit `recalculate(values)` until the prop itself changes. */
export function externalValuesState() {
  let prop: ExternalFnValues | undefined;
  let current: ExternalFnValues | undefined;
  return {
    /** The values for a calculation: a load, an edit, undo or redo. */
    current: () => current,
    /** A load takes the prop's values. */
    loaded(values: ExternalFnValues | undefined) {
      prop = values;
      current = values;
    },
    /** An explicit recalculation; without values it keeps the current ones. */
    recalculated(values: ExternalFnValues | undefined) {
      if (values !== undefined) current = values;
      return current;
    },
    /** Whether a render's prop is new, and so should recalculate the loaded workbook. */
    propChanged(values: ExternalFnValues | undefined) {
      if (values === prop) return false;
      prop = values;
      return true;
    },
  };
}
