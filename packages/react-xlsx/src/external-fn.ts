/** Serializable external-call values cross the worker boundary; missing entries preserve cached workbook values. */

const KEY_SEP = String.fromCharCode(1);

/** Canonical key for one external call. `args` are the engine's stringified argument values
 *  (CCH args are strings/numbers), matching how the host parsed them from the formula text. */
export function externalCallKey(name: string, args: readonly string[]): string {
  return [name, ...args].join(KEY_SEP);
}

/** Serializable map handed to the controller: `externalCallKey(name, args)` -> resolved value. */
export type ExternalFnValues = Record<string, string | number>;

/** Rebuild the engine callback from the serializable map (inside the worker). Returns `null` for
 *  unmapped calls so the engine preserves the cell's cached value. */
export function makeExternalFn(
  values: ExternalFnValues,
): (name: string, args: string[]) => string | number | null {
  return (name, args) => {
    const value = values[externalCallKey(name, args)];
    return value === undefined ? null : value;
  };
}
