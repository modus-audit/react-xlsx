/** External add-in function ([N]!FN(args), e.g. CCH `[1]!TBLink(...)`) host resolution.
 *
 * The forked duke-sheets engine parses `[N]!FN(args)` into a callable node and, during
 * `calculate()`, invokes an `externalFnFn(name, args) -> value | null` host callback for each one
 * (returning null leaves the cell's cached value untouched). A callback is a closure and cannot be
 * `postMessage`d into the worker where the engine runs, so the host passes a *serializable* map of
 * pre-resolved values instead; the worker rebuilds the callback locally from it.
 *
 * {@link externalCallKey} is the single source of truth for how that map is keyed — the host builds
 * the map with it, and the worker looks up with it, so the two never drift. */

// Unit-separator (0x01): won't appear in CCH function names or args, so keys are unambiguous.
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
