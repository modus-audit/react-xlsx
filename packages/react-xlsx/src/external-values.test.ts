import assert from "node:assert/strict";
import { test } from "node:test";
import { externalValuesState } from "./external-values.ts";

test("an explicit recalculation's values hold until the prop itself changes", () => {
  const state = externalValuesState();
  const prop = { key: 1 };
  state.loaded(prop);
  assert.equal(state.propChanged(prop), false);
  const override = { key: 2 };
  state.recalculated(override);
  // The render after the explicit recalculation still sees the same prop: no revert.
  assert.equal(state.propChanged(prop), false);
  assert.equal(state.current(), override);
  assert.equal(state.recalculated(undefined), override);
  const next = { key: 3 };
  assert.equal(state.propChanged(next), true);
  assert.equal(state.recalculated(next), next);
});
