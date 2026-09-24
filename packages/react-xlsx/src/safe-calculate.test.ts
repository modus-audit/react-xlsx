import assert from "node:assert/strict";
import { test } from "node:test";
import type { Workbook } from "@dukelib/sheets-wasm";
import { externalCallKey, makeExternalFn } from "./external-fn.ts";
import { tryRecalculate } from "./safe-calculate.ts";

test("recalculation passes external add-in values to the engine", () => {
  const key = externalCallKey("TBLink", ["Ledger"]);
  const workbook = {
    calculate(options?: { externalFnFn?: (name: string, args: string[]) => string | number | null }) {
      assert.equal(options?.externalFnFn?.("TBLink", ["Ledger"]), 123.45);
      assert.equal(options?.externalFnFn?.("TBLink", ["Missing"]), null);
    },
  } as unknown as Workbook;

  const calcOptions = { externalFnFn: makeExternalFn({ [key]: 123.45 }) };
  assert.deepEqual(tryRecalculate(workbook, calcOptions), {
    calculated: true,
    error: null,
  });
});
