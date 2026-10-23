import assert from "node:assert/strict";
import { test } from "node:test";
import { coerceUserEnteredValue, enteredPercentageFormat } from "./user-entered-value.ts";

test("percentage entry is numeric, while explicit text and existing format precision survive", () => {
  assert.equal(coerceUserEnteredValue("6%"), 0.06);
  assert.equal(coerceUserEnteredValue(" -1.25% "), -0.0125);
  assert.equal(coerceUserEnteredValue(".5%"), 0.005);
  assert.equal(coerceUserEnteredValue("'6%"), "6%");
  assert.equal(coerceUserEnteredValue("six%"), "six%");
  assert.equal(coerceUserEnteredValue("6%%"), "6%%");
  assert.equal(coerceUserEnteredValue("0.06"), 0.06);
  assert.equal(enteredPercentageFormat("6%", { numberFormat: { formatString: "0.00%" } }), null);
  assert.equal(enteredPercentageFormat("6%", { numberFormat: { formatString: "General" } }), "0%");
  assert.equal(enteredPercentageFormat("6.25%"), "0.00%");
  assert.equal(enteredPercentageFormat("'6%"), null);
});
