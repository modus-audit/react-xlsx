import assert from "node:assert/strict";
import { test } from "node:test";
import { resizeHitSlopPx } from "./resize-hit-slop.ts";

test("a default row leaves most of its header for selecting", () => {
  const slop = resizeHitSlopPx(20);
  assert.ok(20 - 2 * slop >= 12, `only ${20 - 2 * slop}px of a 20px row selects`);
});

test("wide columns keep the full grab zone and tiny rows keep a usable one", () => {
  assert.equal(resizeHitSlopPx(161), 8);
  assert.equal(resizeHitSlopPx(6), 2);
});
