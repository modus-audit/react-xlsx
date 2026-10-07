import assert from "node:assert/strict";
import { test } from "node:test";
import { needsMainThreadReload } from "./load-mode.ts";

const loaded = { isWorkerBacked: true, isLoading: false, requestedReadOnly: false, forcedReadOnly: false };

test("turning read-only off reloads a worker-backed workbook on the main thread, once loaded", () => {
  assert.equal(needsMainThreadReload(loaded), true);
  assert.equal(needsMainThreadReload({ ...loaded, isLoading: true }), false);
  assert.equal(needsMainThreadReload({ ...loaded, requestedReadOnly: true }), false);
  assert.equal(needsMainThreadReload({ ...loaded, isWorkerBacked: false }), false);
});

test("a workbook forced read-only by its size stays on the worker", () => {
  assert.equal(needsMainThreadReload({ ...loaded, forcedReadOnly: true }), false);
});
