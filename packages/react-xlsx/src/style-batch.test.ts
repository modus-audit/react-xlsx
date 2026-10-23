import assert from "node:assert/strict";
import { test } from "node:test";
import { applyStyleBatch, requireEditableStyle } from "./style-batch.ts";

test("a later inherited cell refuses a grouped formatting operation before any write", () => {
  let writes = 0;
  let allowed = 0;
  const edits = [{ range: { start: { row: 0, col: 0 }, end: { row: 0, col: 1 } }, style: { font: { bold: true } } }];
  assert.throws(() => applyStyleBatch(edits, {
    allow: () => { allowed++; return true; },
    capture: cell => {
      const own = cell.col === 0 ? { font: { bold: false } } : null;
      requireEditableStyle(own, cell.col === 1 ? { numberFormat: { formatString: "000" } } : null);
      return own;
    },
    apply: () => { writes++; },
    restore: () => { throw new Error("No mutation should need rollback"); }
  }), /inherit row or column formatting/);
  assert.equal(writes, 0);
  assert.equal(allowed, 0);
});

test("plain cells can receive a grouped style and restore their exact null styles together", () => {
  const styles = new Map<number, unknown>([[0, null], [1, null]]);
  const edits = [{ range: { start: { row: 0, col: 0 }, end: { row: 0, col: 1 } }, style: { font: { bold: true } } }];
  const restore = (cell: {col:number}, state: unknown) => { styles.set(cell.col, state); };
  const changes = applyStyleBatch(edits, {
    allow: () => true,
    capture: cell => { const own = styles.get(cell.col); requireEditableStyle(own, null); return own; },
    apply: edit => { for (let col = edit.range.start.col; col <= edit.range.end.col; col++) styles.set(col, edit.style); },
    restore
  });
  assert.equal(changes?.length, 2);
  for (const change of changes ?? []) restore(change.cell, change.before);
  assert.deepEqual([...styles.values()], [null, null]);
});
