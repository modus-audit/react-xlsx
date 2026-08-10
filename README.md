# @extend-ai/react-xlsx — Modus fork (built artifact)

This orphan branch holds the **built** package consumed by peasebell as a git dependency.
Upstream base: `extend-hq/react-xlsx` v0.11.0. Consumed via tag `tb-dist-<ver>-modus.<n>`.

## Why 0.11.0

0.11.0 adds upstream's public `getCellStyle` hook — per-cell style overrides that the canvas
renderer honors (`backgroundColor`, `border*`, `color`, `padding`, `textAlign`). That is the
supported way to shade cells without touching selection, scroll, or focus.

0.11.0 is also the *last* version that still requires `@dukelib/sheets-wasm` 0.1.13, so the
forked WASM engine pin (`modus-audit/duke-sheets` @ `wasm-dist-0.1.15-modus.1`) carries over
unchanged. Upstream 0.13.0 needs 0.1.17 and 0.15.1+ needs 0.1.21 — either would force
re-forking and rebuilding the Rust engine with the TB_LINK patches reapplied.

## The Modus delta

Two patches, neither upstreamed:

1. **`externalFnValues` / `externalFnFn` worker passthrough.** Threads a serializable
   `externalFnValues` map into the Web Worker and rebuilds a synchronous `externalFnFn` there,
   so the forked `@dukelib/sheets-wasm` engine resolves `[N]!FN(args)` CCH add-in calls during
   `calculate()`. Closures can't cross the worker boundary, so the map is passed and the
   callback rebuilt worker-side (`externalCallKey` is the shared key).
2. **`controller.revealCell`.** Select + scroll-into-view + selection repaint, centering a
   target that starts off-screen. Consumed by the beta Excel viewer's in-document search.

## Source

Source is branch **`modus-tb-0.11.0`** — `modus-tb` rebased onto upstream `v0.11.0`, keeping
its four original commits. `modus-tb` itself stays on the 0.10.2 base as the record of the
0.10.2 line.

The rebase took one conflict resolution, in three files: 0.11.0 added a `wasmSource` field to
the same worker `load` payload and dispatch that the TB-link patch extends with
`externalFnValues`. Both are kept.

One fix on top, not a straight replay: 0.11.0 introduced a **deferred worker load path**
(`getWorkerClient().loadWorkbook(deferredBuffer, …)`) that did not exist when the TB-link patch
was written, so the replayed patch left it unthreaded — large workbooks routed to the deferred
path would have silently skipped external-fn resolution. `externalFnValues` is now passed there
too, and added to both enclosing dependency arrays.

## Rebuilding

```bash
git checkout modus-tb-0.11.0
pnpm install
pnpm --filter @extend-ai/react-xlsx run build
```

Then copy `packages/react-xlsx/dist/{index.cjs,index.d.cts,index.d.ts,index.js,xlsx-worker.js}`
and a `version`-bumped `package.json` into an orphan commit and tag it. Source maps and the
`.wasm` are intentionally excluded; the `.wasm` arrives via the `@dukelib/sheets-wasm` dependency.

**Do not hand-edit `dist/`.** Releases through `tb-dist-0.10.2-modus.6` were produced that way,
which is why the 0.10.2 bundles and the `modus-tb` source had drifted apart. Patch `src/`,
rebuild, ship.
