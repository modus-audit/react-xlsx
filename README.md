# @extend-ai/react-xlsx — Modus fork (built artifact)

This orphan branch holds the built package consumed by peasebell as a git dependency.
Upstream base: `extend-hq/react-xlsx` v0.11.0. Consumed via tag `tb-dist-<ver>-modus.<n>`.

## Why 0.11.0

0.11.0 adds upstream's public `getCellStyle` hook for passive per-cell styling while retaining
the forked WASM engine version used for TB-link formula resolution.

## The Modus delta

1. `externalFnValues` worker passthrough resolves CCH add-in calls during calculation.
2. `controller.revealCell` selects, scrolls to, and repaints an off-screen target.
3. `selections` supports non-contiguous Ctrl/Cmd selections while preserving `selection` as
   the active region for backward compatibility.

## Source

Source is branch `codex/0.11-multi-region`, based on `modus-tb-0.11.0` with the multi-region
selection change ported from `feat/multi-region-selection`.

## Rebuilding

```bash
git checkout codex/0.11-multi-region
pnpm install
pnpm --filter @extend-ai/react-xlsx run build
```

Copy `packages/react-xlsx/dist/{index.cjs,index.d.cts,index.d.ts,index.js,xlsx-worker.js}` and
`packages/react-xlsx/package.json` into an orphan commit and tag it. Source maps and the WASM
binary are intentionally excluded; the WASM binary arrives through the package dependency.

Do not hand-edit `dist/`; patch source, rebuild, and publish a new artifact tag.
