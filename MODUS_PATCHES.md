# Modus React XLSX viewer

Based on upstream `0.16.4` (`873cd65bf3d04c4d1ff4f61263c7a2544b6571d0`) and the Modus `0.1.23-modus.1` WASM engine.

Preserved behavior:

- Serializable external add-in values reach the engine in normal, worker, and deferred loads. Formula text remains intact; unresolved calls retain cached values.
- `revealCell` selects and centers off-screen search results.
- `selections`, append/toggle operations, and canvas Ctrl/Cmd drag support non-contiguous selections. Re-adding an existing range preserves the other regions.
- Conditional formatting retains relative references, comparison/ABS/AND expressions, text and blank rules, and cached numeric fallback. These extend upstream's styled rules and retain its priority handling.
- Worker row batches follow the actual scroll viewport and fetch frozen rows separately. Sparse batch coverage and cache invalidation keep distant cells visible without requesting every preceding row.
- Worker sheet bounds retain leading blank axes and merged/content extents. Upstream now supplies hidden axes, dimensions, and absolute batch row indices, replacing the old Vite string patches.

The app configures `initWasm` with a bundler-resolved asset URL and explicitly chooses worker mode. Upstream otherwise routes all read-only workbooks into the worker, bypassing rich conditional formatting.

Validate with `pnpm typecheck`, `pnpm test`, and `pnpm build`. The real-package browser regression suite lives in `peasebell/e2e/tests/xlsx`; it covers DOM and canvas rendering plus main-thread, worker, and deferred loads. Run that suite before tagging a distribution.

After committing the source, run `pnpm build` and `node scripts/package-modus.mjs`. Copy `dist/modus-package/` into a distribution branch and create a new immutable `tb-dist-0.16.4-modus.1` tag. The generated manifest records the exact source and upstream commits. The package and lockfile must reference the same engine artifact.

The release source retains the existing Modus GitHub workflows; importing upstream workflow changes requires separate repository permissions and review.
