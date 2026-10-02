# Modus React XLSX viewer

Based on upstream `0.16.4` (`873cd65bf3d04c4d1ff4f61263c7a2544b6571d0`) and the Modus `0.1.23-modus.1` WASM engine.

Preserved behavior:

- Serializable external add-in values reach the engine in normal, worker, and deferred loads. Formula text remains intact; unresolved calls retain cached values.
- `recalculate` accepts updated external add-in values without reloading the workbook, including
  worker-backed read-only workbooks. Calculation skips the engine's known 5,000-formula trap.
- `revealCell` selects and centers off-screen search results.
- `selections`, append/toggle operations, and canvas Ctrl/Cmd drag support non-contiguous selections. Re-adding an existing range preserves the other regions.
- Conditional formatting retains relative references, negated references (`lessThan -$H$13`, the plus-or-minus flux threshold), comparison/ABS/AND expressions, text and blank rules, and cached numeric fallback. These extend upstream's styled rules and retain its priority handling.
- Copy and paste carry each cell's formatting and its stored (or calculated) value, so pasted numbers stay numbers and formatting travels as in Excel; styles are stored once per distinct style. Pasting or undoing onto a cell that had no style of its own resets it to plain, since the engine reads that style back as `null` and has no call that clears one.
- Worker row batches reuse upstream's viewport calculation, also applying it to worker DOM rendering and excluding frozen rows from the scrolling batch. Frozen rows are fetched separately; sparse batch coverage and cache invalidation keep distant cells visible without requesting every preceding row.
- Worker sheet bounds retain merged/content extents and the initial view of leading blank rows/columns. Upstream supplies hidden axes, dimensions for populated and empty sheets, and absolute batch row indices.
- The atlas build plugin bundles US/world geography as JSON text parsed at runtime. This preserves ESM/CJS delivery without making downstream bundlers analyze hundreds of thousands of coordinate literals. The packaging regression checks data equality and syntax-tree size in both formats.

Retired customizations:

- Engine external-function parsing, callbacks, and unresolved cached-value handling now come from upstream. The viewer still needs the serializable CCH value bridge.
- Classic conditional comparisons, text matching, style merging, and priority handling use upstream's implementation. The custom expression evaluator no longer carries unused classic-operator branches.
- The old hidden-axis, dimension, and row-index string replacements are gone. Redundant blank-sheet sizing fallbacks have also been removed. The forced zero-based used-range minima remain: removing them causes upstream to scroll past leading blank rows and columns on open, as caught by the layout regression.
- The custom viewport-bound calculation is replaced by upstream's implementation with the two worker adjustments above.

The app configures `initWasm` with a bundler-resolved asset URL and explicitly chooses worker mode. Upstream otherwise routes all read-only workbooks into the worker, bypassing rich conditional formatting.

Validate with `pnpm typecheck`, `pnpm test`, and `pnpm build`. The real-package browser regression suite lives in `peasebell/e2e/tests/xlsx`; it covers DOM and canvas rendering plus main-thread, worker, and deferred loads. Run that suite before tagging a distribution.

After committing the source, run `pnpm build` and `node scripts/package-modus.mjs`. Copy `dist/modus-package/` into a distribution branch and create a new immutable `tb-dist-0.16.4-modus.N` tag (currently `modus.7`). The generated manifest records the exact source and upstream commits. The package and lockfile must reference the same engine artifact.

The release source retains the existing Modus GitHub workflows; importing upstream workflow changes requires separate repository permissions and review.

## Read-only review shortcuts

Ctrl/Cmd+Arrow resolves the next data-block boundary, with Shift preserving the selection anchor. Hidden axes are skipped; zero values and formulas returning empty text still count as occupied. Worker workbooks resolve the destination in one worker request. Requests are ordered and cancelled when selection, sheet, or document changes. Navigation stays within the viewer's finite worksheet extent, including its blank padding.

The `showFormulas` rendering prop displays formula text in both DOM and canvas cells without changing calculated values, clipboard values, exports, or aggregates. `getCellStyle` now also applies to blank worker cells, allowing precedents that point at blank cells to be highlighted.

The consuming app owns focus-scoped Find, shortcut bindings for formula view and precedents, and literal A1-reference highlighting. No engine API or workbook mutation is required. Named, table, dynamic, external, and 3D references are explicitly reported as incomplete by the app.
