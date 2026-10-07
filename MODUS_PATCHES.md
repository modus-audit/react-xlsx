# Modus React XLSX viewer

Based on upstream `0.16.4` (`873cd65bf3d04c4d1ff4f61263c7a2544b6571d0`) and the Modus `0.1.23-modus.2` WASM engine.

Preserved behavior:

- Serializable external add-in values reach the engine in normal, worker, and deferred loads, and every later calculation uses the latest `externalFnValues`: recalculation after each edit, undo and redo. Changing them recalculates once, in place, without reloading. Formula text remains intact; unresolved calls retain cached values.
- `readOnly` switches editing on or off in place after load, without reloading the workbook.
- `onBeforeEdit(edit)` sees every workbook edit before it happens and can refuse it: cell content (typing, clearing, pasting, filling), style, merge, unmerge, resize, undo/redo, sheet and name changes, and drawings. Each `XlsxEdit` carries the cells it writes (a paste's extent from the active cell, a fill's target), a lazily computed net formula-count change, and for one formula entry its normalized text and any problem the engine would mishandle silently (unbalanced syntax, a missing sheet). Without a hook, formula entries with a problem are refused.
- Excel entry: text typed into a cell that starts with `=` is a formula, and the cell editor opens a formula cell on its `=formula`.
- The viewer remembers its last copy or cut page-wide (forgotten on any copy outside a grid), so a paste event or async-clipboard paste whose text matches it keeps formulas and formatting. `cutSelection()` copies then clears, and pasting a cut moves formulas unchanged; `paste({ valuesOnly })` is Excel's Paste Values; Ctrl/Cmd+Shift+V pastes values (reading the clipboard when the browser fires no paste event, reporting a refusal through `onClipboardError`); Ctrl/Cmd+X cuts.
- `autoFit(axis, indices?)` fits columns or rows to their text in one undo step (by default every selected one), and double-clicking a header border fits that column or row, or the whole selection when the border belongs to it.
- `recalculate` accepts updated external add-in values without reloading the workbook, including
  worker-backed read-only workbooks. The engine (`wasm-dist-0.1.23-modus.2`) calculates serially, so the old 5,000-formula trap is gone; `autoCalculateFormulaLimit` (default 1,000) sets how many formulas a main-thread workbook may have and still recalculate on load and after edits.
- `controller.calculation` reports execution coverage on initial, deferred, and explicit recalculation.
  It compares the source OOXML formula inventory with the imported inventory, retains the engine
  statistics, and inspects typed formula results because engine `errors` misses some cell errors.
  Failed or skipped calculations retain saved values where no result is available. Calculated errors
  replace saved numbers in cells, with formula-specific corner-click details in DOM and canvas renderers;
  stale async results are ignored.
- `revealCell` selects and centers off-screen search results.
- `selections`, append/toggle operations, and canvas Ctrl/Cmd drag support non-contiguous selections. Re-adding an existing range preserves the other regions.
- Conditional formatting retains relative references, negated references (`lessThan -$H$13`, the plus-or-minus flux threshold), comparison/ABS/AND expressions, text and blank rules, and cached numeric fallback. These extend upstream's styled rules and retain its priority handling.
- Copy and paste carry each cell's formatting and its stored (or calculated) value, so pasted numbers stay numbers and formatting travels as in Excel; styles are stored once per distinct style, and a merged block's covered cells carry their style (Excel keeps a merge's edge borders there). The async `copySelectionToClipboard` writes plain text and HTML only: browsers refuse custom types there, which made it copy nothing.
- The engine reads an unstyled cell's style as `null`, has no call that clears a style, and patches the fields it is given. Pasting a plain cell, filling from one, or undoing a first format writes every field back to the workbook's default format (its first cell format's font), which reads back as `null` again.
- Row and column header borders grab for resizing within a sixth of the header (2 to 8 px) rather than a fixed 8 px, which took 16 of a default 20 px row and turned most clicks on row numbers into resizes. Ctrl/Cmd-click on a header adds that row or column to the selection, and the corner above the row numbers selects every cell. The DOM grid's resize handles use the same grab zone.
- Row AutoFit wraps a merge that spans columns of a single row at the merge's full width (Excel's AutoFit skips merged cells, which left wrapped conclusions cut off). Merges taller than one row are still skipped.
- `resizeRows` and `resizeColumns` apply many sizes as one undo step (one state update on worker-backed sheets), for fitting a whole selection; `resizeRow` and `resizeColumn` use them.
- Unmerge clears every merged block the selection touches, as Excel does; the engine only unmerges an exact merged range.
- Worker row batches reuse upstream's viewport calculation, also applying it to worker DOM rendering and excluding frozen rows from the scrolling batch. Frozen rows are fetched separately; sparse batch coverage and cache invalidation keep distant cells visible without requesting every preceding row.
- Worker sheet bounds retain merged/content extents and the initial view of leading blank rows/columns. Upstream supplies hidden axes, dimensions for populated and empty sheets, and absolute batch row indices.
- The atlas build plugin bundles US/world geography as JSON text parsed at runtime. This preserves ESM/CJS delivery without making downstream bundlers analyze hundreds of thousands of coordinate literals. The packaging regression checks data equality and syntax-tree size in both formats.

Retired customizations:

- Engine external-function parsing, callbacks, and unresolved cached-value handling now come from upstream. The viewer still needs the serializable CCH value bridge.
- Classic conditional comparisons, text matching, style merging, and priority handling use upstream's implementation. The custom expression evaluator no longer carries unused classic-operator branches.
- The old hidden-axis, dimension, and row-index string replacements are gone. Redundant blank-sheet sizing fallbacks have also been removed. The forced zero-based used-range minima remain: removing them causes upstream to scroll past leading blank rows and columns on open, as caught by the layout regression.
- The custom viewport-bound calculation is replaced by upstream's implementation with the two worker adjustments above.

The app configures `initWasm` with a bundler-resolved asset URL and explicitly chooses worker mode. Upstream otherwise routes all read-only workbooks into the worker, bypassing rich conditional formatting.

Validate with `pnpm typecheck`, `pnpm test`, and `pnpm build`. The real-package browser regression suite lives in `peasebell/excel-viewer-e2e`; it covers DOM and canvas rendering plus main-thread, worker, and deferred loads. Run that suite before tagging a distribution.

After committing the source, run `pnpm build` and `node scripts/package-modus.mjs`. Copy `dist/modus-package/` into a distribution branch and use an immutable commit or a new `tb-dist-0.16.4-modus.N` tag (currently `modus.22`). The generated manifest records the exact source and upstream commits. The package and lockfile must reference the same engine artifact.

The release source retains the existing Modus GitHub workflows; importing upstream workflow changes requires separate repository permissions and review.

## Read-only review shortcuts

Ctrl/Cmd+Arrow resolves the next data-block boundary, with Shift preserving the selection anchor. Hidden axes are skipped; zero values and formulas returning empty text still count as occupied. Worker workbooks resolve the destination in one worker request. Requests are ordered and cancelled when selection, sheet, or document changes. Navigation stays within the viewer's finite worksheet extent, including its blank padding.

The `showFormulas` rendering prop displays formula text in both DOM and canvas cells without changing calculated values, clipboard values, exports, or aggregates. `getCellStyle` now also applies to blank worker cells, allowing precedents that point at blank cells to be highlighted.

The consuming app owns focus-scoped Find, shortcut bindings for formula view and precedents, and literal A1-reference highlighting. No engine API or workbook mutation is required. Named, table, dynamic, external, and 3D references are explicitly reported as incomplete by the app.

## Calculation diagnostics

`XlsxCalculationReport` is exported and available through `controller.calculation`. Its revision
identifies each calculation attempt, including updated external values. `complete` means that the
engine finished, the source/import/evaluated formula inventories agree, and inspected formula
results contain no known errors. It does not establish equivalence with Excel, current external
inputs, or financial correctness. `partial`, `skipped`, and `failed` give a reason; unavailable
counts are `null` rather than assumed zero. Legacy XLS source inventories are unavailable.

`errorCount` counts typed error-valued formula cells; `engineErrorCount` preserves the engine's
separate statistic, which can include parsing failures without a typed error cell. `issues` contains
all typed formula errors as `{ sheet, cell, error }`. The application owns presentation, telemetry,
external-input readiness, and issue navigation.

`getCellCalculationDiagnostic` describes the selected main-thread cell or an available worker
snapshot. Worker `getCellSnapshotAsync` responses include the same `diagnostic`: `source` is
`literal`, `calculated`, `saved`, or `unknown`, with the typed `error` if present.
Typed formula errors remain visible instead of using saved numeric fallbacks. DOM and canvas cells
mark them with a small yellow corner. Clicking that corner opens a compact popup anchored to the
cell in both renderers; pointer movement does not open or reposition it. The popup shows only a
short explanation; the code remains in the cell and calculation report. Apps can use their own
tooltip theme through `formulaErrorTooltipClassName`. The popup closes on another click, Escape,
scrolling, resizing, and calculation or sheet changes. Unknown error codes receive a generic explanation.

The source inventory counts only SpreadsheetML formula children of worksheet cells, including
shared formula entries; sparkline, validation, and conditional-format extension references do not
count as formula cells. A new calculating or failed attempt clears earlier execution metrics.
In-memory edits invalidate execution coverage (`workbook-edited` when automatic calculation is
disabled) and make the original source inventory unavailable for later comparison. Partial
coverage and values retained from an earlier calculation have conservative `unknown` provenance.

## Formula copy and fill (.13)

Copy carries the original cell address so relative A1 references relocate on paste; absolute and mixed references retain their locked axes. One lexical helper also serves fill in every direction, preserving quoted strings, sheet qualifiers, names and structured references. The calculator and WASM engine remain unchanged.
