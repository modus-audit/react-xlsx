# @extend-ai/react-xlsx — Modus fork (built artifact)

This orphan branch holds the **built** package consumed by peasebell as a git dependency. It threads
a serializable `externalFnValues` map into the Web Worker and rebuilds a synchronous `externalFnFn`
there, so the (forked) `@dukelib/sheets-wasm` engine resolves `[N]!FN(args)` CCH add-in calls during
`calculate()`. Closures can't cross the worker boundary, so the map is passed and the callback rebuilt
worker-side (`externalCallKey` is the shared key). Source lives on the fork's `modus-tb` branch.
Consumed via tag `tb-dist-<ver>`.
