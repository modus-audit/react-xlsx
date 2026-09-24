import { copyFileSync, mkdirSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname, join } from "node:path";

const reactXlsxRequire = createRequire(new URL("../../../packages/react-xlsx/package.json", import.meta.url));
const dukeEntrypoint = reactXlsxRequire.resolve("@dukelib/sheets-wasm");
const wasmSource = join(dirname(dukeEntrypoint), "duke_sheets_wasm_bg.wasm");
const publicDir = new URL("../public/", import.meta.url);

mkdirSync(publicDir, { recursive: true });
copyFileSync(wasmSource, new URL("duke_sheets_wasm_bg.wasm", publicDir));
