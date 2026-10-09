import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";
import initSheetsWasm, { Workbook } from "@dukelib/sheets-wasm";
import { resolveWorkbookColor } from "./colors.ts";

test("resolves imported engine theme font colors without losing black, white, or tint", async () => {
  await initSheetsWasm({
    module_or_path: readFileSync(new URL(import.meta.resolve("@dukelib/sheets-wasm/duke_sheets_wasm_bg.wasm")))
  });
  const themePalette = { colorsByIndex: { 0: "#ffffff", 1: "#000000", 4: "#0000ff" } };
  const source = new Workbook();
  source.addSheet("Theme colors");
  const sourceSheet = source.getSheet(0);
  const cases = [
    { address: "A1", themeIndex: 0, tint: 0, expected: "#ffffff" },
    { address: "A2", themeIndex: 1, tint: 0, expected: "#000000" },
    { address: "A3", themeIndex: 4, tint: 0.5, expected: "#8080ff" }
  ];
  for (const { address, themeIndex, tint } of cases) {
    sourceSheet.setCellStyle(address, { font: { color: { colorType: "theme", themeIndex, tint } } });
  }
  const workbook = Workbook.fromBytes(source.saveXlsxBytes());
  try {
    const sheet = workbook.getSheet(0);
    for (const { address, themeIndex, tint, expected } of cases) {
      const color = sheet.getCellStyle(address).font.color;
      assert.equal(color.themeIndex, themeIndex);
      assert.equal(resolveWorkbookColor(color, themePalette), expected, address);
      assert.equal(resolveWorkbookColor({ theme: themeIndex, tint }, themePalette), expected);
    }
    assert.equal(resolveWorkbookColor({ theme: 0, themeIndex: 1 }, themePalette), "#ffffff");
    assert.equal(resolveWorkbookColor({ hex: "ff123456", themeIndex: 0 }, themePalette), "#123456");
  } finally {
    workbook.free();
    source.free();
  }
});
