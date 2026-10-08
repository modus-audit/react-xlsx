// src/source-formula-inventory.ts
import { strFromU8, unzipSync } from "fflate";
var spreadsheetNamespaces = /* @__PURE__ */ new Set([
  "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
  "http://purl.oclc.org/ooxml/spreadsheetml/main"
]);
var relationshipNamespaces = /* @__PURE__ */ new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  "http://purl.oclc.org/ooxml/officeDocument/relationships"
]);
var packageRelationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
function decodeAttribute(value) {
  const entities = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" };
  return value.replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (_, entity) => entity.startsWith("#") ? String.fromCodePoint(entity[1] === "x" ? parseInt(entity.slice(2), 16) : Number(entity.slice(1))) : entities[entity]);
}
function scanXml(xml, visit) {
  const stack = [];
  const tags = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<[^>"']*(?:"[^"]*"[^>"']*|'[^']*'[^>"']*)*>/g;
  for (const match of xml.matchAll(tags)) {
    const tag = match[0];
    if (tag.startsWith("<!") || tag.startsWith("<?")) continue;
    if (tag.startsWith("</")) {
      stack.pop();
      continue;
    }
    const qualifiedName = /^<([^\s/>]+)/.exec(tag)?.[1];
    if (!qualifiedName) continue;
    const attributes = Object.fromEntries(Array.from(
      tag.matchAll(/\s([^\s=]+)\s*=\s*(["'])(.*?)\2/g),
      (attribute) => [attribute[1], decodeAttribute(attribute[3])]
    ));
    const namespaces = Object.create(stack[stack.length - 1]?.namespaces ?? null);
    for (const [name2, value] of Object.entries(attributes)) {
      if (name2 === "xmlns") namespaces[""] = value;
      else if (name2.startsWith("xmlns:")) namespaces[name2.slice(6)] = value;
    }
    const separator = qualifiedName.indexOf(":");
    const name = separator < 0 ? qualifiedName : qualifiedName.slice(separator + 1);
    const prefix = separator < 0 ? "" : qualifiedName.slice(0, separator);
    const element = { name, namespace: namespaces[prefix] ?? "", namespaces, attributes };
    stack.push(element);
    visit(stack);
    if (/\/\s*>$/.test(tag)) stack.pop();
  }
}
function spreadsheetPath(path, names) {
  return path.length === names.length && path.every((element, index) => element.name === names[index] && spreadsheetNamespaces.has(element.namespace));
}
function countSourceWorkbookFormulas(bytes) {
  if (bytes[0] !== 80 || bytes[1] !== 75) return null;
  try {
    const metadata = unzipSync(bytes, { filter: (entry) => entry.name === "xl/workbook.xml" || entry.name === "xl/_rels/workbook.xml.rels" });
    if (!metadata["xl/workbook.xml"] || !metadata["xl/_rels/workbook.xml.rels"]) return null;
    const sheetIds = [];
    scanXml(strFromU8(metadata["xl/workbook.xml"]), (path) => {
      if (!spreadsheetPath(path, ["workbook", "sheets", "sheet"])) return;
      const element = path[path.length - 1];
      const id = Object.entries(element.attributes).find(([name]) => {
        const [prefix, local] = name.split(":");
        return local === "id" && relationshipNamespaces.has(element.namespaces[prefix] ?? "");
      })?.[1];
      if (!id) throw new Error("Missing worksheet relationship");
      sheetIds.push(id);
    });
    const relationships = /* @__PURE__ */ new Map();
    scanXml(strFromU8(metadata["xl/_rels/workbook.xml.rels"]), (path) => {
      if (path.length === 2 && path[0]?.name === "Relationships" && path[1]?.name === "Relationship" && path.every((element) => element.namespace === packageRelationshipNamespace)) {
        relationships.set(path[1].attributes.Id, path[1].attributes);
      }
    });
    const worksheetPaths = /* @__PURE__ */ new Set();
    for (const id of sheetIds) {
      const relationship = relationships.get(id);
      if (!relationship?.Type || !relationship.Target || relationship.TargetMode === "External") return null;
      if (!relationship.Type.endsWith("/worksheet")) continue;
      const target = new URL(relationship.Target, "https://xlsx.invalid/xl/workbook.xml");
      if (target.origin !== "https://xlsx.invalid") return null;
      worksheetPaths.add(decodeURIComponent(target.pathname.slice(1)));
    }
    const worksheets = unzipSync(bytes, { filter: (entry) => worksheetPaths.has(entry.name) });
    let count = 0;
    for (const path of worksheetPaths) {
      if (!worksheets[path]) return null;
      scanXml(strFromU8(worksheets[path]), (elements) => {
        if (spreadsheetPath(elements, ["worksheet", "sheetData", "row", "c", "f"])) count += 1;
      });
    }
    return count;
  } catch {
    return null;
  }
}

// src/calculation-diagnostics.ts
function calculationReport(status, reason = null, parsedFormulaCount = 0, sourceFormulaCount2 = null) {
  return {
    status,
    reason,
    parsedFormulaCount,
    sourceFormulaCount: sourceFormulaCount2,
    formulaCount: Math.max(parsedFormulaCount, sourceFormulaCount2 ?? 0),
    evaluatedFormulaCount: null,
    errorCount: null,
    engineErrorCount: null,
    durationMs: null,
    revision: 0,
    issues: []
  };
}
function nonnegativeInteger(value) {
  return typeof value === "number" && Number.isSafeInteger(value) && value >= 0 ? value : null;
}
function cellAddress(row, col) {
  let column = col + 1;
  let label = "";
  while (column > 0) {
    label = String.fromCharCode(65 + (column - 1) % 26) + label;
    column = Math.floor((column - 1) / 26);
  }
  return `${label}${row + 1}`;
}
function inspectCalculation(workbook2, rawStats, parsedFormulaCount, sourceFormulaCount2, durationMs) {
  const stats = rawStats && typeof rawStats === "object" ? rawStats : {};
  const evaluatedFormulaCount = nonnegativeInteger(stats.cellsCalculated);
  const engineErrorCount = nonnegativeInteger(stats.errors);
  const issues = [];
  let resultErrors = 0;
  let inspected = 0;
  let inspectionComplete = true;
  try {
    for (let index = 0; index < workbook2.sheetCount; index += 1) {
      const sheet = workbook2.getSheet(index);
      const cells = sheet.formulaCells;
      if (!Array.isArray(cells)) {
        inspectionComplete = false;
        continue;
      }
      for (const cell of cells) {
        if (!cell || !Number.isInteger(cell.row) || !Number.isInteger(cell.col)) {
          inspectionComplete = false;
          continue;
        }
        const value = sheet.getCalculatedValueAt(cell.row, cell.col);
        inspected += 1;
        if (value.is_error) {
          resultErrors += 1;
          issues.push({
            sheet: workbook2.sheetNames[index] ?? String(index),
            cell: cellAddress(cell.row, cell.col),
            error: value.asError() ?? "Unknown formula error"
          });
        }
        value.free();
      }
    }
  } catch {
    inspectionComplete = false;
  }
  inspectionComplete &&= inspected === parsedFormulaCount;
  const errorCount = inspectionComplete ? resultErrors : null;
  let reason = null;
  if (sourceFormulaCount2 !== null && parsedFormulaCount !== sourceFormulaCount2) reason = "formula-import-mismatch";
  else if (resultErrors > 0 || (engineErrorCount ?? 0) > 0) reason = "formula-errors";
  else if ((nonnegativeInteger(stats.circularReferences) ?? 0) > 0) reason = "circular-references";
  else if (stats.converged === false) reason = "not-converged";
  else if (!inspectionComplete) reason = "result-inspection-incomplete";
  else if (evaluatedFormulaCount === null || engineErrorCount === null || nonnegativeInteger(stats.formulaCount) === null || nonnegativeInteger(stats.circularReferences) === null || typeof stats.converged !== "boolean") reason = "engine-stats-unavailable";
  else if (stats.formulaCount !== parsedFormulaCount || evaluatedFormulaCount < parsedFormulaCount) reason = "evaluation-incomplete";
  else if (sourceFormulaCount2 === null) reason = "source-formula-inventory-unavailable";
  return {
    ...calculationReport(reason ? "partial" : "complete", reason, parsedFormulaCount, sourceFormulaCount2),
    evaluatedFormulaCount,
    engineErrorCount,
    errorCount,
    durationMs,
    issues
  };
}
function cellCalculationDiagnostic(worksheet, row, col, cachedValue, report, hasCalculatedValues2 = true) {
  if (!worksheet.getFormulaAt(row, col)) return { source: "literal", error: null };
  const value = worksheet.getCalculatedValueAt(row, col);
  const error = value.is_error ? value.asError() ?? "Unknown formula error" : null;
  value.free();
  if (error && hasCalculatedValues2) return { source: "calculated", error };
  const source = report.status === "complete" ? "calculated" : report.status === "partial" || report.status === "calculating" ? "unknown" : !hasCalculatedValues2 && cachedValue !== void 0 ? "saved" : "unknown";
  return { source, error };
}

// src/data-navigation.ts
function findDataBoundary(request, sheet, hasContent) {
  const vertical = request.direction === "ArrowUp" || request.direction === "ArrowDown";
  const step = request.direction === "ArrowUp" || request.direction === "ArrowLeft" ? -1 : 1;
  const hidden = new Set((vertical ? sheet.hiddenRows : sheet.hiddenCols) ?? []);
  const max = vertical ? request.maxRow : request.maxCol;
  const usedMax = vertical ? sheet.maxUsedRow : sheet.maxUsedCol;
  const start = vertical ? request.cell.row : request.cell.col;
  const nextVisible = (position) => {
    let next2 = position + step;
    while (next2 >= 0 && next2 <= max && hidden.has(next2)) next2 += step;
    return next2 >= 0 && next2 <= max ? next2 : null;
  };
  const occupied = (position) => position <= usedMax && hasContent(
    vertical ? position : request.cell.row,
    vertical ? request.cell.col : position
  );
  let destination = start;
  let next = nextVisible(start);
  if (next !== null) {
    const contiguous = occupied(start) && occupied(next);
    while (next !== null) {
      const filled = occupied(next);
      if (contiguous && !filled) break;
      destination = next;
      if (!contiguous && filled) break;
      next = nextVisible(next);
    }
  }
  return vertical ? { row: destination, col: request.cell.col } : { row: request.cell.row, col: destination };
}
function worksheetHasContent(worksheet, row, col) {
  if (worksheet.getFormulaAt(row, col)) return true;
  const value = worksheet.getCalculatedValueAt(row, col);
  try {
    return !value.is_empty;
  } finally {
    value.free();
  }
}

// src/xlsx-worker.ts
import { strFromU8 as strFromU84, unzipSync as unzipSync3 } from "fflate";

// src/charts.ts
import { strFromU8 as strFromU82, strToU8 } from "fflate";
var CHART_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
var CHART_EX_REL_TYPE = "http://schemas.microsoft.com/office/2014/relationships/chartEx";
var CHART_STYLE_REL_TYPE = "http://schemas.microsoft.com/office/2011/relationships/chartStyle";
var CHART_COLOR_STYLE_REL_TYPE = "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";
var SERIES_COLORS = [
  "#4472c4",
  "#ed7d31",
  "#a5a5a5",
  "#ffc000",
  "#5b9bd5",
  "#70ad47",
  "#264478",
  "#9e480e",
  "#636363",
  "#997300"
];
function normalizeWorksheetVisibility(value) {
  return value === "hidden" || value === "veryHidden" ? value : "visible";
}
var EMU_PER_PIXEL = 9525;
var THEME_COLOR_INDEX_BY_NAME = {
  accent1: 4,
  accent2: 5,
  accent3: 6,
  accent4: 7,
  accent5: 8,
  accent6: 9,
  dk1: 1,
  dk2: 3,
  folHlink: 11,
  hlink: 10,
  lt1: 0,
  lt2: 2,
  tx1: 1,
  tx2: 3,
  bg1: 0,
  bg2: 2
};
var PRIMARY_CHART_TYPE_LOCAL_NAMES = [
  "barChart",
  "lineChart",
  "line3DChart",
  "stockChart",
  "radarChart",
  "scatterChart",
  "pieChart",
  "pie3DChart",
  "doughnutChart",
  "areaChart",
  "area3DChart",
  "bar3DChart",
  "ofPieChart",
  "bubbleChart",
  "surfaceChart",
  "surface3DChart"
];
function clampUnitInterval(value) {
  return Math.max(0, Math.min(1, value));
}
function isElementNode(node) {
  return node != null && node.nodeType === 1;
}
function normalizeHexColor(value) {
  const hex = value.replace(/^#/, "");
  if (hex.length === 8) {
    return `#${hex.slice(2).toLowerCase()}`;
  }
  if (hex.length === 6) {
    return `#${hex.toLowerCase()}`;
  }
  return null;
}
function resolveColorFromXmlFragment(fragment, themePalette) {
  if (!fragment) {
    return void 0;
  }
  const srgbMatch = fragment.match(/<a:srgbClr\b[^>]*\bval="([0-9a-fA-F]{6,8})"/i);
  if (srgbMatch?.[1]) {
    return normalizeHexColor(srgbMatch[1]) ?? void 0;
  }
  const schemeMatch = fragment.match(/<a:schemeClr\b[^>]*\bval="([^"]+)"[^>]*>([\s\S]*?)<\/a:schemeClr>/i) ?? fragment.match(/<a:schemeClr\b[^>]*\bval="([^"]+)"[^>]*/i);
  if (!schemeMatch?.[1]) {
    return void 0;
  }
  const baseColor = resolveThemeColor(schemeMatch[1], themePalette);
  if (!baseColor) {
    return void 0;
  }
  const transforms = schemeMatch[2] ?? "";
  let lightnessModifier = 1;
  let lightnessOffset = 0;
  for (const match of transforms.matchAll(/<a:(lumMod|lumOff|tint|shade)\b[^>]*\bval="(-?\d+(?:\.\d+)?)"/gi)) {
    const transform = match[1]?.toLowerCase();
    const rawValue = Number(match[2] ?? Number.NaN);
    if (!transform || !Number.isFinite(rawValue)) {
      continue;
    }
    if (transform === "lummod") {
      lightnessModifier *= rawValue / 1e5;
    } else if (transform === "lumoff") {
      lightnessOffset += rawValue / 1e5;
    } else if (transform === "tint") {
      lightnessOffset += (1 - lightnessOffset) * (rawValue / 1e5);
    } else if (transform === "shade") {
      lightnessModifier *= rawValue / 1e5;
    }
  }
  return applyLightnessTransform(baseColor, lightnessModifier, lightnessOffset) ?? void 0;
}
function readHexColorFromXmlFragment(fragment, preferLine = false, themePalette) {
  const source = preferLine ? fragment.match(/<a:ln\b[\s\S]*?<\/a:ln>/i)?.[0] ?? "" : fragment.match(/<a:solidFill\b[\s\S]*?<\/a:solidFill>/i)?.[0] ?? "";
  return resolveColorFromXmlFragment(source, themePalette);
}
function parseFallbackSeriesStylesFromChartXml(chartXml, themePalette) {
  const seriesBlocks = chartXml.match(/<c:ser\b[\s\S]*?<\/c:ser>/gi) ?? [];
  if (seriesBlocks.length === 0) {
    return [];
  }
  return seriesBlocks.map((seriesBlock) => {
    const shapeBlock = seriesBlock.match(/<c:spPr\b[\s\S]*?<\/c:spPr>/i)?.[0] ?? "";
    return {
      color: readHexColorFromXmlFragment(shapeBlock, false, themePalette),
      lineColor: readHexColorFromXmlFragment(shapeBlock, true, themePalette)
    };
  });
}
function parseFallbackPointStylesFromChartXml(chartXml, themePalette) {
  const chartDocument = parseXml(chartXml);
  if (chartDocument) {
    const parsedSeriesStyles = getLocalDescendants(chartDocument, "ser").map((seriesNode) => {
      const styles = [];
      for (const dataPointNode of getLocalChildren(seriesNode, "dPt")) {
        const indexValue = readChartNumericAttribute(dataPointNode, "idx");
        if (indexValue === void 0) {
          continue;
        }
        const shapeProperties = getFirstLocalChild(dataPointNode, "spPr");
        const lineStyle = resolveChartLineStyle(shapeProperties, themePalette);
        styles.push({
          color: resolveChartFillColor(shapeProperties, themePalette) ?? void 0,
          explosion: readChartNumericAttribute(dataPointNode, "explosion"),
          index: indexValue,
          lineColor: lineStyle.color ?? void 0
        });
      }
      return styles;
    });
    if (parsedSeriesStyles.some((styles) => styles.length > 0)) {
      return parsedSeriesStyles;
    }
  }
  const seriesBlocks = chartXml.match(/<c:ser\b[\s\S]*?<\/c:ser>/gi) ?? [];
  if (seriesBlocks.length === 0) {
    return [];
  }
  return seriesBlocks.map((seriesBlock) => {
    const pointBlocks = seriesBlock.match(/<c:dPt\b[\s\S]*?<\/c:dPt>/gi) ?? [];
    if (pointBlocks.length === 0) {
      return [];
    }
    const styles = [];
    for (const pointBlock of pointBlocks) {
      const indexMatch = pointBlock.match(/<c:idx\b[^>]*\bval="(-?\d+)"/i);
      const index = indexMatch?.[1] ? Number(indexMatch[1]) : Number.NaN;
      if (!Number.isFinite(index)) {
        continue;
      }
      const explosionMatch = pointBlock.match(/<c:explosion\b[^>]*\bval="(-?\d+(?:\.\d+)?)"/i);
      const explosionValue = explosionMatch?.[1] ? Number(explosionMatch[1]) : Number.NaN;
      styles.push({
        color: readHexColorFromXmlFragment(pointBlock, false, themePalette),
        explosion: Number.isFinite(explosionValue) ? explosionValue : void 0,
        index,
        lineColor: readHexColorFromXmlFragment(pointBlock, true, themePalette)
      });
    }
    return styles;
  });
}
function parseNumericPointCacheFromXmlFragment(fragment) {
  const pointMatches = Array.from(fragment.matchAll(/<c:pt\b[^>]*\bidx="(-?\d+)"[^>]*>[\s\S]*?<c:v>([^<]*)<\/c:v>[\s\S]*?<\/c:pt>/gi));
  if (pointMatches.length === 0) {
    return [];
  }
  const explicitPointCountMatch = fragment.match(/<c:ptCount\b[^>]*\bval="(\d+)"/i);
  const explicitPointCount = explicitPointCountMatch?.[1] ? Number(explicitPointCountMatch[1]) : Number.NaN;
  const maxIndex = pointMatches.reduce((max, match) => {
    const current = Number(match[1] ?? Number.NaN);
    return Number.isFinite(current) ? Math.max(max, current) : max;
  }, -1);
  const pointCount = Math.max(
    pointMatches.length,
    Number.isFinite(explicitPointCount) ? explicitPointCount : 0,
    maxIndex + 1
  );
  const values = Array.from({ length: pointCount }, () => null);
  for (const match of pointMatches) {
    const index = Number(match[1] ?? Number.NaN);
    const rawValue = (match[2] ?? "").trim();
    const numericValue = Number(rawValue);
    if (!Number.isFinite(index) || index < 0 || !Number.isFinite(numericValue)) {
      continue;
    }
    values[index] = numericValue;
  }
  return values;
}
function parseFallbackBubbleSizesFromChartXml(chartXml) {
  const seriesBlocks = chartXml.match(/<c:ser\b[\s\S]*?<\/c:ser>/gi) ?? [];
  if (seriesBlocks.length === 0) {
    return [];
  }
  return seriesBlocks.map((seriesBlock) => {
    const bubbleSizeBlock = seriesBlock.match(/<c:bubbleSize\b[\s\S]*?<\/c:bubbleSize>/i)?.[0] ?? "";
    if (!bubbleSizeBlock) {
      return [];
    }
    return parseNumericPointCacheFromXmlFragment(bubbleSizeBlock);
  });
}
function decodeChartXmlText(value) {
  return value.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&").replace(/&quot;/g, '"').replace(/&apos;/g, "'");
}
function normalizeChartTitleForMatch(value) {
  return (value ?? "").trim().replace(/\s+/g, " ").toLowerCase();
}
function extractChartTitleFromXml(chartXml) {
  const match = chartXml.match(/<c:title\b[\s\S]*?<a:t>([\s\S]*?)<\/a:t>/i);
  if (!match?.[1]) {
    return null;
  }
  const decoded = decodeChartXmlText(match[1]).trim();
  return decoded.length > 0 ? decoded : null;
}
function resolveArchiveFallbackBubbleSizes(archive, preferredTitle) {
  const preferred = normalizeChartTitleForMatch(preferredTitle);
  let bestScore = Number.NEGATIVE_INFINITY;
  let bestCandidate = [];
  for (const [path, bytes] of Object.entries(archive)) {
    if (!/\/charts\/chart\d+\.xml$/i.test(path)) {
      continue;
    }
    const chartXml = strFromU82(bytes);
    if (!/<c:bubbleChart\b/i.test(chartXml)) {
      continue;
    }
    const candidateBubbleSizes = parseFallbackBubbleSizesFromChartXml(chartXml);
    const hasCandidateValues = candidateBubbleSizes.some((seriesValues) => seriesValues.some((value) => value != null));
    if (!hasCandidateValues) {
      continue;
    }
    let score = 0;
    const candidateTitle = normalizeChartTitleForMatch(extractChartTitleFromXml(chartXml));
    if (preferred.length > 0 && candidateTitle.length > 0 && preferred === candidateTitle) {
      score += 100;
    }
    if (bestCandidate.length === 0) {
      score += 1;
    }
    if (score > bestScore) {
      bestScore = score;
      bestCandidate = candidateBubbleSizes;
      if (score >= 100) {
        break;
      }
    }
  }
  return bestCandidate;
}
function parseChartTypeFromXml(chartXml) {
  for (const chartType of PRIMARY_CHART_TYPE_LOCAL_NAMES) {
    if (new RegExp(`<c:${chartType}\\b`, "i").test(chartXml)) {
      return chartType;
    }
  }
  return "";
}
function findPrimaryChartTypeNode(plotAreaNode) {
  if (!plotAreaNode) {
    return null;
  }
  for (const localName of PRIMARY_CHART_TYPE_LOCAL_NAMES) {
    const node = getLocalChildren(plotAreaNode, localName)[0];
    if (node) {
      return node;
    }
  }
  return null;
}
function resolveScatterChartType(scatterStyle) {
  switch (scatterStyle) {
    case "line":
    case "lineMarker":
      return "ScatterLines";
    case "smooth":
    case "smoothMarker":
      return "ScatterSmooth";
    default:
      return "Scatter";
  }
}
function resolveArchiveFallbackPointStyles(archive, preferredTitle, preferredChartXmlType, themePalette) {
  const preferred = normalizeChartTitleForMatch(preferredTitle);
  const preferredType = (preferredChartXmlType ?? "").trim();
  let bestScore = Number.NEGATIVE_INFINITY;
  let bestCandidate = [];
  for (const [path, bytes] of Object.entries(archive)) {
    if (!/\/charts\/chart\d+\.xml$/i.test(path)) {
      continue;
    }
    const chartXml = strFromU82(bytes);
    const candidateType = parseChartTypeFromXml(chartXml);
    if (!candidateType) {
      continue;
    }
    const candidatePointStyles = parseFallbackPointStylesFromChartXml(chartXml, themePalette);
    const hasCandidateValues = candidatePointStyles.some((seriesStyles) => seriesStyles.some((style) => typeof style.color === "string" && style.color.length > 0 || typeof style.explosion === "number"));
    if (!hasCandidateValues) {
      continue;
    }
    let score = 0;
    const candidateTitle = normalizeChartTitleForMatch(extractChartTitleFromXml(chartXml));
    if (preferred.length > 0 && candidateTitle.length > 0 && preferred === candidateTitle) {
      score += 100;
    }
    if (preferredType && candidateType === preferredType) {
      score += 20;
    }
    if (bestCandidate.length === 0) {
      score += 1;
    }
    if (score > bestScore) {
      bestScore = score;
      bestCandidate = candidatePointStyles;
      if (score >= 120) {
        break;
      }
    }
  }
  return bestCandidate;
}
function parseHexColor(color) {
  const normalized = normalizeHexColor(color);
  if (!normalized) {
    return null;
  }
  const match = /^#([0-9a-f]{6})$/.exec(normalized);
  if (!match) {
    return null;
  }
  return [
    Number.parseInt(match[1].slice(0, 2), 16),
    Number.parseInt(match[1].slice(2, 4), 16),
    Number.parseInt(match[1].slice(4, 6), 16)
  ];
}
function rgbToHsl(red, green, blue) {
  const normalizedRed = red / 255;
  const normalizedGreen = green / 255;
  const normalizedBlue = blue / 255;
  const max = Math.max(normalizedRed, normalizedGreen, normalizedBlue);
  const min = Math.min(normalizedRed, normalizedGreen, normalizedBlue);
  const lightness = (max + min) / 2;
  if (max === min) {
    return [0, 0, lightness];
  }
  const delta = max - min;
  const saturation = lightness > 0.5 ? delta / (2 - max - min) : delta / (max + min);
  let hue = 0;
  switch (max) {
    case normalizedRed:
      hue = (normalizedGreen - normalizedBlue) / delta + (normalizedGreen < normalizedBlue ? 6 : 0);
      break;
    case normalizedGreen:
      hue = (normalizedBlue - normalizedRed) / delta + 2;
      break;
    default:
      hue = (normalizedRed - normalizedGreen) / delta + 4;
      break;
  }
  return [hue / 6, saturation, lightness];
}
function hueToRgb(p, q, t) {
  let nextT = t;
  if (nextT < 0) {
    nextT += 1;
  }
  if (nextT > 1) {
    nextT -= 1;
  }
  if (nextT < 1 / 6) {
    return p + (q - p) * 6 * nextT;
  }
  if (nextT < 1 / 2) {
    return q;
  }
  if (nextT < 2 / 3) {
    return p + (q - p) * (2 / 3 - nextT) * 6;
  }
  return p;
}
function hslToRgb(hue, saturation, lightness) {
  if (saturation === 0) {
    const gray = Math.round(lightness * 255);
    return [gray, gray, gray];
  }
  const q = lightness < 0.5 ? lightness * (1 + saturation) : lightness + saturation - lightness * saturation;
  const p = 2 * lightness - q;
  return [
    Math.round(hueToRgb(p, q, hue + 1 / 3) * 255),
    Math.round(hueToRgb(p, q, hue) * 255),
    Math.round(hueToRgb(p, q, hue - 1 / 3) * 255)
  ];
}
function rgbToHex(red, green, blue) {
  return `#${[red, green, blue].map((channel) => Math.max(0, Math.min(255, Math.round(channel))).toString(16).padStart(2, "0")).join("")}`;
}
function applyLightnessTransform(baseColor, modifier = 1, offset = 0) {
  const rgb = parseHexColor(baseColor);
  if (!rgb) {
    return normalizeHexColor(baseColor);
  }
  const [hue, saturation, lightness] = rgbToHsl(rgb[0], rgb[1], rgb[2]);
  const nextLightness = clampUnitInterval(lightness * modifier + offset);
  const [nextRed, nextGreen, nextBlue] = hslToRgb(hue, saturation, nextLightness);
  return rgbToHex(nextRed, nextGreen, nextBlue);
}
function resolveThemeColor(name, themePalette) {
  if (!name) {
    return null;
  }
  const index = THEME_COLOR_INDEX_BY_NAME[name];
  return index === void 0 ? null : themePalette?.colorsByIndex[index] ?? null;
}
function resolveThemeTypeface(typeface, themePalette) {
  if (!typeface) {
    return null;
  }
  if (typeface === "+mn-lt" || typeface === "+mn-ea" || typeface === "+mn-cs") {
    return themePalette?.minorLatinFont ?? null;
  }
  if (typeface === "+mj-lt" || typeface === "+mj-ea" || typeface === "+mj-cs") {
    return themePalette?.majorLatinFont ?? null;
  }
  return typeface;
}
function readChartTextTypeface(textPropertiesNode, themePalette) {
  if (!textPropertiesNode) {
    return null;
  }
  const defaultRunProperties = getFirstLocalDescendant(textPropertiesNode, "defRPr") ?? getFirstLocalDescendant(textPropertiesNode, "rPr");
  if (!defaultRunProperties) {
    return null;
  }
  const typeface = getFirstLocalChild(defaultRunProperties, "latin")?.getAttribute("typeface") ?? getFirstLocalChild(defaultRunProperties, "ea")?.getAttribute("typeface") ?? getFirstLocalChild(defaultRunProperties, "cs")?.getAttribute("typeface") ?? null;
  const resolved = resolveThemeTypeface(typeface, themePalette)?.trim() ?? "";
  return resolved.length > 0 ? resolved : null;
}
function resolveChartColorNode(node, themePalette) {
  if (!node) {
    return null;
  }
  let baseColor = null;
  if (node.localName === "srgbClr") {
    baseColor = normalizeHexColor(`#${node.getAttribute("val") ?? ""}`);
  } else if (node.localName === "schemeClr") {
    baseColor = resolveThemeColor(node.getAttribute("val"), themePalette);
  } else if (node.localName === "sysClr") {
    baseColor = normalizeHexColor(`#${node.getAttribute("lastClr") ?? ""}`);
  }
  if (!baseColor) {
    return null;
  }
  let lightnessModifier = 1;
  let lightnessOffset = 0;
  for (const transformNode of Array.from(node.childNodes).filter(isElementNode)) {
    const rawValue = Number(transformNode.getAttribute("val") ?? Number.NaN);
    if (!Number.isFinite(rawValue)) {
      continue;
    }
    if (transformNode.localName === "lumMod") {
      lightnessModifier *= rawValue / 1e5;
    } else if (transformNode.localName === "lumOff") {
      lightnessOffset += rawValue / 1e5;
    } else if (transformNode.localName === "tint") {
      lightnessOffset += (1 - lightnessOffset) * (rawValue / 1e5);
    } else if (transformNode.localName === "shade") {
      lightnessModifier *= rawValue / 1e5;
    }
  }
  return applyLightnessTransform(baseColor, lightnessModifier, lightnessOffset);
}
function isChartColorElement(node) {
  return Boolean(node && (node.localName === "schemeClr" || node.localName === "srgbClr" || node.localName === "sysClr"));
}
function findFirstChartColorElement(node) {
  if (!node) {
    return null;
  }
  if (isChartColorElement(node)) {
    return node;
  }
  for (const localName of ["srgbClr", "schemeClr", "sysClr"]) {
    for (const candidate of getLocalDescendants(node, localName)) {
      if (isChartColorElement(candidate)) {
        return candidate;
      }
    }
  }
  return null;
}
function resolveChartFillColor(shapeNode, themePalette) {
  if (!shapeNode || getFirstLocalChild(shapeNode, "noFill")) {
    return null;
  }
  const solidFill = getFirstLocalChild(shapeNode, "solidFill");
  if (solidFill) {
    const colorNode = findFirstChartColorElement(Array.from(solidFill.childNodes).find(isElementNode) ?? null);
    return resolveChartColorNode(colorNode, themePalette);
  }
  const gradientFill = getFirstLocalChild(shapeNode, "gradFill");
  const gradientStops = gradientFill ? getLocalDescendants(gradientFill, "gs").map((stopNode) => ({
    colorNode: Array.from(stopNode.childNodes).find(isElementNode) ?? null,
    position: Number(stopNode.getAttribute("pos") ?? Number.NaN)
  })).filter((stop) => Boolean(stop.colorNode)) : [];
  if (gradientStops.length === 0) {
    return null;
  }
  gradientStops.sort((left, right) => {
    const leftPos = Number.isFinite(left.position) ? left.position : 0;
    const rightPos = Number.isFinite(right.position) ? right.position : 0;
    return leftPos - rightPos;
  });
  const midpointStop = gradientStops.find((stop) => Number.isFinite(stop.position) && stop.position >= 5e4) ?? gradientStops[Math.floor(gradientStops.length / 2)] ?? gradientStops[0];
  return resolveChartColorNode(midpointStop.colorNode, themePalette);
}
function resolveChartLineStyle(shapeNode, themePalette) {
  const lineNode = shapeNode?.localName === "ln" ? shapeNode : shapeNode ? getFirstLocalChild(shapeNode, "ln") : null;
  if (!lineNode) {
    return { color: null, hidden: false, widthPx: void 0 };
  }
  if (getFirstLocalChild(lineNode, "noFill")) {
    return { color: null, hidden: true, widthPx: void 0 };
  }
  const solidFill = getFirstLocalChild(lineNode, "solidFill");
  const colorNode = solidFill ? findFirstChartColorElement(Array.from(solidFill.childNodes).find(isElementNode) ?? null) : null;
  const widthValue = Number(lineNode.getAttribute("w") ?? Number.NaN);
  return {
    color: resolveChartColorNode(colorNode, themePalette),
    hidden: false,
    widthPx: Number.isFinite(widthValue) ? Math.max(1, widthValue / EMU_PER_PIXEL) : void 0
  };
}
function normalizeLegend(raw) {
  if (!raw || typeof raw !== "object") {
    return null;
  }
  const legend = raw;
  return {
    overlay: typeof legend.overlay === "boolean" ? legend.overlay : void 0,
    position: typeof legend.position === "string" ? legend.position : void 0,
    raw: legend
  };
}
function normalizeLegendPosition(position) {
  if (!position) {
    return void 0;
  }
  switch (position) {
    case "bottom":
      return "b";
    case "left":
      return "l";
    case "right":
      return "r";
    case "top":
      return "t";
    default:
      return position;
  }
}
function readChartNumericAttribute(parent, localName) {
  const node = parent ? getFirstLocalChild(parent, localName) : null;
  const value = Number(node?.getAttribute("val") ?? Number.NaN);
  return Number.isFinite(value) ? value : void 0;
}
function readChartBooleanAttribute(parent, localName) {
  const node = parent ? getFirstLocalChild(parent, localName) : null;
  if (!node) {
    return void 0;
  }
  const rawValue = node.getAttribute("val");
  if (rawValue == null) {
    return true;
  }
  if (rawValue === "1" || rawValue === "true") {
    return true;
  }
  if (rawValue === "0" || rawValue === "false") {
    return false;
  }
  return void 0;
}
function readChartLabelFontSizePt(textPropertiesNode) {
  if (!textPropertiesNode) {
    return void 0;
  }
  const runPropertiesNode = getFirstLocalDescendant(textPropertiesNode, "defRPr") ?? getFirstLocalDescendant(textPropertiesNode, "rPr");
  const rawSize = Number(runPropertiesNode?.getAttribute("sz") ?? Number.NaN);
  if (!Number.isFinite(rawSize) || rawSize <= 0) {
    return void 0;
  }
  return rawSize / 100;
}
function parseChartPointDataLabelsFromXml(labelsNode) {
  const fallbackFontSizePt = readChartLabelFontSizePt(getFirstLocalChild(labelsNode, "txPr"));
  const labels = [];
  for (const pointLabelNode of getLocalChildren(labelsNode, "dLbl")) {
    const index = readChartNumericAttribute(pointLabelNode, "idx");
    if (typeof index !== "number" || !Number.isFinite(index)) {
      continue;
    }
    const layoutNode = getFirstLocalChild(pointLabelNode, "layout");
    const manualLayoutNode = layoutNode ? getFirstLocalChild(layoutNode, "manualLayout") : null;
    labels.push({
      deleted: readChartBooleanAttribute(pointLabelNode, "delete"),
      fontSizePt: readChartLabelFontSizePt(getFirstLocalChild(pointLabelNode, "txPr")) ?? fallbackFontSizePt,
      index,
      showBubbleSize: readChartBooleanAttribute(pointLabelNode, "showBubbleSize"),
      showCategoryName: readChartBooleanAttribute(pointLabelNode, "showCatName"),
      showPercent: readChartBooleanAttribute(pointLabelNode, "showPercent"),
      showSeriesName: readChartBooleanAttribute(pointLabelNode, "showSerName"),
      showValue: readChartBooleanAttribute(pointLabelNode, "showVal"),
      x: readChartNumericAttribute(manualLayoutNode, "x"),
      y: readChartNumericAttribute(manualLayoutNode, "y")
    });
  }
  return labels;
}
function parseChartDataLabelsFromXml(labelsNode) {
  if (!labelsNode) {
    return null;
  }
  const pointLabels = parseChartPointDataLabelsFromXml(labelsNode);
  const labels = {
    pointLabels: pointLabels.length > 0 ? pointLabels : void 0,
    raw: {},
    showBubbleSize: readChartBooleanAttribute(labelsNode, "showBubbleSize"),
    showCategoryName: readChartBooleanAttribute(labelsNode, "showCatName"),
    showLegendKey: readChartBooleanAttribute(labelsNode, "showLegendKey"),
    showPercent: readChartBooleanAttribute(labelsNode, "showPercent"),
    showSeriesName: readChartBooleanAttribute(labelsNode, "showSerName"),
    showValue: readChartBooleanAttribute(labelsNode, "showVal")
  };
  const hasValue = labels.showBubbleSize !== void 0 || labels.showCategoryName !== void 0 || labels.showLegendKey !== void 0 || labels.showPercent !== void 0 || (labels.pointLabels?.length ?? 0) > 0 || labels.showSeriesName !== void 0 || labels.showValue !== void 0;
  return hasValue ? labels : null;
}
function readChartRelationships(archive, chartPath) {
  const relsPath = normalizeArchivePath(`${dirname(chartPath)}/_rels/${chartPath.split("/").pop()}.rels`);
  const relsXml = readArchiveText(archive, relsPath);
  if (!relsXml) {
    return /* @__PURE__ */ new Map();
  }
  const relsDocument = parseXml(relsXml);
  if (!relsDocument) {
    return /* @__PURE__ */ new Map();
  }
  const relationships = /* @__PURE__ */ new Map();
  for (const relationshipNode of getLocalDescendants(relsDocument, "Relationship")) {
    const type = relationshipNode.getAttribute("Type");
    const target = relationshipNode.getAttribute("Target");
    if (!type || !target) {
      continue;
    }
    relationships.set(type, resolveRelationshipPath(relsPath, target));
  }
  return relationships;
}
function readChartColorPalette(archive, colorStylePath, themePalette) {
  const colorStyleXml = readArchiveText(archive, colorStylePath);
  if (!colorStyleXml) {
    return [];
  }
  const colorStyleDocument = parseXml(colorStyleXml);
  if (!colorStyleDocument?.documentElement) {
    return [];
  }
  return Array.from(colorStyleDocument.documentElement.childNodes).filter((child) => isElementNode(child) && child.localName !== "variation").map((child) => resolveChartColorNode(child, themePalette) ?? resolveChartColorNode(findFirstChartColorElement(child), themePalette)).filter((color) => typeof color === "string");
}
function readChartStyleAppearance(archive, stylePath, themePalette) {
  const styleXml = readArchiveText(archive, stylePath);
  if (!styleXml) {
    return {};
  }
  const styleDocument = parseXml(styleXml);
  if (!styleDocument) {
    return {};
  }
  const dataPointNode = getFirstLocalDescendant(styleDocument, "dataPoint");
  const fillRefNode = dataPointNode ? getFirstLocalChild(dataPointNode, "fillRef") : null;
  const index = Number(fillRefNode?.getAttribute("idx") ?? Number.NaN);
  const chartAreaNode = getFirstLocalDescendant(styleDocument, "chartArea");
  const chartAreaShapeProperties = chartAreaNode ? getFirstLocalChild(chartAreaNode, "spPr") : null;
  const chartAreaFontRef = chartAreaNode ? getFirstLocalChild(chartAreaNode, "fontRef") : null;
  const chartAreaFontColor = chartAreaFontRef ? resolveChartColorNode(Array.from(chartAreaFontRef.childNodes).find(isElementNode) ?? null, themePalette) : null;
  const titleNode = getFirstLocalDescendant(styleDocument, "title");
  const titleFontRef = titleNode ? getFirstLocalChild(titleNode, "fontRef") : null;
  const titleColor = titleFontRef ? resolveChartColorNode(Array.from(titleFontRef.childNodes).find(isElementNode) ?? null, themePalette) : null;
  const axisStyleNode = getFirstLocalDescendant(styleDocument, "categoryAxis") ?? getFirstLocalDescendant(styleDocument, "valueAxis");
  const axisShapeProperties = axisStyleNode ? getFirstLocalChild(axisStyleNode, "spPr") : null;
  const axisFontRef = axisStyleNode ? getFirstLocalChild(axisStyleNode, "fontRef") : null;
  const chartAreaNoFill = chartAreaShapeProperties ? getFirstLocalChild(chartAreaShapeProperties, "noFill") != null : false;
  return {
    axisLabelColor: axisFontRef ? resolveChartColorNode(Array.from(axisFontRef.childNodes).find(isElementNode) ?? null, themePalette) ?? void 0 : void 0,
    axisLineColor: resolveChartLineStyle(axisShapeProperties, themePalette).color ?? void 0,
    chartAreaBorderColor: resolveChartLineStyle(chartAreaShapeProperties, themePalette).color ?? void 0,
    chartAreaFillColor: resolveChartFillColor(chartAreaShapeProperties, themePalette) ?? void 0,
    chartAreaNoFill,
    paletteOffset: Number.isFinite(index) ? index : void 0,
    textColor: chartAreaFontColor ?? void 0,
    titleColor: titleColor ?? chartAreaFontColor ?? void 0
  };
}
function buildThemeSeriesPalette(themePalette) {
  const themeColors = [4, 5, 6, 7, 8, 9].map((index) => themePalette?.colorsByIndex[index] ?? null).filter((color) => Boolean(color));
  return themeColors.length > 0 ? themeColors : SERIES_COLORS;
}
function normalizeBuiltinSurfaceStyleId(styleId) {
  if (typeof styleId !== "number" || !Number.isFinite(styleId)) {
    return null;
  }
  return styleId >= 100 ? styleId - 100 : styleId;
}
function getBuiltinSurfacePalette(styleId, wireframe) {
  const normalized = normalizeBuiltinSurfaceStyleId(styleId);
  if (normalized === 34 || wireframe === true && normalized == null) {
    return ["#5b9bd5", "#ed7d31", "#a5a5a5"];
  }
  if (normalized === 35 || normalized === 36 || wireframe !== true && normalized == null) {
    return ["#2f5597", "#4472c4", "#5b9bd5", "#8faadc", "#d9e2f3"];
  }
  return null;
}
function applyBuiltinSurfaceDefaults(chart) {
  if (chart.chartType !== "Surface") {
    return;
  }
  const builtinPalette = getBuiltinSurfacePalette(chart.chartStyleId, chart.wireframe);
  if ((!chart.chartColorPalette || chart.chartColorPalette.length === 0) && builtinPalette) {
    chart.chartColorPalette = builtinPalette;
  }
  const wallFill = chart.wireframe ? "#d0d0d0" : "#d9d9df";
  const wallLine = chart.wireframe ? "#a6a6a6" : "#a8adb7";
  chart.floor = {
    ...chart.floor ?? {},
    fillColor: chart.floor?.fillColor ?? wallFill,
    lineColor: chart.floor?.lineColor ?? wallLine
  };
  chart.sideWall = {
    ...chart.sideWall ?? {},
    fillColor: chart.sideWall?.fillColor ?? wallFill,
    lineColor: chart.sideWall?.lineColor ?? wallLine
  };
  chart.backWall = {
    ...chart.backWall ?? {},
    fillColor: chart.backWall?.fillColor ?? wallFill,
    lineColor: chart.backWall?.lineColor ?? wallLine
  };
  if (!chart.surfaceMaterial && chart.wireframe !== true) {
    chart.surfaceMaterial = "flat";
  }
}
function applyBuiltinChartDefaults(chart, themePalette) {
  const darkBuiltInStyle = typeof chart.chartStyleId === "number" && chart.chartStyleId >= 140 && chart.chartStyleId < 150;
  const textColor = themePalette?.colorsByIndex[1] ?? themePalette?.colorsByIndex[3] ?? null;
  const minorTypeface = themePalette?.minorLatinFont?.trim() || void 0;
  const derivedAxisColor = textColor ? applyLightnessTransform(textColor, 0.35, 0.55) : null;
  const derivedBorderColor = textColor ? applyLightnessTransform(textColor, chart.is3d ? 0.28 : 0.22, chart.is3d ? 0.6 : 0.7) : null;
  if (darkBuiltInStyle) {
    chart.chartAreaFillColor = chart.chartAreaFillColor ?? "#1f1f1f";
    chart.chartAreaBorderColor = chart.chartAreaBorderColor ?? "#1f1f1f";
    chart.textColor = chart.textColor ?? "#f5f5f5";
    chart.titleColor = chart.titleColor ?? "#f5f5f5";
    chart.axisLabelColor = chart.axisLabelColor ?? "#d9d9d9";
    chart.axisLineColor = chart.axisLineColor ?? "#8c8c8c";
  }
  chart.chartAreaBorderColor = chart.chartAreaBorderColor ?? derivedBorderColor ?? void 0;
  chart.textColor = chart.textColor ?? textColor ?? void 0;
  chart.titleColor = chart.titleColor ?? textColor ?? void 0;
  chart.axisLabelColor = chart.axisLabelColor ?? derivedAxisColor ?? textColor ?? void 0;
  chart.axisLineColor = chart.axisLineColor ?? derivedAxisColor ?? textColor ?? void 0;
  chart.fontFamily = chart.fontFamily ?? minorTypeface;
  chart.titleFontFamily = chart.titleFontFamily ?? chart.fontFamily ?? minorTypeface;
  const seriesPalette = chart.chartColorPalette && chart.chartColorPalette.length > 0 ? chart.chartColorPalette : buildThemeSeriesPalette(themePalette);
  if (!chart.chartColorPalette || chart.chartColorPalette.length === 0) {
    chart.chartColorPalette = seriesPalette;
  }
  chart.series = chart.series.map((series, index) => {
    const fallbackColor = seriesPalette[index % seriesPalette.length];
    return {
      ...series,
      color: series.color ?? series.lineColor ?? fallbackColor,
      lineColor: series.lineColor ?? series.color ?? fallbackColor,
      markerColor: series.markerColor ?? series.color ?? series.lineColor ?? fallbackColor,
      markerLineColor: series.markerLineColor ?? series.lineColor ?? series.color ?? fallbackColor
    };
  });
  chart.typeGroups = chart.typeGroups?.map((group, groupIndex) => ({
    ...group,
    series: group.series.map((series, seriesIndex) => {
      const fallbackColor = seriesPalette[(groupIndex + seriesIndex) % seriesPalette.length];
      return {
        ...series,
        color: series.color ?? series.lineColor ?? fallbackColor,
        lineColor: series.lineColor ?? series.color ?? fallbackColor,
        markerColor: series.markerColor ?? series.color ?? series.lineColor ?? fallbackColor,
        markerLineColor: series.markerLineColor ?? series.lineColor ?? series.color ?? fallbackColor
      };
    })
  }));
  applyBuiltinSurfaceDefaults(chart);
}
function parseChartPointStyles(seriesNode, themePalette) {
  const pointStyles = [];
  for (const dataPointNode of getLocalChildren(seriesNode, "dPt")) {
    const indexValue = readChartNumericAttribute(dataPointNode, "idx");
    if (indexValue === void 0) {
      continue;
    }
    const shapeProperties = getFirstLocalChild(dataPointNode, "spPr");
    const lineStyle = resolveChartLineStyle(shapeProperties, themePalette);
    pointStyles.push({
      color: resolveChartFillColor(shapeProperties, themePalette) ?? void 0,
      explosion: readChartNumericAttribute(dataPointNode, "explosion"),
      index: indexValue,
      lineColor: lineStyle.color ?? void 0
    });
  }
  return pointStyles;
}
function parseInvertNegativeStyle(seriesNode, themePalette) {
  const invertNode = getFirstLocalDescendant(seriesNode, "invertSolidFillFmt");
  const shapeProperties = invertNode ? getFirstLocalChild(invertNode, "spPr") : null;
  if (!shapeProperties) {
    return {
      color: void 0,
      lineColor: void 0
    };
  }
  const lineStyle = resolveChartLineStyle(shapeProperties, themePalette);
  return {
    color: resolveChartFillColor(shapeProperties, themePalette) ?? void 0,
    lineColor: lineStyle.color ?? void 0
  };
}
function parseChartCacheValues(parentNode, cacheName, mode) {
  if (!parentNode) {
    return null;
  }
  const referenceNode = getFirstLocalChild(parentNode, "numRef") ?? getFirstLocalChild(parentNode, "strRef") ?? parentNode;
  const cacheNode = getFirstLocalChild(referenceNode, cacheName);
  if (!cacheNode) {
    return null;
  }
  const pointCount = readChartNumericAttribute(cacheNode, "ptCount");
  const pointNodes = getLocalChildren(cacheNode, "pt").map((pointNode) => {
    const rawIndex = Number(pointNode.getAttribute("idx") ?? Number.NaN);
    return {
      index: Number.isFinite(rawIndex) ? rawIndex : 0,
      value: getFirstLocalChild(pointNode, "v")?.textContent ?? ""
    };
  }).sort((left, right) => left.index - right.index);
  if (pointNodes.length === 0) {
    return null;
  }
  const maxIndex = pointNodes.reduce((max, point) => Math.max(max, point.index), 0);
  const targetLength = Math.max(
    pointNodes.length,
    Number.isFinite(pointCount ?? Number.NaN) ? Number(pointCount) : 0,
    maxIndex + 1
  );
  const values = Array.from({ length: targetLength }, () => null);
  for (const point of pointNodes) {
    if (mode === "value") {
      values[point.index] = cellValueToNumber(point.value);
    } else {
      values[point.index] = point.value.length > 0 ? point.value : null;
    }
  }
  return values;
}
function parseChartMultiLevelCacheValues(parentNode, mode) {
  if (!parentNode) {
    return null;
  }
  const referenceNode = getFirstLocalChild(parentNode, "multiLvlStrRef") ?? parentNode;
  const cacheNode = getFirstLocalChild(referenceNode, "multiLvlStrCache");
  if (!cacheNode) {
    return null;
  }
  const levelNodes = getLocalChildren(cacheNode, "lvl");
  if (levelNodes.length === 0) {
    return null;
  }
  const pointCount = readChartNumericAttribute(cacheNode, "ptCount");
  const primaryLevelNode = mode === "category" ? levelNodes[levelNodes.length - 1] ?? levelNodes[0] : levelNodes[0];
  const pointNodes = getLocalChildren(primaryLevelNode, "pt").map((pointNode) => {
    const rawIndex = Number(pointNode.getAttribute("idx") ?? Number.NaN);
    return {
      index: Number.isFinite(rawIndex) ? rawIndex : 0,
      value: getFirstLocalChild(pointNode, "v")?.textContent ?? ""
    };
  }).sort((left, right) => left.index - right.index);
  if (pointNodes.length === 0) {
    return null;
  }
  const maxIndex = pointNodes.reduce((max, point) => Math.max(max, point.index), 0);
  const targetLength = Math.max(
    pointNodes.length,
    Number.isFinite(pointCount ?? Number.NaN) ? Number(pointCount) : 0,
    maxIndex + 1
  );
  const values = Array.from({ length: targetLength }, () => null);
  for (const point of pointNodes) {
    if (mode === "value") {
      values[point.index] = cellValueToNumber(point.value);
      continue;
    }
    values[point.index] = point.value.length > 0 ? point.value : null;
  }
  return values;
}
function applyChartSeriesStyleFromXml(chart, chartTypeNode, themePalette) {
  const seriesNodes = getLocalChildren(chartTypeNode, "ser");
  chart.series = chart.series.map((series, index) => {
    const seriesNode = seriesNodes[index];
    if (!seriesNode) {
      return series;
    }
    const shapeProperties = getFirstLocalChild(seriesNode, "spPr");
    const markerNode = getFirstLocalChild(seriesNode, "marker");
    const markerShapeProperties = getFirstLocalChild(markerNode ?? chartTypeNode, "spPr");
    const lineStyle = resolveChartLineStyle(shapeProperties, themePalette);
    const markerLineStyle = resolveChartLineStyle(markerShapeProperties, themePalette);
    const fillColor = resolveChartFillColor(shapeProperties, themePalette);
    const markerSize = readChartNumericAttribute(markerNode, "size");
    const markerSymbolNode = markerNode ? getFirstLocalChild(markerNode, "symbol") : null;
    const markerSymbol = markerSymbolNode?.getAttribute("val") ?? void 0;
    const pointStyles = parseChartPointStyles(seriesNode, themePalette);
    const seriesExplosion = readChartNumericAttribute(seriesNode, "explosion");
    const invertNegativeStyle = parseInvertNegativeStyle(seriesNode, themePalette);
    const invertIfNegative = readChartBooleanAttribute(seriesNode, "invertIfNegative");
    const isScatterChart = chart.chartType === "Scatter" || chart.chartType === "ScatterLines" || chart.chartType === "ScatterSmooth" || chart.chartType === "Bubble";
    const cachedCategories = isScatterChart ? parseChartCacheValues(getFirstLocalChild(seriesNode, "xVal"), "numCache", "value") ?? parseChartMultiLevelCacheValues(getFirstLocalChild(seriesNode, "xVal"), "category") : parseChartCacheValues(getFirstLocalChild(seriesNode, "cat"), "strCache", "category") ?? parseChartCacheValues(getFirstLocalChild(seriesNode, "cat"), "numCache", "category") ?? parseChartMultiLevelCacheValues(getFirstLocalChild(seriesNode, "cat"), "category");
    const cachedValues = isScatterChart ? parseChartCacheValues(getFirstLocalChild(seriesNode, "yVal"), "numCache", "value") : parseChartCacheValues(getFirstLocalChild(seriesNode, "val"), "numCache", "value");
    const cachedBubbleSizes = chart.chartType === "Bubble" ? parseChartCacheValues(getFirstLocalChild(seriesNode, "bubbleSize"), "numCache", "value") : null;
    const existingShapeProperties = series.shapeProperties && typeof series.shapeProperties === "object" ? series.shapeProperties : null;
    const rawFillColor = typeof existingShapeProperties?.solidFillHex === "string" ? normalizeHexColor(existingShapeProperties.solidFillHex) : null;
    const rawLineColor = typeof existingShapeProperties?.lineColorHex === "string" ? normalizeHexColor(existingShapeProperties.lineColorHex) : null;
    const resolvedLineColor = lineStyle.hidden ? void 0 : rawLineColor ?? lineStyle.color ?? rawFillColor ?? fillColor ?? series.lineColor ?? series.color;
    const hasCategoryReference = typeof series.categoriesRef?.formula === "string" && series.categoriesRef.formula.length > 0;
    const hasValueReference = typeof series.valuesRef?.formula === "string" && series.valuesRef.formula.length > 0;
    const hasBubbleSizeReference = typeof series.bubbleSizeRef?.formula === "string" && series.bubbleSizeRef.formula.length > 0;
    return {
      ...series,
      bubbleSizes: !hasBubbleSizeReference && cachedBubbleSizes ? cachedBubbleSizes.map((value) => typeof value === "number" && Number.isFinite(value) ? value : null) : series.bubbleSizes,
      categories: !hasCategoryReference && cachedCategories ? cachedCategories : series.categories,
      color: rawFillColor ?? rawLineColor ?? fillColor ?? lineStyle.color ?? series.color,
      dataPointStyles: pointStyles.length > 0 ? pointStyles : series.dataPointStyles,
      lineColor: resolvedLineColor,
      lineWidthPx: lineStyle.hidden ? void 0 : lineStyle.widthPx ?? series.lineWidthPx,
      markerColor: rawFillColor ?? rawLineColor ?? resolveChartFillColor(markerShapeProperties, themePalette) ?? fillColor ?? lineStyle.color ?? void 0,
      markerLineColor: rawLineColor ?? rawFillColor ?? markerLineStyle.color ?? lineStyle.color ?? fillColor ?? void 0,
      markerSize: markerSize ?? series.markerSize,
      markerSymbol,
      smooth: readChartBooleanAttribute(seriesNode, "smooth") ?? series.smooth,
      invertIfNegative: invertIfNegative ?? series.invertIfNegative,
      shapeProperties: {
        ...series.shapeProperties,
        xmlExplosion: seriesExplosion ?? void 0,
        xmlFillColor: fillColor ?? void 0,
        xmlLineHidden: lineStyle.hidden ? true : void 0,
        xmlLineColor: lineStyle.color ?? void 0,
        xmlLineWidthPx: lineStyle.widthPx ?? void 0,
        xmlNegativeFillColor: invertNegativeStyle.color ?? void 0,
        xmlNegativeLineColor: invertNegativeStyle.lineColor ?? void 0
      },
      negativeColor: invertNegativeStyle.color ?? series.negativeColor,
      negativeLineColor: invertNegativeStyle.lineColor ?? series.negativeLineColor,
      values: !hasValueReference && cachedValues ? cachedValues.map((value) => typeof value === "number" && Number.isFinite(value) ? value : null) : series.values
    };
  });
}
function applyChartStyleFromXml(chart, chartPath, archive, themePalette) {
  const chartXml = readArchiveText(archive, chartPath);
  if (!chartXml) {
    return;
  }
  const relationships = chartPath ? readChartRelationships(archive, chartPath) : /* @__PURE__ */ new Map();
  const fallbackPointStylesBySeries = parseFallbackPointStylesFromChartXml(chartXml, themePalette);
  const fallbackSeriesStyles = parseFallbackSeriesStylesFromChartXml(chartXml, themePalette);
  const fallbackBubbleSizesBySeries = parseFallbackBubbleSizesFromChartXml(chartXml);
  const applyFallbackSeriesStyles = () => {
    if (fallbackBubbleSizesBySeries.length > 0) {
      chart.series = chart.series.map((series, seriesIndex) => {
        const fallbackBubbleSizes = fallbackBubbleSizesBySeries[seriesIndex] ?? [];
        if (fallbackBubbleSizes.length === 0) {
          return series;
        }
        const currentNumericPointCount = (series.bubbleSizes ?? []).filter(
          (value) => typeof value === "number" && Number.isFinite(value)
        ).length;
        const fallbackNumericPointCount = fallbackBubbleSizes.filter(
          (value) => typeof value === "number" && Number.isFinite(value)
        ).length;
        if (currentNumericPointCount >= fallbackNumericPointCount) {
          return series;
        }
        return {
          ...series,
          bubbleSizes: fallbackBubbleSizes
        };
      });
    }
    if (fallbackPointStylesBySeries.length > 0) {
      chart.series = chart.series.map((series, seriesIndex) => {
        const fallbackPointStyles = fallbackPointStylesBySeries[seriesIndex] ?? [];
        if (fallbackPointStyles.length === 0) {
          return series;
        }
        const existingByIndex = new Map((series.dataPointStyles ?? []).map((entry) => [entry.index, entry]));
        for (const fallbackStyle of fallbackPointStyles) {
          const existing = existingByIndex.get(fallbackStyle.index);
          existingByIndex.set(fallbackStyle.index, {
            color: existing?.color ?? fallbackStyle.color,
            explosion: existing?.explosion ?? fallbackStyle.explosion,
            index: fallbackStyle.index,
            lineColor: existing?.lineColor ?? fallbackStyle.lineColor
          });
        }
        return {
          ...series,
          dataPointStyles: Array.from(existingByIndex.values()).sort((left, right) => left.index - right.index)
        };
      });
    }
    if (fallbackSeriesStyles.length > 0) {
      chart.series = chart.series.map((series, seriesIndex) => {
        const fallbackStyle = fallbackSeriesStyles[seriesIndex];
        if (!fallbackStyle) {
          return series;
        }
        const fallbackColor = fallbackStyle.color ?? fallbackStyle.lineColor;
        return {
          ...series,
          color: series.color ?? fallbackColor,
          lineColor: series.lineColor ?? fallbackStyle.lineColor ?? fallbackColor,
          markerColor: series.markerColor ?? fallbackColor ?? series.color,
          markerLineColor: series.markerLineColor ?? fallbackStyle.lineColor ?? fallbackColor ?? series.lineColor
        };
      });
    }
  };
  const applyRelationshipStyles = () => {
    chart.chartColorPalette = readChartColorPalette(archive, relationships.get(CHART_COLOR_STYLE_REL_TYPE), themePalette);
    const styleAppearance2 = readChartStyleAppearance(
      archive,
      relationships.get(CHART_STYLE_REL_TYPE),
      themePalette
    );
    chart.axisLabelColor = styleAppearance2.axisLabelColor ?? chart.axisLabelColor;
    chart.axisLineColor = styleAppearance2.axisLineColor ?? chart.axisLineColor;
    chart.chartAreaBorderColor = styleAppearance2.chartAreaBorderColor ?? chart.chartAreaBorderColor;
    chart.chartAreaFillColor = styleAppearance2.chartAreaFillColor ?? chart.chartAreaFillColor;
    chart.chartColorPaletteOffset = styleAppearance2.paletteOffset ?? chart.chartColorPaletteOffset;
    chart.textColor = styleAppearance2.textColor ?? chart.textColor;
    chart.titleColor = styleAppearance2.titleColor ?? chart.titleColor;
    return styleAppearance2;
  };
  const applyModernChartExStyles = () => {
    const modernPlotAreaNode = chartDocument?.documentElement ? getFirstLocalDescendant(chartDocument.documentElement, "plotArea") : null;
    if (!modernPlotAreaNode) {
      return;
    }
    const parseModernBinning = (seriesNode) => {
      const layoutPrNode = getFirstLocalChild(seriesNode, "layoutPr");
      const binningNode = layoutPrNode ? getFirstLocalChild(layoutPrNode, "binning") : null;
      if (!binningNode) {
        return null;
      }
      const binning = {};
      for (const attribute of Array.from(binningNode.attributes)) {
        const rawValue = attribute.value;
        const numeric = Number(rawValue);
        binning[attribute.localName || attribute.name] = Number.isFinite(numeric) && rawValue.trim() !== "" ? numeric : rawValue;
      }
      return Object.keys(binning).length > 0 ? binning : {};
    };
    const plotAreaShapeProperties2 = getFirstLocalChild(modernPlotAreaNode, "spPr");
    if (plotAreaShapeProperties2) {
      const plotAreaFillColor = resolveChartFillColor(plotAreaShapeProperties2, themePalette);
      const plotAreaLineStyle = resolveChartLineStyle(plotAreaShapeProperties2, themePalette);
      if (plotAreaFillColor) {
        chart.chartAreaFillColor = chart.chartAreaFillColor ?? plotAreaFillColor;
      }
      if (plotAreaLineStyle.color) {
        chart.chartAreaBorderColor = chart.chartAreaBorderColor ?? plotAreaLineStyle.color;
      }
    }
    const modernSeriesNodes = getLocalDescendants(modernPlotAreaNode, "series");
    if (modernSeriesNodes.length === 0) {
      return;
    }
    chart.series = chart.series.map((series, seriesIndex) => {
      const modernSeriesNode = modernSeriesNodes[seriesIndex] ?? null;
      if (!modernSeriesNode) {
        return series;
      }
      const valueColorsNode = getFirstLocalChild(modernSeriesNode, "valueColors");
      const valueColors = valueColorsNode ? Array.from(valueColorsNode.childNodes).filter((node) => node.nodeType === Node.ELEMENT_NODE).map((node) => resolveChartColorNode(findFirstChartColorElement(node) ?? node, themePalette)).filter((value) => typeof value === "string" && value.length > 0) : [];
      const nextRaw = valueColors.length > 0 ? {
        ...series.raw && typeof series.raw === "object" ? series.raw : {},
        valueColors
      } : series.raw;
      const seriesShapeProperties = getFirstLocalChild(modernSeriesNode, "spPr");
      if (!seriesShapeProperties) {
        return nextRaw === series.raw ? series : {
          ...series,
          raw: nextRaw
        };
      }
      const fillColor = resolveChartFillColor(seriesShapeProperties, themePalette);
      const lineStyle = resolveChartLineStyle(seriesShapeProperties, themePalette);
      const fallbackColor = fillColor ?? lineStyle.color ?? void 0;
      return {
        ...series,
        color: series.color ?? fallbackColor,
        lineColor: series.lineColor ?? lineStyle.color ?? fillColor ?? fallbackColor,
        lineWidthPx: series.lineWidthPx ?? (typeof lineStyle.widthPx === "number" ? lineStyle.widthPx : void 0),
        markerColor: series.markerColor ?? fallbackColor ?? series.color,
        markerLineColor: series.markerLineColor ?? lineStyle.color ?? fallbackColor ?? series.lineColor,
        raw: nextRaw
      };
    });
    const seriesLayouts = modernSeriesNodes.map((node) => node.getAttribute("layoutId") ?? node.getAttribute("layout"));
    const clusteredColumnIndex = seriesLayouts.findIndex((layout) => layout === "clusteredColumn");
    if (clusteredColumnIndex >= 0) {
      const clusteredNode = modernSeriesNodes[clusteredColumnIndex] ?? null;
      const parsedBinning = clusteredNode ? parseModernBinning(clusteredNode) : null;
      if (parsedBinning) {
        const syntheticRawSeries = {
          layoutId: "clusteredColumn",
          layoutPr: {
            binning: parsedBinning
          }
        };
        const hasParetoLine = seriesLayouts.includes("paretoLine");
        const replaceColumnSeries = (series) => series ? buildChartExHistogramSeries(series, syntheticRawSeries, hasParetoLine) : null;
        if (chart.typeGroups && chart.typeGroups.length > 0) {
          const nextTypeGroups = chart.typeGroups.map((group) => ({ ...group, series: [...group.series] }));
          const columnGroupIndex = nextTypeGroups.findIndex((group) => group.chartType === "ColumnClustered");
          if (columnGroupIndex >= 0) {
            const originalColumnSeries = nextTypeGroups[columnGroupIndex]?.series[0] ?? null;
            const binnedColumnSeries = replaceColumnSeries(originalColumnSeries);
            if (binnedColumnSeries) {
              nextTypeGroups[columnGroupIndex].series = [binnedColumnSeries];
              const lineGroupIndex = nextTypeGroups.findIndex((group) => group.chartType === "Line");
              if (lineGroupIndex >= 0 && nextTypeGroups[lineGroupIndex]?.series[0]) {
                const originalLineSeries = nextTypeGroups[lineGroupIndex].series[0];
                const recomputedLine = buildChartExParetoLineSeries(
                  binnedColumnSeries,
                  {
                    text: originalLineSeries.name,
                    ...originalLineSeries.raw && typeof originalLineSeries.raw === "object" ? originalLineSeries.raw : {}
                  },
                  0
                );
                nextTypeGroups[lineGroupIndex].series = [
                  {
                    ...originalLineSeries,
                    categories: recomputedLine.categories,
                    categoriesRef: recomputedLine.categoriesRef,
                    raw: recomputedLine.raw,
                    values: recomputedLine.values
                  }
                ];
                chart.series = [binnedColumnSeries, nextTypeGroups[lineGroupIndex].series[0]];
              } else {
                chart.series = [binnedColumnSeries];
              }
              chart.typeGroups = nextTypeGroups;
            }
          } else if (chart.series[0]) {
            const binnedSeries = replaceColumnSeries(chart.series[0]);
            if (binnedSeries) {
              chart.series = [binnedSeries];
            }
          }
        } else if (chart.series[0]) {
          const binnedSeries = replaceColumnSeries(chart.series[0]);
          if (binnedSeries) {
            chart.series = [binnedSeries];
          }
        }
      }
    }
  };
  const chartDocument = parseXml(chartXml);
  const chartNode = chartDocument ? getFirstLocalDescendant(chartDocument, "chart") : null;
  const plotAreaNode = chartNode ? getFirstLocalChild(chartNode, "plotArea") : null;
  const styleIdNode = chartDocument?.documentElement ? getFirstLocalDescendant(chartDocument.documentElement, "style") : null;
  const chartTypeNode = findPrimaryChartTypeNode(plotAreaNode);
  if (!chartNode || !chartTypeNode) {
    applyRelationshipStyles();
    const fallbackStyleId = readChartNumericAttribute(styleIdNode, "style");
    if (typeof fallbackStyleId === "number" && Number.isFinite(fallbackStyleId)) {
      chart.chartStyleId = fallbackStyleId;
    }
    applyModernChartExStyles();
    applyFallbackSeriesStyles();
    applyBuiltinChartDefaults(chart, themePalette);
    return;
  }
  const plotArea = plotAreaNode;
  if (!plotArea) {
    applyRelationshipStyles();
    applyFallbackSeriesStyles();
    applyBuiltinChartDefaults(chart, themePalette);
    return;
  }
  switch (chartTypeNode.localName) {
    case "barChart":
    case "bar3DChart": {
      const grouping = getFirstLocalChild(chartTypeNode, "grouping")?.getAttribute("val");
      const barDir = getFirstLocalChild(chartTypeNode, "barDir")?.getAttribute("val");
      const isHorizontalBar = barDir === "bar";
      chart.is3d = chartTypeNode.localName === "bar3DChart" ? true : chart.is3d;
      if (grouping === "percentStacked") {
        chart.chartType = isHorizontalBar ? "BarPercentStacked" : "ColumnPercentStacked";
      } else if (grouping === "stacked") {
        chart.chartType = isHorizontalBar ? "BarStacked" : "ColumnStacked";
      } else {
        chart.chartType = isHorizontalBar ? "BarClustered" : "ColumnClustered";
      }
      break;
    }
    case "areaChart":
    case "area3DChart": {
      const grouping = getFirstLocalChild(chartTypeNode, "grouping")?.getAttribute("val");
      chart.is3d = chartTypeNode.localName === "area3DChart" ? true : chart.is3d;
      if (grouping === "stacked") {
        chart.chartType = "AreaStacked";
      } else if (grouping === "percentStacked") {
        chart.chartType = "AreaPercentStacked";
      } else {
        chart.chartType = "Area";
      }
      break;
    }
    case "lineChart":
    case "line3DChart": {
      const grouping = getFirstLocalChild(chartTypeNode, "grouping")?.getAttribute("val");
      chart.is3d = chartTypeNode.localName === "line3DChart" ? true : chart.is3d;
      if (grouping === "stacked") {
        chart.chartType = "LineStacked";
      } else if (grouping === "percentStacked") {
        chart.chartType = "LinePercentStacked";
      } else {
        chart.chartType = "Line";
      }
      break;
    }
    case "pieChart":
      chart.chartType = "Pie";
      break;
    case "pie3DChart":
      chart.chartType = "Pie3D";
      chart.is3d = true;
      break;
    case "doughnutChart":
      chart.chartType = "Doughnut";
      break;
    case "ofPieChart":
      chart.chartType = "BarOfPie";
      break;
    case "scatterChart":
      chart.chartType = resolveScatterChartType(getFirstLocalChild(chartTypeNode, "scatterStyle")?.getAttribute("val"));
      break;
    case "radarChart":
      chart.chartType = "Radar";
      break;
    case "surfaceChart":
      chart.chartType = "Surface";
      chart.is3d = false;
      break;
    case "surface3DChart":
      chart.chartType = "Surface";
      chart.is3d = true;
      break;
    case "stockChart":
      chart.chartType = "Stock";
      break;
    case "bubbleChart":
      chart.chartType = "Bubble";
      break;
    default:
      break;
  }
  const legendNode = getFirstLocalChild(chartNode, "legend");
  const legendPosition = legendNode ? getFirstLocalChild(legendNode, "legendPos")?.getAttribute("val") ?? void 0 : void 0;
  const legendOverlay = legendNode ? getFirstLocalChild(legendNode, "overlay")?.getAttribute("val") : void 0;
  chart.legend = legendNode ? {
    overlay: legendOverlay === "1",
    position: normalizeLegendPosition(legendPosition),
    raw: chart.legend?.raw
  } : chart.legend;
  const plotVisibleOnly = readChartBooleanAttribute(chartNode, "plotVisOnly");
  if (plotVisibleOnly !== void 0) {
    chart.plotVisibleOnly = plotVisibleOnly;
  }
  chart.displayBlanksAs = getFirstLocalChild(chartNode, "dispBlanksAs")?.getAttribute("val") ?? chart.displayBlanksAs;
  const styleId = Number(styleIdNode?.getAttribute("val") ?? Number.NaN);
  chart.chartStyleId = Number.isFinite(styleId) ? styleId : chart.chartStyleId;
  chart.firstSliceAngle = readChartNumericAttribute(chartTypeNode, "firstSliceAng") ?? chart.firstSliceAngle;
  chart.gapWidth = readChartNumericAttribute(chartTypeNode, "gapWidth") ?? chart.gapWidth;
  chart.overlap = readChartNumericAttribute(chartTypeNode, "overlap") ?? chart.overlap;
  chart.bubbleScale = readChartNumericAttribute(chartTypeNode, "bubbleScale") ?? chart.bubbleScale;
  chart.varyColors = readChartBooleanAttribute(chartTypeNode, "varyColors") ?? chart.varyColors;
  const bubble3dNode = getFirstLocalChild(chartTypeNode, "bubble3D");
  chart.bubble3d = bubble3dNode ? bubble3dNode.getAttribute("val") !== "0" : chart.bubble3d;
  chart.holeSize = readChartNumericAttribute(chartTypeNode, "holeSize") ?? chart.holeSize;
  chart.radarStyle = getFirstLocalChild(chartTypeNode, "radarStyle")?.getAttribute("val") ?? chart.radarStyle;
  chart.scatterStyle = getFirstLocalChild(chartTypeNode, "scatterStyle")?.getAttribute("val") ?? chart.scatterStyle;
  chart.shape3d = getFirstLocalChild(chartTypeNode, "shape")?.getAttribute("val") ?? chart.shape3d;
  const wireframeNode = getFirstLocalChild(chartTypeNode, "wireframe");
  chart.wireframe = wireframeNode ? wireframeNode.getAttribute("val") !== "0" : chart.wireframe;
  const chartTypeDataLabels = parseChartDataLabelsFromXml(getFirstLocalChild(chartTypeNode, "dLbls"));
  const firstSeriesNode = getLocalChildren(chartTypeNode, "ser")[0] ?? null;
  const seriesDataLabels = parseChartDataLabelsFromXml(getFirstLocalChild(firstSeriesNode, "dLbls"));
  chart.dataLabels = chartTypeDataLabels ?? seriesDataLabels ?? chart.dataLabels;
  const seriesSp3dNode = firstSeriesNode ? getFirstLocalDescendant(firstSeriesNode, "sp3d") : null;
  chart.surfaceMaterial = seriesSp3dNode?.getAttribute("prstMaterial") ?? chart.surfaceMaterial;
  const bandFormatsNode = getLocalChildren(chartTypeNode, "bandFmts")[0] ?? null;
  const bandFormatNodes = bandFormatsNode ? getLocalChildren(bandFormatsNode, "bandFmt") : [];
  const bandFormatColors = bandFormatNodes.map((bandFormatNode) => {
    const shapeProperties = getFirstLocalChild(bandFormatNode, "spPr");
    return resolveChartFillColor(shapeProperties, themePalette) ?? void 0;
  }).filter((color) => typeof color === "string" && color.length > 0);
  const bandFormatLineColors = bandFormatNodes.map((bandFormatNode) => {
    const shapeProperties = getFirstLocalChild(bandFormatNode, "spPr");
    return resolveChartLineStyle(shapeProperties, themePalette).color ?? void 0;
  }).filter((color) => typeof color === "string" && color.length > 0);
  chart.raw = {
    ...chart.raw ?? {},
    bandFormatCount: bandFormatNodes.length > 0 ? bandFormatNodes.length : void 0,
    bandFormatColors: bandFormatColors.length > 0 ? bandFormatColors : void 0,
    bandFormatLineColors: bandFormatLineColors.length > 0 ? bandFormatLineColors : void 0,
    date1904: readChartBooleanAttribute(chartDocument?.documentElement ?? null, "date1904"),
    bubble3d: chart.bubble3d,
    grouping: getFirstLocalChild(chartTypeNode, "grouping")?.getAttribute("val") ?? void 0,
    ofPieType: getFirstLocalChild(chartTypeNode, "ofPieType")?.getAttribute("val") ?? void 0,
    shape: getFirstLocalChild(chartTypeNode, "shape")?.getAttribute("val") ?? void 0,
    secondPieSize: readChartNumericAttribute(chartTypeNode, "secondPieSize"),
    scatterStyle: chart.scatterStyle,
    splitPos: readChartNumericAttribute(chartTypeNode, "splitPos"),
    splitType: getFirstLocalChild(chartTypeNode, "splitType")?.getAttribute("val") ?? void 0,
    xmlChartType: chartTypeNode.localName
  };
  const view3dNode = getFirstLocalDescendant(chartNode, "view3D");
  if (view3dNode) {
    chart.view3d = {
      depthPercent: readChartNumericAttribute(view3dNode, "depthPercent"),
      perspective: readChartNumericAttribute(view3dNode, "perspective"),
      rAngAx: getFirstLocalChild(view3dNode, "rAngAx")?.getAttribute("val") === "1",
      rotX: readChartNumericAttribute(view3dNode, "rotX"),
      rotY: readChartNumericAttribute(view3dNode, "rotY")
    };
  }
  chart.floor = readChartWallFromXml(getFirstLocalChild(chartNode, "floor"), themePalette) ?? chart.floor;
  chart.sideWall = readChartWallFromXml(getFirstLocalChild(chartNode, "sideWall"), themePalette) ?? chart.sideWall;
  chart.backWall = readChartWallFromXml(getFirstLocalChild(chartNode, "backWall"), themePalette) ?? chart.backWall;
  const styleAppearance = applyRelationshipStyles();
  const chartTextTypeface = readChartTextTypeface(getFirstLocalChild(chartNode, "txPr"), themePalette);
  const titleTypeface = readChartTextTypeface(getFirstLocalDescendant(chartNode, "title"), themePalette);
  chart.fontFamily = chartTextTypeface ?? chart.fontFamily;
  chart.titleFontFamily = titleTypeface ?? chart.titleFontFamily ?? chart.fontFamily;
  const chartAreaShapeProperties = chartDocument?.documentElement ? getFirstLocalChild(chartDocument.documentElement, "spPr") : null;
  const plotAreaShapeProperties = getFirstLocalChild(plotArea, "spPr");
  const chartAreaNoFill = chartAreaShapeProperties ? getFirstLocalChild(chartAreaShapeProperties, "noFill") != null : false;
  const plotAreaNoFill = plotAreaShapeProperties ? getFirstLocalChild(plotAreaShapeProperties, "noFill") != null : false;
  chart.raw = {
    ...chart.raw ?? {},
    chartAreaNoFill: styleAppearance.chartAreaNoFill === true || chartAreaNoFill,
    plotAreaNoFill
  };
  if (chartAreaShapeProperties) {
    const chartAreaFillColor = resolveChartFillColor(chartAreaShapeProperties, themePalette);
    if (chartAreaFillColor) {
      chart.chartAreaFillColor = chartAreaFillColor;
    } else if (getFirstLocalChild(chartAreaShapeProperties, "noFill")) {
      chart.chartAreaFillColor = "transparent";
    }
    const chartAreaLineStyle = resolveChartLineStyle(chartAreaShapeProperties, themePalette);
    if (chartAreaLineStyle.hidden) {
      chart.chartAreaBorderColor = "transparent";
    } else if (chartAreaLineStyle.color) {
      chart.chartAreaBorderColor = chartAreaLineStyle.color;
    }
  }
  if (!chart.chartAreaFillColor && (styleAppearance.chartAreaNoFill === true || plotAreaNoFill)) {
    chart.chartAreaFillColor = "transparent";
  }
  const categoryAxisNodes = [
    ...getLocalChildren(plotArea, "catAx"),
    ...getLocalChildren(plotArea, "dateAx")
  ];
  const valueAxisNodes = getLocalChildren(plotArea, "valAx");
  const seriesAxisNode = getLocalChildren(plotArea, "serAx")[0] ?? null;
  const isScatterLikeChart = chart.chartType === "Scatter" || chart.chartType === "ScatterLines" || chart.chartType === "ScatterSmooth" || chart.chartType === "Bubble";
  let categoryAxisNode = categoryAxisNodes[0] ?? null;
  let valueAxisNode = valueAxisNodes[0] ?? null;
  if (!categoryAxisNode && isScatterLikeChart && valueAxisNodes.length >= 2) {
    categoryAxisNode = valueAxisNodes.find((axisNode) => {
      const position = getFirstLocalChild(axisNode, "axPos")?.getAttribute("val");
      return position === "b" || position === "t";
    }) ?? valueAxisNodes[0];
    valueAxisNode = valueAxisNodes.find((axisNode) => {
      const position = getFirstLocalChild(axisNode, "axPos")?.getAttribute("val");
      return position === "l" || position === "r";
    }) ?? valueAxisNodes[1] ?? valueAxisNodes[0];
  }
  chart.categoryAxis = mergeChartAxis(chart.categoryAxis, readChartAxisFromXml(categoryAxisNode));
  chart.valueAxis = mergeChartAxis(chart.valueAxis, readChartAxisFromXml(valueAxisNode));
  chart.seriesAxis = mergeChartAxis(chart.seriesAxis, readChartAxisFromXml(seriesAxisNode));
  chart.axes = chart.axes.length > 0 ? chart.axes.map((axis, index) => index === 0 && categoryAxisNode ? { ...axis, ...readChartAxisFromXml(categoryAxisNode) } : index === 1 && valueAxisNode ? { ...axis, ...readChartAxisFromXml(valueAxisNode) } : axis) : chart.axes;
  if (seriesAxisNode) {
    const seriesAxis = readChartAxisFromXml(seriesAxisNode);
    if (seriesAxis && !chart.axes.some((axis) => axis.id != null && axis.id === seriesAxis.id)) {
      chart.axes = [...chart.axes, seriesAxis];
    }
  }
  applyChartSeriesStyleFromXml(chart, chartTypeNode, themePalette);
  applyFallbackSeriesStyles();
  if (chart.chartType === "Bubble") {
    const archiveFallbackBubbleSizes = resolveArchiveFallbackBubbleSizes(archive, chart.title);
    if (archiveFallbackBubbleSizes.length > 0) {
      chart.series = chart.series.map((series, seriesIndex) => {
        const pointCount = Math.max(series.values.length, series.categories.length);
        if (pointCount <= 1) {
          return series;
        }
        const numericBubbleCount = (series.bubbleSizes ?? []).filter(
          (value) => typeof value === "number" && Number.isFinite(value)
        ).length;
        if (numericBubbleCount >= pointCount) {
          return series;
        }
        const fallbackCandidate = archiveFallbackBubbleSizes[seriesIndex] ?? archiveFallbackBubbleSizes[0] ?? [];
        const fallbackNumericCount = fallbackCandidate.filter(
          (value) => typeof value === "number" && Number.isFinite(value)
        ).length;
        if (fallbackNumericCount < pointCount) {
          return series;
        }
        return {
          ...series,
          bubbleSizes: fallbackCandidate
        };
      });
    }
  }
  if (chart.chartType === "Pie" || chart.chartType === "Pie3D" || chart.chartType === "PieExploded" || chart.chartType === "Doughnut" || chart.chartType === "BarOfPie") {
    const needsPointColorFallback = chart.series.some((series) => {
      const pointCount = Math.max(series.values.length, series.categories.length);
      if (pointCount <= 0) {
        return false;
      }
      const coloredPointCount = (series.dataPointStyles ?? []).filter(
        (style) => typeof style.color === "string" && style.color.length > 0
      ).length;
      return coloredPointCount === 0;
    });
    if (needsPointColorFallback) {
      const archiveFallbackPointStyles = resolveArchiveFallbackPointStyles(
        archive,
        chart.title,
        chartTypeNode.localName,
        themePalette
      );
      if (archiveFallbackPointStyles.length > 0) {
        chart.series = chart.series.map((series, seriesIndex) => {
          const fallbackStyles = archiveFallbackPointStyles[seriesIndex] ?? archiveFallbackPointStyles[0] ?? [];
          if (fallbackStyles.length === 0) {
            return series;
          }
          const existingByIndex = new Map((series.dataPointStyles ?? []).map((entry) => [entry.index, entry]));
          for (const fallbackStyle of fallbackStyles) {
            const existing = existingByIndex.get(fallbackStyle.index);
            existingByIndex.set(fallbackStyle.index, {
              color: existing?.color ?? fallbackStyle.color,
              explosion: existing?.explosion ?? fallbackStyle.explosion,
              index: fallbackStyle.index,
              lineColor: existing?.lineColor ?? fallbackStyle.lineColor
            });
          }
          return {
            ...series,
            dataPointStyles: Array.from(existingByIndex.values()).sort((left, right) => left.index - right.index)
          };
        });
      }
    }
  }
  applyBuiltinChartDefaults(chart, themePalette);
}
function normalizeArchivePath(path) {
  return path.replace(/^\/+/, "").replace(/\\/g, "/");
}
function dirname(path) {
  const normalized = normalizeArchivePath(path);
  const index = normalized.lastIndexOf("/");
  return index >= 0 ? normalized.slice(0, index) : "";
}
function resolveRelationshipPath(basePath, target) {
  if (!target) {
    return "";
  }
  const normalizedTarget = target.replace(/\\/g, "/");
  if (normalizedTarget.startsWith("/")) {
    return normalizeArchivePath(normalizedTarget);
  }
  const normalizedBasePath = normalizeArchivePath(basePath);
  let baseDirectory = dirname(normalizedBasePath);
  if (normalizedBasePath.endsWith(".rels")) {
    const relsMarker = "/_rels/";
    const relsMarkerIndex = normalizedBasePath.lastIndexOf(relsMarker);
    if (relsMarkerIndex >= 0) {
      const ownerPrefix = normalizedBasePath.slice(0, relsMarkerIndex);
      const relFileName = normalizedBasePath.slice(relsMarkerIndex + relsMarker.length);
      const ownerFileName = relFileName.endsWith(".rels") ? relFileName.slice(0, -".rels".length) : relFileName;
      baseDirectory = dirname(`${ownerPrefix}/${ownerFileName}`);
    }
  }
  const segments = [...baseDirectory.split("/").filter(Boolean), ...normalizedTarget.split("/").filter(Boolean)];
  const resolved = [];
  for (const segment of segments) {
    if (segment === ".") {
      continue;
    }
    if (segment === "..") {
      resolved.pop();
      continue;
    }
    resolved.push(segment);
  }
  return resolved.join("/");
}
function readArchiveText(archive, path) {
  if (!path) {
    return null;
  }
  const entry = archive[normalizeArchivePath(path)];
  return entry ? strFromU82(entry) : null;
}
function parseXml(xml) {
  if (typeof DOMParser === "undefined") {
    return null;
  }
  try {
    return new DOMParser().parseFromString(xml, "application/xml");
  } catch {
    return null;
  }
}
function getLocalChildren(parent, localName) {
  return Array.from(parent.childNodes).filter(
    (node) => node.nodeType === Node.ELEMENT_NODE && node.localName === localName
  );
}
function getLocalDescendants(parent, localName) {
  return Array.from(parent.getElementsByTagName("*")).filter(
    (node) => node.localName === localName
  );
}
function getFirstLocalChild(parent, localName) {
  return getLocalChildren(parent, localName)[0] ?? null;
}
function getFirstLocalDescendant(parent, localName) {
  return getLocalDescendants(parent, localName)[0] ?? null;
}
function unquoteSheetName(value) {
  const trimmed = value.trim();
  if (trimmed.startsWith("'") && trimmed.endsWith("'")) {
    return trimmed.slice(1, -1).replace(/''/g, "'");
  }
  return trimmed;
}
function splitSheetReference(reference) {
  let bangIndex = -1;
  let quoted = false;
  for (let index = 0; index < reference.length; index += 1) {
    const char = reference[index];
    if (char === "'") {
      quoted = !quoted;
    } else if (char === "!" && !quoted) {
      bangIndex = index;
      break;
    }
  }
  if (bangIndex < 0) {
    return null;
  }
  return {
    range: reference.slice(bangIndex + 1),
    sheetName: unquoteSheetName(reference.slice(0, bangIndex))
  };
}
function parseA1Cell(reference) {
  const match = /^\$?([A-Z]+)\$?(\d+)$/i.exec(reference.trim());
  if (!match) {
    return null;
  }
  let col = 0;
  for (const char of match[1].toUpperCase()) {
    col = col * 26 + (char.charCodeAt(0) - 64);
  }
  return {
    col: col - 1,
    row: Number(match[2]) - 1
  };
}
function parseA1Range(reference) {
  const [startRef, endRef = startRef] = reference.split(":");
  const start = parseA1Cell(startRef ?? "");
  const end = parseA1Cell(endRef ?? "");
  if (!start || !end) {
    return null;
  }
  return {
    end: {
      col: Math.max(start.col, end.col),
      row: Math.max(start.row, end.row)
    },
    start: {
      col: Math.min(start.col, end.col),
      row: Math.min(start.row, end.row)
    }
  };
}
function formatA1Column(col) {
  let current = col + 1;
  let label = "";
  while (current > 0) {
    const remainder = (current - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    current = Math.floor((current - 1) / 26);
  }
  return label;
}
function buildA1RangeFormula(sheetName, start, end) {
  const escapedSheetName = sheetName.replace(/'/g, "''");
  return `'${escapedSheetName}'!$${formatA1Column(start.col)}$${start.row + 1}:$${formatA1Column(end.col)}$${end.row + 1}`;
}
function resolveReferenceSheet(workbook2, fallbackSheetIndex, formula) {
  if (!formula) {
    return {
      range: null,
      sheet: workbook2.getSheet(fallbackSheetIndex),
      sheetName: workbook2.getSheet(fallbackSheetIndex)?.name ?? ""
    };
  }
  const trimmedFormula = formula.trim();
  if (trimmedFormula.length > 0 && !trimmedFormula.includes("!")) {
    try {
      const namedRange = workbook2.getNamedRange(trimmedFormula);
      if (typeof namedRange === "string" && namedRange.length > 0 && namedRange !== trimmedFormula) {
        return resolveReferenceSheet(workbook2, fallbackSheetIndex, namedRange);
      }
    } catch {
    }
  }
  const split = splitSheetReference(trimmedFormula);
  if (!split) {
    return {
      range: parseA1Range(trimmedFormula),
      sheet: workbook2.getSheet(fallbackSheetIndex),
      sheetName: workbook2.getSheet(fallbackSheetIndex)?.name ?? ""
    };
  }
  try {
    return {
      range: parseA1Range(split.range),
      sheet: workbook2.getSheetByName(split.sheetName),
      sheetName: split.sheetName
    };
  } catch {
    return {
      range: parseA1Range(split.range),
      sheet: workbook2.getSheet(fallbackSheetIndex),
      sheetName: workbook2.getSheet(fallbackSheetIndex)?.name ?? ""
    };
  }
}
function resolveChartReferenceLabel(workbook2, fallbackSheetIndex, reference, fallbackLabel) {
  if (!reference?.formula) {
    return fallbackLabel;
  }
  const resolved = resolveReferenceSheet(workbook2, fallbackSheetIndex, reference.formula);
  if (!resolved.sheet || !resolved.range) {
    return fallbackLabel;
  }
  const { start } = resolved.range;
  if (start.row > 0) {
    const headerDisplay = cellValueToDisplay(
      typeof resolved.sheet.getFormattedValueAt === "function" ? resolved.sheet.getFormattedValueAt(start.row - 1, start.col) : null
    );
    if (headerDisplay.length > 0) {
      return headerDisplay;
    }
  }
  const firstDisplay = cellValueToDisplay(
    typeof resolved.sheet.getFormattedValueAt === "function" ? resolved.sheet.getFormattedValueAt(start.row, start.col) : null
  );
  return firstDisplay.length > 0 ? firstDisplay : fallbackLabel;
}
function resolveReferenceRowPaths(workbook2, fallbackSheetIndex, reference) {
  if (!reference?.formula) {
    return [];
  }
  const resolved = resolveReferenceSheet(workbook2, fallbackSheetIndex, reference.formula);
  if (!resolved.sheet || !resolved.range) {
    return [];
  }
  const rows = [];
  for (let row = resolved.range.start.row; row <= resolved.range.end.row; row += 1) {
    const parts = [];
    for (let col = resolved.range.start.col; col <= resolved.range.end.col; col += 1) {
      const calculated = typeof resolved.sheet.getCalculatedValueAt === "function" ? resolved.sheet.getCalculatedValueAt(row, col) : null;
      const formatted = typeof resolved.sheet.getFormattedValueAt === "function" ? resolved.sheet.getFormattedValueAt(row, col) : calculated;
      const display = cellValueToDisplay(formatted ?? calculated);
      const numeric = cellValueToNumber(calculated ?? formatted);
      const label = display.length > 0 ? display : numeric != null ? String(numeric) : "";
      if (label.length > 0) {
        parts.push(label);
      }
    }
    rows.push(parts);
  }
  return rows;
}
function normalizeChartExLegend(raw) {
  if (!raw || typeof raw !== "object") {
    return null;
  }
  const legend = raw;
  const position = typeof legend.position === "string" ? normalizeLegendPosition(String(legend.position)) : void 0;
  return {
    overlay: typeof legend.overlay === "boolean" ? legend.overlay : void 0,
    position,
    raw: legend
  };
}
function humanizeChartExLayoutLabel(layout) {
  if (!layout) {
    return void 0;
  }
  return layout.replace(/([a-z0-9])([A-Z])/g, "$1 $2").replace(/[-_]+/g, " ").trim().replace(/\b\w/g, (match) => match.toUpperCase());
}
function normalizeChartExAxis(raw) {
  if (!raw || typeof raw !== "object") {
    return null;
  }
  const axis = raw;
  const scaling = axis.scaling && typeof axis.scaling === "object" ? axis.scaling : null;
  const numberFormat = axis.numberFormat && typeof axis.numberFormat === "object" ? axis.numberFormat : null;
  return {
    delete: typeof axis.hidden === "boolean" ? axis.hidden : void 0,
    id: typeof axis.id === "number" && Number.isFinite(axis.id) ? axis.id : void 0,
    crossId: typeof axis.crossId === "number" && Number.isFinite(axis.crossId) ? axis.crossId : void 0,
    majorGridlines: axis.majorGridlines != null ? true : void 0,
    majorUnit: typeof scaling?.majorUnit === "number" ? scaling.majorUnit : void 0,
    max: typeof scaling?.max === "number" ? scaling.max : void 0,
    min: typeof scaling?.min === "number" ? scaling.min : void 0,
    minorGridlines: axis.minorGridlines != null ? true : void 0,
    minorUnit: typeof scaling?.minorUnit === "number" ? scaling.minorUnit : void 0,
    numberFormat: numberFormat ? {
      formatCode: typeof numberFormat.formatCode === "string" ? numberFormat.formatCode : void 0,
      sourceLinked: typeof numberFormat.sourceLinked === "boolean" ? numberFormat.sourceLinked : void 0
    } : void 0,
    raw: axis,
    position: typeof axis.position === "string" ? axis.position : void 0,
    tickLabelSkip: typeof axis.tickLabelSkip === "number" ? axis.tickLabelSkip : void 0,
    tickMarkSkip: typeof axis.tickMarkSkip === "number" ? axis.tickMarkSkip : void 0
  };
}
function resolveChartExLayoutChartType(layout) {
  switch (layout) {
    case "boxWhisker":
      return "BoxWhisker";
    case "clusteredColumn":
      return "ColumnClustered";
    case "funnel":
      return "Funnel";
    case "paretoLine":
      return "Line";
    case "regionMap":
      return "RegionMap";
    case "sunburst":
      return "Sunburst";
    case "treemap":
      return "Treemap";
    case "waterfall":
      return "Waterfall";
    default:
      return layout ? `Unsupported(cx:${layout})` : "ColumnClustered";
  }
}
function resolveChartExSeriesLayout(raw) {
  if (!raw || typeof raw !== "object") {
    return void 0;
  }
  const record = raw;
  return typeof record.layout === "string" ? record.layout : typeof record.layoutId === "string" ? record.layoutId : void 0;
}
function resolveChartExSeriesAxisIds(raw) {
  if (!raw || typeof raw !== "object") {
    return [];
  }
  const record = raw;
  if (Array.isArray(record.axisIds)) {
    return record.axisIds.filter((value) => typeof value === "number" && Number.isFinite(value));
  }
  if (Array.isArray(record.axisId)) {
    return record.axisId.flatMap((value) => {
      if (typeof value === "number" && Number.isFinite(value)) {
        return [value];
      }
      if (value && typeof value === "object" && typeof value.val === "number") {
        return [value.val];
      }
      return [];
    });
  }
  if (typeof record.axisId === "number" && Number.isFinite(record.axisId)) {
    return [record.axisId];
  }
  return [];
}
function niceHistogramStep(value) {
  if (!Number.isFinite(value) || value <= 0) {
    return 1;
  }
  const exponent = Math.floor(Math.log10(value));
  const scale = 10 ** exponent;
  const normalized = value / scale;
  if (normalized <= 1) {
    return scale;
  }
  if (normalized <= 2) {
    return scale * 2;
  }
  if (normalized <= 5) {
    return scale * 5;
  }
  return scale * 10;
}
function formatHistogramBinLabel(lower, upper, index, closedRight) {
  const leftBracket = closedRight ? index === 0 ? "[" : "(" : "[";
  const rightBracket = closedRight ? "]" : ")";
  return `${leftBracket}${Number(lower.toFixed(6))},${Number(upper.toFixed(6))}${rightBracket}`;
}
function buildChartExHistogramBins(values, rawSeries, sortByFrequency) {
  if (values.length === 0) {
    return [];
  }
  const rawRecord = rawSeries && typeof rawSeries === "object" ? rawSeries : null;
  const layoutProperties = rawRecord?.layoutProperties && typeof rawRecord.layoutProperties === "object" ? rawRecord.layoutProperties : null;
  const rawBinning = layoutProperties?.binning && typeof layoutProperties.binning === "object" ? layoutProperties.binning : null;
  const minValue = Math.min(...values);
  const maxValue = Math.max(...values);
  const explicitWidth = typeof rawBinning?.binSize === "number" && Number.isFinite(rawBinning.binSize) && rawBinning.binSize > 0 ? rawBinning.binSize : void 0;
  const explicitCount = typeof rawBinning?.binCount === "number" && Number.isFinite(rawBinning.binCount) && rawBinning.binCount > 0 ? rawBinning.binCount : typeof rawBinning?.count === "number" && Number.isFinite(rawBinning.count) && rawBinning.count > 0 ? rawBinning.count : void 0;
  const closedRight = rawBinning?.intervalClosed === "r" || rawBinning?.intervalClosed === "right";
  const mean = values.reduce((sum, value) => sum + value, 0) / values.length;
  const variance = values.reduce((sum, value) => sum + (value - mean) ** 2, 0) / Math.max(1, values.length);
  const standardDeviation = Math.sqrt(Math.max(0, variance));
  const allIntegers = values.every((value) => Math.abs(value - Math.round(value)) < 1e-9);
  const scottWidth = standardDeviation > 0 ? 3.49 * standardDeviation / Math.cbrt(values.length) : void 0;
  const fallbackWidth = explicitCount != null ? (maxValue - minValue) / Math.max(1, explicitCount) : scottWidth ?? (maxValue - minValue) / Math.max(1, Math.ceil(Math.log2(values.length) + 1));
  const roughWidth = explicitWidth ?? (allIntegers ? Math.max(1, Math.ceil(Math.max(fallbackWidth, 1e-6))) : niceHistogramStep(Math.max(fallbackWidth, 1e-6)));
  const binWidth = Math.max(roughWidth, 1e-6);
  const start = explicitWidth != null || explicitCount != null ? Math.floor(minValue / binWidth) * binWidth : minValue;
  const end = Math.max(start + binWidth, start + Math.ceil((maxValue - start) / binWidth) * binWidth);
  const binCount = Math.max(1, Math.ceil((end - start) / binWidth));
  const bins = Array.from({ length: binCount }, (_, index) => {
    const lower = start + binWidth * index;
    const upper = lower + binWidth;
    return {
      count: 0,
      label: formatHistogramBinLabel(lower, upper, index, closedRight),
      lower,
      upper
    };
  });
  values.forEach((value) => {
    if (!Number.isFinite(value)) {
      return;
    }
    const offset = (value - start) / binWidth;
    let binIndex = Math.floor(offset);
    if (closedRight && Math.abs(offset - Math.round(offset)) < 1e-9 && value > start) {
      binIndex -= 1;
    }
    if (value >= end) {
      binIndex = bins.length - 1;
    }
    if (value <= start) {
      binIndex = 0;
    }
    const target = bins[Math.max(0, Math.min(bins.length - 1, binIndex))];
    if (target) {
      target.count += 1;
    }
  });
  if (sortByFrequency) {
    bins.sort((left, right) => right.count - left.count || left.lower - right.lower);
  }
  return bins;
}
function buildChartExHistogramSeries(series, rawSeries, sortByFrequency) {
  const layout = resolveChartExSeriesLayout(rawSeries);
  const rawRecord = rawSeries && typeof rawSeries === "object" ? rawSeries : null;
  const hasBinning = Boolean(
    layout === "clusteredColumn" && rawRecord?.layoutProperties && typeof rawRecord.layoutProperties === "object" && rawRecord.layoutProperties.binning != null
  );
  if (!hasBinning) {
    return series;
  }
  const numericValues = series.values.filter((value) => typeof value === "number" && Number.isFinite(value));
  if (numericValues.length === 0) {
    return series;
  }
  const bins = buildChartExHistogramBins(numericValues, rawSeries, sortByFrequency);
  if (bins.length === 0) {
    return series;
  }
  return {
    ...series,
    categories: bins.map((bin) => bin.label),
    categoriesRef: null,
    raw: {
      ...series.raw,
      chartExHistogramBins: bins,
      chartExSourceValues: numericValues
    },
    values: bins.map((bin) => bin.count)
  };
}
function buildChartExParetoLineSeries(series, sourceRaw, index) {
  const counts = series.values.map((value) => typeof value === "number" && Number.isFinite(value) ? value : 0);
  const total = counts.reduce((sum, value) => sum + value, 0);
  let running = 0;
  const cumulative = counts.map((value) => {
    running += value;
    return total > 0 ? running / total * 100 : 0;
  });
  return {
    ...series,
    color: void 0,
    lineColor: void 0,
    markerColor: void 0,
    markerLineColor: void 0,
    markerSize: 7,
    markerSymbol: "circle",
    name: typeof sourceRaw?.text === "string" ? sourceRaw.text : "Pareto",
    raw: {
      ...series.raw ?? {},
      chartExLayout: "paretoLine",
      source: sourceRaw && typeof sourceRaw === "object" ? sourceRaw : void 0
    },
    values: cumulative
  };
}
function resolveChartExTextFormula(raw) {
  if (typeof raw === "string" && raw.length > 0) {
    return raw;
  }
  if (!raw || typeof raw !== "object") {
    return void 0;
  }
  const record = raw;
  if (typeof record.formula === "string" && record.formula.length > 0) {
    return record.formula;
  }
  if (typeof record.text === "string" && record.text.length > 0) {
    return record.text;
  }
  if (typeof record.value === "string" && record.value.length > 0) {
    return record.value;
  }
  return void 0;
}
function resolveChartExTitleText(raw) {
  if (typeof raw === "string" && raw.length > 0) {
    return raw;
  }
  if (!raw || typeof raw !== "object") {
    return void 0;
  }
  const record = raw;
  if (typeof record.text === "string" && record.text.length > 0) {
    return record.text;
  }
  const nestedText = record.text && typeof record.text === "object" ? resolveChartExTextFormula(record.text) : void 0;
  if (nestedText) {
    return nestedText;
  }
  return typeof record.value === "string" && record.value.length > 0 ? record.value : void 0;
}
function resolveChartExFallbackCategoryReference(workbook2, fallbackSheetIndex, valueFormula) {
  if (!valueFormula) {
    return null;
  }
  const resolved = resolveReferenceSheet(workbook2, fallbackSheetIndex, valueFormula);
  if (!resolved.sheet || !resolved.range || resolved.range.start.col <= 0) {
    return null;
  }
  return normalizeChartReference({
    formula: buildA1RangeFormula(
      resolved.sheetName,
      {
        col: resolved.range.start.col - 1,
        row: resolved.range.start.row
      },
      {
        col: resolved.range.start.col - 1,
        row: resolved.range.end.row
      }
    )
  });
}
function normalizeChartExSeries(workbook2, workbookSheetIndex, chartId, raw, dataById, index, chartType) {
  const series = raw && typeof raw === "object" ? raw : {};
  const dataId = typeof series.dataId === "number" ? series.dataId : null;
  const dataEntry = dataId != null ? dataById.get(dataId) ?? null : null;
  const dimensions = Array.isArray(dataEntry?.dimensions) ? dataEntry.dimensions.filter((value) => Boolean(value && typeof value === "object")) : [];
  const categoryDimension = dimensions.find((dimension) => dimension.dimType === "cat") ?? dimensions.find((dimension) => dimension.dimType === "name") ?? null;
  const valueDimension = dimensions.find((dimension) => dimension.dimType === "val" || dimension.dimType === "y" || dimension.dimType === "colorVal" || dimension.dimType === "size") ?? dimensions.find((dimension) => dimension !== categoryDimension) ?? categoryDimension;
  const categoryDimensionFormula = typeof categoryDimension?.formula === "string" ? categoryDimension.formula : void 0;
  const valueDimensionFormula = typeof valueDimension?.formula === "string" ? valueDimension.formula : void 0;
  const fallbackCategoryRef = (chartType === "Sunburst" || chartType === "Treemap") && !categoryDimension && typeof valueDimensionFormula === "string" ? resolveChartExFallbackCategoryReference(workbook2, workbookSheetIndex, valueDimensionFormula) : null;
  const categoriesRef = categoryDimension ? normalizeChartReference({
    formula: categoryDimensionFormula
  }) : fallbackCategoryRef;
  const valuesRef = valueDimension ? normalizeChartReference({
    formula: valueDimensionFormula
  }) : null;
  const resolvedValueCells = resolveReferenceValues(workbook2, workbookSheetIndex, valuesRef, "value");
  const values = resolvedValueCells.map((value) => typeof value === "number" && Number.isFinite(value) ? value : null);
  const colorStrings = chartType === "RegionMap" && valueDimension?.dimType === "colorStr" ? resolvedValueCells.map((value) => {
    if (typeof value === "string") {
      const trimmed = value.trim();
      return trimmed.length > 0 ? trimmed : null;
    }
    if (typeof value === "number" && Number.isFinite(value)) {
      return String(value);
    }
    return null;
  }) : [];
  const categories = resolveReferenceValues(workbook2, workbookSheetIndex, categoriesRef, "category");
  const hierarchyCategories = chartType === "Sunburst" || chartType === "Treemap" ? resolveReferenceRowPaths(workbook2, workbookSheetIndex, categoriesRef) : [];
  const seriesTextFormula = resolveChartExTextFormula(series.text);
  const shapeProperties = series.shapeProperties && typeof series.shapeProperties === "object" ? series.shapeProperties : void 0;
  const rawFillColor = typeof shapeProperties?.solidFillHex === "string" ? normalizeHexColor(shapeProperties.solidFillHex) : null;
  const rawLineColor = typeof shapeProperties?.lineColorHex === "string" ? normalizeHexColor(shapeProperties.lineColorHex) : null;
  return {
    bubbleSizeRef: null,
    bubbleSizes: [],
    categories,
    categoriesRef,
    color: rawFillColor ?? void 0,
    dataPoints: Array.isArray(series.dataPoints) ? series.dataPoints : [],
    dataPointStyles: void 0,
    formatIdx: typeof series.formatIdx === "number" ? series.formatIdx : void 0,
    hidden: typeof series.hidden === "boolean" ? series.hidden : void 0,
    id: `${chartId}-series-${index}`,
    invertIfNegative: void 0,
    lineColor: rawLineColor ?? rawFillColor ?? void 0,
    lineWidthPx: typeof shapeProperties?.lineWidth === "number" ? Math.max(1, Number(shapeProperties.lineWidth) / EMU_PER_PIXEL) : void 0,
    marker: void 0,
    markerColor: rawFillColor ?? void 0,
    markerLineColor: rawLineColor ?? rawFillColor ?? void 0,
    markerSize: void 0,
    markerSymbol: void 0,
    name: typeof series.text === "string" ? series.text : seriesTextFormula ? resolveSeriesName(workbook2, workbookSheetIndex, seriesTextFormula) : resolveChartReferenceLabel(workbook2, workbookSheetIndex, valuesRef, `Series ${index + 1}`),
    negativeColor: void 0,
    negativeLineColor: void 0,
    raw: {
      ...series,
      chartExColorStrings: colorStrings,
      chartExHierarchyCategories: hierarchyCategories,
      data: dataEntry,
      dimType: typeof valueDimension?.dimType === "string" ? valueDimension.dimType : void 0
    },
    shapeProperties,
    smooth: void 0,
    values,
    valuesRef
  };
}
function collapseChartExPointSeries(chartType, series) {
  if (chartType !== "Funnel" && chartType !== "Waterfall") {
    if ((chartType === "Sunburst" || chartType === "Treemap") && series.length > 1 && series.every((entry) => {
      const raw = entry.raw && typeof entry.raw === "object" ? entry.raw : null;
      return raw?.dimType === "size";
    })) {
      const primarySeries2 = series.find((entry) => entry.hidden !== true) ?? series[0] ?? null;
      if (!primarySeries2) {
        return series;
      }
      return [
        {
          ...primarySeries2,
          dataPoints: [],
          hidden: false
        }
      ];
    }
    return series;
  }
  const primarySeries = series.find((entry) => entry.hidden !== true) ?? series[0] ?? null;
  if (!primarySeries) {
    return series;
  }
  return [
    {
      ...primarySeries,
      categories: [],
      categoriesRef: null,
      dataPoints: [],
      hidden: false
    }
  ];
}
function normalizeChartExChart(workbook2, workbookSheetIndex, visibleSheetIndex, raw, index, themePalette) {
  const drawing = raw && typeof raw === "object" ? raw : {};
  const chart = drawing.chartEx && typeof drawing.chartEx === "object" ? drawing.chartEx : {};
  const plotArea = chart.plotArea && typeof chart.plotArea === "object" ? chart.plotArea : {};
  const rawSeries = Array.isArray(plotArea.series) ? plotArea.series : [];
  const seriesLayouts = rawSeries.map(resolveChartExSeriesLayout);
  const dataEntries = Array.isArray(chart.data) ? chart.data : [];
  const dataById = /* @__PURE__ */ new Map();
  dataEntries.forEach((entry) => {
    if (!entry || typeof entry !== "object") {
      return;
    }
    const record = entry;
    if (typeof record.id === "number") {
      dataById.set(record.id, record);
    }
  });
  const axes = Array.isArray(plotArea.axes) ? plotArea.axes.map(normalizeChartExAxis).filter((value) => Boolean(value)) : [];
  const primaryLayout = typeof chart.layout === "string" ? chart.layout : seriesLayouts.find((value) => typeof value === "string" && value.length > 0);
  const fallbackTitle = humanizeChartExLayoutLabel(primaryLayout);
  const chartTitle = resolveChartExTitleText(chart.title) ?? (chart.title != null ? "Chart Title" : fallbackTitle);
  const chartType = resolveChartExLayoutChartType(primaryLayout);
  const normalizedSeries = rawSeries.map((entry, seriesIndex) => normalizeChartExSeries(workbook2, workbookSheetIndex, `chart-ex-${workbookSheetIndex}-${index}`, entry, dataById, seriesIndex, chartType));
  const clusteredColumnSeriesIndex = seriesLayouts.findIndex((layout) => layout === "clusteredColumn");
  const hasParetoLine = seriesLayouts.includes("paretoLine");
  const clusteredColumnAxisIds = clusteredColumnSeriesIndex >= 0 ? resolveChartExSeriesAxisIds(rawSeries[clusteredColumnSeriesIndex]) : [];
  const paretoLineSeriesIndex = seriesLayouts.findIndex((layout) => layout === "paretoLine");
  const paretoLineAxisIds = paretoLineSeriesIndex >= 0 ? resolveChartExSeriesAxisIds(rawSeries[paretoLineSeriesIndex]) : [];
  const primaryHistogramSeries = clusteredColumnSeriesIndex >= 0 ? buildChartExHistogramSeries(normalizedSeries[clusteredColumnSeriesIndex] ?? normalizedSeries[0], rawSeries[clusteredColumnSeriesIndex], hasParetoLine) : null;
  const synthesizedParetoSeries = hasParetoLine && primaryHistogramSeries && primaryHistogramSeries.values.length > 0 ? buildChartExParetoLineSeries(primaryHistogramSeries, rawSeries[paretoLineSeriesIndex], paretoLineSeriesIndex) : null;
  const resolvedSeries = synthesizedParetoSeries ? [primaryHistogramSeries, synthesizedParetoSeries] : primaryHistogramSeries ? [
    primaryHistogramSeries,
    ...normalizedSeries.filter((_, seriesIndex) => seriesIndex !== clusteredColumnSeriesIndex)
  ] : collapseChartExPointSeries(chartType, normalizedSeries);
  const resolvedChartType = primaryHistogramSeries ? "ColumnClustered" : chartType;
  const resolvedGapWidth = primaryHistogramSeries ? 0 : void 0;
  const typeGroups = synthesizedParetoSeries ? [
    {
      axisIds: clusteredColumnAxisIds,
      chartType: "ColumnClustered",
      gapWidth: 0,
      raw: {
        gapWidth: 0,
        layout: "clusteredColumn"
      },
      series: [primaryHistogramSeries]
    },
    {
      axisIds: paretoLineAxisIds,
      chartType: "Line",
      raw: {
        layout: "paretoLine"
      },
      series: [synthesizedParetoSeries]
    }
  ] : [];
  const normalizedChart = {
    anchor: normalizeChartAnchor(drawing.anchor),
    autoTitleDeleted: void 0,
    axes,
    axisLabelColor: void 0,
    axisLineColor: void 0,
    categoryAxis: axes[0] ?? null,
    chartAreaBorderColor: void 0,
    chartAreaFillColor: void 0,
    chartColorPalette: void 0,
    chartColorPaletteOffset: void 0,
    chartExLayout: primaryLayout,
    chartPath: void 0,
    chartStyleId: void 0,
    chartType: resolvedChartType,
    dataLabels: rawSeries.length > 0 && rawSeries[0] && typeof rawSeries[0] === "object" ? normalizeChartDataLabels(rawSeries[0].dataLabels) : null,
    displayBlanksAs: void 0,
    editable: true,
    firstSliceAngle: void 0,
    fontFamily: void 0,
    gapWidth: resolvedGapWidth,
    holeSize: void 0,
    id: `chart-ex-${workbookSheetIndex}-${index}`,
    is3d: void 0,
    legend: normalizeChartExLegend(chart.legend),
    name: typeof drawing.name === "string" ? drawing.name : chartTitle,
    overlap: void 0,
    plotVisibleOnly: void 0,
    raw: chart,
    radarStyle: void 0,
    scatterStyle: void 0,
    roundedCorners: void 0,
    shape3d: void 0,
    seriesAxis: null,
    series: resolvedSeries,
    sheetIndex: visibleSheetIndex,
    showDlblsOverMax: void 0,
    sideWall: null,
    backWall: null,
    bubbleScale: void 0,
    bubble3d: void 0,
    floor: null,
    surfaceMaterial: void 0,
    textColor: void 0,
    title: chartTitle,
    titleColor: void 0,
    titleFontFamily: void 0,
    typeGroups,
    valueAxis: axes.find((axis) => axis.numberFormat || axis.majorGridlines) ?? axes[1] ?? null,
    varyColors: typeof chart.valueColors === "boolean" ? chart.valueColors : void 0,
    view3d: void 0,
    wireframe: void 0,
    workbookSheetIndex,
    zIndex: Array.isArray(drawing.drawingPath) && typeof drawing.drawingPath[0] === "number" ? drawing.drawingPath[0] + 1 : index + 1
  };
  applyBuiltinChartDefaults(normalizedChart, themePalette);
  return normalizedChart;
}
function cellValueToNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }
  if (value && typeof value === "object") {
    if (value.is_empty) {
      return null;
    }
    const candidates = [];
    if (typeof value.asNumber === "function") {
      candidates.push(value.asNumber());
    }
    if (typeof value.toJs === "function") {
      candidates.push(value.toJs());
    }
    if (typeof value.asText === "function") {
      candidates.push(value.asText());
    }
    if (typeof value.toString === "function") {
      candidates.push(value.toString());
    }
    for (const candidate of candidates) {
      if (typeof candidate === "number" && Number.isFinite(candidate)) {
        return candidate;
      }
      if (typeof candidate === "string") {
        const parsed = Number(candidate.replace(/,/g, ""));
        if (Number.isFinite(parsed)) {
          return parsed;
        }
      }
    }
  }
  if (typeof value === "string") {
    const parsed = Number(value.replace(/,/g, ""));
    return Number.isFinite(parsed) ? parsed : null;
  }
  return null;
}
function cellValueToDisplay(value) {
  if (value === null || value === void 0) {
    return "";
  }
  if (typeof value === "string") {
    return value;
  }
  if (typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  if (value && typeof value === "object") {
    if (value.is_empty) {
      return "";
    }
    const candidates = [];
    if (typeof value.asText === "function") {
      candidates.push(value.asText());
    }
    if (typeof value.toJs === "function") {
      candidates.push(value.toJs());
    }
    if (typeof value.toString === "function") {
      candidates.push(value.toString());
    }
    for (const candidate of candidates) {
      if (candidate === null || candidate === void 0) {
        continue;
      }
      if (typeof candidate === "string") {
        return candidate;
      }
      return String(candidate);
    }
  }
  return String(value);
}
function resolveReferenceValues(workbook2, fallbackSheetIndex, reference, mode) {
  if (!reference?.formula) {
    return reference?.values ?? [];
  }
  const resolved = resolveReferenceSheet(workbook2, fallbackSheetIndex, reference.formula);
  if (!resolved.sheet || !resolved.range) {
    return reference.values ?? [];
  }
  const values = [];
  for (let row = resolved.range.start.row; row <= resolved.range.end.row; row += 1) {
    for (let col = resolved.range.start.col; col <= resolved.range.end.col; col += 1) {
      const calculated = typeof resolved.sheet.getCalculatedValueAt === "function" ? resolved.sheet.getCalculatedValueAt(row, col) : null;
      const formatted = typeof resolved.sheet.getFormattedValueAt === "function" ? resolved.sheet.getFormattedValueAt(row, col) : calculated;
      if (mode === "value") {
        values.push(cellValueToNumber(calculated ?? formatted));
      } else {
        const display = cellValueToDisplay(formatted ?? calculated);
        const numeric = cellValueToNumber(calculated ?? formatted);
        values.push(display.length > 0 ? display : numeric !== null ? numeric : null);
      }
    }
  }
  return values;
}
function resolveSeriesName(workbook2, fallbackSheetIndex, rawName) {
  if (typeof rawName !== "string" || !rawName) {
    return void 0;
  }
  const resolved = resolveReferenceSheet(workbook2, fallbackSheetIndex, rawName);
  if (!resolved.sheet || !resolved.range) {
    return rawName;
  }
  const value = typeof resolved.sheet.getFormattedValueAt === "function" ? resolved.sheet.getFormattedValueAt(resolved.range.start.row, resolved.range.start.col) : null;
  const display = cellValueToDisplay(value);
  return display || rawName;
}
function normalizeChartReference(raw) {
  if (!raw || typeof raw !== "object") {
    return null;
  }
  const record = raw;
  const values = Array.isArray(record.numbers) ? record.numbers : Array.isArray(record.strings) ? record.strings : void 0;
  return {
    formula: typeof record.formula === "string" ? record.formula : void 0,
    refType: typeof record.refType === "string" ? record.refType : void 0,
    values
  };
}
function normalizeChartAxis(raw) {
  if (!raw || typeof raw !== "object") {
    return null;
  }
  const rawAxis = raw;
  const axis = rawAxis.axis && typeof rawAxis.axis === "object" ? rawAxis.axis : rawAxis;
  const numberFormat = axis.numberFormat && typeof axis.numberFormat === "object" ? axis.numberFormat : null;
  return {
    crossId: typeof rawAxis.crossId === "number" && Number.isFinite(rawAxis.crossId) ? rawAxis.crossId : void 0,
    crosses: typeof axis.crosses === "string" ? axis.crosses : void 0,
    crossBetween: typeof axis.crossBetween === "string" ? axis.crossBetween : void 0,
    delete: typeof axis.delete === "boolean" ? axis.delete : void 0,
    id: typeof rawAxis.id === "number" && Number.isFinite(rawAxis.id) ? rawAxis.id : void 0,
    labelPosition: typeof axis.labelPosition === "string" ? axis.labelPosition : void 0,
    logBase: typeof axis.logBase === "number" ? axis.logBase : void 0,
    orientation: typeof axis.orientation === "string" ? axis.orientation : void 0,
    majorUnit: typeof axis.majorUnit === "number" ? axis.majorUnit : void 0,
    max: typeof axis.maximum === "number" ? axis.maximum : void 0,
    min: typeof axis.minimum === "number" ? axis.minimum : void 0,
    majorGridlines: typeof axis.majorGridlines === "boolean" ? axis.majorGridlines : void 0,
    majorTickMark: typeof axis.majorTickMark === "string" ? axis.majorTickMark : void 0,
    minorUnit: typeof axis.minorUnit === "number" ? axis.minorUnit : void 0,
    minorGridlines: typeof axis.minorGridlines === "boolean" ? axis.minorGridlines : void 0,
    minorTickMark: typeof axis.minorTickMark === "string" ? axis.minorTickMark : void 0,
    numberFormat: numberFormat ? {
      formatCode: typeof numberFormat.formatCode === "string" ? numberFormat.formatCode : void 0,
      sourceLinked: typeof numberFormat.sourceLinked === "boolean" ? numberFormat.sourceLinked : void 0
    } : void 0,
    position: typeof axis.position === "string" ? axis.position : void 0,
    raw: axis,
    shapeProperties: axis.shapeProperties && typeof axis.shapeProperties === "object" ? axis.shapeProperties : void 0,
    tickLabelSkip: typeof axis.tickLabelSkip === "number" && Number.isFinite(axis.tickLabelSkip) ? axis.tickLabelSkip : void 0,
    tickMarkSkip: typeof axis.tickMarkSkip === "number" && Number.isFinite(axis.tickMarkSkip) ? axis.tickMarkSkip : void 0
  };
}
function mergeChartAxis(target, patch) {
  if (!patch) {
    return target ?? null;
  }
  return {
    ...target ?? {},
    ...patch
  };
}
function readChartAxisFromXml(axisNode) {
  if (!axisNode) {
    return null;
  }
  const numFmt = getFirstLocalChild(axisNode, "numFmt");
  const scalingNode = getFirstLocalChild(axisNode, "scaling");
  return {
    crossId: readChartNumericAttribute(axisNode, "crossAx"),
    crosses: getFirstLocalChild(axisNode, "crosses")?.getAttribute("val") ?? void 0,
    crossBetween: getFirstLocalChild(axisNode, "crossBetween")?.getAttribute("val") ?? void 0,
    delete: getFirstLocalChild(axisNode, "delete")?.getAttribute("val") === "1" ? true : getFirstLocalChild(axisNode, "delete")?.getAttribute("val") === "0" ? false : void 0,
    id: readChartNumericAttribute(axisNode, "axId"),
    labelPosition: getFirstLocalChild(axisNode, "tickLblPos")?.getAttribute("val") ?? void 0,
    logBase: readChartNumericAttribute(getFirstLocalChild(axisNode, "scaling"), "logBase"),
    orientation: getFirstLocalChild(scalingNode ?? axisNode, "orientation")?.getAttribute("val") ?? void 0,
    majorGridlines: Boolean(getFirstLocalChild(axisNode, "majorGridlines")),
    majorTickMark: getFirstLocalChild(axisNode, "majorTickMark")?.getAttribute("val") ?? void 0,
    majorUnit: readChartNumericAttribute(axisNode, "majorUnit"),
    max: readChartNumericAttribute(scalingNode, "max"),
    min: readChartNumericAttribute(scalingNode, "min"),
    minorGridlines: Boolean(getFirstLocalChild(axisNode, "minorGridlines")),
    minorTickMark: getFirstLocalChild(axisNode, "minorTickMark")?.getAttribute("val") ?? void 0,
    minorUnit: readChartNumericAttribute(axisNode, "minorUnit"),
    numberFormat: numFmt ? {
      formatCode: numFmt.getAttribute("formatCode") ?? void 0,
      sourceLinked: numFmt.getAttribute("sourceLinked") === "1" ? true : numFmt.getAttribute("sourceLinked") === "0" ? false : void 0
    } : void 0,
    position: getFirstLocalChild(axisNode, "axPos")?.getAttribute("val") ?? void 0,
    tickLabelSkip: readChartNumericAttribute(axisNode, "tickLblSkip"),
    tickMarkSkip: readChartNumericAttribute(axisNode, "tickMarkSkip")
  };
}
function readChartWallFromXml(wallNode, themePalette) {
  if (!wallNode) {
    return null;
  }
  const shapeProperties = getFirstLocalChild(wallNode, "spPr");
  const lineStyle = resolveChartLineStyle(shapeProperties, themePalette);
  return {
    fillColor: resolveChartFillColor(shapeProperties, themePalette) ?? void 0,
    hidden: shapeProperties ? getFirstLocalChild(shapeProperties, "noFill") != null : void 0,
    lineColor: lineStyle.color ?? void 0,
    thickness: readChartNumericAttribute(wallNode, "thickness")
  };
}
function normalizeChartDataLabels(raw) {
  if (!raw || typeof raw !== "object") {
    return null;
  }
  const labels = raw;
  const pointLabels = Array.isArray(labels.pointLabels) ? (() => {
    const normalized = [];
    for (const entry of labels.pointLabels) {
      if (!entry || typeof entry !== "object") {
        continue;
      }
      const point = entry;
      const index = typeof point.index === "number" && Number.isFinite(point.index) ? point.index : null;
      if (index == null) {
        continue;
      }
      const nextPoint = { index };
      if (typeof point.deleted === "boolean") {
        nextPoint.deleted = point.deleted;
      }
      if (typeof point.fontSizePt === "number" && Number.isFinite(point.fontSizePt)) {
        nextPoint.fontSizePt = point.fontSizePt;
      }
      if (typeof point.showBubbleSize === "boolean") {
        nextPoint.showBubbleSize = point.showBubbleSize;
      }
      if (typeof point.showCategoryName === "boolean") {
        nextPoint.showCategoryName = point.showCategoryName;
      }
      if (typeof point.showPercent === "boolean") {
        nextPoint.showPercent = point.showPercent;
      }
      if (typeof point.showSeriesName === "boolean") {
        nextPoint.showSeriesName = point.showSeriesName;
      }
      if (typeof point.showValue === "boolean") {
        nextPoint.showValue = point.showValue;
      }
      if (typeof point.x === "number" && Number.isFinite(point.x)) {
        nextPoint.x = point.x;
      }
      if (typeof point.y === "number" && Number.isFinite(point.y)) {
        nextPoint.y = point.y;
      }
      normalized.push(nextPoint);
    }
    return normalized;
  })() : void 0;
  return {
    pointLabels: pointLabels && pointLabels.length > 0 ? pointLabels : void 0,
    raw: labels,
    showBubbleSize: typeof labels.showBubbleSize === "boolean" ? labels.showBubbleSize : void 0,
    showCategoryName: typeof (labels.showCategoryName ?? labels.visibilityCategoryName) === "boolean" ? Boolean(labels.showCategoryName ?? labels.visibilityCategoryName) : void 0,
    showLegendKey: typeof labels.showLegendKey === "boolean" ? labels.showLegendKey : void 0,
    showPercent: typeof labels.showPercent === "boolean" ? labels.showPercent : void 0,
    showSeriesName: typeof (labels.showSeriesName ?? labels.visibilitySeriesName) === "boolean" ? Boolean(labels.showSeriesName ?? labels.visibilitySeriesName) : void 0,
    showValue: typeof (labels.showValue ?? labels.visibilityValue) === "boolean" ? Boolean(labels.showValue ?? labels.visibilityValue) : void 0
  };
}
function normalizeChartAnchor(raw) {
  if (!raw || typeof raw !== "object") {
    return {
      kind: "two-cell",
      from: { col: 0, colOffsetEmu: 0, row: 0, rowOffsetEmu: 0 },
      to: { col: 8, colOffsetEmu: 0, row: 15, rowOffsetEmu: 0 }
    };
  }
  const anchor = raw;
  const from = anchor.from && typeof anchor.from === "object" ? anchor.from : null;
  if (anchor.type === "oneCell") {
    return {
      from: {
        col: typeof from?.col === "number" ? from.col : 0,
        colOffsetEmu: typeof from?.colOffsetEmu === "number" ? from.colOffsetEmu : 0,
        row: typeof from?.row === "number" ? from.row : 0,
        rowOffsetEmu: typeof from?.rowOffsetEmu === "number" ? from.rowOffsetEmu : 0
      },
      kind: "one-cell",
      sizeEmu: {
        cx: typeof anchor.widthEmu === "number" ? anchor.widthEmu : 0,
        cy: typeof anchor.heightEmu === "number" ? anchor.heightEmu : 0
      }
    };
  }
  if (anchor.type === "absolute") {
    return {
      kind: "absolute",
      positionEmu: {
        x: typeof anchor.xEmu === "number" ? anchor.xEmu : 0,
        y: typeof anchor.yEmu === "number" ? anchor.yEmu : 0
      },
      sizeEmu: {
        cx: typeof anchor.widthEmu === "number" ? anchor.widthEmu : 0,
        cy: typeof anchor.heightEmu === "number" ? anchor.heightEmu : 0
      }
    };
  }
  const to = anchor.to && typeof anchor.to === "object" ? anchor.to : null;
  const fromColValue = from?.col;
  const fromColOffsetValue = from?.colOffsetEmu;
  const fromRowValue = from?.row;
  const fromRowOffsetValue = from?.rowOffsetEmu;
  const toColValue = to?.col;
  const toColOffsetValue = to?.colOffsetEmu;
  const toRowValue = to?.row;
  const toRowOffsetValue = to?.rowOffsetEmu;
  const fromCol = typeof fromColValue === "number" ? fromColValue : 0;
  const fromColOffsetEmu = typeof fromColOffsetValue === "number" ? fromColOffsetValue : 0;
  const fromRow = typeof fromRowValue === "number" ? fromRowValue : 0;
  const fromRowOffsetEmu = typeof fromRowOffsetValue === "number" ? fromRowOffsetValue : 0;
  const rawToCol = typeof toColValue === "number" ? toColValue : 0;
  const rawToColOffsetEmu = typeof toColOffsetValue === "number" ? toColOffsetValue : 0;
  const rawToRow = typeof toRowValue === "number" ? toRowValue : 0;
  const rawToRowOffsetEmu = typeof toRowOffsetValue === "number" ? toRowOffsetValue : 0;
  return {
    kind: "two-cell",
    from: {
      col: fromCol,
      colOffsetEmu: fromColOffsetEmu,
      row: fromRow,
      rowOffsetEmu: fromRowOffsetEmu
    },
    to: {
      col: rawToCol,
      colOffsetEmu: rawToColOffsetEmu,
      row: rawToRow,
      rowOffsetEmu: rawToRowOffsetEmu
    }
  };
}
function parseMarkerNode(node) {
  if (!node) {
    return null;
  }
  const col = Number(getFirstLocalChild(node, "col")?.textContent ?? Number.NaN);
  const row = Number(getFirstLocalChild(node, "row")?.textContent ?? Number.NaN);
  const colOffsetEmu = Number(getFirstLocalChild(node, "colOff")?.textContent ?? 0);
  const rowOffsetEmu = Number(getFirstLocalChild(node, "rowOff")?.textContent ?? 0);
  if (!Number.isFinite(col) || !Number.isFinite(row)) {
    return null;
  }
  return {
    col: Math.max(0, Math.round(col)),
    colOffsetEmu: Number.isFinite(colOffsetEmu) ? Math.max(0, Math.round(colOffsetEmu)) : 0,
    row: Math.max(0, Math.round(row)),
    rowOffsetEmu: Number.isFinite(rowOffsetEmu) ? Math.max(0, Math.round(rowOffsetEmu)) : 0
  };
}
function parseChartAnchorNode(anchorNode) {
  if (anchorNode.localName === "twoCellAnchor") {
    const from = parseMarkerNode(getFirstLocalChild(anchorNode, "from"));
    const to = parseMarkerNode(getFirstLocalChild(anchorNode, "to"));
    return from && to ? { from, kind: "two-cell", to } : null;
  }
  if (anchorNode.localName === "oneCellAnchor") {
    const from = parseMarkerNode(getFirstLocalChild(anchorNode, "from"));
    const ext2 = getFirstLocalChild(anchorNode, "ext");
    const cx2 = Number(ext2?.getAttribute("cx") ?? Number.NaN);
    const cy2 = Number(ext2?.getAttribute("cy") ?? Number.NaN);
    return from && Number.isFinite(cx2) && Number.isFinite(cy2) ? {
      from,
      kind: "one-cell",
      sizeEmu: {
        cx: Math.max(0, Math.round(cx2)),
        cy: Math.max(0, Math.round(cy2))
      }
    } : null;
  }
  const pos = getFirstLocalChild(anchorNode, "pos");
  const ext = getFirstLocalChild(anchorNode, "ext");
  const x = Number(pos?.getAttribute("x") ?? Number.NaN);
  const y = Number(pos?.getAttribute("y") ?? Number.NaN);
  const cx = Number(ext?.getAttribute("cx") ?? Number.NaN);
  const cy = Number(ext?.getAttribute("cy") ?? Number.NaN);
  return Number.isFinite(x) && Number.isFinite(y) && Number.isFinite(cx) && Number.isFinite(cy) ? {
    kind: "absolute",
    positionEmu: {
      x: Math.round(x),
      y: Math.round(y)
    },
    sizeEmu: {
      cx: Math.max(0, Math.round(cx)),
      cy: Math.max(0, Math.round(cy))
    }
  } : null;
}
function isCollapsedChartAnchor(anchor) {
  if (anchor.kind !== "two-cell") {
    return false;
  }
  const collapsedWidth = anchor.to.col < anchor.from.col || anchor.to.col === anchor.from.col && anchor.to.colOffsetEmu <= anchor.from.colOffsetEmu;
  const collapsedHeight = anchor.to.row < anchor.from.row || anchor.to.row === anchor.from.row && anchor.to.rowOffsetEmu <= anchor.from.rowOffsetEmu;
  return collapsedWidth || collapsedHeight;
}
function normalizeChartSeries(workbook2, workbookSheetIndex, chartId, raw, index) {
  const series = raw && typeof raw === "object" ? raw : {};
  const categoriesRef = normalizeChartReference(series.categories);
  const valuesRef = normalizeChartReference(series.values);
  const shapeProperties = series.shapeProperties && typeof series.shapeProperties === "object" ? series.shapeProperties : void 0;
  const rawFillColor = typeof shapeProperties?.solidFillHex === "string" ? normalizeHexColor(shapeProperties.solidFillHex) : null;
  const rawLineColor = typeof shapeProperties?.lineColorHex === "string" ? normalizeHexColor(shapeProperties.lineColorHex) : null;
  const bubbleSizeRef = normalizeChartReference(series.bubbleSize ?? series.bubbleSizes ?? series.bubbles);
  return {
    bubbleSizeRef,
    bubbleSizes: resolveReferenceValues(workbook2, workbookSheetIndex, bubbleSizeRef, "value").map((value) => typeof value === "number" && Number.isFinite(value) ? value : null),
    categories: resolveReferenceValues(workbook2, workbookSheetIndex, categoriesRef, "category"),
    categoriesRef,
    color: rawFillColor ?? void 0,
    dataPoints: Array.isArray(series.dataPoints) ? series.dataPoints : [],
    dataPointStyles: void 0,
    id: `${chartId}-series-${index}`,
    invertIfNegative: typeof series.invertIfNegative === "boolean" ? series.invertIfNegative : void 0,
    lineColor: rawLineColor ?? rawFillColor ?? void 0,
    lineWidthPx: typeof shapeProperties?.lineWidth === "number" ? Math.max(1, Number(shapeProperties.lineWidth) / EMU_PER_PIXEL) : void 0,
    marker: series.marker && typeof series.marker === "object" ? series.marker : void 0,
    markerColor: void 0,
    markerLineColor: void 0,
    markerSize: series.marker && typeof series.marker === "object" && typeof series.marker.size === "number" ? Number(series.marker.size) : void 0,
    markerSymbol: series.marker && typeof series.marker === "object" && typeof series.marker.symbol === "string" ? String(series.marker.symbol) : void 0,
    name: resolveSeriesName(workbook2, workbookSheetIndex, series.name),
    negativeColor: void 0,
    negativeLineColor: void 0,
    raw: series,
    shapeProperties,
    smooth: typeof series.smooth === "boolean" ? series.smooth : void 0,
    values: resolveReferenceValues(workbook2, workbookSheetIndex, valuesRef, "value").map((value) => typeof value === "number" && Number.isFinite(value) ? value : null),
    valuesRef
  };
}
function normalizeChartTypeGroup(workbook2, workbookSheetIndex, chartId, raw, index) {
  if (!raw || typeof raw !== "object") {
    return null;
  }
  const group = raw;
  const rawSeries = Array.isArray(group.series) ? group.series : [];
  return {
    axisIds: Array.isArray(group.axisIds) ? group.axisIds.filter((value) => typeof value === "number" && Number.isFinite(value)) : void 0,
    chartType: typeof group.chartType === "string" ? group.chartType : "ColumnClustered",
    dataLabels: normalizeChartDataLabels(group.dataLabels),
    gapWidth: typeof group.gapWidth === "number" && Number.isFinite(group.gapWidth) ? group.gapWidth : void 0,
    is3d: typeof group.is3D === "boolean" ? group.is3D : void 0,
    overlap: typeof group.overlap === "number" && Number.isFinite(group.overlap) ? group.overlap : void 0,
    raw: group,
    series: rawSeries.map((entry, seriesIndex) => normalizeChartSeries(workbook2, workbookSheetIndex, `${chartId}-group-${index}`, entry, seriesIndex)),
    varyColors: typeof group.varyColors === "boolean" ? group.varyColors : void 0
  };
}
function normalizeChartsheet(raw, index) {
  const chartsheet = raw && typeof raw === "object" ? raw : {};
  return {
    chartIds: Array.isArray(chartsheet.chartIds) ? chartsheet.chartIds.filter((value) => typeof value === "string") : [],
    chartPath: typeof chartsheet.chartPath === "string" ? chartsheet.chartPath : void 0,
    id: `chartsheet-${index}`,
    index,
    name: typeof chartsheet.name === "string" ? chartsheet.name : `Chart ${index + 1}`,
    raw: chartsheet,
    workbookSheetIndex: typeof chartsheet.workbookSheetIndex === "number" ? chartsheet.workbookSheetIndex : void 0
  };
}
function buildTabs(workbook2, chartsheets2, visibleSheetIndexByWorkbookSheetIndex, showHiddenSheets = false) {
  const rawOrder = Array.isArray(workbook2.sheetOrder) ? workbook2.sheetOrder : [];
  if (rawOrder.length === 0) {
    return workbook2.sheetNames.flatMap((name, index) => {
      const worksheet = workbook2.getSheet(index);
      const visibility = normalizeWorksheetVisibility(worksheet.visibility);
      if (!showHiddenSheets && visibility !== "visible") {
        return [];
      }
      return [{
        id: `sheet-${index}`,
        index,
        kind: "sheet",
        name,
        sheetIndex: visibleSheetIndexByWorkbookSheetIndex.get(index) ?? index,
        visibility,
        workbookSheetIndex: index
      }];
    });
  }
  return rawOrder.flatMap((entry, index) => {
    const slotType = typeof entry.slotType === "string" ? entry.slotType : "worksheet";
    const slotIndex = typeof entry.index === "number" ? entry.index : index;
    if (slotType === "chartsheet") {
      const chartsheet = chartsheets2[slotIndex];
      return chartsheet ? [{
        chartsheetIndex: slotIndex,
        id: `chartsheet-${slotIndex}`,
        index,
        kind: "chartsheet",
        name: chartsheet.name
      }] : [];
    }
    const worksheet = workbook2.getSheet(slotIndex);
    const visibility = normalizeWorksheetVisibility(worksheet.visibility);
    if (!showHiddenSheets && visibility !== "visible") {
      return [];
    }
    return [{
      id: `sheet-${slotIndex}`,
      index,
      kind: "sheet",
      name: worksheet.name,
      sheetIndex: visibleSheetIndexByWorkbookSheetIndex.get(slotIndex) ?? slotIndex,
      visibility,
      workbookSheetIndex: slotIndex
    }];
  });
}
function collectChartOriginsForSheet(archive, origin) {
  if (!origin) {
    return [];
  }
  const chartOrigins = [];
  for (const attachment of origin.attachments) {
    const drawingXml = readArchiveText(archive, attachment.drawingPath);
    const relsXml = readArchiveText(archive, attachment.drawingRelsPath);
    if (!drawingXml || !relsXml) {
      continue;
    }
    const drawingDocument = parseXml(drawingXml);
    const relsDocument = parseXml(relsXml);
    if (!drawingDocument || !relsDocument) {
      continue;
    }
    const relationships = /* @__PURE__ */ new Map();
    for (const node of getLocalDescendants(relsDocument, "Relationship")) {
      const id = node.getAttribute("Id");
      const target = node.getAttribute("Target");
      const type = node.getAttribute("Type");
      if (id && target) {
        relationships.set(id, {
          target: resolveRelationshipPath(attachment.drawingRelsPath ?? attachment.drawingPath, target),
          type
        });
      }
    }
    const anchorNodes = Array.from(drawingDocument.documentElement.childNodes).filter(
      (node) => node.nodeType === Node.ELEMENT_NODE && (node.localName === "twoCellAnchor" || node.localName === "oneCellAnchor" || node.localName === "absoluteAnchor")
    );
    let chartAnchorIndex = 0;
    for (const anchorNode of anchorNodes) {
      const graphicFrame = getFirstLocalDescendant(anchorNode, "graphicFrame");
      const chartNode = graphicFrame ? getFirstLocalDescendant(graphicFrame, "chart") : null;
      const relationshipId = chartNode?.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id") ?? chartNode?.getAttribute("r:id") ?? chartNode?.getAttribute("id");
      if (!relationshipId) {
        continue;
      }
      const relationship = relationships.get(relationshipId);
      if (!relationship || relationship.type !== CHART_REL_TYPE && relationship.type !== CHART_EX_REL_TYPE) {
        continue;
      }
      chartOrigins.push({
        anchorIndex: chartAnchorIndex,
        anchor: parseChartAnchorNode(anchorNode),
        chartKind: relationship.type === CHART_EX_REL_TYPE ? "modern" : "classic",
        chartPath: relationship.target,
        drawingPath: attachment.drawingPath,
        workbookSheetIndex: origin.workbookSheetIndex
      });
      chartAnchorIndex += 1;
    }
  }
  return chartOrigins;
}
function applyChartOrigins(chartsByWorkbookSheetIndex2, chartOriginsById, archive, sheetOrigins) {
  for (let workbookSheetIndex = 0; workbookSheetIndex < chartsByWorkbookSheetIndex2.length; workbookSheetIndex += 1) {
    const charts = chartsByWorkbookSheetIndex2[workbookSheetIndex] ?? [];
    const origins = collectChartOriginsForSheet(archive, sheetOrigins[workbookSheetIndex] ?? null);
    const originsByKind = {
      classic: origins.filter((origin) => origin.chartKind === "classic"),
      modern: origins.filter((origin) => origin.chartKind === "modern")
    };
    const chartIndexByKind = {
      classic: 0,
      modern: 0
    };
    charts.forEach((chart) => {
      const chartKind = chart.id.startsWith("chart-ex-") ? "modern" : "classic";
      const origin = originsByKind[chartKind][chartIndexByKind[chartKind]];
      chartIndexByKind[chartKind] += 1;
      if (!origin) {
        return;
      }
      if (origin.anchor && isCollapsedChartAnchor(chart.anchor)) {
        chart.anchor = origin.anchor;
      } else if (origin.anchor && chart.anchor.kind === "two-cell" && chart.anchor.from.col === 0 && chart.anchor.from.row === 0) {
        chart.anchor = origin.anchor;
      }
      chart.chartPath = origin.chartPath ?? void 0;
      chartOriginsById.set(chart.id, origin);
    });
  }
}
function hydrateWorkbookChartStyles(chartsByWorkbookSheetIndex2, imageAssets) {
  const chartOriginsById = /* @__PURE__ */ new Map();
  applyChartOrigins(chartsByWorkbookSheetIndex2, chartOriginsById, imageAssets.archive, imageAssets.sheetOrigins);
  for (const charts of chartsByWorkbookSheetIndex2) {
    for (const chart of charts) {
      applyChartStyleFromXml(chart, chart.chartPath, imageAssets.archive, imageAssets.themePalette);
      applyBuiltinChartDefaults(chart, imageAssets.themePalette);
    }
  }
  return chartOriginsById;
}
function loadWorkbookChartAssets(workbook2, imageAssets, visibleSheetIndexByWorkbookSheetIndex, showHiddenSheets = false) {
  const excludedChartIds = /* @__PURE__ */ new Set();
  const chartsByWorkbookSheetIndex2 = Array.from({ length: workbook2.sheetCount }, (_, workbookSheetIndex) => {
    const worksheet = workbook2.getSheet(workbookSheetIndex);
    const rawCharts = worksheet.charts;
    const rawChartsEx = worksheet.chartsEx;
    const visibleSheetIndex = visibleSheetIndexByWorkbookSheetIndex.get(workbookSheetIndex) ?? workbookSheetIndex;
    const classicCharts = rawCharts.map((rawChart, chartIndex) => {
      const chartId = `chart-${workbookSheetIndex}-${chartIndex}`;
      if (rawChart.hidden || !rawChart.anchor) {
        excludedChartIds.add(chartId);
      }
      const drawing = rawChart && typeof rawChart === "object" ? rawChart : {};
      const chart = drawing.chart && typeof drawing.chart === "object" ? drawing.chart : {};
      const rawView3d = chart.view3D && typeof chart.view3D === "object" ? chart.view3D : null;
      const rawSeries = Array.isArray(chart.series) ? chart.series : [];
      const chartLevelDataLabels = normalizeChartDataLabels(chart.dataLabels);
      const firstSeriesDataLabels = rawSeries.length > 0 && rawSeries[0] && typeof rawSeries[0] === "object" ? normalizeChartDataLabels(rawSeries[0].dataLabels) : null;
      return {
        anchor: normalizeChartAnchor(drawing.anchor),
        autoTitleDeleted: typeof chart.autoTitleDeleted === "boolean" ? chart.autoTitleDeleted : void 0,
        axes: Array.isArray(chart.axes) ? chart.axes.map(normalizeChartAxis).filter((value) => Boolean(value)) : [],
        axisLabelColor: void 0,
        axisLineColor: void 0,
        categoryAxis: normalizeChartAxis(chart.categoryAxis),
        chartAreaBorderColor: void 0,
        chartAreaFillColor: void 0,
        chartColorPalette: void 0,
        chartColorPaletteOffset: void 0,
        chartPath: void 0,
        chartStyleId: void 0,
        chartType: typeof chart.chartType === "string" ? chart.chartType : "ColumnClustered",
        dataLabels: chartLevelDataLabels ?? firstSeriesDataLabels,
        displayBlanksAs: typeof chart.displayBlanksAs === "string" ? chart.displayBlanksAs : void 0,
        editable: true,
        firstSliceAngle: typeof chart.firstSliceAngle === "number" ? chart.firstSliceAngle : void 0,
        fontFamily: void 0,
        gapWidth: typeof chart.gapWidth === "number" ? chart.gapWidth : void 0,
        holeSize: typeof chart.holeSize === "number" ? chart.holeSize : void 0,
        id: chartId,
        is3d: typeof chart.is3D === "boolean" ? chart.is3D : void 0,
        legend: normalizeLegend(chart.legend) ? {
          ...normalizeLegend(chart.legend),
          position: normalizeLegendPosition(normalizeLegend(chart.legend)?.position)
        } : null,
        name: typeof drawing.name === "string" ? drawing.name : void 0,
        overlap: typeof chart.overlap === "number" ? chart.overlap : void 0,
        plotVisibleOnly: typeof chart.plotVisibleOnly === "boolean" ? chart.plotVisibleOnly : void 0,
        raw: chart,
        radarStyle: typeof chart.radarStyle === "string" ? chart.radarStyle : void 0,
        scatterStyle: typeof chart.scatterStyle === "string" ? chart.scatterStyle : void 0,
        roundedCorners: typeof chart.roundedCorners === "boolean" ? chart.roundedCorners : void 0,
        shape3d: typeof chart.shape === "string" ? chart.shape : typeof chart.shape3d === "string" ? chart.shape3d : void 0,
        seriesAxis: null,
        series: rawSeries.map((entry, seriesIndex) => normalizeChartSeries(workbook2, workbookSheetIndex, chartId, entry, seriesIndex)),
        sheetIndex: visibleSheetIndex,
        showDlblsOverMax: typeof chart.showDlblsOverMax === "boolean" ? chart.showDlblsOverMax : void 0,
        sideWall: null,
        backWall: null,
        bubbleScale: typeof chart.bubbleScale === "number" ? chart.bubbleScale : void 0,
        bubble3d: typeof chart.bubble3d === "boolean" ? chart.bubble3d : void 0,
        floor: null,
        surfaceMaterial: void 0,
        textColor: void 0,
        title: typeof chart.title === "string" ? chart.title : void 0,
        titleColor: void 0,
        titleFontFamily: void 0,
        typeGroups: Array.isArray(chart.typeGroups) ? chart.typeGroups.map((entry, groupIndex) => normalizeChartTypeGroup(workbook2, workbookSheetIndex, chartId, entry, groupIndex)).filter((value) => value != null) : [],
        valueAxis: normalizeChartAxis(chart.valueAxis),
        varyColors: typeof chart.varyColors === "boolean" ? chart.varyColors : void 0,
        view3d: rawView3d ? {
          depthPercent: typeof rawView3d.depthPercent === "number" ? rawView3d.depthPercent : void 0,
          perspective: typeof rawView3d.perspective === "number" ? rawView3d.perspective : void 0,
          rAngAx: typeof rawView3d.rAngAx === "boolean" ? rawView3d.rAngAx : typeof rawView3d.rightAngleAxes === "boolean" ? rawView3d.rightAngleAxes : void 0,
          rotX: typeof rawView3d.rotX === "number" ? rawView3d.rotX : typeof rawView3d.rotateX === "number" ? rawView3d.rotateX : void 0,
          rotY: typeof rawView3d.rotY === "number" ? rawView3d.rotY : typeof rawView3d.rotateY === "number" ? rawView3d.rotateY : void 0
        } : void 0,
        wireframe: typeof chart.wireframe === "boolean" ? chart.wireframe : void 0,
        workbookSheetIndex,
        zIndex: Array.isArray(drawing.drawingPath) && typeof drawing.drawingPath[0] === "number" ? drawing.drawingPath[0] + 1 : chartIndex + 1
      };
    });
    const modernCharts = rawChartsEx.map((rawChartEx, chartExIndex) => {
      const chartId = `chart-ex-${workbookSheetIndex}-${chartExIndex}`;
      if (rawChartEx.hidden || !rawChartEx.anchor) {
        excludedChartIds.add(chartId);
      }
      return normalizeChartExChart(
        workbook2,
        workbookSheetIndex,
        visibleSheetIndex,
        rawChartEx,
        chartExIndex,
        imageAssets?.themePalette ?? null
      );
    });
    return [...classicCharts, ...modernCharts];
  });
  const chartsheets2 = Array.isArray(workbook2.chartsheets) ? workbook2.chartsheets.map((entry, index) => normalizeChartsheet(entry, index)) : [];
  const tabs2 = buildTabs(workbook2, chartsheets2, visibleSheetIndexByWorkbookSheetIndex, showHiddenSheets);
  const chartOriginsById = imageAssets ? hydrateWorkbookChartStyles(chartsByWorkbookSheetIndex2, imageAssets) : /* @__PURE__ */ new Map();
  if (imageAssets) {
    for (let index = 0; index < chartsByWorkbookSheetIndex2.length; index += 1) {
      chartsByWorkbookSheetIndex2[index] = (chartsByWorkbookSheetIndex2[index] ?? []).filter((chart) => !excludedChartIds.has(chart.id));
    }
    for (const id of excludedChartIds) {
      chartOriginsById.delete(id);
    }
  } else {
    for (let index = 0; index < chartsByWorkbookSheetIndex2.length; index += 1) {
      chartsByWorkbookSheetIndex2[index] = (chartsByWorkbookSheetIndex2[index] ?? []).filter((chart) => !excludedChartIds.has(chart.id));
    }
    for (const charts of chartsByWorkbookSheetIndex2) {
      for (const chart of charts) {
        applyBuiltinChartDefaults(chart, null);
      }
    }
  }
  return {
    chartOriginsById,
    chartsByWorkbookSheetIndex: chartsByWorkbookSheetIndex2,
    chartsheets: chartsheets2,
    tabs: tabs2
  };
}

// src/images.ts
import { strFromU8 as strFromU83, strToU8 as strToU82, unzipSync as unzipSync2, zipSync } from "fflate";

// src/colors.ts
function normalizeHexColor2(value) {
  const hex = value.replace(/^#/, "");
  if (hex.length === 8) {
    return `#${hex.slice(2).toLowerCase()}`;
  }
  if (hex.length === 6) {
    return `#${hex.toLowerCase()}`;
  }
  return null;
}
function parseHexColor2(color) {
  const normalized = normalizeHexColor2(color);
  const match = normalized ? /^#([0-9a-f]{6})$/.exec(normalized) : null;
  if (!match) {
    return null;
  }
  const hex = match[1];
  return [
    Number.parseInt(hex.slice(0, 2), 16),
    Number.parseInt(hex.slice(2, 4), 16),
    Number.parseInt(hex.slice(4, 6), 16)
  ];
}
function rgbToHsl2(red, green, blue) {
  const normalizedRed = red / 255;
  const normalizedGreen = green / 255;
  const normalizedBlue = blue / 255;
  const max = Math.max(normalizedRed, normalizedGreen, normalizedBlue);
  const min = Math.min(normalizedRed, normalizedGreen, normalizedBlue);
  const lightness = (max + min) / 2;
  if (max === min) {
    return [0, 0, lightness];
  }
  const delta = max - min;
  const saturation = lightness > 0.5 ? delta / (2 - max - min) : delta / (max + min);
  let hue = 0;
  switch (max) {
    case normalizedRed:
      hue = (normalizedGreen - normalizedBlue) / delta + (normalizedGreen < normalizedBlue ? 6 : 0);
      break;
    case normalizedGreen:
      hue = (normalizedBlue - normalizedRed) / delta + 2;
      break;
    default:
      hue = (normalizedRed - normalizedGreen) / delta + 4;
      break;
  }
  return [hue / 6, saturation, lightness];
}
function hueToRgb2(p, q, t) {
  let nextT = t;
  if (nextT < 0) {
    nextT += 1;
  }
  if (nextT > 1) {
    nextT -= 1;
  }
  if (nextT < 1 / 6) {
    return p + (q - p) * 6 * nextT;
  }
  if (nextT < 1 / 2) {
    return q;
  }
  if (nextT < 2 / 3) {
    return p + (q - p) * (2 / 3 - nextT) * 6;
  }
  return p;
}
function hslToRgb2(hue, saturation, lightness) {
  if (saturation === 0) {
    const gray = Math.round(lightness * 255);
    return [gray, gray, gray];
  }
  const q = lightness < 0.5 ? lightness * (1 + saturation) : lightness + saturation - lightness * saturation;
  const p = 2 * lightness - q;
  return [
    Math.round(hueToRgb2(p, q, hue + 1 / 3) * 255),
    Math.round(hueToRgb2(p, q, hue) * 255),
    Math.round(hueToRgb2(p, q, hue - 1 / 3) * 255)
  ];
}
function rgbToHex2(red, green, blue) {
  return `#${[red, green, blue].map((channel) => Math.max(0, Math.min(255, channel)).toString(16).padStart(2, "0")).join("")}`;
}
function applyExcelTint(baseColor, tint) {
  const rgb = parseHexColor2(baseColor);
  if (!rgb || !Number.isFinite(tint) || tint === 0) {
    return normalizeHexColor2(baseColor);
  }
  const [hue, saturation, lightness] = rgbToHsl2(rgb[0], rgb[1], rgb[2]);
  const nextLightness = tint < 0 ? lightness * (1 + tint) : lightness * (1 - tint) + tint;
  const [nextRed, nextGreen, nextBlue] = hslToRgb2(hue, saturation, Math.max(0, Math.min(1, nextLightness)));
  return rgbToHex2(nextRed, nextGreen, nextBlue);
}
function resolveWorkbookColor(color, themePalette) {
  if (!color) {
    return null;
  }
  const directHex = ["hex", "rgb", "argb"].map((key) => color[key]).find((value) => typeof value === "string" && value.trim().length > 0);
  if (directHex) {
    return normalizeHexColor2(directHex);
  }
  const themeValue = color.theme;
  const numericTheme = typeof themeValue === "number" ? themeValue : typeof themeValue === "string" && themeValue.trim().length > 0 ? Number(themeValue) : Number.NaN;
  const themeColor = Number.isFinite(numericTheme) ? themePalette?.colorsByIndex[numericTheme] ?? null : null;
  if (!themeColor) {
    return null;
  }
  const tintValue = color.tint;
  const tint = typeof tintValue === "number" ? tintValue : typeof tintValue === "string" && tintValue.trim().length > 0 ? Number(tintValue) : Number.NaN;
  return Number.isFinite(tint) ? applyExcelTint(themeColor, tint) : themeColor;
}

// src/images.ts
var REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
var SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
var DRAWING_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
var EMU_PER_PIXEL2 = 9525;
var MIN_COL_WIDTH_PX = 30;
var MIN_ROW_HEIGHT_PX = 16;
var DEFAULT_COL_WIDTH_EMU = 64 * EMU_PER_PIXEL2;
var DEFAULT_ROW_HEIGHT_EMU = 20 * EMU_PER_PIXEL2;
var DEFAULT_COLUMN_CHARACTER_WIDTH_PX = 7;
var columnCharacterWidthCache = /* @__PURE__ */ new Map();
function measureColumnCharacterWidthPx(fontFamily, fontSizePt) {
  const normalizedFamily = typeof fontFamily === "string" && fontFamily.trim().length > 0 ? fontFamily.trim() : "Calibri";
  const normalizedSizePt = typeof fontSizePt === "number" && Number.isFinite(fontSizePt) && fontSizePt > 0 ? fontSizePt : 11;
  const cacheKey = `${normalizedFamily}|${normalizedSizePt}`;
  const cached = columnCharacterWidthCache.get(cacheKey);
  if (cached !== void 0) {
    return cached;
  }
  const fontSizePx = normalizedSizePt * (96 / 72);
  const font = `${fontSizePx}px "${normalizedFamily}"`;
  let width = DEFAULT_COLUMN_CHARACTER_WIDTH_PX;
  try {
    const context = typeof document !== "undefined" ? document.createElement("canvas").getContext("2d") : typeof OffscreenCanvas !== "undefined" ? new OffscreenCanvas(32, 32).getContext("2d") : null;
    if (context) {
      context.font = font;
      width = Math.max(1, context.measureText("0").width);
    }
  } catch {
    width = DEFAULT_COLUMN_CHARACTER_WIDTH_PX;
  }
  columnCharacterWidthCache.set(cacheKey, width);
  return width;
}
function sheetColumnWidthToPixels(width, columnCharacterWidthPx = DEFAULT_COLUMN_CHARACTER_WIDTH_PX) {
  if (!Number.isFinite(width) || width <= 0) {
    return MIN_COL_WIDTH_PX;
  }
  const digitWidth = Math.max(1, columnCharacterWidthPx);
  const pixels = width < 1 ? Math.floor(width * (digitWidth + 5) + 0.5) : Math.floor((256 * width + Math.floor(128 / digitWidth)) / 256 * digitWidth);
  return Math.max(MIN_COL_WIDTH_PX, pixels);
}
function resolveWorksheetDefaultColumnWidthPixels(worksheet, columnCharacterWidthPx = DEFAULT_COLUMN_CHARACTER_WIDTH_PX, fallbackPx = sheetColumnWidthToPixels(8.43, columnCharacterWidthPx)) {
  const width = typeof worksheet.defaultColumnWidth === "number" ? worksheet.defaultColumnWidth : Number.NaN;
  return Number.isFinite(width) && width > 0 ? sheetColumnWidthToPixels(width, columnCharacterWidthPx) : fallbackPx;
}
function resolveWorksheetDefaultRowHeightPixels(worksheet, fallbackPx = Math.max(MIN_ROW_HEIGHT_PX, Math.round(15 * 1.33))) {
  const height = typeof worksheet.defaultRowHeight === "number" ? worksheet.defaultRowHeight : Number.NaN;
  return Number.isFinite(height) && height > 0 ? Math.max(MIN_ROW_HEIGHT_PX, Math.round(height * 1.33)) : fallbackPx;
}
function resolveWorksheetMergeMetadata(worksheet) {
  const mergeMetadata = {
    hasHorizontalMerges: false,
    hasVerticalMerges: false,
    maxHorizontalMergeEndCol: -1,
    maxVerticalMergeEndRow: -1
  };
  const mergedRegions = Array.isArray(worksheet.mergedRegions) ? worksheet.mergedRegions : [];
  for (const rawRegion of mergedRegions) {
    let range = null;
    if (typeof rawRegion === "string") {
      range = parseA1RangeReference(rawRegion);
    } else if (rawRegion && typeof rawRegion === "object") {
      const region = rawRegion;
      const startRow = typeof region.startRow === "number" ? region.startRow : Number.NaN;
      const startCol = typeof region.startCol === "number" ? region.startCol : Number.NaN;
      const endRow = typeof region.endRow === "number" ? region.endRow : Number.NaN;
      const endCol = typeof region.endCol === "number" ? region.endCol : Number.NaN;
      if ([startRow, startCol, endRow, endCol].every((value) => Number.isFinite(value) && value >= 0)) {
        range = {
          end: {
            col: Math.max(startCol, endCol),
            row: Math.max(startRow, endRow)
          },
          start: {
            col: Math.min(startCol, endCol),
            row: Math.min(startRow, endRow)
          }
        };
      } else if (typeof region.range === "string") {
        range = parseA1RangeReference(region.range);
      }
    }
    if (!range) {
      continue;
    }
    if (range.end.col > range.start.col) {
      mergeMetadata.hasHorizontalMerges = true;
      mergeMetadata.maxHorizontalMergeEndCol = Math.max(mergeMetadata.maxHorizontalMergeEndCol, range.end.col);
    }
    if (range.end.row > range.start.row) {
      mergeMetadata.hasVerticalMerges = true;
      mergeMetadata.maxVerticalMergeEndRow = Math.max(mergeMetadata.maxVerticalMergeEndRow, range.end.row);
    }
  }
  return mergeMetadata;
}
function buildThemePalette(theme) {
  const themeOrder = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
  const colorsByIndex = {};
  themeOrder.forEach((key, index) => {
    const color = theme.colors.get(key);
    if (color) {
      colorsByIndex[index] = color;
    }
  });
  return {
    colorsByIndex,
    majorLatinFont: theme.majorLatinFont ?? void 0,
    minorLatinFont: theme.minorLatinFont ?? void 0
  };
}
function normalizeArchivePath2(path) {
  return path.replace(/\\/g, "/").replace(/^\/+/, "");
}
function joinArchivePath(...parts) {
  return normalizeArchivePath2(parts.join("/"));
}
function dirname2(path) {
  const normalized = normalizeArchivePath2(path);
  const lastSlash = normalized.lastIndexOf("/");
  return lastSlash >= 0 ? normalized.slice(0, lastSlash) : "";
}
function resolveArchiveTarget(baseDocumentPath, target) {
  if (!target) {
    return normalizeArchivePath2(baseDocumentPath);
  }
  if (target.startsWith("#")) {
    return target;
  }
  if (target.startsWith("/")) {
    return normalizeArchivePath2(target);
  }
  const baseParts = dirname2(baseDocumentPath).split("/").filter(Boolean);
  for (const segment of target.split("/")) {
    if (!segment || segment === ".") {
      continue;
    }
    if (segment === "..") {
      baseParts.pop();
      continue;
    }
    baseParts.push(segment);
  }
  return normalizeArchivePath2(baseParts.join("/"));
}
function relsPathForDocument(documentPath) {
  const baseName = documentPath.split("/").pop();
  const parentDir = dirname2(documentPath);
  return joinArchivePath(parentDir, "_rels", `${baseName}.rels`);
}
function parseXml2(xml) {
  const parser = new DOMParser();
  const document2 = parser.parseFromString(xml, "application/xml");
  if (document2.querySelector("parsererror")) {
    return null;
  }
  return document2;
}
function readArchiveText2(archive, path) {
  const entry = archive[path];
  return entry ? strFromU83(entry) : null;
}
function parseColumnReference(reference) {
  let value = 0;
  for (const character of reference.toUpperCase()) {
    if (character < "A" || character > "Z") {
      return null;
    }
    value = value * 26 + (character.charCodeAt(0) - 64);
  }
  return value > 0 ? value - 1 : null;
}
function parseA1CellReference(reference) {
  const match = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(reference.trim());
  if (!match) {
    return null;
  }
  const col = parseColumnReference(match[1] ?? "");
  const row = Number(match[2] ?? Number.NaN) - 1;
  if (col === null || !Number.isFinite(row) || row < 0) {
    return null;
  }
  return { col, row };
}
function parseA1RangeReference(reference) {
  const [startRef, endRef] = reference.split(":");
  const start = parseA1CellReference(startRef ?? "");
  const end = parseA1CellReference(endRef ?? startRef ?? "");
  return start && end ? { end, start } : null;
}
function stripSheetNameFromFormulaReference(reference) {
  const trimmed = reference.trim();
  const bangIndex = trimmed.lastIndexOf("!");
  return bangIndex >= 0 ? trimmed.slice(bangIndex + 1) : trimmed;
}
function parseFormulaCellReference(reference) {
  const normalized = stripSheetNameFromFormulaReference(reference).split(/\s+/)[0] ?? "";
  return parseA1CellReference(normalized);
}
function parseFormulaRangeReference(reference) {
  return parseA1RangeReference(stripSheetNameFromFormulaReference(reference));
}
function isElementNode2(node) {
  return Boolean(node && node.nodeType === 1);
}
function getLocalElements(parent, localName) {
  return Array.from(parent.getElementsByTagName("*")).filter((node) => isElementNode2(node) && node.localName === localName);
}
function getChildElements(parent, localName) {
  return Array.from(parent.childNodes).filter((node) => isElementNode2(node) && node.localName === localName);
}
function getFirstChild(parent, localName) {
  return getChildElements(parent, localName)[0] ?? null;
}
function getFirstDescendant(parent, localName) {
  return getLocalElements(parent, localName)[0] ?? null;
}
function readFeaturePropertyBagCheckboxComplements(archive) {
  const xml = readArchiveText2(archive, "xl/featurePropertyBag/featurePropertyBag.xml");
  if (!xml) {
    return /* @__PURE__ */ new Set();
  }
  const document2 = parseXml2(xml);
  if (!document2?.documentElement) {
    return /* @__PURE__ */ new Set();
  }
  const bagNodes = getChildElements(document2.documentElement, "bag");
  const bagTypeById = bagNodes.map((node) => node.getAttribute("type") ?? "");
  const checkboxComplementIndices = /* @__PURE__ */ new Set();
  const xfComplementsBag = bagNodes.find((node) => node.getAttribute("type") === "XFComplements") ?? null;
  const mappedBagIds = xfComplementsBag ? getLocalElements(xfComplementsBag, "bagId").map((node) => Number(node.textContent ?? Number.NaN)).filter((value) => Number.isFinite(value)) : [];
  mappedBagIds.forEach((bagId, complementIndex) => {
    const xfComplementBag = bagNodes[bagId];
    if (!xfComplementBag || bagTypeById[bagId] !== "XFComplement") {
      return;
    }
    const xfControlsBagId = getLocalElements(xfComplementBag, "bagId").map((node) => Number(node.textContent ?? Number.NaN)).find((value) => Number.isFinite(value));
    if (xfControlsBagId === void 0) {
      return;
    }
    const xfControlsBag = bagNodes[xfControlsBagId];
    if (!xfControlsBag || bagTypeById[xfControlsBagId] !== "XFControls") {
      return;
    }
    const cellControlBagId = getLocalElements(xfControlsBag, "bagId").map((node) => Number(node.textContent ?? Number.NaN)).find((value) => Number.isFinite(value));
    if (cellControlBagId === void 0) {
      return;
    }
    if (bagTypeById[cellControlBagId] === "Checkbox") {
      checkboxComplementIndices.add(complementIndex);
    }
  });
  return checkboxComplementIndices;
}
function getRelationshipId(element) {
  return element.getAttributeNS(REL_NS, "id") ?? element.getAttribute("r:id") ?? element.getAttribute("id");
}
function parseContentTypes(archive) {
  const xml = readArchiveText2(archive, "[Content_Types].xml");
  const defaultEntries = /* @__PURE__ */ new Map();
  const overrideEntries = /* @__PURE__ */ new Map();
  if (!xml) {
    return { defaultEntries, overrideEntries };
  }
  const document2 = parseXml2(xml);
  if (!document2) {
    return { defaultEntries, overrideEntries };
  }
  for (const defaultNode of getLocalElements(document2, "Default")) {
    const extension = defaultNode.getAttribute("Extension");
    const contentType = defaultNode.getAttribute("ContentType");
    if (extension && contentType) {
      defaultEntries.set(extension.toLowerCase(), contentType);
    }
  }
  for (const overrideNode of getLocalElements(document2, "Override")) {
    const partName = overrideNode.getAttribute("PartName");
    const contentType = overrideNode.getAttribute("ContentType");
    if (partName && contentType) {
      overrideEntries.set(normalizeArchivePath2(partName), contentType);
    }
  }
  return { defaultEntries, overrideEntries };
}
function parseRelationships(archive, relsPath, baseDocumentPath) {
  const xml = readArchiveText2(archive, relsPath);
  const relationships = /* @__PURE__ */ new Map();
  if (!xml) {
    return relationships;
  }
  const document2 = parseXml2(xml);
  if (!document2) {
    return relationships;
  }
  for (const relationshipNode of getLocalElements(document2, "Relationship")) {
    const id = relationshipNode.getAttribute("Id");
    const target = relationshipNode.getAttribute("Target");
    const type = relationshipNode.getAttribute("Type");
    if (!id || !target || !type) {
      continue;
    }
    relationships.set(id, {
      id,
      target: resolveArchiveTarget(baseDocumentPath, target),
      targetMode: relationshipNode.getAttribute("TargetMode"),
      type
    });
  }
  return relationships;
}
function parseWorkbookSheets(archive) {
  const workbookXml = readArchiveText2(archive, "xl/workbook.xml");
  if (!workbookXml) {
    return [];
  }
  const workbookDocument = parseXml2(workbookXml);
  if (!workbookDocument) {
    return [];
  }
  const workbookRelationships = parseRelationships(archive, "xl/_rels/workbook.xml.rels", "xl/workbook.xml");
  const sheets2 = [];
  for (const sheetNode of getLocalElements(workbookDocument, "sheet")) {
    const relationshipId = getRelationshipId(sheetNode);
    if (!relationshipId) {
      continue;
    }
    const relationship = workbookRelationships.get(relationshipId);
    if (!relationship) {
      continue;
    }
    sheets2.push({
      name: sheetNode.getAttribute("name") ?? `Sheet ${sheets2.length + 1}`,
      path: relationship.target
    });
  }
  return sheets2;
}
function parseWorkbookTheme(archive) {
  const defaultTheme = {
    colors: /* @__PURE__ */ new Map([
      ["accent1", "#5b9bd5"],
      ["accent2", "#ed7d31"],
      ["accent3", "#a5a5a5"],
      ["accent4", "#ffc000"],
      ["accent5", "#4472c4"],
      ["accent6", "#70ad47"],
      ["bg1", "#ffffff"],
      ["bg2", "#e7e6e6"],
      ["dk1", "#000000"],
      ["dk2", "#6e747a"],
      ["folHlink", "#993366"],
      ["hlink", "#085296"],
      ["lt1", "#ffffff"],
      ["lt2", "#e7e6e6"],
      ["tx1", "#000000"],
      ["tx2", "#6e747a"]
    ]),
    majorLatinFont: null,
    minorLatinFont: null
  };
  const themeXml = readArchiveText2(archive, "xl/theme/theme1.xml");
  if (!themeXml) {
    return defaultTheme;
  }
  const themeDocument = parseXml2(themeXml);
  if (!themeDocument) {
    return defaultTheme;
  }
  const colors = new Map(defaultTheme.colors);
  const colorSchemeNode = getLocalElements(themeDocument, "clrScheme")[0] ?? null;
  if (colorSchemeNode) {
    for (const colorNode of Array.from(colorSchemeNode.childNodes).filter(isElementNode2)) {
      const key = colorNode.localName;
      const srgbNode = getFirstChild(colorNode, "srgbClr");
      const sysNode = getFirstChild(colorNode, "sysClr");
      const hex = srgbNode?.getAttribute("val") ?? sysNode?.getAttribute("lastClr");
      if (hex) {
        colors.set(key, normalizeHexColor3(hex));
      }
    }
  }
  const fontSchemeNode = getLocalElements(themeDocument, "fontScheme")[0] ?? null;
  const majorLatinFont = getFirstChild(getFirstChild(fontSchemeNode, "majorFont"), "latin")?.getAttribute("typeface") ?? null;
  const minorLatinFont = getFirstChild(getFirstChild(fontSchemeNode, "minorFont"), "latin")?.getAttribute("typeface") ?? null;
  colors.set("bg1", colors.get("lt1") ?? defaultTheme.colors.get("bg1") ?? "#ffffff");
  colors.set("tx1", colors.get("dk1") ?? defaultTheme.colors.get("tx1") ?? "#000000");
  colors.set("bg2", colors.get("lt2") ?? defaultTheme.colors.get("bg2") ?? "#e7e6e6");
  colors.set("tx2", colors.get("dk2") ?? defaultTheme.colors.get("tx2") ?? "#6e747a");
  return {
    colors,
    majorLatinFont,
    minorLatinFont
  };
}
function parseSpreadsheetColor(node) {
  if (!node) {
    return void 0;
  }
  const color = {};
  const rgb = node.getAttribute("rgb");
  const theme = node.getAttribute("theme");
  const tint = node.getAttribute("tint");
  const indexed = node.getAttribute("indexed");
  if (rgb) {
    color.rgb = normalizeHexColor3(rgb);
  }
  if (theme !== null) {
    color.theme = Number(theme);
  }
  if (tint !== null) {
    color.tint = Number(tint);
  }
  if (indexed !== null) {
    color.indexed = Number(indexed);
  }
  return Object.keys(color).length > 0 ? color : void 0;
}
function hasEnabledSpreadsheetFlag(node) {
  if (!node) {
    return false;
  }
  const value = node.getAttribute("val");
  return value === null || value !== "0" && value !== "false";
}
function parseSheetSparklines(document2, themePalette) {
  const sparklines = [];
  for (const groupNode of getLocalElements(document2, "sparklineGroup")) {
    const rawType = groupNode.getAttribute("type");
    const sparklineType = rawType === "column" ? "column" : rawType === "stacked" ? "winLoss" : "line";
    const markersNode = getFirstChild(groupNode, "markers");
    const negativeNode = getFirstChild(groupNode, "negative");
    const colorSeries = resolveWorkbookColor(parseSpreadsheetColor(getFirstChild(groupNode, "colorSeries")), themePalette);
    const colorNegative = resolveWorkbookColor(parseSpreadsheetColor(getFirstChild(groupNode, "colorNegative")), themePalette);
    const colorMarkers = resolveWorkbookColor(parseSpreadsheetColor(getFirstChild(groupNode, "colorMarkers")), themePalette);
    const colorFirst = resolveWorkbookColor(parseSpreadsheetColor(getFirstChild(groupNode, "colorFirst")), themePalette);
    const colorLast = resolveWorkbookColor(parseSpreadsheetColor(getFirstChild(groupNode, "colorLast")), themePalette);
    const colorHigh = resolveWorkbookColor(parseSpreadsheetColor(getFirstChild(groupNode, "colorHigh")), themePalette);
    const colorLow = resolveWorkbookColor(parseSpreadsheetColor(getFirstChild(groupNode, "colorLow")), themePalette);
    const sparklineCollectionNode = getFirstChild(groupNode, "sparklines");
    if (!sparklineCollectionNode) {
      continue;
    }
    for (const sparklineNode of getChildElements(sparklineCollectionNode, "sparkline")) {
      const formula = getFirstChild(sparklineNode, "f")?.textContent ?? "";
      const targetReference = getFirstChild(sparklineNode, "sqref")?.textContent ?? "";
      const range = parseFormulaRangeReference(formula);
      const target = parseFormulaCellReference(targetReference);
      if (!range || !target) {
        continue;
      }
      sparklines.push({
        color: colorSeries ?? void 0,
        firstColor: colorFirst ?? void 0,
        highColor: colorHigh ?? void 0,
        lastColor: colorLast ?? void 0,
        lowColor: colorLow ?? void 0,
        markerColor: colorMarkers ?? void 0,
        markers: hasEnabledSpreadsheetFlag(markersNode),
        negative: hasEnabledSpreadsheetFlag(negativeNode),
        negativeColor: colorNegative ?? void 0,
        range,
        target,
        type: sparklineType
      });
    }
  }
  return sparklines;
}
function parseSpreadsheetFont(node) {
  if (!node) {
    return void 0;
  }
  const font = {};
  const size = getFirstChild(node, "sz")?.getAttribute("val");
  const name = getFirstChild(node, "name")?.getAttribute("val");
  const family = getFirstChild(node, "family")?.getAttribute("val");
  const scheme = getFirstChild(node, "scheme")?.getAttribute("val");
  const charset = getFirstChild(node, "charset")?.getAttribute("val");
  const verticalAlign = getFirstChild(node, "vertAlign")?.getAttribute("val");
  const color = parseSpreadsheetColor(getFirstChild(node, "color"));
  if (hasEnabledSpreadsheetFlag(getFirstChild(node, "b"))) {
    font.bold = true;
  }
  if (hasEnabledSpreadsheetFlag(getFirstChild(node, "i"))) {
    font.italic = true;
  }
  if (hasEnabledSpreadsheetFlag(getFirstChild(node, "strike"))) {
    font.strikethrough = true;
  }
  if (getFirstChild(node, "u")) {
    font.underline = getFirstChild(node, "u")?.getAttribute("val") ?? "single";
  }
  if (size !== null && size !== void 0) {
    font.size = Number(size);
  }
  if (name) {
    font.name = name;
  }
  if (family !== null && family !== void 0) {
    font.family = Number(family);
  }
  if (scheme) {
    font.scheme = scheme;
  }
  if (charset !== null && charset !== void 0) {
    font.charset = Number(charset);
  }
  if (verticalAlign) {
    font.verticalAlign = verticalAlign;
  }
  if (hasEnabledSpreadsheetFlag(getFirstChild(node, "shadow"))) {
    font.shadow = true;
  }
  if (hasEnabledSpreadsheetFlag(getFirstChild(node, "outline"))) {
    font.outline = true;
  }
  if (hasEnabledSpreadsheetFlag(getFirstChild(node, "condense"))) {
    font.condense = true;
  }
  if (hasEnabledSpreadsheetFlag(getFirstChild(node, "extend"))) {
    font.extend = true;
  }
  if (color) {
    font.color = color;
  }
  return Object.keys(font).length > 0 ? font : void 0;
}
function parseSpreadsheetFill(node) {
  if (!node) {
    return void 0;
  }
  const gradientFill = getFirstChild(node, "gradientFill");
  if (gradientFill) {
    const stops = Array.from(gradientFill.childNodes).filter(isElementNode2).filter((child) => child.localName === "stop").map((stopNode) => ({
      color: parseSpreadsheetColor(Array.from(stopNode.childNodes).find(isElementNode2) ?? null),
      position: Number(stopNode.getAttribute("position") ?? Number.NaN)
    })).filter((stop) => stop.color && Number.isFinite(stop.position));
    if (stops.length > 0) {
      return {
        degree: Number(gradientFill.getAttribute("degree") ?? 0),
        fillType: "gradient",
        gradientType: gradientFill.getAttribute("type") ?? "linear",
        stops
      };
    }
  }
  const patternFill = getFirstChild(node, "patternFill");
  if (!patternFill) {
    return void 0;
  }
  const patternType = patternFill.getAttribute("patternType") ?? "none";
  const foreground = parseSpreadsheetColor(getFirstChild(patternFill, "fgColor"));
  const background = parseSpreadsheetColor(getFirstChild(patternFill, "bgColor"));
  const solidColor = foreground ?? background;
  if (patternType === "solid" && solidColor) {
    return {
      color: solidColor,
      fillType: "solid"
    };
  }
  if ((patternType === "none" || patternType === "gray125") && (foreground || background)) {
    return {
      background,
      fillType: "pattern",
      foreground,
      patternType
    };
  }
  if (patternType !== "none" && patternType !== "gray125" && (foreground || background)) {
    return {
      background,
      fillType: "pattern",
      foreground,
      patternType
    };
  }
  return void 0;
}
function parseSpreadsheetBorderEdge(node) {
  if (!node) {
    return void 0;
  }
  const style = node.getAttribute("style");
  const color = parseSpreadsheetColor(getFirstChild(node, "color"));
  if (!style || style === "none") {
    return void 0;
  }
  return {
    color,
    style
  };
}
function parseSpreadsheetBorder(node) {
  if (!node) {
    return void 0;
  }
  const border = {};
  ["top", "right", "bottom", "left", "horizontal", "vertical"].forEach((edge) => {
    const parsedEdge = parseSpreadsheetBorderEdge(getFirstChild(node, edge));
    if (parsedEdge) {
      border[edge] = parsedEdge;
    }
  });
  return Object.keys(border).length > 0 ? border : void 0;
}
function parseSpreadsheetAlignment(node) {
  if (!node) {
    return void 0;
  }
  const alignment = {};
  const horizontal = node.getAttribute("horizontal");
  const vertical = node.getAttribute("vertical");
  const wrapText = node.getAttribute("wrapText");
  const indent = node.getAttribute("indent");
  const shrinkToFit = node.getAttribute("shrinkToFit");
  const textRotation = node.getAttribute("textRotation");
  if (horizontal) {
    alignment.horizontal = horizontal;
  }
  if (vertical) {
    alignment.vertical = vertical;
  }
  if (wrapText !== null) {
    alignment.wrapText = wrapText === "1";
  }
  if (shrinkToFit !== null) {
    alignment.shrinkToFit = shrinkToFit === "1";
  }
  if (indent !== null) {
    alignment.indent = Number(indent);
  }
  if (textRotation !== null) {
    const parsedRotation = Number(textRotation);
    if (Number.isFinite(parsedRotation)) {
      alignment.textRotation = parsedRotation;
    }
  }
  return Object.keys(alignment).length > 0 ? alignment : void 0;
}
function parseDifferentialStyle(node) {
  if (!node) {
    return {};
  }
  const style = {};
  const font = parseSpreadsheetFont(getFirstChild(node, "font"));
  const fill = parseSpreadsheetFill(getFirstChild(node, "fill"));
  const border = parseSpreadsheetBorder(getFirstChild(node, "border"));
  const alignment = parseSpreadsheetAlignment(getFirstChild(node, "alignment"));
  if (font) {
    style.font = font;
  }
  if (fill) {
    style.fill = fill;
  }
  if (border) {
    style.border = border;
  }
  if (alignment) {
    style.alignment = alignment;
  }
  return style;
}
function parseResolvedXfStyle(xfNode, fonts, fills, borders, checkboxComplementIndices) {
  const style = {};
  const fontId = Number(xfNode.getAttribute("fontId") ?? Number.NaN);
  const fillId = Number(xfNode.getAttribute("fillId") ?? Number.NaN);
  const borderId = Number(xfNode.getAttribute("borderId") ?? Number.NaN);
  const alignment = parseSpreadsheetAlignment(getFirstChild(xfNode, "alignment"));
  if (Number.isFinite(fontId) && fonts[fontId]) {
    style.font = fonts[fontId];
  }
  if (Number.isFinite(fillId) && fills[fillId]) {
    style.fill = fills[fillId];
  }
  if (Number.isFinite(borderId) && borders[borderId]) {
    style.border = borders[borderId];
  }
  if (alignment) {
    style.alignment = alignment;
  }
  const xfComplementNode = getFirstDescendant(xfNode, "xfComplement");
  const xfComplementIndex = Number(xfComplementNode?.getAttribute("i") ?? Number.NaN);
  if (Number.isFinite(xfComplementIndex) && checkboxComplementIndices?.has(xfComplementIndex)) {
    style.cellControl = { kind: "checkbox" };
  }
  return style;
}
function parseWorkbookStyles(archive) {
  const xml = readArchiveText2(archive, "xl/styles.xml");
  if (!xml) {
    return {
      differentialStyles: [],
      defaultFont: null,
      namedCellStyleByName: {},
      styleById: {},
      tableStyleByName: {}
    };
  }
  const document2 = parseXml2(xml);
  if (!document2) {
    return {
      differentialStyles: [],
      defaultFont: null,
      namedCellStyleByName: {},
      styleById: {},
      tableStyleByName: {}
    };
  }
  const fontsNode = getFirstDescendant(document2, "fonts");
  const fillsNode = getFirstDescendant(document2, "fills");
  const bordersNode = getFirstDescendant(document2, "borders");
  const cellStyleXfsNode = getFirstDescendant(document2, "cellStyleXfs");
  const cellStylesNode = getFirstDescendant(document2, "cellStyles");
  const cellXfsNode = getFirstDescendant(document2, "cellXfs");
  const dxfsNode = getFirstDescendant(document2, "dxfs");
  const tableStylesNode = getFirstDescendant(document2, "tableStyles");
  if (!cellXfsNode) {
    return {
      differentialStyles: [],
      defaultFont: null,
      namedCellStyleByName: {},
      styleById: {},
      tableStyleByName: {}
    };
  }
  const checkboxComplementIndices = readFeaturePropertyBagCheckboxComplements(archive);
  const fonts = getChildElements(fontsNode ?? document2.documentElement, "font").map((node) => parseSpreadsheetFont(node));
  const fills = getChildElements(fillsNode ?? document2.documentElement, "fill").map((node) => parseSpreadsheetFill(node));
  const borders = getChildElements(bordersNode ?? document2.documentElement, "border").map((node) => parseSpreadsheetBorder(node));
  const differentialStyles = getChildElements(dxfsNode ?? document2.documentElement, "dxf").map((node) => parseDifferentialStyle(node));
  const cellStyleXfs = getChildElements(cellStyleXfsNode ?? document2.documentElement, "xf").map(
    (node) => parseResolvedXfStyle(node, fonts, fills, borders, checkboxComplementIndices)
  );
  const namedCellStyleByName = {};
  const styleById = {};
  const tableStyleByName = {};
  getChildElements(cellXfsNode, "xf").forEach((xfNode, index) => {
    styleById[index] = parseResolvedXfStyle(xfNode, fonts, fills, borders, checkboxComplementIndices);
  });
  getChildElements(cellStylesNode ?? document2.documentElement, "cellStyle").forEach((cellStyleNode) => {
    const name = cellStyleNode.getAttribute("name");
    const xfId = Number(cellStyleNode.getAttribute("xfId") ?? Number.NaN);
    if (!name || !Number.isFinite(xfId)) {
      return;
    }
    const resolvedStyle = cellStyleXfs[xfId];
    if (resolvedStyle) {
      namedCellStyleByName[name] = resolvedStyle;
    }
  });
  getChildElements(tableStylesNode ?? document2.documentElement, "tableStyle").forEach((tableStyleNode) => {
    const name = tableStyleNode.getAttribute("name");
    if (!name) {
      return;
    }
    const elements = {};
    getChildElements(tableStyleNode, "tableStyleElement").forEach((elementNode) => {
      const type = elementNode.getAttribute("type");
      const dxfId = Number(elementNode.getAttribute("dxfId") ?? Number.NaN);
      if (!type || !Number.isFinite(dxfId)) {
        return;
      }
      const differentialStyle = differentialStyles[dxfId];
      if (differentialStyle) {
        elements[type] = differentialStyle;
      }
    });
    tableStyleByName[name] = elements;
  });
  const normalFont = namedCellStyleByName.Normal?.font ?? styleById[0]?.font ?? fonts[0];
  const defaultFont = normalFont ? {
    family: typeof normalFont.name === "string" ? normalFont.name : void 0,
    sizePt: typeof normalFont.size === "number" ? normalFont.size : void 0
  } : null;
  return {
    differentialStyles,
    defaultFont,
    namedCellStyleByName,
    styleById,
    tableStyleByName
  };
}
function parseSqrefRanges(sqref) {
  if (!sqref) {
    return [];
  }
  return sqref.trim().split(/\s+/).flatMap((reference) => {
    const range = parseA1RangeReference(reference);
    return range ? [range] : [];
  });
}
function parseConditionalFormatValueObject(node) {
  if (!node) {
    return null;
  }
  const type = node.getAttribute("type");
  if (!type) {
    return null;
  }
  const rawValue = node.getAttribute("val") ?? getFirstChild(node, "f")?.textContent ?? void 0;
  const numericValue = rawValue !== void 0 ? Number(rawValue) : Number.NaN;
  return {
    type,
    value: Number.isFinite(numericValue) ? numericValue : void 0
  };
}
function parseSpreadsheetBooleanAttribute(node, name) {
  if (!node) {
    return void 0;
  }
  const value = node.getAttribute(name);
  if (value === null) {
    return void 0;
  }
  return value !== "0" && value !== "false";
}
function parseStandardConditionalFormatRule(cfRuleNode, ranges, differentialStyles = []) {
  const type = cfRuleNode.getAttribute("type");
  const rawPriority = Number(cfRuleNode.getAttribute("priority") ?? Number.NaN);
  const priority = Number.isFinite(rawPriority) ? rawPriority : Number.MAX_SAFE_INTEGER;
  const formulas = getChildElements(cfRuleNode, "formula").map((formulaNode) => (formulaNode.textContent ?? "").trim()).filter((formula) => formula.length > 0);
  if (type === "colorScale") {
    const colorScaleNode = getFirstChild(cfRuleNode, "colorScale");
    if (!colorScaleNode) {
      return null;
    }
    const cfvos = getChildElements(colorScaleNode, "cfvo").map((node) => parseConditionalFormatValueObject(node)).filter((value) => Boolean(value));
    const colors = getChildElements(colorScaleNode, "color").map((node) => parseSpreadsheetColor(node)).filter((value) => Boolean(value));
    if (cfvos.length === 0 || colors.length === 0) {
      return null;
    }
    return {
      cfvos,
      colors,
      kind: "colorScale",
      priority,
      ranges
    };
  }
  if (type === "dataBar") {
    const dataBarNode = getFirstChild(cfRuleNode, "dataBar");
    if (!dataBarNode) {
      return null;
    }
    const cfvos = getChildElements(dataBarNode, "cfvo").map((node) => parseConditionalFormatValueObject(node)).filter((value) => Boolean(value));
    if (cfvos.length === 0) {
      return null;
    }
    const extId = getFirstDescendant(cfRuleNode, "id")?.textContent?.trim() || void 0;
    return {
      cfvos,
      color: parseSpreadsheetColor(getFirstChild(dataBarNode, "color")),
      kind: "dataBar",
      priority,
      ranges,
      id: extId
    };
  }
  if (type === "iconSet") {
    const iconSetNode = getFirstChild(cfRuleNode, "iconSet");
    if (!iconSetNode) {
      return null;
    }
    const iconSetName = iconSetNode.getAttribute("iconSet");
    const cfvos = getChildElements(iconSetNode, "cfvo").map((node) => parseConditionalFormatValueObject(node)).filter((value) => Boolean(value));
    if (!iconSetName || cfvos.length === 0) {
      return null;
    }
    return {
      cfvos,
      icons: cfvos.map((_, index) => ({
        iconId: index,
        iconSet: iconSetName
      })),
      kind: "iconSet",
      priority,
      ranges,
      reverse: parseSpreadsheetBooleanAttribute(iconSetNode, "reverse"),
      showValue: parseSpreadsheetBooleanAttribute(iconSetNode, "showValue")
    };
  }
  const rawDxfId = Number(cfRuleNode.getAttribute("dxfId") ?? Number.NaN);
  if (!type || !Number.isFinite(rawDxfId)) {
    return null;
  }
  const style = differentialStyles[rawDxfId];
  if (!style) {
    return null;
  }
  const rawRank = Number(cfRuleNode.getAttribute("rank") ?? Number.NaN);
  const rawStdDev = Number(cfRuleNode.getAttribute("stdDev") ?? Number.NaN);
  return {
    aboveAverage: parseSpreadsheetBooleanAttribute(cfRuleNode, "aboveAverage"),
    bottom: parseSpreadsheetBooleanAttribute(cfRuleNode, "bottom"),
    equalAverage: parseSpreadsheetBooleanAttribute(cfRuleNode, "equalAverage"),
    formulas,
    kind: "styled",
    operator: cfRuleNode.getAttribute("operator") ?? void 0,
    percent: parseSpreadsheetBooleanAttribute(cfRuleNode, "percent"),
    priority,
    rank: Number.isFinite(rawRank) ? rawRank : void 0,
    ranges,
    ruleType: type,
    stdDev: Number.isFinite(rawStdDev) ? rawStdDev : void 0,
    stopIfTrue: parseSpreadsheetBooleanAttribute(cfRuleNode, "stopIfTrue"),
    style,
    text: cfRuleNode.getAttribute("text") ?? void 0,
    timePeriod: cfRuleNode.getAttribute("timePeriod") ?? void 0
  };
}
function parseExtendedConditionalFormatRule(cfRuleNode, ranges) {
  const type = cfRuleNode.getAttribute("type");
  const ruleId = cfRuleNode.getAttribute("id") ?? void 0;
  const rawPriority = Number(cfRuleNode.getAttribute("priority") ?? Number.NaN);
  const priority = Number.isFinite(rawPriority) ? rawPriority : Number.MAX_SAFE_INTEGER;
  if (type === "dataBar") {
    const dataBarNode = getFirstChild(cfRuleNode, "dataBar");
    if (!dataBarNode) {
      return null;
    }
    const cfvos = getChildElements(dataBarNode, "cfvo").map((node) => parseConditionalFormatValueObject(node)).filter((value) => Boolean(value));
    if (cfvos.length === 0) {
      return null;
    }
    return {
      axisColor: parseSpreadsheetColor(getFirstChild(dataBarNode, "axisColor")),
      border: parseSpreadsheetBooleanAttribute(dataBarNode, "border"),
      borderColor: parseSpreadsheetColor(getFirstChild(dataBarNode, "borderColor")),
      cfvos,
      color: parseSpreadsheetColor(getFirstChild(dataBarNode, "fillColor")),
      gradient: parseSpreadsheetBooleanAttribute(dataBarNode, "gradient"),
      kind: "dataBar",
      maxLength: Number(dataBarNode.getAttribute("maxLength") ?? Number.NaN),
      minLength: Number(dataBarNode.getAttribute("minLength") ?? Number.NaN),
      negativeBarBorderColorSameAsPositive: parseSpreadsheetBooleanAttribute(dataBarNode, "negativeBarBorderColorSameAsPositive"),
      negativeBorderColor: parseSpreadsheetColor(getFirstChild(dataBarNode, "negativeBorderColor")),
      negativeFillColor: parseSpreadsheetColor(getFirstChild(dataBarNode, "negativeFillColor")),
      priority,
      ranges,
      showValue: parseSpreadsheetBooleanAttribute(dataBarNode, "showValue"),
      id: ruleId
    };
  }
  if (type === "iconSet") {
    const iconSetNode = getFirstChild(cfRuleNode, "iconSet");
    if (!iconSetNode) {
      return null;
    }
    const cfvos = getChildElements(iconSetNode, "cfvo").map((node) => parseConditionalFormatValueObject(node)).filter((value) => Boolean(value));
    const icons = getChildElements(iconSetNode, "cfIcon").map((iconNode) => {
      const iconSet = iconNode.getAttribute("iconSet");
      const rawIconId = Number(iconNode.getAttribute("iconId") ?? Number.NaN);
      if (!iconSet || !Number.isFinite(rawIconId)) {
        return null;
      }
      return {
        iconId: rawIconId,
        iconSet
      };
    }).filter((icon) => Boolean(icon));
    if (cfvos.length === 0 || icons.length === 0) {
      return null;
    }
    return {
      cfvos,
      icons,
      kind: "iconSet",
      priority,
      ranges,
      reverse: parseSpreadsheetBooleanAttribute(iconSetNode, "reverse"),
      showValue: parseSpreadsheetBooleanAttribute(iconSetNode, "showValue"),
      id: ruleId
    };
  }
  return null;
}
function mergeConditionalFormatRule(baseRule, extendedRule) {
  if (baseRule.kind !== extendedRule.kind) {
    return baseRule;
  }
  if (baseRule.kind === "colorScale" && extendedRule.kind === "colorScale") {
    return {
      ...baseRule,
      ...extendedRule,
      cfvos: extendedRule.cfvos.length > 0 ? extendedRule.cfvos : baseRule.cfvos,
      colors: extendedRule.colors.length > 0 ? extendedRule.colors : baseRule.colors,
      priority: Number.isFinite(extendedRule.priority) ? extendedRule.priority : baseRule.priority,
      ranges: extendedRule.ranges.length > 0 ? extendedRule.ranges : baseRule.ranges
    };
  }
  if (baseRule.kind === "dataBar" && extendedRule.kind === "dataBar") {
    const merged = {
      ...baseRule,
      ...extendedRule,
      axisColor: extendedRule.axisColor ?? baseRule.axisColor,
      border: extendedRule.border ?? baseRule.border,
      cfvos: extendedRule.cfvos.length > 0 ? extendedRule.cfvos : baseRule.cfvos,
      color: extendedRule.color ?? baseRule.color,
      negativeBarBorderColorSameAsPositive: extendedRule.negativeBarBorderColorSameAsPositive ?? baseRule.negativeBarBorderColorSameAsPositive,
      negativeBorderColor: extendedRule.negativeBorderColor ?? baseRule.negativeBorderColor,
      negativeFillColor: extendedRule.negativeFillColor ?? baseRule.negativeFillColor,
      priority: Number.isFinite(extendedRule.priority) ? extendedRule.priority : baseRule.priority,
      ranges: extendedRule.ranges.length > 0 ? extendedRule.ranges : baseRule.ranges
    };
    return merged;
  }
  if (baseRule.kind === "iconSet" && extendedRule.kind === "iconSet") {
    const merged = {
      ...baseRule,
      ...extendedRule,
      cfvos: extendedRule.cfvos.length > 0 ? extendedRule.cfvos : baseRule.cfvos,
      icons: extendedRule.icons.length > 0 ? extendedRule.icons : baseRule.icons,
      priority: Number.isFinite(extendedRule.priority) ? extendedRule.priority : baseRule.priority,
      ranges: extendedRule.ranges.length > 0 ? extendedRule.ranges : baseRule.ranges
    };
    return merged;
  }
  return baseRule;
}
function parseConditionalFormatRules(document2, differentialStyles = []) {
  const standardRules = [];
  const extendedRules = [];
  getLocalElements(document2, "conditionalFormatting").forEach((conditionalFormattingNode) => {
    const isExtended = conditionalFormattingNode.namespaceURI !== SPREADSHEET_NS;
    const ranges = isExtended ? parseSqrefRanges(getFirstChild(conditionalFormattingNode, "sqref")?.textContent ?? "") : parseSqrefRanges(conditionalFormattingNode.getAttribute("sqref"));
    getChildElements(conditionalFormattingNode, "cfRule").forEach((cfRuleNode) => {
      const parsedRule = isExtended ? parseExtendedConditionalFormatRule(cfRuleNode, ranges) : parseStandardConditionalFormatRule(cfRuleNode, ranges, differentialStyles);
      if (parsedRule) {
        if (isExtended) {
          extendedRules.push(parsedRule);
        } else {
          standardRules.push(parsedRule);
        }
      }
    });
  });
  const mergedRules = [];
  const usedExtendedRuleIds = /* @__PURE__ */ new Set();
  const extendedRulesById = new Map(
    extendedRules.filter((rule) => typeof rule.id === "string" && rule.id.length > 0).map((rule) => [rule.id, rule])
  );
  standardRules.forEach((rule) => {
    const matchingExtendedRule = rule.id ? extendedRulesById.get(rule.id) : void 0;
    if (matchingExtendedRule) {
      usedExtendedRuleIds.add(rule.id);
      mergedRules.push(mergeConditionalFormatRule(rule, matchingExtendedRule));
      return;
    }
    mergedRules.push(rule);
  });
  extendedRules.forEach((rule) => {
    if (rule.id && usedExtendedRuleIds.has(rule.id)) {
      return;
    }
    mergedRules.push(rule);
  });
  return mergedRules.map((rule) => {
    const nextRule = { ...rule };
    delete nextRule.id;
    return nextRule;
  }).filter((rule) => rule.ranges.length > 0).sort((left, right) => left.priority - right.priority);
}
function parseSheetState(archive, path, options) {
  const xml = readArchiveText2(archive, path);
  if (!xml) {
    return null;
  }
  const document2 = parseXml2(xml);
  if (!document2) {
    return null;
  }
  const includeCachedFormulaValues = options?.includeCachedFormulaValues ?? true;
  const autoFilterRanges = parseSqrefRanges(getLocalElements(document2, "autoFilter")[0]?.getAttribute("ref"));
  const cachedFormulaValues = {};
  const conditionalFormatRules = parseConditionalFormatRules(document2, options?.differentialStyles ?? []);
  const sparklines = parseSheetSparklines(document2, options?.themePalette);
  const sheetFormatNode = getLocalElements(document2, "sheetFormatPr")[0] ?? null;
  const sheetViewNode = getLocalElements(document2, "sheetView")[0] ?? null;
  const rowHeightOverridesPx = {};
  const colWidthOverridesPx = {};
  const rowStyleIds = {};
  const colStyleIds = {};
  let minContentCol = Number.POSITIVE_INFINITY;
  let minContentRow = Number.POSITIVE_INFINITY;
  let maxContentCol = -1;
  let maxContentRow = -1;
  const columnWidthCharacterWidthPx = measureColumnCharacterWidthPx(
    options?.defaultFont?.family,
    options?.defaultFont?.sizePt
  );
  const defaultRowHeight = Number(sheetFormatNode?.getAttribute("defaultRowHeight") ?? 15);
  const defaultColWidth = Number(
    sheetFormatNode?.getAttribute("defaultColWidth") ?? sheetFormatNode?.getAttribute("baseColWidth") ?? 8.43
  );
  const rawZoomScale = Number(
    sheetViewNode?.getAttribute("zoomScale") ?? sheetViewNode?.getAttribute("zoomScaleNormal") ?? Number.NaN
  );
  const zoomScale = Number.isFinite(rawZoomScale) && rawZoomScale > 0 ? rawZoomScale : 100;
  const trackContentCell = (cellRef) => {
    if (!cellRef) {
      return;
    }
    const cell = parseA1CellReference(cellRef);
    if (!cell) {
      return;
    }
    minContentCol = Math.min(minContentCol, cell.col);
    minContentRow = Math.min(minContentRow, cell.row);
    maxContentCol = Math.max(maxContentCol, cell.col);
    maxContentRow = Math.max(maxContentRow, cell.row);
  };
  const isMeaningfulCellNode = (cellNode) => {
    if (getFirstChild(cellNode, "f") || getFirstChild(cellNode, "is")) {
      return true;
    }
    const valueNode = getFirstChild(cellNode, "v");
    return Boolean(valueNode && (valueNode.textContent ?? "").length > 0);
  };
  getLocalElements(document2, "row").forEach((rowNode) => {
    const rowIndex = Number(rowNode.getAttribute("r") ?? 0) - 1;
    const height = Number(rowNode.getAttribute("ht") ?? Number.NaN);
    const styleId = Number(rowNode.getAttribute("s") ?? Number.NaN);
    if (rowIndex >= 0 && Number.isFinite(height)) {
      rowHeightOverridesPx[rowIndex] = Math.max(MIN_ROW_HEIGHT_PX, Math.round(height * 1.33));
    }
    if (rowIndex >= 0 && Number.isFinite(styleId)) {
      rowStyleIds[rowIndex] = styleId;
    }
    getChildElements(rowNode, "c").forEach((cellNode) => {
      const cellRef = cellNode.getAttribute("r");
      if (isMeaningfulCellNode(cellNode)) {
        trackContentCell(cellRef);
      }
      if (includeCachedFormulaValues) {
        const formulaNode = getFirstChild(cellNode, "f");
        const valueNode = getFirstChild(cellNode, "v");
        if (formulaNode && valueNode && cellRef) {
          cachedFormulaValues[cellRef] = valueNode.textContent ?? "";
        }
      }
    });
  });
  const maxMetadataCol = Math.max(maxContentCol, 0) + 256;
  getLocalElements(document2, "col").forEach((colNode) => {
    const min = Number(colNode.getAttribute("min") ?? 0) - 1;
    const max = Number(colNode.getAttribute("max") ?? 0) - 1;
    const width = Number(colNode.getAttribute("width") ?? Number.NaN);
    const styleId = Number(colNode.getAttribute("style") ?? Number.NaN);
    if (!Number.isFinite(width)) {
      if (!Number.isFinite(styleId)) {
        return;
      }
    }
    for (let col = min; col <= Math.min(max, maxMetadataCol); col += 1) {
      if (col >= 0) {
        if (Number.isFinite(width)) {
          const widthPx = sheetColumnWidthToPixels(width, columnWidthCharacterWidthPx);
          colWidthOverridesPx[col] = widthPx;
        }
        if (Number.isFinite(styleId)) {
          colStyleIds[col] = styleId;
        }
      }
    }
  });
  return {
    autoFilterRanges,
    cachedFormulaValues,
    columnWidthCharacterWidthPx,
    colWidthOverridesPx,
    colStyleIds,
    conditionalFormatRules,
    defaultColWidthPx: sheetColumnWidthToPixels(defaultColWidth, columnWidthCharacterWidthPx),
    defaultRowHeightPx: Math.max(MIN_ROW_HEIGHT_PX, Math.round(defaultRowHeight * 1.33)),
    hasHorizontalMerges: false,
    hasVerticalMerges: false,
    maxHorizontalMergeEndCol: -1,
    maxVerticalMergeEndRow: -1,
    maxContentCol,
    maxContentRow,
    minContentCol: Number.isFinite(minContentCol) ? minContentCol : -1,
    minContentRow: Number.isFinite(minContentRow) ? minContentRow : -1,
    hiddenCols: [],
    hiddenRows: [],
    rowHeightOverridesPx,
    rowStyleIds,
    showGridLines: (sheetViewNode?.getAttribute("showGridLines") ?? "1") !== "0",
    sparklines,
    zoomScale
  };
}
function normalizeHexColor3(value) {
  const hex = value.replace(/^#/, "");
  if (hex.length === 8) {
    return `#${hex.slice(2).toLowerCase()}`;
  }
  if (hex.length === 6) {
    return `#${hex.toLowerCase()}`;
  }
  return "#000000";
}
function dukeDrawingAnchorToXlsxAnchor(anchor) {
  if (anchor.type === "absolute") {
    return {
      kind: "absolute",
      positionEmu: { x: anchor.xEmu, y: anchor.yEmu },
      sizeEmu: { cx: anchor.widthEmu, cy: anchor.heightEmu }
    };
  }
  const from = {
    col: anchor.from.col,
    colOffsetEmu: anchor.from.colOffsetEmu ?? 0,
    row: anchor.from.row,
    rowOffsetEmu: anchor.from.rowOffsetEmu ?? 0
  };
  if (anchor.type === "oneCell") {
    return {
      from,
      kind: "one-cell",
      sizeEmu: { cx: anchor.widthEmu, cy: anchor.heightEmu }
    };
  }
  return {
    from,
    kind: "two-cell",
    to: {
      col: anchor.to.col,
      colOffsetEmu: anchor.to.colOffsetEmu ?? 0,
      row: anchor.to.row,
      rowOffsetEmu: anchor.to.rowOffsetEmu ?? 0
    }
  };
}
function mapDukeDrawingColor(color) {
  if (!color) {
    return void 0;
  }
  switch (color.colorType) {
    case "auto":
      return { colorType: "auto" };
    case "rgb":
      return { b: color.b, colorType: "rgb", g: color.g, r: color.r };
    case "argb":
      return { a: color.a, b: color.b, colorType: "argb", g: color.g, r: color.r };
    case "theme":
      return { colorType: "theme", themeIndex: color.index, tint: color.tint };
    case "indexed":
      return { colorType: "indexed", paletteIndex: color.index };
  }
}
function mapDukeDrawingFont(font) {
  if (!font) {
    return void 0;
  }
  return {
    bold: font.bold,
    charset: font.charset,
    color: mapDukeDrawingColor(font.color),
    family: font.family,
    italic: font.italic,
    name: font.name,
    scheme: font.scheme,
    size: font.size,
    strikethrough: font.strikethrough,
    underline: font.underline,
    verticalAlign: font.verticalAlign
  };
}
function mapDukeDrawingText(text) {
  return {
    horizontalAlignment: text.horizontalAlignment,
    runs: text.runs.map((run) => ({ font: mapDukeDrawingFont(run.font), text: run.text })),
    verticalAlignment: text.verticalAlignment
  };
}
function resolveDukeDrawingColor(color, themePalette) {
  if (!color) {
    return void 0;
  }
  switch (color.colorType) {
    case "rgb":
      return resolveWorkbookColor({ rgb: [color.r, color.g, color.b].map((value) => value.toString(16).padStart(2, "0")).join("") }) ?? void 0;
    case "argb":
      return resolveWorkbookColor({ argb: [color.a, color.r, color.g, color.b].map((value) => value.toString(16).padStart(2, "0")).join("") }) ?? void 0;
    case "theme":
      return resolveWorkbookColor({ theme: color.index, tint: color.tint }, themePalette) ?? void 0;
    default:
      return void 0;
  }
}
function resolveDukeDrawingText(text, themePalette) {
  const caption = mapDukeDrawingText(text);
  const label = normalizeControlLabel(text.runs.map((run) => run.text).join(""));
  const firstStyledRun = text.runs.find((run) => run.font)?.font;
  const horizontalAlignment = text.horizontalAlignment;
  return {
    caption,
    fontFamily: firstStyledRun?.name,
    fontSizePt: firstStyledRun?.size,
    label,
    textAlign: horizontalAlignment === "left" || horizontalAlignment === "center" || horizontalAlignment === "right" ? horizontalAlignment : void 0,
    textColor: resolveDukeDrawingColor(firstStyledRun?.color, themePalette)
  };
}
function mapDukeFormControlKind(kind, themePalette) {
  switch (kind.kind) {
    case "button":
      return { ...resolveDukeDrawingText(kind.caption, themePalette), kind: "button" };
    case "checkbox":
      return {
        ...resolveDukeDrawingText(kind.caption, themePalette),
        checked: kind.state === "checked",
        kind: "checkbox",
        linkedCell: kind.cellLink,
        no3D: kind.no3D,
        state: kind.state
      };
    case "optionButton":
      return {
        ...resolveDukeDrawingText(kind.caption, themePalette),
        checked: kind.state === "checked",
        firstInGroup: kind.firstInGroup,
        kind: "radio",
        linkedCell: kind.cellLink,
        no3D: kind.no3D,
        state: kind.state
      };
    case "label":
      return { ...resolveDukeDrawingText(kind.caption, themePalette), kind: "label" };
    case "groupBox":
      return {
        ...resolveDukeDrawingText(kind.caption, themePalette),
        kind: "group-box",
        no3D: kind.no3D
      };
    case "listBox":
      return {
        inputRange: kind.inputRange,
        kind: "listbox",
        linkedCell: kind.cellLink,
        no3D: kind.no3D,
        selected: [...kind.selected],
        selection: kind.selection
      };
    case "dropdown":
      return {
        inputRange: kind.inputRange,
        kind: "dropdown",
        lines: kind.lines,
        linkedCell: kind.cellLink,
        no3D: kind.no3D,
        selected: kind.selected
      };
    case "scrollbar":
      return {
        horizontal: kind.horizontal,
        increment: kind.increment,
        kind: "scrollbar",
        linkedCell: kind.cellLink,
        max: kind.max,
        min: kind.min,
        page: kind.page,
        value: kind.value
      };
    case "spinner":
      return {
        increment: kind.increment,
        kind: "spinner",
        linkedCell: kind.cellLink,
        max: kind.max,
        min: kind.min,
        value: kind.value
      };
    case "unknown":
      return {
        ...resolveDukeDrawingText(kind.caption, themePalette),
        kind: kind.objectType.toLowerCase() === "editbox" ? "editbox" : "unknown",
        legacyObjectType: kind.legacyObjectType,
        objectType: kind.objectType
      };
  }
}
function mapDukeFormControl(control, controlIndex, workbookSheetIndex, themePalette) {
  return {
    altText: control.altText,
    anchor: dukeDrawingAnchorToXlsxAnchor(control.anchor),
    controlIndex,
    editAs: control.anchor.type === "twoCell" ? control.anchor.editAs : void 0,
    hidden: control.hidden,
    id: `form-control-${workbookSheetIndex}-${controlIndex}`,
    locked: control.locked,
    macroName: control.formControl.macroName,
    name: control.name,
    printable: control.printable,
    rawClientData: control.formControl.rawClientData,
    rawObj: control.formControl.rawObj,
    rawProperties: control.formControl.rawProperties,
    sheetIndex: workbookSheetIndex,
    title: control.title,
    workbookSheetIndex,
    zIndex: (control.drawingPath[0] ?? controlIndex) + 1,
    ...mapDukeFormControlKind(control.formControl.kind, themePalette)
  };
}
function collectWorkbookFormControls(workbook2, themePalette) {
  return Array.from({ length: workbook2.sheetCount }, (_, workbookSheetIndex) => {
    try {
      const controls = workbook2.getSheet(workbookSheetIndex).formControls;
      return Array.isArray(controls) ? controls.flatMap((control, controlIndex) => control.anchor ? [mapDukeFormControl(control, controlIndex, workbookSheetIndex, themePalette)] : []) : [];
    } catch {
      return [];
    }
  });
}
function normalizeControlLabel(label) {
  if (!label) {
    return void 0;
  }
  const normalized = label.replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
  return normalized.length > 0 ? normalized : void 0;
}
function parseWorkbookStructureAssetsFromArchive(archive, options) {
  const contentTypes = parseContentTypes(archive);
  const workbookSheets = parseWorkbookSheets(archive);
  const theme = parseWorkbookTheme(archive);
  const themePalette = buildThemePalette(theme);
  const { defaultFont, differentialStyles, namedCellStyleByName, styleById, tableStyleByName } = parseWorkbookStyles(archive);
  return {
    contentTypes,
    namedCellStyleByName,
    sheetStatesByWorkbookSheetIndex: workbookSheets.map((sheet) => parseSheetState(archive, sheet.path, {
      ...options,
      differentialStyles,
      defaultFont,
      themePalette
    })),
    styleById,
    tableMetadataByWorkbookSheetIndex: workbookSheets.map(() => []),
    tableStyleByName,
    theme,
    themePalette,
    workbookSheets
  };
}
function parseWorkbookStructureAssets(bytes, options) {
  const archive = unzipSync2(bytes);
  const {
    namedCellStyleByName,
    sheetStatesByWorkbookSheetIndex,
    styleById,
    tableMetadataByWorkbookSheetIndex,
    tableStyleByName,
    themePalette
  } = parseWorkbookStructureAssetsFromArchive(archive, options);
  return {
    namedCellStyleByName,
    sheetStatesByWorkbookSheetIndex,
    styleById,
    tableMetadataByWorkbookSheetIndex,
    tableStyleByName,
    themePalette
  };
}
function parseWorkbookChartStyleAssets(bytes) {
  const archive = unzipSync2(bytes);
  const {
    themePalette,
    workbookSheets
  } = parseWorkbookStructureAssetsFromArchive(archive);
  const sheetOrigins = [];
  workbookSheets.forEach((sheet, workbookSheetIndex) => {
    const sheetRelationships = parseRelationships(archive, relsPathForDocument(sheet.path), sheet.path);
    const attachments = [];
    for (const relationship of sheetRelationships.values()) {
      if (relationship.type !== DRAWING_REL_TYPE) {
        continue;
      }
      const drawingPath = relationship.target;
      const drawingRelsPath = relsPathForDocument(drawingPath);
      attachments.push({
        drawingPath,
        drawingRelsPath: archive[drawingRelsPath] ? drawingRelsPath : null,
        mediaPaths: []
      });
    }
    sheetOrigins[workbookSheetIndex] = attachments.length > 0 ? {
      attachments,
      workbookSheetIndex
    } : null;
  });
  return {
    archive,
    sheetOrigins,
    themePalette
  };
}
function resolveSheetColumnWidthPixels(width, columnWidthCharacterWidthPx) {
  return sheetColumnWidthToPixels(width, columnWidthCharacterWidthPx);
}

// src/external-fn.ts
var KEY_SEP = String.fromCharCode(1);
function externalCallKey(name, args) {
  return [name, ...args].join(KEY_SEP);
}
function makeExternalFn(values) {
  return (name, args) => {
    const value = values[externalCallKey(name, args)];
    return value === void 0 ? null : value;
  };
}
function externalCalcOptions(values) {
  return values ? { externalFnFn: makeExternalFn(values) } : void 0;
}

// src/safe-calculate.ts
var AUTO_CALCULATE_FORMULA_THRESHOLD = 1e3;
var SHEET_REF_REGEX = /'((?:[^']|'')+)'!|(?<![#\w.\u0080-\uFFFF])([A-Za-z_\u0080-\uFFFF][\w.\u0080-\uFFFF]*)!/g;
function collectReferencedSheetNames(workbook2) {
  const referenced = /* @__PURE__ */ new Set();
  for (let sheetIdx = 0; sheetIdx < workbook2.sheetCount; sheetIdx += 1) {
    let sheet;
    try {
      sheet = workbook2.getSheet(sheetIdx);
    } catch {
      continue;
    }
    const cells = sheet.formulaCells;
    if (!Array.isArray(cells)) {
      continue;
    }
    for (const cell of cells) {
      const formula = cell?.formula;
      if (!formula) {
        continue;
      }
      SHEET_REF_REGEX.lastIndex = 0;
      let match;
      while ((match = SHEET_REF_REGEX.exec(formula)) !== null) {
        const raw = match[1] ?? match[2];
        if (!raw) {
          continue;
        }
        referenced.add(raw.replace(/''/g, "'"));
      }
    }
  }
  return referenced;
}
function hasUnresolvedSheetReferences(workbook2) {
  let names;
  try {
    names = workbook2.sheetNames;
  } catch {
    return false;
  }
  const known = new Set(names);
  const referenced = collectReferencedSheetNames(workbook2);
  for (const name of referenced) {
    if (!known.has(name)) {
      return true;
    }
  }
  return false;
}
function countWorkbookFormulas(workbook2) {
  let total = 0;
  for (let index = 0; index < workbook2.sheetCount; index += 1) {
    total += workbook2.getSheet(index).formulaCount;
  }
  return total;
}
function safeCalculate(workbook2, options = {}) {
  const formulaCount = countWorkbookFormulas(workbook2);
  const sourceFormulaCount2 = options.sourceFormulaCount ?? null;
  const skipped = (reason, nextWorkbook = workbook2) => ({
    workbook: nextWorkbook,
    calculated: false,
    skipReason: reason,
    calculation: calculationReport(reason === "calculate-trapped" ? "failed" : "skipped", reason, formulaCount, sourceFormulaCount2)
  });
  if (hasUnresolvedSheetReferences(workbook2)) {
    return skipped("unresolved-sheet-refs");
  }
  try {
    const start = performance.now();
    const stats = workbook2.calculate(options.calcOptions);
    return {
      workbook: workbook2,
      calculated: true,
      skipReason: null,
      calculation: inspectCalculation(workbook2, stats, formulaCount, sourceFormulaCount2, performance.now() - start)
    };
  } catch (err) {
    console.warn("[react-xlsx] workbook.calculate() trapped; falling back to cached formula values", err);
    if (options.reparse) {
      try {
        return skipped("calculate-trapped", options.reparse());
      } catch (reparseErr) {
        console.warn("[react-xlsx] workbook reparse after calculate trap failed", reparseErr);
      }
    }
    return skipped("calculate-trapped");
  }
}

// src/wasm.ts
var wasmModulePromise = null;
var hasConfiguredWasmSource = false;
var configuredWasmSource;
var configuredWorkerWasmSource;
function bufferSourceToArrayBuffer(source) {
  if (source instanceof ArrayBuffer) {
    return source.slice(0);
  }
  const bytes = new Uint8Array(source.buffer, source.byteOffset, source.byteLength);
  const copy = new Uint8Array(bytes);
  return copy.buffer;
}
function sourceToWorkerSource(source) {
  if (typeof source === "string") {
    return source;
  }
  if (typeof URL !== "undefined" && source instanceof URL) {
    return source.href;
  }
  if (typeof Request !== "undefined" && source instanceof Request) {
    return source.url;
  }
  if (source instanceof ArrayBuffer || ArrayBuffer.isView(source)) {
    return bufferSourceToArrayBuffer(source);
  }
  if (typeof WebAssembly !== "undefined" && source instanceof WebAssembly.Module) {
    return source;
  }
  return void 0;
}
function setWasmSource(source) {
  hasConfiguredWasmSource = true;
  configuredWasmSource = source;
  configuredWorkerWasmSource = sourceToWorkerSource(source);
}
function getSheetsWasmModule() {
  if (!wasmModulePromise) {
    wasmModulePromise = import("@dukelib/sheets-wasm").then(async (mod) => {
      if (configuredWasmSource !== void 0) {
        await mod.default({ module_or_path: configuredWasmSource });
      } else {
        await mod.default();
      }
      return mod;
    });
  }
  return wasmModulePromise;
}

// src/xlsx-worker.ts
var DEFAULT_ROW_HEIGHT = 24;
var DEFAULT_COL_WIDTH = 80;
var DEFAULT_ZOOM_SCALE = 100;
var FAST_STRUCTURE_PARSE_THRESHOLD_BYTES = 5 * 1024 * 1024;
var MIN_ROW_HEIGHT_PX2 = 16;
function isLegacyXlsWorkbook(bytes) {
  return bytes.byteLength >= 8 && bytes[0] === 208 && bytes[1] === 207 && bytes[2] === 17 && bytes[3] === 224 && bytes[4] === 161 && bytes[5] === 177 && bytes[6] === 26 && bytes[7] === 225;
}
function shouldSkipXmlParsingForWorkbook(bytes, skipXmlParsing = false) {
  return skipXmlParsing || isLegacyXlsWorkbook(bytes);
}
function normalizeWorksheetVisibility2(value) {
  return value === "hidden" || value === "veryHidden" ? value : "visible";
}
var workbook = null;
var workbookSourceBytes = null;
var sourceFormulaCount = null;
var calculation = calculationReport("idle");
var hasCalculatedValues = false;
var chartsByWorkbookSheetIndex = [];
var chartsheets = [];
var formControlsByWorkbookSheetIndex = [];
var sheets = [];
var tablesByWorkbookSheetIndex = [];
var tabs = [];
function canParseXmlInWorker() {
  return typeof DOMParser !== "undefined";
}
function decodeXmlAttribute(value) {
  return value.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");
}
function readXmlAttribute(tag, name) {
  const escapedName = name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const match = new RegExp(`(?:^|\\s)${escapedName}="([^"]*)"`).exec(tag);
  return match ? decodeXmlAttribute(match[1] ?? "") : null;
}
function readArchiveText3(archive, path) {
  const entry = archive[path];
  return entry ? strFromU84(entry) : "";
}
function normalizeWorkbookRelationshipTarget(target) {
  if (target.startsWith("/")) {
    return target.replace(/^\/+/, "");
  }
  return target.startsWith("xl/") ? target : `xl/${target.replace(/^\.?\//, "")}`;
}
function parseWorkbookSheetPathsFromArchive(archive) {
  const workbookXml = readArchiveText3(archive, "xl/workbook.xml");
  const workbookRelationshipsXml = readArchiveText3(archive, "xl/_rels/workbook.xml.rels");
  if (!workbookXml || !workbookRelationshipsXml) {
    return [];
  }
  const relationshipTargetById = /* @__PURE__ */ new Map();
  for (const match of workbookRelationshipsXml.matchAll(/<Relationship\b[^>]*>/g)) {
    const tag = match[0];
    const id = readXmlAttribute(tag, "Id");
    const target = readXmlAttribute(tag, "Target");
    if (id && target) {
      relationshipTargetById.set(id, normalizeWorkbookRelationshipTarget(target));
    }
  }
  const paths = [];
  for (const match of workbookXml.matchAll(/<sheet\b[^>]*>/g)) {
    const tag = match[0];
    const relationshipId = readXmlAttribute(tag, "r:id") ?? readXmlAttribute(tag, "id");
    const target = relationshipId ? relationshipTargetById.get(relationshipId) : null;
    if (target) {
      paths.push(target);
    }
  }
  return paths;
}
function parseWorkerSheetLayoutAssets(bytes, sheetCount) {
  try {
    const archive = unzipSync3(bytes);
    const workbookSheetPaths = parseWorkbookSheetPathsFromArchive(archive);
    const sheetPaths = workbookSheetPaths.length > 0 ? workbookSheetPaths : Array.from({ length: sheetCount }, (_, index) => `xl/worksheets/sheet${index + 1}.xml`);
    return sheetPaths.slice(0, sheetCount).map((path) => {
      const xml = readArchiveText3(archive, path);
      if (!xml) {
        return null;
      }
      const rowHeightOverridesPx = {};
      for (const match of xml.matchAll(/<row\b[^>]*>/g)) {
        const tag = match[0];
        const rowNumber = Number(readXmlAttribute(tag, "r") ?? Number.NaN);
        const height = Number(readXmlAttribute(tag, "ht") ?? Number.NaN);
        const rowIndex = rowNumber - 1;
        if (rowIndex >= 0 && Number.isFinite(height)) {
          rowHeightOverridesPx[rowIndex] = Math.max(MIN_ROW_HEIGHT_PX2, Math.round(height * 1.33));
        }
      }
      return { rowHeightOverridesPx };
    });
  } catch {
    return [];
  }
}
function buildVisibleSheetIndexByWorkbookSheetIndex(nextWorkbook, showHiddenSheets = false) {
  const mapping = /* @__PURE__ */ new Map();
  let visibleIndex = 0;
  for (let workbookSheetIndex = 0; workbookSheetIndex < nextWorkbook.sheetCount; workbookSheetIndex += 1) {
    const worksheet = nextWorkbook.getSheet(workbookSheetIndex);
    const visibility = normalizeWorksheetVisibility2(worksheet.visibility);
    if (!showHiddenSheets && visibility !== "visible") {
      continue;
    }
    mapping.set(workbookSheetIndex, visibleIndex);
    visibleIndex += 1;
  }
  return mapping;
}
function normalizeRange(range) {
  return {
    start: {
      col: Math.min(range.start.col, range.end.col),
      row: Math.min(range.start.row, range.end.row)
    },
    end: {
      col: Math.max(range.start.col, range.end.col),
      row: Math.max(range.start.row, range.end.row)
    }
  };
}
function parseA1CellReference2(reference) {
  const match = /^([A-Z]+)(\d+)$/i.exec(reference.trim());
  if (!match) {
    return null;
  }
  const [, columnPart, rowPart] = match;
  let col = 0;
  for (const char of columnPart.toUpperCase()) {
    col = col * 26 + (char.charCodeAt(0) - 64);
  }
  return {
    col: col - 1,
    row: Number(rowPart) - 1
  };
}
function parseA1RangeReference2(reference) {
  const [startRef, endRef = startRef] = reference.split(":");
  const start = parseA1CellReference2(startRef ?? "");
  const end = parseA1CellReference2(endRef ?? "");
  if (!start || !end) {
    return null;
  }
  return normalizeRange({ end, start });
}
function resolveWorkbookReference(targetWorkbook, defaultWorkbookSheetIndex, rawReference, resolvingNamedRange = false) {
  const reference = rawReference.trim().replace(/^=/, "");
  if (!reference) {
    return null;
  }
  let workbookSheetIndex = defaultWorkbookSheetIndex;
  let rangeReference = reference;
  const bangIndex = reference.lastIndexOf("!");
  if (bangIndex >= 0) {
    let sheetName = reference.slice(0, bangIndex).trim();
    rangeReference = reference.slice(bangIndex + 1).trim();
    if (sheetName.startsWith("'") && sheetName.endsWith("'")) {
      sheetName = sheetName.slice(1, -1).replace(/''/g, "'");
    }
    const resolvedSheetIndex = targetWorkbook.sheetIndex(sheetName);
    if (resolvedSheetIndex === void 0) {
      return null;
    }
    workbookSheetIndex = resolvedSheetIndex;
  } else if (!resolvingNamedRange) {
    const namedRange = targetWorkbook.getNamedRange(reference);
    if (namedRange) {
      return resolveWorkbookReference(targetWorkbook, defaultWorkbookSheetIndex, namedRange, true);
    }
  }
  const range = parseA1RangeReference2(rangeReference.replace(/\$/g, ""));
  if (!range || workbookSheetIndex < 0 || workbookSheetIndex >= targetWorkbook.sheetCount) {
    return null;
  }
  try {
    return {
      range,
      worksheet: targetWorkbook.getSheet(workbookSheetIndex)
    };
  } catch {
    return null;
  }
}
function resolveFormControlItems(targetWorkbook, workbookSheetIndex, control) {
  if (control.kind !== "dropdown" && control.kind !== "listbox" || !control.inputRange) {
    return [];
  }
  const source = resolveWorkbookReference(targetWorkbook, workbookSheetIndex, control.inputRange);
  if (!source) {
    return [];
  }
  const items = [];
  for (let row = source.range.start.row; row <= source.range.end.row; row += 1) {
    for (let col = source.range.start.col; col <= source.range.end.col; col += 1) {
      items.push(source.worksheet.getFormattedValueAt(row, col));
    }
  }
  return items;
}
function parseWorksheetFreezePanes(worksheet) {
  const rawFreezePanes = worksheet.freezePanes;
  const row = typeof rawFreezePanes?.row === "number" && rawFreezePanes.row >= 0 ? rawFreezePanes.row : null;
  const col = typeof rawFreezePanes?.col === "number" && rawFreezePanes.col >= 0 ? rawFreezePanes.col : null;
  if (row === null && col === null) {
    return null;
  }
  return {
    col: col ?? 0,
    row: row ?? 0
  };
}
function parseWorksheetDataValidations(worksheet) {
  const rawDataValidations = Array.isArray(worksheet.dataValidations) ? worksheet.dataValidations : [];
  return rawDataValidations.flatMap((entry) => {
    if (!entry || typeof entry !== "object") {
      return [];
    }
    const validation = entry;
    const ranges = Array.isArray(validation.ranges) ? validation.ranges.flatMap((range) => {
      if (typeof range !== "string") {
        return [];
      }
      const parsedRange = parseA1RangeReference2(range);
      return parsedRange ? [parsedRange] : [];
    }) : [];
    const validationType = typeof validation.validationType === "string" ? validation.validationType : null;
    if (!validationType || ranges.length === 0) {
      return [];
    }
    return [{
      allowBlank: typeof validation.allowBlank === "boolean" ? validation.allowBlank : void 0,
      errorMessage: typeof validation.errorMessage === "string" ? validation.errorMessage : void 0,
      errorStyle: typeof validation.errorStyle === "string" ? validation.errorStyle : void 0,
      inputMessage: typeof validation.inputMessage === "string" ? validation.inputMessage : void 0,
      listSource: typeof validation.listSource === "string" ? validation.listSource : void 0,
      ranges,
      showDropdown: typeof validation.showDropdown === "boolean" ? validation.showDropdown : void 0,
      showErrorAlert: typeof validation.showErrorAlert === "boolean" ? validation.showErrorAlert : void 0,
      showInputMessage: typeof validation.showInputMessage === "boolean" ? validation.showInputMessage : void 0,
      validationType
    }];
  });
}
function resolveWorksheetZoomScale(worksheet, sheetState) {
  const candidates = [
    sheetState?.zoomScale,
    typeof worksheet.zoomScale === "number" ? worksheet.zoomScale : void 0
  ];
  const value = candidates.find((entry) => typeof entry === "number" && Number.isFinite(entry) && entry > 0);
  return value ?? DEFAULT_ZOOM_SCALE;
}
function resolveSheetDisplayUsedRange(usedRange, sheetState) {
  const [minRow, minCol, maxRow, maxCol] = usedRange;
  const maxContentRow = sheetState?.maxContentRow ?? -1;
  const maxContentCol = sheetState?.maxContentCol ?? -1;
  const maxVerticalMergeEndRow = sheetState?.maxVerticalMergeEndRow ?? -1;
  const maxHorizontalMergeEndCol = sheetState?.maxHorizontalMergeEndCol ?? -1;
  const maxMeaningfulRow = Math.max(maxContentRow, maxVerticalMergeEndRow);
  const maxMeaningfulCol = Math.max(maxContentCol, maxHorizontalMergeEndCol);
  if (maxMeaningfulRow < 0 && maxMeaningfulCol < 0) {
    return usedRange;
  }
  return [
    sheetState?.minContentRow !== void 0 && sheetState.minContentRow >= 0 ? Math.min(minRow, sheetState.minContentRow) : minRow,
    sheetState?.minContentCol !== void 0 && sheetState.minContentCol >= 0 ? Math.min(minCol, sheetState.minContentCol) : minCol,
    maxMeaningfulRow >= 0 ? maxContentRow >= 0 ? Math.min(maxRow, maxMeaningfulRow) : Math.max(maxRow, maxMeaningfulRow) : maxRow,
    maxMeaningfulCol >= 0 ? maxContentCol >= 0 ? Math.min(maxCol, maxMeaningfulCol) : Math.max(maxCol, maxMeaningfulCol) : maxCol
  ];
}
function buildSheetList(nextWorkbook, structureAssets, sheetLayoutStates, showHiddenSheets = false) {
  const sheetsByWorkbookSheetIndex = [];
  for (let index = 0; index < nextWorkbook.sheetCount; index += 1) {
    const worksheet = nextWorkbook.getSheet(index);
    const sheetState = structureAssets?.sheetStatesByWorkbookSheetIndex[index] ?? sheetLayoutStates?.[index] ?? null;
    const mergeMetadata = resolveWorksheetMergeMetadata(worksheet);
    const effectiveSheetState = {
      ...sheetState,
      ...mergeMetadata
    };
    const defaultColWidthPx = resolveWorksheetDefaultColumnWidthPixels(
      worksheet,
      sheetState?.columnWidthCharacterWidthPx,
      sheetState?.defaultColWidthPx ?? DEFAULT_COL_WIDTH
    );
    const defaultRowHeightPx = resolveWorksheetDefaultRowHeightPixels(
      worksheet,
      sheetState?.defaultRowHeightPx ?? DEFAULT_ROW_HEIGHT
    );
    const visibility = normalizeWorksheetVisibility2(worksheet.visibility);
    if (!showHiddenSheets && visibility !== "visible") {
      continue;
    }
    const resolveColumnWidthPx = (col) => {
      const width = worksheet.getColumnWidth(col);
      if (width !== void 0 && width !== null) {
        return resolveSheetColumnWidthPixels(width, sheetState?.columnWidthCharacterWidthPx);
      }
      return sheetState?.colWidthOverridesPx?.[col] ?? defaultColWidthPx;
    };
    const resolveRowHeightPx = (row) => {
      const height = worksheet.getRowHeight(row);
      if (height !== void 0 && height !== null) {
        return Math.max(Math.round(height * 1.33), 16);
      }
      return sheetState?.rowHeightOverridesPx?.[row] ?? defaultRowHeightPx;
    };
    const usedRange = worksheet.usedRange();
    if (!usedRange) {
      sheetsByWorkbookSheetIndex.push({
        autoFilterRanges: sheetState?.autoFilterRanges ?? [],
        cachedFormulaValues: sheetState?.cachedFormulaValues ?? {},
        columnWidthCharacterWidthPx: sheetState?.columnWidthCharacterWidthPx,
        colCount: 0,
        colStyleIds: sheetState?.colStyleIds ?? {},
        colWidthOverridesPx: sheetState?.colWidthOverridesPx ?? {},
        colWidths: [],
        conditionalFormatRules: sheetState?.conditionalFormatRules ?? [],
        dataValidations: parseWorksheetDataValidations(worksheet),
        defaultColWidthPx,
        defaultRowHeightPx,
        freezePanes: parseWorksheetFreezePanes(worksheet),
        hasHorizontalMerges: mergeMetadata.hasHorizontalMerges,
        hasVerticalMerges: mergeMetadata.hasVerticalMerges,
        maxHorizontalMergeEndCol: mergeMetadata.maxHorizontalMergeEndCol,
        maxVerticalMergeEndRow: mergeMetadata.maxVerticalMergeEndRow,
        hiddenCols: sheetState?.hiddenCols ?? [],
        hiddenRows: sheetState?.hiddenRows ?? [],
        minUsedCol: -1,
        minUsedRow: -1,
        maxUsedCol: -1,
        maxUsedRow: -1,
        name: worksheet.name,
        visibility,
        namedCellStyleByName: structureAssets?.namedCellStyleByName ?? {},
        rowCount: 0,
        rowHeightOverridesPx: sheetState?.rowHeightOverridesPx ?? {},
        rowHeights: [],
        rowStyleIds: sheetState?.rowStyleIds ?? {},
        showGridLines: sheetState?.showGridLines ?? true,
        sparklines: sheetState?.sparklines ?? [],
        styleById: structureAssets?.styleById ?? {},
        tableStyleByName: structureAssets?.tableStyleByName ?? {},
        themePalette: structureAssets?.themePalette ?? { colorsByIndex: {} },
        visibleCols: [],
        visibleRows: [],
        workbookSheetIndex: index,
        zoomScale: resolveWorksheetZoomScale(worksheet, sheetState)
      });
      continue;
    }
    const [rawMinRow, rawMinCol, resolvedMaxRow, resolvedMaxCol] = resolveSheetDisplayUsedRange(usedRange, effectiveSheetState);
    const maxRow = Math.max(resolvedMaxRow, sheetState?.maxContentRow ?? -1, effectiveSheetState.maxVerticalMergeEndRow ?? -1);
    const maxCol = Math.max(resolvedMaxCol, sheetState?.maxContentCol ?? -1, effectiveSheetState.maxHorizontalMergeEndCol ?? -1);
    const minRow = structureAssets ? rawMinRow : 0;
    const minCol = structureAssets ? rawMinCol : 0;
    const visibleRows = [];
    const hiddenRows = [];
    for (let row = 0; row <= maxRow; row += 1) {
      if (worksheet.isRowHidden(row)) {
        hiddenRows.push(row);
      } else {
        visibleRows.push(row);
      }
    }
    const visibleCols = [];
    const hiddenCols = [];
    for (let col = 0; col <= maxCol; col += 1) {
      if (worksheet.isColumnHidden(col)) {
        hiddenCols.push(col);
      } else {
        visibleCols.push(col);
      }
    }
    sheetsByWorkbookSheetIndex.push({
      autoFilterRanges: sheetState?.autoFilterRanges ?? [],
      cachedFormulaValues: sheetState?.cachedFormulaValues ?? {},
      columnWidthCharacterWidthPx: sheetState?.columnWidthCharacterWidthPx,
      colCount: visibleCols.length,
      colStyleIds: sheetState?.colStyleIds ?? {},
      colWidthOverridesPx: sheetState?.colWidthOverridesPx ?? {},
      colWidths: visibleCols.map(resolveColumnWidthPx),
      conditionalFormatRules: sheetState?.conditionalFormatRules ?? [],
      dataValidations: parseWorksheetDataValidations(worksheet),
      defaultColWidthPx,
      defaultRowHeightPx,
      freezePanes: parseWorksheetFreezePanes(worksheet),
      hasHorizontalMerges: mergeMetadata.hasHorizontalMerges,
      hasVerticalMerges: mergeMetadata.hasVerticalMerges,
      maxHorizontalMergeEndCol: mergeMetadata.maxHorizontalMergeEndCol,
      maxVerticalMergeEndRow: mergeMetadata.maxVerticalMergeEndRow,
      hiddenCols,
      hiddenRows,
      minUsedCol: minCol,
      minUsedRow: minRow,
      maxUsedCol: maxCol,
      maxUsedRow: maxRow,
      name: worksheet.name,
      visibility,
      namedCellStyleByName: structureAssets?.namedCellStyleByName ?? {},
      rowCount: visibleRows.length,
      rowHeightOverridesPx: sheetState?.rowHeightOverridesPx ?? {},
      rowHeights: visibleRows.map(resolveRowHeightPx),
      rowStyleIds: sheetState?.rowStyleIds ?? {},
      showGridLines: sheetState?.showGridLines ?? true,
      sparklines: sheetState?.sparklines ?? [],
      styleById: structureAssets?.styleById ?? {},
      tableStyleByName: structureAssets?.tableStyleByName ?? {},
      themePalette: structureAssets?.themePalette ?? { colorsByIndex: {} },
      visibleCols,
      visibleRows,
      workbookSheetIndex: index,
      zoomScale: resolveWorksheetZoomScale(worksheet, sheetState)
    });
  }
  return sheetsByWorkbookSheetIndex;
}
function mapWorksheetTables(worksheet, autoFilterRanges = []) {
  const rawTables = worksheet?.tables ?? [];
  const mappedTables = rawTables.flatMap((table, index) => {
    const rawColumns = Array.isArray(table.columns) ? table.columns : [];
    const rawName = typeof table.name === "string" ? table.name : `Table${index + 1}`;
    const rawDisplayName = typeof table.displayName === "string" ? table.displayName : typeof table.name === "string" ? table.name : `Table ${index + 1}`;
    const rawReference = typeof table.reference === "string" ? table.reference : "";
    const reference = rawReference;
    const parsedRange = parseA1RangeReference2(reference);
    if (!parsedRange) {
      return [];
    }
    return [{
      columns: rawColumns.map((column, columnIndex) => ({
        id: typeof column.id === "number" ? column.id ?? columnIndex + 1 : columnIndex + 1,
        index: columnIndex,
        name: typeof column.name === "string" ? column.name ?? `Column ${columnIndex + 1}` : `Column ${columnIndex + 1}`
      })),
      displayName: rawDisplayName,
      end: parsedRange.end,
      headerRowCount: resolveWorkbookTableCount(table.headerRowCount, 1),
      headerRowCellStyle: typeof table.headerRowCellStyle === "string" ? table.headerRowCellStyle : void 0,
      name: rawName,
      reference,
      start: parsedRange.start,
      styleInfo: table.styleInfo,
      totalsRowCount: resolveWorkbookTableCount(table.totalsRowCount, 0),
      totalsRowShown: resolveWorkbookTableBoolean(table.totalsRowShown)
    }];
  });
  const existingReferences = new Set(mappedTables.map((table) => table.reference));
  const mappedAutoFilterTables = autoFilterRanges.flatMap((range, index) => {
    const reference = `${cellAddressToA1(range.start)}:${cellAddressToA1(range.end)}`;
    if (existingReferences.has(reference)) {
      return [];
    }
    const columnCount = Math.max(0, range.end.col - range.start.col + 1);
    const columns = Array.from({ length: columnCount }, (_, columnIndex) => {
      const headerValue = worksheet ? decodeHtmlEntities(worksheet.getFormattedValueAt(range.start.row, range.start.col + columnIndex) ?? "") : "";
      return {
        id: columnIndex + 1,
        index: columnIndex,
        name: headerValue.trim().length > 0 ? headerValue : `Column ${columnIndex + 1}`
      };
    });
    return [{
      columns,
      displayName: `AutoFilter ${index + 1}`,
      end: range.end,
      headerRowCount: 1,
      name: `__autofilter_${index + 1}_${reference}`,
      reference,
      start: range.start,
      totalsRowCount: 0,
      totalsRowShown: false
    }];
  });
  return [...mappedTables, ...mappedAutoFilterTables];
}
function resolveWorkbookTableCount(value, fallback) {
  if (typeof value === "number" && Number.isFinite(value) && value >= 0) {
    return value;
  }
  if (typeof value === "string") {
    const parsed = Number.parseInt(value, 10);
    if (Number.isFinite(parsed) && parsed >= 0) {
      return parsed;
    }
  }
  return fallback;
}
function resolveWorkbookTableBoolean(value) {
  if (typeof value === "boolean") {
    return value;
  }
  if (typeof value === "number") {
    return value !== 0;
  }
  if (typeof value === "string") {
    const normalized = value.trim().toLowerCase();
    if (normalized === "0" || normalized === "false" || normalized === "") {
      return false;
    }
    if (normalized === "1" || normalized === "true") {
      return true;
    }
  }
  return false;
}
function decodeHtmlEntities(value) {
  return value.replace(/&quot;/g, '"').replace(/&#34;/g, '"').replace(/&apos;/g, "'").replace(/&#39;/g, "'").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");
}
function getCellDisplayValue(worksheet, row, col) {
  const formatted = worksheet.getFormattedValueAt(row, col);
  if (formatted && !formatted.startsWith("#")) {
    return decodeHtmlEntities(formatted);
  }
  const cellValue = worksheet.getCalculatedValueAt(row, col);
  if (cellValue.is_error) {
    return cellValue.asError() ?? "";
  }
  if (cellValue.is_empty) {
    return "";
  }
  return cellValue.toString();
}
function cellAddressToA1(cell) {
  let col = cell.col + 1;
  let label = "";
  while (col > 0) {
    const remainder = (col - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    col = Math.floor((col - 1) / 26);
  }
  return `${label}${cell.row + 1}`;
}
async function loadWorkbook(buffer, skipXmlParsing = false, showHiddenSheets = false, externalFnValues) {
  const wasmModule = await getSheetsWasmModule();
  const bytes = new Uint8Array(buffer);
  const effectiveSkipXmlParsing = shouldSkipXmlParsingForWorkbook(bytes, skipXmlParsing);
  let activeWorkbook = wasmModule.Workbook.fromBytes(bytes);
  const totalFormulas = countWorkbookFormulas(activeWorkbook);
  sourceFormulaCount = countSourceWorkbookFormulas(bytes);
  calculation = calculationReport("skipped", "auto-formula-limit", totalFormulas, sourceFormulaCount);
  hasCalculatedValues = false;
  if (totalFormulas <= AUTO_CALCULATE_FORMULA_THRESHOLD) {
    const result = safeCalculate(activeWorkbook, {
      reparse: () => wasmModule.Workbook.fromBytes(bytes),
      calcOptions: externalCalcOptions(externalFnValues),
      sourceFormulaCount
    });
    calculation = result.calculation;
    hasCalculatedValues = result.calculated;
    activeWorkbook = result.workbook;
  }
  const nextWorkbook = activeWorkbook;
  const shouldUseFastStructureParse = bytes.byteLength >= FAST_STRUCTURE_PARSE_THRESHOLD_BYTES && totalFormulas <= AUTO_CALCULATE_FORMULA_THRESHOLD;
  const structureAssets = effectiveSkipXmlParsing || shouldUseFastStructureParse || !canParseXmlInWorker() ? null : parseWorkbookStructureAssets(bytes, {
    includeCachedFormulaValues: true
  });
  formControlsByWorkbookSheetIndex = collectWorkbookFormControls(
    nextWorkbook,
    structureAssets?.themePalette
  ).map(
    (controls, workbookSheetIndex) => controls.map((control) => ({
      ...control,
      items: resolveFormControlItems(nextWorkbook, workbookSheetIndex, control)
    }))
  );
  const sheetLayoutStates = structureAssets ? void 0 : parseWorkerSheetLayoutAssets(bytes, nextWorkbook.sheetCount);
  workbook = nextWorkbook;
  workbookSourceBytes = bytes;
  sheets = buildSheetList(nextWorkbook, structureAssets, sheetLayoutStates, showHiddenSheets);
  tablesByWorkbookSheetIndex = Array.from(
    { length: nextWorkbook.sheetCount },
    (_, workbookSheetIndex) => mapWorksheetTables(
      nextWorkbook.getSheet(workbookSheetIndex),
      structureAssets?.sheetStatesByWorkbookSheetIndex[workbookSheetIndex]?.autoFilterRanges ?? []
    )
  );
  const visibleSheetIndexByWorkbookSheetIndex = new Map(sheets.map((sheet, index) => [sheet.workbookSheetIndex, index]));
  const hasCharts = Array.from({ length: nextWorkbook.sheetCount }, (_, workbookSheetIndex) => {
    const worksheet = nextWorkbook.getSheet(workbookSheetIndex);
    const hasClassicCharts = Array.isArray(worksheet.charts) && worksheet.charts.length > 0;
    const hasModernCharts = Array.isArray(worksheet.chartsEx) && worksheet.chartsEx.length > 0;
    return hasClassicCharts || hasModernCharts;
  }).some(Boolean);
  const chartStyleAssets = effectiveSkipXmlParsing || !hasCharts || !canParseXmlInWorker() ? null : parseWorkbookChartStyleAssets(bytes);
  const chartAssets = loadWorkbookChartAssets(
    nextWorkbook,
    chartStyleAssets,
    visibleSheetIndexByWorkbookSheetIndex,
    showHiddenSheets
  );
  chartsByWorkbookSheetIndex = chartAssets.chartsByWorkbookSheetIndex;
  chartsheets = chartAssets.chartsheets;
  tabs = chartAssets.tabs;
  return {
    calculation,
    chartsByWorkbookSheetIndex,
    chartsheets,
    formControlsByWorkbookSheetIndex,
    sheets,
    tablesByWorkbookSheetIndex,
    tabs
  };
}
async function parseCharts(buffer, skipXmlParsing = false, showHiddenSheets = false, autoCalculateFormulaLimit = AUTO_CALCULATE_FORMULA_THRESHOLD) {
  const wasmModule = await getSheetsWasmModule();
  const bytes = new Uint8Array(buffer);
  const effectiveSkipXmlParsing = shouldSkipXmlParsingForWorkbook(bytes, skipXmlParsing);
  let activeWorkbook = wasmModule.Workbook.fromBytes(bytes);
  const totalFormulas = countWorkbookFormulas(activeWorkbook);
  if (totalFormulas <= autoCalculateFormulaLimit) {
    const result = safeCalculate(activeWorkbook, {
      reparse: () => wasmModule.Workbook.fromBytes(bytes)
    });
    activeWorkbook = result.workbook;
  }
  const nextWorkbook = activeWorkbook;
  const visibleSheetIndexByWorkbookSheetIndex = buildVisibleSheetIndexByWorkbookSheetIndex(nextWorkbook, showHiddenSheets);
  const chartStyleAssets = effectiveSkipXmlParsing || !canParseXmlInWorker() ? null : parseWorkbookChartStyleAssets(bytes);
  const chartAssets = loadWorkbookChartAssets(
    nextWorkbook,
    chartStyleAssets,
    visibleSheetIndexByWorkbookSheetIndex,
    showHiddenSheets
  );
  return {
    chartsByWorkbookSheetIndex: chartAssets.chartsByWorkbookSheetIndex,
    chartsheets: chartAssets.chartsheets,
    tabs: chartAssets.tabs
  };
}
async function recalculateWorkbook(externalFnValues) {
  if (!workbook) {
    return { calculated: false, skipReason: null, calculation: calculationReport("idle") };
  }
  const wasmModule = await getSheetsWasmModule();
  const sourceBytes = workbookSourceBytes;
  const result = safeCalculate(workbook, {
    calcOptions: externalCalcOptions(externalFnValues),
    sourceFormulaCount,
    reparse: sourceBytes ? () => wasmModule.Workbook.fromBytes(sourceBytes) : void 0
  });
  hasCalculatedValues = result.workbook !== workbook ? false : hasCalculatedValues || result.calculated;
  workbook = result.workbook;
  calculation = result.calculation;
  return { calculated: result.calculated, skipReason: result.skipReason, calculation };
}
function respond(message) {
  self.postMessage(message);
}
async function handleMessage(message) {
  switch (message.type) {
    case "load": {
      if (message.payload.wasmSource !== void 0) {
        setWasmSource(message.payload.wasmSource);
      }
      return loadWorkbook(
        message.payload.buffer,
        message.payload.skipXmlParsing,
        message.payload.showHiddenSheets,
        message.payload.externalFnValues
      );
    }
    case "parseCharts": {
      if (message.payload.wasmSource !== void 0) {
        setWasmSource(message.payload.wasmSource);
      }
      return parseCharts(
        message.payload.buffer,
        message.payload.skipXmlParsing,
        message.payload.showHiddenSheets,
        message.payload.autoCalculateFormulaLimit
      );
    }
    case "recalculate": {
      return recalculateWorkbook(message.payload.externalFnValues);
    }
    case "getCellSnapshot": {
      if (!workbook) {
        return {
          displayValue: "",
          formula: "",
          diagnostic: { source: "unknown", error: null }
        };
      }
      const targetSheet = sheets.find((sheet) => sheet.workbookSheetIndex === message.payload.workbookSheetIndex) ?? null;
      const worksheet = workbook.getSheet(message.payload.workbookSheetIndex);
      return {
        displayValue: getCellDisplayValue(worksheet, message.payload.row, message.payload.col),
        formula: worksheet.getFormulaAt(message.payload.row, message.payload.col) ?? "",
        diagnostic: cellCalculationDiagnostic(worksheet, message.payload.row, message.payload.col, targetSheet?.cachedFormulaValues?.[cellAddressToA1(message.payload)], calculation, hasCalculatedValues)
      };
    }
    case "findDataBoundary": {
      const sheet = sheets.find((entry) => entry.workbookSheetIndex === message.payload.workbookSheetIndex);
      if (!workbook || !sheet) throw new Error("Worksheet unavailable");
      const worksheet = workbook.getSheet(sheet.workbookSheetIndex);
      return findDataBoundary(
        message.payload,
        sheet,
        (row, col) => worksheetHasContent(worksheet, row, col)
      );
    }
    case "getRowsBatch": {
      if (!workbook) {
        return null;
      }
      const worksheet = workbook.getSheet(message.payload.workbookSheetIndex);
      if (typeof worksheet.getRowsBatch !== "function") {
        return null;
      }
      return worksheet.getRowsBatch(message.payload.startRow, message.payload.rowCount, {
        includeFormulas: true,
        includeHyperlinks: true,
        includeMergeInfo: true,
        includeStyles: true,
        useFormattedValues: true
      });
    }
    default:
      return null;
  }
}
self.addEventListener("message", (event) => {
  const message = event.data;
  void handleMessage(message).then((result) => {
    respond({
      id: message.id,
      result,
      success: true
    });
  }).catch((error) => {
    respond({
      error: error instanceof Error ? error.message : "Worker request failed.",
      id: message.id,
      success: false
    });
  });
});
//# sourceMappingURL=xlsx-worker.js.map