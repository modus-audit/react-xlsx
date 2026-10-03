import * as React from 'react';
import * as _dukelib_sheets_wasm from '@dukelib/sheets-wasm';
import { Workbook, Worksheet } from '@dukelib/sheets-wasm';
import * as react_jsx_runtime from 'react/jsx-runtime';

type DataDirection = "ArrowUp" | "ArrowDown" | "ArrowLeft" | "ArrowRight";
interface DataNavigationRequest {
    cell: XlsxCellAddress;
    direction: DataDirection;
    maxRow: number;
    maxCol: number;
}

/** Serializable external-call values cross the worker boundary; missing entries preserve cached workbook values. */
/** Canonical key for one external call. `args` are the engine's stringified argument values
 *  (CCH args are strings/numbers), matching how the host parsed them from the formula text. */
declare function externalCallKey(name: string, args: readonly string[]): string;
/** Serializable map handed to the controller: `externalCallKey(name, args)` -> resolved value. */
type ExternalFnValues = Record<string, string | number>;

interface XlsxThemePalette {
    colorsByIndex: Record<number, string>;
    majorLatinFont?: string;
    minorLatinFont?: string;
}
interface XlsxResolvedCellStyle {
    [key: string]: unknown;
    alignment?: Record<string, unknown>;
    border?: Record<string, Record<string, unknown>>;
    cellControl?: {
        kind: "checkbox";
    };
    fill?: Record<string, unknown>;
    font?: Record<string, unknown>;
}
interface XlsxTableStyleDefinition {
    [elementType: string]: XlsxResolvedCellStyle;
}
interface XlsxConditionalFormatValueObject {
    type: string;
    value?: number;
}
interface XlsxConditionalDataBarRule {
    axisColor?: Record<string, unknown>;
    border?: boolean;
    color?: Record<string, unknown>;
    borderColor?: Record<string, unknown>;
    cfvos: XlsxConditionalFormatValueObject[];
    gradient?: boolean;
    kind: "dataBar";
    maxLength?: number;
    minLength?: number;
    negativeBarBorderColorSameAsPositive?: boolean;
    negativeBorderColor?: Record<string, unknown>;
    negativeFillColor?: Record<string, unknown>;
    priority: number;
    ranges: XlsxCellRange[];
    showValue?: boolean;
}
interface XlsxConditionalColorScaleRule {
    cfvos: XlsxConditionalFormatValueObject[];
    colors: Record<string, unknown>[];
    kind: "colorScale";
    priority: number;
    ranges: XlsxCellRange[];
}
interface XlsxConditionalFormatIcon {
    iconId: number;
    iconSet: string;
}
interface XlsxConditionalIconSetRule {
    cfvos: XlsxConditionalFormatValueObject[];
    icons: XlsxConditionalFormatIcon[];
    kind: "iconSet";
    priority: number;
    ranges: XlsxCellRange[];
    reverse?: boolean;
    showValue?: boolean;
}
interface XlsxConditionalStyledRule {
    aboveAverage?: boolean;
    bottom?: boolean;
    equalAverage?: boolean;
    formulas: string[];
    kind: "styled";
    operator?: string;
    percent?: boolean;
    priority: number;
    rank?: number;
    ranges: XlsxCellRange[];
    ruleType: string;
    stdDev?: number;
    stopIfTrue?: boolean;
    style: XlsxResolvedCellStyle;
    text?: string;
    timePeriod?: string;
}
type XlsxConditionalFormatRule = XlsxConditionalColorScaleRule | XlsxConditionalStyledRule | XlsxConditionalDataBarRule | XlsxConditionalIconSetRule;
interface XlsxDataValidation {
    allowBlank?: boolean;
    errorMessage?: string;
    errorStyle?: string;
    inputMessage?: string;
    listSource?: string;
    ranges: XlsxCellRange[];
    showDropdown?: boolean;
    showErrorAlert?: boolean;
    showInputMessage?: boolean;
    validationType: string;
}
interface XlsxFreezePanes {
    col: number;
    row: number;
}
type XlsxSheetVisibility = "hidden" | "veryHidden" | "visible";
interface XlsxSheetData {
    autoFilterRanges: XlsxCellRange[];
    cachedFormulaValues: Record<string, string>;
    colWidthOverridesPx: Record<number, number>;
    colStyleIds: Record<number, number>;
    hiddenCols?: number[];
    hiddenRows?: number[];
    conditionalFormatRules: XlsxConditionalFormatRule[];
    dataValidations: XlsxDataValidation[];
    name: string;
    visibility: XlsxSheetVisibility;
    columnWidthCharacterWidthPx?: number;
    defaultColWidthPx: number;
    defaultRowHeightPx: number;
    freezePanes: XlsxFreezePanes | null;
    hasHorizontalMerges: boolean;
    hasVerticalMerges: boolean;
    maxHorizontalMergeEndCol: number;
    maxVerticalMergeEndRow: number;
    minUsedCol: number;
    minUsedRow: number;
    maxUsedCol: number;
    maxUsedRow: number;
    rowCount: number;
    colCount: number;
    rowHeightOverridesPx: Record<number, number>;
    rowStyleIds: Record<number, number>;
    namedCellStyleByName: Record<string, XlsxResolvedCellStyle>;
    styleById: Record<number, XlsxResolvedCellStyle>;
    tableStyleByName: Record<string, XlsxTableStyleDefinition>;
    visibleRows: number[];
    visibleCols: number[];
    colWidths: number[];
    rowHeights: number[];
    showGridLines: boolean;
    sparklines: XlsxSparkline[];
    themePalette: XlsxThemePalette;
    workbookSheetIndex: number;
    zoomScale?: number;
}
interface XlsxCellAddress {
    col: number;
    row: number;
}
/** A column width or row height in pixels, for `resizeColumns` and `resizeRows`. */
type XlsxAxisSize = {
    index: number;
    sizePx: number;
};
interface XlsxCellRange {
    end: XlsxCellAddress;
    start: XlsxCellAddress;
}
/**
 * Color value accepted by persisted workbook style mutation APIs.
 *
 * Use `rgb` with a six-character `hex` value such as `"2563EB"` for the
 * simplest browser-to-Excel mapping. `argb`, `theme`, and `indexed` are
 * available for callers that need to preserve Excel-native color references.
 */
interface XlsxCellStyleColorInput {
    /** Excel color representation to write. Defaults are handled by the workbook engine. */
    colorType?: "auto" | "rgb" | "argb" | "theme" | "indexed";
    /** Hex color string, for example `"2563EB"` for RGB or `"FF2563EB"` for ARGB. */
    hex?: string;
    /** Red channel, 0-255. */
    r?: number;
    /** Green channel, 0-255. */
    g?: number;
    /** Blue channel, 0-255. */
    b?: number;
    /** Alpha channel, 0-255, used by ARGB colors. */
    a?: number;
    /** Theme color index for Excel theme colors. */
    themeIndex?: number;
    /** Theme tint/shade adjustment. */
    tint?: number;
    /** Indexed color palette entry. */
    paletteIndex?: number;
}
/** Font styling to persist on a cell or range. */
interface XlsxCellFontStyleInput {
    /** Font face name, such as `"Aptos"` or `"Calibri"`. */
    name?: string;
    /** Font size in points. */
    size?: number;
    /** Whether the font is bold. */
    bold?: boolean;
    /** Whether the font is italic. */
    italic?: boolean;
    /** Underline style. */
    underline?: "none" | "single" | "double" | "singleAccounting" | "doubleAccounting";
    /** Whether the font is struck through. */
    strikethrough?: boolean;
    /** Font color. */
    color?: XlsxCellStyleColorInput;
    /** Baseline, superscript, or subscript positioning. */
    verticalAlign?: "baseline" | "superscript" | "subscript";
    /** Excel font family classification. */
    family?: number;
    /** Font charset identifier. */
    charset?: number;
    /** Excel font scheme value. */
    scheme?: string;
}
/** A color stop used by gradient fills. */
interface XlsxCellGradientStopInput {
    /** Stop position from 0 to 1. */
    position: number;
    /** Stop color. */
    color: XlsxCellStyleColorInput;
}
/** Fill styling to persist on a cell or range. */
interface XlsxCellFillStyleInput {
    /** Fill type. Use `"solid"` with `color` for typical Excel fill color. */
    fillType?: "none" | "solid" | "pattern" | "gradient";
    /** Primary fill color. */
    color?: XlsxCellStyleColorInput;
    /** Excel pattern name for pattern fills. */
    pattern?: string;
    /** Foreground color for pattern fills. */
    foreground?: XlsxCellStyleColorInput;
    /** Background color for pattern fills. */
    background?: XlsxCellStyleColorInput;
    /** Gradient type for gradient fills. */
    gradientType?: "linear" | "path";
    /** Gradient angle in degrees. */
    angle?: number;
    /** Gradient color stops. */
    stops?: XlsxCellGradientStopInput[];
}
/** A single border edge style. */
interface XlsxCellBorderEdgeInput {
    /** Excel border line style. */
    style?: "none" | "thin" | "medium" | "thick" | "dashed" | "dotted" | "double" | "hair" | "mediumDashed" | "dashDot" | "mediumDashDot" | "dashDotDot" | "mediumDashDotDot" | "slantDashDot";
    /** Border line color. */
    color?: XlsxCellStyleColorInput;
}
/** Border styling to persist on a cell or range. */
interface XlsxCellBorderStyleInput {
    /** Left border edge. */
    left?: XlsxCellBorderEdgeInput;
    /** Right border edge. */
    right?: XlsxCellBorderEdgeInput;
    /** Top border edge. */
    top?: XlsxCellBorderEdgeInput;
    /** Bottom border edge. */
    bottom?: XlsxCellBorderEdgeInput;
    /** Diagonal border edge. */
    diagonal?: XlsxCellBorderEdgeInput;
    /** Diagonal border direction. */
    diagonalDirection?: "none" | "down" | "up" | "both";
}
/** Alignment styling to persist on a cell or range. */
interface XlsxCellAlignmentInput {
    /** Horizontal alignment. */
    horizontal?: "general" | "left" | "center" | "right" | "fill" | "justify" | "centerContinuous" | "distributed";
    /** Vertical alignment. */
    vertical?: "top" | "center" | "bottom" | "justify" | "distributed";
    /** Whether text wraps inside the cell. */
    wrapText?: boolean;
    /** Whether text shrinks to fit the cell. */
    shrinkToFit?: boolean;
    /** Indentation level. */
    indent?: number;
    /** Text rotation in degrees. */
    rotation?: number;
    /** Text reading order. */
    readingOrder?: "contextDependent" | "leftToRight" | "rightToLeft";
}
/** Number format styling to persist on a cell or range. */
interface XlsxCellNumberFormatInput {
    /** Number format source. Use `"custom"` with `formatString` for custom Excel formats. */
    formatType?: "general" | "builtin" | "custom";
    /** Built-in Excel number format id. */
    id?: number;
    /** Excel number format string, for example `"$#,##0.00"` or `"m/d/yyyy"`. */
    formatString?: string;
}
/** Protection flags to persist on a cell or range. */
interface XlsxCellProtectionInput {
    /** Whether the cell is locked when sheet protection is enabled. */
    locked?: boolean;
    /** Whether the cell formula is hidden when sheet protection is enabled. */
    hidden?: boolean;
}
/**
 * Persisted Excel cell style patch.
 *
 * Passing a partial style updates the supplied style groups on the target cell
 * or range. Unlike the `getCellStyle` render override prop, this mutates the
 * workbook and is included in exported XLSX bytes.
 */
interface XlsxCellStyleInput {
    /** Font formatting such as bold, italic, size, family, and font color. */
    font?: XlsxCellFontStyleInput;
    /** Cell fill formatting such as background color or gradient fill. */
    fill?: XlsxCellFillStyleInput;
    /** Border formatting for cell edges. */
    border?: XlsxCellBorderStyleInput;
    /** Alignment, wrapping, indentation, rotation, and reading order. */
    alignment?: XlsxCellAlignmentInput;
    /** Excel number/date/currency/percent format metadata. */
    numberFormat?: XlsxCellNumberFormatInput;
    /** Cell protection flags. */
    protection?: XlsxCellProtectionInput;
}
interface XlsxSparkline {
    color?: string;
    firstColor?: string;
    highColor?: string;
    lastColor?: string;
    lowColor?: string;
    markerColor?: string;
    markers?: boolean;
    negative?: boolean;
    negativeColor?: string;
    range: XlsxCellRange;
    sheetName?: string;
    target: XlsxCellAddress;
    type: "column" | "line" | "winLoss";
}
interface XlsxClipboardData {
    html: string;
    structured: string;
    text: string;
}
interface XlsxTableColumn {
    id: number;
    index: number;
    name: string;
}
interface XlsxTableStyleInfo {
    name?: string;
    showColumnStripes?: boolean;
    showFirstColumn?: boolean;
    showLastColumn?: boolean;
    showRowStripes?: boolean;
}
interface XlsxTable {
    columns: XlsxTableColumn[];
    displayName: string;
    end: XlsxCellAddress;
    headerRowCount: number;
    headerRowCellStyle?: string;
    name: string;
    reference: string;
    start: XlsxCellAddress;
    styleInfo?: XlsxTableStyleInfo;
    totalsRowCount: number;
    totalsRowShown: boolean;
}
type XlsxTableSortDirection = "ascending" | "descending";
interface XlsxTableSortState {
    columnIndex: number;
    direction: XlsxTableSortDirection;
    tableName: string;
}
interface XlsxImageMarker {
    col: number;
    colOffsetEmu: number;
    row: number;
    rowOffsetEmu: number;
}
type XlsxImageAnchor = {
    from: XlsxImageMarker;
    kind: "one-cell";
    sizeEmu: {
        cx: number;
        cy: number;
    };
} | {
    kind: "absolute";
    positionEmu: {
        x: number;
        y: number;
    };
    sizeEmu: {
        cx: number;
        cy: number;
    };
} | {
    from: XlsxImageMarker;
    kind: "two-cell";
    to: XlsxImageMarker;
};
interface XlsxImage {
    anchor: XlsxImageAnchor;
    description?: string;
    editable?: boolean;
    hyperlink?: string;
    id: string;
    mediaPath?: string;
    mimeType: string;
    name?: string;
    sheetIndex: number;
    src: string;
    workbookSheetIndex: number;
    zIndex: number;
}
interface XlsxShapeFill {
    color?: string;
    none?: boolean;
    opacity?: number;
}
interface XlsxShapeStroke {
    color?: string;
    dash?: string;
    headEndType?: string;
    none?: boolean;
    opacity?: number;
    tailEndType?: string;
    widthPx?: number;
}
interface XlsxShapeTextRun {
    bold?: boolean;
    color?: string;
    fontFamily?: string;
    fontSizePt?: number;
    italic?: boolean;
    text: string;
    underline?: boolean;
}
interface XlsxShapeParagraph {
    align?: "center" | "justify" | "left" | "right";
    runs: XlsxShapeTextRun[];
}
interface XlsxShapeTextBox {
    horizontalAlign?: "center" | "left";
    insetPx?: {
        bottom: number;
        left: number;
        right: number;
        top: number;
    };
    verticalAlign?: "bottom" | "middle" | "top";
}
interface XlsxShape {
    anchor: XlsxImageAnchor;
    description?: string;
    fill?: XlsxShapeFill;
    flipH?: boolean;
    flipV?: boolean;
    geometry: string;
    geometryAdjustments?: Record<string, number>;
    hidden?: boolean;
    hyperlink?: string;
    id: string;
    name?: string;
    paragraphs: XlsxShapeParagraph[];
    rotationDeg?: number;
    scaleX?: number;
    scaleY?: number;
    sheetIndex: number;
    svgPath?: string;
    svgViewBox?: {
        height: number;
        width: number;
    };
    stroke?: XlsxShapeStroke;
    textBox?: XlsxShapeTextBox;
    workbookSheetIndex: number;
    zIndex: number;
}
type XlsxFormControlKind = "button" | "checkbox" | "dropdown" | "editbox" | "group-box" | "label" | "listbox" | "radio" | "scrollbar" | "spinner" | "unknown";
type XlsxFormControlState = "unchecked" | "checked" | "mixed";
type XlsxFormControlSelectionMode = "single" | "multi" | "extend";
interface XlsxFormControlCaptionRun {
    font?: XlsxCellFontStyleInput;
    text: string;
}
interface XlsxFormControlCaption {
    horizontalAlignment?: "general" | "left" | "center" | "right" | "fill" | "justify" | "centerContinuous" | "distributed";
    runs: XlsxFormControlCaptionRun[];
    verticalAlignment?: "top" | "center" | "bottom" | "justify" | "distributed";
}
type XlsxFormControlCaptionInput = string | XlsxFormControlCaption;
type XlsxFormControlKindInput = {
    kind: "button";
    caption: XlsxFormControlCaptionInput;
} | {
    kind: "checkbox";
    caption: XlsxFormControlCaptionInput;
    state: XlsxFormControlState;
    cellLink?: string;
    no3D?: boolean;
} | {
    kind: "optionButton";
    caption: XlsxFormControlCaptionInput;
    state: Exclude<XlsxFormControlState, "mixed">;
    cellLink?: string;
    no3D?: boolean;
} | {
    kind: "label";
    caption: XlsxFormControlCaptionInput;
} | {
    kind: "groupBox";
    caption: XlsxFormControlCaptionInput;
    no3D?: boolean;
} | {
    kind: "listBox";
    inputRange?: string;
    cellLink?: string;
    selection: XlsxFormControlSelectionMode;
    selected?: number[];
    no3D?: boolean;
} | {
    kind: "dropdown";
    inputRange?: string;
    cellLink?: string;
    selected?: number;
    lines: number;
    no3D?: boolean;
} | {
    kind: "editbox";
    caption: XlsxFormControlCaptionInput;
    legacyObjectType?: number;
    rawObj?: Uint8Array | number[];
    rawProperties?: Array<[string, string]>;
} | {
    kind: "scrollbar";
    value: number;
    min: number;
    max: number;
    increment: number;
    page: number;
    horizontal?: boolean;
    cellLink?: string;
} | {
    kind: "spinner";
    value: number;
    min: number;
    max: number;
    increment: number;
    cellLink?: string;
} | {
    kind: "unknown";
    objectType: string;
    legacyObjectType?: number;
    caption?: XlsxFormControlCaptionInput;
    rawObj?: Uint8Array | number[];
    rawProperties?: Array<[string, string]>;
};
interface XlsxFormControlMetadataInput {
    altText?: string;
    hidden?: boolean;
    kind: XlsxFormControlKindInput;
    locked?: boolean;
    macroName?: string;
    name?: string;
    printable?: boolean;
    rawClientData?: Array<Uint8Array | number[]>;
    title?: string;
}
type XlsxFormControlInput = XlsxFormControlMetadataInput & ({
    anchor: Extract<XlsxImageAnchor, {
        kind: "two-cell";
    }>;
    editAs?: "twoCell" | "oneCell" | "absolute";
} | {
    anchor: Exclude<XlsxImageAnchor, {
        kind: "two-cell";
    }>;
    editAs?: never;
});
type XlsxFormControlPatch = Partial<Omit<XlsxFormControlMetadataInput, "kind">> & {
    anchor?: XlsxImageAnchor;
    editAs?: "twoCell" | "oneCell" | "absolute";
    kind?: XlsxFormControlKindInput;
};
interface XlsxFormControl {
    anchor: XlsxImageAnchor;
    altText?: string;
    /** Original rich caption reported by Duke Sheets. */
    caption?: XlsxFormControlCaption;
    checked?: boolean;
    /** Duke Sheets' stable zero-based index for this worksheet load. */
    controlIndex?: number;
    /** Original two-cell anchor behavior. */
    editAs?: "twoCell" | "oneCell" | "absolute";
    firstInGroup?: boolean;
    fontFamily?: string;
    fontSizePt?: number;
    hidden?: boolean;
    horizontal?: boolean;
    id: string;
    increment?: number;
    inputRange?: string;
    /** Current formatted labels resolved from a dropdown or list box input range. */
    items?: string[];
    kind: XlsxFormControlKind;
    label?: string;
    lines?: number;
    linkedCell?: string;
    locked?: boolean;
    macroName?: string;
    max?: number;
    min?: number;
    name?: string;
    no3D?: boolean;
    page?: number;
    printable?: boolean;
    objectType?: string;
    legacyObjectType?: number;
    rawClientData?: number[][];
    rawObj?: number[];
    rawProperties?: Array<[string, string]>;
    /** Zero-based selected item index or indexes, matching Duke Sheets. */
    selected?: number | number[];
    selection?: XlsxFormControlSelectionMode;
    sheetIndex: number;
    state?: XlsxFormControlState;
    textAlign?: "center" | "left" | "right";
    textColor?: string;
    title?: string;
    value?: number;
    workbookSheetIndex: number;
    zIndex: number;
}
interface XlsxFormControlActionEvent {
    control: XlsxFormControl;
    type: "activate";
}
interface XlsxFormControlChangeEvent {
    control: XlsxFormControl;
    previousControl: XlsxFormControl;
    type: "change";
}
interface XlsxFormControlRenderProps {
    /** Activates a button control. No-ops in read-only mode. */
    activate: () => void;
    /** Current checked state for checkbox and radio controls. */
    checked: boolean;
    /** Full parsed control metadata. Switch on `control.kind` to render a kind-specific component. */
    control: XlsxFormControl;
    /** True when this control cannot be changed through the viewer. */
    disabled: boolean;
    /** Current formatted dropdown or list box items resolved from the control input range. */
    items: string[];
    /** Display label after placeholder Excel names have been removed. */
    label: string;
    /** True when a checkbox is in Excel's indeterminate state. */
    mixed: boolean;
    /** True when the viewer is rendering without workbook mutation support. */
    readOnly: boolean;
    /** Calculated worksheet-relative control rectangle in CSS pixels. */
    rect: XlsxImageRect;
    /** Persists a dropdown or list box selection. No-ops for other control kinds or in read-only mode. */
    setSelected: (selected: number | number[] | undefined) => boolean;
    /** Persists a checkbox or radio state. No-ops for other control kinds or in read-only mode. */
    setState: (state: XlsxFormControlState) => boolean;
    /** Persists a scrollbar or spinner value. No-ops for other control kinds or in read-only mode. */
    setValue: (value: number) => boolean;
    /** Stops an event from leaking into worksheet selection. */
    stopPropagation: (event: React.SyntheticEvent) => void;
    /** Apply this style to the custom control root to preserve worksheet positioning and sizing. */
    style: React.CSSProperties;
    /** Current worksheet zoom multiplier, where `1` is 100%. */
    zoomFactor: number;
}
interface XlsxChartReference {
    formula?: string;
    refType?: string;
    values?: Array<number | string | null>;
}
interface XlsxChartDataLabels {
    pointLabels?: XlsxChartPointDataLabel[];
    raw?: Record<string, unknown>;
    showBubbleSize?: boolean;
    showCategoryName?: boolean;
    showLegendKey?: boolean;
    showPercent?: boolean;
    showSeriesName?: boolean;
    showValue?: boolean;
}
interface XlsxChartPointDataLabel {
    deleted?: boolean;
    fontSizePt?: number;
    index: number;
    showBubbleSize?: boolean;
    showCategoryName?: boolean;
    showPercent?: boolean;
    showSeriesName?: boolean;
    showValue?: boolean;
    x?: number;
    y?: number;
}
interface XlsxChartLegend {
    overlay?: boolean;
    position?: string;
    raw?: Record<string, unknown>;
}
interface XlsxChartAxis {
    crossId?: number;
    crosses?: string;
    crossBetween?: string;
    delete?: boolean;
    id?: number;
    labelPosition?: string;
    logBase?: number;
    orientation?: string;
    majorUnit?: number;
    max?: number;
    min?: number;
    majorGridlines?: boolean;
    majorTickMark?: string;
    minorUnit?: number;
    minorGridlines?: boolean;
    minorTickMark?: string;
    numberFormat?: {
        formatCode?: string;
        sourceLinked?: boolean;
    };
    position?: string;
    raw?: Record<string, unknown>;
    shapeProperties?: Record<string, unknown>;
    tickLabelSkip?: number;
    tickMarkSkip?: number;
}
interface XlsxChartPointStyle {
    color?: string;
    explosion?: number;
    index: number;
    lineColor?: string;
}
interface XlsxChartSeries {
    bubbleSizeRef?: XlsxChartReference | null;
    bubbleSizes?: Array<number | null>;
    categories: Array<number | string | null>;
    categoriesRef?: XlsxChartReference | null;
    color?: string;
    dataPoints: unknown[];
    dataPointStyles?: XlsxChartPointStyle[];
    formatIdx?: number;
    hidden?: boolean;
    id: string;
    invertIfNegative?: boolean;
    lineColor?: string;
    lineWidthPx?: number;
    marker?: Record<string, unknown>;
    markerColor?: string;
    markerLineColor?: string;
    markerSize?: number;
    markerSymbol?: string;
    name?: string;
    negativeColor?: string;
    negativeLineColor?: string;
    raw?: Record<string, unknown>;
    shapeProperties?: Record<string, unknown>;
    smooth?: boolean;
    values: Array<number | null>;
    valuesRef?: XlsxChartReference | null;
}
type XlsxChartElementSelection = {
    kind: "chart";
    chartId: string;
} | {
    kind: "series";
    chartId: string;
    seriesId: string;
    seriesIndex: number;
} | {
    kind: "point";
    chartId: string;
    seriesId: string;
    seriesIndex: number;
    pointIndex: number;
} | {
    kind: "legendEntry";
    chartId: string;
    seriesId: string;
    seriesIndex: number;
};
type XlsxFormulaTarget = {
    kind: "cell";
    cell: XlsxCellAddress | null;
} | {
    kind: "chartSeries";
    chartId: string;
    seriesId: string;
    seriesIndex: number;
} | null;
interface XlsxChartTypeGroup {
    axisIds?: number[];
    chartType: string;
    dataLabels?: XlsxChartDataLabels | null;
    gapWidth?: number;
    is3d?: boolean;
    overlap?: number;
    raw?: Record<string, unknown>;
    series: XlsxChartSeries[];
    varyColors?: boolean;
}
interface XlsxChartWall {
    fillColor?: string;
    hidden?: boolean;
    lineColor?: string;
    thickness?: number;
}
interface XlsxChart {
    anchor: XlsxImageAnchor;
    autoTitleDeleted?: boolean;
    axes: XlsxChartAxis[];
    axisLabelColor?: string;
    axisLineColor?: string;
    categoryAxis?: XlsxChartAxis | null;
    chartExLayout?: string;
    chartAreaBorderColor?: string;
    chartAreaFillColor?: string;
    chartColorPalette?: string[];
    chartColorPaletteOffset?: number;
    chartPath?: string;
    chartStyleId?: number;
    chartType: string;
    dataLabels?: XlsxChartDataLabels | null;
    displayBlanksAs?: string;
    editable?: boolean;
    firstSliceAngle?: number;
    fontFamily?: string;
    gapWidth?: number;
    holeSize?: number;
    id: string;
    is3d?: boolean;
    legend?: XlsxChartLegend | null;
    name?: string;
    overlap?: number;
    plotVisibleOnly?: boolean;
    raw?: Record<string, unknown>;
    radarStyle?: string;
    scatterStyle?: string;
    roundedCorners?: boolean;
    shape3d?: string;
    seriesAxis?: XlsxChartAxis | null;
    series: XlsxChartSeries[];
    sheetIndex: number;
    showDlblsOverMax?: boolean;
    sideWall?: XlsxChartWall | null;
    backWall?: XlsxChartWall | null;
    bubbleScale?: number;
    bubble3d?: boolean;
    floor?: XlsxChartWall | null;
    surfaceMaterial?: string;
    textColor?: string;
    title?: string;
    titleColor?: string;
    titleFontFamily?: string;
    typeGroups?: XlsxChartTypeGroup[];
    valueAxis?: XlsxChartAxis | null;
    varyColors?: boolean;
    view3d?: {
        depthPercent?: number;
        perspective?: number;
        rAngAx?: boolean;
        rotX?: number;
        rotY?: number;
    };
    wireframe?: boolean;
    workbookSheetIndex: number;
    zIndex: number;
}
interface XlsxChartsheet {
    chartIds: string[];
    chartPath?: string;
    id: string;
    index: number;
    name: string;
    raw?: Record<string, unknown>;
    workbookSheetIndex?: number;
}
interface XlsxWorkbookTab {
    chartsheetIndex?: number;
    id: string;
    index: number;
    kind: "chartsheet" | "sheet";
    name: string;
    sheetIndex?: number;
    visibility?: XlsxSheetVisibility;
    workbookSheetIndex?: number;
}
interface XlsxImageRect {
    height: number;
    left: number;
    top: number;
    width: number;
}
/** Rendered column widths and row heights in logical pixels, indexed by worksheet coordinates. */
interface XlsxDrawingLayout {
    columnWidths: readonly number[];
    rowHeights: readonly number[];
}
type XlsxImageResizeHandlePosition = "n" | "s" | "e" | "w" | "ne" | "nw" | "se" | "sw";
interface XlsxImageRenderProps {
    /** The built-in image element that react-xlsx would render without customization. */
    defaultNode: React.ReactNode;
    /** Workbook image metadata, including source, anchor, name, alt text, and editability. */
    image: XlsxImage;
    /** The image rectangle in viewer pixels, including the current zoom level. */
    rect: XlsxImageRect;
    /** Absolute positioning styles to apply when rendering a replacement image node. */
    style: React.CSSProperties;
}
interface XlsxImageSelectionRenderProps {
    /** The built-in selected-image outline and resize handles. */
    defaultNode: React.ReactNode;
    /** Returns pointer handlers and styles for a custom resize handle. */
    getHandleProps: (position: XlsxImageResizeHandlePosition) => {
        onPointerDown: (event: React.PointerEvent<HTMLElement>) => void;
        style: React.CSSProperties;
    };
    /** The currently selected image. */
    image: XlsxImage;
    /** The selected image rectangle in viewer pixels, including the current zoom level. */
    rect: XlsxImageRect;
}
interface XlsxChartLoadingRenderProps {
    /** The chart that is waiting for its renderer or data. */
    chart: XlsxChart;
    /** The built-in chart loading placeholder. */
    defaultNode: React.ReactNode;
    /** The chart rectangle in viewer pixels, including the current zoom level. */
    rect: XlsxImageRect;
}
interface XlsxFileTooLargeRenderProps {
    /** The built-in file-too-large message. */
    defaultNode: React.ReactNode;
    /** File name displayed in the viewer UI. */
    displayFileName: string;
    /** Actual file size in bytes. */
    fileSizeBytes: number;
    /** Configured maximum file size in bytes. */
    maxFileSizeBytes: number;
}
interface XlsxScrollerRenderProps {
    /** Workbook grid content that should be rendered inside the scrollable viewport. */
    children: React.ReactNode;
    /** Props that must be applied to the actual scrollable viewport element. */
    viewportProps: React.HTMLAttributes<HTMLDivElement> & {
        ref: React.Ref<HTMLDivElement>;
        style: React.CSSProperties;
        tabIndex: number;
    };
}
interface UseXlsxViewerControllerOptions {
    /**
     * Allows row and column resizing even while editing is disabled by `readOnly`.
     *
     * @default false
     */
    allowResizeInReadOnly?: boolean;
    /**
     * Values keyed by `externalCallKey(name, args)` resolve add-in formulas without rewriting them.
     * Unmapped calls preserve their cached values.
     */
    externalFnValues?: Record<string, string | number>;
    /**
     * Defers loading until `continueDeferredLoad()` is called when the file is larger than this byte threshold.
     * Set to `0` to parse immediately.
     *
     * @default 0
     * @example
     * ```tsx
     * useXlsxViewerController({ file, deferLoadingAboveBytes: 10 * 1024 * 1024 })
     * ```
     */
    deferLoadingAboveBytes?: number;
    /**
     * Local workbook bytes to load. Use either `file` or `src`.
     *
     * @example
     * ```tsx
     * <XlsxViewer file={await uploadedFile.arrayBuffer()} fileName={uploadedFile.name} />
     * ```
     */
    file?: ArrayBuffer;
    /**
     * Optional display and download name for the workbook.
     *
     * @example
     * ```tsx
     * <XlsxViewer file={buffer} fileName="forecast.xlsx" />
     * ```
     */
    fileName?: string;
    /**
     * Rejects files larger than this byte limit and renders `fileTooLargeState`.
     *
     * @default 25 * 1024 * 1024
     */
    maxFileSizeBytes?: number;
    /**
     * Disables workbook edits, paste, fill, undo/redo, and other mutation actions.
     *
     * @default false
     */
    readOnly?: boolean;
    /**
     * Automatically enables read-only mode for files larger than this byte threshold.
     * Set to `0` to disable the automatic switch.
     *
     * @default 0
     */
    readOnlyAboveBytes?: number;
    /**
     * Includes hidden and very hidden workbook sheets in the sheet tabs and controller state.
     *
     * @default false
     */
    showHiddenSheets?: boolean;
    /**
     * Skips OOXML ZIP/XML parsing and relies on `@dukelib/sheets-wasm` metadata only.
     * This is automatically used for legacy `.xls` files.
     *
     * @default false
     */
    skipXmlParsing?: boolean;
    /**
     * Remote workbook URL to fetch and load. Use either `src` or `file`.
     *
     * @example
     * ```tsx
     * <XlsxViewer src="/workbooks/report.xlsx" />
     * ```
     */
    src?: string;
    /**
     * Parses supported workbook data in a Web Worker so large files do not block React rendering.
     *
     * @default true
     */
    useWorker?: boolean;
}
interface XlsxViewerController {
    activeCell: XlsxCellAddress | null;
    activeCellAddress: string | null;
    activeSheet: XlsxSheetData | null;
    activeSheetIndex: number;
    activeTab: XlsxWorkbookTab | null;
    activeTabIndex: number;
    canDownload: boolean;
    canExport: boolean;
    canLoadDeferred: boolean;
    canRedo: boolean;
    canUndo: boolean;
    canZoomIn: boolean;
    canZoomOut: boolean;
    /** Adds a Duke-supported form control to a visible worksheet and returns its worksheet-local index. */
    addFormControl: (input: XlsxFormControlInput, sheetIndex?: number) => number | null;
    clearSelectedCells: () => void;
    clearSelection: () => void;
    continueDeferredLoad: () => void;
    copySelectionToClipboard: () => Promise<boolean>;
    defaultZoomScale: number;
    deferredLoadFileSize: number | null;
    defineNamedRange: (name: string, range?: XlsxCellRange | null) => void;
    displayFileName: string;
    download: () => void;
    charts: XlsxChart[];
    chartsheets: XlsxChartsheet[];
    exportCsv: () => void;
    exportXlsx: () => void;
    error: Error | null;
    file?: ArrayBuffer;
    fillSelection: (targetRange: XlsxCellRange) => void;
    clearSelectedChart: () => void;
    clearSelectedChartElement: () => void;
    clearSelectedImage: () => void;
    getChartById: (id: string) => XlsxChart | null;
    getChartSeriesFormula: (chartId: string, seriesIndex: number) => string;
    getChartsheetById: (id: string) => XlsxChartsheet | null;
    formControls: XlsxFormControl[];
    getSheetCharts: (sheetIndex?: number) => XlsxChart[];
    getSheetFormControls: (sheetIndex?: number) => XlsxFormControl[];
    /** Resolves a list/dropdown input range to its current formatted item labels. */
    getFormControlItems: (controlIndex: number, sheetIndex?: number) => string[];
    getImageById: (id: string) => XlsxImage | null;
    getSheetImages: (sheetIndex?: number) => XlsxImage[];
    getSheetShapes: (sheetIndex?: number) => XlsxShape[];
    getClipboardData: () => XlsxClipboardData | null;
    getCellDisplayValue: (cell?: XlsxCellAddress | null) => string;
    getCellFormula: (cell?: XlsxCellAddress | null) => string;
    getCellSnapshotAsync?: (workbookSheetIndex: number, row: number, col: number) => Promise<{
        displayValue: string;
        formula: string;
    }>;
    isLoadDeferred: boolean;
    isLoading: boolean;
    isChartsLoading: boolean;
    isWorkerBacked?: boolean;
    images: XlsxImage[];
    moveChartBy: (id: string, deltaX: number, deltaY: number) => void;
    shapes: XlsxShape[];
    mergeSelection: () => void;
    maxZoomScale: number;
    minZoomScale: number;
    moveImageBy: (id: string, deltaX: number, deltaY: number) => void;
    removeActiveSheet: () => void;
    /** Removes a Duke-supported control by its worksheet-local index. */
    removeFormControl: (controlIndex: number, sheetIndex?: number) => boolean;
    readOnly: boolean;
    /** Recalculates formulas, optionally resolving external add-in calls. */
    recalculate: (externalFnValues?: ExternalFnValues) => void;
    revision: number;
    resetZoom: () => void;
    resizeChartBy: (id: string, handle: XlsxImageResizeHandlePosition, deltaX: number, deltaY: number) => void;
    resizeImageBy: (id: string, handle: XlsxImageResizeHandlePosition, deltaX: number, deltaY: number) => void;
    resizeColumn: (col: number, widthPx: number) => void;
    resizeRow: (row: number, heightPx: number) => void;
    /** Resizes several columns as one undo step, e.g. fitting every selected column. */
    resizeColumns: (sizes: ReadonlyArray<XlsxAxisSize>) => void;
    /** Resizes several rows as one undo step, e.g. fitting every selected row. */
    resizeRows: (sizes: ReadonlyArray<XlsxAxisSize>) => void;
    redo: () => void;
    pasteFromClipboard: () => Promise<boolean>;
    pasteStructuredClipboardData: (payload: string) => boolean;
    pasteText: (text: string) => boolean;
    selectedRangeAddress: string | null;
    selectedValue: string;
    selectedFormula: string;
    /**
     * Sets the formula for a cell in the active worksheet.
     *
     * Pass an A1-style formula body or formula string, for example `"SUM(A1:A3)"`
     * or `"=SUM(A1:A3)"`. The mutation is recorded in undo history and is included
     * in exported workbook bytes.
     *
     * @param cell Zero-based row and column address in the active worksheet.
     * @param formula Formula text to write. Pass an empty string to clear the formula.
     */
    setCellFormula: (cell: XlsxCellAddress, formula: string) => void;
    /**
     * Applies persisted Excel style properties to a cell in the active worksheet.
     *
     * This mutates workbook data, records undo history, refreshes the viewer, and
     * is included in exported XLSX bytes. Use the `getCellStyle` viewer prop only
     * for render-time CSS overrides that should not modify the workbook.
     *
     * @param cell Zero-based row and column address in the active worksheet.
     * @param style Partial Excel style patch to apply to the target cell.
     */
    setCellStyle: (cell: XlsxCellAddress, style: XlsxCellStyleInput) => void;
    /**
     * Sets the value for a cell in the active worksheet.
     *
     * User-entered strings are coerced to booleans or numbers when they match
     * Excel-like input. Prefix text with an apostrophe to force literal text.
     *
     * @param cell Zero-based row and column address in the active worksheet.
     * @param value Text entered by the user.
     */
    setCellValue: (cell: XlsxCellAddress, value: string) => void;
    /**
     * Applies persisted Excel style properties to every cell in a range.
     *
     * This mutates workbook data, records undo history, refreshes the viewer, and
     * is included in exported XLSX bytes.
     *
     * @param range Zero-based cell range in the active worksheet.
     * @param style Partial Excel style patch to apply to the target range.
     */
    setRangeStyle: (range: XlsxCellRange, style: XlsxCellStyleInput) => void;
    setZoomScale: (zoomScale: number) => void;
    selectCell: (cell: XlsxCellAddress, options?: {
        extend?: boolean;
        append?: boolean;
    }) => void;
    /** Resolve a data-block boundary without changing selection; worker sheets stay off-thread. */
    findDataBoundary: (request: DataNavigationRequest) => Promise<XlsxCellAddress>;
    /** Select a cell and center it in the viewport when it is off-screen. */
    revealCell: (cell: XlsxCellAddress) => void;
    /** @internal Grid wiring: XlsxGrid registers revealCell's scroll + overlay impl here. */
    registerRevealCellImpl: (impl: ((cell: XlsxCellAddress) => void) | null) => void;
    selectChart: (id: string | null) => void;
    selectChartElement: (selection: XlsxChartElementSelection | null) => void;
    selectRange: (range: XlsxCellRange, options?: {
        append?: boolean;
        toggle?: boolean;
    }) => void;
    selection: XlsxCellRange | null;
    /** All selected regions; `selection` is the active, last region. */
    selections: XlsxCellRange[];
    setActiveSheetIndex: (index: number) => void;
    setActiveTabIndex: (index: number) => void;
    selectedChart: XlsxChart | null;
    selectedChartElement: XlsxChartElementSelection | null;
    selectedChartFormula: string | null;
    selectedChartId: string | null;
    selectedCellFormula: string;
    selectedFormulaTarget: XlsxFormulaTarget;
    selectedImage: XlsxImage | null;
    selectedImageId: string | null;
    setSelectedFormula: (formula: string) => boolean;
    setSelectedCellFormula: (formula: string) => void;
    /**
     * Applies persisted Excel style properties to the active cell.
     *
     * No-ops when there is no active cell, when editing is disabled, or when no
     * workbook is loaded. The mutation is recorded in undo history and is included
     * in exported XLSX bytes.
     *
     * @param style Partial Excel style patch to apply to the active cell.
     */
    setSelectedCellStyle: (style: XlsxCellStyleInput) => void;
    setSelectedCellValue: (value: string) => void;
    sheets: XlsxSheetData[];
    src?: string;
    sortState: XlsxTableSortState | null;
    sortTable: (tableName: string, columnIndex: number, direction: XlsxTableSortDirection) => void;
    setChartSeriesFormula: (chartId: string, seriesIndex: number, formula: string) => boolean;
    selectImage: (id: string | null) => void;
    setChartRect: (id: string, rect: XlsxImageRect, layout?: XlsxDrawingLayout) => void;
    setImageRect: (id: string, rect: XlsxImageRect, layout?: XlsxDrawingLayout) => void;
    getRowsBatchAsync?: (workbookSheetIndex: number, startRow: number, rowCount: number) => Promise<unknown[] | null>;
    tables: XlsxTable[];
    tabs: XlsxWorkbookTab[];
    undo: () => void;
    unmergeSelection: () => void;
    updateChart: (id: string, patch: Partial<XlsxChart>) => void;
    /** Updates a Duke-supported control and participates in undo/redo and export. */
    updateFormControl: (controlIndex: number, patch: XlsxFormControlPatch, sheetIndex?: number) => boolean;
    workbook: Workbook | null;
    zoomIn: () => void;
    zoomOut: () => void;
    zoomScale: number;
    getActiveWorksheet: () => Worksheet | null;
    addSheet: (name?: string) => void;
}
interface XlsxViewerSelection {
    activeCell: XlsxCellAddress | null;
    activeCellAddress: string | null;
    clearSelection: () => void;
    selectedRangeAddress: string | null;
    selectCell: (cell: XlsxCellAddress, options?: {
        extend?: boolean;
        append?: boolean;
    }) => void;
    selectRange: (range: XlsxCellRange, options?: {
        append?: boolean;
        toggle?: boolean;
    }) => void;
    selection: XlsxCellRange | null;
    /** All selected regions; `selection` is the active/last one. Length 1 for a normal selection. */
    selections: XlsxCellRange[];
}
interface XlsxViewerZoom {
    canZoomIn: boolean;
    canZoomOut: boolean;
    defaultZoomScale: number;
    maxZoomScale: number;
    minZoomScale: number;
    resetZoom: () => void;
    setZoomScale: (zoomScale: number) => void;
    zoomIn: () => void;
    zoomOut: () => void;
    zoomScale: number;
}
interface XlsxViewerEditing {
    addSheet: (name?: string) => void;
    addFormControl: (input: XlsxFormControlInput, sheetIndex?: number) => number | null;
    canRedo: boolean;
    canUndo: boolean;
    clearSelectedCells: () => void;
    copySelectionToClipboard: () => Promise<boolean>;
    defineNamedRange: (name: string, range?: XlsxCellRange | null) => void;
    fillSelection: (targetRange: XlsxCellRange) => void;
    formControls: XlsxFormControl[];
    getClipboardData: () => XlsxClipboardData | null;
    getCellDisplayValue: (cell?: XlsxCellAddress | null) => string;
    getCellFormula: (cell?: XlsxCellAddress | null) => string;
    getFormControlItems: (controlIndex: number, sheetIndex?: number) => string[];
    getSheetFormControls: (sheetIndex?: number) => XlsxFormControl[];
    mergeSelection: () => void;
    pasteFromClipboard: () => Promise<boolean>;
    pasteStructuredClipboardData: (payload: string) => boolean;
    pasteText: (text: string) => boolean;
    removeActiveSheet: () => void;
    removeFormControl: (controlIndex: number, sheetIndex?: number) => boolean;
    readOnly: boolean;
    redo: () => void;
    selectedCellFormula: string;
    selectedChartFormula: string | null;
    selectedFormula: string;
    selectedFormulaTarget: XlsxFormulaTarget;
    selectedValue: string;
    /**
     * Sets the formula for a cell in the active worksheet.
     *
     * @param cell Zero-based row and column address in the active worksheet.
     * @param formula Formula text to write. Pass an empty string to clear the formula.
     */
    setCellFormula: (cell: XlsxCellAddress, formula: string) => void;
    /**
     * Applies persisted Excel style properties to a cell in the active worksheet.
     *
     * The mutation is recorded in undo history and is included in exported XLSX
     * bytes. This is different from the `getCellStyle` render override prop.
     *
     * @param cell Zero-based row and column address in the active worksheet.
     * @param style Partial Excel style patch to apply to the target cell.
     */
    setCellStyle: (cell: XlsxCellAddress, style: XlsxCellStyleInput) => void;
    /**
     * Sets the value for a cell in the active worksheet.
     *
     * @param cell Zero-based row and column address in the active worksheet.
     * @param value Text entered by the user.
     */
    setCellValue: (cell: XlsxCellAddress, value: string) => void;
    /**
     * Applies persisted Excel style properties to every cell in a range.
     *
     * @param range Zero-based cell range in the active worksheet.
     * @param style Partial Excel style patch to apply to the target range.
     */
    setRangeStyle: (range: XlsxCellRange, style: XlsxCellStyleInput) => void;
    setSelectedFormula: (formula: string) => boolean;
    setSelectedCellFormula: (formula: string) => void;
    /**
     * Applies persisted Excel style properties to the active cell.
     *
     * @param style Partial Excel style patch to apply to the active cell.
     */
    setSelectedCellStyle: (style: XlsxCellStyleInput) => void;
    setSelectedCellValue: (value: string) => void;
    undo: () => void;
    unmergeSelection: () => void;
    updateFormControl: (controlIndex: number, patch: XlsxFormControlPatch, sheetIndex?: number) => boolean;
}
interface XlsxViewerTables {
    sortState: XlsxTableSortState | null;
    sortTable: (tableName: string, columnIndex: number, direction: XlsxTableSortDirection) => void;
    tables: XlsxTable[];
}
interface XlsxViewerImages {
    charts: XlsxChart[];
    clearSelectedChart: () => void;
    clearSelectedChartElement: () => void;
    clearSelectedImage: () => void;
    getChartById: (id: string) => XlsxChart | null;
    getChartSeriesFormula: (chartId: string, seriesIndex: number) => string;
    getSheetCharts: (sheetIndex?: number) => XlsxChart[];
    getImageById: (id: string) => XlsxImage | null;
    getSheetImages: (sheetIndex?: number) => XlsxImage[];
    images: XlsxImage[];
    moveChartBy: (id: string, deltaX: number, deltaY: number) => void;
    moveImageBy: (id: string, deltaX: number, deltaY: number) => void;
    isChartsLoading: boolean;
    readOnly: boolean;
    resizeChartBy: (id: string, handle: XlsxImageResizeHandlePosition, deltaX: number, deltaY: number) => void;
    resizeImageBy: (id: string, handle: XlsxImageResizeHandlePosition, deltaX: number, deltaY: number) => void;
    selectedChart: XlsxChart | null;
    selectedChartElement: XlsxChartElementSelection | null;
    selectedChartFormula: string | null;
    selectedChartId: string | null;
    selectedImage: XlsxImage | null;
    selectedImageId: string | null;
    selectChart: (id: string | null) => void;
    selectChartElement: (selection: XlsxChartElementSelection | null) => void;
    selectImage: (id: string | null) => void;
    setChartSeriesFormula: (chartId: string, seriesIndex: number, formula: string) => boolean;
    setChartRect: (id: string, rect: XlsxImageRect, layout?: XlsxDrawingLayout) => void;
    setImageRect: (id: string, rect: XlsxImageRect, layout?: XlsxDrawingLayout) => void;
    updateChart: (id: string, patch: Partial<XlsxChart>) => void;
}
interface XlsxViewerCharts {
    activeTab: XlsxWorkbookTab | null;
    activeTabIndex: number;
    charts: XlsxChart[];
    chartsheets: XlsxChartsheet[];
    clearSelectedChart: () => void;
    clearSelectedChartElement: () => void;
    getChartById: (id: string) => XlsxChart | null;
    getChartSeriesFormula: (chartId: string, seriesIndex: number) => string;
    getChartsheetById: (id: string) => XlsxChartsheet | null;
    getSheetCharts: (sheetIndex?: number) => XlsxChart[];
    isChartsLoading: boolean;
    moveChartBy: (id: string, deltaX: number, deltaY: number) => void;
    readOnly: boolean;
    resizeChartBy: (id: string, handle: XlsxImageResizeHandlePosition, deltaX: number, deltaY: number) => void;
    selectChart: (id: string | null) => void;
    selectedChart: XlsxChart | null;
    selectedChartElement: XlsxChartElementSelection | null;
    selectedChartFormula: string | null;
    selectedChartId: string | null;
    selectChartElement: (selection: XlsxChartElementSelection | null) => void;
    setChartSeriesFormula: (chartId: string, seriesIndex: number, formula: string) => boolean;
    setActiveTabIndex: (index: number) => void;
    setChartRect: (id: string, rect: XlsxImageRect, layout?: XlsxDrawingLayout) => void;
    tabs: XlsxWorkbookTab[];
    updateChart: (id: string, patch: Partial<XlsxChart>) => void;
}
type XlsxSheetThumbnailResolution = number | {
    maxHeight?: number;
    maxWidth?: number;
};
interface UseXlsxViewerThumbnailsOptions {
    includeHeaders?: boolean;
    resolution?: XlsxSheetThumbnailResolution;
}
interface XlsxSheetThumbnail {
    aspectRatio: number;
    contentHeight: number;
    contentWidth: number;
    height: number;
    paint: (canvas: HTMLCanvasElement | null) => boolean;
    sheet: XlsxSheetData;
    sheetIndex: number;
    sourceRange: XlsxCellRange;
    width: number;
    workbookSheetIndex: number;
}
interface XlsxViewerThumbnails {
    paintThumbnail: (sheetIndex: number, canvas: HTMLCanvasElement | null) => boolean;
    thumbnails: XlsxSheetThumbnail[];
}
interface XlsxTableHeaderMenuRenderProps {
    /** Address of the table header cell that owns this menu. */
    cell: XlsxCellAddress;
    /** Table column metadata for the active header. */
    column: XlsxTableColumn;
    /** Current sort direction for this column, or `null` when unsorted. */
    direction: XlsxTableSortDirection | null;
    /** Sorts the table by this column in ascending order. */
    sortAscending: () => void;
    /** Sorts the table by this column in descending order. */
    sortDescending: () => void;
    /** Table metadata for the active header. */
    table: XlsxTable;
    /** Built-in trigger text/icon. Useful when composing your own trigger button. */
    triggerIcon: string;
    /** Props that must be applied to your menu trigger button so grid selection does not receive the click. */
    triggerProps: React.ButtonHTMLAttributes<HTMLButtonElement>;
}
interface XlsxCellStyleContext {
    /** Address of the cell being styled. */
    cell: XlsxCellAddress;
    /** True when the cell is part of a selected chart's highlighted source range. */
    hasChartHighlight: boolean;
    /** True when a conditional format (color scale, data bar, or icon set) applies to the cell. */
    hasConditionalFormat: boolean;
    /** True when the cell has a hyperlink. */
    hasHyperlink: boolean;
    /** True when the cell has a data validation rule. */
    hasValidation: boolean;
    /** True when the cell is the anchor of a merged range. */
    isMerged: boolean;
    /** True when the cell is a table header cell. */
    isTableHeader: boolean;
    /**
     * The fully resolved style the viewer computed for this cell, combining the workbook's
     * own formatting with the viewer's built-in styling. Treat this as read-only; returning
     * a partial style from `getCellStyle` merges on top of it.
     */
    resolvedStyle: React.CSSProperties;
    /** Display name of the sheet the cell belongs to. */
    sheetName: string;
    /** The cell's resolved display value. */
    value: string;
    /** Workbook sheet index of the cell's sheet. */
    workbookSheetIndex: number;
}
interface XlsxViewerProviderProps extends UseXlsxViewerControllerOptions {
    /** Viewer UI and hooks that should share this provider's workbook controller. */
    children: React.ReactNode;
    /**
     * Existing controller to provide instead of creating one from `file`, `src`, or other controller options.
     *
     * @example
     * ```tsx
     * const controller = useXlsxViewerController({ src: "/report.xlsx" });
     * return <XlsxViewerProvider controller={controller}><XlsxViewer /></XlsxViewerProvider>;
     * ```
     */
    controller?: XlsxViewerController;
    /**
     * Uses the built-in dark worksheet/viewer palette.
     *
     * @default false
     */
    isDark?: boolean;
}
interface XlsxViewerProps extends UseXlsxViewerControllerOptions {
    /**
     * Allows row and column resizing even while editing is disabled by `readOnly`.
     *
     * @default false
     */
    allowResizeInReadOnly?: boolean;
    /** Class name applied to the root viewer shell. */
    className?: string;
    /**
     * Existing controller to render. Takes precedence over provider context and source props on this viewer.
     *
     * @example
     * ```tsx
     * const controller = useXlsxViewerController({ file, fileName: "model.xlsx" });
     * return <XlsxViewer controller={controller} />;
     * ```
     */
    controller?: XlsxViewerController;
    /** Content shown when no workbook has been loaded. */
    emptyState?: React.ReactNode;
    /**
     * Animates the selection rectangle in canvas mode.
     *
     * @default true
     */
    enableCanvasSelectionAnimation?: boolean;
    /**
     * Enables pinch zoom and modifier-key wheel zoom inside the viewer.
     *
     * @default true
     */
    enableGestureZoom?: boolean;
    /**
     * Renders worksheets with the canvas renderer. Set to `false` to use the DOM table renderer.
     *
     * @default true
     * @example
     * ```tsx
     * <XlsxViewer experimentalCanvas={false} />
     * ```
     */
    experimentalCanvas?: boolean;
    /** Content shown for non-size load errors, or a function that receives the thrown error. */
    errorState?: React.ReactNode | ((error: Error) => React.ReactNode);
    /** Content shown when `maxFileSizeBytes` rejects a file. */
    fileTooLargeState?: React.ReactNode | ((props: XlsxFileTooLargeRenderProps) => React.ReactNode);
    /**
     * Returns extra CSS style overrides for an individual cell. Called for every rendered cell
     * with the cell's address, sheet, resolved style, and contextual flags. Return `undefined`
     * (or `null`) to leave a cell untouched, or a partial style object that merges on top of the
     * viewer's resolved style. Use this as an escape hatch for custom per-cell styling such as
     * highlights, outlines, or status tints, without forking the workbook data.
     *
     * The DOM renderer applies every returned CSS property. The canvas renderer
     * (`experimentalCanvas`, the default) honors the subset it can paint: `backgroundColor`,
     * `backgroundImage` gradients, `color`, the four `border*` sides, `padding`, `textAlign`,
     * `textDecoration`, `textOverflow`, and font properties. CSS-only effects such as `boxShadow`,
     * `outline`, or `animation` apply in the DOM renderer (`experimentalCanvas={false}`).
     *
     * Keep this callback stable (e.g. wrap it in `useCallback`) so cell styling is not recomputed
     * on every render. When the callback identity changes, the viewer re-resolves and repaints
     * cell styles.
     *
     * @example
     * ```tsx
     * const getCellStyle = React.useCallback<NonNullable<XlsxViewerProps["getCellStyle"]>>(
     *   ({ cell }) => (cell.row === 3 && cell.col === 1 ? { backgroundColor: "#dbeafe" } : undefined),
     *   []
     * );
     * return <XlsxViewer file={buffer} getCellStyle={getCellStyle} />;
     * ```
     */
    getCellStyle?: (context: XlsxCellStyleContext) => React.CSSProperties | null | undefined;
    /** Display formula text while retaining calculated values for workbook operations. */
    showFormulas?: boolean;
    /** Background color used for row-number, column-letter, and corner headers. */
    headerBackgroundColor?: string;
    /** Text color used for row-number and column-letter headers. */
    headerTextColor?: string;
    /**
     * CSS height for the viewer container.
     *
     * @example
     * ```tsx
     * <XlsxViewer height="70vh" />
     * ```
     */
    height?: React.CSSProperties["height"];
    /**
     * Uses the built-in dark worksheet/viewer palette.
     *
     * @default false
     */
    isDark?: boolean;
    /** Legacy loading element rendered while the workbook is being parsed. Prefer `loadingState` for new code. */
    loadingComponent?: React.ReactElement;
    /** Content shown while the workbook is being parsed. */
    loadingState?: React.ReactNode;
    /** Called when an editable button form control is activated. Duke preserves buttons but does not execute macros. */
    onFormControlAction?: (event: XlsxFormControlActionEvent) => void;
    /** Called after the built-in editor persists a checkbox, radio, list, dropdown, scrollbar, or spinner change. */
    onFormControlChange?: (event: XlsxFormControlChangeEvent) => void;
    /** Replaces the chart loading placeholder. */
    renderChartLoading?: (props: XlsxChartLoadingRenderProps) => React.ReactNode;
    /**
     * Replaces built-in worksheet form-control rendering.
     * Apply `style` to the custom root and use `stopPropagation` on pointer/click handlers.
     * The supplied setters preserve Duke mutation, linked-cell, undo/redo, and change-event behavior.
     */
    renderFormControl?: (props: XlsxFormControlRenderProps) => React.ReactNode;
    /**
     * Replaces worksheet images while preserving their calculated rect and style.
     *
     * @example
     * ```tsx
     * <XlsxViewer renderImage={({ image, style }) => <img src={image.src} style={style} alt={image.description ?? ""} />} />
     * ```
     */
    renderImage?: (props: XlsxImageRenderProps) => React.ReactNode;
    /** Replaces the selected-image outline and resize handles. */
    renderImageSelection?: (props: XlsxImageSelectionRenderProps) => React.ReactNode;
    /**
     * Replaces the scroll viewport shell. Spread `viewportProps` onto the actual scrollable viewport element.
     *
     * @example
     * ```tsx
     * <XlsxViewer
     *   renderScroller={({ children, viewportProps }) => (
     *     <ScrollAreaPrimitive.Root className="h-full min-h-0 w-full min-w-0 flex-1">
     *       <ScrollAreaPrimitive.Viewport {...viewportProps}>{children}</ScrollAreaPrimitive.Viewport>
     *       <ScrollBar />
     *       <ScrollAreaPrimitive.Corner />
     *     </ScrollAreaPrimitive.Root>
     *   )}
     * />
     * ```
     */
    renderScroller?: (props: XlsxScrollerRenderProps) => React.ReactNode;
    /**
     * Toggles the default rounded outer shell.
     *
     * @default true
     */
    rounded?: boolean;
    /**
     * Disables workbook edits, paste, fill, undo/redo, and other mutation actions.
     *
     * @default false
     */
    readOnly?: boolean;
    /**
     * Border and accent color for the current selection.
     *
     * @example
     * ```tsx
     * <XlsxViewer selectionColor="#2563eb" />
     * ```
     */
    selectionColor?: string;
    /** Fill color used for the translucent selection overlay. */
    selectionFillColor?: string;
    /** Accent color used for selected row and column headers. */
    selectionHeaderColor?: string;
    /**
     * Replaces the table header sort/filter trigger menu.
     * Apply `triggerProps` to your actual trigger button so clicks do not leak into grid selection.
     */
    renderTableHeaderMenu?: (props: XlsxTableHeaderMenuRenderProps) => React.ReactNode;
    /**
     * Shows worksheet images, charts, shapes, and form controls.
     *
     * @default true
     */
    showImages?: boolean;
    /**
     * Shows the built-in toolbar above the workbook grid.
     *
     * @default true
     */
    showDefaultToolbar?: boolean;
    /**
     * Replaces the toolbar area with a node or render function.
     *
     * @example
     * ```tsx
     * <XlsxViewer toolbar={(controller) => <button onClick={controller.download}>Download</button>} />
     * ```
     */
    toolbar?: React.ReactNode | ((controller: XlsxViewerController) => React.ReactNode);
}

declare class XlsxFileSizeLimitExceededError extends Error {
    fileSizeBytes: number;
    maxFileSizeBytes: number;
    constructor(fileSizeBytes: number, maxFileSizeBytes: number);
}
declare function useXlsxViewerController(options: UseXlsxViewerControllerOptions): XlsxViewerController;

type XlsxWasmSource = string | URL | Request | Response | BufferSource | WebAssembly.Module;
declare function setWasmSource(source: XlsxWasmSource): void;
declare function initWasm(source?: XlsxWasmSource): Promise<typeof _dukelib_sheets_wasm>;

declare function XlsxViewerProvider({ children, controller, isDark, ...options }: XlsxViewerProviderProps): react_jsx_runtime.JSX.Element;
declare function useXlsxViewer(): XlsxViewerController;
declare function useXlsxViewerSelection(): XlsxViewerSelection;
declare function useXlsxViewerZoom(): XlsxViewerZoom;
declare function useXlsxViewerEditing(): XlsxViewerEditing;
declare function useXlsxViewerTables(): XlsxViewerTables;
declare function useXlsxViewerImages(): XlsxViewerImages;
declare function useXlsxViewerCharts(): XlsxViewerCharts;
declare function useXlsxViewerThumbnails(options?: UseXlsxViewerThumbnailsOptions): XlsxViewerThumbnails;
declare function XlsxViewer(props: XlsxViewerProps): react_jsx_runtime.JSX.Element;
declare function DefaultXlsxToolbar(): react_jsx_runtime.JSX.Element;

export { DefaultXlsxToolbar, type ExternalFnValues, type UseXlsxViewerControllerOptions, type UseXlsxViewerThumbnailsOptions, type XlsxAxisSize, type XlsxCellAddress, type XlsxCellAlignmentInput, type XlsxCellBorderEdgeInput, type XlsxCellBorderStyleInput, type XlsxCellFillStyleInput, type XlsxCellFontStyleInput, type XlsxCellGradientStopInput, type XlsxCellNumberFormatInput, type XlsxCellProtectionInput, type XlsxCellRange, type XlsxCellStyleColorInput, type XlsxCellStyleContext, type XlsxCellStyleInput, type XlsxChart, type XlsxChartAxis, type XlsxChartDataLabels, type XlsxChartElementSelection, type XlsxChartLoadingRenderProps, type XlsxChartReference, type XlsxChartSeries, type XlsxChartsheet, type XlsxDrawingLayout, XlsxFileSizeLimitExceededError, type XlsxFileTooLargeRenderProps, type XlsxFormControl, type XlsxFormControlActionEvent, type XlsxFormControlCaption, type XlsxFormControlCaptionInput, type XlsxFormControlCaptionRun, type XlsxFormControlChangeEvent, type XlsxFormControlInput, type XlsxFormControlKind, type XlsxFormControlKindInput, type XlsxFormControlPatch, type XlsxFormControlRenderProps, type XlsxFormControlSelectionMode, type XlsxFormControlState, type XlsxFormulaTarget, type XlsxImage, type XlsxImageAnchor, type XlsxImageRect, type XlsxImageRenderProps, type XlsxImageResizeHandlePosition, type XlsxImageSelectionRenderProps, type XlsxScrollerRenderProps, type XlsxShape, type XlsxShapeFill, type XlsxShapeParagraph, type XlsxShapeStroke, type XlsxShapeTextBox, type XlsxShapeTextRun, type XlsxSheetData, type XlsxSheetThumbnail, type XlsxSheetThumbnailResolution, type XlsxSheetVisibility, type XlsxTable, type XlsxTableColumn, type XlsxTableHeaderMenuRenderProps, type XlsxTableSortDirection, type XlsxTableSortState, type XlsxThemePalette, XlsxViewer, type XlsxViewerCharts, type XlsxViewerController, type XlsxViewerEditing, type XlsxViewerImages, type XlsxViewerProps, XlsxViewerProvider, type XlsxViewerProviderProps, type XlsxViewerSelection, type XlsxViewerTables, type XlsxViewerThumbnails, type XlsxViewerZoom, type XlsxWasmSource, type XlsxWorkbookTab, externalCallKey, initWasm, setWasmSource, useXlsxViewer, useXlsxViewerCharts, useXlsxViewerController, useXlsxViewerEditing, useXlsxViewerImages, useXlsxViewerSelection, useXlsxViewerTables, useXlsxViewerThumbnails, useXlsxViewerZoom };
