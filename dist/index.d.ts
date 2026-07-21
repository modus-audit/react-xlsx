import * as React from 'react';
import { Workbook, Worksheet } from '@dukelib/sheets-wasm';
import * as react_jsx_runtime from 'react/jsx-runtime';

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
type XlsxConditionalFormatRule = XlsxConditionalColorScaleRule | XlsxConditionalDataBarRule | XlsxConditionalIconSetRule;
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
interface XlsxCellRange {
    end: XlsxCellAddress;
    start: XlsxCellAddress;
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
interface XlsxFormControl {
    anchor: XlsxImageAnchor;
    checked?: boolean;
    fontFamily?: string;
    fontSizePt?: number;
    hidden?: boolean;
    id: string;
    kind: XlsxFormControlKind;
    label?: string;
    linkedCell?: string;
    name?: string;
    sheetIndex: number;
    textAlign?: "center" | "left" | "right";
    textColor?: string;
    workbookSheetIndex: number;
    zIndex: number;
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
        key: React.Key;
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
     * Pre-resolved values for external add-in calls (`[N]!FN(args)`, e.g. CCH `[1]!TBLink(...)`),
     * keyed by `externalCallKey(name, args)`. Passed (as a serializable map) into the engine so it
     * resolves those calls during `calculate()` — keeping the formula text untouched. Unmapped calls
     * keep the cell's cached value. Omit to disable (default behavior, identical to upstream).
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
    clearSelectedImage: () => void;
    getChartById: (id: string) => XlsxChart | null;
    getChartsheetById: (id: string) => XlsxChartsheet | null;
    formControls: XlsxFormControl[];
    getSheetCharts: (sheetIndex?: number) => XlsxChart[];
    getSheetFormControls: (sheetIndex?: number) => XlsxFormControl[];
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
    readOnly: boolean;
    recalculate: () => void;
    revision: number;
    resetZoom: () => void;
    resizeChartBy: (id: string, handle: XlsxImageResizeHandlePosition, deltaX: number, deltaY: number) => void;
    resizeImageBy: (id: string, handle: XlsxImageResizeHandlePosition, deltaX: number, deltaY: number) => void;
    resizeColumn: (col: number, widthPx: number) => void;
    resizeRow: (row: number, heightPx: number) => void;
    redo: () => void;
    pasteFromClipboard: () => Promise<boolean>;
    pasteStructuredClipboardData: (payload: string) => boolean;
    pasteText: (text: string) => boolean;
    selectedRangeAddress: string | null;
    selectedValue: string;
    selectedFormula: string;
    setCellFormula: (cell: XlsxCellAddress, formula: string) => void;
    setCellValue: (cell: XlsxCellAddress, value: string) => void;
    setZoomScale: (zoomScale: number) => void;
    selectCell: (cell: XlsxCellAddress, options?: {
        extend?: boolean;
        append?: boolean;
    }) => void;
    /** Select a cell AND scroll it into view + repaint the selection — the same path
        keyboard navigation uses. Prefer this over selectCell for programmatic "go to cell"
        (e.g. in-document search), which otherwise leaves the highlight unmoved. */
    revealCell: (cell: XlsxCellAddress) => void;
    /** @internal Grid wiring: XlsxGrid registers revealCell's scroll + overlay impl here. */
    registerRevealCellImpl: (impl: ((cell: XlsxCellAddress) => void) | null) => void;
    selectChart: (id: string | null) => void;
    selectRange: (range: XlsxCellRange, options?: {
        append?: boolean;
        toggle?: boolean;
    }) => void;
    selection: XlsxCellRange | null;
    /** All selected regions (non-contiguous, ENG multi-region). `selection` is the active/last
        region for back-compat; `selections` is the full set (length 1 for a normal selection). */
    selections: XlsxCellRange[];
    setActiveSheetIndex: (index: number) => void;
    setActiveTabIndex: (index: number) => void;
    selectedChart: XlsxChart | null;
    selectedChartId: string | null;
    selectedImage: XlsxImage | null;
    selectedImageId: string | null;
    setSelectedCellFormula: (formula: string) => void;
    setSelectedCellValue: (value: string) => void;
    sheets: XlsxSheetData[];
    src?: string;
    sortState: XlsxTableSortState | null;
    sortTable: (tableName: string, columnIndex: number, direction: XlsxTableSortDirection) => void;
    selectImage: (id: string | null) => void;
    setChartRect: (id: string, rect: XlsxImageRect) => void;
    setImageRect: (id: string, rect: XlsxImageRect) => void;
    getRowsBatchAsync?: (workbookSheetIndex: number, startRow: number, rowCount: number) => Promise<unknown[] | null>;
    tables: XlsxTable[];
    tabs: XlsxWorkbookTab[];
    undo: () => void;
    unmergeSelection: () => void;
    updateChart: (id: string, patch: Partial<XlsxChart>) => void;
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
    canRedo: boolean;
    canUndo: boolean;
    clearSelectedCells: () => void;
    copySelectionToClipboard: () => Promise<boolean>;
    defineNamedRange: (name: string, range?: XlsxCellRange | null) => void;
    fillSelection: (targetRange: XlsxCellRange) => void;
    getClipboardData: () => XlsxClipboardData | null;
    getCellDisplayValue: (cell?: XlsxCellAddress | null) => string;
    getCellFormula: (cell?: XlsxCellAddress | null) => string;
    mergeSelection: () => void;
    pasteFromClipboard: () => Promise<boolean>;
    pasteStructuredClipboardData: (payload: string) => boolean;
    pasteText: (text: string) => boolean;
    removeActiveSheet: () => void;
    readOnly: boolean;
    redo: () => void;
    selectedFormula: string;
    selectedValue: string;
    setCellFormula: (cell: XlsxCellAddress, formula: string) => void;
    setCellValue: (cell: XlsxCellAddress, value: string) => void;
    setSelectedCellFormula: (formula: string) => void;
    setSelectedCellValue: (value: string) => void;
    undo: () => void;
    unmergeSelection: () => void;
}
interface XlsxViewerTables {
    sortState: XlsxTableSortState | null;
    sortTable: (tableName: string, columnIndex: number, direction: XlsxTableSortDirection) => void;
    tables: XlsxTable[];
}
interface XlsxViewerImages {
    charts: XlsxChart[];
    clearSelectedChart: () => void;
    clearSelectedImage: () => void;
    getChartById: (id: string) => XlsxChart | null;
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
    selectedChartId: string | null;
    selectedImage: XlsxImage | null;
    selectedImageId: string | null;
    selectChart: (id: string | null) => void;
    selectImage: (id: string | null) => void;
    setChartRect: (id: string, rect: XlsxImageRect) => void;
    setImageRect: (id: string, rect: XlsxImageRect) => void;
    updateChart: (id: string, patch: Partial<XlsxChart>) => void;
}
interface XlsxViewerCharts {
    activeTab: XlsxWorkbookTab | null;
    activeTabIndex: number;
    charts: XlsxChart[];
    chartsheets: XlsxChartsheet[];
    clearSelectedChart: () => void;
    getChartById: (id: string) => XlsxChart | null;
    getChartsheetById: (id: string) => XlsxChartsheet | null;
    getSheetCharts: (sheetIndex?: number) => XlsxChart[];
    isChartsLoading: boolean;
    moveChartBy: (id: string, deltaX: number, deltaY: number) => void;
    readOnly: boolean;
    resizeChartBy: (id: string, handle: XlsxImageResizeHandlePosition, deltaX: number, deltaY: number) => void;
    selectChart: (id: string | null) => void;
    selectedChart: XlsxChart | null;
    selectedChartId: string | null;
    setActiveTabIndex: (index: number) => void;
    setChartRect: (id: string, rect: XlsxImageRect) => void;
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
    /** Replaces the chart loading placeholder. */
    renderChartLoading?: (props: XlsxChartLoadingRenderProps) => React.ReactNode;
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

/** External add-in function ([N]!FN(args), e.g. CCH `[1]!TBLink(...)`) host resolution.
 *
 * The forked duke-sheets engine parses `[N]!FN(args)` into a callable node and, during
 * `calculate()`, invokes an `externalFnFn(name, args) -> value | null` host callback for each one
 * (returning null leaves the cell's cached value untouched). A callback is a closure and cannot be
 * `postMessage`d into the worker where the engine runs, so the host passes a *serializable* map of
 * pre-resolved values instead; the worker rebuilds the callback locally from it.
 *
 * {@link externalCallKey} is the single source of truth for how that map is keyed — the host builds
 * the map with it, and the worker looks up with it, so the two never drift. */
/** Canonical key for one external call. `args` are the engine's stringified argument values
 *  (CCH args are strings/numbers), matching how the host parsed them from the formula text. */
declare function externalCallKey(name: string, args: readonly string[]): string;
/** Serializable map handed to the controller: `externalCallKey(name, args)` -> resolved value. */
type ExternalFnValues = Record<string, string | number>;

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

export { DefaultXlsxToolbar, type ExternalFnValues, type UseXlsxViewerControllerOptions, type UseXlsxViewerThumbnailsOptions, type XlsxCellAddress, type XlsxCellRange, type XlsxChart, type XlsxChartAxis, type XlsxChartDataLabels, type XlsxChartLoadingRenderProps, type XlsxChartReference, type XlsxChartSeries, type XlsxChartsheet, XlsxFileSizeLimitExceededError, type XlsxFileTooLargeRenderProps, type XlsxImage, type XlsxImageAnchor, type XlsxImageRect, type XlsxImageRenderProps, type XlsxImageResizeHandlePosition, type XlsxImageSelectionRenderProps, type XlsxScrollerRenderProps, type XlsxShape, type XlsxShapeFill, type XlsxShapeParagraph, type XlsxShapeStroke, type XlsxShapeTextBox, type XlsxShapeTextRun, type XlsxSheetData, type XlsxSheetThumbnail, type XlsxSheetThumbnailResolution, type XlsxSheetVisibility, type XlsxTable, type XlsxTableColumn, type XlsxTableHeaderMenuRenderProps, type XlsxTableSortDirection, type XlsxTableSortState, type XlsxThemePalette, XlsxViewer, type XlsxViewerCharts, type XlsxViewerController, type XlsxViewerEditing, type XlsxViewerImages, type XlsxViewerProps, XlsxViewerProvider, type XlsxViewerProviderProps, type XlsxViewerSelection, type XlsxViewerTables, type XlsxViewerThumbnails, type XlsxViewerZoom, type XlsxWorkbookTab, externalCallKey, useXlsxViewer, useXlsxViewerCharts, useXlsxViewerController, useXlsxViewerEditing, useXlsxViewerImages, useXlsxViewerSelection, useXlsxViewerTables, useXlsxViewerThumbnails, useXlsxViewerZoom };
