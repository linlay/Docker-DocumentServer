export type AiBridgeEditorType = "word" | "slide" | "cell";
export type AiBridgeControl = "save" | "history" | "undo" | "redo";
export type AiBridgeErrorCode =
  | "AI_BRIDGE_ERROR"
  | "NOT_READY"
  | "CONNECTION_TIMEOUT"
  | "TIMEOUT"
  | "INVALID_TARGET_WINDOW"
  | "INVALID_TARGET_ORIGIN"
  | "INVALID_CLIENT_ID"
  | "INVALID_REQUEST_ID"
  | "REQUEST_IN_FLIGHT"
  | "INVALID_COMMAND"
  | "INVALID_TOOL_CALL"
  | "INVALID_ARGUMENTS"
  | "ARGUMENTS_TOO_LARGE"
  | "TOO_MANY_CALLS"
  | "TOOL_NOT_ALLOWED"
  | "CONTROL_NOT_ALLOWED"
  | "METHOD_NOT_ALLOWED"
  | "DOCUMENT_NOT_CONFIGURED"
  | "DOCUMENT_MISMATCH"
  | "EDITOR_MISMATCH"
  | "PERSISTENCE_FAILED"
  | "EXECUTION_FAILED"
  | "INVALID_LISTENER"
  | "CLIENT_DESTROYED"
  | string;

export interface AiBridgeRequestOptions {
  /** 100-300000 ms; defaults to 90000 ms. */
  timeoutMs?: number;
  /** Stable idempotency key: 1-200 characters from [A-Za-z0-9._:-]. */
  requestId?: string;
}

export interface AiBridgeContext {
  documentKey: string;
  fileName: string;
  fileType: string;
  editorType: AiBridgeEditorType | "";
  userId: string;
  /** Available only to the direct host API, not the cross-origin client state. */
  callbackUrl?: string;
}

export interface AiBridgeCapabilities {
  editorType: AiBridgeEditorType;
  tools: AiBridgeToolName[];
  controls: AiBridgeControl[];
}

export interface AiBridgeState {
  version: "0.2.0";
  protocolVersion: 1;
  pluginGuid: "asc.{A17E5F31-64AA-4E37-9A42-8D430814C2F6}";
  ready: boolean;
  editorType: AiBridgeEditorType | null;
  context: AiBridgeContext;
  capabilities: AiBridgeCapabilities | null;
}

export interface AiBridgeExecutionResult {
  ok?: boolean;
  editorType?: AiBridgeEditorType;
  changed: number;
  needsSave?: boolean;
  results: Array<Record<string, unknown>>;
  persisted: boolean;
  forceSave?: Record<string, unknown>;
}

export interface AiBridgeServiceResult extends Record<string, unknown> {
  persisted?: boolean;
  key?: string;
  fileName?: string;
}

export class AiBridgeError extends Error {
  readonly code: AiBridgeErrorCode;
  readonly requestId: string | null;
  readonly details?: unknown;
}

export interface BasicTextFormat {
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;
}
export interface TextFormat extends BasicTextFormat {
  underline?: boolean;
  strikeout?: boolean;
  doubleStrikeout?: boolean;
  caps?: boolean;
  smallCaps?: boolean;
  highlightColor?: string;
  characterSpacing?: number;
  vertAlign?: "baseline" | "subscript" | "superscript";
}
export interface WordParagraphFormat extends TextFormat {
  align?: "left" | "center" | "right" | "both";
  styleName?: string;
  headingLevel?: number;
  outlineLevel?: number;
  spacingBefore?: number;
  spacingAfter?: number;
  lineSpacing?: number;
  lineRule?: "auto" | "exact" | "atLeast";
  firstLineIndent?: number;
  leftIndent?: number;
  rightIndent?: number;
  keepLines?: boolean;
  keepNext?: boolean;
  widowControl?: boolean;
  contextualSpacing?: boolean;
  pageBreakBefore?: boolean;
}
export interface WordParagraphTarget {
  /** One-based indexes from word_inspect.paragraphDetails. */
  paragraphIndexes?: number[];
  search?: string;
  matchCase?: boolean;
  matchMode?: "contains" | "exact";
  occurrence?: number;
  maxParagraphs?: number;
  all?: boolean;
  current?: boolean;
}
export interface WordSearchTarget {
  search: string;
  matchCase?: boolean;
  occurrence?: number;
  maxMatches?: number;
}

export interface WordInspectArgs {
  maxChars?: number;
  maxParagraphs?: number;
  includeStructure?: boolean;
  includeComments?: boolean;
}
export interface WordReplaceTextArgs { search: string; replace: string; matchCase?: boolean; }
export interface WordAppendParagraphArgs extends WordParagraphFormat {
  text: string;
  listType?: "bullet" | "numbered";
  listLevel?: number;
}
export interface WordInsertParagraphArgs extends WordParagraphFormat {
  text: string;
  inline?: boolean;
  listType?: "bullet" | "numbered";
  listLevel?: number;
}
export interface WordFormatDocumentArgs extends WordParagraphFormat {}
export interface WordFormatSelectionArgs extends TextFormat {}
export interface WordFormatMatchesArgs extends WordSearchTarget, TextFormat {}
export interface WordDeleteMatchesArgs extends WordSearchTarget {}
export interface WordAddHyperlinkArgs extends WordSearchTarget {
  url: string;
  screenTip?: string;
  bookmarkName?: string;
}
export interface WordAddCommentArgs extends WordSearchTarget {
  text: string;
  author?: string;
  userId?: string;
}
export interface WordAddBookmarkArgs extends WordSearchTarget {
  occurrence: number;
  name: string;
}
export interface WordFormatParagraphsArgs extends WordParagraphTarget, WordParagraphFormat {}
export interface WordSetParagraphTextArgs extends WordParagraphTarget, WordParagraphFormat { text: string; }
export interface WordDeleteParagraphsArgs extends WordParagraphTarget {}
export interface WordSetListArgs extends WordParagraphTarget {
  listType: "bullet" | "numbered";
  level?: number;
  contextualSpacing?: boolean;
}
export interface WordInsertPageBreakArgs extends WordParagraphTarget {
  position?: "before" | "after";
}
export interface WordNavigateArgs {
  target?: "start" | "end" | "page" | "next" | "previous" | "relative" | "current" | "search";
  /** One-based page number; required when target is "page". */
  page?: number;
  pageDelta?: number;
  search?: string;
  matchCase?: boolean;
  occurrence?: number;
}
export interface WordScrollArgs { direction?: "up" | "down"; pages?: number; }
export interface WordScaleFontArgs { scale: number; }
export interface WordAddTableArgs extends BasicTextFormat {
  rows: number;
  cols: number;
  data?: unknown[][];
  widthPercent?: number;
  styleName?: string;
  firstRow?: boolean;
  lastRow?: boolean;
  firstColumn?: boolean;
  lastColumn?: boolean;
  horizontalBanding?: boolean;
  verticalBanding?: boolean;
  title?: string;
  description?: string;
}
export interface WordSetTableCellArgs extends WordParagraphFormat {
  tableIndex: number;
  row: number;
  column: number;
  text?: string;
  backgroundColor?: string;
  verticalAlign?: "top" | "center" | "bottom";
  widthPercent?: number;
}
export interface WordFormatTableArgs extends BasicTextFormat {
  tableIndex: number;
  widthPercent?: number;
  align?: "left" | "center" | "right";
  styleName?: string;
  backgroundColor?: string;
  title?: string;
  description?: string;
  firstRow?: boolean;
  lastRow?: boolean;
  firstColumn?: boolean;
  lastColumn?: boolean;
  horizontalBanding?: boolean;
  verticalBanding?: boolean;
  verticalAlign?: "top" | "center" | "bottom";
}
export interface WordEditTableArgs {
  tableIndex: number;
  action: "addRow" | "addColumn" | "removeRow" | "removeColumn" | "mergeCells" | "splitCell" | "clear" | "delete";
  row?: number;
  column?: number;
  position?: "before" | "after";
  rowStart?: number;
  rowEnd?: number;
  columnStart?: number;
  columnEnd?: number;
  rows?: number;
  columns?: number;
}
export interface WordSetPageLayoutArgs {
  sectionIndex?: number;
  pageSize?: "A4" | "Letter" | "Legal" | "custom";
  orientation?: "portrait" | "landscape";
  widthMm?: number;
  heightMm?: number;
  marginLeftMm?: number;
  marginTopMm?: number;
  marginRightMm?: number;
  marginBottomMm?: number;
  headerDistanceMm?: number;
  footerDistanceMm?: number;
  titlePage?: boolean;
}
export interface WordSetHeaderFooterArgs extends BasicTextFormat {
  kind: "header" | "footer";
  type?: "default" | "first" | "even";
  action?: "set" | "remove";
  sectionIndex?: number;
  text?: string;
  replace?: boolean;
  pageNumber?: boolean;
  pagesCount?: boolean;
  align?: "left" | "center" | "right" | "both";
}
export interface WordSetDocumentTextArgs { text: string; }

export interface AiBridgeGradientStop {
  /** Position along the gradient, from 0 to 100 percent. */
  position: number;
  color: string;
}
export type AiBridgeFill =
  | { type: "none" }
  | { type: "solid"; color: string }
  | { type: "linearGradient"; stops: AiBridgeGradientStop[]; angleDeg?: number }
  | { type: "radialGradient"; stops: AiBridgeGradientStop[] }
  | { type: "pattern"; pattern: string; backgroundColor: string; foregroundColor: string }
  | { type?: "raw"; raw: unknown };
export interface AiBridgeLine {
  type?: "solid" | "none";
  enabled?: boolean;
  widthPt?: number;
  color?: string;
  fill?: AiBridgeFill;
  raw?: unknown;
}
export interface AiBridgePadding {
  left?: number;
  top?: number;
  right?: number;
  bottom?: number;
}
export interface AiBridgeChartDataLabels {
  showSeriesName?: boolean;
  showCategoryName?: boolean;
  showValue?: boolean;
  showPercent?: boolean;
}
export interface AiBridgeChartAxis {
  title?: string;
  titleFontSize?: number;
  labelsFontSize?: number;
  normalOrder?: boolean;
  majorTickMark?: string;
  minorTickMark?: string;
  tickLabelPosition?: "none" | "nextTo" | "low" | "high";
  numberFormat?: string;
  /** Axis position used by ONLYOFFICE SetAxieNumFormat, e.g. left or bottom. */
  position?: string;
}
export interface AiBridgeChartLegend {
  position?: "left" | "top" | "right" | "bottom" | "none";
  fontSize?: number;
  fill?: AiBridgeFill;
  line?: AiBridgeLine;
}
export interface AiBridgeChartPointUpdate {
  index: number;
  fill?: AiBridgeFill;
  line?: AiBridgeLine;
  markerFill?: AiBridgeFill;
  markerLine?: AiBridgeLine;
  numberFormat?: string;
  allSeries?: boolean;
  allMarkers?: boolean;
  dataLabels?: AiBridgeChartDataLabels;
}
export interface AiBridgeChartSeriesUpdate {
  /** Zero-based series index. */
  index: number;
  type?: AiBridgeChartType;
  name?: string;
  /** PPT chart values. */
  values?: number[];
  /** XLSX source range for series values. */
  valuesRange?: string;
  /** PPT scatter-chart X values. */
  xValues?: number[];
  /** XLSX source range for scatter-chart X values. */
  xValuesRange?: string;
  numberFormat?: string;
  fill?: AiBridgeFill;
  line?: AiBridgeLine;
  allSeries?: boolean;
  points?: AiBridgeChartPointUpdate[];
}
export interface AiBridgeChartGridlines {
  majorHorizontal?: { line: AiBridgeLine };
  minorHorizontal?: { line: AiBridgeLine };
  majorVertical?: { line: AiBridgeLine };
  minorVertical?: { line: AiBridgeLine };
}
export type AiBridgeChartType =
  | "bar" | "barStacked" | "barStackedPercent" | "bar3D"
  | "barStacked3D" | "barStackedPercent3D" | "barStackedPercent3DPerspective"
  | "horizontalBar" | "horizontalBarStacked" | "horizontalBarStackedPercent"
  | "horizontalBar3D" | "horizontalBarStacked3D" | "horizontalBarStackedPercent3D"
  | "lineNormal" | "lineStacked" | "lineStackedPercent" | "lineNormalMarker"
  | "lineStackedMarker" | "lineStackedPerMarker" | "line3D"
  | "pie" | "pie3D" | "doughnut"
  | "scatter" | "scatterLine" | "scatterLineMarker" | "scatterSmooth" | "scatterSmoothMarker"
  | "stock" | "area" | "areaStacked" | "areaStackedPercent"
  | "comboCustom" | "comboBarLine" | "comboBarLineSecondary"
  | "radar" | "radarMarker" | "radarFilled"
  /** Friendly aliases normalized by the bridge. */
  | "line" | "lineMarker" | "stackedBar" | "stackedBarPercent"
  | "stackedLine" | "stackedLinePercent" | "column";
export interface AiBridgeChartFormatting {
  name?: string;
  style?: number;
  title?: string;
  titleFontSize?: number;
  rotationDeg?: number;
  fill?: AiBridgeFill;
  line?: AiBridgeLine;
  plotAreaFill?: AiBridgeFill;
  plotAreaLine?: AiBridgeLine;
  titleFill?: AiBridgeFill;
  titleLine?: AiBridgeLine;
  legend?: AiBridgeChartLegend;
  horizontalAxis?: AiBridgeChartAxis;
  verticalAxis?: AiBridgeChartAxis;
  dataLabels?: AiBridgeChartDataLabels;
  gridlines?: AiBridgeChartGridlines;
  seriesUpdates?: AiBridgeChartSeriesUpdate[];
  removeSeries?: number[];
  includeRaw?: boolean;
}
export interface SlidesObjectTarget {
  slide: number;
  objectId?: string;
  /** Zero-based index returned by slides_inspect_objects. */
  objectIndex?: number;
  name?: string;
}
export interface SlidesChartTarget {
  slide: number;
  chartId?: string;
  /** Zero-based index returned by slides_inspect_charts. */
  chartIndex?: number;
  name?: string;
}
export interface SlidesShapeFormatting extends BasicTextFormat {
  shapeType?: string;
  text?: string;
  align?: string;
  verticalAlign?: "top" | "center" | "bottom";
  xMm?: number;
  yMm?: number;
  widthMm?: number;
  heightMm?: number;
  rotationDeg?: number;
  flipH?: boolean;
  flipV?: boolean;
  name?: string;
  fill?: AiBridgeFill;
  line?: AiBridgeLine;
  paddingMm?: AiBridgePadding;
  /** Legacy solid fill. Prefer fill. */
  fillColor?: string;
  lineColor?: string;
  lineWidthPt?: number;
  includeRaw?: boolean;
}
export interface SlidesInspectArgs { maxChars?: number; }
export interface SlidesInspectObjectsArgs {
  slide?: number;
  kinds?: Array<"shape" | "chart" | "image" | "table" | "oleObject" | "group" | string>;
  maxObjects?: number;
  includeRaw?: boolean;
  includeSlideRaw?: boolean;
}
export interface SlidesReplaceTextArgs { search: string; replace: string; matchCase?: boolean; slide?: number; }
export interface SlidesScaleFontArgs { scale: number; slide?: number; }
export interface SlidesFormatTextArgs extends TextFormat { slide?: number; }
export interface SlidesFormatSelectionArgs extends TextFormat {}
export interface SlidesAddSlideArgs { index?: number; title?: string; titleFontSize?: number; backgroundColor?: string; }
export interface SlidesDuplicateSlideArgs { slide: number; }
export interface SlidesDeleteSlideArgs { slide: number; }
export interface SlidesAddTextBoxArgs extends SlidesShapeFormatting {
  slide: number;
  text: string;
}
export interface SlidesSetBackgroundArgs {
  slide: number;
  mode?: "custom" | "clear" | "layout" | "master";
  fill?: AiBridgeFill;
}
export interface SlidesAddShapeArgs extends SlidesShapeFormatting {
  slide: number;
  shapeType: string;
}
export interface SlidesUpdateShapeArgs extends SlidesObjectTarget, SlidesShapeFormatting {}
export interface SlidesDeleteObjectArgs extends SlidesObjectTarget {}
export interface SlidesInspectChartsArgs {
  slide?: number;
  maxCharts?: number;
  includeRaw?: boolean;
}
export interface SlidesAddChartArgs extends AiBridgeChartFormatting {
  slide: number;
  type?: AiBridgeChartType;
  series: number[][];
  seriesNames: Array<string | number>;
  categories: Array<string | number>;
  numFormats?: string[];
  xMm?: number;
  yMm?: number;
  widthMm?: number;
  heightMm?: number;
}
export interface SlidesUpdateChartArgs extends SlidesChartTarget, AiBridgeChartFormatting {
  categories?: Array<string | number>;
  xMm?: number;
  yMm?: number;
  widthMm?: number;
  heightMm?: number;
}
export interface SlidesDeleteChartArgs extends SlidesChartTarget {}

export interface SheetsInspectArgs { maxCells?: number; }
export interface SheetTarget { sheet?: string; }
export interface SheetsSetValuesArgs extends SheetTarget { range: string; values: unknown; }
export interface SheetsSetFormulaArgs extends SheetTarget { range: string; formula: string; }
export interface SheetsReplaceTextArgs extends SheetTarget { range?: string; search: string; replace: string; }
export interface SheetsFormatRangeArgs extends SheetTarget {
  range: string;
  fontSize?: number;
  fontName?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontColor?: string;
  fillColor?: string;
  horizontalAlign?: string;
  verticalAlign?: string;
  numberFormat?: string;
  wrap?: boolean;
  columnWidth?: number;
  rowHeight?: number;
}
export interface SheetsAddSheetArgs { name: string; }
export interface SheetsRenameSheetArgs extends SheetTarget { newName: string; }
export interface SheetsDeleteSheetArgs { sheet: string; }
export interface SheetsAddChartArgs extends SheetTarget {
  range: string;
  type?: AiBridgeChartType;
  title?: string;
  titleFontSize?: number;
  inRows?: boolean;
  style?: number;
  widthMm?: number;
  heightMm?: number;
  fromColumn?: number;
  fromRow?: number;
  columnOffsetMm?: number;
  rowOffsetMm?: number;
  categoryRange?: string;
  addSeries?: Array<{ name: string; valuesRange: string; xValuesRange?: string }>;
  rotationDeg?: number;
  name?: string;
  fill?: AiBridgeFill;
  line?: AiBridgeLine;
  plotAreaFill?: AiBridgeFill;
  plotAreaLine?: AiBridgeLine;
  titleFill?: AiBridgeFill;
  titleLine?: AiBridgeLine;
  legend?: AiBridgeChartLegend;
  horizontalAxis?: AiBridgeChartAxis;
  verticalAxis?: AiBridgeChartAxis;
  dataLabels?: AiBridgeChartDataLabels;
  gridlines?: AiBridgeChartGridlines;
  seriesUpdates?: AiBridgeChartSeriesUpdate[];
  removeSeries?: number[];
  includeRaw?: boolean;
}
export interface SheetsInspectChartsArgs extends SheetTarget { maxCharts?: number; includeRaw?: boolean; }
export interface SheetsChartTarget extends SheetTarget {
  /** Zero-based index returned by sheets_inspect_charts. */
  chartIndex?: number;
  name?: string;
}
export interface SheetsUpdateChartArgs extends SheetsChartTarget, AiBridgeChartFormatting {
  widthMm?: number;
  heightMm?: number;
  fromColumn?: number;
  fromRow?: number;
  columnOffsetMm?: number;
  rowOffsetMm?: number;
  categoryRange?: string;
  addSeries?: Array<{ name: string; valuesRange: string; xValuesRange?: string }>;
}
/** Requires ApiDrawing.Delete, a paid capability in some ONLYOFFICE Docs editions. */
export interface SheetsDeleteChartArgs extends SheetsChartTarget {}

export interface AiBridgeToolArgumentsMap {
  word_inspect: WordInspectArgs;
  word_replace_text: WordReplaceTextArgs;
  word_append_paragraph: WordAppendParagraphArgs;
  word_insert_paragraph: WordInsertParagraphArgs;
  word_format_document: WordFormatDocumentArgs;
  word_format_selection: WordFormatSelectionArgs;
  word_format_matches: WordFormatMatchesArgs;
  word_delete_matches: WordDeleteMatchesArgs;
  word_add_hyperlink: WordAddHyperlinkArgs;
  word_add_comment: WordAddCommentArgs;
  word_add_bookmark: WordAddBookmarkArgs;
  word_format_paragraphs: WordFormatParagraphsArgs;
  word_set_paragraph_text: WordSetParagraphTextArgs;
  word_delete_paragraphs: WordDeleteParagraphsArgs;
  word_set_list: WordSetListArgs;
  word_insert_page_break: WordInsertPageBreakArgs;
  word_navigate: WordNavigateArgs;
  word_scroll: WordScrollArgs;
  word_scale_font: WordScaleFontArgs;
  word_add_table: WordAddTableArgs;
  word_set_table_cell: WordSetTableCellArgs;
  word_format_table: WordFormatTableArgs;
  word_edit_table: WordEditTableArgs;
  word_set_page_layout: WordSetPageLayoutArgs;
  word_set_header_footer: WordSetHeaderFooterArgs;
  word_set_document_text: WordSetDocumentTextArgs;
  slides_inspect: SlidesInspectArgs;
  slides_inspect_objects: SlidesInspectObjectsArgs;
  slides_replace_text: SlidesReplaceTextArgs;
  slides_scale_font: SlidesScaleFontArgs;
  slides_format_text: SlidesFormatTextArgs;
  slides_format_selection: SlidesFormatSelectionArgs;
  slides_add_slide: SlidesAddSlideArgs;
  slides_duplicate_slide: SlidesDuplicateSlideArgs;
  slides_delete_slide: SlidesDeleteSlideArgs;
  slides_add_textbox: SlidesAddTextBoxArgs;
  slides_set_background: SlidesSetBackgroundArgs;
  slides_add_shape: SlidesAddShapeArgs;
  slides_update_shape: SlidesUpdateShapeArgs;
  slides_delete_object: SlidesDeleteObjectArgs;
  slides_inspect_charts: SlidesInspectChartsArgs;
  slides_add_chart: SlidesAddChartArgs;
  slides_update_chart: SlidesUpdateChartArgs;
  slides_delete_chart: SlidesDeleteChartArgs;
  sheets_inspect: SheetsInspectArgs;
  sheets_set_values: SheetsSetValuesArgs;
  sheets_set_formula: SheetsSetFormulaArgs;
  sheets_replace_text: SheetsReplaceTextArgs;
  sheets_format_range: SheetsFormatRangeArgs;
  sheets_add_sheet: SheetsAddSheetArgs;
  sheets_rename_sheet: SheetsRenameSheetArgs;
  sheets_delete_sheet: SheetsDeleteSheetArgs;
  sheets_add_chart: SheetsAddChartArgs;
  sheets_inspect_charts: SheetsInspectChartsArgs;
  sheets_update_chart: SheetsUpdateChartArgs;
  sheets_delete_chart: SheetsDeleteChartArgs;
}

export type AiBridgeToolName = keyof AiBridgeToolArgumentsMap;
export type AiBridgeToolCall = {
  [Name in AiBridgeToolName]: {
    id?: string;
    name: Name;
    arguments: AiBridgeToolArgumentsMap[Name];
  }
}[AiBridgeToolName];

export interface AiBridgeCommand { toolCalls: AiBridgeToolCall[]; }

export interface AiBridgeWordApi {
  inspect(args?: WordInspectArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  replaceText(args: WordReplaceTextArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  appendParagraph(args: WordAppendParagraphArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  insertParagraph(args: WordInsertParagraphArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  formatDocument(args: WordFormatDocumentArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  formatSelection(args: WordFormatSelectionArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  formatMatches(args: WordFormatMatchesArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  deleteMatches(args: WordDeleteMatchesArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  addHyperlink(args: WordAddHyperlinkArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  addComment(args: WordAddCommentArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  addBookmark(args: WordAddBookmarkArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  formatParagraphs(args: WordFormatParagraphsArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  setParagraphText(args: WordSetParagraphTextArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  deleteParagraphs(args: WordDeleteParagraphsArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  setList(args: WordSetListArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  insertPageBreak(args: WordInsertPageBreakArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  navigate(args?: WordNavigateArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  scroll(args?: WordScrollArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  scaleFont(args: WordScaleFontArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  addTable(args: WordAddTableArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  setTableCell(args: WordSetTableCellArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  formatTable(args: WordFormatTableArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  editTable(args: WordEditTableArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  setPageLayout(args: WordSetPageLayoutArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  setHeaderFooter(args: WordSetHeaderFooterArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  setDocumentText(args: WordSetDocumentTextArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
}

export interface AiBridgeSlidesApi {
  inspect(args?: SlidesInspectArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  inspectObjects(args?: SlidesInspectObjectsArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  replaceText(args: SlidesReplaceTextArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  scaleFont(args: SlidesScaleFontArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  formatText(args: SlidesFormatTextArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  formatSelection(args: SlidesFormatSelectionArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  addSlide(args: SlidesAddSlideArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  duplicateSlide(args: SlidesDuplicateSlideArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  deleteSlide(args: SlidesDeleteSlideArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  addTextBox(args: SlidesAddTextBoxArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  setBackground(args: SlidesSetBackgroundArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  addShape(args: SlidesAddShapeArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  updateShape(args: SlidesUpdateShapeArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  deleteObject(args: SlidesDeleteObjectArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  inspectCharts(args?: SlidesInspectChartsArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  addChart(args: SlidesAddChartArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  updateChart(args: SlidesUpdateChartArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  deleteChart(args: SlidesDeleteChartArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
}

export interface AiBridgeSheetsApi {
  inspect(args?: SheetsInspectArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  setValues(args: SheetsSetValuesArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  setFormula(args: SheetsSetFormulaArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  replaceText(args: SheetsReplaceTextArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  formatRange(args: SheetsFormatRangeArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  addSheet(args: SheetsAddSheetArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  renameSheet(args: SheetsRenameSheetArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  deleteSheet(args: SheetsDeleteSheetArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  addChart(args: SheetsAddChartArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  inspectCharts(args?: SheetsInspectChartsArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  updateChart(args: SheetsUpdateChartArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  deleteChart(args: SheetsDeleteChartArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
}

export type AiBridgeEventName = "ready" | "reload" | "error";

export interface AiBridgeApi {
  readonly version: "0.2.0";
  readonly protocolVersion: 1;
  readonly pluginGuid: "asc.{A17E5F31-64AA-4E37-9A42-8D430814C2F6}";
  readonly isReady: boolean;
  readonly editorType: AiBridgeEditorType | null;
  readonly context: AiBridgeContext;
  readonly capabilities: AiBridgeCapabilities | null;
  readonly AiBridgeError: typeof AiBridgeError;
  readonly word: AiBridgeWordApi;
  readonly slides: AiBridgeSlidesApi;
  readonly sheets: AiBridgeSheetsApi;
  ready(options?: Pick<AiBridgeRequestOptions, "timeoutMs">): Promise<void>;
  getState(): AiBridgeState;
  execute(command: AiBridgeCommand | AiBridgeToolCall, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  executeTool<Name extends AiBridgeToolName>(name: Name, args: AiBridgeToolArgumentsMap[Name], options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  executeBatch(toolCalls: AiBridgeToolCall[], options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  save(options?: AiBridgeRequestOptions): Promise<AiBridgeServiceResult>;
  history(options?: AiBridgeRequestOptions): Promise<AiBridgeServiceResult>;
  undo(options?: AiBridgeRequestOptions): Promise<AiBridgeServiceResult>;
  redo(options?: AiBridgeRequestOptions): Promise<AiBridgeServiceResult>;
  on(name: AiBridgeEventName, listener: (detail: unknown) => void): () => void;
  off(name: AiBridgeEventName, listener: (detail: unknown) => void): void;
}

export interface AiBridgeClientOptions {
  targetWindow: Window;
  /** Exact origin only, e.g. https://editor.example.com. Wildcard is forbidden. */
  targetOrigin: string;
  timeoutMs?: number;
  clientId?: string;
}

export interface AiBridgeClientInstance extends Omit<AiBridgeApi, "ready" | "getState" | "AiBridgeError" | "context" | "capabilities" | "editorType" | "pluginGuid"> {
  readonly state: AiBridgeState | null;
  readonly context: AiBridgeContext | null;
  readonly capabilities: AiBridgeCapabilities | null;
  readonly editorType: AiBridgeEditorType | null;
  readonly pluginGuid: AiBridgeState["pluginGuid"] | null;
  ready(options?: Pick<AiBridgeRequestOptions, "timeoutMs">): Promise<AiBridgeState>;
  connect(options?: Pick<AiBridgeRequestOptions, "timeoutMs">): Promise<AiBridgeState>;
  refreshState(options?: AiBridgeRequestOptions): Promise<AiBridgeState>;
  destroy(): void;
}

export interface AiBridgeClientConstructor {
  new(options: AiBridgeClientOptions): AiBridgeClientInstance;
}

export interface AiBridgeHostOptions {
  getEditorConfig?: () => Record<string, unknown>;
  /** Exact origins allowed to use client-sdk.js. There is no default allow-list. */
  clientOrigins?: string[];
}

declare global {
  interface Window {
    aiBridge: AiBridgeApi;
    onlyofficeAI: AiBridgeApi;
    AiBridgeClient: AiBridgeClientConstructor;
    AiBridgeError: typeof AiBridgeError;
    aiBridgeOptions?: AiBridgeHostOptions;
  }
}
