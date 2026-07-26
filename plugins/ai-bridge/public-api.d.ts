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
  | "INVALID_IMAGE_SOURCE"
  | "IMAGE_FETCH_BLOCKED"
  | "IMAGE_FETCH_FAILED"
  | "IMAGE_TOO_LARGE"
  | "UNSUPPORTED_IMAGE_FORMAT"
  | "IMAGE_ASSET_EXPIRED"
  | "IMAGE_API_UNSUPPORTED"
  | "WORD_API_UNSUPPORTED"
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
  version: "0.4.0";
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
export type AiBridgeImageSource =
  | { type: "url"; url: string }
  | { type: "dataUrl"; dataUrl: string };
export interface TextFormat extends BasicTextFormat {
  /** Name of a character style. ONLYOFFICE stores this as a run style. */
  characterStyleName?: string;
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
export interface WordAddImageArgs {
  source: AiBridgeImageSource;
  widthMm?: number;
  heightMm?: number;
  preserveAspectRatio?: boolean;
  current?: boolean;
  /** One-based index from word_inspect.paragraphDetails. */
  paragraphIndex?: number;
  search?: string;
  occurrence?: number;
  matchCase?: boolean;
  wrapping?: "inline" | "square" | "tight" | "through" | "topAndBottom" | "behind" | "inFront";
  name?: string;
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
  /** Defaults to "end". before/after require exactly one anchor below. */
  insertAt?: "end" | "current" | "before" | "after";
  /** One-based top-level paragraph anchor. */
  paragraphIndex?: number;
  /** One-based top-level table anchor. */
  tableIndex?: number;
  /** Paragraph text anchor. */
  search?: string;
  matchCase?: boolean;
  matchMode?: "contains" | "exact";
  /** One-based search occurrence; defaults to 1. */
  occurrence?: number;
  /** Insert a page-break paragraph immediately before the table. */
  pageBreakBefore?: boolean;
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
  /** Field instruction codes such as `TIME \@ "yyyy-MM-dd"` or `FILENAME`. */
  fields?: string[];
  align?: "left" | "center" | "right" | "both";
}
export interface WordSetDocumentTextArgs { text: string; }
export interface WordInspectAdvancedArgs {
  includeProperties?: boolean;
  /** Custom property names to read; ONLYOFFICE does not expose name enumeration. */
  customPropertyNames?: string[];
  includeSections?: boolean;
  includeStyles?: boolean;
  includeNumbering?: boolean;
  includeDrawings?: boolean;
  includeBookmarks?: boolean;
  includeNotes?: boolean;
  includeComments?: boolean;
  includeRevisions?: boolean;
  includeContentControls?: boolean;
  includeCustomXml?: boolean;
  maxItems?: number;
}
export interface WordSetDocumentPropertiesArgs {
  title?: string;
  subject?: string;
  creator?: string;
  description?: string;
  keywords?: string;
  category?: string;
  language?: string;
  identifier?: string;
  lastModifiedBy?: string;
  revision?: string;
  created?: string;
  modified?: string;
  custom?: Array<{
    name: string;
    value: string | number | boolean;
    valueType?: "string" | "number" | "boolean" | "date";
  }>;
}
export interface WordSectionColumn {
  widthMm: number;
  spaceMm?: number;
}
export interface WordManageSectionArgs {
  action?: "configure" | "create";
  /** One-based section index for configure. */
  sectionIndex?: number;
  /** One-based paragraph index that ends the new section. */
  paragraphIndex?: number;
  type?: "continuous" | "nextPage" | "evenPage" | "oddPage";
  startPageNumber?: number;
  titlePage?: boolean;
  evenAndOddHeaders?: boolean;
  columns?: {
    count?: number;
    spaceMm?: number;
    entries?: WordSectionColumn[];
  };
}
export interface WordManageStyleArgs extends WordParagraphFormat {
  action?: "create" | "update";
  name: string;
  type?: "paragraph" | "character" | "table" | "numbering";
  basedOn?: string;
}
export interface WordTabStop {
  positionMm: number;
  align?: "left" | "center" | "right" | "decimal" | "bar" | "clear";
}
export interface WordSetTabsArgs extends WordParagraphTarget {
  tabs: WordTabStop[];
  clearAll?: boolean;
}
export interface WordNumberingLevel {
  level: number;
  format?: string;
  text?: string;
  start?: number;
  align?: "left" | "center" | "right";
  restart?: number;
}
export interface WordSetNumberingArgs extends WordParagraphTarget {
  kind?: "bullet" | "numbered" | "multilevel";
  level?: number;
  levels?: WordNumberingLevel[];
  restartAt?: number;
}
export interface WordBorderFormat {
  style?: "none" | "single" | "double" | "dotted" | "dashed" | "thick" | "wave";
  color?: string;
  widthPt?: number;
  spacePt?: number;
}
export interface WordFormatTableAdvancedArgs {
  tableIndex: number;
  row?: number;
  column?: number;
  rowHeightMm?: number;
  rowHeightRule?: "auto" | "atLeast";
  columnWidthMm?: number;
  repeatHeader?: boolean;
  cellMarginsMm?: { top?: number; right?: number; bottom?: number; left?: number };
  borders?: {
    top?: WordBorderFormat;
    right?: WordBorderFormat;
    bottom?: WordBorderFormat;
    left?: WordBorderFormat;
    insideHorizontal?: WordBorderFormat;
    insideVertical?: WordBorderFormat;
  };
  wrapping?: "none" | "around";
}
export interface WordAddNestedTableArgs extends BasicTextFormat {
  tableIndex: number;
  row: number;
  column: number;
  rows: number;
  cols: number;
  data?: unknown[][];
  widthPercent?: number;
  styleName?: string;
}
export interface WordDrawingTarget {
  drawingIndex?: number;
  name?: string;
  kind?: "image" | "shape" | "chart" | "oleObject" | "smartArt" | "any";
}
export interface WordManageDrawingArgs extends WordDrawingTarget {
  action: "update" | "delete";
  widthMm?: number;
  heightMm?: number;
  rotationDeg?: number;
  wrapping?: "inline" | "square" | "tight" | "through" | "topAndBottom" | "behind" | "inFront";
  xMm?: number;
  yMm?: number;
  relativeFromH?: string;
  relativeFromV?: string;
  alignH?: string;
  alignV?: string;
  nameUpdate?: string;
  border?: WordBorderFormat;
}
export interface WordAddShapeArgs extends BasicTextFormat {
  shapeType: string;
  text?: string;
  widthMm?: number;
  heightMm?: number;
  fillColor?: string;
  lineColor?: string;
  lineWidthPt?: number;
  wrapping?: "inline" | "square" | "tight" | "through" | "topAndBottom" | "behind" | "inFront";
  rotationDeg?: number;
  name?: string;
  paragraphIndex?: number;
  current?: boolean;
}
export interface WordAddChartArgs {
  chartType?: AiBridgeChartType;
  data: number[][];
  seriesNames?: string[];
  categories?: Array<string | number>;
  numberFormats?: string[];
  widthMm?: number;
  heightMm?: number;
  style?: number;
  title?: string;
  legendPosition?: "left" | "top" | "right" | "bottom" | "none";
  showValues?: boolean;
  showCategories?: boolean;
  showSeriesNames?: boolean;
  wrapping?: "inline" | "square" | "tight" | "through" | "topAndBottom" | "behind" | "inFront";
  paragraphIndex?: number;
  current?: boolean;
}
export interface WordAddMathArgs {
  equation: string;
  format?: "unicode" | "latex" | "mathml";
  paragraphIndex?: number;
  current?: boolean;
}
export interface WordAddOleObjectArgs {
  /** Preview image, imported through the same protected image pipeline. */
  source: AiBridgeImageSource;
  data: string;
  applicationId: string;
  widthMm?: number;
  heightMm?: number;
  paragraphIndex?: number;
  current?: boolean;
  name?: string;
}
export interface WordManageFieldsArgs {
  action: "add" | "updateAll" | "clearForms";
  instruction?: string;
  search?: string;
  occurrence?: number;
  matchCase?: boolean;
  paragraphIndex?: number;
}
export interface WordManageLongDocumentArgs {
  action:
    | "addToc" | "updateToc"
    | "addCaption" | "addTableOfFigures" | "updateTableOfFigures"
    | "addCrossReference" | "addFootnote" | "addEndnote"
    | "deleteBookmark";
  paragraphIndex?: number;
  label?: string;
  text?: string;
  excludeLabel?: boolean;
  numberFormat?: string;
  before?: boolean;
  headingLevel?: number;
  separator?: string;
  showPageNumbers?: boolean;
  rightAlignPageNumbers?: boolean;
  leaderType?: string;
  formatAsLinks?: boolean;
  outlineLevels?: number;
  referenceKind?: "caption" | "bookmark" | "heading" | "numbered" | "footnote" | "endnote";
  targetIndex?: number;
  captionIndex?: number;
  referenceType?: string;
  hyperlink?: boolean;
  aboveBelow?: boolean;
  numberSeparator?: string;
  noteText?: string;
  bookmarkName?: string;
}
export interface WordManageCommentsArgs {
  action: "reply" | "edit" | "remove" | "removeAll" | "resolve" | "reopen";
  commentId?: string;
  text?: string;
  author?: string;
  userId?: string;
}
export interface WordManageRevisionsArgs {
  action: "start" | "stop" | "acceptAll" | "rejectAll";
}
export interface WordSetProtectionArgs {
  action: "protect" | "unprotect";
  mode?: "readOnly" | "comments" | "forms";
}
export interface WordContentControlItem { display: string; value?: string; }
export interface WordContentControlDataBinding {
  prefixMapping?: string;
  storeItemId: string;
  xpath: string;
}
export interface WordManageContentControlArgs {
  action: "add" | "update" | "remove" | "clear" | "check";
  kind?: "block" | "inline" | "checkbox" | "comboBox" | "dropDown" | "datePicker" | "picture";
  index?: number;
  tag?: string;
  title?: string;
  text?: string;
  placeholder?: string;
  items?: WordContentControlItem[];
  selectedValue?: string;
  checked?: boolean;
  lock?: "none" | "content" | "control" | "both";
  color?: string;
  appearance?: "boundingBox" | "hidden";
  date?: string;
  dateFormat?: string;
  dataBinding?: WordContentControlDataBinding;
  updateFromXml?: boolean;
  paragraphIndex?: number;
  current?: boolean;
  /** Inline-only table cell target; tableIndex, row, and column must be supplied together. */
  tableIndex?: number;
  row?: number;
  column?: number;
  /** Inline controls default to append; replace clears the selected paragraph or cell first. */
  contentMode?: "append" | "replace";
}
export interface WordManageCustomXmlArgs {
  action:
    | "add" | "replace" | "remove"
    | "insertElement" | "updateElement" | "deleteElement"
    | "insertAttribute" | "updateAttribute" | "deleteAttribute";
  partId?: string;
  xml?: string;
  xpath?: string;
  name?: string;
  value?: string;
}
export interface WordInspectMacrosArgs {
  kind?: "onlyoffice" | "vba";
}
export interface WordMacroDefinition {
  name: string;
  value: string;
  guid?: string;
}
export interface WordSetMacrosArgs {
  macros: WordMacroDefinition[];
  current?: number;
}
export interface WordSetWatermarkArgs extends BasicTextFormat {
  action?: "setText" | "setImage" | "remove";
  text?: string;
  source?: AiBridgeImageSource;
  opacity?: number;
  diagonal?: boolean;
  scale?: number;
}
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
export interface SlidesDrawingSelector {
  objectId?: string;
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
export interface SlidesInspectLayoutsArgs {
  /** One-based master index; omit to inspect all slide masters. */
  masterIndex?: number;
  includeObjects?: boolean;
  includeRaw?: boolean;
}
export interface SlidesInspectThemesArgs {
  /** One-based master index; omit to inspect every slide master. */
  masterIndex?: number;
  /** One-based slide number; when set, inspect that slide's effective theme. */
  slide?: number;
  includeRaw?: boolean;
}
export interface SlidesInspectBuiltinThemesArgs {}
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
export interface SlidesMoveSlideArgs {
  /** One-based source slide number. */
  slide: number;
  /** One-based destination position. */
  toIndex: number;
}
export interface SlidesSetVisibilityArgs { slide: number; visible: boolean; }
export interface SlidesSetSizeArgs {
  preset?: "wide" | "standard" | "custom";
  widthMm?: number;
  heightMm?: number;
  orientation?: "landscape" | "portrait";
}
export interface SlidesApplyLayoutArgs {
  slide: number;
  /** One-based master index; defaults to 1. */
  masterIndex?: number;
  /** One-based layout index returned by slides_inspect_layouts. */
  layoutIndex: number;
  includeRaw?: boolean;
}
export interface SlidesSetShowSettingsArgs { loop: boolean; }
export interface SlidesApplyThemeArgs {
  /** One-based source master index; defaults to 1 when sourceSlide is omitted. */
  sourceMasterIndex?: number;
  /** One-based source slide whose effective theme should be copied. */
  sourceSlide?: number;
  /** One-based destination slide; omit to apply to the presentation. */
  targetSlide?: number;
  includeRaw?: boolean;
}
export interface SlidesApplyBuiltinThemeArgs {
  /** Theme identifier returned by slides_inspect_builtin_themes. */
  theme: number | string;
}
export interface SlidesThemeFonts {
  majorLatin: string;
  minorLatin: string;
  majorEastAsian?: string;
  majorComplex?: string;
  minorEastAsian?: string;
  minorComplex?: string;
  name?: string;
}
export interface SlidesSetThemeArgs {
  /** One-based master index; defaults to 1 when slide is omitted. */
  masterIndex?: number;
  /** One-based slide whose effective theme should be changed. */
  slide?: number;
  /** Exactly 12 colors: dk1, lt1, dk2, lt2, accent1-6, hyperlink, followed hyperlink. */
  colors?: string[];
  colorSchemeName?: string;
  fonts?: SlidesThemeFonts;
  includeRaw?: boolean;
}
export interface SlidesLayoutPlaceholder extends SlidesShapeFormatting {
  type: string;
}
export interface SlidesCreateLayoutArgs {
  /** One-based master index; defaults to 1. */
  masterIndex?: number;
  name?: string;
  fill?: AiBridgeFill;
  followMasterBackground?: boolean;
  placeholders?: SlidesLayoutPlaceholder[];
  /** One-based slide numbers that should immediately use the new layout. */
  applyToSlides?: number[];
  includeRaw?: boolean;
}
export interface SlidesAddTemplateShapeArgs extends SlidesShapeFormatting {
  scope: "master" | "layout";
  masterIndex?: number;
  layoutIndex?: number;
  shapeType: string;
  placeholderType?: string;
}
export interface SlidesManageTemplateObjectArgs extends SlidesShapeFormatting {
  scope: "master" | "layout";
  masterIndex?: number;
  layoutIndex?: number;
  action: "update" | "delete";
  objectId?: string;
  /** Zero-based index returned by slides_inspect_layouts({includeObjects:true}). */
  objectIndex?: number;
  /** Existing object name used as a selector. */
  name?: string;
  /** Replacement name when action is update. */
  newName?: string;
}
export interface SlidesSetTemplateBackgroundArgs {
  scope: "master" | "layout";
  masterIndex?: number;
  layoutIndex?: number;
  mode?: "custom" | "clear" | "master";
  fill?: AiBridgeFill;
}
export interface SlidesParagraphFormat extends TextFormat {
  align?: "left" | "center" | "right" | "both";
  firstLineIndentMm?: number;
  leftIndentMm?: number;
  rightIndentMm?: number;
  spacingBeforePt?: number;
  spacingAfterPt?: number;
  /** Line multiplier for auto, points for exact/atLeast. */
  lineSpacing?: number;
  lineRule?: "auto" | "exact" | "atLeast";
  /** Outline/list level from 0 to 8. */
  level?: number;
  listType?: "none" | "bullet" | "number";
  bulletSymbol?: string;
  numberingType?: string;
  startAt?: number;
}
export interface SlidesTextParagraph extends SlidesParagraphFormat { text: string; }
export interface SlidesSetTextContentArgs extends SlidesObjectTarget {
  paragraphs: SlidesTextParagraph[];
  includeRaw?: boolean;
}
export interface SlidesFormatParagraphsArgs extends SlidesObjectTarget, SlidesParagraphFormat {
  /** One-based paragraph indexes; omit to format every paragraph in the shape. */
  paragraphIndexes?: number[];
  includeRaw?: boolean;
}
export interface SlidesUpdateObjectArgs extends SlidesObjectTarget {
  xMm?: number;
  yMm?: number;
  widthMm?: number;
  heightMm?: number;
  rotationDeg?: number;
  flipH?: boolean;
  flipV?: boolean;
  name?: string;
  line?: AiBridgeLine;
  lineColor?: string;
  lineWidthPt?: number;
  includeRaw?: boolean;
}
export interface SlidesSetHyperlinkArgs extends SlidesObjectTarget {
  action?: "external" | "firstSlide" | "lastSlide" | "nextSlide" | "previousSlide" | "slide" | "remove";
  url?: string;
  targetSlide?: number;
  tooltip?: string;
}
export interface SlidesSetNotesArgs { slide: number; text: string; append?: boolean; }
export interface SlidesAddCommentArgs {
  slide: number;
  text: string;
  xMm?: number;
  yMm?: number;
  author?: string;
  userId?: string;
}
export interface SlidesInspectCommentsArgs {}
export interface SlidesManageCommentArgs {
  action: "update" | "addReply" | "removeReplies" | "delete";
  /** Comment identifier returned by slides_inspect_comments. */
  commentId?: string;
  /** Zero-based comment index returned by slides_inspect_comments. */
  commentIndex?: number;
  text?: string;
  author?: string;
  userId?: string;
  solved?: boolean;
  xMm?: number;
  yMm?: number;
  start?: number;
  count?: number;
}
export interface SlidesSetTransitionArgs {
  slide: number;
  clear?: boolean;
  effect?: string;
  speed?: "slow" | "medium" | "fast";
  durationMs?: number;
  advanceOnClick?: boolean;
  advanceOnTime?: boolean;
  advanceTimeMs?: number;
}
export interface SlidesInspectAnimationsArgs { slide?: number; }
export interface SlidesManageAnimationArgs extends SlidesObjectTarget {
  action: "add" | "update" | "delete" | "clear";
  /** Zero-based flattened effect index returned by slides_inspect_animations. */
  effectIndex?: number;
  sequence?: "main" | "interactive";
  triggerObject?: SlidesDrawingSelector;
  effectType?: string;
  trigger?: "onclick" | "withprevious" | "afterprevious";
  durationMs?: number;
  delayMs?: number;
  repeatCount?: number;
  /** Zero-based destination index within the sequence. */
  toIndex?: number;
}
export interface SlidesTableBorder {
  widthMm?: number;
  color?: string;
  fill?: AiBridgeFill;
  sides?: Array<"top" | "right" | "bottom" | "left">;
}
export interface SlidesTableCellFormat extends SlidesParagraphFormat {
  text?: string;
  fill?: AiBridgeFill;
  backgroundColor?: string;
  verticalAlign?: "top" | "center" | "bottom";
  border?: SlidesTableBorder;
}
export interface SlidesTableLook {
  firstColumn?: boolean;
  firstRow?: boolean;
  lastColumn?: boolean;
  lastRow?: boolean;
  horizontalBanding?: boolean;
  verticalBanding?: boolean;
}
export interface SlidesAddTableArgs {
  slide: number;
  rows: number;
  columns: number;
  data?: unknown[][];
  xMm?: number;
  yMm?: number;
  widthMm?: number;
  heightMm?: number;
  name?: string;
  header?: SlidesTableCellFormat;
  includeRaw?: boolean;
}
export interface SlidesSetTableCellArgs extends SlidesObjectTarget, SlidesTableCellFormat {
  row: number;
  column: number;
  includeRaw?: boolean;
}
export interface SlidesEditTableArgs extends SlidesObjectTarget {
  action: "addRow" | "addColumn" | "removeRow" | "removeColumn" | "mergeCells" | "splitCell";
  row?: number;
  column?: number;
  position?: "before" | "after";
  rowStart?: number;
  rowEnd?: number;
  columnStart?: number;
  columnEnd?: number;
  rows?: number;
  columns?: number;
  includeRaw?: boolean;
}
export interface SlidesFormatTableArgs extends SlidesObjectTarget, SlidesTableCellFormat {
  xMm?: number;
  yMm?: number;
  widthMm?: number;
  heightMm?: number;
  rotationDeg?: number;
  flipH?: boolean;
  flipV?: boolean;
  rowStart?: number;
  rowEnd?: number;
  columnStart?: number;
  columnEnd?: number;
  columnWidthsMm?: number[];
  rowHeightsMm?: number[];
  tableLook?: SlidesTableLook;
  includeRaw?: boolean;
}
export interface SlidesAlignObjectsArgs {
  slide: number;
  targets: SlidesDrawingSelector[];
  align?: "left" | "center" | "right" | "top" | "middle" | "bottom";
  distribute?: "horizontal" | "vertical";
  relativeTo?: "selection" | "slide";
}
export interface SlidesGroupObjectsArgs extends SlidesDrawingSelector {
  slide: number;
  action?: "group" | "ungroup";
  targets?: SlidesDrawingSelector[];
  name?: string;
  includeRaw?: boolean;
}
export interface SlidesReorderObjectArgs extends SlidesObjectTarget {
  action: "front" | "back" | "forward" | "backward";
}
export interface SlidesAddConnectorArgs {
  slide: number;
  connectorType?: string;
  startXmm: number;
  startYmm: number;
  endXmm: number;
  endYmm: number;
  name?: string;
  line?: AiBridgeLine;
  lineColor?: string;
  lineWidthPt?: number;
  includeRaw?: boolean;
}
export type SlidesFreeformCommand =
  | { type: "moveTo" | "lineTo"; xMm: number; yMm: number }
  | { type: "quadBezTo"; controlXmm: number; controlYmm: number; xMm: number; yMm: number }
  | { type: "cubicBezTo"; control1Xmm: number; control1Ymm: number; control2Xmm: number; control2Ymm: number; xMm: number; yMm: number }
  | { type: "arcTo"; widthRadiusMm: number; heightRadiusMm: number; startAngleDeg: number; sweepAngleDeg: number }
  | { type: "close" };
export interface SlidesFreeformPath {
  fill?: string;
  stroke?: boolean;
  commands: SlidesFreeformCommand[];
}
export interface SlidesAddFreeformArgs {
  slide: number;
  paths: SlidesFreeformPath[];
  xMm?: number;
  yMm?: number;
  widthMm?: number;
  heightMm?: number;
  name?: string;
  fill?: AiBridgeFill;
  line?: AiBridgeLine;
  fillColor?: string;
  lineColor?: string;
  lineWidthPt?: number;
  includeRaw?: boolean;
}
export interface SlidesAddTextBoxArgs extends SlidesShapeFormatting {
  slide: number;
  text: string;
}
export interface SlidesAddWordArtArgs extends SlidesShapeFormatting {
  slide: number;
  text: string;
  transform?: string;
}
export interface SlidesAddMathArgs extends SlidesObjectTarget {
  text: string;
  format?: "latex" | "unicode" | "mathml";
  xMm?: number;
  yMm?: number;
  widthMm?: number;
  heightMm?: number;
  rotationDeg?: number;
  name?: string;
  includeRaw?: boolean;
}
export interface SlidesAddImageArgs {
  source: AiBridgeImageSource;
  slide: number;
  xMm?: number;
  yMm?: number;
  widthMm?: number;
  heightMm?: number;
  preserveAspectRatio?: boolean;
  rotationDeg?: number;
  flipH?: boolean;
  flipV?: boolean;
  name?: string;
}
export interface SlidesAddImageShapeArgs {
  source: AiBridgeImageSource;
  slide: number;
  shapeType?: string;
  fillMode?: "stretch" | "tile";
  xMm?: number;
  yMm?: number;
  widthMm?: number;
  heightMm?: number;
  preserveAspectRatio?: boolean;
  rotationDeg?: number;
  flipH?: boolean;
  flipV?: boolean;
  name?: string;
  line?: AiBridgeLine;
  lineColor?: string;
  lineWidthPt?: number;
  includeRaw?: boolean;
}
export interface SlidesAddOleObjectArgs {
  source: AiBridgeImageSource;
  slide: number;
  data: string;
  appId: string;
  xMm?: number;
  yMm?: number;
  widthMm?: number;
  heightMm?: number;
  rotationDeg?: number;
  name?: string;
  includeRaw?: boolean;
}
export interface SlidesInspectMacrosArgs { kind?: "onlyoffice" | "vba"; }
export interface SlidesSetMacrosArgs { content: Record<string, unknown>; }
export interface SlidesControlSlideshowArgs {
  action: "start" | "end" | "pause" | "resume" | "next" | "previous" | "goto";
  /** One-based slide number required by goto. */
  slide?: number;
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
export interface SheetsInspectRangeArgs extends SheetTarget { range: string; includeValues?: boolean; }
export interface SheetsSetValuesArgs extends SheetTarget { range: string; values: unknown; }
export interface SheetsSetFormulaArgs extends SheetTarget { range: string; formula: string; }
export interface SheetsSetArrayFormulaArgs extends SheetTarget { range: string; formula: string; }
export interface SheetsReplaceTextArgs extends SheetTarget { range?: string; search: string; replace: string; }
export interface SheetsFormatRangeArgs extends SheetTarget {
  range: string;
  fontSize?: number;
  fontName?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikeout?: boolean;
  fontColor?: string;
  fillColor?: string;
  horizontalAlign?: string;
  verticalAlign?: string;
  numberFormat?: string;
  wrap?: boolean;
  orientation?: string | number;
  columnWidth?: number;
  rowHeight?: number;
  borders?: Array<{ side: string; style: string; color: string }>;
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
export interface SheetsManageSheetArgs extends SheetTarget {
  action: "setVisibility" | "setActive" | "moveBefore" | "copy";
  visible?: boolean;
  beforeSheet?: string;
  newName?: string;
}
export interface SheetsManageRangeArgs extends SheetTarget {
  action:
    | "merge" | "unmerge" | "insert" | "delete" | "copy" | "cut"
    | "clear" | "clearContents" | "clearFormats" | "clearHyperlinks"
    | "autoFit" | "setHidden" | "fillDown" | "fillUp" | "fillLeft" | "fillRight" | "select";
  range: string;
  destinationSheet?: string;
  destination?: string;
  shift?: string;
  across?: boolean;
  hidden?: boolean;
  rows?: boolean;
  columns?: boolean;
}
export interface SheetsRichTextRun {
  start: number;
  length: number;
  text?: string;
  fontName?: string;
  fontSize?: number;
  fontColor?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean | string;
  strikeout?: boolean;
}
export interface SheetsSetRichTextArgs extends SheetTarget { range: string; text?: string; runs: SheetsRichTextRun[]; }
export interface SheetsInspectNamesArgs {}
export interface SheetsManageNamesArgs {
  action: "add" | "update" | "delete";
  name: string;
  newName?: string;
  refersTo?: string;
}
export interface SheetsRecalculateArgs { mode?: "formulas" | "pivots" | "all"; }
export interface SheetsSortArgs extends SheetTarget {
  range: string;
  keys: Array<{ range: string; order?: "xlAscending" | "xlDescending" }>;
  header?: string;
  orientation?: string;
}
export interface SheetsFilterArgs extends SheetTarget {
  action: "set" | "showAll" | "reapply";
  range?: string;
  field?: number;
  criteria1?: unknown;
  operator?: string;
  criteria2?: unknown;
  visibleDropDown?: boolean;
}
export interface SheetsInspectTablesArgs extends SheetTarget {
  tableIndex?: number;
  name?: string;
  maxTables?: number;
}
export interface SheetsManageTableArgs extends SheetTarget {
  action: "create" | "format" | "update" | "resize" | "delete" | "unlist";
  range?: string;
  sourceType?: string;
  tableIndex?: number;
  tableName?: string;
  name?: string;
  newName?: string;
  style?: string;
  showTotals?: boolean;
  showHeaders?: boolean;
  rowStripes?: boolean;
  columnStripes?: boolean;
  firstColumn?: boolean;
  lastColumn?: boolean;
  showAutoFilter?: boolean;
  showAutoFilterDropDown?: boolean;
  summary?: string;
  alternativeText?: string;
}
export interface SheetsManageConditionalFormatArgs extends SheetTarget {
  action: "add" | "deleteAll";
  range: string;
  type?: "cellValue" | "expression" | "uniqueValues" | "duplicateValues" | "colorScale" | "dataBar" | "iconSet" | "top10" | "aboveAverage";
  operator?: string;
  formula1?: unknown;
  formula2?: unknown;
  scale?: 2 | 3;
  fillColor?: string;
  fontColor?: string;
  bold?: boolean;
}
export interface SheetsManageValidationArgs extends SheetTarget {
  action: "add" | "modify" | "delete";
  range: string;
  type?: string;
  alertStyle?: string;
  operator?: string;
  formula1?: unknown;
  formula2?: unknown;
  ignoreBlank?: boolean;
  showInput?: boolean;
  showError?: boolean;
  inputTitle?: string;
  inputMessage?: string;
  errorTitle?: string;
  errorMessage?: string;
}
export interface SheetsInspectPivotsArgs {}
export interface SheetsManagePivotArgs extends SheetTarget {
  action: "createNewSheet" | "createExisting" | "update" | "refresh" | "refreshAll" | "clear";
  name?: string;
  sourceSheet?: string;
  sourceRange?: string;
  destinationSheet?: string;
  destinationRange?: string;
  fields?: Record<string, unknown>;
  dataFields?: Array<string | { name: string; function?: string }>;
  style?: string;
  title?: string;
  description?: string;
  rowGrand?: boolean;
  columnGrand?: boolean;
}
export interface SheetsInspectDrawingsArgs extends SheetTarget { maxDrawings?: number; includeRaw?: boolean; }
export interface SheetsManageDrawingArgs extends SheetTarget {
  action: "addImage" | "addShape" | "addTextBox" | "addOleObject" | "update" | "delete" | "copy";
  destinationSheet?: string;
  drawingIndex?: number;
  name?: string;
  source?: AiBridgeImageSource;
  shapeType?: string;
  text?: string;
  data?: string;
  appId?: string;
  widthMm?: number;
  heightMm?: number;
  fromColumn?: number;
  fromRow?: number;
  columnOffsetMm?: number;
  rowOffsetMm?: number;
  rotationDeg?: number;
  flipH?: boolean;
  flipV?: boolean;
  fill?: AiBridgeFill;
  line?: AiBridgeLine;
  includeRaw?: boolean;
}
export interface SheetsManageHyperlinkArgs extends SheetTarget {
  action: "set" | "delete";
  range: string;
  url?: string;
  location?: string;
  displayText?: string;
  tooltip?: string;
}
export interface SheetsInspectCommentsArgs extends SheetTarget {}
export interface SheetsManageCommentsArgs extends SheetTarget {
  action: "add" | "update" | "delete" | "setSolved" | "addReply" | "removeReplies";
  range?: string;
  commentId?: string;
  text?: string;
  author?: string;
  userId?: string;
  solved?: boolean;
  start?: number;
  count?: number;
  removeAll?: boolean;
}
export interface SheetsInspectFreezePanesArgs extends SheetTarget {}
export interface SheetsManageFreezePanesArgs extends SheetTarget {
  action: "unfreeze" | "freezeAt" | "freezeRows" | "freezeColumns";
  range?: string;
  count?: number;
}
export interface SheetsInspectPropertiesArgs { customNames?: string[]; }
export interface SheetsCoreProperties {
  category?: unknown;
  contentStatus?: unknown;
  created?: unknown;
  creator?: unknown;
  description?: unknown;
  identifier?: unknown;
  keywords?: unknown;
  language?: unknown;
  lastModifiedBy?: unknown;
  lastPrinted?: unknown;
  modified?: unknown;
  revision?: unknown;
  subject?: unknown;
  title?: unknown;
  version?: unknown;
}
export interface SheetsManagePropertiesArgs {
  core?: SheetsCoreProperties;
  custom?: Record<string, unknown>;
}
export interface SheetsInspectProtectedRangesArgs extends SheetTarget {}
export interface SheetsManageProtectedRangesArgs extends SheetTarget {
  action: "add" | "update" | "addUser" | "deleteUser";
  title: string;
  newTitle?: string;
  range?: string;
  anyoneType?: string;
  userId?: string;
  userName?: string;
  permission?: string;
}
export interface SheetsInspectPageLayoutArgs extends SheetTarget {}
export type SheetsManagePageLayoutArgs = SheetTarget & {
  orientation?: string;
  margins?: { top?: number; right?: number; bottom?: number; left?: number };
  printGridlines?: boolean;
  printHeadings?: boolean;
} & (
  | { orientation: string }
  | { margins: { top?: number; right?: number; bottom?: number; left?: number } }
  | { printGridlines: boolean }
  | { printHeadings: boolean }
);
export interface SheetsInspectMacrosArgs { kind?: "onlyoffice" | "vba"; }
export interface SheetsSetMacrosArgs { content: Record<string, unknown>; }

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
  word_add_image: WordAddImageArgs;
  word_inspect_advanced: WordInspectAdvancedArgs;
  word_set_document_properties: WordSetDocumentPropertiesArgs;
  word_manage_section: WordManageSectionArgs;
  word_manage_style: WordManageStyleArgs;
  word_set_tabs: WordSetTabsArgs;
  word_set_numbering: WordSetNumberingArgs;
  word_format_table_advanced: WordFormatTableAdvancedArgs;
  word_add_nested_table: WordAddNestedTableArgs;
  word_manage_drawing: WordManageDrawingArgs;
  word_add_shape: WordAddShapeArgs;
  word_add_chart: WordAddChartArgs;
  word_add_math: WordAddMathArgs;
  word_add_ole_object: WordAddOleObjectArgs;
  word_manage_fields: WordManageFieldsArgs;
  word_manage_long_document: WordManageLongDocumentArgs;
  word_manage_comments: WordManageCommentsArgs;
  word_manage_revisions: WordManageRevisionsArgs;
  word_set_protection: WordSetProtectionArgs;
  word_manage_content_control: WordManageContentControlArgs;
  word_manage_custom_xml: WordManageCustomXmlArgs;
  word_inspect_macros: WordInspectMacrosArgs;
  word_set_macros: WordSetMacrosArgs;
  word_set_watermark: WordSetWatermarkArgs;
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
  slides_inspect_layouts: SlidesInspectLayoutsArgs;
  slides_inspect_themes: SlidesInspectThemesArgs;
  slides_inspect_builtin_themes: SlidesInspectBuiltinThemesArgs;
  slides_inspect_objects: SlidesInspectObjectsArgs;
  slides_replace_text: SlidesReplaceTextArgs;
  slides_scale_font: SlidesScaleFontArgs;
  slides_format_text: SlidesFormatTextArgs;
  slides_format_selection: SlidesFormatSelectionArgs;
  slides_add_slide: SlidesAddSlideArgs;
  slides_duplicate_slide: SlidesDuplicateSlideArgs;
  slides_delete_slide: SlidesDeleteSlideArgs;
  slides_move_slide: SlidesMoveSlideArgs;
  slides_set_visibility: SlidesSetVisibilityArgs;
  slides_set_size: SlidesSetSizeArgs;
  slides_apply_layout: SlidesApplyLayoutArgs;
  slides_set_show_settings: SlidesSetShowSettingsArgs;
  slides_apply_theme: SlidesApplyThemeArgs;
  slides_apply_builtin_theme: SlidesApplyBuiltinThemeArgs;
  slides_set_theme: SlidesSetThemeArgs;
  slides_create_layout: SlidesCreateLayoutArgs;
  slides_add_template_shape: SlidesAddTemplateShapeArgs;
  slides_manage_template_object: SlidesManageTemplateObjectArgs;
  slides_set_template_background: SlidesSetTemplateBackgroundArgs;
  slides_set_text_content: SlidesSetTextContentArgs;
  slides_format_paragraphs: SlidesFormatParagraphsArgs;
  slides_update_object: SlidesUpdateObjectArgs;
  slides_set_hyperlink: SlidesSetHyperlinkArgs;
  slides_set_notes: SlidesSetNotesArgs;
  slides_add_comment: SlidesAddCommentArgs;
  slides_inspect_comments: SlidesInspectCommentsArgs;
  slides_manage_comment: SlidesManageCommentArgs;
  slides_set_transition: SlidesSetTransitionArgs;
  slides_inspect_animations: SlidesInspectAnimationsArgs;
  slides_manage_animation: SlidesManageAnimationArgs;
  slides_add_table: SlidesAddTableArgs;
  slides_set_table_cell: SlidesSetTableCellArgs;
  slides_edit_table: SlidesEditTableArgs;
  slides_format_table: SlidesFormatTableArgs;
  slides_align_objects: SlidesAlignObjectsArgs;
  slides_group_objects: SlidesGroupObjectsArgs;
  slides_reorder_object: SlidesReorderObjectArgs;
  slides_add_connector: SlidesAddConnectorArgs;
  slides_add_freeform: SlidesAddFreeformArgs;
  slides_add_textbox: SlidesAddTextBoxArgs;
  slides_add_word_art: SlidesAddWordArtArgs;
  slides_add_math: SlidesAddMathArgs;
  slides_add_image: SlidesAddImageArgs;
  slides_add_image_shape: SlidesAddImageShapeArgs;
  slides_add_ole_object: SlidesAddOleObjectArgs;
  slides_set_background: SlidesSetBackgroundArgs;
  slides_add_shape: SlidesAddShapeArgs;
  slides_update_shape: SlidesUpdateShapeArgs;
  slides_delete_object: SlidesDeleteObjectArgs;
  slides_inspect_charts: SlidesInspectChartsArgs;
  slides_add_chart: SlidesAddChartArgs;
  slides_update_chart: SlidesUpdateChartArgs;
  slides_delete_chart: SlidesDeleteChartArgs;
  slides_inspect_macros: SlidesInspectMacrosArgs;
  slides_set_macros: SlidesSetMacrosArgs;
  slides_control_slideshow: SlidesControlSlideshowArgs;
  sheets_inspect: SheetsInspectArgs;
  sheets_inspect_range: SheetsInspectRangeArgs;
  sheets_set_values: SheetsSetValuesArgs;
  sheets_set_formula: SheetsSetFormulaArgs;
  sheets_set_array_formula: SheetsSetArrayFormulaArgs;
  sheets_replace_text: SheetsReplaceTextArgs;
  sheets_format_range: SheetsFormatRangeArgs;
  sheets_add_sheet: SheetsAddSheetArgs;
  sheets_rename_sheet: SheetsRenameSheetArgs;
  sheets_delete_sheet: SheetsDeleteSheetArgs;
  sheets_manage_sheet: SheetsManageSheetArgs;
  sheets_manage_range: SheetsManageRangeArgs;
  sheets_set_rich_text: SheetsSetRichTextArgs;
  sheets_inspect_names: SheetsInspectNamesArgs;
  sheets_manage_names: SheetsManageNamesArgs;
  sheets_recalculate: SheetsRecalculateArgs;
  sheets_sort: SheetsSortArgs;
  sheets_filter: SheetsFilterArgs;
  sheets_inspect_tables: SheetsInspectTablesArgs;
  sheets_manage_table: SheetsManageTableArgs;
  sheets_manage_conditional_format: SheetsManageConditionalFormatArgs;
  sheets_manage_validation: SheetsManageValidationArgs;
  sheets_inspect_pivots: SheetsInspectPivotsArgs;
  sheets_manage_pivot: SheetsManagePivotArgs;
  sheets_inspect_drawings: SheetsInspectDrawingsArgs;
  sheets_manage_drawing: SheetsManageDrawingArgs;
  sheets_manage_hyperlink: SheetsManageHyperlinkArgs;
  sheets_inspect_comments: SheetsInspectCommentsArgs;
  sheets_manage_comments: SheetsManageCommentsArgs;
  sheets_inspect_freeze_panes: SheetsInspectFreezePanesArgs;
  sheets_manage_freeze_panes: SheetsManageFreezePanesArgs;
  sheets_inspect_properties: SheetsInspectPropertiesArgs;
  sheets_manage_properties: SheetsManagePropertiesArgs;
  sheets_inspect_protected_ranges: SheetsInspectProtectedRangesArgs;
  sheets_manage_protected_ranges: SheetsManageProtectedRangesArgs;
  sheets_inspect_page_layout: SheetsInspectPageLayoutArgs;
  sheets_manage_page_layout: SheetsManagePageLayoutArgs;
  sheets_inspect_macros: SheetsInspectMacrosArgs;
  sheets_set_macros: SheetsSetMacrosArgs;
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
  addImage(args: WordAddImageArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  inspectAdvanced(args?: WordInspectAdvancedArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  setDocumentProperties(args: WordSetDocumentPropertiesArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  manageSection(args: WordManageSectionArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  manageStyle(args: WordManageStyleArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  setTabs(args: WordSetTabsArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  setNumbering(args: WordSetNumberingArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  formatTableAdvanced(args: WordFormatTableAdvancedArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  addNestedTable(args: WordAddNestedTableArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  manageDrawing(args: WordManageDrawingArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  addShape(args: WordAddShapeArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  addChart(args: WordAddChartArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  addMath(args: WordAddMathArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  addOleObject(args: WordAddOleObjectArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  manageFields(args: WordManageFieldsArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  manageLongDocument(args: WordManageLongDocumentArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  manageComments(args: WordManageCommentsArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  manageRevisions(args: WordManageRevisionsArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  setProtection(args: WordSetProtectionArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  manageContentControl(args: WordManageContentControlArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  manageCustomXml(args: WordManageCustomXmlArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  inspectMacros(args?: WordInspectMacrosArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  setMacros(args: WordSetMacrosArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  setWatermark(args: WordSetWatermarkArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
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
  inspectLayouts(args?: SlidesInspectLayoutsArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  inspectThemes(args?: SlidesInspectThemesArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  inspectBuiltinThemes(args?: SlidesInspectBuiltinThemesArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  inspectObjects(args?: SlidesInspectObjectsArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  replaceText(args: SlidesReplaceTextArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  scaleFont(args: SlidesScaleFontArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  formatText(args: SlidesFormatTextArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  formatSelection(args: SlidesFormatSelectionArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  addSlide(args: SlidesAddSlideArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  duplicateSlide(args: SlidesDuplicateSlideArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  deleteSlide(args: SlidesDeleteSlideArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  moveSlide(args: SlidesMoveSlideArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  setVisibility(args: SlidesSetVisibilityArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  setSize(args: SlidesSetSizeArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  applyLayout(args: SlidesApplyLayoutArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  setShowSettings(args: SlidesSetShowSettingsArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  applyTheme(args: SlidesApplyThemeArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  applyBuiltinTheme(args: SlidesApplyBuiltinThemeArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  setTheme(args: SlidesSetThemeArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  createLayout(args: SlidesCreateLayoutArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  addTemplateShape(args: SlidesAddTemplateShapeArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  manageTemplateObject(args: SlidesManageTemplateObjectArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  setTemplateBackground(args: SlidesSetTemplateBackgroundArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  setTextContent(args: SlidesSetTextContentArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  formatParagraphs(args: SlidesFormatParagraphsArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  updateObject(args: SlidesUpdateObjectArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  setHyperlink(args: SlidesSetHyperlinkArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  setNotes(args: SlidesSetNotesArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  addComment(args: SlidesAddCommentArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  inspectComments(args?: SlidesInspectCommentsArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  manageComment(args: SlidesManageCommentArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  setTransition(args: SlidesSetTransitionArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  inspectAnimations(args?: SlidesInspectAnimationsArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  manageAnimation(args: SlidesManageAnimationArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  addTable(args: SlidesAddTableArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  setTableCell(args: SlidesSetTableCellArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  editTable(args: SlidesEditTableArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  formatTable(args: SlidesFormatTableArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  alignObjects(args: SlidesAlignObjectsArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  groupObjects(args: SlidesGroupObjectsArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  reorderObject(args: SlidesReorderObjectArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  addConnector(args: SlidesAddConnectorArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  addFreeform(args: SlidesAddFreeformArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  addTextBox(args: SlidesAddTextBoxArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  addWordArt(args: SlidesAddWordArtArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  addMath(args: SlidesAddMathArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  addImage(args: SlidesAddImageArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  addImageShape(args: SlidesAddImageShapeArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  addOleObject(args: SlidesAddOleObjectArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  setBackground(args: SlidesSetBackgroundArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  addShape(args: SlidesAddShapeArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  updateShape(args: SlidesUpdateShapeArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  deleteObject(args: SlidesDeleteObjectArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  inspectCharts(args?: SlidesInspectChartsArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  addChart(args: SlidesAddChartArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  updateChart(args: SlidesUpdateChartArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  deleteChart(args: SlidesDeleteChartArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  inspectMacros(args?: SlidesInspectMacrosArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  setMacros(args: SlidesSetMacrosArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  controlSlideshow(args: SlidesControlSlideshowArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
}

export interface AiBridgeSheetsApi {
  inspect(args?: SheetsInspectArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  inspectRange(args: SheetsInspectRangeArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  setValues(args: SheetsSetValuesArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  setFormula(args: SheetsSetFormulaArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  setArrayFormula(args: SheetsSetArrayFormulaArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  replaceText(args: SheetsReplaceTextArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  formatRange(args: SheetsFormatRangeArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  addSheet(args: SheetsAddSheetArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  renameSheet(args: SheetsRenameSheetArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  deleteSheet(args: SheetsDeleteSheetArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  manageSheet(args: SheetsManageSheetArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  manageRange(args: SheetsManageRangeArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  setRichText(args: SheetsSetRichTextArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  inspectNames(args?: SheetsInspectNamesArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  manageNames(args: SheetsManageNamesArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  recalculate(args?: SheetsRecalculateArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  sort(args: SheetsSortArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  filter(args: SheetsFilterArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  inspectTables(args?: SheetsInspectTablesArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  manageTable(args: SheetsManageTableArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  manageConditionalFormat(args: SheetsManageConditionalFormatArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  manageValidation(args: SheetsManageValidationArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  inspectPivots(args?: SheetsInspectPivotsArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  managePivot(args: SheetsManagePivotArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  inspectDrawings(args?: SheetsInspectDrawingsArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  manageDrawing(args: SheetsManageDrawingArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  manageHyperlink(args: SheetsManageHyperlinkArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  inspectComments(args?: SheetsInspectCommentsArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  manageComments(args: SheetsManageCommentsArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  inspectFreezePanes(args?: SheetsInspectFreezePanesArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  manageFreezePanes(args: SheetsManageFreezePanesArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  inspectProperties(args?: SheetsInspectPropertiesArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  manageProperties(args: SheetsManagePropertiesArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  inspectProtectedRanges(args?: SheetsInspectProtectedRangesArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  manageProtectedRanges(args: SheetsManageProtectedRangesArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  inspectPageLayout(args?: SheetsInspectPageLayoutArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  managePageLayout(args: SheetsManagePageLayoutArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  inspectMacros(args?: SheetsInspectMacrosArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  setMacros(args: SheetsSetMacrosArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  addChart(args: SheetsAddChartArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  inspectCharts(args?: SheetsInspectChartsArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  updateChart(args: SheetsUpdateChartArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  deleteChart(args: SheetsDeleteChartArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
}

export type AiBridgeEventName = "ready" | "reload" | "error";

export interface AiBridgeApi {
  readonly version: "0.4.0";
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
