export type AiBridgeEditorType = "word" | "slide" | "cell";
export type AiBridgeControl = "save" | "history" | "undo" | "redo";
export type AiBridgeErrorCode =
  | "AI_BRIDGE_ERROR"
  | "NOT_READY"
  | "CONNECTION_TIMEOUT"
  | "TIMEOUT"
  | "BRIDGE_TIMEOUT"
  | "HTTP_RELAY_NETWORK_ERROR"
  | "HTTP_RELAY_INVALID_RESPONSE"
  | "EDITOR_TOKEN_REQUIRED"
  | "INVALID_EDITOR_TOKEN"
  | "EDITOR_TOKEN_EXPIRED"
  | "INVALID_BRIDGE_RESUME_TOKEN"
  | "BRIDGE_RESUME_TOKEN_EXPIRED"
  | "INVALID_RELAY_SESSION"
  | "NO_ACTIVE_EDITOR"
  | "INVALID_TARGET_WINDOW"
  | "INVALID_TARGET_ORIGIN"
  | "INVALID_CLIENT_ID"
  | "INVALID_REQUEST_ID"
  | "REQUEST_ID_CONFLICT"
  | "REQUEST_IN_FLIGHT"
  | "INVALID_COMMAND"
  | "INVALID_TOOL_CALL"
  | "INVALID_ARGUMENTS"
  | "INVALID_TOOL_ARGUMENTS"
  | "INVALID_TARGET"
  | "ARGUMENTS_TOO_LARGE"
  | "TOO_MANY_CALLS"
  | "TOOL_NOT_ALLOWED"
  | "CONTROL_NOT_ALLOWED"
  | "METHOD_NOT_ALLOWED"
  | "MESSAGE_NOT_SUPPORTED"
  | "DOCUMENT_NOT_CONFIGURED"
  | "DOCUMENT_MISMATCH"
  | "EDITOR_MISMATCH"
  | "PERSISTENCE_NOT_AVAILABLE"
  | "PERSISTENCE_INVALID_RESPONSE"
  | "PERSISTENCE_TIMEOUT"
  | "PERSISTENCE_FAILED"
  | "INVALID_IMAGE_SOURCE"
  | "IMAGE_FETCH_BLOCKED"
  | "IMAGE_FETCH_FAILED"
  | "IMAGE_TOO_LARGE"
  | "UNSUPPORTED_IMAGE_FORMAT"
  | "IMAGE_ASSET_EXPIRED"
  | "IMAGE_API_UNSUPPORTED"
  | "WORD_API_UNSUPPORTED"
  | "SLIDES_API_UNSUPPORTED"
  | "SHEETS_API_UNSUPPORTED"
  | "BATCH_DEPENDENCY_REQUIRES_SPLIT"
  | "SHEETS_CHART_CREATE_FAILED"
  | "SHEETS_RUNTIME_INCOMPATIBLE"
  | "SHEETS_CHART_PARTIAL_MUTATION"
  | "EXECUTION_FAILED"
  | "INVALID_LISTENER"
  | "CLIENT_DESTROYED"
  | "CONTRACT_MISMATCH"
  | "CONTRACT_VERSION_MISMATCH"
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
  contractVersion?: "0.2.4";
  contractSha256?: string;
  runtime?: {
    product: "ONLYOFFICE";
    version: string | null;
    edition: "community" | "enterprise" | "unknown";
  };
  features?: {
    sheets?: {
      nativeTables: { create: boolean; inspect: boolean };
      rangeStyleTables: { create: boolean };
      conditionalFormatting: { create: boolean };
      rangeFill: { fillDown: boolean; fillUp: boolean; fillLeft: boolean; fillRight: boolean };
      arrayFormula: { set: boolean };
      validation: { manage: boolean };
      comments: { create: boolean; inspect: boolean; update: boolean; delete: boolean };
      freezePanes: { inspect: boolean; manage: boolean };
      charts: { create: boolean; inspect: boolean; update: boolean; addSeriesOnCreate: boolean; delete: boolean };
    };
  };
}

export interface AiBridgeState {
  version: "0.2.4";
  protocolVersion: 1;
  pluginGuid: "asc.{A17E5F31-64AA-4E37-9A42-8D430814C2F6}";
  contractVersion: "0.2.4";
  contractSha256: string;
  ready: boolean;
  documentReady?: boolean;
  capabilityProbeComplete?: boolean;
  capabilityProbedAt?: number | null;
  saveReady?: boolean;
  saveStatus?: "idle" | "editor-saving" | "editor-saved" | "saving" | "saved" | "failed" | string;
  editorType: AiBridgeEditorType | null;
  context: AiBridgeContext;
  capabilities: AiBridgeCapabilities | null;
}

export interface AiBridgeArgumentNormalization {
  toolCallIndex: number;
  /** Source path only; parameter values are never echoed. */
  path: string;
  canonicalPath: string;
  kind: "propertyAlias" | "canonicalWins" | "enumAlias" | "enumCanonicalization";
}

export interface AiBridgeExecutionResult {
  ok?: boolean;
  editorType?: AiBridgeEditorType;
  changed: number;
  needsSave?: boolean;
  results: Array<Record<string, unknown>>;
  persisted: boolean;
  forceSave?: Record<string, unknown>;
  persistence?: {
    status: "saved" | "no_changes" | "failed";
    editorSaved: boolean;
    forceSave: {
      accepted: boolean;
      noChanges: boolean;
      beforeMtime: number | null;
      afterMtime: number | null;
      commandError: number | null;
    };
  };
  argumentNormalizations?: AiBridgeArgumentNormalization[];
}

export interface AiBridgeWordValidationResult {
  ok: true;
  valid: true;
  editorType: "word";
  toolCalls: number;
  contractVersion: "0.2.4";
  contractSha256: string;
  argumentNormalizations?: AiBridgeArgumentNormalization[];
}

export interface AiBridgeServiceResult extends Record<string, unknown> {
  persisted?: boolean;
  key?: string;
  fileName?: string;
}

export interface AiBridgeToolValidationError {
  toolCallIndex: number;
  tool: string;
  path: string;
  keyword: string;
  message: string;
}

export interface AiBridgeToolArgumentsErrorDetails {
  validationErrors: AiBridgeToolValidationError[];
  completedToolCalls: 0;
  partialMutationPossible: false;
}

export class AiBridgeError extends Error {
  readonly code: AiBridgeErrorCode;
  readonly requestId: string | null;
  readonly details?: AiBridgeToolArgumentsErrorDetails | Record<string, unknown>;
}

export interface BasicTextFormat {
  /** Font size in points (pt). */
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;
}
export type AiBridgeTableCellScalar = string | number | boolean | null;
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
  /** @deprecated Point-valued legacy alias. Prefer characterSpacingPt. Never pass twips. */
  characterSpacing?: number;
  /** Additional character spacing in points (pt). */
  characterSpacingPt?: number;
  vertAlign?: "baseline" | "subscript" | "superscript";
}
export interface WordParagraphFormat extends TextFormat {
  align?: "left" | "center" | "right" | "both";
  styleName?: string;
  headingLevel?: number;
  outlineLevel?: number;
  /** @deprecated Point-valued legacy alias. Prefer spacingBeforePt. Never pass twips. */
  spacingBefore?: number;
  /** @deprecated Point-valued legacy alias. Prefer spacingAfterPt. Never pass twips. */
  spacingAfter?: number;
  /** Paragraph spacing before in points (pt). */
  spacingBeforePt?: number;
  /** Paragraph spacing after in points (pt). */
  spacingAfterPt?: number;
  /** Line multiplier for auto; points for exact/atLeast. */
  lineSpacing?: number;
  lineRule?: "auto" | "exact" | "atLeast";
  /** @deprecated Point-valued legacy alias. Prefer firstLineIndentPt. Never pass twips. */
  firstLineIndent?: number;
  /** @deprecated Point-valued legacy alias. Prefer leftIndentPt. Never pass twips. */
  leftIndent?: number;
  /** @deprecated Point-valued legacy alias. Prefer rightIndentPt. Never pass twips. */
  rightIndent?: number;
  /** First-line or hanging indent in points (pt); negative values create a hanging indent. */
  firstLineIndentPt?: number;
  /** Left paragraph indent in points (pt). */
  leftIndentPt?: number;
  /** Right paragraph indent in points (pt). */
  rightIndentPt?: number;
  keepLines?: boolean;
  keepNext?: boolean;
  widowControl?: boolean;
  contextualSpacing?: boolean;
  pageBreakBefore?: boolean;
}
export interface WordParagraphTarget {
  /** Prefer stable IDs returned by word_inspect. */
  paragraphId?: string | number;
  internalId?: string | number;
  paragraphIndex?: number;
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
export interface WordAddBookmarkArgs {
  search: string;
  occurrence: number;
  name: string;
  matchCase?: boolean;
}
export interface WordAddImageArgs extends WordStableParagraphTarget {
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
export interface WordTableCellInput extends BasicTextFormat {
  text: string;
  paragraphStyleName?: string;
  /** @deprecated Compatibility alias for paragraphStyleName. */
  styleName?: string;
  characterStyleName?: string;
  /** @deprecated Compatibility alias for color. Prefer color. */
  textColor?: string;
  underline?: boolean;
  highlightColor?: string;
  /** @deprecated Point-valued legacy alias. Prefer characterSpacingPt. Never pass twips. */
  characterSpacing?: number;
  characterSpacingPt?: number;
  backgroundColor?: string;
  align?: "left" | "center" | "right" | "both";
  /** @deprecated Point-valued legacy aliases. Prefer the corresponding *Pt fields. */
  spacingBefore?: number;
  spacingAfter?: number;
  spacingBeforePt?: number;
  spacingAfterPt?: number;
  lineSpacing?: number;
  lineRule?: "auto" | "exact" | "atLeast";
  /** @deprecated Point-valued legacy aliases. Prefer the corresponding *Pt fields. */
  firstLineIndent?: number;
  leftIndent?: number;
  rightIndent?: number;
  firstLineIndentPt?: number;
  leftIndentPt?: number;
  rightIndentPt?: number;
  verticalAlign?: "top" | "center" | "bottom";
  widthPercent?: number;
}
export interface WordAddTableArgs extends BasicTextFormat, WordStableParagraphTarget {
  rows: number;
  cols: number;
  /** @deprecated Compatibility alias for cols. Prefer cols. */
  columns?: number;
  data?: Array<Array<AiBridgeTableCellScalar | WordTableCellInput>>;
  widthPercent?: number;
  /** Whole-table alignment; cell paragraph alignment belongs in data cell objects. */
  align?: "left" | "center" | "right";
  /** @deprecated Compatibility alias for tableStyleName. */
  styleName?: string;
  tableStyleName?: string;
  cellParagraphStyleName?: string;
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
  paragraphStyleName?: string;
  backgroundColor?: string;
  verticalAlign?: "top" | "center" | "bottom";
  widthPercent?: number;
}
export interface WordFormatTableArgs extends BasicTextFormat {
  tableIndex: number;
  widthPercent?: number;
  align?: "left" | "center" | "right";
  /** @deprecated Compatibility alias for tableStyleName. */
  styleName?: string;
  tableStyleName?: string;
  cellParagraphStyleName?: string;
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
  /** Custom page width in millimetres (mm). */
  widthMm?: number;
  /** Custom page height in millimetres (mm). */
  heightMm?: number;
  /** Page and header/footer distances in millimetres (mm). */
  marginLeftMm?: number;
  marginTopMm?: number;
  marginRightMm?: number;
  marginBottomMm?: number;
  headerDistanceMm?: number;
  footerDistanceMm?: number;
  titlePage?: boolean;
  /** @deprecated Compatibility alias for titlePage. Prefer titlePage. */
  differentFirstPage?: boolean;
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
  /** @deprecated Compatibility alias for creator. Prefer creator. */
  author?: string;
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
  /** @deprecated One-based paragraph index that ends the current section. */
  paragraphIndex?: number;
  endParagraphIndex?: number;
  boundary?: {
    position: "before" | "after";
    paragraphId?: string | number;
    internalId?: string | number;
    paragraphIndex?: number;
    search?: string;
    occurrence?: number;
    matchCase?: boolean;
    matchMode?: "contains" | "exact";
  };
  type?: "continuous" | "nextPage" | "evenPage" | "oddPage";
  startPageNumber?: number;
  titlePage?: boolean;
  /** @deprecated Compatibility alias for titlePage. Prefer titlePage. */
  differentFirstPage?: boolean;
  evenAndOddHeaders?: boolean;
  columns?: {
    count?: number;
    spaceMm?: number;
    entries?: WordSectionColumn[];
  };
}
export interface WordManageStyleArgs {
  action?: "create" | "update";
  name: string;
  type?: "paragraph" | "character" | "table" | "numbering";
  basedOn?: string;
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikeout?: boolean;
  color?: string;
  highlightColor?: string;
  /** @deprecated Point-valued legacy alias. Prefer characterSpacingPt. */
  characterSpacing?: number;
  characterSpacingPt?: number;
  align?: "left" | "center" | "right" | "both";
  /** @deprecated Point-valued legacy aliases. Prefer the corresponding *Pt fields. */
  spacingBefore?: number;
  spacingAfter?: number;
  spacingBeforePt?: number;
  spacingAfterPt?: number;
  lineSpacing?: number;
  lineRule?: "auto" | "exact" | "atLeast";
  /** @deprecated Point-valued legacy aliases. Prefer the corresponding *Pt fields. */
  firstLineIndent?: number;
  leftIndent?: number;
  rightIndent?: number;
  firstLineIndentPt?: number;
  leftIndentPt?: number;
  rightIndentPt?: number;
  keepLines?: boolean;
  keepNext?: boolean;
  pageBreakBefore?: boolean;
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
export interface WordNumberingAssignment {
  level: number;
  paragraphId?: string | number;
  internalId?: string | number;
  paragraphIndex?: number;
  search?: string;
  occurrence?: number;
  matchCase?: boolean;
  matchMode?: "contains" | "exact";
}
export interface WordStableParagraphTarget {
  paragraphId?: string | number;
  internalId?: string | number;
  paragraphIndex?: number;
  search?: string;
  occurrence?: number;
  matchCase?: boolean;
  matchMode?: "contains" | "exact";
}
export interface WordSetNumberingArgs extends WordParagraphTarget {
  kind?: "bullet" | "numbered" | "multilevel";
  level?: number;
  levels?: WordNumberingLevel[];
  restartAt?: number;
  assignments?: WordNumberingAssignment[];
  continueFrom?: WordStableParagraphTarget;
}
export interface WordBorderFormat {
  style?: "none" | "single" | "double" | "dotted" | "dashed" | "thick" | "wave";
  color?: string;
  /** Border width in points (pt). */
  widthPt?: number;
  /** Space between border and content in points (pt). */
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
  /** @deprecated Compatibility alias for cols. Prefer cols. */
  columns?: number;
  data?: Array<Array<AiBridgeTableCellScalar | WordTableCellInput>>;
  widthPercent?: number;
  /** Whole-table alignment; cell paragraph alignment belongs in data cell objects. */
  align?: "left" | "center" | "right";
  /** @deprecated Compatibility alias for tableStyleName. */
  styleName?: string;
  tableStyleName?: string;
  cellParagraphStyleName?: string;
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
export interface WordAddShapeArgs extends BasicTextFormat, WordStableParagraphTarget {
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
export interface WordAddChartArgs extends WordStableParagraphTarget {
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
export interface WordAddMathArgs extends WordStableParagraphTarget {
  equation: string;
  format?: "unicode" | "latex" | "mathml";
  paragraphIndex?: number;
  current?: boolean;
}
export interface WordAddOleObjectArgs extends WordStableParagraphTarget {
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
  paragraphId?: string | number;
  internalId?: string | number;
}
export interface WordManageLongDocumentArgs extends WordStableParagraphTarget {
  action:
    | "addToc" | "updateToc"
    | "addCaption" | "addTableOfFigures" | "updateTableOfFigures"
    | "addCrossReference" | "addFootnote" | "addEndnote"
    | "deleteBookmark";
  paragraphIndex?: number;
  paragraphId?: string | number;
  internalId?: string | number;
  tableIndex?: number;
  insertAt?: "before" | "after" | "replaceEmpty";
  styleName?: string;
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
  targetParagraphId?: string | number;
  targetInternalId?: string | number;
  captionIndex?: number;
  referenceType?: string;
  hyperlink?: boolean;
  aboveBelow?: boolean;
  numberSeparator?: string;
  replaceSearch?: string;
  occurrence?: number;
  matchCase?: boolean;
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
export interface WordManageContentControlArgs extends WordStableParagraphTarget {
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
  /** Line width in points (pt). */
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
  titleBold?: boolean;
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
  titleBold?: boolean;
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
  /** Include paragraph/run style readback; unsupported getters are reported as unavailable. */
  includeTextStyles?: boolean;
}
export interface SlidesAllowedOverlapPair {
  firstName: string;
  secondName: string;
}
export interface SlidesValidateLayoutArgs {
  slide?: number;
  /** Uniform or per-edge structural safe area in millimetres. */
  safeMarginMm?: number | AiBridgePadding;
  minFontSize?: number;
  maxObjects?: number;
  expectedMasterIndex?: number;
  expectedLayoutIndex?: number;
  allowedOverlapPairs?: SlidesAllowedOverlapPair[];
  requireUniqueNames?: boolean;
}
export interface SlidesReplaceTextArgs { search: string; replace: string; matchCase?: boolean; slide?: number; }
export interface SlidesScaleFontArgs { scale: number; slide?: number; }
export interface SlidesFormatTextArgs extends BasicTextFormat {
  underline?: boolean;
  slide?: number;
}
export interface SlidesFormatSelectionArgs extends BasicTextFormat {
  underline?: boolean;
}
export interface SlidesAddSlideArgs {
  index?: number;
  title?: string;
  titleFontSize?: number;
  backgroundColor?: string;
  /** One-based master index; used only when layoutIndex is provided. */
  masterIndex?: number;
  /** One-based layout index applied atomically to the new slide. */
  layoutIndex?: number;
}
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
export interface SlidesTemplateShapeFormatting extends SlidesShapeFormatting {
  underline?: boolean;
}
export interface SlidesAddTemplateShapeArgs extends SlidesTemplateShapeFormatting {
  scope: "master" | "layout";
  masterIndex?: number;
  layoutIndex?: number;
  shapeType: string;
  placeholderType?: string;
}
export interface SlidesManageTemplateObjectArgs extends SlidesTemplateShapeFormatting {
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
export type SlidesSetTemplateBackgroundArgs = {
  scope: "master" | "layout";
  masterIndex?: number;
  layoutIndex?: number;
} & (
  | { mode?: "custom"; fill: AiBridgeFill; source?: never; fillMode?: never }
  | { mode: "image"; source: AiBridgeImageSource; fillMode?: "stretch" | "tile"; fill?: never }
  | { mode: "clear" | "master"; fill?: never; source?: never; fillMode?: never }
);
export interface SlidesParagraphFormat extends BasicTextFormat {
  underline?: boolean;
  align?: "left" | "center" | "right" | "both";
  /** Paragraph indents in millimetres (mm). */
  firstLineIndentMm?: number;
  leftIndentMm?: number;
  rightIndentMm?: number;
  /** Paragraph spacing in points (pt). */
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
export interface SlidesTableCellStyle extends BasicTextFormat {
  underline?: boolean;
  fill?: AiBridgeFill;
  backgroundColor?: string;
  align?: "left" | "center" | "right" | "both";
  verticalAlign?: "top" | "center" | "bottom";
  border?: SlidesTableBorder;
}
export interface SlidesTableCellFormat extends SlidesTableCellStyle {
  text?: string;
}
export interface SlidesTableCellInput extends SlidesTableCellFormat {
  text: string;
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
  data?: Array<Array<AiBridgeTableCellScalar | SlidesTableCellInput>>;
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
export interface SlidesFormatTableArgs extends SlidesObjectTarget, SlidesTableCellStyle {
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
export type SlidesAddTextBoxArgs =
  Omit<SlidesShapeFormatting, "text"> & { slide: number } & (
    | { text: string; paragraphs?: never }
    | { paragraphs: SlidesTextParagraph[]; text?: never }
  );
export interface SlidesAddWordArtArgs extends BasicTextFormat {
  slide: number;
  text: string;
  transform?: string;
  xMm?: number;
  yMm?: number;
  widthMm?: number;
  heightMm?: number;
  rotationDeg?: number;
  name?: string;
  fill?: AiBridgeFill;
  line?: AiBridgeLine;
  fillColor?: string;
  lineColor?: string;
  lineWidthPt?: number;
  includeRaw?: boolean;
}
export interface SlidesAddMathArgs {
  slide: number;
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
export type SlidesSetBackgroundArgs =
  | { slide: number; mode?: "custom"; fill: AiBridgeFill; source?: never; fillMode?: never }
  | { slide: number; mode: "image"; source: AiBridgeImageSource; fillMode?: "stretch" | "tile"; fill?: never }
  | { slide: number; mode: "clear" | "layout" | "master"; fill?: never; source?: never; fillMode?: never };
export type SlidesAddShapeArgs =
  Omit<SlidesShapeFormatting, "text"> & { slide: number; shapeType: string } & (
    | { text?: string; paragraphs?: never }
    | { paragraphs: SlidesTextParagraph[]; text?: never }
  );
export interface SlidesUpdateShapeArgs extends SlidesObjectTarget, SlidesShapeFormatting {}
export interface SlidesDeleteObjectArgs extends SlidesObjectTarget {}
export type SlidesInspectSmartArtsArgs = AiBridgeGeneratedSlidesInspectSmartartsArgs;
export type SlidesAddSmartArtArgs = AiBridgeGeneratedSlidesAddSmartartArgs;
export type SlidesUpdateSmartArtArgs = AiBridgeGeneratedSlidesUpdateSmartartArgs;
export type SlidesDeleteSmartArtArgs = AiBridgeGeneratedSlidesDeleteSmartartArgs;
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
export interface SheetsInspectRangeArgs extends SheetTarget {
  range: string;
  includeValues?: boolean;
  /** Include cell font, fill, alignment, number format, wrapping, and dimensions in the readback. */
  includeFormat?: boolean;
  /** Include conditional-format count and readable rule properties in the readback. */
  includeConditionalFormats?: boolean;
  /** Include the readable data-validation rule, or null when no rule exists. */
  includeValidation?: boolean;
}
export type SheetCellValue = string | number | boolean | null;
export interface SheetsSetValuesArgs extends SheetTarget {
  range: string;
  /**
   * Scalars are valid only for a single-cell target. Multi-cell targets require
   * a non-empty rectangular matrix with exactly the same dimensions.
   */
  values: SheetCellValue | SheetCellValue[][];
}
export interface SheetsSetFormulaArgs extends SheetTarget {
  range: string;
  /**
   * Formula strings are valid only for a single-cell target. Multi-cell targets
   * require a target-sized matrix, or an anchor followed by fillDown/fillRight.
   */
  formula: string | string[][];
}
export interface SheetsSetArrayFormulaArgs extends SheetTarget { range: string; formula: string; }
export interface SheetsReplaceTextArgs extends SheetTarget { range?: string; search: string; replace: string; }
export interface SheetsFormatRangeArgs extends SheetTarget {
  range: string;
  /** Font size in points (pt). */
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
  /** Text rotation accepted by ONLYOFFICE, commonly degrees or an xl* orientation token. */
  orientation?: string | number;
  /** @deprecated Native character-width alias. Prefer columnWidthChars. */
  columnWidth?: number;
  /** @deprecated Point-valued alias. Prefer rowHeightPt. */
  rowHeight?: number;
  /** Column width in spreadsheet character units (roughly seven pixels per unit). */
  columnWidthChars?: number;
  /** Row height in points (pt). */
  rowHeightPt?: number;
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
  header?: boolean | "yes" | "no" | "guess" | "xlYes" | "xlNo" | "xlGuess";
  orientation?: "rows" | "columns" | "xlSortRows" | "xlSortColumns";
}
export type SheetsFilterOperator =
  | "and" | "or" | "filterValues" | "values" | "top10Items" | "bottom10Items"
  | "top10Percent" | "bottom10Percent" | "filterCellColor" | "filterFontColor"
  | "filterIcon" | "dynamic" | "xlAnd" | "xlOr" | "xlFilterValues"
  | "xlTop10Items" | "xlBottom10Items" | "xlTop10Percent" | "xlBottom10Percent"
  | "xlFilterCellColor" | "xlFilterFontColor" | "xlFilterIcon" | "xlFilterDynamic";
export interface SheetsFilterArgs extends SheetTarget {
  action: "set" | "showAll" | "reapply";
  range?: string;
  field?: number;
  criteria1?: unknown;
  operator?: SheetsFilterOperator;
  criteria2?: unknown;
  visibleDropDown?: boolean;
}
export interface SheetsInspectTablesArgs extends SheetTarget {
  tableIndex?: number;
  name?: string;
  maxTables?: number;
}
export interface SheetsManageTableArgs extends SheetTarget {
  /** "format" is deprecated and only accepts sheet plus range. */
  action: "create" | "format" | "update" | "resize" | "delete" | "unlist";
  range?: string;
  /**
   * Applies only to create and defaults to auto. structured requires
   * ApiListObject; rangeStyle only formats the range; basic is a compatibility
   * alias for rangeStyle; auto reports degradation explicitly.
   */
  tableMode?: "auto" | "structured" | "rangeStyle" | "basic";
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
  operator?: SheetsComparisonOperator;
  formula1?: unknown;
  formula2?: unknown;
  scale?: 2 | 3;
  fillColor?: string;
  fontColor?: string;
  bold?: boolean;
}
export type SheetsComparisonOperator =
  | "between" | "notBetween" | "equal" | "notEqual" | "greaterThan" | "lessThan"
  | "greaterThanOrEqual" | "lessThanOrEqual" | "xlBetween" | "xlNotBetween"
  | "xlEqual" | "xlNotEqual" | "xlGreater" | "xlLess" | "xlGreaterEqual" | "xlLessEqual";
export type SheetsValidationType =
  | "inputOnly" | "wholeNumber" | "decimal" | "list" | "date" | "time" | "textLength"
  | "custom" | "xlValidateInputOnly" | "xlValidateWholeNumber" | "xlValidateDecimal"
  | "xlValidateList" | "xlValidateDate" | "xlValidateTime" | "xlValidateTextLength"
  | "xlValidateCustom";
export type SheetsValidationAlertStyle =
  | "stop" | "warning" | "information" | "info"
  | "xlValidAlertStop" | "xlValidAlertWarning" | "xlValidAlertInformation";
export interface SheetsManageValidationArgs extends SheetTarget {
  action: "add" | "modify" | "delete";
  range: string;
  type?: SheetsValidationType;
  alertStyle?: SheetsValidationAlertStyle;
  operator?: SheetsComparisonOperator;
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
export type SheetsManageFreezePanesArgs = SheetTarget & (
  | { action: "unfreeze"; range?: never; count?: never }
  /** The range is the exact leading worksheet area to freeze, not the first scrollable cell. */
  | { action: "freezeAt"; range: string; count?: never }
  /** count must be a positive integer; use unfreeze to remove panes. */
  | { action: "freezeRows"; count: number; range?: never }
  /** count must be a positive integer; use unfreeze to remove panes. */
  | { action: "freezeColumns"; count: number; range?: never }
);
export interface SheetsFreezePanesLocation {
  address: string;
  rows: number | null;
  columns: number | null;
}
/** Readback returned by freeze-pane inspect/manage calls after view verification. */
export interface SheetsFreezePanesResult {
  name: "sheets_inspect_freeze_panes" | "sheets_manage_freeze_panes";
  action?: "unfreeze" | "freezeAt" | "freezeRows" | "freezeColumns";
  sheet: string;
  location: SheetsFreezePanesLocation | null;
  frozenRows: number;
  frozenColumns: number;
  topLeftCell: string | null;
  verified: boolean;
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
/** Page-layout readback; display fields are normalized worksheet-screen booleans. */
export interface SheetsPageLayoutResult {
  name: "sheets_inspect_page_layout" | "sheets_manage_page_layout";
  sheet: string;
  orientation: string | null;
  marginsPt: { top: number | null; right: number | null; bottom: number | null; left: number | null };
  margins: { top: number | null; right: number | null; bottom: number | null; left: number | null };
  printGridlines: boolean | null;
  printHeadings: boolean | null;
  displayGridlines: boolean;
  displayHeadings: boolean;
  verified: boolean;
}
export type SheetsManagePageLayoutArgs = SheetTarget & {
  orientation?: "portrait" | "landscape" | "xlPortrait" | "xlLandscape";
  /** @deprecated Point-valued alias. Prefer marginsPt. */
  margins?: { top?: number; right?: number; bottom?: number; left?: number };
  /** Printed page margins in points (pt). */
  marginsPt?: { top?: number; right?: number; bottom?: number; left?: number };
  printGridlines?: boolean;
  printHeadings?: boolean;
  /** Controls and verifies gridlines in the worksheet screen view, not printed output. */
  displayGridlines?: boolean;
  /** Controls and verifies row/column headings in the worksheet screen view, not printed output. */
  displayHeadings?: boolean;
} & (
  | { orientation: "portrait" | "landscape" | "xlPortrait" | "xlLandscape" }
  | { margins: { top?: number; right?: number; bottom?: number; left?: number } }
  | { marginsPt: { top?: number; right?: number; bottom?: number; left?: number } }
  | { printGridlines: boolean }
  | { printHeadings: boolean }
  | { displayGridlines: boolean }
  | { displayHeadings: boolean }
);
export interface SheetsInspectMacrosArgs { kind?: "onlyoffice" | "vba"; }
export interface SheetsSetMacrosArgs { content: Record<string, unknown>; }

// <ai-bridge-generated:tool-arguments>
// Generated by tools/sync_contract.py. Do not edit this region.
export type AiBridgeGeneratedWordInspectArgs = {
  maxChars?: number;
  maxParagraphs?: number;
  includeStructure?: boolean;
  includeComments?: boolean;
};
export type AiBridgeGeneratedWordReplaceTextArgs = {
  search: string;
  replace: string;
  matchCase?: boolean;
};
export type AiBridgeGeneratedWordAppendParagraphArgs = {
  text: string;
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikeout?: boolean;
  doubleStrikeout?: boolean;
  caps?: boolean;
  smallCaps?: boolean;
  color?: string;
  highlightColor?: string;
  characterSpacing?: number;
  characterSpacingPt?: number;
  vertAlign?: "baseline" | "subscript" | "superscript";
  characterStyleName?: string;
  align?: "left" | "center" | "right" | "both";
  spacingBefore?: number;
  spacingAfter?: number;
  spacingBeforePt?: number;
  spacingAfterPt?: number;
  lineSpacing?: number;
  lineRule?: "auto" | "exact" | "atLeast";
  firstLineIndent?: number;
  leftIndent?: number;
  rightIndent?: number;
  firstLineIndentPt?: number;
  leftIndentPt?: number;
  rightIndentPt?: number;
  styleName?: string;
  headingLevel?: number;
  outlineLevel?: number;
  keepLines?: boolean;
  keepNext?: boolean;
  widowControl?: boolean;
  contextualSpacing?: boolean;
  pageBreakBefore?: boolean;
  listType?: "bullet" | "numbered";
  listLevel?: number;
};
export type AiBridgeGeneratedWordInsertParagraphArgs = {
  text: string;
  inline?: boolean;
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikeout?: boolean;
  doubleStrikeout?: boolean;
  caps?: boolean;
  smallCaps?: boolean;
  color?: string;
  highlightColor?: string;
  characterSpacing?: number;
  characterSpacingPt?: number;
  vertAlign?: "baseline" | "subscript" | "superscript";
  characterStyleName?: string;
  align?: "left" | "center" | "right" | "both";
  spacingBefore?: number;
  spacingAfter?: number;
  spacingBeforePt?: number;
  spacingAfterPt?: number;
  lineSpacing?: number;
  lineRule?: "auto" | "exact" | "atLeast";
  firstLineIndent?: number;
  leftIndent?: number;
  rightIndent?: number;
  firstLineIndentPt?: number;
  leftIndentPt?: number;
  rightIndentPt?: number;
  styleName?: string;
  headingLevel?: number;
  outlineLevel?: number;
  keepLines?: boolean;
  keepNext?: boolean;
  widowControl?: boolean;
  contextualSpacing?: boolean;
  pageBreakBefore?: boolean;
  listType?: "bullet" | "numbered";
  listLevel?: number;
};
export type AiBridgeGeneratedWordFormatDocumentArgs = {
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikeout?: boolean;
  doubleStrikeout?: boolean;
  caps?: boolean;
  smallCaps?: boolean;
  color?: string;
  highlightColor?: string;
  characterSpacing?: number;
  characterSpacingPt?: number;
  vertAlign?: "baseline" | "subscript" | "superscript";
  characterStyleName?: string;
  align?: "left" | "center" | "right" | "both";
  styleName?: string;
  headingLevel?: number;
  outlineLevel?: number;
  spacingBefore?: number;
  spacingAfter?: number;
  spacingBeforePt?: number;
  spacingAfterPt?: number;
  lineSpacing?: number;
  lineRule?: "auto" | "exact" | "atLeast";
  firstLineIndent?: number;
  leftIndent?: number;
  rightIndent?: number;
  firstLineIndentPt?: number;
  leftIndentPt?: number;
  rightIndentPt?: number;
  keepLines?: boolean;
  keepNext?: boolean;
  widowControl?: boolean;
  contextualSpacing?: boolean;
  pageBreakBefore?: boolean;
};
export type AiBridgeGeneratedWordFormatSelectionArgs = {
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikeout?: boolean;
  doubleStrikeout?: boolean;
  caps?: boolean;
  smallCaps?: boolean;
  color?: string;
  highlightColor?: string;
  characterSpacing?: number;
  characterSpacingPt?: number;
  vertAlign?: "baseline" | "subscript" | "superscript";
  characterStyleName?: string;
};
export type AiBridgeGeneratedWordFormatMatchesArgs = {
  search: string;
  matchCase?: boolean;
  occurrence?: number;
  maxMatches?: number;
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikeout?: boolean;
  doubleStrikeout?: boolean;
  caps?: boolean;
  smallCaps?: boolean;
  color?: string;
  highlightColor?: string;
  characterSpacing?: number;
  characterSpacingPt?: number;
  vertAlign?: "baseline" | "subscript" | "superscript";
  characterStyleName?: string;
};
export type AiBridgeGeneratedWordDeleteMatchesArgs = {
  search: string;
  matchCase?: boolean;
  occurrence?: number;
  maxMatches?: number;
};
export type AiBridgeGeneratedWordAddHyperlinkArgs = {
  search: string;
  url: string;
  screenTip?: string;
  bookmarkName?: string;
  matchCase?: boolean;
  occurrence?: number;
  maxMatches?: number;
};
export type AiBridgeGeneratedWordAddCommentArgs = {
  search: string;
  text: string;
  author?: string;
  userId?: string;
  matchCase?: boolean;
  occurrence?: number;
  maxMatches?: number;
};
export type AiBridgeGeneratedWordAddBookmarkArgs = {
  search: string;
  occurrence: number;
  name: string;
  matchCase?: boolean;
};
export type AiBridgeGeneratedWordAddImageArgs = {
  source: ({
    type: "url";
    url: string;
  } | {
    type: "dataUrl";
    dataUrl: string;
  } | {
    type: "relayAsset";
    path: string;
    assetToken: string;
    mimeType: "image/png" | "image/jpeg" | "image/gif" | "image/webp" | "image/svg+xml";
    widthPx: number;
    heightPx: number;
    assetId: string;
  });
  widthMm?: number;
  heightMm?: number;
  preserveAspectRatio?: boolean;
  current?: boolean;
  paragraphIndex?: number;
  paragraphId?: string | number;
  internalId?: string | number;
  search?: string;
  occurrence?: number;
  matchCase?: boolean;
  matchMode?: "contains" | "exact";
  wrapping?: "inline" | "square" | "tight" | "through" | "topAndBottom" | "behind" | "inFront";
  name?: string;
};
export type AiBridgeGeneratedWordInspectAdvancedArgs = {
  includeProperties?: boolean;
  customPropertyNames?: Array<string>;
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
};
export type AiBridgeGeneratedWordSetDocumentPropertiesArgs = {
  title?: string;
  subject?: string;
  creator?: string;
  author?: string;
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
};
export type AiBridgeGeneratedWordManageSectionArgs = {
  action?: "configure" | "create";
  sectionIndex?: number;
  paragraphIndex?: number;
  endParagraphIndex?: number;
  boundary?: {
    position: "before" | "after";
    paragraphId?: string | number;
    internalId?: string | number;
    paragraphIndex?: number;
    search?: string;
    occurrence?: number;
    matchCase?: boolean;
    matchMode?: "contains" | "exact";
  };
  type?: "continuous" | "nextPage" | "evenPage" | "oddPage";
  startPageNumber?: number;
  titlePage?: boolean;
  differentFirstPage?: boolean;
  evenAndOddHeaders?: boolean;
  columns?: {
    count?: number;
    spaceMm?: number;
    entries?: Array<{
      widthMm: number;
      spaceMm?: number;
    }>;
  };
} & ({ action: "create"; } | { action?: Exclude<"configure" | "create", "create">; });
export type AiBridgeGeneratedWordManageStyleArgs = {
  action?: "create" | "update";
  name: string;
  type?: "paragraph" | "character" | "table" | "numbering";
  basedOn?: string;
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikeout?: boolean;
  color?: string;
  highlightColor?: string;
  characterSpacing?: number;
  characterSpacingPt?: number;
  align?: "left" | "center" | "right" | "both";
  spacingBefore?: number;
  spacingAfter?: number;
  spacingBeforePt?: number;
  spacingAfterPt?: number;
  lineSpacing?: number;
  lineRule?: "auto" | "exact" | "atLeast";
  firstLineIndent?: number;
  leftIndent?: number;
  rightIndent?: number;
  firstLineIndentPt?: number;
  leftIndentPt?: number;
  rightIndentPt?: number;
  keepLines?: boolean;
  keepNext?: boolean;
  pageBreakBefore?: boolean;
};
export type AiBridgeGeneratedWordSetTabsArgs = {
  paragraphIndex?: number;
  paragraphId?: string | number;
  internalId?: string | number;
  paragraphIndexes?: Array<number>;
  search?: string;
  matchCase?: boolean;
  matchMode?: "contains" | "exact";
  occurrence?: number;
  maxParagraphs?: number;
  all?: boolean;
  current?: boolean;
  clearAll?: boolean;
  tabs: Array<{
    positionMm: number;
    align?: "left" | "center" | "right" | "decimal" | "bar" | "clear";
  }>;
};
export type AiBridgeGeneratedWordSetNumberingArgs = {
  kind?: "bullet" | "numbered" | "multilevel";
  level?: number;
  restartAt?: number;
  levels?: Array<{
    level?: number;
    format?: "none" | "bullet" | "decimal" | "lowerRoman" | "upperRoman" | "lowerLetter" | "upperLetter" | "decimalZero";
    numberFormat?: "none" | "bullet" | "decimal" | "lowerRoman" | "upperRoman" | "lowerLetter" | "upperLetter" | "decimalZero";
    text?: string;
    start?: number;
    align?: "left" | "center" | "right";
    restart?: (boolean | 0 | 1);
    leftIndentPt?: number;
    firstLineIndentPt?: number;
    hangingIndentPt?: number;
  }>;
  assignments?: Array<{
    level: number;
    paragraphId?: string | number;
    internalId?: string | number;
    paragraphIndex?: number;
    search?: string;
    occurrence?: number;
    matchCase?: boolean;
    matchMode?: "contains" | "exact";
  }>;
  continueFrom?: {
    paragraphId?: string | number;
    internalId?: string | number;
    paragraphIndex?: number;
    search?: string;
    occurrence?: number;
    matchCase?: boolean;
    matchMode?: "contains" | "exact";
  };
  paragraphIndex?: number;
  paragraphId?: string | number;
  internalId?: string | number;
  paragraphIndexes?: Array<number>;
  search?: string;
  matchCase?: boolean;
  matchMode?: "contains" | "exact";
  occurrence?: number;
  maxParagraphs?: number;
  all?: boolean;
  current?: boolean;
} & ({ assignments: Array<{
  level: number;
  paragraphId?: string | number;
  internalId?: string | number;
  paragraphIndex?: number;
  search?: string;
  occurrence?: number;
  matchCase?: boolean;
  matchMode?: "contains" | "exact";
}>; } | { paragraphIndex: number; } | { paragraphId: string | number; } | { internalId: string | number; } | { paragraphIndexes: Array<number>; } | { search: string; } | { all: boolean; } | { current: boolean; });
export type AiBridgeGeneratedWordFormatTableAdvancedArgs = {
  tableIndex: number;
  row?: number;
  column?: number;
  rowHeightMm?: number;
  rowHeightRule?: "auto" | "atLeast";
  columnWidthMm?: number;
  repeatHeader?: boolean;
  cellMarginsMm?: {
    top?: number;
    right?: number;
    bottom?: number;
    left?: number;
  };
  borders?: {
    top?: {
      style?: "none" | "single" | "double" | "dotted" | "dashed" | "thick" | "wave";
      color?: string;
      widthPt?: number;
      spacePt?: number;
    };
    right?: {
      style?: "none" | "single" | "double" | "dotted" | "dashed" | "thick" | "wave";
      color?: string;
      widthPt?: number;
      spacePt?: number;
    };
    bottom?: {
      style?: "none" | "single" | "double" | "dotted" | "dashed" | "thick" | "wave";
      color?: string;
      widthPt?: number;
      spacePt?: number;
    };
    left?: {
      style?: "none" | "single" | "double" | "dotted" | "dashed" | "thick" | "wave";
      color?: string;
      widthPt?: number;
      spacePt?: number;
    };
    insideHorizontal?: {
      style?: "none" | "single" | "double" | "dotted" | "dashed" | "thick" | "wave";
      color?: string;
      widthPt?: number;
      spacePt?: number;
    };
    insideVertical?: {
      style?: "none" | "single" | "double" | "dotted" | "dashed" | "thick" | "wave";
      color?: string;
      widthPt?: number;
      spacePt?: number;
    };
  };
  wrapping?: "none" | "around";
} & ({ repeatHeader: boolean; row: number; } | { repeatHeader?: never; });
export type AiBridgeGeneratedWordAddNestedTableArgs = {
  tableIndex: number;
  row: number;
  column: number;
  rows: number;
  cols: number;
  columns?: number;
  data?: Array<Array<(string | number | boolean | null | {
    text: string;
    paragraphStyleName?: string;
    styleName?: string;
    characterStyleName?: string;
    fontSize?: number;
    fontFamily?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    color?: string;
    textColor?: string;
    highlightColor?: string;
    characterSpacing?: number;
    characterSpacingPt?: number;
    backgroundColor?: string;
    align?: "left" | "center" | "right" | "both";
    spacingBefore?: number;
    spacingAfter?: number;
    spacingBeforePt?: number;
    spacingAfterPt?: number;
    lineSpacing?: number;
    lineRule?: "auto" | "exact" | "atLeast";
    firstLineIndent?: number;
    leftIndent?: number;
    rightIndent?: number;
    firstLineIndentPt?: number;
    leftIndentPt?: number;
    rightIndentPt?: number;
    verticalAlign?: "top" | "center" | "bottom";
    widthPercent?: number;
  })>>;
  widthPercent?: number;
  align?: "left" | "center" | "right";
  styleName?: string;
  tableStyleName?: string;
  cellParagraphStyleName?: string;
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;
};
export type AiBridgeGeneratedWordManageDrawingArgs = {
  action: "update" | "delete";
  drawingIndex?: number;
  name?: string;
  kind?: "image" | "shape" | "chart" | "oleObject" | "smartArt" | "any";
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
  border?: {
    style?: "none" | "single" | "double" | "dotted" | "dashed" | "thick" | "wave";
    color?: string;
    widthPt?: number;
    spacePt?: number;
  };
} & ({ drawingIndex: number; } | { name: string; });
export type AiBridgeGeneratedWordAddShapeArgs = {
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
  paragraphId?: string | number;
  internalId?: string | number;
  search?: string;
  matchCase?: boolean;
  matchMode?: "contains" | "exact";
  occurrence?: number;
  current?: boolean;
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;
};
export type AiBridgeGeneratedWordAddChartArgs = {
  chartType?: "bar" | "barStacked" | "barStackedPercent" | "bar3D" | "barStacked3D" | "barStackedPercent3D" | "barStackedPercent3DPerspective" | "horizontalBar" | "horizontalBarStacked" | "horizontalBarStackedPercent" | "horizontalBar3D" | "horizontalBarStacked3D" | "horizontalBarStackedPercent3D" | "lineNormal" | "lineStacked" | "lineStackedPercent" | "lineNormalMarker" | "lineStackedMarker" | "lineStackedPerMarker" | "line3D" | "pie" | "pie3D" | "doughnut" | "scatter" | "scatterLine" | "scatterLineMarker" | "scatterSmooth" | "scatterSmoothMarker" | "stock" | "area" | "areaStacked" | "areaStackedPercent" | "comboCustom" | "comboBarLine" | "comboBarLineSecondary" | "radar" | "radarMarker" | "radarFilled" | "line" | "lineMarker" | "stackedBar" | "stackedBarPercent" | "stackedLine" | "stackedLinePercent" | "column";
  data: Array<Array<number>>;
  seriesNames?: Array<string>;
  categories?: Array<string | number>;
  numberFormats?: Array<string>;
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
  paragraphId?: string | number;
  internalId?: string | number;
  search?: string;
  matchCase?: boolean;
  matchMode?: "contains" | "exact";
  occurrence?: number;
  current?: boolean;
};
export type AiBridgeGeneratedWordAddMathArgs = {
  equation: string;
  format?: "unicode" | "latex" | "mathml";
  paragraphIndex?: number;
  paragraphId?: string | number;
  internalId?: string | number;
  search?: string;
  matchCase?: boolean;
  matchMode?: "contains" | "exact";
  occurrence?: number;
  current?: boolean;
};
export type AiBridgeGeneratedWordAddOleObjectArgs = {
  source: ({
    type: "url";
    url: string;
  } | {
    type: "dataUrl";
    dataUrl: string;
  } | {
    type: "relayAsset";
    path: string;
    assetToken: string;
    mimeType: "image/png" | "image/jpeg" | "image/gif" | "image/webp" | "image/svg+xml";
    widthPx: number;
    heightPx: number;
    assetId: string;
  });
  data: string;
  applicationId: string;
  widthMm?: number;
  heightMm?: number;
  paragraphIndex?: number;
  paragraphId?: string | number;
  internalId?: string | number;
  search?: string;
  matchCase?: boolean;
  matchMode?: "contains" | "exact";
  occurrence?: number;
  current?: boolean;
  name?: string;
};
export type AiBridgeGeneratedWordManageFieldsArgs = {
  action: "add" | "updateAll" | "clearForms";
  instruction?: string;
  search?: string;
  occurrence?: number;
  matchCase?: boolean;
  paragraphIndex?: number;
  paragraphId?: string | number;
  internalId?: string | number;
} & ({ action: "add"; instruction: string; } | { action?: Exclude<"add" | "updateAll" | "clearForms", "add">; });
export type AiBridgeGeneratedWordManageLongDocumentArgs = {
  action: "addToc" | "updateToc" | "addCaption" | "addTableOfFigures" | "updateTableOfFigures" | "addCrossReference" | "addFootnote" | "addEndnote" | "deleteBookmark";
  paragraphIndex?: number;
  paragraphId?: string | number;
  internalId?: string | number;
  search?: string;
  matchMode?: "contains" | "exact";
  tableIndex?: number;
  insertAt?: "before" | "after" | "replaceEmpty";
  styleName?: string;
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
  targetParagraphId?: string | number;
  targetInternalId?: string | number;
  captionIndex?: number;
  referenceType?: string;
  hyperlink?: boolean;
  aboveBelow?: boolean;
  numberSeparator?: string;
  replaceSearch?: string;
  occurrence?: number;
  matchCase?: boolean;
  noteText?: string;
  bookmarkName?: string;
};
export type AiBridgeGeneratedWordManageCommentsArgs = {
  action: "reply" | "edit" | "remove" | "removeAll" | "resolve" | "reopen";
  commentId?: string;
  text?: string;
  author?: string;
  userId?: string;
};
export type AiBridgeGeneratedWordManageRevisionsArgs = {
  action: "start" | "stop" | "acceptAll" | "rejectAll" | "setDisplay";
  displayMode?: "edit" | "simple" | "final" | "original";
} & ({ action: "setDisplay"; displayMode: "edit" | "simple" | "final" | "original"; } | { action?: Exclude<"start" | "stop" | "acceptAll" | "rejectAll" | "setDisplay", "setDisplay">; });
export type AiBridgeGeneratedWordSetProtectionArgs = {
  action: "protect" | "unprotect";
  mode?: "readOnly" | "comments" | "forms";
};
export type AiBridgeGeneratedWordManageContentControlArgs = {
  action: "add" | "update" | "remove" | "clear" | "check";
  kind?: "block" | "inline" | "checkbox" | "comboBox" | "dropDown" | "datePicker" | "picture";
  index?: number;
  tag?: string;
  title?: string;
  text?: string;
  placeholder?: string;
  items?: Array<{
    display: string;
    value?: string;
  }>;
  selectedValue?: string;
  checked?: boolean;
  lock?: "none" | "content" | "control" | "both";
  color?: string;
  appearance?: "boundingBox" | "hidden";
  date?: string;
  dateFormat?: string;
  dataBinding?: {
    prefixMapping?: string;
    storeItemId: string;
    xpath: string;
  };
  updateFromXml?: boolean;
  paragraphIndex?: number;
  paragraphId?: string | number;
  internalId?: string | number;
  search?: string;
  matchCase?: boolean;
  matchMode?: "contains" | "exact";
  occurrence?: number;
  current?: boolean;
  tableIndex?: number;
  row?: number;
  column?: number;
  contentMode?: "append" | "replace";
};
export type AiBridgeGeneratedWordManageCustomXmlArgs = {
  action: "add" | "replace" | "remove" | "insertElement" | "updateElement" | "deleteElement" | "insertAttribute" | "updateAttribute" | "deleteAttribute";
  partId?: string;
  xml?: string;
  xpath?: string;
  name?: string;
  value?: string;
};
export type AiBridgeGeneratedWordInspectMacrosArgs = {
  kind?: "office" | "vba";
};
export type AiBridgeGeneratedWordSetMacrosArgs = {
  macros: Array<{
    name: string;
    value: string;
    guid?: string;
  }>;
  current?: number;
};
export type AiBridgeGeneratedWordSetWatermarkArgs = {
  action?: "setText" | "setImage" | "remove";
  text?: string;
  source?: ({
    type: "url";
    url: string;
  } | {
    type: "dataUrl";
    dataUrl: string;
  } | {
    type: "relayAsset";
    path: string;
    assetToken: string;
    mimeType: "image/png" | "image/jpeg" | "image/gif" | "image/webp" | "image/svg+xml";
    widthPx: number;
    heightPx: number;
    assetId: string;
  });
  opacity?: number;
  diagonal?: boolean;
  scale?: number;
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;
};
export type AiBridgeGeneratedWordFormatParagraphsArgs = {
  paragraphIndex?: number;
  paragraphId?: string | number;
  internalId?: string | number;
  paragraphIndexes?: Array<number>;
  search?: string;
  matchCase?: boolean;
  matchMode?: "contains" | "exact";
  occurrence?: number;
  maxParagraphs?: number;
  all?: boolean;
  current?: boolean;
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikeout?: boolean;
  doubleStrikeout?: boolean;
  caps?: boolean;
  smallCaps?: boolean;
  color?: string;
  highlightColor?: string;
  characterSpacing?: number;
  characterSpacingPt?: number;
  vertAlign?: "baseline" | "subscript" | "superscript";
  characterStyleName?: string;
  align?: "left" | "center" | "right" | "both";
  styleName?: string;
  headingLevel?: number;
  outlineLevel?: number;
  spacingBefore?: number;
  spacingAfter?: number;
  spacingBeforePt?: number;
  spacingAfterPt?: number;
  lineSpacing?: number;
  lineRule?: "auto" | "exact" | "atLeast";
  firstLineIndent?: number;
  leftIndent?: number;
  rightIndent?: number;
  firstLineIndentPt?: number;
  leftIndentPt?: number;
  rightIndentPt?: number;
  keepLines?: boolean;
  keepNext?: boolean;
  widowControl?: boolean;
  contextualSpacing?: boolean;
  pageBreakBefore?: boolean;
};
export type AiBridgeGeneratedWordSetParagraphTextArgs = {
  text: string;
  paragraphIndex?: number;
  paragraphId?: string | number;
  internalId?: string | number;
  paragraphIndexes?: Array<number>;
  search?: string;
  matchCase?: boolean;
  matchMode?: "contains" | "exact";
  occurrence?: number;
  maxParagraphs?: number;
  all?: boolean;
  current?: boolean;
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikeout?: boolean;
  doubleStrikeout?: boolean;
  caps?: boolean;
  smallCaps?: boolean;
  color?: string;
  highlightColor?: string;
  characterSpacing?: number;
  characterSpacingPt?: number;
  vertAlign?: "baseline" | "subscript" | "superscript";
  characterStyleName?: string;
  align?: "left" | "center" | "right" | "both";
  styleName?: string;
  headingLevel?: number;
  outlineLevel?: number;
  spacingBefore?: number;
  spacingAfter?: number;
  spacingBeforePt?: number;
  spacingAfterPt?: number;
  lineSpacing?: number;
  lineRule?: "auto" | "exact" | "atLeast";
  firstLineIndent?: number;
  leftIndent?: number;
  rightIndent?: number;
  firstLineIndentPt?: number;
  leftIndentPt?: number;
  rightIndentPt?: number;
  keepLines?: boolean;
  keepNext?: boolean;
  widowControl?: boolean;
  contextualSpacing?: boolean;
  pageBreakBefore?: boolean;
};
export type AiBridgeGeneratedWordDeleteParagraphsArgs = {
  paragraphIndex?: number;
  paragraphId?: string | number;
  internalId?: string | number;
  paragraphIndexes?: Array<number>;
  search?: string;
  matchCase?: boolean;
  matchMode?: "contains" | "exact";
  occurrence?: number;
  maxParagraphs?: number;
  all?: boolean;
  current?: boolean;
};
export type AiBridgeGeneratedWordSetListArgs = {
  listType: "bullet" | "numbered";
  level?: number;
  paragraphIndex?: number;
  paragraphId?: string | number;
  internalId?: string | number;
  paragraphIndexes?: Array<number>;
  search?: string;
  matchCase?: boolean;
  matchMode?: "contains" | "exact";
  occurrence?: number;
  maxParagraphs?: number;
  all?: boolean;
  current?: boolean;
  contextualSpacing?: boolean;
};
export type AiBridgeGeneratedWordInsertPageBreakArgs = {
  position?: "before" | "after";
  paragraphIndex?: number;
  paragraphId?: string | number;
  internalId?: string | number;
  paragraphIndexes?: Array<number>;
  search?: string;
  matchCase?: boolean;
  matchMode?: "contains" | "exact";
  occurrence?: number;
  maxParagraphs?: number;
  all?: boolean;
  current?: boolean;
};
export type AiBridgeGeneratedWordNavigateArgs = {
  target?: "start" | "end" | "page" | "next" | "previous" | "relative" | "current" | "search";
  page?: number;
  pageDelta?: number;
  search?: string;
  matchCase?: boolean;
  occurrence?: number;
};
export type AiBridgeGeneratedWordScrollArgs = {
  direction?: "up" | "down";
  pages?: number;
};
export type AiBridgeGeneratedWordScaleFontArgs = {
  scale: number;
};
export type AiBridgeGeneratedWordAddTableArgs = {
  rows: number;
  cols: number;
  columns?: number;
  data?: Array<Array<(string | number | boolean | null | {
    text: string;
    paragraphStyleName?: string;
    styleName?: string;
    characterStyleName?: string;
    fontSize?: number;
    fontFamily?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    color?: string;
    textColor?: string;
    highlightColor?: string;
    characterSpacing?: number;
    characterSpacingPt?: number;
    backgroundColor?: string;
    align?: "left" | "center" | "right" | "both";
    spacingBefore?: number;
    spacingAfter?: number;
    spacingBeforePt?: number;
    spacingAfterPt?: number;
    lineSpacing?: number;
    lineRule?: "auto" | "exact" | "atLeast";
    firstLineIndent?: number;
    leftIndent?: number;
    rightIndent?: number;
    firstLineIndentPt?: number;
    leftIndentPt?: number;
    rightIndentPt?: number;
    verticalAlign?: "top" | "center" | "bottom";
    widthPercent?: number;
  })>>;
  widthPercent?: number;
  align?: "left" | "center" | "right";
  styleName?: string;
  tableStyleName?: string;
  cellParagraphStyleName?: string;
  firstRow?: boolean;
  lastRow?: boolean;
  firstColumn?: boolean;
  lastColumn?: boolean;
  horizontalBanding?: boolean;
  verticalBanding?: boolean;
  title?: string;
  description?: string;
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  insertAt?: "end" | "current" | "before" | "after";
  paragraphIndex?: number;
  paragraphId?: string | number;
  internalId?: string | number;
  tableIndex?: number;
  search?: string;
  matchCase?: boolean;
  matchMode?: "contains" | "exact";
  occurrence?: number;
  pageBreakBefore?: boolean;
};
export type AiBridgeGeneratedWordSetTableCellArgs = {
  tableIndex: number;
  row: number;
  column: number;
  text?: string;
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikeout?: boolean;
  doubleStrikeout?: boolean;
  caps?: boolean;
  smallCaps?: boolean;
  color?: string;
  highlightColor?: string;
  characterSpacing?: number;
  characterSpacingPt?: number;
  vertAlign?: "baseline" | "subscript" | "superscript";
  characterStyleName?: string;
  backgroundColor?: string;
  align?: "left" | "center" | "right" | "both";
  styleName?: string;
  paragraphStyleName?: string;
  headingLevel?: number;
  outlineLevel?: number;
  spacingBefore?: number;
  spacingAfter?: number;
  spacingBeforePt?: number;
  spacingAfterPt?: number;
  lineSpacing?: number;
  lineRule?: "auto" | "exact" | "atLeast";
  firstLineIndent?: number;
  leftIndent?: number;
  rightIndent?: number;
  firstLineIndentPt?: number;
  leftIndentPt?: number;
  rightIndentPt?: number;
  keepLines?: boolean;
  keepNext?: boolean;
  widowControl?: boolean;
  contextualSpacing?: boolean;
  pageBreakBefore?: boolean;
  verticalAlign?: "top" | "center" | "bottom";
  widthPercent?: number;
};
export type AiBridgeGeneratedWordFormatTableArgs = {
  tableIndex: number;
  widthPercent?: number;
  align?: "left" | "center" | "right";
  styleName?: string;
  tableStyleName?: string;
  cellParagraphStyleName?: string;
  backgroundColor?: string;
  title?: string;
  description?: string;
  firstRow?: boolean;
  lastRow?: boolean;
  firstColumn?: boolean;
  lastColumn?: boolean;
  horizontalBanding?: boolean;
  verticalBanding?: boolean;
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  verticalAlign?: "top" | "center" | "bottom";
};
export type AiBridgeGeneratedWordEditTableArgs = {
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
} & ({ action: "removeRow"; row: number; } | { action?: Exclude<"addRow" | "addColumn" | "removeRow" | "removeColumn" | "mergeCells" | "splitCell" | "clear" | "delete", "removeRow">; }) & ({ action: "removeColumn"; column: number; } | { action?: Exclude<"addRow" | "addColumn" | "removeRow" | "removeColumn" | "mergeCells" | "splitCell" | "clear" | "delete", "removeColumn">; }) & ({ action: "splitCell"; row: number; column: number; } | { action?: Exclude<"addRow" | "addColumn" | "removeRow" | "removeColumn" | "mergeCells" | "splitCell" | "clear" | "delete", "splitCell">; }) & ({ action: "mergeCells"; rowStart: number; rowEnd: number; columnStart: number; columnEnd: number; } | { action?: Exclude<"addRow" | "addColumn" | "removeRow" | "removeColumn" | "mergeCells" | "splitCell" | "clear" | "delete", "mergeCells">; });
export type AiBridgeGeneratedWordSetPageLayoutArgs = {
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
  differentFirstPage?: boolean;
};
export type AiBridgeGeneratedWordSetHeaderFooterArgs = {
  kind: "header" | "footer";
  type?: "default" | "first" | "even";
  action?: "set" | "remove";
  sectionIndex?: number;
  text?: string;
  replace?: boolean;
  pageNumber?: boolean;
  pagesCount?: boolean;
  fields?: Array<string>;
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  align?: "left" | "center" | "right" | "both";
};
export type AiBridgeGeneratedWordSetDocumentTextArgs = {
  text: string;
};
export type AiBridgeGeneratedSlidesInspectArgs = {
  maxChars?: number;
};
export type AiBridgeGeneratedSlidesInspectLayoutsArgs = {
  masterIndex?: number;
  includeObjects?: boolean;
  includeRaw?: boolean;
};
export type AiBridgeGeneratedSlidesInspectBackgroundsArgs = {
  slide?: number;
  includeTemplates?: boolean;
  includeRaw?: boolean;
};
export type AiBridgeGeneratedSlidesInspectThemesArgs = {
  masterIndex?: number;
  slide?: number;
  includeRaw?: boolean;
};
export type AiBridgeGeneratedSlidesInspectBuiltinThemesArgs = Record<string, never>;
export type AiBridgeGeneratedSlidesInspectObjectsArgs = {
  slide?: number;
  kinds?: Array<string>;
  maxObjects?: number;
  includeRaw?: boolean;
  includeSlideRaw?: boolean;
  includeTextStyles?: boolean;
};
export type AiBridgeGeneratedSlidesValidateLayoutArgs = {
  slide?: number;
  safeMarginMm?: (number | {
    left?: number;
    top?: number;
    right?: number;
    bottom?: number;
  });
  minFontSize?: number;
  maxObjects?: number;
  expectedMasterIndex?: number;
  expectedLayoutIndex?: number;
  allowedOverlapPairs?: Array<{
    firstName: string;
    secondName: string;
  }>;
  requireUniqueNames?: boolean;
};
export type AiBridgeGeneratedSlidesReplaceTextArgs = {
  search: string;
  replace: string;
  matchCase?: boolean;
  slide?: number;
};
export type AiBridgeGeneratedSlidesScaleFontArgs = {
  scale: number;
  slide?: number;
};
export type AiBridgeGeneratedSlidesFormatTextArgs = {
  slide?: number;
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
};
export type AiBridgeGeneratedSlidesFormatSelectionArgs = {
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
};
export type AiBridgeGeneratedSlidesAddSlideArgs = {
  index?: number;
  title?: string;
  titleFontSize?: number;
  titlePlacement?: "auto" | "placeholder" | "textbox";
  backgroundColor?: string;
  masterIndex?: number;
  layoutIndex?: number;
} & ({ masterIndex: number; layoutIndex: number; } | { masterIndex?: never; });
export type AiBridgeGeneratedSlidesDuplicateSlideArgs = {
  slide: number;
};
export type AiBridgeGeneratedSlidesDeleteSlideArgs = {
  slide: number;
};
export type AiBridgeGeneratedSlidesMoveSlideArgs = {
  slide: number;
  toIndex: number;
};
export type AiBridgeGeneratedSlidesSetVisibilityArgs = {
  slide: number;
  visible: boolean;
};
export type AiBridgeGeneratedSlidesSetSizeArgs = {
  preset?: "wide" | "standard" | "custom";
  widthMm?: number;
  heightMm?: number;
  orientation?: "landscape" | "portrait";
};
export type AiBridgeGeneratedSlidesApplyLayoutArgs = {
  slide: number;
  masterIndex?: number;
  layoutIndex: number;
  includeRaw?: boolean;
};
export type AiBridgeGeneratedSlidesSetShowSettingsArgs = {
  loop: boolean;
};
export type AiBridgeGeneratedSlidesApplyThemeArgs = {
  sourceMasterIndex?: number;
  sourceSlide?: number;
  targetSlide?: number;
  includeRaw?: boolean;
};
export type AiBridgeGeneratedSlidesApplyBuiltinThemeArgs = {
  theme: number | string;
};
export type AiBridgeGeneratedSlidesSetThemeArgs = {
  masterIndex?: number;
  slide?: number;
  colors?: Array<string>;
  colorSchemeName?: string;
  fonts?: {
    majorLatin: string;
    minorLatin: string;
    majorEastAsian?: string;
    majorComplex?: string;
    minorEastAsian?: string;
    minorComplex?: string;
    name?: string;
  };
  includeRaw?: boolean;
} & ({ colors: Array<string>; } | { fonts: {
  majorLatin: string;
  minorLatin: string;
  majorEastAsian?: string;
  majorComplex?: string;
  minorEastAsian?: string;
  minorComplex?: string;
  name?: string;
}; });
export type AiBridgeGeneratedSlidesCreateLayoutArgs = {
  masterIndex?: number;
  name?: string;
  fill?: ({
    type: "none";
  } | {
    type: "solid";
    color: string;
  } | {
    type: "linearGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
    angleDeg?: number;
  } | {
    type: "radialGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
  } | {
    type: "pattern";
    pattern: string;
    backgroundColor: string;
    foregroundColor: string;
  } | {
    type?: "raw";
    raw: unknown;
  });
  followMasterBackground?: boolean;
  placeholders?: Array<{
    type: string;
    shapeType?: string;
    text?: string;
    xMm?: number;
    yMm?: number;
    widthMm?: number;
    heightMm?: number;
    rotationDeg?: number;
    flipH?: boolean;
    flipV?: boolean;
    name?: string;
    fontSize?: number;
    fontFamily?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    color?: string;
    align?: string;
    verticalAlign?: "top" | "center" | "bottom";
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    line?: {
      type?: "solid" | "none";
      enabled?: boolean;
      widthPt?: number;
      color?: string;
      fill?: ({
        type: "none";
      } | {
        type: "solid";
        color: string;
      } | {
        type: "linearGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
        angleDeg?: number;
      } | {
        type: "radialGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
      } | {
        type: "pattern";
        pattern: string;
        backgroundColor: string;
        foregroundColor: string;
      } | {
        type?: "raw";
        raw: unknown;
      });
      raw?: unknown;
    };
    paddingMm?: {
      left?: number;
      top?: number;
      right?: number;
      bottom?: number;
    };
    fillColor?: string;
    lineColor?: string;
    lineWidthPt?: number;
  }>;
  applyToSlides?: Array<number>;
  includeRaw?: boolean;
};
export type AiBridgeGeneratedSlidesAddTemplateShapeArgs = {
  scope: "master" | "layout";
  masterIndex?: number;
  layoutIndex?: number;
  shapeType: string;
  placeholderType?: string;
  text?: string;
  xMm?: number;
  yMm?: number;
  widthMm?: number;
  heightMm?: number;
  rotationDeg?: number;
  flipH?: boolean;
  flipV?: boolean;
  name?: string;
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
  align?: string;
  verticalAlign?: "top" | "center" | "bottom";
  fill?: ({
    type: "none";
  } | {
    type: "solid";
    color: string;
  } | {
    type: "linearGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
    angleDeg?: number;
  } | {
    type: "radialGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
  } | {
    type: "pattern";
    pattern: string;
    backgroundColor: string;
    foregroundColor: string;
  } | {
    type?: "raw";
    raw: unknown;
  });
  line?: {
    type?: "solid" | "none";
    enabled?: boolean;
    widthPt?: number;
    color?: string;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    raw?: unknown;
  };
  paddingMm?: {
    left?: number;
    top?: number;
    right?: number;
    bottom?: number;
  };
  fillColor?: string;
  lineColor?: string;
  lineWidthPt?: number;
  includeRaw?: boolean;
};
export type AiBridgeGeneratedSlidesManageTemplateObjectArgs = {
  scope: "master" | "layout";
  masterIndex?: number;
  layoutIndex?: number;
  action: "update" | "delete";
  objectId?: string;
  objectIndex?: number;
  name?: string;
  newName?: string;
  shapeType?: string;
  text?: string;
  xMm?: number;
  yMm?: number;
  widthMm?: number;
  heightMm?: number;
  rotationDeg?: number;
  flipH?: boolean;
  flipV?: boolean;
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
  align?: string;
  verticalAlign?: "top" | "center" | "bottom";
  fill?: ({
    type: "none";
  } | {
    type: "solid";
    color: string;
  } | {
    type: "linearGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
    angleDeg?: number;
  } | {
    type: "radialGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
  } | {
    type: "pattern";
    pattern: string;
    backgroundColor: string;
    foregroundColor: string;
  } | {
    type?: "raw";
    raw: unknown;
  });
  line?: {
    type?: "solid" | "none";
    enabled?: boolean;
    widthPt?: number;
    color?: string;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    raw?: unknown;
  };
  paddingMm?: {
    left?: number;
    top?: number;
    right?: number;
    bottom?: number;
  };
  fillColor?: string;
  lineColor?: string;
  lineWidthPt?: number;
  includeRaw?: boolean;
} & ({ objectId: string; } | { objectIndex: number; } | { name: string; });
export type AiBridgeGeneratedSlidesSetTemplateBackgroundArgs = {
  scope: "master" | "layout";
  masterIndex?: number;
  layoutIndex?: number;
  mode?: "custom" | "image" | "clear" | "master";
  fill?: ({
    type: "none";
  } | {
    type: "solid";
    color: string;
  } | {
    type: "linearGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
    angleDeg?: number;
  } | {
    type: "radialGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
  } | {
    type: "pattern";
    pattern: string;
    backgroundColor: string;
    foregroundColor: string;
  } | {
    type?: "raw";
    raw: unknown;
  });
  source?: ({
    type: "url";
    url: string;
  } | {
    type: "dataUrl";
    dataUrl: string;
  } | {
    type: "relayAsset";
    path: string;
    assetToken: string;
    mimeType: "image/png" | "image/jpeg" | "image/gif" | "image/webp" | "image/svg+xml";
    widthPx: number;
    heightPx: number;
    assetId: string;
  });
  fillMode?: "stretch" | "tile";
} & ({ mode: "image"; source: ({
  type: "url";
  url: string;
} | {
  type: "dataUrl";
  dataUrl: string;
} | {
  type: "relayAsset";
  path: string;
  assetToken: string;
  mimeType: "image/png" | "image/jpeg" | "image/gif" | "image/webp" | "image/svg+xml";
  widthPx: number;
  heightPx: number;
  assetId: string;
}); } | { mode?: Exclude<"custom" | "image" | "clear" | "master", "image">; }) & ({ fill: ({
  type: "none";
} | {
  type: "solid";
  color: string;
} | {
  type: "linearGradient";
  stops: Array<{
    position: number;
    color: string;
  }>;
  angleDeg?: number;
} | {
  type: "radialGradient";
  stops: Array<{
    position: number;
    color: string;
  }>;
} | {
  type: "pattern";
  pattern: string;
  backgroundColor: string;
  foregroundColor: string;
} | {
  type?: "raw";
  raw: unknown;
}); } | {  }) & ({ mode: "custom" | "image" | "clear" | "master"; } | { mode?: never; });
export type AiBridgeGeneratedSlidesSetTextContentArgs = {
  slide: number;
  objectId?: string;
  objectIndex?: number;
  name?: string;
  paragraphs: Array<{
    text: string;
    fontSize?: number;
    fontFamily?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    color?: string;
    align?: "left" | "center" | "right" | "both";
    firstLineIndentMm?: number;
    leftIndentMm?: number;
    rightIndentMm?: number;
    spacingBeforePt?: number;
    spacingAfterPt?: number;
    lineSpacing?: number;
    lineRule?: "auto" | "exact" | "atLeast";
    level?: number;
    listType?: "none" | "bullet" | "number";
    bulletSymbol?: string;
    numberingType?: string;
    startAt?: number;
  }>;
  includeRaw?: boolean;
} & ({ objectId: string; } | { objectIndex: number; } | { name: string; });
export type AiBridgeGeneratedSlidesFormatParagraphsArgs = {
  slide: number;
  objectId?: string;
  objectIndex?: number;
  name?: string;
  paragraphIndexes?: Array<number>;
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
  align?: "left" | "center" | "right" | "both";
  firstLineIndentMm?: number;
  leftIndentMm?: number;
  rightIndentMm?: number;
  spacingBeforePt?: number;
  spacingAfterPt?: number;
  lineSpacing?: number;
  lineRule?: "auto" | "exact" | "atLeast";
  level?: number;
  listType?: "none" | "bullet" | "number";
  bulletSymbol?: string;
  numberingType?: string;
  startAt?: number;
  includeRaw?: boolean;
} & ({ objectId: string; } | { objectIndex: number; } | { name: string; });
export type AiBridgeGeneratedSlidesUpdateObjectArgs = {
  slide: number;
  objectId?: string;
  objectIndex?: number;
  name?: string;
  xMm?: number;
  yMm?: number;
  widthMm?: number;
  heightMm?: number;
  rotationDeg?: number;
  flipH?: boolean;
  flipV?: boolean;
  line?: {
    type?: "solid" | "none";
    enabled?: boolean;
    widthPt?: number;
    color?: string;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    raw?: unknown;
  };
  lineColor?: string;
  lineWidthPt?: number;
  includeRaw?: boolean;
} & ({ objectId: string; } | { objectIndex: number; } | { name: string; });
export type AiBridgeGeneratedSlidesSetHyperlinkArgs = {
  slide: number;
  objectId?: string;
  objectIndex?: number;
  name?: string;
  action?: "external" | "firstSlide" | "lastSlide" | "nextSlide" | "previousSlide" | "slide" | "remove";
  url?: string;
  targetSlide?: number;
  tooltip?: string;
} & ({ objectId: string; } | { objectIndex: number; } | { name: string; });
export type AiBridgeGeneratedSlidesSetNotesArgs = {
  slide: number;
  text: string;
  append?: boolean;
};
export type AiBridgeGeneratedSlidesAddCommentArgs = {
  slide: number;
  text: string;
  xMm?: number;
  yMm?: number;
  author?: string;
  userId?: string;
};
export type AiBridgeGeneratedSlidesInspectCommentsArgs = Record<string, never>;
export type AiBridgeGeneratedSlidesManageCommentArgs = {
  action: "update" | "addReply" | "removeReplies" | "delete";
  commentId?: string;
  commentIndex?: number;
  text?: string;
  author?: string;
  userId?: string;
  solved?: boolean;
  xMm?: number;
  yMm?: number;
  start?: number;
  count?: number;
} & ({ commentId: string; } | { commentIndex: number; });
export type AiBridgeGeneratedSlidesSetTransitionArgs = {
  slide: number;
  clear?: boolean;
  effect?: string;
  speed?: "slow" | "medium" | "fast";
  durationMs?: number;
  advanceOnClick?: boolean;
  advanceOnTime?: boolean;
  advanceTimeMs?: number;
};
export type AiBridgeGeneratedSlidesInspectAnimationsArgs = {
  slide?: number;
};
export type AiBridgeGeneratedSlidesManageAnimationArgs = {
  slide: number;
  action: "add" | "update" | "delete" | "clear";
  objectId?: string;
  objectIndex?: number;
  name?: string;
  effectIndex?: number;
  sequence?: "main" | "interactive";
  triggerObject?: {
    objectId?: string;
    objectIndex?: number;
    name?: string;
  } & ({ objectId: string; } | { objectIndex: number; } | { name: string; });
  effectType?: string;
  trigger?: "onclick" | "withprevious" | "afterprevious";
  durationMs?: number;
  delayMs?: number;
  repeatCount?: number;
  toIndex?: number;
};
export type AiBridgeGeneratedSlidesAddTableArgs = {
  slide: number;
  rows: number;
  columns: number;
  data?: Array<Array<(string | number | boolean | null | {
    text: string;
    fontSize?: number;
    fontFamily?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    color?: string;
    align?: "left" | "center" | "right" | "both";
    firstLineIndentMm?: number;
    leftIndentMm?: number;
    rightIndentMm?: number;
    spacingBeforePt?: number;
    spacingAfterPt?: number;
    lineSpacing?: number;
    lineRule?: "auto" | "exact" | "atLeast";
    level?: number;
    listType?: "none" | "bullet" | "number";
    bulletSymbol?: string;
    numberingType?: string;
    startAt?: number;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    backgroundColor?: string;
    verticalAlign?: "top" | "center" | "bottom";
    border?: {
      widthMm?: number;
      color?: string;
      fill?: ({
        type: "none";
      } | {
        type: "solid";
        color: string;
      } | {
        type: "linearGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
        angleDeg?: number;
      } | {
        type: "radialGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
      } | {
        type: "pattern";
        pattern: string;
        backgroundColor: string;
        foregroundColor: string;
      } | {
        type?: "raw";
        raw: unknown;
      });
      sides?: Array<"top" | "right" | "bottom" | "left">;
    };
  })>>;
  xMm?: number;
  yMm?: number;
  widthMm?: number;
  heightMm?: number;
  name?: string;
  header?: {
    text?: string;
    fontSize?: number;
    fontFamily?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    color?: string;
    align?: "left" | "center" | "right" | "both";
    firstLineIndentMm?: number;
    leftIndentMm?: number;
    rightIndentMm?: number;
    spacingBeforePt?: number;
    spacingAfterPt?: number;
    lineSpacing?: number;
    lineRule?: "auto" | "exact" | "atLeast";
    level?: number;
    listType?: "none" | "bullet" | "number";
    bulletSymbol?: string;
    numberingType?: string;
    startAt?: number;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    backgroundColor?: string;
    verticalAlign?: "top" | "center" | "bottom";
    border?: {
      widthMm?: number;
      color?: string;
      fill?: ({
        type: "none";
      } | {
        type: "solid";
        color: string;
      } | {
        type: "linearGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
        angleDeg?: number;
      } | {
        type: "radialGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
      } | {
        type: "pattern";
        pattern: string;
        backgroundColor: string;
        foregroundColor: string;
      } | {
        type?: "raw";
        raw: unknown;
      });
      sides?: Array<"top" | "right" | "bottom" | "left">;
    };
  };
  includeRaw?: boolean;
};
export type AiBridgeGeneratedSlidesSetTableCellArgs = {
  slide: number;
  objectId?: string;
  objectIndex?: number;
  name?: string;
  row: number;
  column: number;
  text?: string;
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
  align?: "left" | "center" | "right" | "both";
  fill?: ({
    type: "none";
  } | {
    type: "solid";
    color: string;
  } | {
    type: "linearGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
    angleDeg?: number;
  } | {
    type: "radialGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
  } | {
    type: "pattern";
    pattern: string;
    backgroundColor: string;
    foregroundColor: string;
  } | {
    type?: "raw";
    raw: unknown;
  });
  backgroundColor?: string;
  verticalAlign?: "top" | "center" | "bottom";
  border?: {
    widthMm?: number;
    color?: string;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    sides?: Array<"top" | "right" | "bottom" | "left">;
  };
  includeRaw?: boolean;
} & ({ objectId: string; } | { objectIndex: number; } | { name: string; });
export type AiBridgeGeneratedSlidesEditTableArgs = {
  slide: number;
  objectId?: string;
  objectIndex?: number;
  name?: string;
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
} & ({ objectId: string; } | { objectIndex: number; } | { name: string; });
export type AiBridgeGeneratedSlidesFormatTableArgs = {
  slide: number;
  objectId?: string;
  objectIndex?: number;
  name?: string;
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
  columnWidthsMm?: Array<number>;
  rowHeightsMm?: Array<number>;
  tableLook?: {
    firstColumn?: boolean;
    firstRow?: boolean;
    lastColumn?: boolean;
    lastRow?: boolean;
    horizontalBanding?: boolean;
    verticalBanding?: boolean;
  };
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
  align?: "left" | "center" | "right" | "both";
  fill?: ({
    type: "none";
  } | {
    type: "solid";
    color: string;
  } | {
    type: "linearGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
    angleDeg?: number;
  } | {
    type: "radialGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
  } | {
    type: "pattern";
    pattern: string;
    backgroundColor: string;
    foregroundColor: string;
  } | {
    type?: "raw";
    raw: unknown;
  });
  backgroundColor?: string;
  verticalAlign?: "top" | "center" | "bottom";
  border?: {
    widthMm?: number;
    color?: string;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    sides?: Array<"top" | "right" | "bottom" | "left">;
  };
  includeRaw?: boolean;
} & ({ objectId: string; } | { objectIndex: number; } | { name: string; });
export type AiBridgeGeneratedSlidesAlignObjectsArgs = {
  slide: number;
  targets: Array<{
    objectId?: string;
    objectIndex?: number;
    name?: string;
  } & ({ objectId: string; } | { objectIndex: number; } | { name: string; })>;
  align?: "left" | "center" | "right" | "top" | "middle" | "bottom";
  distribute?: "horizontal" | "vertical";
  relativeTo?: "selection" | "slide";
} & ({ align: "left" | "center" | "right" | "top" | "middle" | "bottom"; } | { distribute: "horizontal" | "vertical"; });
export type AiBridgeGeneratedSlidesGroupObjectsArgs = {
  slide: number;
  action?: "group" | "ungroup";
  targets?: Array<{
    objectId?: string;
    objectIndex?: number;
    name?: string;
  } & ({ objectId: string; } | { objectIndex: number; } | { name: string; })>;
  objectId?: string;
  objectIndex?: number;
  name?: string;
  includeRaw?: boolean;
};
export type AiBridgeGeneratedSlidesReorderObjectArgs = {
  slide: number;
  objectId?: string;
  objectIndex?: number;
  name?: string;
  action: "front" | "back" | "forward" | "backward";
} & ({ objectId: string; } | { objectIndex: number; } | { name: string; });
export type AiBridgeGeneratedSlidesAddConnectorArgs = {
  slide: number;
  connectorType?: string;
  startXmm: number;
  startYmm: number;
  endXmm: number;
  endYmm: number;
  name?: string;
  line?: {
    type?: "solid" | "none";
    enabled?: boolean;
    widthPt?: number;
    color?: string;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    raw?: unknown;
  };
  lineColor?: string;
  lineWidthPt?: number;
  includeRaw?: boolean;
};
export type AiBridgeGeneratedSlidesAddFreeformArgs = {
  slide: number;
  paths: Array<{
    fill?: string;
    stroke?: boolean;
    commands: Array<{
      type: "moveTo" | "lineTo" | "quadBezTo" | "cubicBezTo" | "arcTo" | "close";
      xMm?: number;
      yMm?: number;
      controlXmm?: number;
      controlYmm?: number;
      control1Xmm?: number;
      control1Ymm?: number;
      control2Xmm?: number;
      control2Ymm?: number;
      widthRadiusMm?: number;
      heightRadiusMm?: number;
      startAngleDeg?: number;
      sweepAngleDeg?: number;
    }>;
  }>;
  xMm?: number;
  yMm?: number;
  widthMm?: number;
  heightMm?: number;
  name?: string;
  fill?: ({
    type: "none";
  } | {
    type: "solid";
    color: string;
  } | {
    type: "linearGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
    angleDeg?: number;
  } | {
    type: "radialGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
  } | {
    type: "pattern";
    pattern: string;
    backgroundColor: string;
    foregroundColor: string;
  } | {
    type?: "raw";
    raw: unknown;
  });
  line?: {
    type?: "solid" | "none";
    enabled?: boolean;
    widthPt?: number;
    color?: string;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    raw?: unknown;
  };
  fillColor?: string;
  lineColor?: string;
  lineWidthPt?: number;
  includeRaw?: boolean;
};
export type AiBridgeGeneratedSlidesAddTextboxArgs = {
  slide: number;
  text?: string;
  paragraphs?: Array<{
    text: string;
    fontSize?: number;
    fontFamily?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    color?: string;
    align?: "left" | "center" | "right" | "both";
    firstLineIndentMm?: number;
    leftIndentMm?: number;
    rightIndentMm?: number;
    spacingBeforePt?: number;
    spacingAfterPt?: number;
    lineSpacing?: number;
    lineRule?: "auto" | "exact" | "atLeast";
    level?: number;
    listType?: "none" | "bullet" | "number";
    bulletSymbol?: string;
    numberingType?: string;
    startAt?: number;
  }>;
  xMm?: number;
  yMm?: number;
  widthMm?: number;
  heightMm?: number;
  rotationDeg?: number;
  flipH?: boolean;
  flipV?: boolean;
  name?: string;
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  fill?: ({
    type: "none";
  } | {
    type: "solid";
    color: string;
  } | {
    type: "linearGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
    angleDeg?: number;
  } | {
    type: "radialGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
  } | {
    type: "pattern";
    pattern: string;
    backgroundColor: string;
    foregroundColor: string;
  } | {
    type?: "raw";
    raw: unknown;
  });
  line?: {
    type?: "solid" | "none";
    enabled?: boolean;
    widthPt?: number;
    color?: string;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    raw?: unknown;
  };
  paddingMm?: {
    left?: number;
    top?: number;
    right?: number;
    bottom?: number;
  };
  fillColor?: string;
  lineColor?: string;
  lineWidthPt?: number;
  align?: string;
  verticalAlign?: "top" | "center" | "bottom";
  includeRaw?: boolean;
} & ({ text: string; } | { paragraphs: Array<{
  text: string;
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
  align?: "left" | "center" | "right" | "both";
  firstLineIndentMm?: number;
  leftIndentMm?: number;
  rightIndentMm?: number;
  spacingBeforePt?: number;
  spacingAfterPt?: number;
  lineSpacing?: number;
  lineRule?: "auto" | "exact" | "atLeast";
  level?: number;
  listType?: "none" | "bullet" | "number";
  bulletSymbol?: string;
  numberingType?: string;
  startAt?: number;
}>; });
export type AiBridgeGeneratedSlidesAddWordArtArgs = {
  slide: number;
  text: string;
  transform?: string;
  xMm?: number;
  yMm?: number;
  widthMm?: number;
  heightMm?: number;
  rotationDeg?: number;
  name?: string;
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  fill?: ({
    type: "none";
  } | {
    type: "solid";
    color: string;
  } | {
    type: "linearGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
    angleDeg?: number;
  } | {
    type: "radialGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
  } | {
    type: "pattern";
    pattern: string;
    backgroundColor: string;
    foregroundColor: string;
  } | {
    type?: "raw";
    raw: unknown;
  });
  line?: {
    type?: "solid" | "none";
    enabled?: boolean;
    widthPt?: number;
    color?: string;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    raw?: unknown;
  };
  fillColor?: string;
  lineColor?: string;
  lineWidthPt?: number;
  includeRaw?: boolean;
};
export type AiBridgeGeneratedSlidesAddMathArgs = {
  slide: number;
  text: string;
  format?: "latex" | "unicode" | "mathml";
  xMm?: number;
  yMm?: number;
  widthMm?: number;
  heightMm?: number;
  rotationDeg?: number;
  name?: string;
  includeRaw?: boolean;
};
export type AiBridgeGeneratedSlidesAddImageArgs = {
  source: ({
    type: "url";
    url: string;
  } | {
    type: "dataUrl";
    dataUrl: string;
  } | {
    type: "relayAsset";
    path: string;
    assetToken: string;
    mimeType: "image/png" | "image/jpeg" | "image/gif" | "image/webp" | "image/svg+xml";
    widthPx: number;
    heightPx: number;
    assetId: string;
  });
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
};
export type AiBridgeGeneratedSlidesAddImageShapeArgs = {
  source: ({
    type: "url";
    url: string;
  } | {
    type: "dataUrl";
    dataUrl: string;
  } | {
    type: "relayAsset";
    path: string;
    assetToken: string;
    mimeType: "image/png" | "image/jpeg" | "image/gif" | "image/webp" | "image/svg+xml";
    widthPx: number;
    heightPx: number;
    assetId: string;
  });
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
  line?: {
    type?: "solid" | "none";
    enabled?: boolean;
    widthPt?: number;
    color?: string;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    raw?: unknown;
  };
  lineColor?: string;
  lineWidthPt?: number;
  includeRaw?: boolean;
};
export type AiBridgeGeneratedSlidesAddOleObjectArgs = {
  source: ({
    type: "url";
    url: string;
  } | {
    type: "dataUrl";
    dataUrl: string;
  } | {
    type: "relayAsset";
    path: string;
    assetToken: string;
    mimeType: "image/png" | "image/jpeg" | "image/gif" | "image/webp" | "image/svg+xml";
    widthPx: number;
    heightPx: number;
    assetId: string;
  });
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
};
export type AiBridgeGeneratedSlidesSetBackgroundArgs = {
  slide: number;
  mode?: "custom" | "image" | "clear" | "layout" | "master";
  fill?: ({
    type: "none";
  } | {
    type: "solid";
    color: string;
  } | {
    type: "linearGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
    angleDeg?: number;
  } | {
    type: "radialGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
  } | {
    type: "pattern";
    pattern: string;
    backgroundColor: string;
    foregroundColor: string;
  } | {
    type?: "raw";
    raw: unknown;
  });
  source?: ({
    type: "url";
    url: string;
  } | {
    type: "dataUrl";
    dataUrl: string;
  } | {
    type: "relayAsset";
    path: string;
    assetToken: string;
    mimeType: "image/png" | "image/jpeg" | "image/gif" | "image/webp" | "image/svg+xml";
    widthPx: number;
    heightPx: number;
    assetId: string;
  });
  fillMode?: "stretch" | "tile";
} & ({ mode: "image"; source: ({
  type: "url";
  url: string;
} | {
  type: "dataUrl";
  dataUrl: string;
} | {
  type: "relayAsset";
  path: string;
  assetToken: string;
  mimeType: "image/png" | "image/jpeg" | "image/gif" | "image/webp" | "image/svg+xml";
  widthPx: number;
  heightPx: number;
  assetId: string;
}); } | { mode?: Exclude<"custom" | "image" | "clear" | "layout" | "master", "image">; }) & ({ fill: ({
  type: "none";
} | {
  type: "solid";
  color: string;
} | {
  type: "linearGradient";
  stops: Array<{
    position: number;
    color: string;
  }>;
  angleDeg?: number;
} | {
  type: "radialGradient";
  stops: Array<{
    position: number;
    color: string;
  }>;
} | {
  type: "pattern";
  pattern: string;
  backgroundColor: string;
  foregroundColor: string;
} | {
  type?: "raw";
  raw: unknown;
}); } | {  }) & ({ mode: "custom" | "image" | "clear" | "layout" | "master"; } | { mode?: never; });
export type AiBridgeGeneratedSlidesAddShapeArgs = {
  slide: number;
  shapeType: string;
  text?: string;
  paragraphs?: Array<{
    text: string;
    fontSize?: number;
    fontFamily?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    color?: string;
    align?: "left" | "center" | "right" | "both";
    firstLineIndentMm?: number;
    leftIndentMm?: number;
    rightIndentMm?: number;
    spacingBeforePt?: number;
    spacingAfterPt?: number;
    lineSpacing?: number;
    lineRule?: "auto" | "exact" | "atLeast";
    level?: number;
    listType?: "none" | "bullet" | "number";
    bulletSymbol?: string;
    numberingType?: string;
    startAt?: number;
  }>;
  xMm?: number;
  yMm?: number;
  widthMm?: number;
  heightMm?: number;
  rotationDeg?: number;
  flipH?: boolean;
  flipV?: boolean;
  name?: string;
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  fill?: ({
    type: "none";
  } | {
    type: "solid";
    color: string;
  } | {
    type: "linearGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
    angleDeg?: number;
  } | {
    type: "radialGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
  } | {
    type: "pattern";
    pattern: string;
    backgroundColor: string;
    foregroundColor: string;
  } | {
    type?: "raw";
    raw: unknown;
  });
  line?: {
    type?: "solid" | "none";
    enabled?: boolean;
    widthPt?: number;
    color?: string;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    raw?: unknown;
  };
  paddingMm?: {
    left?: number;
    top?: number;
    right?: number;
    bottom?: number;
  };
  fillColor?: string;
  lineColor?: string;
  lineWidthPt?: number;
  align?: string;
  verticalAlign?: "top" | "center" | "bottom";
  includeRaw?: boolean;
};
export type AiBridgeGeneratedSlidesUpdateShapeArgs = {
  slide: number;
  objectId?: string;
  objectIndex?: number;
  name?: string;
  shapeType?: string;
  text?: string;
  xMm?: number;
  yMm?: number;
  widthMm?: number;
  heightMm?: number;
  rotationDeg?: number;
  flipH?: boolean;
  flipV?: boolean;
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  fill?: ({
    type: "none";
  } | {
    type: "solid";
    color: string;
  } | {
    type: "linearGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
    angleDeg?: number;
  } | {
    type: "radialGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
  } | {
    type: "pattern";
    pattern: string;
    backgroundColor: string;
    foregroundColor: string;
  } | {
    type?: "raw";
    raw: unknown;
  });
  line?: {
    type?: "solid" | "none";
    enabled?: boolean;
    widthPt?: number;
    color?: string;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    raw?: unknown;
  };
  paddingMm?: {
    left?: number;
    top?: number;
    right?: number;
    bottom?: number;
  };
  fillColor?: string;
  lineColor?: string;
  lineWidthPt?: number;
  align?: string;
  verticalAlign?: "top" | "center" | "bottom";
  includeRaw?: boolean;
} & ({ objectId: string; } | { objectIndex: number; } | { name: string; });
export type AiBridgeGeneratedSlidesDeleteObjectArgs = {
  slide: number;
  objectId?: string;
  objectIndex?: number;
  name?: string;
} & ({ objectId: string; } | { objectIndex: number; } | { name: string; });
export type AiBridgeGeneratedSlidesInspectSmartartsArgs = {
  slide?: number;
  maxSmartArts?: number;
  includeRaw?: boolean;
  includeTextStyles?: boolean;
};
export type AiBridgeGeneratedSlidesAddSmartartArgs = {
  slide: number;
  type: ("AccentedPicture" | "Balance" | "TitledPictureBlocks" | "PictureAccentBlocks" | "BlockCycle" | "StackedVenn" | "VerticalEquation" | "VerticalBlockList" | "VerticalBendingProcess" | "VerticalBulletList" | "VerticalCurvedList" | "VerticalProcess" | "VerticalBoxList" | "VerticalPictureList" | "VerticalCircleList" | "VerticalPictureAccentList" | "VerticalArrowList" | "VerticalChevronList" | "VerticalAccentList" | "NestedTarget" | "Funnel" | "UpwardArrow" | "IncreasingArrowsProcess" | "StepUpProcess" | "CircularPictureCallout" | "HorizontalHierarchy" | "HorizontalLabeledHierarchy" | "HorizontalMultiLevelHierarchy" | "HorizontalOrganizationChart" | "HorizontalBulletList" | "HorizontalPictureList" | "ClosedChevronProcess" | "HierarchyList" | "Hierarchy" | "CirclePictureHierarchy" | "LabeledHierarchy" | "InvertedPyramid" | "HexagonCluster" | "CircleRelationship" | "CircleAccentTimeline" | "CircularBendingProcess" | "ArrowRibbon" | "LinearVenn" | "PictureLineup" | "TitlePictureLineup" | "BendingPictureCaptionList" | "BendingPictureAccentList" | "TitledMatrix" | "IncreasingCircleProcess" | "BendingPictureBlocks" | "BendingPictureCaption" | "BendingPictureSemiTransparentText" | "NonDirectionalCycle" | "ContinuousBlockProcess" | "ContinuousPictureList" | "ContinuousCycle" | "DescendingBlockList" | "StepDownProcess" | "ReverseList" | "OrganizationChart" | "NameAndTitleOrganizationChart" | "AlternatingFlow" | "PyramidList" | "PlusAndMinus" | "RepeatingBendingProcess" | "CaptionedPictures" | "DetailedProcess" | "PictureStrips" | "HalfCircleOrganizationChart" | "PhasedProcess" | "BasicVenn" | "BasicTimeline" | "BasicPie" | "BasicMatrix" | "BasicPyramid" | "BasicRadial" | "BasicTarget" | "BasicBlockList" | "BasicBendingProcess" | "BasicProcess" | "BasicChevronProcess" | "BasicCycle" | "OpposingIdeas" | "OpposingArrows" | "RandomToResultProcess" | "SubStepProcess" | "PieProcess" | "AccentProcess" | "AscendingPictureAccentProcess" | "PictureAccentProcess" | "RadialVenn" | "RadialCycle" | "RadialCluster" | "RadialList" | "MultiDirectionalCycle" | "DivergingRadial" | "DivergingArrows" | "FramedTextPicture" | "GroupedList" | "SegmentedPyramid" | "SegmentedProcess" | "SegmentedCycle" | "PictureGrid" | "GridMatrix" | "SpiralPicture" | "StackedList" | "PictureCaptionList" | "ProcessList" | "BubblePictureList" | "SquareAccentList" | "LinedList" | "PictureAccentList" | "TitledPictureAccentList" | "SnapshotPictureList" | "ContinuousArrowProcess" | "CircleArrowProcess" | "ProcessArrows" | "StaggeredProcess" | "ConvergingRadial" | "ConvergingArrows" | "TableHierarchy" | "TableList" | "TextCycle" | "TrapezoidList" | "DescendingProcess" | "ChevronList" | "Equation" | "CounterbalanceArrows" | "TargetList" | "CycleMatrix" | "AlternatingPictureBlocks" | "AlternatingPictureCircles" | "AlternatingHexagonList" | "Gear" | "ArchitectureLayout" | "ChevronAccentProcess" | "CircleProcess" | "ConvergingText" | "HexagonRadial" | "InterconnectedBlockProcess" | "InterconnectedRings" | "PictureFrame" | "PictureOrganizationChart" | "RadialPictureList" | "TabList" | "TabbedArc" | "ThemePictureAccent" | "ThemePictureAlternatingAccent" | "ThemePictureGrid" | "VaryingWidthList" | "VerticalBracketList" | number);
  xMm?: number;
  yMm?: number;
  widthMm?: number;
  heightMm?: number;
  rotationDeg?: number;
  flipH?: boolean;
  flipV?: boolean;
  name?: string;
  nodeFill?: ({
    type: "none";
  } | {
    type: "solid";
    color: string;
  } | {
    type: "linearGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
    angleDeg?: number;
  } | {
    type: "radialGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
  } | {
    type: "pattern";
    pattern: string;
    backgroundColor: string;
    foregroundColor: string;
  } | {
    type?: "raw";
    raw: unknown;
  });
  nodeLine?: {
    type?: "solid" | "none";
    enabled?: boolean;
    widthPt?: number;
    color?: string;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    raw?: unknown;
  };
  nodes?: Array<{
    index?: number;
    nodeId?: string;
    text?: string;
    paragraphs?: Array<{
      text: string;
      fontSize?: number;
      fontFamily?: string;
      bold?: boolean;
      italic?: boolean;
      underline?: boolean;
      color?: string;
      align?: "left" | "center" | "right" | "both";
      firstLineIndentMm?: number;
      leftIndentMm?: number;
      rightIndentMm?: number;
      spacingBeforePt?: number;
      spacingAfterPt?: number;
      lineSpacing?: number;
      lineRule?: "auto" | "exact" | "atLeast";
      level?: number;
      listType?: "none" | "bullet" | "number";
      bulletSymbol?: string;
      numberingType?: string;
      startAt?: number;
    }>;
    fontSize?: number;
    fontFamily?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    color?: string;
    align?: "left" | "center" | "right" | "both";
    verticalAlign?: "top" | "center" | "bottom";
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    line?: {
      type?: "solid" | "none";
      enabled?: boolean;
      widthPt?: number;
      color?: string;
      fill?: ({
        type: "none";
      } | {
        type: "solid";
        color: string;
      } | {
        type: "linearGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
        angleDeg?: number;
      } | {
        type: "radialGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
      } | {
        type: "pattern";
        pattern: string;
        backgroundColor: string;
        foregroundColor: string;
      } | {
        type?: "raw";
        raw: unknown;
      });
      raw?: unknown;
    };
  } & ({ index: number; } | { nodeId: string; })>;
  includeRaw?: boolean;
  includeTextStyles?: boolean;
};
export type AiBridgeGeneratedSlidesUpdateSmartartArgs = {
  slide: number;
  smartArtId?: string;
  smartArtIndex?: number;
  name?: string;
  xMm?: number;
  yMm?: number;
  widthMm?: number;
  heightMm?: number;
  rotationDeg?: number;
  flipH?: boolean;
  flipV?: boolean;
  nodeFill?: ({
    type: "none";
  } | {
    type: "solid";
    color: string;
  } | {
    type: "linearGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
    angleDeg?: number;
  } | {
    type: "radialGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
  } | {
    type: "pattern";
    pattern: string;
    backgroundColor: string;
    foregroundColor: string;
  } | {
    type?: "raw";
    raw: unknown;
  });
  nodeLine?: {
    type?: "solid" | "none";
    enabled?: boolean;
    widthPt?: number;
    color?: string;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    raw?: unknown;
  };
  nodes?: Array<{
    index?: number;
    nodeId?: string;
    text?: string;
    paragraphs?: Array<{
      text: string;
      fontSize?: number;
      fontFamily?: string;
      bold?: boolean;
      italic?: boolean;
      underline?: boolean;
      color?: string;
      align?: "left" | "center" | "right" | "both";
      firstLineIndentMm?: number;
      leftIndentMm?: number;
      rightIndentMm?: number;
      spacingBeforePt?: number;
      spacingAfterPt?: number;
      lineSpacing?: number;
      lineRule?: "auto" | "exact" | "atLeast";
      level?: number;
      listType?: "none" | "bullet" | "number";
      bulletSymbol?: string;
      numberingType?: string;
      startAt?: number;
    }>;
    fontSize?: number;
    fontFamily?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    color?: string;
    align?: "left" | "center" | "right" | "both";
    verticalAlign?: "top" | "center" | "bottom";
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    line?: {
      type?: "solid" | "none";
      enabled?: boolean;
      widthPt?: number;
      color?: string;
      fill?: ({
        type: "none";
      } | {
        type: "solid";
        color: string;
      } | {
        type: "linearGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
        angleDeg?: number;
      } | {
        type: "radialGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
      } | {
        type: "pattern";
        pattern: string;
        backgroundColor: string;
        foregroundColor: string;
      } | {
        type?: "raw";
        raw: unknown;
      });
      raw?: unknown;
    };
  } & ({ index: number; } | { nodeId: string; })>;
  includeRaw?: boolean;
  includeTextStyles?: boolean;
} & ({ smartArtId: string; } | { smartArtIndex: number; } | { name: string; });
export type AiBridgeGeneratedSlidesDeleteSmartartArgs = {
  slide: number;
  smartArtId?: string;
  smartArtIndex?: number;
  name?: string;
} & ({ smartArtId: string; } | { smartArtIndex: number; } | { name: string; });
export type AiBridgeGeneratedSlidesInspectChartsArgs = {
  slide?: number;
  maxCharts?: number;
  includeRaw?: boolean;
};
export type AiBridgeGeneratedSlidesAddChartArgs = {
  slide: number;
  type?: "bar" | "barStacked" | "barStackedPercent" | "bar3D" | "barStacked3D" | "barStackedPercent3D" | "barStackedPercent3DPerspective" | "horizontalBar" | "horizontalBarStacked" | "horizontalBarStackedPercent" | "horizontalBar3D" | "horizontalBarStacked3D" | "horizontalBarStackedPercent3D" | "lineNormal" | "lineStacked" | "lineStackedPercent" | "lineNormalMarker" | "lineStackedMarker" | "lineStackedPerMarker" | "line3D" | "pie" | "pie3D" | "doughnut" | "scatter" | "scatterLine" | "scatterLineMarker" | "scatterSmooth" | "scatterSmoothMarker" | "stock" | "area" | "areaStacked" | "areaStackedPercent" | "comboCustom" | "comboBarLine" | "comboBarLineSecondary" | "radar" | "radarMarker" | "radarFilled" | "line" | "lineMarker" | "stackedBar" | "stackedBarPercent" | "stackedLine" | "stackedLinePercent" | "column";
  series: Array<Array<number>>;
  seriesNames: Array<string | number>;
  categories: Array<string | number>;
  numFormats?: Array<string>;
  xMm?: number;
  yMm?: number;
  widthMm?: number;
  heightMm?: number;
  name?: string;
  style?: number;
  holeSizePercent?: number;
  title?: string;
  titleFontSize?: number;
  titleBold?: boolean;
  rotationDeg?: number;
  fill?: ({
    type: "none";
  } | {
    type: "solid";
    color: string;
  } | {
    type: "linearGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
    angleDeg?: number;
  } | {
    type: "radialGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
  } | {
    type: "pattern";
    pattern: string;
    backgroundColor: string;
    foregroundColor: string;
  } | {
    type?: "raw";
    raw: unknown;
  });
  line?: {
    type?: "solid" | "none";
    enabled?: boolean;
    widthPt?: number;
    color?: string;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    raw?: unknown;
  };
  plotAreaFill?: ({
    type: "none";
  } | {
    type: "solid";
    color: string;
  } | {
    type: "linearGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
    angleDeg?: number;
  } | {
    type: "radialGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
  } | {
    type: "pattern";
    pattern: string;
    backgroundColor: string;
    foregroundColor: string;
  } | {
    type?: "raw";
    raw: unknown;
  });
  plotAreaLine?: {
    type?: "solid" | "none";
    enabled?: boolean;
    widthPt?: number;
    color?: string;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    raw?: unknown;
  };
  titleFill?: ({
    type: "none";
  } | {
    type: "solid";
    color: string;
  } | {
    type: "linearGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
    angleDeg?: number;
  } | {
    type: "radialGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
  } | {
    type: "pattern";
    pattern: string;
    backgroundColor: string;
    foregroundColor: string;
  } | {
    type?: "raw";
    raw: unknown;
  });
  titleLine?: {
    type?: "solid" | "none";
    enabled?: boolean;
    widthPt?: number;
    color?: string;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    raw?: unknown;
  };
  legend?: {
    position?: "left" | "top" | "right" | "bottom" | "none";
    fontSize?: number;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    line?: {
      type?: "solid" | "none";
      enabled?: boolean;
      widthPt?: number;
      color?: string;
      fill?: ({
        type: "none";
      } | {
        type: "solid";
        color: string;
      } | {
        type: "linearGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
        angleDeg?: number;
      } | {
        type: "radialGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
      } | {
        type: "pattern";
        pattern: string;
        backgroundColor: string;
        foregroundColor: string;
      } | {
        type?: "raw";
        raw: unknown;
      });
      raw?: unknown;
    };
  };
  horizontalAxis?: {
    title?: string;
    titleFontSize?: number;
    titleBold?: boolean;
    labelsFontSize?: number;
    normalOrder?: boolean;
    majorTickMark?: string;
    minorTickMark?: string;
    tickLabelPosition?: "none" | "nextTo" | "low" | "high";
    numberFormat?: string;
    position?: string;
  };
  verticalAxis?: {
    title?: string;
    titleFontSize?: number;
    titleBold?: boolean;
    labelsFontSize?: number;
    normalOrder?: boolean;
    majorTickMark?: string;
    minorTickMark?: string;
    tickLabelPosition?: "none" | "nextTo" | "low" | "high";
    numberFormat?: string;
    position?: string;
  };
  dataLabels?: {
    showSeriesName?: boolean;
    showCategoryName?: boolean;
    showValue?: boolean;
    showPercent?: boolean;
  };
  gridlines?: {
    majorHorizontal?: {
      line: {
        type?: "solid" | "none";
        enabled?: boolean;
        widthPt?: number;
        color?: string;
        fill?: ({
          type: "none";
        } | {
          type: "solid";
          color: string;
        } | {
          type: "linearGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
          angleDeg?: number;
        } | {
          type: "radialGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
        } | {
          type: "pattern";
          pattern: string;
          backgroundColor: string;
          foregroundColor: string;
        } | {
          type?: "raw";
          raw: unknown;
        });
        raw?: unknown;
      };
    };
    minorHorizontal?: {
      line: {
        type?: "solid" | "none";
        enabled?: boolean;
        widthPt?: number;
        color?: string;
        fill?: ({
          type: "none";
        } | {
          type: "solid";
          color: string;
        } | {
          type: "linearGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
          angleDeg?: number;
        } | {
          type: "radialGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
        } | {
          type: "pattern";
          pattern: string;
          backgroundColor: string;
          foregroundColor: string;
        } | {
          type?: "raw";
          raw: unknown;
        });
        raw?: unknown;
      };
    };
    majorVertical?: {
      line: {
        type?: "solid" | "none";
        enabled?: boolean;
        widthPt?: number;
        color?: string;
        fill?: ({
          type: "none";
        } | {
          type: "solid";
          color: string;
        } | {
          type: "linearGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
          angleDeg?: number;
        } | {
          type: "radialGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
        } | {
          type: "pattern";
          pattern: string;
          backgroundColor: string;
          foregroundColor: string;
        } | {
          type?: "raw";
          raw: unknown;
        });
        raw?: unknown;
      };
    };
    minorVertical?: {
      line: {
        type?: "solid" | "none";
        enabled?: boolean;
        widthPt?: number;
        color?: string;
        fill?: ({
          type: "none";
        } | {
          type: "solid";
          color: string;
        } | {
          type: "linearGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
          angleDeg?: number;
        } | {
          type: "radialGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
        } | {
          type: "pattern";
          pattern: string;
          backgroundColor: string;
          foregroundColor: string;
        } | {
          type?: "raw";
          raw: unknown;
        });
        raw?: unknown;
      };
    };
  };
  seriesUpdates?: Array<{
    index: number;
    type?: "bar" | "barStacked" | "barStackedPercent" | "bar3D" | "barStacked3D" | "barStackedPercent3D" | "barStackedPercent3DPerspective" | "horizontalBar" | "horizontalBarStacked" | "horizontalBarStackedPercent" | "horizontalBar3D" | "horizontalBarStacked3D" | "horizontalBarStackedPercent3D" | "lineNormal" | "lineStacked" | "lineStackedPercent" | "lineNormalMarker" | "lineStackedMarker" | "lineStackedPerMarker" | "line3D" | "pie" | "pie3D" | "doughnut" | "scatter" | "scatterLine" | "scatterLineMarker" | "scatterSmooth" | "scatterSmoothMarker" | "stock" | "area" | "areaStacked" | "areaStackedPercent" | "comboCustom" | "comboBarLine" | "comboBarLineSecondary" | "radar" | "radarMarker" | "radarFilled" | "line" | "lineMarker" | "stackedBar" | "stackedBarPercent" | "stackedLine" | "stackedLinePercent" | "column";
    name?: string;
    values?: Array<number>;
    valuesRange?: string;
    xValues?: Array<number>;
    xValuesRange?: string;
    numberFormat?: string;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    line?: {
      type?: "solid" | "none";
      enabled?: boolean;
      widthPt?: number;
      color?: string;
      fill?: ({
        type: "none";
      } | {
        type: "solid";
        color: string;
      } | {
        type: "linearGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
        angleDeg?: number;
      } | {
        type: "radialGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
      } | {
        type: "pattern";
        pattern: string;
        backgroundColor: string;
        foregroundColor: string;
      } | {
        type?: "raw";
        raw: unknown;
      });
      raw?: unknown;
    };
    allSeries?: boolean;
    points?: Array<{
      index: number;
      fill?: ({
        type: "none";
      } | {
        type: "solid";
        color: string;
      } | {
        type: "linearGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
        angleDeg?: number;
      } | {
        type: "radialGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
      } | {
        type: "pattern";
        pattern: string;
        backgroundColor: string;
        foregroundColor: string;
      } | {
        type?: "raw";
        raw: unknown;
      });
      line?: {
        type?: "solid" | "none";
        enabled?: boolean;
        widthPt?: number;
        color?: string;
        fill?: ({
          type: "none";
        } | {
          type: "solid";
          color: string;
        } | {
          type: "linearGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
          angleDeg?: number;
        } | {
          type: "radialGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
        } | {
          type: "pattern";
          pattern: string;
          backgroundColor: string;
          foregroundColor: string;
        } | {
          type?: "raw";
          raw: unknown;
        });
        raw?: unknown;
      };
      markerFill?: ({
        type: "none";
      } | {
        type: "solid";
        color: string;
      } | {
        type: "linearGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
        angleDeg?: number;
      } | {
        type: "radialGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
      } | {
        type: "pattern";
        pattern: string;
        backgroundColor: string;
        foregroundColor: string;
      } | {
        type?: "raw";
        raw: unknown;
      });
      markerLine?: {
        type?: "solid" | "none";
        enabled?: boolean;
        widthPt?: number;
        color?: string;
        fill?: ({
          type: "none";
        } | {
          type: "solid";
          color: string;
        } | {
          type: "linearGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
          angleDeg?: number;
        } | {
          type: "radialGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
        } | {
          type: "pattern";
          pattern: string;
          backgroundColor: string;
          foregroundColor: string;
        } | {
          type?: "raw";
          raw: unknown;
        });
        raw?: unknown;
      };
      numberFormat?: string;
      allSeries?: boolean;
      allMarkers?: boolean;
      dataLabels?: {
        showSeriesName?: boolean;
        showCategoryName?: boolean;
        showValue?: boolean;
        showPercent?: boolean;
      };
    }>;
  }>;
  removeSeries?: Array<number>;
  includeRaw?: boolean;
} & ({ holeSizePercent: number; type: "doughnut"; } | { holeSizePercent?: never; });
export type AiBridgeGeneratedSlidesUpdateChartArgs = {
  slide: number;
  chartId?: string;
  chartIndex?: number;
  name?: string;
  categories?: Array<string | number>;
  xMm?: number;
  yMm?: number;
  widthMm?: number;
  heightMm?: number;
  style?: number;
  holeSizePercent?: number;
  title?: string;
  titleFontSize?: number;
  titleBold?: boolean;
  rotationDeg?: number;
  fill?: ({
    type: "none";
  } | {
    type: "solid";
    color: string;
  } | {
    type: "linearGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
    angleDeg?: number;
  } | {
    type: "radialGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
  } | {
    type: "pattern";
    pattern: string;
    backgroundColor: string;
    foregroundColor: string;
  } | {
    type?: "raw";
    raw: unknown;
  });
  line?: {
    type?: "solid" | "none";
    enabled?: boolean;
    widthPt?: number;
    color?: string;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    raw?: unknown;
  };
  plotAreaFill?: ({
    type: "none";
  } | {
    type: "solid";
    color: string;
  } | {
    type: "linearGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
    angleDeg?: number;
  } | {
    type: "radialGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
  } | {
    type: "pattern";
    pattern: string;
    backgroundColor: string;
    foregroundColor: string;
  } | {
    type?: "raw";
    raw: unknown;
  });
  plotAreaLine?: {
    type?: "solid" | "none";
    enabled?: boolean;
    widthPt?: number;
    color?: string;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    raw?: unknown;
  };
  titleFill?: ({
    type: "none";
  } | {
    type: "solid";
    color: string;
  } | {
    type: "linearGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
    angleDeg?: number;
  } | {
    type: "radialGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
  } | {
    type: "pattern";
    pattern: string;
    backgroundColor: string;
    foregroundColor: string;
  } | {
    type?: "raw";
    raw: unknown;
  });
  titleLine?: {
    type?: "solid" | "none";
    enabled?: boolean;
    widthPt?: number;
    color?: string;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    raw?: unknown;
  };
  legend?: {
    position?: "left" | "top" | "right" | "bottom" | "none";
    fontSize?: number;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    line?: {
      type?: "solid" | "none";
      enabled?: boolean;
      widthPt?: number;
      color?: string;
      fill?: ({
        type: "none";
      } | {
        type: "solid";
        color: string;
      } | {
        type: "linearGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
        angleDeg?: number;
      } | {
        type: "radialGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
      } | {
        type: "pattern";
        pattern: string;
        backgroundColor: string;
        foregroundColor: string;
      } | {
        type?: "raw";
        raw: unknown;
      });
      raw?: unknown;
    };
  };
  horizontalAxis?: {
    title?: string;
    titleFontSize?: number;
    titleBold?: boolean;
    labelsFontSize?: number;
    normalOrder?: boolean;
    majorTickMark?: string;
    minorTickMark?: string;
    tickLabelPosition?: "none" | "nextTo" | "low" | "high";
    numberFormat?: string;
    position?: string;
  };
  verticalAxis?: {
    title?: string;
    titleFontSize?: number;
    titleBold?: boolean;
    labelsFontSize?: number;
    normalOrder?: boolean;
    majorTickMark?: string;
    minorTickMark?: string;
    tickLabelPosition?: "none" | "nextTo" | "low" | "high";
    numberFormat?: string;
    position?: string;
  };
  dataLabels?: {
    showSeriesName?: boolean;
    showCategoryName?: boolean;
    showValue?: boolean;
    showPercent?: boolean;
  };
  gridlines?: {
    majorHorizontal?: {
      line: {
        type?: "solid" | "none";
        enabled?: boolean;
        widthPt?: number;
        color?: string;
        fill?: ({
          type: "none";
        } | {
          type: "solid";
          color: string;
        } | {
          type: "linearGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
          angleDeg?: number;
        } | {
          type: "radialGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
        } | {
          type: "pattern";
          pattern: string;
          backgroundColor: string;
          foregroundColor: string;
        } | {
          type?: "raw";
          raw: unknown;
        });
        raw?: unknown;
      };
    };
    minorHorizontal?: {
      line: {
        type?: "solid" | "none";
        enabled?: boolean;
        widthPt?: number;
        color?: string;
        fill?: ({
          type: "none";
        } | {
          type: "solid";
          color: string;
        } | {
          type: "linearGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
          angleDeg?: number;
        } | {
          type: "radialGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
        } | {
          type: "pattern";
          pattern: string;
          backgroundColor: string;
          foregroundColor: string;
        } | {
          type?: "raw";
          raw: unknown;
        });
        raw?: unknown;
      };
    };
    majorVertical?: {
      line: {
        type?: "solid" | "none";
        enabled?: boolean;
        widthPt?: number;
        color?: string;
        fill?: ({
          type: "none";
        } | {
          type: "solid";
          color: string;
        } | {
          type: "linearGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
          angleDeg?: number;
        } | {
          type: "radialGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
        } | {
          type: "pattern";
          pattern: string;
          backgroundColor: string;
          foregroundColor: string;
        } | {
          type?: "raw";
          raw: unknown;
        });
        raw?: unknown;
      };
    };
    minorVertical?: {
      line: {
        type?: "solid" | "none";
        enabled?: boolean;
        widthPt?: number;
        color?: string;
        fill?: ({
          type: "none";
        } | {
          type: "solid";
          color: string;
        } | {
          type: "linearGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
          angleDeg?: number;
        } | {
          type: "radialGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
        } | {
          type: "pattern";
          pattern: string;
          backgroundColor: string;
          foregroundColor: string;
        } | {
          type?: "raw";
          raw: unknown;
        });
        raw?: unknown;
      };
    };
  };
  seriesUpdates?: Array<{
    index: number;
    type?: "bar" | "barStacked" | "barStackedPercent" | "bar3D" | "barStacked3D" | "barStackedPercent3D" | "barStackedPercent3DPerspective" | "horizontalBar" | "horizontalBarStacked" | "horizontalBarStackedPercent" | "horizontalBar3D" | "horizontalBarStacked3D" | "horizontalBarStackedPercent3D" | "lineNormal" | "lineStacked" | "lineStackedPercent" | "lineNormalMarker" | "lineStackedMarker" | "lineStackedPerMarker" | "line3D" | "pie" | "pie3D" | "doughnut" | "scatter" | "scatterLine" | "scatterLineMarker" | "scatterSmooth" | "scatterSmoothMarker" | "stock" | "area" | "areaStacked" | "areaStackedPercent" | "comboCustom" | "comboBarLine" | "comboBarLineSecondary" | "radar" | "radarMarker" | "radarFilled" | "line" | "lineMarker" | "stackedBar" | "stackedBarPercent" | "stackedLine" | "stackedLinePercent" | "column";
    name?: string;
    values?: Array<number>;
    valuesRange?: string;
    xValues?: Array<number>;
    xValuesRange?: string;
    numberFormat?: string;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    line?: {
      type?: "solid" | "none";
      enabled?: boolean;
      widthPt?: number;
      color?: string;
      fill?: ({
        type: "none";
      } | {
        type: "solid";
        color: string;
      } | {
        type: "linearGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
        angleDeg?: number;
      } | {
        type: "radialGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
      } | {
        type: "pattern";
        pattern: string;
        backgroundColor: string;
        foregroundColor: string;
      } | {
        type?: "raw";
        raw: unknown;
      });
      raw?: unknown;
    };
    allSeries?: boolean;
    points?: Array<{
      index: number;
      fill?: ({
        type: "none";
      } | {
        type: "solid";
        color: string;
      } | {
        type: "linearGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
        angleDeg?: number;
      } | {
        type: "radialGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
      } | {
        type: "pattern";
        pattern: string;
        backgroundColor: string;
        foregroundColor: string;
      } | {
        type?: "raw";
        raw: unknown;
      });
      line?: {
        type?: "solid" | "none";
        enabled?: boolean;
        widthPt?: number;
        color?: string;
        fill?: ({
          type: "none";
        } | {
          type: "solid";
          color: string;
        } | {
          type: "linearGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
          angleDeg?: number;
        } | {
          type: "radialGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
        } | {
          type: "pattern";
          pattern: string;
          backgroundColor: string;
          foregroundColor: string;
        } | {
          type?: "raw";
          raw: unknown;
        });
        raw?: unknown;
      };
      markerFill?: ({
        type: "none";
      } | {
        type: "solid";
        color: string;
      } | {
        type: "linearGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
        angleDeg?: number;
      } | {
        type: "radialGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
      } | {
        type: "pattern";
        pattern: string;
        backgroundColor: string;
        foregroundColor: string;
      } | {
        type?: "raw";
        raw: unknown;
      });
      markerLine?: {
        type?: "solid" | "none";
        enabled?: boolean;
        widthPt?: number;
        color?: string;
        fill?: ({
          type: "none";
        } | {
          type: "solid";
          color: string;
        } | {
          type: "linearGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
          angleDeg?: number;
        } | {
          type: "radialGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
        } | {
          type: "pattern";
          pattern: string;
          backgroundColor: string;
          foregroundColor: string;
        } | {
          type?: "raw";
          raw: unknown;
        });
        raw?: unknown;
      };
      numberFormat?: string;
      allSeries?: boolean;
      allMarkers?: boolean;
      dataLabels?: {
        showSeriesName?: boolean;
        showCategoryName?: boolean;
        showValue?: boolean;
        showPercent?: boolean;
      };
    }>;
  }>;
  removeSeries?: Array<number>;
  includeRaw?: boolean;
} & ({ chartId: string; } | { chartIndex: number; } | { name: string; });
export type AiBridgeGeneratedSlidesDeleteChartArgs = {
  slide: number;
  chartId?: string;
  chartIndex?: number;
  name?: string;
} & ({ chartId: string; } | { chartIndex: number; } | { name: string; });
export type AiBridgeGeneratedSlidesInspectMacrosArgs = {
  kind?: "office" | "vba";
};
export type AiBridgeGeneratedSlidesSetMacrosArgs = {
  content: Record<string, unknown>;
};
export type AiBridgeGeneratedSlidesControlSlideshowArgs = {
  action: "start" | "end" | "pause" | "resume" | "next" | "previous" | "goto";
  slide?: number;
};
export type AiBridgeGeneratedSheetsInspectArgs = {
  maxCells?: number;
};
export type AiBridgeGeneratedSheetsSetValuesArgs = {
  sheet?: string;
  range: string;
  values: (string | number | boolean | null | Array<Array<string | number | boolean | null>>);
};
export type AiBridgeGeneratedSheetsSetFormulaArgs = {
  sheet?: string;
  range: string;
  formula: (string | Array<Array<string>>);
};
export type AiBridgeGeneratedSheetsInspectRangeArgs = {
  sheet?: string;
  range: string;
  includeValues?: boolean;
  includeFormat?: boolean;
  includeConditionalFormats?: boolean;
  includeValidation?: boolean;
};
export type AiBridgeGeneratedSheetsSetArrayFormulaArgs = {
  sheet?: string;
  range: string;
  formula: string;
};
export type AiBridgeGeneratedSheetsReplaceTextArgs = {
  sheet?: string;
  range?: string;
  search: string;
  replace: string;
};
export type AiBridgeGeneratedSheetsFormatRangeArgs = {
  sheet?: string;
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
  columnWidthChars?: number;
  rowHeight?: number;
  rowHeightPt?: number;
  borders?: Array<{
    side: string;
    style: string;
    color: string;
  }>;
} & ({ fontSize: number; } | { fontName: string; } | { bold: boolean; } | { italic: boolean; } | { underline: boolean; } | { strikeout: boolean; } | { fontColor: string; } | { fillColor: string; } | { horizontalAlign: string; } | { verticalAlign: string; } | { numberFormat: string; } | { wrap: boolean; } | { orientation: string | number; } | { columnWidth: number; } | { columnWidthChars: number; } | { rowHeight: number; } | { rowHeightPt: number; } | { borders: Array<{
  side: string;
  style: string;
  color: string;
}>; });
export type AiBridgeGeneratedSheetsAddSheetArgs = {
  name: string;
};
export type AiBridgeGeneratedSheetsRenameSheetArgs = {
  sheet?: string;
  newName: string;
};
export type AiBridgeGeneratedSheetsDeleteSheetArgs = {
  sheet: string;
};
export type AiBridgeGeneratedSheetsManageSheetArgs = {
  action: "setVisibility" | "setActive" | "moveBefore" | "copy";
  sheet?: string;
  visible?: boolean;
  beforeSheet?: string;
  newName?: string;
};
export type AiBridgeGeneratedSheetsManageRangeArgs = {
  action: "merge" | "unmerge" | "insert" | "delete" | "copy" | "cut" | "clear" | "clearContents" | "clearFormats" | "clearHyperlinks" | "autoFit" | "setHidden" | "fillDown" | "fillUp" | "fillLeft" | "fillRight" | "select";
  sheet?: string;
  range: string;
  destinationSheet?: string;
  destination?: string;
  shift?: string;
  across?: boolean;
  hidden?: boolean;
  rows?: boolean;
  columns?: boolean;
};
export type AiBridgeGeneratedSheetsSetRichTextArgs = {
  sheet?: string;
  range: string;
  text?: string;
  runs: Array<{
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
  }>;
};
export type AiBridgeGeneratedSheetsInspectNamesArgs = Record<string, never>;
export type AiBridgeGeneratedSheetsManageNamesArgs = {
  action: "add" | "update" | "delete";
  name: string;
  newName?: string;
  refersTo?: string;
};
export type AiBridgeGeneratedSheetsRecalculateArgs = {
  mode?: "formulas" | "pivots" | "all";
};
export type AiBridgeGeneratedSheetsSortArgs = {
  sheet?: string;
  range: string;
  keys: Array<{
    range: string;
    order?: "xlAscending" | "xlDescending";
  }>;
  header?: (boolean | "yes" | "no" | "guess" | "xlYes" | "xlNo" | "xlGuess");
  orientation?: "rows" | "columns" | "xlSortRows" | "xlSortColumns";
};
export type AiBridgeGeneratedSheetsFilterArgs = {
  action: "set" | "showAll" | "reapply";
  sheet?: string;
  range?: string;
  field?: number;
  criteria1?: unknown;
  operator?: "and" | "or" | "filterValues" | "values" | "top10Items" | "bottom10Items" | "top10Percent" | "bottom10Percent" | "filterCellColor" | "filterFontColor" | "filterIcon" | "dynamic" | "xlAnd" | "xlOr" | "xlFilterValues" | "xlTop10Items" | "xlBottom10Items" | "xlTop10Percent" | "xlBottom10Percent" | "xlFilterCellColor" | "xlFilterFontColor" | "xlFilterIcon" | "xlFilterDynamic";
  criteria2?: unknown;
  visibleDropDown?: boolean;
};
export type AiBridgeGeneratedSheetsInspectTablesArgs = {
  sheet?: string;
  tableIndex?: number;
  name?: string;
  maxTables?: number;
};
export type AiBridgeGeneratedSheetsManageTableArgs = {
  action: "create" | "format" | "update" | "resize" | "delete" | "unlist";
  sheet?: string;
  range?: string;
  tableMode?: "auto" | "structured" | "rangeStyle" | "basic";
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
};
export type AiBridgeGeneratedSheetsManageConditionalFormatArgs = {
  action: "add" | "deleteAll";
  sheet?: string;
  range: string;
  type?: "cellValue" | "expression" | "uniqueValues" | "duplicateValues" | "colorScale" | "dataBar" | "iconSet" | "top10" | "aboveAverage";
  operator?: "between" | "notBetween" | "equal" | "notEqual" | "greaterThan" | "lessThan" | "greaterThanOrEqual" | "lessThanOrEqual" | "xlBetween" | "xlNotBetween" | "xlEqual" | "xlNotEqual" | "xlGreater" | "xlLess" | "xlGreaterEqual" | "xlLessEqual";
  formula1?: unknown;
  formula2?: unknown;
  scale?: number;
  fillColor?: string;
  fontColor?: string;
  bold?: boolean;
};
export type AiBridgeGeneratedSheetsManageValidationArgs = {
  action: "add" | "modify" | "delete";
  sheet?: string;
  range: string;
  type?: "inputOnly" | "wholeNumber" | "decimal" | "list" | "date" | "time" | "textLength" | "custom" | "xlValidateInputOnly" | "xlValidateWholeNumber" | "xlValidateDecimal" | "xlValidateList" | "xlValidateDate" | "xlValidateTime" | "xlValidateTextLength" | "xlValidateCustom";
  alertStyle?: "stop" | "warning" | "information" | "info" | "xlValidAlertStop" | "xlValidAlertWarning" | "xlValidAlertInformation";
  operator?: "between" | "notBetween" | "equal" | "notEqual" | "greaterThan" | "lessThan" | "greaterThanOrEqual" | "lessThanOrEqual" | "xlBetween" | "xlNotBetween" | "xlEqual" | "xlNotEqual" | "xlGreater" | "xlLess" | "xlGreaterEqual" | "xlLessEqual";
  formula1?: unknown;
  formula2?: unknown;
  ignoreBlank?: boolean;
  showInput?: boolean;
  showError?: boolean;
  inputTitle?: string;
  inputMessage?: string;
  errorTitle?: string;
  errorMessage?: string;
};
export type AiBridgeGeneratedSheetsInspectPivotsArgs = Record<string, never>;
export type AiBridgeGeneratedSheetsManagePivotArgs = {
  action: "createNewSheet" | "createExisting" | "update" | "refresh" | "refreshAll" | "clear";
  sheet?: string;
  name?: string;
  sourceSheet?: string;
  sourceRange?: string;
  destinationSheet?: string;
  destinationRange?: string;
  fields?: Record<string, unknown>;
  dataFields?: Array<(string | {
    name: string;
    function?: string;
  })>;
  style?: string;
  title?: string;
  description?: string;
  rowGrand?: boolean;
  columnGrand?: boolean;
};
export type AiBridgeGeneratedSheetsInspectDrawingsArgs = {
  sheet?: string;
  maxDrawings?: number;
  includeRaw?: boolean;
};
export type AiBridgeGeneratedSheetsManageDrawingArgs = {
  action: "addImage" | "addShape" | "addTextBox" | "addOleObject" | "update" | "delete" | "copy";
  sheet?: string;
  destinationSheet?: string;
  drawingIndex?: number;
  name?: string;
  source?: ({
    type: "url";
    url: string;
  } | {
    type: "dataUrl";
    dataUrl: string;
  } | {
    type: "relayAsset";
    path: string;
    assetToken: string;
    mimeType: "image/png" | "image/jpeg" | "image/gif" | "image/webp" | "image/svg+xml";
    widthPx: number;
    heightPx: number;
    assetId: string;
  });
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
  fill?: ({
    type: "none";
  } | {
    type: "solid";
    color: string;
  } | {
    type: "linearGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
    angleDeg?: number;
  } | {
    type: "radialGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
  } | {
    type: "pattern";
    pattern: string;
    backgroundColor: string;
    foregroundColor: string;
  } | {
    type?: "raw";
    raw: unknown;
  });
  line?: {
    type?: "solid" | "none";
    enabled?: boolean;
    widthPt?: number;
    color?: string;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    raw?: unknown;
  };
  includeRaw?: boolean;
};
export type AiBridgeGeneratedSheetsManageHyperlinkArgs = {
  action: "set" | "delete";
  sheet?: string;
  range: string;
  url?: string;
  location?: string;
  displayText?: string;
  tooltip?: string;
};
export type AiBridgeGeneratedSheetsInspectCommentsArgs = {
  sheet?: string;
};
export type AiBridgeGeneratedSheetsManageCommentsArgs = {
  action: "add" | "update" | "delete" | "setSolved" | "addReply" | "removeReplies";
  sheet?: string;
  range?: string;
  commentId?: string;
  text?: string;
  author?: string;
  userId?: string;
  solved?: boolean;
  start?: number;
  count?: number;
  removeAll?: boolean;
};
export type AiBridgeGeneratedSheetsInspectFreezePanesArgs = {
  sheet?: string;
};
export type AiBridgeGeneratedSheetsManageFreezePanesArgs = {
  action: "unfreeze" | "freezeAt" | "freezeRows" | "freezeColumns";
  sheet?: string;
  range?: string;
  count?: number;
} & ({
  action?: "unfreeze";
} | {
  action?: "freezeAt";
} | {
  action?: "freezeRows";
} | {
  action?: "freezeColumns";
});
export type AiBridgeGeneratedSheetsInspectPropertiesArgs = {
  customNames?: Array<string>;
};
export type AiBridgeGeneratedSheetsManagePropertiesArgs = {
  core?: {
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
  };
  custom?: Record<string, unknown>;
};
export type AiBridgeGeneratedSheetsInspectProtectedRangesArgs = {
  sheet?: string;
};
export type AiBridgeGeneratedSheetsManageProtectedRangesArgs = {
  action: "add" | "update" | "addUser" | "deleteUser";
  sheet?: string;
  title: string;
  newTitle?: string;
  range?: string;
  anyoneType?: string;
  userId?: string;
  userName?: string;
  permission?: string;
};
export type AiBridgeGeneratedSheetsInspectPageLayoutArgs = {
  sheet?: string;
};
export type AiBridgeGeneratedSheetsManagePageLayoutArgs = {
  sheet?: string;
  orientation?: "portrait" | "landscape" | "xlPortrait" | "xlLandscape";
  margins?: {
    top?: number;
    right?: number;
    bottom?: number;
    left?: number;
  };
  marginsPt?: {
    top?: number;
    right?: number;
    bottom?: number;
    left?: number;
  };
  printGridlines?: boolean;
  printHeadings?: boolean;
  displayGridlines?: boolean;
  displayHeadings?: boolean;
} & ({ orientation: "portrait" | "landscape" | "xlPortrait" | "xlLandscape"; } | { margins: {
  top?: number;
  right?: number;
  bottom?: number;
  left?: number;
}; } | { marginsPt: {
  top?: number;
  right?: number;
  bottom?: number;
  left?: number;
}; } | { printGridlines: boolean; } | { printHeadings: boolean; } | { displayGridlines: boolean; } | { displayHeadings: boolean; });
export type AiBridgeGeneratedSheetsInspectMacrosArgs = {
  kind?: "office" | "vba";
};
export type AiBridgeGeneratedSheetsSetMacrosArgs = {
  content: Record<string, unknown>;
};
export type AiBridgeGeneratedSheetsAddChartArgs = {
  sheet?: string;
  range: string;
  type?: "bar" | "barStacked" | "barStackedPercent" | "bar3D" | "barStacked3D" | "barStackedPercent3D" | "barStackedPercent3DPerspective" | "horizontalBar" | "horizontalBarStacked" | "horizontalBarStackedPercent" | "horizontalBar3D" | "horizontalBarStacked3D" | "horizontalBarStackedPercent3D" | "lineNormal" | "lineStacked" | "lineStackedPercent" | "lineNormalMarker" | "lineStackedMarker" | "lineStackedPerMarker" | "line3D" | "pie" | "pie3D" | "doughnut" | "scatter" | "scatterLine" | "scatterLineMarker" | "scatterSmooth" | "scatterSmoothMarker" | "stock" | "area" | "areaStacked" | "areaStackedPercent" | "comboCustom" | "comboBarLine" | "comboBarLineSecondary" | "radar" | "radarMarker" | "radarFilled" | "line" | "lineMarker" | "stackedBar" | "stackedBarPercent" | "stackedLine" | "stackedLinePercent" | "column";
  title?: string;
  titleFontSize?: number;
  titleBold?: boolean;
  inRows?: boolean;
  style?: number;
  widthMm?: number;
  heightMm?: number;
  fromColumn?: number;
  fromRow?: number;
  columnOffsetMm?: number;
  rowOffsetMm?: number;
  categoryRange?: string;
  addSeries?: Array<{
    name: string;
    valuesRange: string;
    xValuesRange?: string;
  }>;
  rotationDeg?: number;
  name?: string;
  fill?: ({
    type: "none";
  } | {
    type: "solid";
    color: string;
  } | {
    type: "linearGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
    angleDeg?: number;
  } | {
    type: "radialGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
  } | {
    type: "pattern";
    pattern: string;
    backgroundColor: string;
    foregroundColor: string;
  } | {
    type?: "raw";
    raw: unknown;
  });
  line?: {
    type?: "solid" | "none";
    enabled?: boolean;
    widthPt?: number;
    color?: string;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    raw?: unknown;
  };
  plotAreaFill?: ({
    type: "none";
  } | {
    type: "solid";
    color: string;
  } | {
    type: "linearGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
    angleDeg?: number;
  } | {
    type: "radialGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
  } | {
    type: "pattern";
    pattern: string;
    backgroundColor: string;
    foregroundColor: string;
  } | {
    type?: "raw";
    raw: unknown;
  });
  plotAreaLine?: {
    type?: "solid" | "none";
    enabled?: boolean;
    widthPt?: number;
    color?: string;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    raw?: unknown;
  };
  titleFill?: ({
    type: "none";
  } | {
    type: "solid";
    color: string;
  } | {
    type: "linearGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
    angleDeg?: number;
  } | {
    type: "radialGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
  } | {
    type: "pattern";
    pattern: string;
    backgroundColor: string;
    foregroundColor: string;
  } | {
    type?: "raw";
    raw: unknown;
  });
  titleLine?: {
    type?: "solid" | "none";
    enabled?: boolean;
    widthPt?: number;
    color?: string;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    raw?: unknown;
  };
  legend?: {
    position?: "left" | "top" | "right" | "bottom" | "none";
    fontSize?: number;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    line?: {
      type?: "solid" | "none";
      enabled?: boolean;
      widthPt?: number;
      color?: string;
      fill?: ({
        type: "none";
      } | {
        type: "solid";
        color: string;
      } | {
        type: "linearGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
        angleDeg?: number;
      } | {
        type: "radialGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
      } | {
        type: "pattern";
        pattern: string;
        backgroundColor: string;
        foregroundColor: string;
      } | {
        type?: "raw";
        raw: unknown;
      });
      raw?: unknown;
    };
  };
  horizontalAxis?: {
    title?: string;
    titleFontSize?: number;
    titleBold?: boolean;
    labelsFontSize?: number;
    normalOrder?: boolean;
    majorTickMark?: string;
    minorTickMark?: string;
    tickLabelPosition?: "none" | "nextTo" | "low" | "high";
    numberFormat?: string;
    position?: string;
  };
  verticalAxis?: {
    title?: string;
    titleFontSize?: number;
    titleBold?: boolean;
    labelsFontSize?: number;
    normalOrder?: boolean;
    majorTickMark?: string;
    minorTickMark?: string;
    tickLabelPosition?: "none" | "nextTo" | "low" | "high";
    numberFormat?: string;
    position?: string;
  };
  dataLabels?: {
    showSeriesName?: boolean;
    showCategoryName?: boolean;
    showValue?: boolean;
    showPercent?: boolean;
  };
  gridlines?: {
    majorHorizontal?: {
      line: {
        type?: "solid" | "none";
        enabled?: boolean;
        widthPt?: number;
        color?: string;
        fill?: ({
          type: "none";
        } | {
          type: "solid";
          color: string;
        } | {
          type: "linearGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
          angleDeg?: number;
        } | {
          type: "radialGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
        } | {
          type: "pattern";
          pattern: string;
          backgroundColor: string;
          foregroundColor: string;
        } | {
          type?: "raw";
          raw: unknown;
        });
        raw?: unknown;
      };
    };
    minorHorizontal?: {
      line: {
        type?: "solid" | "none";
        enabled?: boolean;
        widthPt?: number;
        color?: string;
        fill?: ({
          type: "none";
        } | {
          type: "solid";
          color: string;
        } | {
          type: "linearGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
          angleDeg?: number;
        } | {
          type: "radialGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
        } | {
          type: "pattern";
          pattern: string;
          backgroundColor: string;
          foregroundColor: string;
        } | {
          type?: "raw";
          raw: unknown;
        });
        raw?: unknown;
      };
    };
    majorVertical?: {
      line: {
        type?: "solid" | "none";
        enabled?: boolean;
        widthPt?: number;
        color?: string;
        fill?: ({
          type: "none";
        } | {
          type: "solid";
          color: string;
        } | {
          type: "linearGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
          angleDeg?: number;
        } | {
          type: "radialGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
        } | {
          type: "pattern";
          pattern: string;
          backgroundColor: string;
          foregroundColor: string;
        } | {
          type?: "raw";
          raw: unknown;
        });
        raw?: unknown;
      };
    };
    minorVertical?: {
      line: {
        type?: "solid" | "none";
        enabled?: boolean;
        widthPt?: number;
        color?: string;
        fill?: ({
          type: "none";
        } | {
          type: "solid";
          color: string;
        } | {
          type: "linearGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
          angleDeg?: number;
        } | {
          type: "radialGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
        } | {
          type: "pattern";
          pattern: string;
          backgroundColor: string;
          foregroundColor: string;
        } | {
          type?: "raw";
          raw: unknown;
        });
        raw?: unknown;
      };
    };
  };
  seriesUpdates?: Array<{
    index: number;
    type?: "bar" | "barStacked" | "barStackedPercent" | "bar3D" | "barStacked3D" | "barStackedPercent3D" | "barStackedPercent3DPerspective" | "horizontalBar" | "horizontalBarStacked" | "horizontalBarStackedPercent" | "horizontalBar3D" | "horizontalBarStacked3D" | "horizontalBarStackedPercent3D" | "lineNormal" | "lineStacked" | "lineStackedPercent" | "lineNormalMarker" | "lineStackedMarker" | "lineStackedPerMarker" | "line3D" | "pie" | "pie3D" | "doughnut" | "scatter" | "scatterLine" | "scatterLineMarker" | "scatterSmooth" | "scatterSmoothMarker" | "stock" | "area" | "areaStacked" | "areaStackedPercent" | "comboCustom" | "comboBarLine" | "comboBarLineSecondary" | "radar" | "radarMarker" | "radarFilled" | "line" | "lineMarker" | "stackedBar" | "stackedBarPercent" | "stackedLine" | "stackedLinePercent" | "column";
    name?: string;
    values?: Array<number>;
    valuesRange?: string;
    xValues?: Array<number>;
    xValuesRange?: string;
    numberFormat?: string;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    line?: {
      type?: "solid" | "none";
      enabled?: boolean;
      widthPt?: number;
      color?: string;
      fill?: ({
        type: "none";
      } | {
        type: "solid";
        color: string;
      } | {
        type: "linearGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
        angleDeg?: number;
      } | {
        type: "radialGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
      } | {
        type: "pattern";
        pattern: string;
        backgroundColor: string;
        foregroundColor: string;
      } | {
        type?: "raw";
        raw: unknown;
      });
      raw?: unknown;
    };
    allSeries?: boolean;
    points?: Array<{
      index: number;
      fill?: ({
        type: "none";
      } | {
        type: "solid";
        color: string;
      } | {
        type: "linearGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
        angleDeg?: number;
      } | {
        type: "radialGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
      } | {
        type: "pattern";
        pattern: string;
        backgroundColor: string;
        foregroundColor: string;
      } | {
        type?: "raw";
        raw: unknown;
      });
      line?: {
        type?: "solid" | "none";
        enabled?: boolean;
        widthPt?: number;
        color?: string;
        fill?: ({
          type: "none";
        } | {
          type: "solid";
          color: string;
        } | {
          type: "linearGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
          angleDeg?: number;
        } | {
          type: "radialGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
        } | {
          type: "pattern";
          pattern: string;
          backgroundColor: string;
          foregroundColor: string;
        } | {
          type?: "raw";
          raw: unknown;
        });
        raw?: unknown;
      };
      markerFill?: ({
        type: "none";
      } | {
        type: "solid";
        color: string;
      } | {
        type: "linearGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
        angleDeg?: number;
      } | {
        type: "radialGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
      } | {
        type: "pattern";
        pattern: string;
        backgroundColor: string;
        foregroundColor: string;
      } | {
        type?: "raw";
        raw: unknown;
      });
      markerLine?: {
        type?: "solid" | "none";
        enabled?: boolean;
        widthPt?: number;
        color?: string;
        fill?: ({
          type: "none";
        } | {
          type: "solid";
          color: string;
        } | {
          type: "linearGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
          angleDeg?: number;
        } | {
          type: "radialGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
        } | {
          type: "pattern";
          pattern: string;
          backgroundColor: string;
          foregroundColor: string;
        } | {
          type?: "raw";
          raw: unknown;
        });
        raw?: unknown;
      };
      numberFormat?: string;
      allSeries?: boolean;
      allMarkers?: boolean;
      dataLabels?: {
        showSeriesName?: boolean;
        showCategoryName?: boolean;
        showValue?: boolean;
        showPercent?: boolean;
      };
    }>;
  }>;
  removeSeries?: Array<number>;
  includeRaw?: boolean;
};
export type AiBridgeGeneratedSheetsInspectChartsArgs = {
  sheet?: string;
  maxCharts?: number;
  includeRaw?: boolean;
};
export type AiBridgeGeneratedSheetsUpdateChartArgs = {
  sheet?: string;
  chartIndex?: number;
  name?: string;
  style?: number;
  title?: string;
  titleFontSize?: number;
  titleBold?: boolean;
  rotationDeg?: number;
  widthMm?: number;
  heightMm?: number;
  fromColumn?: number;
  fromRow?: number;
  columnOffsetMm?: number;
  rowOffsetMm?: number;
  categoryRange?: string;
  addSeries?: Array<{
    name: string;
    valuesRange: string;
    xValuesRange?: string;
  }>;
  fill?: ({
    type: "none";
  } | {
    type: "solid";
    color: string;
  } | {
    type: "linearGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
    angleDeg?: number;
  } | {
    type: "radialGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
  } | {
    type: "pattern";
    pattern: string;
    backgroundColor: string;
    foregroundColor: string;
  } | {
    type?: "raw";
    raw: unknown;
  });
  line?: {
    type?: "solid" | "none";
    enabled?: boolean;
    widthPt?: number;
    color?: string;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    raw?: unknown;
  };
  plotAreaFill?: ({
    type: "none";
  } | {
    type: "solid";
    color: string;
  } | {
    type: "linearGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
    angleDeg?: number;
  } | {
    type: "radialGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
  } | {
    type: "pattern";
    pattern: string;
    backgroundColor: string;
    foregroundColor: string;
  } | {
    type?: "raw";
    raw: unknown;
  });
  plotAreaLine?: {
    type?: "solid" | "none";
    enabled?: boolean;
    widthPt?: number;
    color?: string;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    raw?: unknown;
  };
  titleFill?: ({
    type: "none";
  } | {
    type: "solid";
    color: string;
  } | {
    type: "linearGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
    angleDeg?: number;
  } | {
    type: "radialGradient";
    stops: Array<{
      position: number;
      color: string;
    }>;
  } | {
    type: "pattern";
    pattern: string;
    backgroundColor: string;
    foregroundColor: string;
  } | {
    type?: "raw";
    raw: unknown;
  });
  titleLine?: {
    type?: "solid" | "none";
    enabled?: boolean;
    widthPt?: number;
    color?: string;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    raw?: unknown;
  };
  legend?: {
    position?: "left" | "top" | "right" | "bottom" | "none";
    fontSize?: number;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    line?: {
      type?: "solid" | "none";
      enabled?: boolean;
      widthPt?: number;
      color?: string;
      fill?: ({
        type: "none";
      } | {
        type: "solid";
        color: string;
      } | {
        type: "linearGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
        angleDeg?: number;
      } | {
        type: "radialGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
      } | {
        type: "pattern";
        pattern: string;
        backgroundColor: string;
        foregroundColor: string;
      } | {
        type?: "raw";
        raw: unknown;
      });
      raw?: unknown;
    };
  };
  horizontalAxis?: {
    title?: string;
    titleFontSize?: number;
    titleBold?: boolean;
    labelsFontSize?: number;
    normalOrder?: boolean;
    majorTickMark?: string;
    minorTickMark?: string;
    tickLabelPosition?: "none" | "nextTo" | "low" | "high";
    numberFormat?: string;
    position?: string;
  };
  verticalAxis?: {
    title?: string;
    titleFontSize?: number;
    titleBold?: boolean;
    labelsFontSize?: number;
    normalOrder?: boolean;
    majorTickMark?: string;
    minorTickMark?: string;
    tickLabelPosition?: "none" | "nextTo" | "low" | "high";
    numberFormat?: string;
    position?: string;
  };
  dataLabels?: {
    showSeriesName?: boolean;
    showCategoryName?: boolean;
    showValue?: boolean;
    showPercent?: boolean;
  };
  gridlines?: {
    majorHorizontal?: {
      line: {
        type?: "solid" | "none";
        enabled?: boolean;
        widthPt?: number;
        color?: string;
        fill?: ({
          type: "none";
        } | {
          type: "solid";
          color: string;
        } | {
          type: "linearGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
          angleDeg?: number;
        } | {
          type: "radialGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
        } | {
          type: "pattern";
          pattern: string;
          backgroundColor: string;
          foregroundColor: string;
        } | {
          type?: "raw";
          raw: unknown;
        });
        raw?: unknown;
      };
    };
    minorHorizontal?: {
      line: {
        type?: "solid" | "none";
        enabled?: boolean;
        widthPt?: number;
        color?: string;
        fill?: ({
          type: "none";
        } | {
          type: "solid";
          color: string;
        } | {
          type: "linearGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
          angleDeg?: number;
        } | {
          type: "radialGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
        } | {
          type: "pattern";
          pattern: string;
          backgroundColor: string;
          foregroundColor: string;
        } | {
          type?: "raw";
          raw: unknown;
        });
        raw?: unknown;
      };
    };
    majorVertical?: {
      line: {
        type?: "solid" | "none";
        enabled?: boolean;
        widthPt?: number;
        color?: string;
        fill?: ({
          type: "none";
        } | {
          type: "solid";
          color: string;
        } | {
          type: "linearGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
          angleDeg?: number;
        } | {
          type: "radialGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
        } | {
          type: "pattern";
          pattern: string;
          backgroundColor: string;
          foregroundColor: string;
        } | {
          type?: "raw";
          raw: unknown;
        });
        raw?: unknown;
      };
    };
    minorVertical?: {
      line: {
        type?: "solid" | "none";
        enabled?: boolean;
        widthPt?: number;
        color?: string;
        fill?: ({
          type: "none";
        } | {
          type: "solid";
          color: string;
        } | {
          type: "linearGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
          angleDeg?: number;
        } | {
          type: "radialGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
        } | {
          type: "pattern";
          pattern: string;
          backgroundColor: string;
          foregroundColor: string;
        } | {
          type?: "raw";
          raw: unknown;
        });
        raw?: unknown;
      };
    };
  };
  seriesUpdates?: Array<{
    index: number;
    type?: "bar" | "barStacked" | "barStackedPercent" | "bar3D" | "barStacked3D" | "barStackedPercent3D" | "barStackedPercent3DPerspective" | "horizontalBar" | "horizontalBarStacked" | "horizontalBarStackedPercent" | "horizontalBar3D" | "horizontalBarStacked3D" | "horizontalBarStackedPercent3D" | "lineNormal" | "lineStacked" | "lineStackedPercent" | "lineNormalMarker" | "lineStackedMarker" | "lineStackedPerMarker" | "line3D" | "pie" | "pie3D" | "doughnut" | "scatter" | "scatterLine" | "scatterLineMarker" | "scatterSmooth" | "scatterSmoothMarker" | "stock" | "area" | "areaStacked" | "areaStackedPercent" | "comboCustom" | "comboBarLine" | "comboBarLineSecondary" | "radar" | "radarMarker" | "radarFilled" | "line" | "lineMarker" | "stackedBar" | "stackedBarPercent" | "stackedLine" | "stackedLinePercent" | "column";
    name?: string;
    values?: Array<number>;
    valuesRange?: string;
    xValues?: Array<number>;
    xValuesRange?: string;
    numberFormat?: string;
    fill?: ({
      type: "none";
    } | {
      type: "solid";
      color: string;
    } | {
      type: "linearGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
      angleDeg?: number;
    } | {
      type: "radialGradient";
      stops: Array<{
        position: number;
        color: string;
      }>;
    } | {
      type: "pattern";
      pattern: string;
      backgroundColor: string;
      foregroundColor: string;
    } | {
      type?: "raw";
      raw: unknown;
    });
    line?: {
      type?: "solid" | "none";
      enabled?: boolean;
      widthPt?: number;
      color?: string;
      fill?: ({
        type: "none";
      } | {
        type: "solid";
        color: string;
      } | {
        type: "linearGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
        angleDeg?: number;
      } | {
        type: "radialGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
      } | {
        type: "pattern";
        pattern: string;
        backgroundColor: string;
        foregroundColor: string;
      } | {
        type?: "raw";
        raw: unknown;
      });
      raw?: unknown;
    };
    allSeries?: boolean;
    points?: Array<{
      index: number;
      fill?: ({
        type: "none";
      } | {
        type: "solid";
        color: string;
      } | {
        type: "linearGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
        angleDeg?: number;
      } | {
        type: "radialGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
      } | {
        type: "pattern";
        pattern: string;
        backgroundColor: string;
        foregroundColor: string;
      } | {
        type?: "raw";
        raw: unknown;
      });
      line?: {
        type?: "solid" | "none";
        enabled?: boolean;
        widthPt?: number;
        color?: string;
        fill?: ({
          type: "none";
        } | {
          type: "solid";
          color: string;
        } | {
          type: "linearGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
          angleDeg?: number;
        } | {
          type: "radialGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
        } | {
          type: "pattern";
          pattern: string;
          backgroundColor: string;
          foregroundColor: string;
        } | {
          type?: "raw";
          raw: unknown;
        });
        raw?: unknown;
      };
      markerFill?: ({
        type: "none";
      } | {
        type: "solid";
        color: string;
      } | {
        type: "linearGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
        angleDeg?: number;
      } | {
        type: "radialGradient";
        stops: Array<{
          position: number;
          color: string;
        }>;
      } | {
        type: "pattern";
        pattern: string;
        backgroundColor: string;
        foregroundColor: string;
      } | {
        type?: "raw";
        raw: unknown;
      });
      markerLine?: {
        type?: "solid" | "none";
        enabled?: boolean;
        widthPt?: number;
        color?: string;
        fill?: ({
          type: "none";
        } | {
          type: "solid";
          color: string;
        } | {
          type: "linearGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
          angleDeg?: number;
        } | {
          type: "radialGradient";
          stops: Array<{
            position: number;
            color: string;
          }>;
        } | {
          type: "pattern";
          pattern: string;
          backgroundColor: string;
          foregroundColor: string;
        } | {
          type?: "raw";
          raw: unknown;
        });
        raw?: unknown;
      };
      numberFormat?: string;
      allSeries?: boolean;
      allMarkers?: boolean;
      dataLabels?: {
        showSeriesName?: boolean;
        showCategoryName?: boolean;
        showValue?: boolean;
        showPercent?: boolean;
      };
    }>;
  }>;
  removeSeries?: Array<number>;
  includeRaw?: boolean;
} & ({ chartIndex: number; } | { name: string; });
export type AiBridgeGeneratedSheetsDeleteChartArgs = {
  sheet?: string;
  chartIndex?: number;
  name?: string;
} & ({ chartIndex: number; } | { name: string; });

export interface AiBridgeToolArgumentsMap {
  word_inspect: AiBridgeGeneratedWordInspectArgs;
  word_replace_text: AiBridgeGeneratedWordReplaceTextArgs;
  word_append_paragraph: AiBridgeGeneratedWordAppendParagraphArgs;
  word_insert_paragraph: AiBridgeGeneratedWordInsertParagraphArgs;
  word_format_document: AiBridgeGeneratedWordFormatDocumentArgs;
  word_format_selection: AiBridgeGeneratedWordFormatSelectionArgs;
  word_format_matches: AiBridgeGeneratedWordFormatMatchesArgs;
  word_delete_matches: AiBridgeGeneratedWordDeleteMatchesArgs;
  word_add_hyperlink: AiBridgeGeneratedWordAddHyperlinkArgs;
  word_add_comment: AiBridgeGeneratedWordAddCommentArgs;
  word_add_bookmark: AiBridgeGeneratedWordAddBookmarkArgs;
  word_add_image: AiBridgeGeneratedWordAddImageArgs;
  word_inspect_advanced: AiBridgeGeneratedWordInspectAdvancedArgs;
  word_set_document_properties: AiBridgeGeneratedWordSetDocumentPropertiesArgs;
  word_manage_section: AiBridgeGeneratedWordManageSectionArgs;
  word_manage_style: AiBridgeGeneratedWordManageStyleArgs;
  word_set_tabs: AiBridgeGeneratedWordSetTabsArgs;
  word_set_numbering: AiBridgeGeneratedWordSetNumberingArgs;
  word_format_table_advanced: AiBridgeGeneratedWordFormatTableAdvancedArgs;
  word_add_nested_table: AiBridgeGeneratedWordAddNestedTableArgs;
  word_manage_drawing: AiBridgeGeneratedWordManageDrawingArgs;
  word_add_shape: AiBridgeGeneratedWordAddShapeArgs;
  word_add_chart: AiBridgeGeneratedWordAddChartArgs;
  word_add_math: AiBridgeGeneratedWordAddMathArgs;
  word_add_ole_object: AiBridgeGeneratedWordAddOleObjectArgs;
  word_manage_fields: AiBridgeGeneratedWordManageFieldsArgs;
  word_manage_long_document: AiBridgeGeneratedWordManageLongDocumentArgs;
  word_manage_comments: AiBridgeGeneratedWordManageCommentsArgs;
  word_manage_revisions: AiBridgeGeneratedWordManageRevisionsArgs;
  word_set_protection: AiBridgeGeneratedWordSetProtectionArgs;
  word_manage_content_control: AiBridgeGeneratedWordManageContentControlArgs;
  word_manage_custom_xml: AiBridgeGeneratedWordManageCustomXmlArgs;
  word_inspect_macros: AiBridgeGeneratedWordInspectMacrosArgs;
  word_set_macros: AiBridgeGeneratedWordSetMacrosArgs;
  word_set_watermark: AiBridgeGeneratedWordSetWatermarkArgs;
  word_format_paragraphs: AiBridgeGeneratedWordFormatParagraphsArgs;
  word_set_paragraph_text: AiBridgeGeneratedWordSetParagraphTextArgs;
  word_delete_paragraphs: AiBridgeGeneratedWordDeleteParagraphsArgs;
  word_set_list: AiBridgeGeneratedWordSetListArgs;
  word_insert_page_break: AiBridgeGeneratedWordInsertPageBreakArgs;
  word_navigate: AiBridgeGeneratedWordNavigateArgs;
  word_scroll: AiBridgeGeneratedWordScrollArgs;
  word_scale_font: AiBridgeGeneratedWordScaleFontArgs;
  word_add_table: AiBridgeGeneratedWordAddTableArgs;
  word_set_table_cell: AiBridgeGeneratedWordSetTableCellArgs;
  word_format_table: AiBridgeGeneratedWordFormatTableArgs;
  word_edit_table: AiBridgeGeneratedWordEditTableArgs;
  word_set_page_layout: AiBridgeGeneratedWordSetPageLayoutArgs;
  word_set_header_footer: AiBridgeGeneratedWordSetHeaderFooterArgs;
  word_set_document_text: AiBridgeGeneratedWordSetDocumentTextArgs;
  slides_inspect: AiBridgeGeneratedSlidesInspectArgs;
  slides_inspect_layouts: AiBridgeGeneratedSlidesInspectLayoutsArgs;
  slides_inspect_backgrounds: AiBridgeGeneratedSlidesInspectBackgroundsArgs;
  slides_inspect_themes: AiBridgeGeneratedSlidesInspectThemesArgs;
  slides_inspect_builtin_themes: AiBridgeGeneratedSlidesInspectBuiltinThemesArgs;
  slides_inspect_objects: AiBridgeGeneratedSlidesInspectObjectsArgs;
  slides_validate_layout: AiBridgeGeneratedSlidesValidateLayoutArgs;
  slides_replace_text: AiBridgeGeneratedSlidesReplaceTextArgs;
  slides_scale_font: AiBridgeGeneratedSlidesScaleFontArgs;
  slides_format_text: AiBridgeGeneratedSlidesFormatTextArgs;
  slides_format_selection: AiBridgeGeneratedSlidesFormatSelectionArgs;
  slides_add_slide: AiBridgeGeneratedSlidesAddSlideArgs;
  slides_duplicate_slide: AiBridgeGeneratedSlidesDuplicateSlideArgs;
  slides_delete_slide: AiBridgeGeneratedSlidesDeleteSlideArgs;
  slides_move_slide: AiBridgeGeneratedSlidesMoveSlideArgs;
  slides_set_visibility: AiBridgeGeneratedSlidesSetVisibilityArgs;
  slides_set_size: AiBridgeGeneratedSlidesSetSizeArgs;
  slides_apply_layout: AiBridgeGeneratedSlidesApplyLayoutArgs;
  slides_set_show_settings: AiBridgeGeneratedSlidesSetShowSettingsArgs;
  slides_apply_theme: AiBridgeGeneratedSlidesApplyThemeArgs;
  slides_apply_builtin_theme: AiBridgeGeneratedSlidesApplyBuiltinThemeArgs;
  slides_set_theme: AiBridgeGeneratedSlidesSetThemeArgs;
  slides_create_layout: AiBridgeGeneratedSlidesCreateLayoutArgs;
  slides_add_template_shape: AiBridgeGeneratedSlidesAddTemplateShapeArgs;
  slides_manage_template_object: AiBridgeGeneratedSlidesManageTemplateObjectArgs;
  slides_set_template_background: AiBridgeGeneratedSlidesSetTemplateBackgroundArgs;
  slides_set_text_content: AiBridgeGeneratedSlidesSetTextContentArgs;
  slides_format_paragraphs: AiBridgeGeneratedSlidesFormatParagraphsArgs;
  slides_update_object: AiBridgeGeneratedSlidesUpdateObjectArgs;
  slides_set_hyperlink: AiBridgeGeneratedSlidesSetHyperlinkArgs;
  slides_set_notes: AiBridgeGeneratedSlidesSetNotesArgs;
  slides_add_comment: AiBridgeGeneratedSlidesAddCommentArgs;
  slides_inspect_comments: AiBridgeGeneratedSlidesInspectCommentsArgs;
  slides_manage_comment: AiBridgeGeneratedSlidesManageCommentArgs;
  slides_set_transition: AiBridgeGeneratedSlidesSetTransitionArgs;
  slides_inspect_animations: AiBridgeGeneratedSlidesInspectAnimationsArgs;
  slides_manage_animation: AiBridgeGeneratedSlidesManageAnimationArgs;
  slides_add_table: AiBridgeGeneratedSlidesAddTableArgs;
  slides_set_table_cell: AiBridgeGeneratedSlidesSetTableCellArgs;
  slides_edit_table: AiBridgeGeneratedSlidesEditTableArgs;
  slides_format_table: AiBridgeGeneratedSlidesFormatTableArgs;
  slides_align_objects: AiBridgeGeneratedSlidesAlignObjectsArgs;
  slides_group_objects: AiBridgeGeneratedSlidesGroupObjectsArgs;
  slides_reorder_object: AiBridgeGeneratedSlidesReorderObjectArgs;
  slides_add_connector: AiBridgeGeneratedSlidesAddConnectorArgs;
  slides_add_freeform: AiBridgeGeneratedSlidesAddFreeformArgs;
  slides_add_textbox: AiBridgeGeneratedSlidesAddTextboxArgs;
  slides_add_word_art: AiBridgeGeneratedSlidesAddWordArtArgs;
  slides_add_math: AiBridgeGeneratedSlidesAddMathArgs;
  slides_add_image: AiBridgeGeneratedSlidesAddImageArgs;
  slides_add_image_shape: AiBridgeGeneratedSlidesAddImageShapeArgs;
  slides_add_ole_object: AiBridgeGeneratedSlidesAddOleObjectArgs;
  slides_set_background: AiBridgeGeneratedSlidesSetBackgroundArgs;
  slides_add_shape: AiBridgeGeneratedSlidesAddShapeArgs;
  slides_update_shape: AiBridgeGeneratedSlidesUpdateShapeArgs;
  slides_delete_object: AiBridgeGeneratedSlidesDeleteObjectArgs;
  slides_inspect_smartarts: AiBridgeGeneratedSlidesInspectSmartartsArgs;
  slides_add_smartart: AiBridgeGeneratedSlidesAddSmartartArgs;
  slides_update_smartart: AiBridgeGeneratedSlidesUpdateSmartartArgs;
  slides_delete_smartart: AiBridgeGeneratedSlidesDeleteSmartartArgs;
  slides_inspect_charts: AiBridgeGeneratedSlidesInspectChartsArgs;
  slides_add_chart: AiBridgeGeneratedSlidesAddChartArgs;
  slides_update_chart: AiBridgeGeneratedSlidesUpdateChartArgs;
  slides_delete_chart: AiBridgeGeneratedSlidesDeleteChartArgs;
  slides_inspect_macros: AiBridgeGeneratedSlidesInspectMacrosArgs;
  slides_set_macros: AiBridgeGeneratedSlidesSetMacrosArgs;
  slides_control_slideshow: AiBridgeGeneratedSlidesControlSlideshowArgs;
  sheets_inspect: AiBridgeGeneratedSheetsInspectArgs;
  sheets_set_values: AiBridgeGeneratedSheetsSetValuesArgs;
  sheets_set_formula: AiBridgeGeneratedSheetsSetFormulaArgs;
  sheets_inspect_range: AiBridgeGeneratedSheetsInspectRangeArgs;
  sheets_set_array_formula: AiBridgeGeneratedSheetsSetArrayFormulaArgs;
  sheets_replace_text: AiBridgeGeneratedSheetsReplaceTextArgs;
  sheets_format_range: AiBridgeGeneratedSheetsFormatRangeArgs;
  sheets_add_sheet: AiBridgeGeneratedSheetsAddSheetArgs;
  sheets_rename_sheet: AiBridgeGeneratedSheetsRenameSheetArgs;
  sheets_delete_sheet: AiBridgeGeneratedSheetsDeleteSheetArgs;
  sheets_manage_sheet: AiBridgeGeneratedSheetsManageSheetArgs;
  sheets_manage_range: AiBridgeGeneratedSheetsManageRangeArgs;
  sheets_set_rich_text: AiBridgeGeneratedSheetsSetRichTextArgs;
  sheets_inspect_names: AiBridgeGeneratedSheetsInspectNamesArgs;
  sheets_manage_names: AiBridgeGeneratedSheetsManageNamesArgs;
  sheets_recalculate: AiBridgeGeneratedSheetsRecalculateArgs;
  sheets_sort: AiBridgeGeneratedSheetsSortArgs;
  sheets_filter: AiBridgeGeneratedSheetsFilterArgs;
  sheets_inspect_tables: AiBridgeGeneratedSheetsInspectTablesArgs;
  sheets_manage_table: AiBridgeGeneratedSheetsManageTableArgs;
  sheets_manage_conditional_format: AiBridgeGeneratedSheetsManageConditionalFormatArgs;
  sheets_manage_validation: AiBridgeGeneratedSheetsManageValidationArgs;
  sheets_inspect_pivots: AiBridgeGeneratedSheetsInspectPivotsArgs;
  sheets_manage_pivot: AiBridgeGeneratedSheetsManagePivotArgs;
  sheets_inspect_drawings: AiBridgeGeneratedSheetsInspectDrawingsArgs;
  sheets_manage_drawing: AiBridgeGeneratedSheetsManageDrawingArgs;
  sheets_manage_hyperlink: AiBridgeGeneratedSheetsManageHyperlinkArgs;
  sheets_inspect_comments: AiBridgeGeneratedSheetsInspectCommentsArgs;
  sheets_manage_comments: AiBridgeGeneratedSheetsManageCommentsArgs;
  sheets_inspect_freeze_panes: AiBridgeGeneratedSheetsInspectFreezePanesArgs;
  sheets_manage_freeze_panes: AiBridgeGeneratedSheetsManageFreezePanesArgs;
  sheets_inspect_properties: AiBridgeGeneratedSheetsInspectPropertiesArgs;
  sheets_manage_properties: AiBridgeGeneratedSheetsManagePropertiesArgs;
  sheets_inspect_protected_ranges: AiBridgeGeneratedSheetsInspectProtectedRangesArgs;
  sheets_manage_protected_ranges: AiBridgeGeneratedSheetsManageProtectedRangesArgs;
  sheets_inspect_page_layout: AiBridgeGeneratedSheetsInspectPageLayoutArgs;
  sheets_manage_page_layout: AiBridgeGeneratedSheetsManagePageLayoutArgs;
  sheets_inspect_macros: AiBridgeGeneratedSheetsInspectMacrosArgs;
  sheets_set_macros: AiBridgeGeneratedSheetsSetMacrosArgs;
  sheets_add_chart: AiBridgeGeneratedSheetsAddChartArgs;
  sheets_inspect_charts: AiBridgeGeneratedSheetsInspectChartsArgs;
  sheets_update_chart: AiBridgeGeneratedSheetsUpdateChartArgs;
  sheets_delete_chart: AiBridgeGeneratedSheetsDeleteChartArgs;
}
// </ai-bridge-generated:tool-arguments>

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
  validateLayout(args?: SlidesValidateLayoutArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
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
  inspectSmartArts(args?: SlidesInspectSmartArtsArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  addSmartArt(args: SlidesAddSmartArtArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  updateSmartArt(args: SlidesUpdateSmartArtArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  deleteSmartArt(args: SlidesDeleteSmartArtArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
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
  readonly version: "0.2.4";
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
  /** Enable the document-hub long-poll Relay. relayBaseUrl must also be configured. */
  httpRelay?: boolean;
  /** Same-origin document-hub Relay base path, e.g. /api/v1/editor-relay. */
  relayBaseUrl?: string;
  /** Same-origin image service base path. */
  imageBaseUrl?: string;
  /** Same-origin persistence service base path. */
  persistenceBaseUrl?: string;
  /** document-hub editor session identifier forwarded to Relay requests. */
  editorSessionId?: string;
  /** Stable document-hub document identifier. */
  documentId?: string;
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
