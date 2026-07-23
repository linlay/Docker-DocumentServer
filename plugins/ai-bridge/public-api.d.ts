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
  version: "0.1.0";
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
export interface TextFormat extends BasicTextFormat { underline?: boolean; }

export interface WordInspectArgs { maxChars?: number; }
export interface WordReplaceTextArgs { search: string; replace: string; matchCase?: boolean; }
export interface WordAppendParagraphArgs extends BasicTextFormat {
  text: string;
  align?: "left" | "center" | "right" | "both";
  spacingBefore?: number;
  spacingAfter?: number;
}
export interface WordInsertParagraphArgs extends BasicTextFormat {
  text: string;
  inline?: boolean;
  align?: "left" | "center" | "right" | "both" | string;
}
export interface WordFormatDocumentArgs extends TextFormat {
  strikeout?: boolean;
  align?: "left" | "center" | "right" | "both";
  spacingBefore?: number;
  spacingAfter?: number;
}
export interface WordFormatSelectionArgs extends TextFormat { strikeout?: boolean; }
export interface WordScaleFontArgs { scale: number; }
export interface WordAddTableArgs {
  rows: number;
  cols: number;
  data?: unknown[][];
  widthPercent?: number;
}
export interface WordSetDocumentTextArgs { text: string; }

export interface SlidesInspectArgs { maxChars?: number; }
export interface SlidesReplaceTextArgs { search: string; replace: string; matchCase?: boolean; slide?: number; }
export interface SlidesScaleFontArgs { scale: number; slide?: number; }
export interface SlidesFormatTextArgs extends TextFormat { slide?: number; }
export interface SlidesFormatSelectionArgs extends TextFormat {}
export interface SlidesAddSlideArgs { index?: number; title?: string; titleFontSize?: number; backgroundColor?: string; }
export interface SlidesDuplicateSlideArgs { slide: number; }
export interface SlidesDeleteSlideArgs { slide: number; }
export interface SlidesAddTextBoxArgs extends BasicTextFormat {
  slide: number;
  text: string;
  xMm?: number;
  yMm?: number;
  widthMm?: number;
  heightMm?: number;
  fillColor?: string;
  align?: string;
}

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
  type?: string;
  title?: string;
  titleFontSize?: number;
  inRows?: boolean;
  style?: number;
  widthMm?: number;
  heightMm?: number;
  fromColumn?: number;
  fromRow?: number;
}

export interface AiBridgeToolArgumentsMap {
  word_inspect: WordInspectArgs;
  word_replace_text: WordReplaceTextArgs;
  word_append_paragraph: WordAppendParagraphArgs;
  word_insert_paragraph: WordInsertParagraphArgs;
  word_format_document: WordFormatDocumentArgs;
  word_format_selection: WordFormatSelectionArgs;
  word_scale_font: WordScaleFontArgs;
  word_add_table: WordAddTableArgs;
  word_set_document_text: WordSetDocumentTextArgs;
  slides_inspect: SlidesInspectArgs;
  slides_replace_text: SlidesReplaceTextArgs;
  slides_scale_font: SlidesScaleFontArgs;
  slides_format_text: SlidesFormatTextArgs;
  slides_format_selection: SlidesFormatSelectionArgs;
  slides_add_slide: SlidesAddSlideArgs;
  slides_duplicate_slide: SlidesDuplicateSlideArgs;
  slides_delete_slide: SlidesDeleteSlideArgs;
  slides_add_textbox: SlidesAddTextBoxArgs;
  sheets_inspect: SheetsInspectArgs;
  sheets_set_values: SheetsSetValuesArgs;
  sheets_set_formula: SheetsSetFormulaArgs;
  sheets_replace_text: SheetsReplaceTextArgs;
  sheets_format_range: SheetsFormatRangeArgs;
  sheets_add_sheet: SheetsAddSheetArgs;
  sheets_rename_sheet: SheetsRenameSheetArgs;
  sheets_delete_sheet: SheetsDeleteSheetArgs;
  sheets_add_chart: SheetsAddChartArgs;
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
  scaleFont(args: WordScaleFontArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  addTable(args: WordAddTableArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  setDocumentText(args: WordSetDocumentTextArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
}

export interface AiBridgeSlidesApi {
  inspect(args?: SlidesInspectArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  replaceText(args: SlidesReplaceTextArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  scaleFont(args: SlidesScaleFontArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  formatText(args: SlidesFormatTextArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  formatSelection(args: SlidesFormatSelectionArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  addSlide(args: SlidesAddSlideArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  duplicateSlide(args: SlidesDuplicateSlideArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  deleteSlide(args: SlidesDeleteSlideArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
  addTextBox(args: SlidesAddTextBoxArgs, options?: AiBridgeRequestOptions): Promise<AiBridgeExecutionResult>;
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
}

export type AiBridgeEventName = "ready" | "reload" | "error";

export interface AiBridgeApi {
  readonly version: "0.1.0";
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
