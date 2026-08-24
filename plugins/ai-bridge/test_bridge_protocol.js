const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");

const pluginSource = fs.readFileSync(path.join(__dirname, "plugin.js"), "utf8");
const hostSource = fs.readFileSync(path.join(__dirname, "host-bridge.js"), "utf8");
const clientSource = fs.readFileSync(path.join(__dirname, "client-sdk.js"), "utf8");
const bootstrapSource = fs.readFileSync(path.join(__dirname, "bootstrap.js"), "utf8");
const wordBridgeSource = fs.readFileSync(path.join(__dirname, "bridges/word-bridge.js"), "utf8");
const publicContract = JSON.parse(fs.readFileSync(path.join(__dirname, "public-api.json"), "utf8"));
const TEST_PLUGIN_GUID = "asc.{A17E5F31-64AA-4E37-9A42-8D430814C2F6}";
const docxCapabilityMatrix = fs.readFileSync(
  path.join(__dirname, "DOCX-CAPABILITIES.zh-CN.md"),
  "utf8",
);

function eventTarget(target) {
  const listeners = new Map();
  target.dispatchedEvents = [];
  target.addEventListener = (name, listener) => {
    const entries = listeners.get(name) || [];
    entries.push(listener);
    listeners.set(name, entries);
  };
  target.removeEventListener = (name, listener) => {
    const entries = listeners.get(name) || [];
    listeners.set(name, entries.filter(entry => entry !== listener));
  };
  target.dispatch = (name, event) => {
    for (const listener of listeners.get(name) || []) listener(event);
  };
  target.dispatchEvent = event => {
    target.dispatchedEvents.push(event);
    target.dispatch(event.type, event);
    return true;
  };
  return target;
}

function createHarness(options = {}) {
  const editorType = options.editorType || "word";
  const channelId = options.channelId || "11111111-2222-4333-8444-555555555555";
  const defaultPluginOptions = {
    hostOrigin: "https://app.test",
    channelId,
  };
  const hostPluginOptions = Object.prototype.hasOwnProperty.call(options, "hostPluginOptions")
    ? options.hostPluginOptions
    : (options.strictHandshake || options.nestedHost ? defaultPluginOptions : null);
  const ascPluginOptions = Object.prototype.hasOwnProperty.call(options, "ascPluginOptions")
    ? options.ascPluginOptions
    : (hostPluginOptions || {});
  const editorConfig = {
    documentId: options.documentId,
    documentType: editorType,
    document: { key: "doc-key-v1", title: `demo.${editorType === "word" ? "docx" : editorType === "slide" ? "pptx" : "xlsx"}`, fileType: editorType === "word" ? "docx" : editorType === "slide" ? "pptx" : "xlsx" },
    editorConfig: {
      callbackUrl: "https://app.test/callback",
      lang: options.interfaceLanguage || "en",
      region: options.region || "en-US",
      user: { id: "user-1" },
    },
    events: options.editorEvents || {},
    token: options.editorToken || "editor-token",
  };
  if (hostPluginOptions) {
    editorConfig.editorConfig.plugins = {
      options: {
        [TEST_PLUGIN_GUID]: hostPluginOptions,
      },
    };
  }
  const executedToolCalls = [];
  const initializationCommands = [];
  const servicePaths = [];
  const serviceRequests = [];
  const hostRelayPaths = [];
  const imageBaseUrl = Object.prototype.hasOwnProperty.call(options, "imageBaseUrl")
    ? options.imageBaseUrl
    : "/api/v1/editor-relay/images";
  const persistenceBaseUrl = Object.prototype.hasOwnProperty.call(options, "persistenceBaseUrl")
    ? options.persistenceBaseUrl
    : "/api/v1/editor-relay/persistence";
  let reloadCount = 0;
  const documentElement = { dataset: {} };
  const sessionValues = new Map(Object.entries(options.sessionValues || {}));
  const localValues = new Map(Object.entries(options.localValues || {}));
  const sessionStorage = {
    getItem(key) { return sessionValues.has(key) ? sessionValues.get(key) : null; },
    setItem(key, value) { sessionValues.set(key, String(value)); },
    removeItem(key) { sessionValues.delete(key); },
  };
  const localStorage = {
    getItem(key) { return localValues.has(key) ? localValues.get(key) : null; },
    setItem(key, value) { localValues.set(key, String(value)); },
    removeItem(key) { localValues.delete(key); },
  };

  const hostWindow = eventTarget({
    location: {
      origin: "https://app.test",
      href: "https://app.test/editor",
      pathname: options.pathname || "/editor",
      hostname: options.hostname || "app.test",
      protocol: options.protocol || "https:",
      reload() { reloadCount += 1; },
    },
    crypto: { randomUUID: () => "00000000-0000-4000-8000-000000000001" },
    btoa: value => Buffer.from(value, "binary").toString("base64"),
    setTimeout: options.hostSetTimeout || setTimeout,
    clearTimeout: options.hostClearTimeout || clearTimeout,
    history: { replaceState() {} },
    navigator: {
      onLine: Object.prototype.hasOwnProperty.call(options, "navigatorOnline")
        ? options.navigatorOnline
        : true,
      sendBeacon() { return true; },
    },
    sessionStorage,
    localStorage,
    aiBridgeOptions: {
      getEditorConfig: () => editorConfig,
      clientOrigins: options.clientOrigins || [],
      httpRelay: options.httpRelay,
      relayBaseUrl: options.relayBaseUrl,
      imageBaseUrl,
      persistenceBaseUrl,
      editorSessionId: options.editorSessionId,
      documentId: options.optionsDocumentId,
    },
  });
  hostWindow.fetch = async (requestPath, requestOptions) => {
    if (
      options.interceptPersistence !== true
      && persistenceBaseUrl
      && requestPath.startsWith(`${persistenceBaseUrl}/`)
    ) {
      servicePaths.push(requestPath);
      serviceRequests.push({ path: requestPath, requestOptions });
      const configured = typeof options.serviceResponse === "function"
        ? options.serviceResponse(requestPath)
        : options.serviceResponse;
      return {
        ok: true,
        status: 200,
        json: async () => configured || {
          ok: true,
          accepted: true,
          persisted: true,
          status: "saved",
        },
      };
    }
    if (typeof options.hostFetch === "function") {
      hostRelayPaths.push(requestPath);
      return options.hostFetch(requestPath, requestOptions);
    }
    throw new Error(`unexpected host request: ${requestPath}`);
  };

  const pluginWindow = eventTarget({
    location: {
      origin: "https://docs.test",
      href: "https://docs.test/sdkjs-plugins/ai-bridge/index.html",
      ancestorOrigins: options.nestedHost
        ? ["https://docs.test", "https://app.test", "https://outer.test"]
        : ["https://app.test"],
    },
    setTimeout,
    clearTimeout,
    setInterval: options.pluginSetInterval || setInterval,
    clearInterval: options.pluginClearInterval || clearInterval,
    fetch: async (path, requestOptions) => {
      servicePaths.push(path);
      serviceRequests.push({ path, requestOptions });
      const configured = typeof options.serviceResponse === "function"
        ? options.serviceResponse(path)
        : options.serviceResponse;
      return {
        ok: true,
        status: 200,
        json: async () => configured || { persisted: true, path },
      };
    },
  });
  if (options.bootstrapState) pluginWindow.__aiBridgeBootstrapState = options.bootstrapState;

  const ancestorMessages = [];
  const hostMessages = [];
  const topProxy = {
    parent: null,
    postMessage(message, targetOrigin) {
      ancestorMessages.push({ ancestor: "host", message, targetOrigin });
      if (
        targetOrigin !== "https://app.test"
        || options.dropPluginHello
        || (options.delayPluginHello && message.type === "hello")
      ) return;
      queueMicrotask(() => hostWindow.dispatch("message", {
        origin: "https://docs.test",
        source: pluginHandle,
        data: message,
      }));
    },
  };
  topProxy.parent = topProxy;

  const outerProxy = {
    parent: null,
    postMessage(message, targetOrigin) {
      ancestorMessages.push({ ancestor: "outer", message, targetOrigin });
    },
  };
  outerProxy.parent = outerProxy;

  const hostProxy = {
    parent: outerProxy,
    postMessage(message, targetOrigin) {
      ancestorMessages.push({ ancestor: "host", message, targetOrigin });
      if (
        targetOrigin !== "https://app.test"
        || options.dropPluginHello
        || (options.delayPluginHello && message.type === "hello")
      ) return;
      queueMicrotask(() => hostWindow.dispatch("message", {
        origin: "https://docs.test",
        source: pluginHandle,
        data: message,
      }));
    },
  };

  const editorProxy = {
    parent: hostProxy,
    postMessage(message, targetOrigin) {
      ancestorMessages.push({ ancestor: "editor", message, targetOrigin });
    },
  };
  const hostSourceProxy = options.nestedHost ? hostProxy : topProxy;

  const pluginHandle = {
    postMessage(message, targetOrigin) {
      assert.equal(targetOrigin, "https://docs.test");
      hostMessages.push({ message, targetOrigin });
      queueMicrotask(() => pluginWindow.dispatch("message", {
        origin: "https://app.test",
        source: hostSourceProxy,
        data: message,
      }));
    },
  };

  pluginWindow.parent = options.nestedHost ? editorProxy : topProxy;
  pluginWindow.top = options.nestedHost ? outerProxy : topProxy;
  pluginWindow.AICopilotBridges = options.missingBridge ? {} : {
    [editorType]: {
      execute: async toolCalls => {
        executedToolCalls.push(...toolCalls);
        return {
          changed: toolCalls.filter(call => (
            !call.name.endsWith("_inspect") &&
            call.name !== "word_navigate" &&
            call.name !== "word_scroll"
          )).length,
          results: toolCalls.map(call => ({ name: call.name, document: "demo" })),
          needsSave: toolCalls.some(call => (
            !call.name.endsWith("_inspect") &&
            call.name !== "word_navigate" &&
            call.name !== "word_scroll"
          )),
        };
      },
    },
  };
  if (editorType === "cell" && !options.missingBridge) {
    pluginWindow.AICopilotBridges.cell.probeCapabilities = async () => ({
      runtime: { product: "ONLYOFFICE", version: "test", edition: "enterprise" },
      features: {
        sheets: {
          nativeTables: { create: true, inspect: true },
          rangeStyleTables: { create: true },
          conditionalFormatting: { create: true },
          charts: { create: true, addSeriesOnCreate: false, delete: true },
        },
      },
    });
  }
  if (typeof options.bridgePreflight === "function" && !options.missingBridge) {
    pluginWindow.AICopilotBridges[editorType].preflight = options.bridgePreflight;
  }

  const Asc = {
    scope: {},
    plugin: {
      info: { editorType, options: ascPluginOptions },
      callCommand(command, close, calc, callback) {
        initializationCommands.push(command.toString());
        callback(JSON.stringify(
          options.saveError
            ? { ok: false, error: options.saveError }
            : { ok: true },
        ));
      },
    },
  };

  const hostDocument = options.hostDocument || eventTarget({
    documentElement,
    readyState: options.documentReadyState,
    visibilityState: options.pageVisibilityState,
    querySelector() { return null; },
    currentScript: {
      src: "https://docs.test/sdkjs-plugins/ai-bridge/host-bridge.js",
    },
  });
  vm.runInNewContext(hostSource, {
    window: hostWindow,
    document: hostDocument,
    URL,
    Date: options.hostDate || Date,
    CustomEvent: class CustomEvent {
      constructor(name, options) { this.type = name; this.detail = options && options.detail; }
    },
    console,
  });

  if (typeof options.beforePluginInit === "function") {
    options.beforePluginInit({ editorConfig, hostWindow });
  }

  vm.runInNewContext(pluginSource, {
    window: pluginWindow,
    Asc,
    fetch: pluginWindow.fetch,
    URL,
    console,
  });

  Asc.plugin.init();
  return {
    hostWindow,
    editorConfig,
    executedToolCalls,
    initializationCommands,
    servicePaths,
    serviceRequests,
    hostRelayPaths,
    documentElement,
    sessionValues,
    localValues,
    pluginWindow,
    Asc,
    pluginHandle,
    topProxy,
    hostProxy,
    outerProxy,
    editorProxy,
    ancestorMessages,
    hostMessages,
    get reloadCount() { return reloadCount; },
  };
}

async function waitFor(check, timeoutMs = 1000) {
  const deadline = Date.now() + timeoutMs;
  while (!check()) {
    if (Date.now() >= deadline) throw new Error("timed out waiting for condition");
    await new Promise(resolve => setTimeout(resolve, 5));
  }
}

function relayResponse(status, body) {
  return {
    ok: status >= 200 && status < 300,
    status,
    json: async () => body,
  };
}

function importedImageResponse(overrides = {}) {
  const assetId = overrides.assetId || `${"a".repeat(64)}.png`;
  return relayResponse(200, {
    ok: true,
    asset: {
      assetId,
      path: `/api/v1/editor-relay/images/${assetId}?token=signed-token`,
      mimeType: "image/png",
      widthPx: 400,
      heightPx: 200,
      expiresAt: 2_000_000_000,
      ...overrides,
    },
  });
}

function replaceLiteral(text, properties) {
  const source = String(text);
  const search = String(properties.searchString);
  const replacement = String(properties.replaceString);
  const haystack = properties.matchCase ? source : source.toLowerCase();
  const needle = properties.matchCase ? search : search.toLowerCase();
  let result = "";
  let offset = 0;
  let match;
  while (needle && (match = haystack.indexOf(needle, offset)) !== -1) {
    result += source.slice(offset, match) + replacement;
    offset = match + needle.length;
  }
  return result + source.slice(offset);
}

function createWordBridgeHarness(options = {}) {
  let documentText = String(options.text || "");
  let documentParagraphs = [];
  let documentElements = [];
  let nextParagraphId = 1;
  let currentPage = Number(options.currentPage) || 0;
  const createdImages = [];
  const createdTables = [];
  const createdContentControls = [];
  const executeMethodCalls = [];
  const nativeSearchAndReplaceCalls = [];
  const callCommandCalls = [];
  let capabilityProbeResultUsed = false;
  const insertContentCalls = [];
  const addElementCalls = [];
  const formattedRanges = [];
  const navigatedPages = [];
  const rangeActions = [];
  const createdStyles = [];
  const createdCharts = [];
  const createdSectionParagraphs = [];
  const appliedStyles = [];
  const fieldInstructions = [];
  const revisionEvents = [];
  const scope = {};
  const styles = new Map();
  const internalStyles = [];
  let historyPoints = 0;
  let revisionTracking = Boolean(options.revisionTracking);

  const coreFields = [
    "Title", "Subject", "Creator", "Description", "Keywords", "Category",
    "Language", "Identifier", "LastModifiedBy", "Revision", "Created", "Modified",
  ];

  function createCoreBacking(initial = {}) {
    const backing = {};
    for (const field of coreFields) {
      const key = field.charAt(0).toLowerCase() + field.slice(1);
      backing[key] = Object.prototype.hasOwnProperty.call(initial, key) ? initial[key] : null;
      backing[`set${field}`] = function (value) { this[key] = value; };
      backing[`asc_get${field}`] = function () { return this[key]; };
    }
    return backing;
  }

  function createApiCore(backing) {
    const core = { Core: backing };
    for (const field of coreFields) {
      core[`Set${field}`] = function (value) { this.Core[`set${field}`](value); };
      core[`Get${field}`] = function () { return this.Core[`asc_get${field}`](); };
    }
    if (options.missingCoreSetter) delete core[String(options.missingCoreSetter)];
    return core;
  }

  function createCustomPropertiesBacking(initial = {}) {
    return { values: { ...initial } };
  }

  function createApiCustomProperties(backing) {
    return {
      CustomProperties: backing,
      Add(name, value) { this.CustomProperties.values[String(name)] = value; return true; },
      Get(name) { return this.CustomProperties.values[String(name)]; },
    };
  }

  function createStyle(name, type) {
    const style = {
      id: `1_${createdStyles.length + 30}`,
      name: String(name),
      type: String(type),
      basedOn: null,
      GetName() { return this.name; },
      GetType() { return this.type; },
      GetBasedOn() { return this.basedOn; },
      SetBasedOn(value) { this.basedOn = value; return true; },
      GetTextPr() { return {}; },
      GetParaPr() { return {}; },
    };
    styles.set(style.name, style);
    internalStyles[style.id] = style;
    createdStyles.push(style);
    return style;
  }

  const internalStyleCollection = {
    GetAllStyles() {
      return internalStyles;
    },
  };

  function elementParagraphs(element) {
    if (!element) return [];
    if (typeof element.GetClassType === "function" && element.GetClassType() === "paragraph") {
      return [element];
    }
    if (typeof element.GetRowsCount === "function" && typeof element.GetRow === "function") {
      const paragraphs = [];
      for (let rowIndex = 0; rowIndex < element.GetRowsCount(); rowIndex += 1) {
        const row = element.GetRow(rowIndex);
        for (let columnIndex = 0; columnIndex < row.GetCellsCount(); columnIndex += 1) {
          const cell = row.GetCell(columnIndex);
          if (!cell || typeof cell.GetContent !== "function") continue;
          const content = cell.GetContent();
          if (content && typeof content.GetAllParagraphs === "function") {
            paragraphs.push(...content.GetAllParagraphs());
          }
        }
      }
      return paragraphs;
    }
    return [];
  }

  function refreshDocumentParagraphs() {
    if (documentElements.length) {
      documentParagraphs = documentElements.flatMap(elementParagraphs);
    }
  }

  function syncDocumentText() {
    refreshDocumentParagraphs();
    documentText = documentParagraphs.map(paragraph => paragraph.text).join("\n");
  }

  function createTextRun(value) {
    return {
      text: String(value),
      style: null,
      GetClassType() { return "run"; },
      SetStyle(style) { this.style = style; return true; },
      SetFontSize(size) { this.fontSize = size; return true; },
      SetFontFamily(family) { this.fontFamily = family; return true; },
      SetBold(bold) { this.bold = Boolean(bold); return true; },
      SetItalic(italic) { this.italic = Boolean(italic); return true; },
      SetUnderline(underline) { this.underline = Boolean(underline); return true; },
      SetColor(color) { this.color = color; return true; },
      SetHighlight(color) { this.highlightColor = color; return true; },
      SetSpacing(spacing) { this.characterSpacing = spacing; return true; },
    };
  }

  function createParagraph(text = "") {
    const paragraphId = nextParagraphId++;
    const paragraph = {
      text: String(text),
      drawings: [],
      runs: [],
      inlineControls: [],
      parentTable: null,
      parentContentControl: null,
      GetClassType() { return "paragraph"; },
      GetParaId() { return `para-${paragraphId}`; },
      GetInternalId() { return `internal-${paragraphId}`; },
      GetText() { return this.text; },
      GetStyle() { return this.style || null; },
      GetNumPr() { return this.numbering || null; },
      GetParentTable() { return this.parentTable; },
      GetParentContentControl() { return this.parentContentControl; },
      AddText(value) {
        const run = createTextRun(value);
        run.GetRange = function () {
          return {
            AddField(instruction) {
              fieldInstructions.push(String(instruction));
              return true;
            },
          };
        };
        this.runs.push(run);
        this.text += run.text;
        syncDocumentText();
        return run;
      },
      SetText(value) {
        this.text = String(value);
        this.runs = [];
        if (this.text) {
          this.runs.push(createTextRun(this.text));
        }
        syncDocumentText();
        return this.runs[0] || {};
      },
      RemoveAllElements() {
        this.text = "";
        this.runs = [];
        this.inlineControls = [];
        syncDocumentText();
        return true;
      },
      GetElementsCount() { return this.runs.length; },
      GetElement(index) { return this.runs[index] || null; },
      SetJc(value) { this.align = value; return true; },
      SetStyle(value) { this.style = value; return true; },
      SetSpacingBefore(value) { this.spacingBefore = value; return true; },
      SetSpacingAfter(value) { this.spacingAfter = value; return true; },
      GetSpacingBefore() { return this.spacingBefore === undefined ? 0 : this.spacingBefore; },
      GetSpacingAfter() { return this.spacingAfter === undefined ? 0 : this.spacingAfter; },
      SetSpacingLine(value, rule) { this.spacingLine = { value, rule }; return true; },
      SetIndFirstLine(value) { this.firstLineIndent = value; return true; },
      SetIndLeft(value) { this.leftIndent = value; return true; },
      SetIndRight(value) { this.rightIndent = value; return true; },
      SetNumbering(level) { this.numbering = level; return true; },
      SetPageBreakBefore(value) { this.pageBreakBefore = Boolean(value); return true; },
      AddPageBreak() { this.pageBreakAfter = true; return true; },
      AddInlineLvlSdt(control) {
        this.inlineControls.push(control);
        control.paragraph = this;
        this.text += String(control.text || "");
        syncDocumentText();
        return control;
      },
      Select() { this.selected = true; return true; },
      AddDrawing(drawing) {
        this.drawings.push(drawing);
        drawing.paragraph = this;
        return drawing;
      },
    };
    if (options.getPosInParentSupported !== false) {
      paragraph.GetPosInParent = function () { return documentElements.indexOf(this); };
    }
    if (paragraph.text) {
      paragraph.runs.push(createTextRun(paragraph.text));
    }
    return paragraph;
  }

  function createTable(rowsCount, columnsCount) {
    const table = {
      rows: [],
      width: null,
      style: null,
      parentTable: null,
      parentContentControl: null,
      GetClassType() { return "table"; },
      GetParentTable() { return this.parentTable; },
      GetParentContentControl() { return this.parentContentControl; },
      GetRowsCount() { return this.rows.length; },
      GetRow(index) { return this.rows[index] || null; },
      SetWidth(kind, value) { this.width = { kind, value }; return true; },
      SetJc(value) { this.align = value; return true; },
      GetJc() { return this.align || "left"; },
      SetStyle(value) { this.style = value; return true; },
      SetTableLook(...values) { this.tableLook = values; return true; },
      SetTableTitle(value) { this.title = String(value); return true; },
      SetTableDescription(value) { this.description = String(value); return true; },
    };
    if (options.getPosInParentSupported !== false) {
      table.GetPosInParent = function () { return documentElements.indexOf(this); };
    }
    for (let rowIndex = 0; rowIndex < rowsCount; rowIndex += 1) {
      const row = {
        cells: [],
        GetCellsCount() { return this.cells.length; },
        GetCell(index) { return this.cells[index] || null; },
      };
      for (let columnIndex = 0; columnIndex < columnsCount; columnIndex += 1) {
        let cellParagraphs = [createParagraph()];
        cellParagraphs[0].parentTable = table;
        const cell = {
          backgroundColor: null,
          verticalAlign: null,
          width: null,
          nestedTables: [],
          GetContent() {
            return {
              GetAllParagraphs() { return cellParagraphs; },
              GetElement(index) { return cellParagraphs[index] || null; },
              Push(element) {
                cell.nestedTables.push(element);
                element.parentTable = table;
                return true;
              },
            };
          },
          SetText(value) {
            const paragraph = createParagraph(String(value));
            paragraph.parentTable = table;
            cellParagraphs = [paragraph];
            syncDocumentText();
            return paragraph.runs[0] || {};
          },
          SetBackgroundColor(value) { this.backgroundColor = value; return true; },
          SetVerticalAlign(value) { this.verticalAlign = value; return true; },
          SetWidth(kind, value) { this.width = { kind, value }; return true; },
        };
        row.cells.push(cell);
      }
      table.rows.push(row);
    }
    createdTables.push(table);
    return table;
  }

  function createInlineContentControl() {
    const control = {
      id: `inline-${createdContentControls.length + 1}`,
      text: "",
      tag: "",
      title: "",
      placeholder: "",
      lock: "unlocked",
      appearance: "boundingBox",
      GetClassType() { return "inlineLvlSdt"; },
      GetId() { return this.id; },
      GetTag() { return this.tag; },
      GetAlias() { return this.title; },
      GetPlaceholderText() { return this.placeholder; },
      GetLock() { return this.lock; },
      GetAppearance() { return this.appearance; },
      SetTag(value) { this.tag = String(value); return true; },
      SetAlias(value) { this.title = String(value); return true; },
      SetPlaceholderText(value) { this.placeholder = String(value); return true; },
      SetLock(value) { this.lock = String(value); return true; },
      SetColor(value) { this.color = value; return true; },
      SetAppearance(value) { this.appearance = String(value); return true; },
      SetDateFormat(value) { this.dateFormat = String(value); return true; },
      SetDate(value) { this.date = value; return true; },
      SetDataBinding(value) { this.dataBinding = value; return true; },
      UpdateFromXmlMapping() { this.updatedFromXml = true; return true; },
      RemoveAllElements() { this.text = ""; return true; },
      AddText(value) { this.text += String(value); return true; },
    };
    createdContentControls.push(control);
    return control;
  }

  documentParagraphs = (
    Array.isArray(options.paragraphs)
      ? options.paragraphs
      : documentText.split("\n")
  ).map(createParagraph);
  if (!documentParagraphs.length) documentParagraphs = [createParagraph()];
  documentElements = [
    ...documentParagraphs,
    ...(Array.isArray(options.tables) ? options.tables : []),
  ];
  syncDocumentText();
  const initialDocumentElements = [...documentElements];
  const initialDocumentTables = initialDocumentElements.filter(element => (
    element && typeof element.GetRowsCount === "function" && typeof element.GetRow === "function"
  ));
  const initialCore = Object.prototype.hasOwnProperty.call(options, "core")
    ? options.core
    : createCoreBacking(options.coreValues || {});
  const initialCustomProperties = Object.prototype.hasOwnProperty.call(options, "customProperties")
    ? options.customProperties
    : createCustomPropertiesBacking(options.customPropertyValues || {});
  const document = {
    Document: {
      Core: initialCore,
      CustomProperties: initialCustomProperties,
      Get_Styles() { return internalStyleCollection; },
    },
    GetCore() { return createApiCore(this.Document.Core); },
    GetCustomProperties() { return createApiCustomProperties(this.Document.CustomProperties); },
    GetText() { return documentText; },
    GetContent() {
      return [
        ...(options.staleTableCollections ? initialDocumentElements : documentElements),
      ];
    },
    GetElementsCount() { return documentElements.length; },
    GetElement(index) { return documentElements[index] || null; },
    GetAllParagraphs() { return documentParagraphs; },
    GetAllNumberedParagraphs() { return documentParagraphs.filter(paragraph => paragraph.numbering); },
    GetAllTables() {
      if (options.staleTableCollections) return initialDocumentTables;
      return documentElements.filter(element => (
        element && typeof element.GetRowsCount === "function" && typeof element.GetRow === "function"
      ));
    },
    GetAllContentControls() {
      return [
        ...(Array.isArray(options.contentControls) ? options.contentControls : []),
        ...createdContentControls,
      ];
    },
    GetContentControlsByTag(tag) {
      return this.GetAllContentControls().filter(control => (
        typeof control.GetTag === "function" && String(control.GetTag()) === String(tag)
      ));
    },
    GetCurrentContentControl() {
      return options.currentContentControl || null;
    },
    GetCurrentParagraph() {
      if (options.currentParagraphIndex !== undefined) {
        return documentParagraphs[Number(options.currentParagraphIndex) - 1] || null;
      }
      return options.currentParagraph || null;
    },
    AddComboBoxContentControl(list, selected) {
      return typeof options.addComboBoxContentControl === "function"
        ? options.addComboBoxContentControl(list, selected)
        : null;
    },
    AddDropDownListContentControl(list, selected) {
      return typeof options.addDropDownListContentControl === "function"
        ? options.addDropDownListContentControl(list, selected)
        : null;
    },
    GetSections() { return Array.isArray(options.sections) ? options.sections : []; },
    GetFinalSection() {
      const sections = Array.isArray(options.sections) ? options.sections : [];
      return sections.length ? sections[sections.length - 1] : null;
    },
    CreateSection(paragraph) {
      createdSectionParagraphs.push(paragraph);
      return typeof options.createSection === "function" ? options.createSection(paragraph) : {};
    },
    RemoveAllElements() {
      if (options.removeAllElementsAccepted === false) return false;
      documentElements = [createParagraph()];
      syncDocumentText();
      return true;
    },
    Push(element) {
      if (options.pushAccepted === false) return false;
      documentElements.push(element);
      syncDocumentText();
      return true;
    },
    InsertContent(elements) {
      if (options.insertContentAccepted === false) return false;
      insertContentCalls.push([...elements]);
      documentElements.push(...elements);
      syncDocumentText();
      return true;
    },
    AddElement(position, element) {
      if (options.addElementAccepted === false) return false;
      addElementCalls.push({ position, element });
      documentElements.splice(position, 0, element);
      syncDocumentText();
      return true;
    },
    Search(search, matchCase) {
      const source = matchCase ? documentText : documentText.toLowerCase();
      const needle = matchCase ? String(search) : String(search).toLowerCase();
      const ranges = [];
      let offset = 0;
      while (needle && (offset = source.indexOf(needle, offset)) !== -1) {
        const format = { start: offset, end: offset + needle.length };
        formattedRanges.push(format);
        ranges.push({
          SetFontSize(value) { format.fontSize = value; },
          SetFontFamily(value) { format.fontFamily = value; },
          SetBold(value) { format.bold = value; },
          SetItalic(value) { format.italic = value; },
          SetUnderline(value) { format.underline = value; },
          SetStrikeout(value) { format.strikeout = value; },
          SetColor(value) { format.color = value; },
          SetHighlight(value) { format.highlightColor = value; },
          SetStyle(value) {
            appliedStyles.push({ start: format.start, style: value });
            return true;
          },
          Delete() { rangeActions.push({ action: "delete", start: format.start, end: format.end }); return true; },
          AddHyperlink(url, screenTip, bookmarkName) {
            rangeActions.push({ action: "hyperlink", start: format.start, url, screenTip, bookmarkName });
            return {};
          },
          AddComment(text, author, userId) {
            rangeActions.push({ action: "comment", start: format.start, text, author, userId });
            return {
              GetId() { return `comment-${format.start}`; },
              GetQuoteText() { return documentText.slice(format.start, format.end); },
            };
          },
          AddBookmark(name) {
            rangeActions.push({ action: "bookmark", start: format.start, name });
            return true;
          },
          Select() { rangeActions.push({ action: "select", start: format.start }); return true; },
          GetStartPage() { return Number(options.searchPage) || 0; },
        });
        offset += needle.length;
      }
      return ranges;
    },
    GetPageCount() { return Number(options.pageCount) || 1; },
    GetCurrentPage() { return currentPage; },
    GetCurrentVisiblePages() { return [currentPage]; },
    GetAllStyles() {
      if (options.omitCreatedStylesFromGetAllStyles) return [];
      return Array.from(styles.values());
    },
    GetStyle(name) { return styles.get(String(name)) || null; },
    CreateStyle(name, type) { return createStyle(name, type); },
    CreateNumbering(kind) {
      return typeof options.createNumbering === "function" ? options.createNumbering(kind) : null;
    },
    GoToPage(index) {
      currentPage = index;
      navigatedPages.push(index);
      return true;
    },
    CreateNewHistoryPoint() {
      historyPoints += 1;
      return true;
    },
    SetAssistantTrackRevisions(value) {
      if (options.setTrackRevisionsAccepted === false) return false;
      revisionTracking = Boolean(value);
      revisionEvents.push({ type: "tracking", value: revisionTracking });
      return true;
    },
    SetTrackRevisions(value) {
      if (options.setTrackRevisionsAccepted === false) return false;
      revisionTracking = Boolean(value);
      revisionEvents.push({ type: "tracking", value: revisionTracking });
      return true;
    },
    IsTrackRevisions() { return revisionTracking; },
    AcceptAllRevisionChanges() {
      revisionEvents.push({ type: "acceptAll" });
      return options.acceptAllRevisionsAccepted !== false;
    },
    RejectAllRevisionChanges() {
      revisionEvents.push({ type: "rejectAll" });
      return options.rejectAllRevisionsAccepted !== false;
    },
  };
  if (options.nativeSearchAndReplace) {
    document.SearchAndReplace = function (properties) {
      nativeSearchAndReplaceCalls.push(JSON.parse(JSON.stringify(properties)));
      if (options.nativeSearchAndReplaceAccepted === false) return false;
      for (const paragraph of documentParagraphs) {
        paragraph.text = replaceLiteral(paragraph.text, properties);
      }
      syncDocumentText();
      return true;
    };
  }
  const officeContext = vm.createContext({
    Asc: { scope },
    AscCommon: {
      CCore: options.coreFactorySupported === false
        ? undefined
        : function () { return createCoreBacking(); },
      CCustomProperties: options.customPropertiesFactorySupported === false
        ? undefined
        : function () { return createCustomPropertiesBacking(); },
    },
    Api: {
      GetDocument: () => document,
      CreateParagraph: () => createParagraph(),
      CreateTable: (rows, columns) => createTable(rows, columns),
      CreateInlineLvlSdt: () => createInlineContentControl(),
      CreateImage: (url, width, height) => {
        if (options.createImageSupported === false) return undefined;
        const image = {
          url,
          width,
          height,
          wrapping: "inline",
          name: "",
          SetWrappingStyle(value) { this.wrapping = value; },
          SetName(value) { this.name = value; },
        };
        createdImages.push(image);
        return image;
      },
      CreateChart: (...args) => {
        const chart = {
          arguments: args,
          SetTitle(value, size) { this.title = { value, size }; },
          SetLegendPos(value) { this.legendPosition = value; },
          SetShowDataLabels(...values) { this.dataLabels = values; },
          SetWrappingStyle(value) { this.wrapping = value; },
        };
        createdCharts.push(chart);
        return chart;
      },
      Color: value => value,
    },
    JSON,
    String,
  });
  const pluginWindow = {};
  const Asc = {
    scope,
    plugin: {
      info: {},
      callCommand(command, close, recalculate, callback) {
        assert.equal(close, false);
        callCommandCalls.push({ command: command.toString(), recalculate: Boolean(recalculate) });
        if (
          !recalculate
          && !capabilityProbeResultUsed
          && Object.prototype.hasOwnProperty.call(options, "capabilityProbeRawResult")
        ) {
          capabilityProbeResultUsed = true;
          callback(options.capabilityProbeRawResult);
          return;
        }
        try {
          callback(vm.runInContext(`(${command.toString()})()`, officeContext));
        } catch (error) {
          callback(JSON.stringify({ ok: false, error: error.message }));
        }
      },
      executeMethod(name, args, callback) {
        executeMethodCalls.push({ name, args: JSON.parse(JSON.stringify(args)) });
        const acceptedByName = options.executeMethodAcceptedByName || {};
        if (
          options.executeMethodAccepted === false
          || acceptedByName[name] === false
        ) return false;
        if (name === "SearchAndReplace") {
          for (const paragraph of documentParagraphs) {
            paragraph.text = replaceLiteral(paragraph.text, args[0]);
          }
          syncDocumentText();
        }
        if (name === "SetDisplayModeInReview") {
          revisionEvents.push({ type: "display", value: args[0] });
        }
        const responses = options.executeMethodResponses || {};
        queueMicrotask(() => callback(Object.prototype.hasOwnProperty.call(responses, name) ? responses[name] : undefined));
        return true;
      },
    },
  };
  if (options.executeMethodSupported === false) delete Asc.plugin.executeMethod;
  pluginWindow.Asc = Asc;
  vm.runInNewContext(wordBridgeSource, { window: pluginWindow, Asc, console, Promise, JSON, String });
  return {
    bridge: pluginWindow.AICopilotBridges.word,
    document,
    executeMethodCalls,
    nativeSearchAndReplaceCalls,
    callCommandCalls,
    insertContentCalls,
    addElementCalls,
    formattedRanges,
    navigatedPages,
    rangeActions,
    createdImages,
    createdTables,
    createdContentControls,
    createdCharts,
    createdSectionParagraphs,
    createdStyles,
    appliedStyles,
    fieldInstructions,
    revisionEvents,
    get text() { return documentText; },
    get paragraphs() { return documentParagraphs.map(paragraph => paragraph.text); },
    get paragraphObjects() { return documentParagraphs; },
    get elements() { return documentElements; },
    get tables() {
      return documentElements.filter(element => (
        element && typeof element.GetRowsCount === "function" && typeof element.GetRow === "function"
      ));
    },
    get historyPoints() { return historyPoints; },
  };
}

test("headless plugin handshakes with one host instance and executes a read-only tool", async () => {
  const { hostWindow } = createHarness();
  await hostWindow.aiBridge.ready({ timeoutMs: 1000 });

  assert.equal(hostWindow.aiBridge.version, publicContract.version);
  assert.equal(hostWindow.aiBridge.editorType, "word");
  assert.equal(hostWindow.aiBridge.context.documentKey, "doc-key-v1");
  assert.ok(hostWindow.aiBridge.capabilities.tools.includes("word_inspect"));

  const result = await hostWindow.aiBridge.word.inspect({}, { timeoutMs: 1000 });
  assert.equal(result.changed, 0);
  assert.equal(result.results[0].name, "word_inspect");
});

test("document-hub documentId is published from signed config or editor URL", async () => {
  const signed = createHarness({
    documentId: "11111111-1111-4111-8111-111111111111",
  });
  await signed.hostWindow.aiBridge.ready({ timeoutMs: 1000 });
  assert.equal(
    signed.hostWindow.aiBridge.context.documentId,
    "11111111-1111-4111-8111-111111111111",
  );

  const fallback = createHarness({
    pathname: "/documents/22222222-2222-4222-8222-222222222222",
  });
  await fallback.hostWindow.aiBridge.ready({ timeoutMs: 1000 });
  assert.equal(
    fallback.hostWindow.aiBridge.context.documentId,
    "22222222-2222-4222-8222-222222222222",
  );
});

test("document-hub relay options survive ONLYOFFICE replacing window globals", async () => {
  const documentId = "33333333-3333-4333-8333-333333333333";
  const harness = createHarness({ optionsDocumentId: documentId });
  harness.hostWindow.aiBridgeOptions = null;

  await harness.hostWindow.aiBridge.ready({ timeoutMs: 1000 });
  assert.equal(harness.hostWindow.aiBridge.context.documentId, documentId);
});

test("strict plugin handshake reaches a document host nested below a cross-origin outer frame", async () => {
  const harness = createHarness({ nestedHost: true });
  await harness.hostWindow.aiBridge.ready({ timeoutMs: 1000 });

  const hellos = harness.ancestorMessages.filter(
    entry => entry.message && entry.message.type === "hello",
  );
  assert.deepEqual(hellos.map(entry => entry.ancestor), ["editor", "host", "outer"]);
  assert.ok(hellos.every(entry => entry.targetOrigin === "https://app.test"));
  assert.ok(hellos.every(
    entry => entry.message.channelId === "11111111-2222-4333-8444-555555555555",
  ));

  const result = await harness.hostWindow.aiBridge.word.inspect({}, { timeoutMs: 1000 });
  assert.equal(result.results[0].name, "word_inspect");
  assert.equal(harness.documentElement.dataset.aiBridgeState, "ready");
});

test("strict plugin locks the first valid host window for all later commands", async () => {
  const harness = createHarness({ nestedHost: true });
  await harness.hostWindow.aiBridge.ready({ timeoutMs: 1000 });
  const configure = harness.hostMessages.find(
    entry => entry.message && entry.message.type === "configure",
  );
  assert.ok(configure);

  harness.pluginWindow.dispatch("message", {
    origin: "https://app.test",
    source: harness.editorProxy,
    data: {
      source: "ai-bridge-host",
      protocolVersion: 1,
      pluginGuid: TEST_PLUGIN_GUID,
      channelId: "11111111-2222-4333-8444-555555555555",
      sessionId: configure.message.sessionId,
      type: "execute",
      requestId: "forged-after-lock",
      method: "executeTool",
      params: { name: "word_inspect", arguments: {} },
    },
  });
  await new Promise(resolve => setTimeout(resolve, 5));

  assert.equal(harness.executedToolCalls.length, 0);
  assert.equal(
    harness.hostMessages.some(
      entry => entry.message && entry.message.requestId === "forged-after-lock",
    ),
    false,
  );
});

test("strict host rejects a legacy channel-less plugin handshake", async () => {
  const harness = createHarness({
    strictHandshake: true,
    ascPluginOptions: {},
    pluginSetInterval() { return 1; },
    pluginClearInterval() {},
  });

  await assert.rejects(
    harness.hostWindow.aiBridge.ready({ timeoutMs: 20 }),
    error => error && error.code === "NOT_READY",
  );
  assert.notEqual(harness.documentElement.dataset.aiBridgeState, "ready");
});

test("HTTP Relay reports an invalid strict handshake before registration", async () => {
  let diagnosticReport = null;
  const harness = createHarness({
    hostname: "office.test",
    httpRelay: true,
    relayBaseUrl: "/api/v1/editor-relay",
    hostPluginOptions: {
      hostOrigin: "https://wrong-host.test",
      channelId: "11111111-2222-4333-8444-555555555555",
    },
    ascPluginOptions: {
      hostOrigin: "https://wrong-host.test",
      channelId: "11111111-2222-4333-8444-555555555555",
    },
    pluginSetInterval() { return 1; },
    pluginClearInterval() {},
    hostFetch: async (requestPath, requestOptions) => {
      assert.ok(requestPath.endsWith("/startup-failure"));
      diagnosticReport = JSON.parse(requestOptions.body);
      return relayResponse(202, { ok: true, diagnosticId: diagnosticReport.diagnosticId });
    },
  });

  await waitFor(() => harness.documentElement.dataset.aiBridgeRelayState === "startup-failed");
  assert.equal(diagnosticReport.code, "PLUGIN_HANDSHAKE_CONFIG_INVALID");
  assert.equal(diagnosticReport.details.stalledPhase, "invalid-handshake-config");
  assert.equal(diagnosticReport.details.environment.handshakeConfigValid, false);
  assert.equal(diagnosticReport.details.environment.handshakeOriginValid, true);
  assert.equal(diagnosticReport.details.environment.handshakeOriginMatches, false);
  assert.equal(
    harness.hostRelayPaths.some(requestPath => requestPath.endsWith("/register")),
    false,
  );
});

test("HTTP Relay distinguishes a rejected plugin handshake from a missing plugin", async () => {
  let startupTimeout = null;
  let diagnosticReport = null;
  const harness = createHarness({
    hostname: "office.test",
    httpRelay: true,
    relayBaseUrl: "/api/v1/editor-relay",
    strictHandshake: true,
    ascPluginOptions: {
      hostOrigin: "https://app.test",
      channelId: "aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee",
    },
    pluginSetInterval() { return 1; },
    pluginClearInterval() {},
    hostSetTimeout(callback, delay) {
      if (delay === 180000) {
        startupTimeout = callback;
        return 180000;
      }
      return setTimeout(callback, delay);
    },
    hostClearTimeout(timer) {
      if (timer !== 180000) clearTimeout(timer);
    },
    hostFetch: async (requestPath, requestOptions) => {
      assert.ok(requestPath.endsWith("/startup-failure"));
      diagnosticReport = JSON.parse(requestOptions.body);
      return relayResponse(202, { ok: true, diagnosticId: diagnosticReport.diagnosticId });
    },
  });

  startupTimeout();
  await waitFor(() => harness.documentElement.dataset.aiBridgeRelayState === "startup-failed");
  assert.equal(diagnosticReport.code, "PLUGIN_HANDSHAKE_REJECTED");
  assert.equal(diagnosticReport.details.stalledPhase, "handshake-rejected");
  assert.equal(diagnosticReport.details.handshakeRejection.reason, "channel-mismatch");
  assert.equal(diagnosticReport.details.handshakeRejection.channelPresent, true);
  assert.equal(Number.isFinite(diagnosticReport.details.startup.pluginMessageSeenMs), true);
  assert.equal(Number.isFinite(diagnosticReport.details.startup.handshakeRejectedMs), true);
});

test("strict plugin ignores wrong origin, wrong channel, and non-ancestor configuration", async () => {
  const harness = createHarness({
    nestedHost: true,
    dropPluginHello: true,
    pluginSetInterval() { return 1; },
    pluginClearInterval() {},
  });
  const configure = overrides => ({
    source: "ai-bridge-host",
    protocolVersion: 1,
    pluginGuid: TEST_PLUGIN_GUID,
    type: "configure",
    sessionId: "session:forged",
    channelId: "11111111-2222-4333-8444-555555555555",
    config: { editorType: "word", documentKey: "doc-key-v1" },
    ...overrides,
  });
  const evilMessages = [];
  const evilWindow = {
    postMessage(message, targetOrigin) {
      evilMessages.push({ message, targetOrigin });
    },
  };
  const messagesBefore = harness.ancestorMessages.length;

  harness.pluginWindow.dispatch("message", {
    origin: "https://evil.test",
    source: harness.hostProxy,
    data: configure({}),
  });
  harness.pluginWindow.dispatch("message", {
    origin: "https://app.test",
    source: harness.hostProxy,
    data: configure({ channelId: "aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee" }),
  });
  harness.pluginWindow.dispatch("message", {
    origin: "https://app.test",
    source: harness.hostProxy,
    data: configure({ channelId: undefined }),
  });
  harness.pluginWindow.dispatch("message", {
    origin: "https://app.test",
    source: evilWindow,
    data: configure({}),
  });

  await new Promise(resolve => setTimeout(resolve, 5));
  assert.equal(harness.ancestorMessages.length, messagesBefore);
  assert.equal(evilMessages.length, 0);
  await assert.rejects(
    harness.hostWindow.aiBridge.ready({ timeoutMs: 20 }),
    error => error && error.code === "NOT_READY",
  );
});

test("plugin reports a missing editor bridge instead of waiting for a generic timeout", async () => {
  const harness = createHarness({
    missingBridge: true,
    pluginSetInterval() { return 1; },
    pluginClearInterval() {},
  });

  await assert.rejects(
    harness.hostWindow.aiBridge.ready({ timeoutMs: 1000 }),
    error => (
      error
      && error.code === "PLUGIN_BRIDGE_MISSING"
      && error.details?.pluginError?.editorType === "word"
      && error.details?.pluginError?.expectedAsset === "bridges/word-bridge.js"
    ),
  );
  assert.equal(harness.documentElement.dataset.aiBridgeState, "startup-error");
  assert.equal(
    harness.ancestorMessages.some(
      entry => entry.message?.type === "startup-error"
        && entry.message.error?.code === "PLUGIN_BRIDGE_MISSING",
    ),
    true,
  );
});

test("plugin exposes browser asset loading failures in startup diagnostics", async () => {
  const harness = createHarness({
    missingBridge: true,
    bootstrapState: {
      loaded: ["bridges/slides-bridge.js", "bridges/sheets-bridge.js"],
      failures: [{
        asset: "bridges/word-bridge.js",
        reason: "network-or-policy-error",
      }],
    },
    pluginSetInterval() { return 1; },
    pluginClearInterval() {},
  });

  await assert.rejects(
    harness.hostWindow.aiBridge.ready({ timeoutMs: 1000 }),
    error => (
      error
      && error.code === "PLUGIN_ASSET_LOAD_FAILED"
      && error.details?.pluginError?.asset === "bridges/word-bridge.js"
      && error.details?.pluginError?.reason === "network-or-policy-error"
    ),
  );
});

test("bootstrap records versioned browser asset loads and failures", async () => {
  const requestedURLs = [];
  const bootstrapWindow = {
    location: {
      href: "https://docs.test/sdkjs-plugins/ai-bridge/index.html",
      origin: "https://docs.test",
      ancestorOrigins: ["https://app.test"],
    },
    Asc: { plugin: { info: { options: {} }, init() {} } },
    setInterval() { return 1; },
  };
  bootstrapWindow.parent = bootstrapWindow;
  bootstrapWindow.top = bootstrapWindow;
  const bootstrapDocument = {
    currentScript: {
      src: "https://docs.test/sdkjs-plugins/ai-bridge/bootstrap.js?v=revision-123",
    },
    createElement() { return {}; },
    head: {
      appendChild(script) {
        requestedURLs.push(script.src);
        queueMicrotask(() => {
          if (new URL(script.src).pathname.endsWith("/bridges/word-bridge.js")) script.onerror();
          else script.onload();
        });
      },
    },
  };

  vm.runInNewContext(bootstrapSource, {
    window: bootstrapWindow,
    document: bootstrapDocument,
    URL,
    console,
  });

  await waitFor(() => requestedURLs.length === 4);
  assert.deepEqual(
    requestedURLs.map(value => new URL(value).searchParams.get("v")),
    ["revision-123", "revision-123", "revision-123", "revision-123"],
  );
  assert.deepEqual(
    Array.from(bootstrapWindow.__aiBridgeBootstrapState.failures, value => ({ ...value })),
    [{ asset: "bridges/word-bridge.js", reason: "network-or-policy-error" }],
  );
  assert.deepEqual(
    Array.from(bootstrapWindow.__aiBridgeBootstrapState.loaded),
    ["bridges/slides-bridge.js", "bridges/sheets-bridge.js", "plugin.js"],
  );
});

test("simplified Chinese Word editor initializes the document proofing language", async () => {
  const harness = createHarness({ interfaceLanguage: "zh", region: "zh-CN" });
  await harness.hostWindow.aiBridge.ready({ timeoutMs: 1000 });

  assert.equal(harness.initializationCommands.length, 1);
  assert.match(harness.initializationCommands[0], /asc_setDefaultLanguage/);
  assert.equal(harness.editorConfig.editorConfig.lang, "zh");
  assert.equal(harness.editorConfig.editorConfig.region, "zh-CN");
});

test("non-Chinese Word editor preserves the document proofing language", async () => {
  const harness = createHarness({ interfaceLanguage: "en", region: "en-US" });
  await harness.hostWindow.aiBridge.ready({ timeoutMs: 1000 });

  assert.equal(harness.initializationCommands.length, 0);
});

test("local Relay reports private startup timings and preserves editor event callbacks", async () => {
  const appContext = { source: "app" };
  const documentContext = { source: "document" };
  const appEvent = { type: "app-ready" };
  const documentEvent = { type: "document-ready" };
  const callbackCalls = [];
  let registerState = null;

  const harness = createHarness({
    hostname: "localhost",
    httpRelay: true,
    relayBaseUrl: "/api/v1/editor-relay",
    editorEvents: {
      onAppReady(event) {
        callbackCalls.push(["app", this, event]);
        return "app-result";
      },
      onDocumentReady(event) {
        callbackCalls.push(["document", this, event]);
        return "document-result";
      },
    },
    beforePluginInit({ editorConfig, hostWindow }) {
      hostWindow.dispatch("load", {});
      assert.equal(editorConfig.events.onAppReady.call(appContext, appEvent), "app-result");
      assert.equal(
        editorConfig.events.onDocumentReady.call(documentContext, documentEvent),
        "document-result",
      );
    },
    hostFetch: async (requestPath, requestOptions) => {
      const payload = JSON.parse(requestOptions.body);
      if (requestPath.endsWith("/register")) {
        registerState = payload.state;
        return relayResponse(200, {
          ok: true,
          relayKey: "relay-key",
          resumeToken: "resume-token",
        });
      }
      if (requestPath.endsWith("/poll")) {
        return relayResponse(409, {
          ok: false,
          error: { code: "SESSION_SUPERSEDED", message: "done" },
        });
      }
      throw new Error(`unexpected Relay request: ${requestPath}`);
    },
  });

  await waitFor(() => registerState !== null);
  await waitFor(() => harness.documentElement.dataset.aiBridgeRelayState === "superseded");

  assert.equal(registerState.documentReady, true);
  assert.equal(registerState.capabilityProbeComplete, true);
  assert.equal(Number.isFinite(registerState.capabilityProbedAt), true);
  assert.equal(registerState.saveReady, true);
  assert.equal(registerState.saveStatus, "idle");
  assert.deepEqual(callbackCalls, [
    ["app", appContext, appEvent],
    ["document", documentContext, documentEvent],
  ]);
  assert.equal("_diagnostics" in harness.hostWindow.aiBridge.getState(), false);
  const startup = registerState._diagnostics.startup;
  for (const field of [
    "hostScriptMs",
    "windowLoadMs",
    "appReadyMs",
    "documentReadyMs",
    "bridgeReadyMs",
  ]) {
    assert.equal(Number.isFinite(startup[field]), true, field);
    assert.ok(startup[field] >= 0, field);
  }
});

test("explicit HTTPS opt-in starts HTTP Relay on a public host", async () => {
  let registerCalls = 0;
  let pollCalls = 0;
  const harness = createHarness({
    hostname: "office.test",
    protocol: "https:",
    httpRelay: true,
    relayBaseUrl: "/api/v1/editor-relay",
    hostFetch: async requestPath => {
      if (requestPath.endsWith("/register")) {
        registerCalls += 1;
        return relayResponse(200, {
          ok: true,
          relayKey: "relay-key",
          resumeToken: "resume-token",
        });
      }
      if (requestPath.endsWith("/poll")) {
        pollCalls += 1;
        return relayResponse(409, {
          ok: false,
          error: {
            code: "SESSION_SUPERSEDED",
            message: "test complete",
          },
        });
      }
      throw new Error(`unexpected Relay request: ${requestPath}`);
    },
  });

  await waitFor(() => harness.documentElement.dataset.aiBridgeRelayState === "superseded");

  assert.equal(registerCalls, 1);
  assert.equal(pollCalls, 1);
  assert.deepEqual(
    harness.hostWindow.dispatchedEvents
      .filter(event => event.type === "ai-bridge-relay-state")
      .map(event => event.detail.code
        ? { state: event.detail.state, code: event.detail.code }
        : { state: event.detail.state }),
    [
      { state: "registering" },
      { state: "ready" },
      { state: "superseded", code: "SESSION_SUPERSEDED" },
    ],
  );
});

test("HTTP Relay reports a fetch rejection with runtime network diagnostics", async () => {
  let pollCalls = 0;
  let runtimeReport = null;
  const harness = createHarness({
    hostname: "localhost",
    httpRelay: true,
    relayBaseUrl: "/api/v1/editor-relay",
    editorSessionId: "33333333-3333-4333-8333-333333333333",
    documentReadyState: "complete",
    pageVisibilityState: "visible",
    hostSetTimeout(callback, delay) {
      return setTimeout(callback, delay === 1000 ? 0 : delay);
    },
    hostFetch: async (requestPath, requestOptions) => {
      if (requestPath.endsWith("/register")) {
        return relayResponse(200, {
          ok: true,
          relayKey: "relay-key",
          resumeToken: "resume-token",
        });
      }
      if (requestPath.endsWith("/runtime-failure")) {
        runtimeReport = JSON.parse(requestOptions.body);
        return relayResponse(202, { ok: true, diagnosticId: runtimeReport.diagnosticId });
      }
      if (requestPath.endsWith("/poll")) {
        pollCalls += 1;
        if (pollCalls === 1) throw new TypeError("Failed to fetch");
        return relayResponse(409, {
          ok: false,
          error: { code: "SESSION_SUPERSEDED", message: "test complete" },
        });
      }
      throw new Error(`unexpected Relay request: ${requestPath}`);
    },
  });

  await waitFor(() => runtimeReport !== null);
  await waitFor(() => harness.documentElement.dataset.aiBridgeRelayState === "superseded");
  const retrying = harness.hostWindow.dispatchedEvents.find(
    event => event.type === "ai-bridge-relay-state" && event.detail.state === "retrying",
  );
  assert.ok(retrying);
  assert.equal(retrying.detail.code, "HTTP_RELAY_NETWORK_ERROR");
  assert.match(retrying.detail.details.diagnosticId, /^diagnostic:/);
  assert.equal(retrying.detail.details.path, "poll");
  assert.equal(retrying.detail.details.classification, "fetch-rejected");
  assert.equal(retrying.detail.details.requestSequence, 2);
  assert.equal(retrying.detail.details.consecutiveFailures, 1);
  assert.equal(retrying.detail.details.errorName, "TypeError");
  assert.equal(retrying.detail.details.network.navigatorOnline, true);
  assert.equal(retrying.detail.details.network.documentReadyState, "complete");
  assert.equal(retrying.detail.details.network.pageVisibilityState, "visible");
  assert.equal(runtimeReport.diagnosticId, retrying.detail.details.diagnosticId);
  assert.equal(runtimeReport.sessionId.startsWith("http-session:"), true);
  assert.equal(runtimeReport.code, "HTTP_RELAY_NETWORK_ERROR");
  assert.deepEqual(runtimeReport.details, JSON.parse(JSON.stringify(retrying.detail.details)));
});

test("HTTP Relay classifies a poll canceled during unload without showing a retry state", async () => {
  let rejectPoll = null;
  let runtimeReport = null;
  const harness = createHarness({
    hostname: "localhost",
    httpRelay: true,
    relayBaseUrl: "/api/v1/editor-relay",
    editorSessionId: "33333333-3333-4333-8333-333333333333",
    hostFetch: async (requestPath, requestOptions) => {
      if (requestPath.endsWith("/register")) {
        return relayResponse(200, {
          ok: true,
          relayKey: "relay-key",
          resumeToken: "resume-token",
        });
      }
      if (requestPath.endsWith("/poll")) {
        return new Promise((resolve, reject) => {
          rejectPoll = reject;
        });
      }
      if (requestPath.endsWith("/unregister")) {
        return relayResponse(200, { ok: true });
      }
      if (requestPath.endsWith("/runtime-failure")) {
        runtimeReport = JSON.parse(requestOptions.body);
        return relayResponse(202, { ok: true, diagnosticId: runtimeReport.diagnosticId });
      }
      throw new Error(`unexpected Relay request: ${requestPath}`);
    },
  });

  await waitFor(() => typeof rejectPoll === "function");
  harness.hostWindow.dispatch("beforeunload", {});
  rejectPoll(new TypeError("Failed to fetch"));
  await waitFor(() => runtimeReport !== null);
  await new Promise(resolve => setTimeout(resolve, 10));

  assert.equal(runtimeReport.details.path, "poll");
  assert.equal(runtimeReport.details.classification, "page-unloading");
  assert.equal(runtimeReport.details.network.lifecycleEvent, "beforeunload");
  assert.equal(runtimeReport.details.network.navigatorOnline, true);
  assert.equal(
    harness.hostWindow.dispatchedEvents.some(
      event => event.type === "ai-bridge-relay-state" && event.detail.state === "retrying",
    ),
    false,
  );
});

test("HTTP Relay keeps registering when the plugin becomes ready after 120 seconds", async () => {
  let now = 1_000_000;
  let startupTimer = null;
  const FakeDate = class extends Date {
    static now() { return now; }
  };
  const harness = createHarness({
    hostname: "localhost",
    httpRelay: true,
    relayBaseUrl: "/api/v1/editor-relay",
    delayPluginHello: true,
    pluginSetInterval: () => 0,
    pluginClearInterval: () => {},
    hostDate: FakeDate,
    hostSetTimeout(callback, delay) {
      if (delay === 180000) {
        startupTimer = { callback, delay };
        return 180000;
      }
      return setTimeout(callback, delay);
    },
    hostClearTimeout(timer) {
      if (timer !== 180000) clearTimeout(timer);
    },
    hostFetch: async requestPath => {
      if (requestPath.endsWith("/register")) {
        return relayResponse(200, {
          ok: true,
          relayKey: "relay-key",
          resumeToken: "resume-token",
        });
      }
      if (requestPath.endsWith("/poll")) {
        return relayResponse(409, {
          ok: false,
          error: { code: "SESSION_SUPERSEDED", message: "done" },
        });
      }
      throw new Error(`unexpected Relay request: ${requestPath}`);
    },
  });

  assert.equal(harness.documentElement.dataset.aiBridgeRelayState, "registering");
  assert.equal(startupTimer.delay, 180000);
  assert.deepEqual(harness.hostRelayPaths, []);

  now += 120000;
  const hello = harness.ancestorMessages.find(entry => entry.message?.type === "hello").message;
  harness.hostWindow.dispatch("message", {
    origin: "https://docs.test",
    source: harness.pluginHandle,
    data: hello,
  });

  await waitFor(() => harness.documentElement.dataset.aiBridgeRelayState === "superseded");
  const registerRequest = harness.hostRelayPaths.find(path => path.endsWith("/register"));
  assert.ok(registerRequest);
  const registration = harness.hostWindow.dispatchedEvents.find(
    event => event.type === "ai-bridge-relay-state" && event.detail.state === "ready",
  );
  assert.ok(registration);
  assert.equal(
    harness.hostWindow.aiBridge.getState().ready,
    true,
  );
  assert.equal(
    harness.hostWindow.dispatchedEvents.some(
      event => event.type === "ai-bridge-relay-state" && event.detail.code === "NOT_READY",
    ),
    false,
  );
});

test("HTTP Relay stops with startup diagnostics after 180 seconds", async () => {
  let now = 1_000_000;
  let startupTimeout = null;
  let diagnosticReport = null;
  const FakeDate = class extends Date {
    static now() { return now; }
  };
  const harness = createHarness({
    hostname: "localhost",
    httpRelay: true,
    relayBaseUrl: "/api/v1/editor-relay",
    dropPluginHello: true,
    pluginSetInterval: () => 0,
    pluginClearInterval: () => {},
    hostDate: FakeDate,
    hostSetTimeout(callback, delay) {
      if (delay === 180000) {
        startupTimeout = callback;
        return 180000;
      }
      return setTimeout(callback, delay);
    },
    hostClearTimeout(timer) {
      if (timer !== 180000) clearTimeout(timer);
    },
    beforePluginInit({ editorConfig }) {
      editorConfig.events.onAppReady({ type: "app-ready" });
      editorConfig.events.onDocumentReady({ type: "document-ready" });
    },
    hostFetch: async (requestPath, requestOptions) => {
      assert.ok(requestPath.endsWith("/startup-failure"));
      diagnosticReport = JSON.parse(requestOptions.body);
      return relayResponse(202, {
        ok: true,
        diagnosticId: diagnosticReport.diagnosticId,
      });
    },
  });

  assert.equal(harness.documentElement.dataset.aiBridgeRelayState, "registering");
  assert.equal(typeof startupTimeout, "function");
  now += 180000;
  startupTimeout();

  await waitFor(() => harness.documentElement.dataset.aiBridgeRelayState === "startup-failed");
  await waitFor(() => diagnosticReport !== null);
  const detail = harness.hostWindow.dispatchedEvents
    .filter(event => event.type === "ai-bridge-relay-state")
    .at(-1).detail;
  assert.equal(detail.code, "PLUGIN_NOT_LOADED");
  assert.equal(detail.details.timeoutMs, 180000);
  assert.equal(detail.details.stalledPhase, "waiting-plugin-hello");
  assert.equal(detail.details.startup.hostScriptMs, 0);
  assert.equal(detail.details.startup.appReadyMs, 0);
  assert.equal(detail.details.startup.documentReadyMs, 0);
  assert.match(detail.details.diagnosticId, /^diagnostic:/);
  assert.equal(diagnosticReport.code, "PLUGIN_NOT_LOADED");
  assert.equal(diagnosticReport.details.diagnosticId, detail.details.diagnosticId);
  assert.deepEqual(
    harness.hostRelayPaths.map(requestPath => requestPath.split("/").at(-1)),
    ["startup-failure"],
  );
});

test("HTTP Relay distinguishes an editor app startup timeout", async () => {
  let startupTimeout = null;
  let diagnosticReport = null;
  const harness = createHarness({
    hostname: "localhost",
    httpRelay: true,
    relayBaseUrl: "/api/v1/editor-relay",
    dropPluginHello: true,
    pluginSetInterval: () => 0,
    pluginClearInterval: () => {},
    hostSetTimeout(callback, delay) {
      if (delay === 180000) {
        startupTimeout = callback;
        return 180000;
      }
      return setTimeout(callback, delay);
    },
    hostClearTimeout(timer) {
      if (timer !== 180000) clearTimeout(timer);
    },
    hostFetch: async (requestPath, requestOptions) => {
      assert.ok(requestPath.endsWith("/startup-failure"));
      diagnosticReport = JSON.parse(requestOptions.body);
      return relayResponse(202, { ok: true, diagnosticId: diagnosticReport.diagnosticId });
    },
  });

  startupTimeout();
  await waitFor(() => harness.documentElement.dataset.aiBridgeRelayState === "startup-failed");
  assert.equal(diagnosticReport.code, "EDITOR_APP_STARTUP_TIMEOUT");
  assert.equal(diagnosticReport.details.stalledPhase, "waiting-editor-app");
  assert.equal(diagnosticReport.details.environment.documentReadyState, "unknown");
});

test("contract mismatch is terminal and reports one Relay state without registering", async () => {
  let diagnosticReport = null;
  const harness = createHarness({
    hostname: "localhost",
    httpRelay: true,
    relayBaseUrl: "/api/v1/editor-relay",
    dropPluginHello: true,
    pluginSetInterval: () => 0,
    pluginClearInterval: () => {},
    hostFetch: async (requestPath, requestOptions) => {
      assert.ok(requestPath.endsWith("/startup-failure"));
      diagnosticReport = JSON.parse(requestOptions.body);
      return relayResponse(202, { ok: true, diagnosticId: diagnosticReport.diagnosticId });
    },
  });
  const mismatchMessage = {
    source: "ai-bridge-plugin",
    pluginGuid: TEST_PLUGIN_GUID,
    protocolVersion: 1,
    contractSha256: "0".repeat(64),
  };

  harness.hostWindow.dispatch("message", {
    origin: "https://docs.test",
    source: harness.pluginHandle,
    data: mismatchMessage,
  });
  harness.hostWindow.dispatch("message", {
    origin: "https://docs.test",
    source: harness.pluginHandle,
    data: mismatchMessage,
  });

  await waitFor(() => harness.documentElement.dataset.aiBridgeRelayState === "contract-mismatch");
  await assert.rejects(
    harness.hostWindow.aiBridge.ready({ timeoutMs: 1000 }),
    error => error && error.code === "CONTRACT_MISMATCH",
  );
  await new Promise(resolve => setTimeout(resolve, 25));

  assert.equal(harness.documentElement.dataset.aiBridgeState, "contract-mismatch");
  assert.deepEqual(
    harness.hostRelayPaths.map(requestPath => requestPath.split("/").at(-1)),
    ["startup-failure"],
  );
  assert.equal(diagnosticReport.code, "CONTRACT_MISMATCH");
  assert.equal(diagnosticReport.details.stalledPhase, "contract-mismatch");
  assert.deepEqual(
    harness.hostWindow.dispatchedEvents
      .filter(event => event.type === "ai-bridge-relay-state")
      .map(event => event.detail.code
        ? { state: event.detail.state, code: event.detail.code }
        : { state: event.detail.state }),
    [
      { state: "registering" },
      { state: "contract-mismatch", code: "CONTRACT_MISMATCH" },
    ],
  );
});

test("Relay contract version mismatch is terminal and preserves server diagnostics", async () => {
  let registerCalls = 0;
  let diagnosticReport = null;
  const mismatchDetails = {
    expectedContractVersion: "0.2.2",
    expectedContractSha256: "a".repeat(64),
    receivedContractVersion: "0.2.2",
    receivedContractSha256: "b".repeat(64),
  };
  const harness = createHarness({
    hostname: "localhost",
    httpRelay: true,
    relayBaseUrl: "/api/v1/editor-relay",
    hostFetch: async (requestPath, requestOptions) => {
      if (requestPath.endsWith("/register")) {
        registerCalls += 1;
        return relayResponse(409, {
          ok: false,
          error: {
            code: "CONTRACT_VERSION_MISMATCH",
            message: "编辑器页面与 HTTP Relay 使用了不同的 ai-bridge 契约",
            details: mismatchDetails,
          },
        });
      }
      assert.ok(requestPath.endsWith("/startup-failure"));
      diagnosticReport = JSON.parse(requestOptions.body);
      return relayResponse(202, { ok: true, diagnosticId: diagnosticReport.diagnosticId });
    },
  });

  await waitFor(() => harness.documentElement.dataset.aiBridgeRelayState === "contract-mismatch");
  await waitFor(() => diagnosticReport !== null);

  assert.equal(registerCalls, 1);
  assert.equal(harness.hostRelayPaths.length, 2);
  assert.equal(diagnosticReport.code, "CONTRACT_VERSION_MISMATCH");
  const terminalDetail = JSON.parse(JSON.stringify(
    harness.hostWindow.dispatchedEvents
      .filter(event => event.type === "ai-bridge-relay-state")
      .at(-1).detail,
  ));
  assert.deepEqual(
    terminalDetail,
    {
      state: "contract-mismatch",
      code: "CONTRACT_VERSION_MISMATCH",
      message: "编辑器页面与 HTTP Relay 使用了不同的 ai-bridge 契约",
      details: {
        ...mismatchDetails,
        diagnosticId: diagnosticReport.diagnosticId,
        httpStatus: 409,
        path: "register",
      },
    },
  );
});

test("HTTP Relay does not start without an explicit base path", async () => {
  const harness = createHarness({
    hostname: "office.test",
    protocol: "http:",
    httpRelay: true,
    hostFetch: async () => {
      throw new Error("Relay without a configured base path must not start");
    },
  });

  await harness.hostWindow.aiBridge.ready({ timeoutMs: 1000 });
  await new Promise(resolve => setTimeout(resolve, 25));

  assert.deepEqual(harness.hostRelayPaths, []);
  assert.equal(harness.documentElement.dataset.aiBridgeRelayState, undefined);
});

test("superseded HTTP Relay stops without re-registering or reloading", async () => {
  let registerCalls = 0;
  let pollCalls = 0;
  const harness = createHarness({
    hostname: "localhost",
    httpRelay: true,
    relayBaseUrl: "/api/v1/editor-relay",
    hostFetch: async requestPath => {
      if (requestPath.endsWith("/register")) {
        registerCalls += 1;
        return relayResponse(200, {
          ok: true,
          relayKey: "relay-key",
          resumeToken: "resume-token",
        });
      }
      if (requestPath.endsWith("/poll")) {
        pollCalls += 1;
        return relayResponse(409, {
          ok: false,
          error: {
            code: "SESSION_SUPERSEDED",
            message: "new page is authoritative",
          },
        });
      }
      throw new Error(`unexpected Relay request: ${requestPath}`);
    },
  });

  await waitFor(() => harness.documentElement.dataset.aiBridgeRelayState === "superseded");
  const requestCount = harness.hostRelayPaths.length;
  await new Promise(resolve => setTimeout(resolve, 25));

  assert.equal(registerCalls, 1);
  assert.equal(pollCalls, 1);
  assert.equal(harness.hostRelayPaths.length, requestCount);
  assert.equal(harness.reloadCount, 0);
});

test("first invalid editor session schedules only one page reload", async () => {
  const harness = createHarness({
    hostname: "localhost",
    httpRelay: true,
    relayBaseUrl: "/api/v1/editor-relay",
    hostFetch: async requestPath => {
      assert.ok(requestPath.endsWith("/register"));
      return relayResponse(403, {
        ok: false,
        error: {
          code: "editor_session_invalid",
          message: "invalid session",
        },
      });
    },
  });

  await waitFor(() => harness.reloadCount === 1);

  assert.equal(harness.documentElement.dataset.aiBridgeRelayState, "credential-reload");
  assert.ok(
    [...harness.localValues.entries()].some(
      ([key, value]) => key.startsWith("aiBridgeCredentialReloadAt:") && Number(value) > 0,
    ),
  );
  assert.equal(harness.hostRelayPaths.length, 1);
});

test("repeated invalid editor session inside the guard window becomes terminal", async () => {
  const harness = createHarness({
    hostname: "localhost",
    httpRelay: true,
    relayBaseUrl: "/api/v1/editor-relay",
    localValues: {
      "aiBridgeCredentialReloadAt:demo.docx|docx|word|user-1": String(Date.now()),
    },
    hostFetch: async requestPath => {
      assert.ok(requestPath.endsWith("/register"));
      return relayResponse(403, {
        ok: false,
        error: {
          code: "editor_session_invalid",
          message: "invalid session again",
        },
      });
    },
  });

  await waitFor(() => harness.documentElement.dataset.aiBridgeRelayState === "credential-error");
  const requestCount = harness.hostRelayPaths.length;
  await new Promise(resolve => setTimeout(resolve, 75));

  assert.equal(harness.reloadCount, 0);
  assert.equal(harness.hostRelayPaths.length, requestCount);
});

test("successful Relay registration clears the credential reload guard", async () => {
  const guardKey = "aiBridgeCredentialReloadAt:demo.docx|docx|word|user-1";
  const harness = createHarness({
    hostname: "localhost",
    httpRelay: true,
    relayBaseUrl: "/api/v1/editor-relay",
    localValues: { [guardKey]: String(Date.now()) },
    sessionValues: { [guardKey]: String(Date.now()) },
    hostFetch: async requestPath => {
      if (requestPath.endsWith("/register")) {
        return relayResponse(200, {
          ok: true,
          relayKey: "relay-key",
          resumeToken: "resume-token",
        });
      }
      if (requestPath.endsWith("/poll")) {
        return relayResponse(409, {
          ok: false,
          error: { code: "SESSION_SUPERSEDED", message: "done" },
        });
      }
      throw new Error(`unexpected Relay request: ${requestPath}`);
    },
  });

  await waitFor(() => harness.documentElement.dataset.aiBridgeRelayState === "superseded");
  assert.equal(harness.localValues.has(guardKey), false);
  assert.equal(harness.sessionValues.has(guardKey), false);
});

test("plugin rejects a command when the host document key changes", async () => {
  const { hostWindow, editorConfig } = createHarness();
  await hostWindow.aiBridge.ready({ timeoutMs: 1000 });
  editorConfig.document.key = "another-doc-key";

  await assert.rejects(hostWindow.aiBridge.word.inspect({}, {
    timeoutMs: 1000,
    requestId: "mismatch-request",
  }), error => {
    assert.equal(error.name, "AiBridgeError");
    assert.equal(error.code, "DOCUMENT_MISMATCH");
    assert.equal(error.requestId, "mismatch-request");
    assert.match(error.message, /命令目标不是当前文档/);
    return true;
  });
});

test("plugin aggregates generated schema and bridge semantic errors before checkpoint", async () => {
  const toolCalls = [
    {
      name: "word_append_paragraph",
      arguments: { text: 123, internalNote: "敏感正文" },
    },
    {
      name: "word_manage_section",
      arguments: { action: "create" },
    },
    {
      name: "word_insert_page_break",
      arguments: {},
    },
  ];
  const harness = createHarness({
    bridgePreflight(calls) {
      return {
        toolCalls: calls,
        argumentNormalizations: [],
        validationErrors: [
          {
            toolCallIndex: 1,
            tool: "word_manage_section",
            path: "arguments.paragraphIndex",
            keyword: "semantic",
            message: "paragraphIndex is required when action is create",
          },
          {
            toolCallIndex: 2,
            tool: "word_insert_page_break",
            path: "arguments.target",
            keyword: "semantic",
            message: "paragraphIndexes, search, all=true, or current=true is required",
          },
        ],
      };
    },
  });
  await harness.hostWindow.aiBridge.ready({ timeoutMs: 1000 });

  await assert.rejects(
    harness.hostWindow.aiBridge.executeBatch(toolCalls, {
      timeoutMs: 1000,
      requestId: "direct-preflight-invalid",
    }),
    error => {
      assert.equal(error.code, "INVALID_TOOL_ARGUMENTS");
      assert.equal(error.details.completedToolCalls, 0);
      assert.equal(error.details.partialMutationPossible, false);
      assert.deepEqual(
        Array.from(error.details.validationErrors, item => [
          item.toolCallIndex,
          item.path,
          item.keyword,
        ]),
        [
          [0, "arguments.internalNote", "additionalProperties"],
          [0, "arguments.text", "type"],
          [1, "arguments", "anyOf"],
          [1, "arguments.paragraphIndex", "semantic"],
          [2, "arguments.target", "semantic"],
        ],
      );
      assert.doesNotMatch(JSON.stringify(error.details), /敏感正文|internalNote.*敏感/);
      return true;
    },
  );

  assert.equal(harness.executedToolCalls.length, 0);
  assert.deepEqual(harness.servicePaths, []);
});

test("plugin rejects unsafe sheets_filter batches before checkpoint or execution", async () => {
  const cases = [
    [{
      name: "sheets_filter",
      arguments: { action: "set", range: "A1:B3" },
    }],
    [
      { name: "sheets_inspect", arguments: {} },
      {
        name: "sheets_filter",
        arguments: { action: "set", range: "A1:B3" },
      },
    ],
  ];

  for (let index = 0; index < cases.length; index += 1) {
    const harness = createHarness({ editorType: "cell" });
    await harness.hostWindow.aiBridge.ready({ timeoutMs: 1000 });
    await assert.rejects(
      harness.hostWindow.aiBridge.executeBatch(cases[index], {
        timeoutMs: 1000,
        requestId: `unsafe-filter-${index}`,
      }),
      error => {
        assert.equal(error.code, "INVALID_TOOL_ARGUMENTS");
        assert.equal(error.details.completedToolCalls, 0);
        assert.equal(error.details.partialMutationPossible, false);
        return true;
      },
    );
    assert.equal(harness.executedToolCalls.length, 0);
    assert.deepEqual(harness.servicePaths, []);
  }
});

test("plugin preserves all 30 ordered calls but keeps the public 20-call error message", async () => {
  const harness = createHarness();
  await harness.hostWindow.aiBridge.ready({ timeoutMs: 1000 });
  const calls = Array.from({ length: 30 }, (_unused, index) => ({
    name: "word_inspect",
    arguments: { maxChars: 500 + index },
  }));

  await harness.hostWindow.aiBridge.executeBatch(calls, {
    timeoutMs: 1000,
    requestId: "batch-tolerance-30",
  });
  assert.deepEqual(
    harness.executedToolCalls.map(({ name, arguments: args }) => ({ name, arguments: args })),
    calls,
  );

  await assert.rejects(
    harness.hostWindow.aiBridge.executeBatch(
      [...calls, { name: "word_inspect", arguments: { maxChars: 530 } }],
      { timeoutMs: 1000, requestId: "batch-tolerance-31" },
    ),
    error => {
      assert.equal(error.code, "TOO_MANY_CALLS");
      assert.match(error.message, /20/);
      return true;
    },
  );
  assert.deepEqual(
    harness.executedToolCalls.map(({ name, arguments: args }) => ({ name, arguments: args })),
    calls,
  );
});

test("public contract, plugin allow-lists, and all convenience methods stay aligned", async () => {
  const cases = {
    word: {
      group: "word",
      methods: {
        inspect: ["word_inspect", {}],
        replaceText: ["word_replace_text", { search: "a", replace: "b" }],
        appendParagraph: ["word_append_paragraph", { text: "a" }],
        insertParagraph: ["word_insert_paragraph", { text: "a" }],
        formatDocument: ["word_format_document", { bold: true }],
        formatSelection: ["word_format_selection", { italic: true }],
        formatMatches: ["word_format_matches", { search: "a", bold: true }],
        deleteMatches: ["word_delete_matches", { search: "a" }],
        addHyperlink: ["word_add_hyperlink", { search: "a", url: "https://example.test" }],
        addComment: ["word_add_comment", { search: "a", text: "review" }],
        addBookmark: ["word_add_bookmark", { search: "a", occurrence: 1, name: "a1" }],
        addImage: ["word_add_image", { source: { type: "url", url: "https://images.test/a.png" } }],
        inspectAdvanced: ["word_inspect_advanced", {}],
        setDocumentProperties: ["word_set_document_properties", { title: "Sample" }],
        manageSection: ["word_manage_section", { action: "configure", sectionIndex: 1 }],
        manageStyle: ["word_manage_style", { action: "create", name: "Sample", type: "paragraph" }],
        setTabs: ["word_set_tabs", { tabs: [] }],
        setNumbering: ["word_set_numbering", { paragraphIndexes: [1], kind: "bullet" }],
        formatTableAdvanced: ["word_format_table_advanced", { tableIndex: 1 }],
        addNestedTable: ["word_add_nested_table", {
          tableIndex: 1,
          row: 1,
          column: 1,
          rows: 1,
          cols: 1,
        }],
        manageDrawing: ["word_manage_drawing", { action: "delete", drawingIndex: 1 }],
        addShape: ["word_add_shape", { shapeType: "rect" }],
        addChart: ["word_add_chart", { data: [[1]] }],
        addMath: ["word_add_math", { equation: "x" }],
        addOleObject: ["word_add_ole_object", {
          source: { type: "url", url: "https://images.test/ole.png" },
          data: "payload",
          applicationId: "sample",
        }],
        manageFields: ["word_manage_fields", { action: "updateAll" }],
        manageLongDocument: ["word_manage_long_document", { action: "updateToc" }],
        manageComments: ["word_manage_comments", { action: "removeAll" }],
        manageRevisions: ["word_manage_revisions", { action: "stop" }],
        setProtection: ["word_set_protection", { action: "unprotect" }],
        manageContentControl: ["word_manage_content_control", { action: "add", kind: "checkbox" }],
        manageCustomXml: ["word_manage_custom_xml", { action: "add", xml: "<sample/>" }],
        inspectMacros: ["word_inspect_macros", {}],
        setMacros: ["word_set_macros", { macros: [] }],
        setWatermark: ["word_set_watermark", { action: "remove" }],
        formatParagraphs: ["word_format_paragraphs", { paragraphIndexes: [1], headingLevel: 1 }],
        setParagraphText: ["word_set_paragraph_text", { paragraphIndexes: [1], text: "a" }],
        deleteParagraphs: ["word_delete_paragraphs", { paragraphIndexes: [1] }],
        setList: ["word_set_list", { paragraphIndexes: [1], listType: "bullet" }],
        insertPageBreak: ["word_insert_page_break", { current: true }],
        navigate: ["word_navigate", { target: "end" }],
        scroll: ["word_scroll", { direction: "down", pages: 1 }],
        scaleFont: ["word_scale_font", { scale: 0.9 }],
        addTable: ["word_add_table", { rows: 1, cols: 1 }],
        setTableCell: ["word_set_table_cell", { tableIndex: 1, row: 1, column: 1, text: "a" }],
        formatTable: ["word_format_table", { tableIndex: 1, styleName: "Bordered" }],
        editTable: ["word_edit_table", { tableIndex: 1, action: "addRow" }],
        setPageLayout: ["word_set_page_layout", { pageSize: "A4" }],
        setHeaderFooter: ["word_set_header_footer", { kind: "footer", pageNumber: true }],
        setDocumentText: ["word_set_document_text", { text: "a" }],
      },
    },
    slide: {
      group: "slides",
      methods: {
        inspect: ["slides_inspect", {}],
        inspectObjects: ["slides_inspect_objects", {}],
        validateLayout: ["slides_validate_layout", { slide: 1 }],
        replaceText: ["slides_replace_text", { search: "a", replace: "b" }],
        scaleFont: ["slides_scale_font", { scale: 0.9 }],
        formatText: ["slides_format_text", { bold: true }],
        formatSelection: ["slides_format_selection", { italic: true }],
        addSlide: ["slides_add_slide", { title: "a" }],
        duplicateSlide: ["slides_duplicate_slide", { slide: 1 }],
        deleteSlide: ["slides_delete_slide", { slide: 1 }],
        addTextBox: ["slides_add_textbox", { slide: 1, text: "a" }],
        addImage: ["slides_add_image", { slide: 1, source: { type: "dataUrl", dataUrl: "data:image/png;base64,AAAA" } }],
        setBackground: ["slides_set_background", { slide: 1, fill: { type: "solid", color: "#FFFFFF" } }],
        addShape: ["slides_add_shape", { slide: 1, shapeType: "rect" }],
        updateShape: ["slides_update_shape", { slide: 1, objectIndex: 1, fill: { type: "none" } }],
        deleteObject: ["slides_delete_object", { slide: 1, objectIndex: 1 }],
        inspectCharts: ["slides_inspect_charts", {}],
        addChart: ["slides_add_chart", { slide: 1, series: [[1]], seriesNames: ["S1"], categories: ["C1"] }],
        updateChart: ["slides_update_chart", { slide: 1, chartIndex: 1, title: "T" }],
        deleteChart: ["slides_delete_chart", { slide: 1, chartIndex: 1 }],
      },
    },
    cell: {
      group: "sheets",
      methods: {
        inspect: ["sheets_inspect", {}],
        setValues: ["sheets_set_values", { range: "A1", values: 1 }],
        setFormula: ["sheets_set_formula", { range: "A1", formula: "=1" }],
        replaceText: ["sheets_replace_text", { search: "a", replace: "b" }],
        formatRange: ["sheets_format_range", { range: "A1", bold: true }],
        addSheet: ["sheets_add_sheet", { name: "S2" }],
        renameSheet: ["sheets_rename_sheet", { newName: "S2" }],
        deleteSheet: ["sheets_delete_sheet", { sheet: "S2" }],
        addChart: ["sheets_add_chart", { range: "A1:B2" }],
        inspectCharts: ["sheets_inspect_charts", {}],
        updateChart: ["sheets_update_chart", { chartIndex: 1, title: "T" }],
        deleteChart: ["sheets_delete_chart", { chartIndex: 1 }],
        inspectRange: ["sheets_inspect_range", { range: "A1:B2" }],
        setArrayFormula: ["sheets_set_array_formula", { range: "A1:A2", formula: "=ROW(A1:A2)" }],
        manageSheet: ["sheets_manage_sheet", { action: "setActive" }],
        manageRange: ["sheets_manage_range", { action: "clear", range: "A1" }],
        setRichText: ["sheets_set_rich_text", {
          range: "A1",
          text: "a",
          runs: [{ start: 0, length: 1, bold: true }],
        }],
        inspectNames: ["sheets_inspect_names", {}],
        manageNames: ["sheets_manage_names", {
          action: "add",
          name: "SampleName",
          refersTo: "=Sheet1!$A$1",
        }],
        recalculate: ["sheets_recalculate", {}],
        sort: ["sheets_sort", {
          range: "A1:B2",
          keys: [{ range: "A1:A2", order: "xlAscending" }],
        }],
        filter: ["sheets_filter", { action: "showAll" }],
        inspectTables: ["sheets_inspect_tables", {}],
        manageTable: ["sheets_manage_table", { action: "create", range: "A1:B2" }],
        manageConditionalFormat: ["sheets_manage_conditional_format", {
          action: "deleteAll",
          range: "A1:B2",
        }],
        manageValidation: ["sheets_manage_validation", {
          action: "delete",
          range: "A1:B2",
        }],
        inspectPivots: ["sheets_inspect_pivots", {}],
        managePivot: ["sheets_manage_pivot", { action: "refreshAll" }],
        inspectDrawings: ["sheets_inspect_drawings", {}],
        manageDrawing: ["sheets_manage_drawing", { action: "addShape", shapeType: "rect" }],
        manageHyperlink: ["sheets_manage_hyperlink", {
          action: "set",
          range: "A1",
          url: "https://example.test",
        }],
        inspectComments: ["sheets_inspect_comments", {}],
        manageComments: ["sheets_manage_comments", {
          action: "add",
          range: "A1",
          text: "review",
        }],
        inspectFreezePanes: ["sheets_inspect_freeze_panes", {}],
        manageFreezePanes: ["sheets_manage_freeze_panes", { action: "unfreeze" }],
        inspectProperties: ["sheets_inspect_properties", {}],
        manageProperties: ["sheets_manage_properties", { core: { title: "Sample" } }],
        inspectProtectedRanges: ["sheets_inspect_protected_ranges", {}],
        manageProtectedRanges: ["sheets_manage_protected_ranges", {
          action: "add",
          title: "InputArea",
          range: "A1:B2",
        }],
        inspectPageLayout: ["sheets_inspect_page_layout", {}],
        managePageLayout: ["sheets_manage_page_layout", { orientation: "landscape" }],
        inspectMacros: ["sheets_inspect_macros", {}],
        setMacros: ["sheets_set_macros", { content: { macros: [], current: -1 } }],
      },
    },
  };

  for (const [editorType, entry] of Object.entries(cases)) {
    const harness = createHarness({
      editorType,
      hostFetch: async requestPath => {
        assert.equal(requestPath, "/api/v1/editor-relay/images/import");
        return importedImageResponse();
      },
    });
    await harness.hostWindow.aiBridge.ready({ timeoutMs: 1000 });
    for (const [method, [toolName, args]] of Object.entries(entry.methods)) {
      const result = await harness.hostWindow.aiBridge[entry.group][method](args, {
        timeoutMs: 1000,
        requestId: `${editorType}:${method}`,
      });
      assert.equal(result.results[0].name, toolName);
    }
    assert.deepEqual(
      harness.executedToolCalls.map(call => call.name),
      Object.values(entry.methods).map(([name]) => name),
    );
    assert.deepEqual(
      Array.from(harness.hostWindow.aiBridge.capabilities.tools).sort(),
      Object.keys(publicContract.tools[editorType]).sort(),
    );
  }
});

test("public Word contract exposes positioned tables and inline content replacement", () => {
  const tableProperties = publicContract.tools.word.word_add_table.properties;
  assert.deepEqual(tableProperties.insertAt.enum, ["end", "current", "before", "after"]);
  assert.equal(tableProperties.paragraphIndex.minimum, 1);
  assert.equal(tableProperties.tableIndex.minimum, 1);
  assert.deepEqual(tableProperties.matchMode.enum, ["contains", "exact"]);
  assert.equal(tableProperties.pageBreakBefore.type, "boolean");

  const controlProperties = publicContract.tools.word.word_manage_content_control.properties;
  assert.deepEqual(controlProperties.contentMode.enum, ["append", "replace"]);
  assert.equal(controlProperties.tableIndex.minimum, 1);
  assert.equal(controlProperties.row.minimum, 1);
  assert.equal(controlProperties.column.minimum, 1);
  assert.equal(
    publicContract.tools.word.word_insert_page_break.properties.all.type,
    "boolean",
  );
});

test("public table contracts expose scalar and formatted cell inputs", () => {
  assert.deepEqual(
    publicContract.$defs.tableCellScalar.type,
    ["string", "number", "boolean", "null"],
  );
  assert.deepEqual(publicContract.$defs.wordTableCell.required, ["text"]);
  assert.equal(publicContract.$defs.wordTableCell.additionalProperties, false);
  assert.equal(publicContract.$defs.wordTableCell.properties.text.type, "string");
  assert.equal(publicContract.$defs.wordTableCell.properties.textColor.deprecated, true);
  assert.equal(
    publicContract.$defs.wordTableCell.properties.textColor["x-canonicalProperty"],
    "color",
  );
  assert.deepEqual(publicContract.$defs.slideTableCell.required, ["text"]);
  assert.deepEqual(
    publicContract.tools.word.word_add_table.properties.align.enum,
    ["left", "center", "right"],
  );
  assert.deepEqual(
    publicContract.tools.word.word_add_nested_table.properties.align.enum,
    ["left", "center", "right"],
  );

  const wordCell = publicContract.tools.word.word_add_table
    .properties.data.items.items.anyOf;
  const nestedWordCell = publicContract.tools.word.word_add_nested_table
    .properties.data.items.items.anyOf;
  const slideCell = publicContract.tools.slide.slides_add_table
    .properties.data.items.items.anyOf;
  assert.deepEqual(wordCell, [
    { $ref: "#/$defs/tableCellScalar" },
    { $ref: "#/$defs/wordTableCell" },
  ]);
  assert.deepEqual(nestedWordCell, wordCell);
  assert.deepEqual(slideCell, [
    { $ref: "#/$defs/tableCellScalar" },
    { $ref: "#/$defs/slideTableCell" },
  ]);
});

test("public Word contract makes paragraph point units explicit", () => {
  for (const tool of [
    "word_append_paragraph",
    "word_insert_paragraph",
    "word_format_document",
    "word_manage_style",
    "word_format_paragraphs",
    "word_set_paragraph_text",
    "word_set_table_cell",
  ]) {
    const properties = publicContract.tools.word[tool].properties;
    assert.match(properties.leftIndentPt.description, /points \(pt\)/);
    assert.match(properties.firstLineIndentPt.description, /points \(pt\)/);
    assert.equal(properties.leftIndent.deprecated, true);
    assert.match(properties.leftIndent.description, /Never pass twips/);
  }
  assert.match(publicContract.unitConventions.word, /Never pass OOXML twips or EMU/);
});

test("Word schema, TypeScript arguments, and runtime paragraph fields stay aligned", () => {
  const declarations = fs.readFileSync(
    path.join(__dirname, "public-api.d.ts"),
    "utf8",
  );
  const interfaces = new Map();
  const declarationPattern = /export interface\s+(\w+)(?:\s+extends\s+([^{]+))?\s*\{/g;
  let match;
  while ((match = declarationPattern.exec(declarations))) {
    let depth = 1;
    let cursor = declarationPattern.lastIndex;
    while (cursor < declarations.length && depth) {
      if (declarations[cursor] === "{") depth += 1;
      else if (declarations[cursor] === "}") depth -= 1;
      cursor += 1;
    }
    const body = declarations.slice(declarationPattern.lastIndex, cursor - 1);
    const fields = new Set();
    let member = "";
    let memberDepth = 0;
    for (const character of body) {
      if (character === "{") memberDepth += 1;
      if (character === "}") memberDepth -= 1;
      member += character;
      if (character === ";" && memberDepth === 0) {
        const field = member.match(/([A-Za-z]\w*)\??\s*:/);
        if (field) fields.add(field[1]);
        member = "";
      }
    }
    interfaces.set(match[1], {
      extends: String(match[2] || "")
        .split(",")
        .map(value => value.trim())
        .filter(Boolean),
      fields,
    });
    declarationPattern.lastIndex = cursor;
  }

  function flattenedFields(name, seen = new Set()) {
    if (seen.has(name)) return new Set();
    seen.add(name);
    const declaration = interfaces.get(name);
    assert.ok(declaration, "missing TypeScript interface " + name);
    const fields = new Set(declaration.fields);
    for (const parent of declaration.extends) {
      for (const field of flattenedFields(parent, seen)) fields.add(field);
    }
    return fields;
  }

  const alignment = {
    word_append_paragraph: "WordAppendParagraphArgs",
    word_insert_paragraph: "WordInsertParagraphArgs",
    word_format_document: "WordFormatDocumentArgs",
    word_add_bookmark: "WordAddBookmarkArgs",
    word_manage_style: "WordManageStyleArgs",
    word_set_tabs: "WordSetTabsArgs",
    word_set_numbering: "WordSetNumberingArgs",
    word_format_paragraphs: "WordFormatParagraphsArgs",
    word_set_paragraph_text: "WordSetParagraphTextArgs",
    word_set_list: "WordSetListArgs",
    word_add_table: "WordAddTableArgs",
    word_set_table_cell: "WordSetTableCellArgs",
    word_add_nested_table: "WordAddNestedTableArgs",
  };
  for (const [tool, interfaceName] of Object.entries(alignment)) {
    assert.deepEqual(
      [...flattenedFields(interfaceName)].sort(),
      Object.keys(publicContract.tools.word[tool].properties || {}).sort(),
      tool + " must expose the same fields in JSON schema and TypeScript",
    );
  }

  const wordBridgeSource = fs.readFileSync(
    path.join(__dirname, "bridges", "word-bridge.js"),
    "utf8",
  );
  const runtimeTextMatch = wordBridgeSource.match(
    /var textFormatKeys = \[([\s\S]*?)\];/,
  );
  const runtimeParagraphMatch = wordBridgeSource.match(
    /var paragraphFormatKeys = textFormatKeys\.concat\(\[([\s\S]*?)\]\);/,
  );
  assert.ok(runtimeTextMatch);
  assert.ok(runtimeParagraphMatch);
  const runtimeFields = [
    ...runtimeTextMatch[1].matchAll(/"([^"]+)"/g),
    ...runtimeParagraphMatch[1].matchAll(/"([^"]+)"/g),
  ].map(value => value[1]);
  for (const tool of [
    "word_append_paragraph",
    "word_insert_paragraph",
    "word_format_document",
    "word_format_paragraphs",
    "word_set_paragraph_text",
    "word_set_table_cell",
  ]) {
    const schemaFields = publicContract.tools.word[tool].properties;
    for (const field of runtimeFields) {
      assert.ok(schemaFields[field], tool + " schema is missing runtime field " + field);
    }
  }
});

test("DOCX capability matrix covers D01 through D70 exactly once", () => {
  const codes = Array.from(
    docxCapabilityMatrix.matchAll(/^\| D(\d{2}) \|/gm),
    match => Number(match[1]),
  );
  assert.deepEqual(codes, Array.from({ length: 70 }, (_, index) => index + 1));
});

test("save/history/undo/redo controls are public and use the persistence service", async () => {
  const harness = createHarness();
  await harness.hostWindow.aiBridge.ready({ timeoutMs: 1000 });

  await harness.hostWindow.aiBridge.save({ timeoutMs: 1000, requestId: "control-save" });
  await harness.hostWindow.aiBridge.history({ timeoutMs: 1000, requestId: "control-history" });
  await harness.hostWindow.aiBridge.undo({ timeoutMs: 1000, requestId: "control-undo" });
  await harness.hostWindow.aiBridge.redo({ timeoutMs: 1000, requestId: "control-redo" });

  assert.deepEqual(harness.servicePaths, [
    "/api/v1/editor-relay/persistence/forcesave",
    "/api/v1/editor-relay/persistence/history",
    "/api/v1/editor-relay/persistence/undo",
    "/api/v1/editor-relay/persistence/redo",
  ]);
});

test("persistence controls fail explicitly when the host service is not configured", async () => {
  const harness = createHarness({ persistenceBaseUrl: "" });
  await harness.hostWindow.aiBridge.ready({ timeoutMs: 1000 });

  await assert.rejects(
    harness.hostWindow.aiBridge.save({ timeoutMs: 1000, requestId: "missing-persistence" }),
    error => error && error.code === "PERSISTENCE_NOT_AVAILABLE",
  );
  assert.deepEqual(harness.servicePaths, []);
});

test("document-hub mode routes persistence through the authenticated host page", async () => {
  const requests = [];
  const harness = createHarness({
    persistenceBaseUrl: "/api/v1/editor-relay/persistence",
    interceptPersistence: true,
    editorSessionId: "editor-session-1234567890",
    hostFetch: async (requestPath, requestOptions) => {
      requests.push({ path: requestPath, options: requestOptions });
      return {
        ok: true,
        status: 200,
        json: async () => ({ ok: true, accepted: true, persisted: true, status: "saved" }),
      };
    },
  });
  await harness.hostWindow.aiBridge.ready({ timeoutMs: 1000 });
  await harness.hostWindow.aiBridge.word.appendParagraph(
    { text: "saved through hub" },
    { timeoutMs: 1000, requestId: "hub-persistence" },
  );

  assert.deepEqual(requests.map(request => request.path), [
    "/api/v1/editor-relay/persistence/checkpoint",
    "/api/v1/editor-relay/persistence/forcesave",
  ]);
  assert.equal(harness.servicePaths.length, 0);
  for (const request of requests) {
    assert.equal(request.options.credentials, "same-origin");
    assert.match(request.options.headers["X-Editor-Session-ID"] || "", /./);
  }
});

test("requestId replays only the identical cached request and rejects changed parameters, tools, or controls", async () => {
  const harness = createHarness();
  await harness.hostWindow.aiBridge.ready({ timeoutMs: 1000 });
  const options = { timeoutMs: 1000, requestId: "stable-agent-request-1" };

  const first = await harness.hostWindow.aiBridge.word.appendParagraph({ text: "once" }, options);
  const second = await harness.hostWindow.aiBridge.word.appendParagraph({ text: "once" }, options);

  assert.equal(first.results[0].name, "word_append_paragraph");
  assert.equal(second.results[0].name, "word_append_paragraph");
  assert.equal(harness.executedToolCalls.length, 1);
  const persistenceCallCount = harness.servicePaths.length;

  for (const invoke of [
    () => harness.hostWindow.aiBridge.word.appendParagraph({ text: "changed" }, options),
    () => harness.hostWindow.aiBridge.word.replaceText({ search: "old", replace: "new" }, options),
    () => harness.hostWindow.aiBridge.save(options),
  ]) {
    await assert.rejects(invoke(), error => {
      assert.equal(error.code, "REQUEST_ID_CONFLICT");
      assert.equal(error.requestId, options.requestId);
      assert.equal(error.details.requestId, options.requestId);
      assert.equal(error.details.retryable, false);
      assert.equal(error.details.reuseAllowed, false);
      assert.equal(error.details.requiredAction, "use_new_request_id");
      assert.equal(error.details.partialMutationPossible, false);
      return true;
    });
  }

  assert.equal(harness.executedToolCalls.length, 1);
  assert.equal(harness.servicePaths.length, persistenceCallCount);
});

test("host imports an image before sending the internal source to the plugin", async () => {
  let imports = 0;
  const harness = createHarness({
    hostFetch: async (requestPath, requestOptions) => {
      imports += 1;
      assert.equal(requestPath, "/api/v1/editor-relay/images/import");
      assert.equal(requestOptions.headers.Authorization, "Bearer editor-token");
      assert.deepEqual(JSON.parse(requestOptions.body), {
        source: { type: "url", url: "https://images.test/photo.png" },
      });
      return importedImageResponse();
    },
  });
  await harness.hostWindow.aiBridge.ready({ timeoutMs: 1000 });

  await harness.hostWindow.aiBridge.word.addImage({
    source: { type: "url", url: "https://images.test/photo.png" },
    widthMm: 50,
  }, {
    timeoutMs: 1000,
    requestId: "image-import-1",
  });

  assert.equal(imports, 1);
  assert.equal(harness.executedToolCalls.length, 1);
  const args = harness.executedToolCalls[0].arguments;
  assert.equal(args.source, undefined);
  assert.equal(args._image.url, `https://app.test/api/v1/editor-relay/images/${"a".repeat(64)}.png?token=signed-token`);
  assert.equal(args._image.widthPx, 400);
  assert.equal(args._image.heightPx, 200);
});

test("host safely imports slide, master, and layout background images before plugin execution", async () => {
  let imports = 0;
  const harness = createHarness({
    editorType: "slide",
    hostFetch: async (requestPath, requestOptions) => {
      assert.equal(requestPath, "/api/v1/editor-relay/images/import");
      imports += 1;
      assert.equal(requestOptions.headers.Authorization, "Bearer editor-token");
      return importedImageResponse({
        assetId: `${String(imports).padStart(64, "0")}.png`,
        widthPx: 1600,
        heightPx: 900,
      });
    },
  });
  await harness.hostWindow.aiBridge.ready({ timeoutMs: 1000 });

  await harness.hostWindow.aiBridge.slides.setBackground({
    slide: 1,
    mode: "image",
    source: { type: "url", url: "https://images.test/slide-background.png" },
  }, {
    timeoutMs: 1000,
    requestId: "slide-background-import",
  });
  await harness.hostWindow.aiBridge.slides.setTemplateBackground({
    scope: "master",
    masterIndex: 1,
    mode: "image",
    fillMode: "tile",
    source: { type: "dataUrl", dataUrl: "data:image/png;base64,AAAA" },
  }, {
    timeoutMs: 1000,
    requestId: "master-background-import",
  });
  await harness.hostWindow.aiBridge.slides.setTemplateBackground({
    scope: "layout",
    masterIndex: 1,
    layoutIndex: 1,
    mode: "image",
    source: { type: "url", url: "https://images.test/layout-background.png" },
  }, {
    timeoutMs: 1000,
    requestId: "layout-background-import",
  });

  assert.equal(imports, 3);
  assert.equal(harness.executedToolCalls.length, 3);
  assert.deepEqual(
    harness.executedToolCalls.map(call => call.name),
    ["slides_set_background", "slides_set_template_background", "slides_set_template_background"],
  );
  assert.ok(harness.executedToolCalls.every(call => call.arguments.source === undefined));
  assert.ok(harness.executedToolCalls.every(call => call.arguments._image.widthPx === 1600));
  assert.equal(harness.executedToolCalls[0].arguments.fillMode, undefined);
  assert.equal(harness.executedToolCalls[1].arguments.fillMode, "tile");
});

test("host reuses a leased relay image asset without posting the image again", async () => {
  const assetId = `${"b".repeat(64)}.png`;
  const assetPath = `/api/v1/editor-relay/images/${assetId}`;
  const harness = createHarness({
    editorType: "slide",
    imageBaseUrl: "/api/v1/editor-relay/images",
    hostFetch: async (requestPath, requestOptions) => {
      assert.equal(requestPath, assetPath);
      assert.equal(requestOptions.method, "GET");
      assert.equal(requestOptions.headers["X-AI-Asset-Token"], "leased-asset-token");
      return {
        ok: true,
        status: 200,
        headers: { get: name => name === "Content-Type" ? "image/png" : null },
        arrayBuffer: async () => Uint8Array.from([137, 80, 78, 71]).buffer,
      };
    },
  });
  await harness.hostWindow.aiBridge.ready({ timeoutMs: 1000 });

  await harness.hostWindow.aiBridge.slides.setBackground({
    slide: 1,
    mode: "image",
    source: {
      type: "relayAsset",
      path: assetPath,
      assetToken: "leased-asset-token",
      assetId,
      mimeType: "image/png",
      widthPx: 1600,
      heightPx: 900,
    },
  }, {
    timeoutMs: 1000,
    requestId: "slide-background-relay-asset",
  });

  assert.deepEqual(harness.hostRelayPaths, [assetPath]);
  assert.equal(harness.executedToolCalls.length, 1);
  const args = harness.executedToolCalls[0].arguments;
  assert.equal(args.source, undefined);
  assert.equal(args._image.assetId, assetId);
  assert.equal(args._image.transport, "dataUrl");
  assert.match(args._image.url, /^data:image\/png;base64,/);
});

test("failed background image import prevents checkpoint and plugin execution", async () => {
  const harness = createHarness({
    editorType: "slide",
    hostFetch: async requestPath => {
      assert.equal(requestPath, "/api/v1/editor-relay/images/import");
      return relayResponse(400, {
        ok: false,
        error: { code: "UNSUPPORTED_IMAGE_FORMAT", message: "不支持该背景图片格式" },
      });
    },
  });
  await harness.hostWindow.aiBridge.ready({ timeoutMs: 1000 });

  await assert.rejects(
    harness.hostWindow.aiBridge.slides.setBackground({
      slide: 1,
      mode: "image",
      source: { type: "url", url: "https://images.test/background.bmp" },
    }, {
      timeoutMs: 1000,
      requestId: "background-import-failure",
    }),
    error => error.code === "UNSUPPORTED_IMAGE_FORMAT",
  );

  assert.equal(harness.executedToolCalls.length, 0);
  assert.deepEqual(harness.servicePaths, []);
});

test("non-image background modes skip optional image import and invalid image mode stops before execution", async () => {
  let imports = 0;
  const harness = createHarness({
    editorType: "slide",
    hostFetch: async requestPath => {
      if (requestPath === "/api/v1/editor-relay/images/import") imports += 1;
      return importedImageResponse();
    },
  });
  await harness.hostWindow.aiBridge.ready({ timeoutMs: 1000 });

  await harness.hostWindow.aiBridge.executeBatch([
    {
      name: "slides_set_background",
      arguments: { slide: 1, fill: { type: "solid", color: "#FFFFFF" } },
    },
    {
      name: "slides_set_background",
      arguments: { slide: 1, mode: "layout" },
    },
    {
      name: "slides_set_template_background",
      arguments: { scope: "layout", masterIndex: 1, layoutIndex: 1, mode: "master" },
    },
  ], {
    timeoutMs: 1000,
    requestId: "non-image-backgrounds",
  });

  assert.equal(imports, 0);
  assert.equal(harness.executedToolCalls.length, 3);

  await assert.rejects(
    harness.hostWindow.aiBridge.slides.setBackground({ slide: 1, mode: "image" }, {
      timeoutMs: 1000,
      requestId: "missing-background-image",
    }),
    error => {
      assert.equal(error.code, "INVALID_TOOL_ARGUMENTS");
      assert.equal(error.details.completedToolCalls, 0);
      assert.equal(error.details.partialMutationPossible, false);
      assert.ok(error.details.validationErrors.some(item => item.path === "arguments.source"));
      return true;
    },
  );

  assert.equal(imports, 0);
  assert.equal(harness.executedToolCalls.length, 3);
});

test("localhost host resolves the signed asset to an internal Data URL", async () => {
  const png = Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9Z3xkAAAAASUVORK5CYII=",
    "base64",
  );
  const harness = createHarness({
    hostname: "localhost",
    httpRelay: false,
    hostFetch: async requestPath => {
      if (requestPath === "/api/v1/editor-relay/images/import") return importedImageResponse({
        widthPx: 1,
        heightPx: 1,
      });
      if (requestPath.startsWith("/api/v1/editor-relay/images/")) {
        return {
          ok: true,
          status: 200,
          headers: { get: name => name.toLowerCase() === "content-type" ? "image/png" : null },
          arrayBuffer: async () => png.buffer.slice(png.byteOffset, png.byteOffset + png.byteLength),
        };
      }
      if (requestPath.endsWith("/register")) {
        return relayResponse(200, { ok: true, relayKey: "relay-key", resumeToken: "resume-token" });
      }
      if (requestPath.endsWith("/poll")) return relayResponse(200, { ok: true, command: null });
      throw new Error(`unexpected request: ${requestPath}`);
    },
  });
  await harness.hostWindow.aiBridge.ready({ timeoutMs: 1000 });

  await harness.hostWindow.aiBridge.word.addImage({
    source: { type: "dataUrl", dataUrl: `data:image/png;base64,${png.toString("base64")}` },
  }, {
    timeoutMs: 1000,
    requestId: "loopback-image-data",
  });

  const image = harness.executedToolCalls[0].arguments._image;
  assert.equal(image.transport, "dataUrl");
  assert.match(image.url, /^data:image\/png;base64,/);
  assert.equal(image.url.includes("signed-token"), false);
});

test("image batch imports atomically and retries the same requestId without re-importing", async () => {
  let imports = 0;
  const harness = createHarness({
    editorType: "slide",
    hostFetch: async () => {
      imports += 1;
      return importedImageResponse({
        assetId: `${String(imports).padStart(64, "0")}.png`,
      });
    },
  });
  await harness.hostWindow.aiBridge.ready({ timeoutMs: 1000 });
  const command = {
    toolCalls: [
      {
        name: "slides_add_image",
        arguments: { slide: 1, source: { type: "url", url: "https://images.test/one.png" } },
      },
      {
        name: "slides_add_image",
        arguments: { slide: 1, source: { type: "dataUrl", dataUrl: "data:image/png;base64,AAAA" } },
      },
    ],
  };
  const options = { timeoutMs: 1000, requestId: "image-batch-retry" };

  await harness.hostWindow.aiBridge.executeBatch(command.toolCalls, options);
  await harness.hostWindow.aiBridge.executeBatch(command.toolCalls, options);

  assert.equal(imports, 2);
  assert.equal(harness.executedToolCalls.length, 2);
  assert.ok(harness.executedToolCalls.every(call => call.arguments.source === undefined));
});

test("one failed image import prevents the entire batch from reaching the plugin", async () => {
  let imports = 0;
  const harness = createHarness({
    hostFetch: async () => {
      imports += 1;
      if (imports === 2) {
        return relayResponse(400, {
          ok: false,
          error: { code: "UNSUPPORTED_IMAGE_FORMAT", message: "不支持该图片格式" },
        });
      }
      return importedImageResponse();
    },
  });
  await harness.hostWindow.aiBridge.ready({ timeoutMs: 1000 });

  await assert.rejects(
    harness.hostWindow.aiBridge.executeBatch([
      {
        name: "word_add_image",
        arguments: { source: { type: "url", url: "https://images.test/one.png" } },
      },
      {
        name: "word_add_image",
        arguments: { source: { type: "url", url: "https://images.test/two.svg" } },
      },
    ], {
      timeoutMs: 1000,
      requestId: "image-batch-failure",
    }),
    error => error.code === "UNSUPPORTED_IMAGE_FORMAT",
  );

  assert.equal(imports, 2);
  assert.equal(harness.executedToolCalls.length, 0);
  assert.deepEqual(harness.servicePaths, []);
});

test("editor save failures include a safe diagnostic phase", async () => {
  const harness = createHarness({ saveError: "编辑器保存失败" });
  await harness.hostWindow.aiBridge.ready({ timeoutMs: 1000 });

  await assert.rejects(
    harness.hostWindow.aiBridge.word.setDocumentText(
      { text: "敏感正文" },
      { timeoutMs: 1000, requestId: "save-phase-test" },
    ),
    error => {
      assert.equal(error.code, "EXECUTION_FAILED");
      assert.equal(error.details.phase, "editor-save");
      assert.equal(error.details.partialMutationPossible, true);
      assert.doesNotMatch(error.message, /敏感正文/);
      return true;
    },
  );
});

test("force-save failures report persistence phase and possible document mutation", async () => {
  const harness = createHarness({
    serviceResponse(path) {
      if (path === "/api/v1/editor-relay/persistence/forcesave") {
        return {
          ok: false,
          error: {
            code: "PERSISTENCE_FAILED",
            message: "持久化服务暂时失败",
          },
        };
      }
      return { persisted: true, path };
    },
  });
  await harness.hostWindow.aiBridge.ready({ timeoutMs: 1000 });

  await assert.rejects(
    harness.hostWindow.aiBridge.word.setDocumentText(
      { text: "敏感正文" },
      { timeoutMs: 1000, requestId: "persistence-failure" },
    ),
    error => {
      assert.equal(error.code, "PERSISTENCE_FAILED");
      assert.equal(error.message, "持久化服务暂时失败");
      assert.equal(error.details.phase, "persistence");
      assert.equal(error.details.partialMutationPossible, true);
      assert.equal(error.details.path, "forcesave");
      assert.doesNotMatch(JSON.stringify(error.details), /敏感正文/);
      return true;
    },
  );
});

test("force-save failed status is normalized as PERSISTENCE_FAILED", async () => {
  const harness = createHarness({
    serviceResponse(path) {
      if (path === "/api/v1/editor-relay/persistence/forcesave") {
        return {
          accepted: false,
          noChanges: false,
          persisted: false,
          status: "failed",
          commandError: 4,
          beforeMtime: 100,
          afterMtime: 100,
        };
      }
      return { persisted: true, path };
    },
  });
  await harness.hostWindow.aiBridge.ready({ timeoutMs: 1000 });

  await assert.rejects(
    harness.hostWindow.aiBridge.word.replaceText(
      { search: "old", replace: "new" },
      { timeoutMs: 1000, requestId: "persistence-failed-status" },
    ),
    error => {
      assert.equal(error.code, "PERSISTENCE_FAILED");
      assert.equal(error.details.persistence.status, "failed");
      assert.equal(error.details.persistence.forceSave.commandError, 4);
      return true;
    },
  );
});

test("force-save persists without reloading the live editor", async () => {
  const harness = createHarness({
    serviceResponse(path) {
      if (path === "/api/v1/editor-relay/persistence/forcesave") {
        return {
          accepted: true,
          persisted: true,
          status: "saved",
          commandError: 0,
          promoted: true,
          beforeMtime: 100,
          afterMtime: 101,
        };
      }
      return { persisted: true, path };
    },
  });
  await harness.hostWindow.aiBridge.ready({ timeoutMs: 1000 });

  const result = await harness.hostWindow.aiBridge.word.replaceText(
    { search: "old", replace: "new" },
    { timeoutMs: 1000, requestId: "persist-and-rebind-1" },
  );

  assert.equal(result.persisted, true);
  assert.equal(result.forceSave.promoted, true);
  assert.equal(result.persistence.status, "saved");
  assert.equal(result.persistence.editorSaved, true);
  assert.equal(result.persistence.forceSave.commandError, 0);
  const forceSaveRequest = harness.serviceRequests.find(item => item.path === "/api/v1/editor-relay/persistence/forcesave");
  assert.ok(forceSaveRequest);
  const forceSavePayload = JSON.parse(forceSaveRequest.requestOptions.body).payload;
  assert.equal(Number.isFinite(forceSavePayload.editorSaveConfirmedAt), true);
  assert.ok(forceSavePayload.editorSaveConfirmedAt > 0);
  assert.equal(harness.reloadCount, 0);
  await new Promise(resolve => setTimeout(resolve, 80));
  assert.equal(harness.reloadCount, 0);
});

test("explicit save persists without reloading the live editor", async () => {
  const harness = createHarness({
    serviceResponse(path) {
      if (path === "/api/v1/editor-relay/persistence/forcesave") {
        return {
          accepted: true,
          persisted: true,
          status: "saved",
          commandError: 0,
          promoted: true,
          beforeMtime: 100,
          afterMtime: 101,
        };
      }
      return { persisted: true, path };
    },
  });
  await harness.hostWindow.aiBridge.ready({ timeoutMs: 1000 });

  const result = await harness.hostWindow.aiBridge.save({
    timeoutMs: 1000,
    requestId: "save-without-reload-1",
  });

  assert.equal(result.persisted, true);
  assert.equal(result.persistence.status, "saved");
  assert.equal(result.persistence.editorSaved, true);
  assert.equal(harness.reloadCount, 0);
  await new Promise(resolve => setTimeout(resolve, 80));
  assert.equal(harness.reloadCount, 0);
});

test("explicit save treats CommandService no-changes as persisted", async () => {
  const harness = createHarness({
    serviceResponse(path) {
      if (path === "/api/v1/editor-relay/persistence/forcesave") {
        return {
          accepted: false,
          noChanges: true,
          persisted: true,
          status: "no_changes",
          commandError: 4,
          beforeMtime: 100,
          afterMtime: 100,
        };
      }
      return { persisted: true, path };
    },
  });
  await harness.hostWindow.aiBridge.ready({ timeoutMs: 1000 });

  const result = await harness.hostWindow.aiBridge.save({
    timeoutMs: 1000,
    requestId: "save-no-changes-1",
  });

  assert.equal(result.persisted, true);
  assert.equal(result.persistence.status, "no_changes");
  assert.equal(result.persistence.forceSave.noChanges, true);
  assert.equal(result.persistence.forceSave.commandError, 4);
});

test("cross-origin client SDK connects through an explicit relay allow-list", async () => {
  const harness = createHarness({ clientOrigins: ["https://copilot.test"] });
  await harness.hostWindow.aiBridge.ready({ timeoutMs: 1000 });

  const clientWindow = eventTarget({
    location: { origin: "https://copilot.test", href: "https://copilot.test/agent" },
    crypto: { randomUUID: () => "10000000-0000-4000-8000-000000000001" },
    setTimeout,
    clearTimeout,
  });
  const hostHandle = {
    postMessage(message, targetOrigin) {
      assert.equal(targetOrigin, "https://app.test");
      queueMicrotask(() => harness.hostWindow.dispatch("message", {
        origin: "https://copilot.test",
        source: clientHandle,
        data: message,
      }));
    },
  };
  const clientHandle = {
    postMessage(message, targetOrigin) {
      assert.equal(targetOrigin, "https://copilot.test");
      queueMicrotask(() => clientWindow.dispatch("message", {
        origin: "https://app.test",
        source: hostHandle,
        data: message,
      }));
    },
  };

  vm.runInNewContext(clientSource, { window: clientWindow, URL, console });
  const office = new clientWindow.AiBridgeClient({
    targetWindow: hostHandle,
    targetOrigin: "https://app.test",
    timeoutMs: 1000,
  });

  const state = await office.connect({ timeoutMs: 1000 });
  assert.equal(state.context.documentKey, "doc-key-v1");
  assert.equal(state.context.callbackUrl, undefined);
  assert.equal(office.editorType, "word");
  assert.deepEqual(
    Object.keys(office.word).sort(),
    Object.keys(harness.hostWindow.aiBridge.word).sort(),
  );
  assert.deepEqual(
    Object.keys(office.slides).sort(),
    Object.keys(harness.hostWindow.aiBridge.slides).sort(),
  );
  assert.deepEqual(
    Object.keys(office.sheets).sort(),
    Object.keys(harness.hostWindow.aiBridge.sheets).sort(),
  );
  assert.equal(
    Object.keys(office.word).length +
    Object.keys(office.slides).length +
    Object.keys(office.sheets).length,
    Object.keys(harness.hostWindow.aiBridge.word).length +
    Object.keys(harness.hostWindow.aiBridge.slides).length +
    Object.keys(harness.hostWindow.aiBridge.sheets).length,
  );

  const result = await office.word.replaceText(
    { search: "old", replace: "new" },
    { timeoutMs: 1000, requestId: "external:replace:1" },
  );
  assert.equal(result.results[0].name, "word_replace_text");
  assert.equal(harness.executedToolCalls.at(-1).name, "word_replace_text");
  const retry = await office.word.replaceText(
    { search: "old", replace: "new" },
    { timeoutMs: 1000, requestId: "external:replace:1" },
  );
  assert.equal(retry.results[0].name, "word_replace_text");
  assert.equal(harness.executedToolCalls.filter(call => call.name === "word_replace_text").length, 1);
  const scroll = await office.word.scroll(
    { direction: "down", pages: 1 },
    { timeoutMs: 1000, requestId: "external:scroll:1" },
  );
  assert.equal(scroll.results[0].name, "word_scroll");
  assert.equal(scroll.changed, 0);
  assert.equal(scroll.persisted, false);
  office.destroy();
});

test("cross-origin client SDK times out when its origin is not allow-listed", async () => {
  const harness = createHarness({ clientOrigins: [] });
  await harness.hostWindow.aiBridge.ready({ timeoutMs: 1000 });
  const clientWindow = eventTarget({
    location: { origin: "https://blocked.test", href: "https://blocked.test/agent" },
    crypto: { randomUUID: () => "20000000-0000-4000-8000-000000000001" },
    setTimeout,
    clearTimeout,
  });
  const hostHandle = {
    postMessage(message) {
      queueMicrotask(() => harness.hostWindow.dispatch("message", {
        origin: "https://blocked.test",
        source: clientHandle,
        data: message,
      }));
    },
  };
  const clientHandle = { postMessage() { throw new Error("blocked client must receive no reply"); } };
  vm.runInNewContext(clientSource, { window: clientWindow, URL, console });
  const office = new clientWindow.AiBridgeClient({
    targetWindow: hostHandle,
    targetOrigin: "https://app.test",
    timeoutMs: 100,
  });

  await assert.rejects(office.connect({ timeoutMs: 100 }), error => {
    assert.equal(error.code, "CONNECTION_TIMEOUT");
    return true;
  });
  office.destroy();
});

test("word bridge replaces text through the compatible plugin method", async () => {
  const harness = createWordBridgeHarness({ text: "主任委员：姜涛\n记录人：姜涛" });
  assert.equal(typeof harness.document.SearchAndReplace, "undefined");

  const result = await harness.bridge.execute([{
    name: "word_replace_text",
    arguments: { search: "姜涛", replace: "张三", matchCase: false },
  }]);

  assert.equal(harness.text, "主任委员：张三\n记录人：张三");
  assert.equal(result.changed, 2);
  assert.equal(result.needsSave, true);
  assert.equal(result.results[0].replacedCount, 2);
  assert.deepEqual(harness.executeMethodCalls, [{
    name: "SearchAndReplace",
    args: [{ searchString: "姜涛", replaceString: "张三", matchCase: false }],
  }]);
});

test("word bridge falls back when native replacement capability probing returns no result", async () => {
  const harness = createWordBridgeHarness({
    text: "A",
    capabilityProbeRawResult: "",
  });

  const result = await harness.bridge.execute([{
    name: "word_replace_text",
    arguments: { search: "A", replace: "B", matchCase: true },
  }]);

  assert.equal(harness.text, "B");
  assert.equal(result.changed, 1);
  assert.deepEqual(harness.executeMethodCalls, [{
    name: "SearchAndReplace",
    args: [{ searchString: "A", replaceString: "B", matchCase: true }],
  }]);
});

test("word bridge reports capability probe failure when no safe fallback exists", async () => {
  const harness = createWordBridgeHarness({
    text: "A",
    capabilityProbeRawResult: "",
    executeMethodSupported: false,
  });

  await assert.rejects(
    harness.bridge.execute([{
      name: "word_replace_text",
      arguments: { search: "A", replace: "不得泄露", matchCase: true },
    }]),
    error => {
      assert.equal(error.code, "EDITOR_CAPABILITY_PROBE_FAILED");
      assert.equal(error.details.editorType, "word");
      assert.equal(error.details.operation, "probe_search_and_replace");
      assert.equal(error.details.reason, "probe_failed");
      assert.equal(error.details.partialMutationPossible, false);
      assert.equal(error.details.requiredAction, "report_bridge_failure");
      assert.doesNotMatch(JSON.stringify(error.details), /不得泄露/);
      return true;
    },
  );
});

test("word bridge skips SearchAndReplace when the source text is absent", async () => {
  const harness = createWordBridgeHarness({ text: "主任委员：李四" });

  const result = await harness.bridge.execute([{
    name: "word_replace_text",
    arguments: { search: "姜涛", replace: "张三", matchCase: false },
  }]);

  assert.equal(result.changed, 0);
  assert.equal(result.needsSave, false);
  assert.equal(result.results[0].replaced, false);
  assert.equal(harness.executeMethodCalls.length, 0);
});

test("word bridge keeps native replacement chains in one mutation callCommand", async () => {
  const harness = createWordBridgeHarness({
    text: "A",
    nativeSearchAndReplace: true,
  });
  const replacements = [
    { name: "word_replace_text", arguments: { search: "A", replace: "B", matchCase: true } },
    { name: "word_replace_text", arguments: { search: "B", replace: "C", matchCase: true } },
  ];
  for (let index = 0; index < 18; index += 1) {
    replacements.push({
      name: "word_replace_text",
      arguments: { search: `missing-${index}`, replace: `unused-${index}`, matchCase: true },
    });
  }

  const result = await harness.bridge.execute(replacements);

  assert.equal(harness.text, "C");
  assert.equal(result.changed, 2);
  assert.equal(harness.historyPoints, 1);
  assert.equal(harness.executeMethodCalls.length, 0);
  assert.equal(harness.callCommandCalls.filter(call => call.recalculate).length, 1);
  assert.equal(harness.callCommandCalls.filter(call => !call.recalculate).length, 1);
  assert.deepEqual(harness.nativeSearchAndReplaceCalls, [
    { searchString: "A", replaceString: "B", matchCase: true },
    { searchString: "B", replaceString: "C", matchCase: true },
  ]);
  assert.deepEqual(Array.from(result.results, entry => entry.search), [
    "A", "B",
    ...Array.from({ length: 18 }, (_, index) => `missing-${index}`),
  ]);
});

test("word bridge does not fall back after a native SearchAndReplace failure", async () => {
  const harness = createWordBridgeHarness({
    text: "A",
    nativeSearchAndReplace: true,
    nativeSearchAndReplaceAccepted: false,
  });

  await assert.rejects(
    harness.bridge.execute([{
      name: "word_replace_text",
      arguments: { search: "A", replace: "B", matchCase: true },
    }]),
    /ONLYOFFICE 拒绝执行 SearchAndReplace/,
  );

  assert.equal(harness.text, "A");
  assert.equal(harness.nativeSearchAndReplaceCalls.length, 1);
  assert.equal(harness.executeMethodCalls.length, 0);
});

test("word bridge formats only matching text ranges", async () => {
  const harness = createWordBridgeHarness({ text: "主任委员：匿名\n副主任委员：匿名、匿名、匿名" });

  const result = await harness.bridge.execute([{
    name: "word_format_matches",
    arguments: { search: "匿名", bold: true, italic: true, matchCase: false },
  }]);

  assert.equal(result.changed, 4);
  assert.equal(result.needsSave, true);
  assert.equal(result.results[0].matches, 4);
  assert.equal(result.results[0].formattedMatches, 4);
  assert.equal(harness.formattedRanges.length, 4);
  assert.ok(harness.formattedRanges.every(range => range.bold === true && range.italic === true));
});

test("word bridge normalizes safe point aliases, lets canonical fields win, and rejects ambiguous twips", async () => {
  const harness = createWordBridgeHarness({ text: "已有正文" });

  await harness.bridge.execute([{
    name: "word_append_paragraph",
    arguments: {
      text: "悬挂缩进",
      spacingBeforePt: 6,
      spacingAfterPt: 4,
      leftIndentPt: 36,
      firstLineIndentPt: -18,
      lineSpacing: 1.5,
      lineRule: "auto",
    },
  }]);

  const paragraph = harness.paragraphObjects.at(-1);
  assert.equal(paragraph.spacingBefore, 120);
  assert.equal(paragraph.spacingAfter, 80);
  assert.equal(paragraph.leftIndent, 720);
  assert.equal(paragraph.firstLineIndent, -360);
  assert.deepEqual(paragraph.spacingLine, { value: 360, rule: "auto" });

  await assert.rejects(
    harness.bridge.execute([{
      name: "word_append_paragraph",
      arguments: { text: "错误缩进", leftIndent: 720 },
    }]),
    error => error.code === "INVALID_TOOL_ARGUMENTS"
      && error.details.validationErrors[0].path === "arguments.leftIndent",
  );
  await assert.rejects(
    harness.bridge.execute([{
      name: "word_append_paragraph",
      arguments: { text: "错误悬挂缩进", firstLineIndent: -360 },
    }]),
    error => error.code === "INVALID_TOOL_ARGUMENTS"
      && error.details.validationErrors[0].path === "arguments.firstLineIndent",
  );
  const canonicalWins = await harness.bridge.execute([{
    name: "word_append_paragraph",
    arguments: { text: "字段兼容", leftIndent: 720, leftIndentPt: 36 },
  }]);
  assert.equal(harness.paragraphObjects.at(-1).leftIndent, 720);
  assert.deepEqual(JSON.parse(JSON.stringify(canonicalWins.argumentNormalizations)), [{
    toolCallIndex: 0,
    path: "arguments.leftIndent",
    canonicalPath: "arguments.leftIndentPt",
    kind: "canonicalWins",
  }]);
  const legacyLineRule = await harness.bridge.execute([{
    name: "word_append_paragraph",
    arguments: { text: "兼容行距", lineSpacing: 1.5, lineRule: "multiple" },
  }]);
  assert.deepEqual(harness.paragraphObjects.at(-1).spacingLine, {
    value: 360,
    rule: "auto",
  });
  assert.equal(legacyLineRule.argumentNormalizations[0].kind, "enumAlias");
});

test("word bridge creates ONLYOFFICE run styles for public character styles and applies them", async () => {
  const harness = createWordBridgeHarness({ text: "术语 术语" });

  const result = await harness.bridge.execute([
    {
      name: "word_manage_style",
      arguments: { action: "create", name: "术语强调", type: "character", bold: true },
    },
    {
      name: "word_format_matches",
      arguments: { search: "术语", characterStyleName: "术语强调" },
    },
  ]);

  assert.equal(result.changed, 3);
  assert.equal(harness.createdStyles.length, 1);
  assert.equal(harness.createdStyles[0].type, "run");
  assert.equal(result.results[0].type, "character");
  assert.equal(harness.appliedStyles.length, 2);
  assert.ok(harness.appliedStyles.every(entry => entry.style === harness.createdStyles[0]));
});

test("word document properties initialize an empty core and round-trip through advanced inspection", async () => {
  const harness = createWordBridgeHarness({
    text: "空白文档",
    core: null,
    customProperties: null,
  });

  const updated = await harness.bridge.execute([{
    name: "word_set_document_properties",
    arguments: {
      title: "雨停之前",
      subject: "短篇小说",
      creator: "佚名",
      keywords: "故乡,车站,雨",
      custom: [{ name: "ReviewState", value: "Final", valueType: "string" }],
    },
  }]);
  const inspected = await harness.bridge.execute([{
    name: "word_inspect_advanced",
    arguments: {
      includeProperties: true,
      customPropertyNames: ["ReviewState"],
    },
  }]);

  assert.equal(updated.changed, 5);
  assert.deepEqual(JSON.parse(JSON.stringify(updated.results[0].updated)), [
    "title", "subject", "creator", "keywords", "custom:ReviewState",
  ]);
  assert.ok(harness.document.Document.Core);
  assert.ok(harness.document.Document.CustomProperties);
  assert.deepEqual(JSON.parse(JSON.stringify(inspected.results[0].properties.core)), {
    title: "雨停之前",
    subject: "短篇小说",
    creator: "佚名",
    description: null,
    keywords: "故乡,车站,雨",
    category: null,
    language: null,
    identifier: null,
    lastModifiedBy: null,
    revision: null,
    created: null,
    modified: null,
  });
  assert.deepEqual(JSON.parse(JSON.stringify(inspected.results[0].properties.custom)), {
    ReviewState: "Final",
  });
});

test("word document properties validate every setter before mutating the core", async () => {
  const harness = createWordBridgeHarness({
    text: "属性校验",
    coreValues: { title: "原始标题", subject: "原始主题" },
    missingCoreSetter: "SetSubject",
  });

  await assert.rejects(
    harness.bridge.execute([{
      name: "word_set_document_properties",
      arguments: { title: "新标题", subject: "新主题" },
    }]),
    error => error.code === "WORD_API_UNSUPPORTED"
      && error.details.partialMutationPossible === false,
  );

  assert.equal(harness.document.Document.Core.title, "原始标题");
  assert.equal(harness.document.Document.Core.subject, "原始主题");
});

test("word document properties return a stable capability error when an empty core cannot initialize", async () => {
  const harness = createWordBridgeHarness({
    text: "属性能力",
    core: null,
    coreFactorySupported: false,
  });

  await assert.rejects(
    harness.bridge.execute([{
      name: "word_set_document_properties",
      arguments: { title: "无法写入" },
    }]),
    error => error.code === "WORD_API_UNSUPPORTED"
      && /无法初始化文档核心属性/.test(error.message)
      && error.details.partialMutationPossible === false,
  );
});

test("word advanced inspection includes styles omitted by ONLYOFFICE 9.4 GetAllStyles", async () => {
  const harness = createWordBridgeHarness({
    text: "自定义样式",
    omitCreatedStylesFromGetAllStyles: true,
  });

  await harness.bridge.execute([
    {
      name: "word_manage_style",
      arguments: {
        action: "create",
        name: "正文_仿宋",
        type: "paragraph",
        fontFamily: "仿宋",
        fontSize: 11,
      },
    },
    {
      name: "word_manage_style",
      arguments: {
        action: "create",
        name: "BodyFangsong",
        type: "paragraph",
        fontFamily: "仿宋",
        fontSize: 11,
      },
    },
    {
      name: "word_manage_style",
      arguments: {
        action: "create",
        name: "术语强调",
        type: "character",
        bold: true,
      },
    },
  ]);
  const inspected = await harness.bridge.execute([{
    name: "word_inspect_advanced",
    arguments: { includeStyles: true },
  }]);

  assert.equal(inspected.changed, 0);
  assert.equal(inspected.needsSave, false);
  assert.deepEqual(JSON.parse(JSON.stringify(inspected.results[0].styles)), [
    {
      index: 1,
      name: "正文_仿宋",
      type: "paragraph",
      basedOn: null,
      usageCount: 0,
      tableParagraphUsageCount: 0,
    },
    {
      index: 2,
      name: "BodyFangsong",
      type: "paragraph",
      basedOn: null,
      usageCount: 0,
      tableParagraphUsageCount: 0,
    },
    {
      index: 3,
      name: "术语强调",
      type: "run",
      basedOn: null,
      usageCount: 0,
      tableParagraphUsageCount: 0,
    },
  ]);
});

test("word advanced inspection de-duplicates styles returned by public and internal collections", async () => {
  const harness = createWordBridgeHarness({ text: "样式去重" });

  await harness.bridge.execute([{
    name: "word_manage_style",
    arguments: {
      action: "create",
      name: "SharedStyle",
      type: "paragraph",
    },
  }]);
  const inspected = await harness.bridge.execute([{
    name: "word_inspect_advanced",
    arguments: { includeStyles: true },
  }]);

  assert.deepEqual(
    inspected.results[0].styles.filter(style => style.name === "SharedStyle"),
    [{
      index: 1,
      name: "SharedStyle",
      type: "paragraph",
      basedOn: null,
      usageCount: 0,
      tableParagraphUsageCount: 0,
    }],
  );
});

test("word bridge applies advanced borders and margins to the targeted table row cells", async () => {
  const calls = [];
  function createCell(index) {
    return {
      index,
      SetCellMarginTop(value) { calls.push(["marginTop", index, value]); return true; },
      SetCellMarginLeft(value) { calls.push(["marginLeft", index, value]); return true; },
      SetCellBorderTop(...args) { calls.push(["borderTop", index, ...args]); return true; },
      SetWidth(unit, value) { calls.push(["width", index, unit, value]); return true; },
    };
  }
  const firstRowCells = [createCell(0), createCell(1)];
  const secondRowCells = [createCell(2), createCell(3)];
  function createRow(index, cells) {
    return {
      GetCellsCount() { return cells.length; },
      GetCell(cellIndex) { return cells[cellIndex]; },
      SetHeight(rule, value) { calls.push(["height", index, rule, value]); return true; },
      SetTableHeader(value) { calls.push(["header", index, value]); return true; },
    };
  }
  const rows = [createRow(0, firstRowCells), createRow(1, secondRowCells)];
  const table = {
    GetRowsCount() { return rows.length; },
    GetRow(index) { return rows[index]; },
    GetAllCells() { return firstRowCells.concat(secondRowCells); },
  };
  const harness = createWordBridgeHarness({ text: "表格", tables: [table] });

  const result = await harness.bridge.execute([{
    name: "word_format_table_advanced",
    arguments: {
      tableIndex: 1,
      row: 1,
      rowHeightMm: 10,
      rowHeightRule: "atLeast",
      repeatHeader: true,
      columnWidthMm: 20,
      cellMarginsMm: { top: 2, left: 3 },
      borders: { top: { style: "single", widthPt: 1, color: "#112233" } },
    },
  }]);

  assert.equal(result.changed, 1);
  assert.equal(result.results[0].cells, 2);
  assert.deepEqual(calls.filter(call => call[0] === "width").map(call => call[1]), [0, 1]);
  assert.deepEqual(calls.filter(call => call[0] === "marginTop").map(call => call[1]), [0, 1]);
  assert.deepEqual(calls.filter(call => call[0] === "borderTop").map(call => call[1]), [0, 1]);
  assert.equal(calls.some(call => call[1] === 2 || call[1] === 3), false);
});

test("word bridge failures identify the failed batch tool without echoing arguments", async () => {
  const harness = createWordBridgeHarness({
    paragraphs: ["第一段"],
    currentParagraphIndex: 1,
  });

  await assert.rejects(
    harness.bridge.execute([
      {
        name: "word_replace_text",
        arguments: { search: "不存在的文本", replace: "不会发生" },
      },
      {
        name: "word_insert_page_break",
        arguments: {
          current: true,
          position: "end",
          internalNote: "敏感正文",
        },
      },
    ]),
    error => {
      assert.equal(error.code, "INVALID_ARGUMENTS");
      assert.equal(error.message, "word_insert_page_break.position 必须是 before 或 after");
      assert.equal(error.details.phase, "word-command");
      assert.equal(error.details.tool, "word_insert_page_break");
      assert.equal(error.details.toolCallIndex, 1);
      assert.equal(error.details.completedToolCalls, 1);
      assert.equal(error.details.partialMutationPossible, false);
      assert.doesNotMatch(JSON.stringify(error.details), /敏感正文|internalNote/);
      return true;
    },
  );
});

test("word bridge classifies the observed page-break and table argument failures", async () => {
  const pageBreakHarness = createWordBridgeHarness({ paragraphs: ["第一段"] });
  await assert.rejects(
    pageBreakHarness.bridge.execute([{
      name: "word_insert_page_break",
      arguments: {},
    }]),
    error => {
      assert.equal(error.code, "INVALID_TOOL_ARGUMENTS");
      assert.equal(error.message, "Word 工具参数校验失败");
      assert.equal(error.details.completedToolCalls, 0);
      assert.equal(error.details.partialMutationPossible, false);
      assert.equal(
        JSON.stringify(error.details.validationErrors.map(item => item.path)),
        JSON.stringify(["arguments.target"]),
      );
      return true;
    },
  );
  assert.equal(pageBreakHarness.historyPoints, 0);

  const cells = [{}, {}];
  const row = {
    GetCellsCount() { return cells.length; },
    GetCell(index) { return cells[index] || null; },
  };
  const table = {
    GetRowsCount() { return 1; },
    GetRow(index) { return index === 0 ? row : null; },
    GetAllCells() { return cells; },
  };
  const tableHarness = createWordBridgeHarness({
    paragraphs: ["表格前"],
    tables: [table],
  });

  await assert.rejects(
    tableHarness.bridge.execute([{
      name: "word_format_table_advanced",
      arguments: { tableIndex: 1, repeatHeader: true },
    }]),
    error => {
      assert.equal(error.code, "INVALID_TOOL_ARGUMENTS");
      assert.equal(error.details.completedToolCalls, 0);
      assert.equal(error.details.partialMutationPossible, false);
      assert.equal(
        JSON.stringify(error.details.validationErrors.map(item => item.path)),
        JSON.stringify(["arguments.row"]),
      );
      return true;
    },
  );

  await assert.rejects(
    tableHarness.bridge.execute([{
      name: "word_edit_table",
      arguments: {
        tableIndex: 1,
        action: "mergeCells",
        startRow: 1,
        startColumn: 1,
        endRow: 1,
        endColumn: 2,
      },
    }]),
    error => {
      assert.equal(error.code, "INVALID_TOOL_ARGUMENTS");
      assert.equal(error.details.completedToolCalls, 0);
      assert.equal(error.details.partialMutationPossible, false);
      assert.equal(
        JSON.stringify(error.details.validationErrors.map(item => item.path)),
        JSON.stringify([
          "arguments.rowStart",
          "arguments.rowEnd",
          "arguments.columnStart",
          "arguments.columnEnd",
        ]),
      );
      return true;
    },
  );
  assert.equal(tableHarness.historyPoints, 0);
});

test("word bridge rejects an empty first-page header before entering ONLYOFFICE", async () => {
  let getHeaderCalls = 0;
  const harness = createWordBridgeHarness({
    sections: [{
      GetHeader() {
        getHeaderCalls += 1;
        return null;
      },
    }],
  });

  await assert.rejects(
    harness.bridge.execute([{
      name: "word_set_header_footer",
      arguments: { kind: "header", type: "first", text: "" },
    }]),
    error => {
      assert.equal(error.code, "INVALID_TOOL_ARGUMENTS");
      assert.equal(error.details.completedToolCalls, 0);
      assert.equal(error.details.partialMutationPossible, false);
      assert.equal(error.details.validationErrors.length, 1);
      assert.equal(error.details.validationErrors[0].path, "arguments.text");
      assert.equal(error.details.validationErrors[0].keyword, "semantic");
      assert.match(error.details.validationErrors[0].message, /titlePage=true/);
      assert.match(error.details.validationErrors[0].message, /action=remove/);
      return true;
    },
  );

  assert.equal(harness.historyPoints, 0);
  assert.equal(harness.callCommandCalls.length, 0);
  assert.equal(getHeaderCalls, 0);
});

test("word bridge reports null header and footer content without dereferencing it", async () => {
  for (const kind of ["header", "footer"]) {
    let contentCalls = 0;
    const section = {
      GetHeader() {
        contentCalls += 1;
        return null;
      },
      GetFooter() {
        contentCalls += 1;
        return null;
      },
    };
    const harness = createWordBridgeHarness({ sections: [section] });

    await assert.rejects(
      harness.bridge.execute([{
        name: "word_set_header_footer",
        arguments: { kind, type: "first", text: "首页内容" },
      }]),
      error => {
        assert.equal(error.code, "WORD_API_UNSUPPORTED");
        assert.doesNotMatch(error.message, /Cannot read properties of null/);
        assert.match(error.message, /titlePage=true/);
        assert.equal(error.details.partialMutationPossible, true);
        assert.deepEqual(error.details.headerFooter, {
          kind,
          type: "first",
          sectionIndex: 1,
          remediation: "Use word_set_page_layout with titlePage=true for an empty first page, or create non-empty first-page content after enabling titlePage.",
        });
        return true;
      },
    );

    assert.equal(contentCalls, 1);
  }
});

test("word bridge reports exact footnote locator misses without possible mutation", async () => {
  const harness = createWordBridgeHarness({
    paragraphs: ["恢复时间目标(RTO)应不超过四小时"],
  });
  const originalParagraphs = [...harness.paragraphs];

  await assert.rejects(
    harness.bridge.execute([{
      name: "word_manage_long_document",
      arguments: {
        action: "addFootnote",
        search: "恢复时间目标(RTO)",
        matchMode: "exact",
        noteText: "RTO 定义",
      },
    }]),
    error => {
      assert.equal(error.code, "INVALID_TARGET");
      assert.match(error.message, /完整段落文本/);
      assert.match(error.message, /contains/);
      assert.equal(error.details.tool, "word_manage_long_document");
      assert.equal(error.details.completedToolCalls, 0);
      assert.equal(error.details.partialMutationPossible, false);
      assert.equal(error.details.locator, "search");
      assert.equal(error.details.matchMode, "exact");
      assert.equal(error.details.matchCount, 0);
      assert.equal(Object.prototype.hasOwnProperty.call(error.details, "search"), false);
      assert.doesNotMatch(JSON.stringify(error.details), /恢复时间目标/);
      return true;
    },
  );

  assert.deepEqual(harness.paragraphs, originalParagraphs);
});

test("word bridge classifies shared paragraph locator failures before mutation", async () => {
  const cases = [
    {
      name: "endnote search miss",
      tool: {
        name: "word_manage_long_document",
        arguments: { action: "addEndnote", search: "不存在", noteText: "说明" },
      },
      code: "INVALID_TARGET",
      locator: "search",
    },
    {
      name: "searched occurrence out of range",
      tool: {
        name: "word_manage_long_document",
        arguments: { action: "addFootnote", search: "正文", occurrence: 2, noteText: "说明" },
      },
      code: "INVALID_TARGET",
      locator: "search",
      occurrence: 2,
      matchCount: 1,
    },
    {
      name: "page break index out of range",
      tool: {
        name: "word_insert_page_break",
        arguments: { paragraphIndex: 9, position: "after" },
      },
      code: "INVALID_TARGET",
      locator: "paragraphIndex",
    },
    {
      name: "numbering paragraph id missing",
      tool: {
        name: "word_set_numbering",
        arguments: {
          assignments: [{ paragraphId: "missing-paragraph", level: 0 }],
        },
      },
      code: "INVALID_TARGET",
      locator: "paragraphId",
    },
    {
      name: "invalid paragraph index",
      tool: {
        name: "word_insert_page_break",
        arguments: { paragraphIndex: 0, position: "after" },
      },
      code: "INVALID_TOOL_ARGUMENTS",
      locator: undefined,
      schemaValidation: true,
    },
  ];

  for (const scenario of cases) {
    const harness = createWordBridgeHarness({ paragraphs: ["正文"] });
    await assert.rejects(
      harness.bridge.execute([scenario.tool]),
      error => {
        assert.equal(error.code, scenario.code, scenario.name);
        assert.equal(error.details.completedToolCalls, 0, scenario.name);
        assert.equal(error.details.partialMutationPossible, false, scenario.name);
        if (!scenario.schemaValidation) {
          assert.equal(error.details.locator, scenario.locator, scenario.name);
        }
        if (scenario.occurrence !== undefined) {
          assert.equal(error.details.occurrence, scenario.occurrence, scenario.name);
        }
        if (scenario.matchCount !== undefined) {
          assert.equal(error.details.matchCount, scenario.matchCount, scenario.name);
        }
        return true;
      },
    );
    assert.deepEqual(harness.paragraphs, ["正文"], scenario.name);
  }
});

test("word bridge reports ambiguous footnote targets without possible mutation", async () => {
  const harness = createWordBridgeHarness({ paragraphs: ["RTO 说明一", "RTO 说明二"] });

  await assert.rejects(
    harness.bridge.execute([{
      name: "word_manage_long_document",
      arguments: { action: "addFootnote", search: "RTO", noteText: "说明" },
    }]),
    error => {
      assert.equal(error.code, "INVALID_TARGET");
      assert.match(error.message, /实际匹配 2 个/);
      assert.equal(error.details.matchCount, 2);
      assert.equal(error.details.partialMutationPossible, false);
      return true;
    },
  );

  assert.deepEqual(harness.paragraphs, ["RTO 说明一", "RTO 说明二"]);
});

test("word bridge keeps possible mutation true after an earlier completed write", async () => {
  const harness = createWordBridgeHarness({ paragraphs: ["原正文"] });

  await assert.rejects(
    harness.bridge.execute([
      { name: "word_append_paragraph", arguments: { text: "已写入" } },
      {
        name: "word_manage_long_document",
        arguments: { action: "addFootnote", search: "不存在", noteText: "说明" },
      },
    ]),
    error => {
      assert.equal(error.code, "INVALID_TARGET");
      assert.equal(error.details.toolCallIndex, 1);
      assert.equal(error.details.completedToolCalls, 1);
      assert.equal(error.details.partialMutationPossible, true);
      return true;
    },
  );

  assert.deepEqual(harness.paragraphs, ["原正文", "已写入"]);
});

test("word bridge keeps title-page, non-empty header, page fields, and removal paths compatible", async () => {
  const calls = [];
  const paragraph = {
    AddText(value) { calls.push(["appendText", value]); },
    AddPageNumber() { calls.push(["pageNumber"]); },
    AddPagesCount() { calls.push(["pagesCount"]); },
  };
  const content = {
    SetText(value) { calls.push(["setText", value]); },
    GetElement(index) { return index === 0 ? paragraph : null; },
    Push(value) { calls.push(["push", value]); },
  };
  const section = {
    SetTitlePage(value) { calls.push(["titlePage", value]); },
    GetHeader(type, create) { calls.push(["getHeader", type, create]); return content; },
    GetFooter(type, create) { calls.push(["getFooter", type, create]); return content; },
    RemoveHeader(type) { calls.push(["removeHeader", type]); },
  };
  const harness = createWordBridgeHarness({ sections: [section] });

  await harness.bridge.execute([{
    name: "word_set_page_layout",
    arguments: { titlePage: true },
  }]);
  assert.deepEqual(calls, [["titlePage", true]]);

  await harness.bridge.execute([{
    name: "word_set_header_footer",
    arguments: { kind: "header", type: "first", text: "首页内容" },
  }]);
  await harness.bridge.execute([{
    name: "word_set_header_footer",
    arguments: { kind: "footer", type: "default", pageNumber: true, pagesCount: true },
  }]);
  await harness.bridge.execute([{
    name: "word_set_header_footer",
    arguments: { kind: "header", type: "first", action: "remove" },
  }]);

  assert.deepEqual(calls.slice(1), [
    ["getHeader", "first", true],
    ["setText", "首页内容"],
    ["getFooter", "default", true],
    ["pageNumber"],
    ["appendText", " / "],
    ["pagesCount"],
    ["removeHeader", "first"],
  ]);
});

test("word bridge exact paragraph targets ignore the terminal paragraph mark", async () => {
  const harness = createWordBridgeHarness({
    paragraphs: ["内部测试材料\r\n", "下一页"],
  });

  const result = await harness.bridge.execute([{
    name: "word_insert_page_break",
    arguments: {
      search: "内部测试材料",
      matchMode: "exact",
      occurrence: 1,
      position: "after",
    },
  }]);

  assert.equal(result.changed, 1);
  assert.equal(result.results[0].pageBreaks, 1);
  assert.equal(harness.paragraphObjects[0].pageBreakAfter, true);
  assert.equal(harness.paragraphObjects[1].pageBreakAfter, undefined);
});

test("word inspection reports paragraph spacing and whole-table alignment", async () => {
  const harness = createWordBridgeHarness({ paragraphs: ["表格前"] });

  const result = await harness.bridge.execute([
    {
      name: "word_format_paragraphs",
      arguments: { paragraphIndexes: [1], spacingAfterPt: 6 },
    },
    {
      name: "word_add_table",
      arguments: { rows: 1, cols: 1, align: "center", data: [["值"]] },
    },
    {
      name: "word_inspect",
      arguments: { includeStructure: true },
    },
  ]);

  const inspection = result.results[2];
  assert.equal(inspection.paragraphDetails[0].spacingAfterPt, 6);
  assert.equal(inspection.tableDetails[0].align, "center");
});

test("word bridge aggregates all static semantic failures before creating history", async () => {
  const harness = createWordBridgeHarness({ paragraphs: ["第一段"] });

  await assert.rejects(
    harness.bridge.execute([
      {
        name: "word_manage_section",
        arguments: { action: "create" },
      },
      {
        name: "word_manage_fields",
        arguments: { action: "add" },
      },
      {
        name: "word_insert_page_break",
        arguments: {},
      },
    ]),
    error => {
      assert.equal(error.code, "INVALID_TOOL_ARGUMENTS");
      assert.equal(error.details.completedToolCalls, 0);
      assert.equal(error.details.partialMutationPossible, false);
      assert.equal(
        JSON.stringify(error.details.validationErrors.map(item => [
          item.toolCallIndex,
          item.tool,
          item.path,
        ])),
        JSON.stringify([
          [0, "word_manage_section", "arguments.boundary"],
          [1, "word_manage_fields", "arguments.instruction"],
          [2, "word_insert_page_break", "arguments.target"],
        ]),
      );
      return true;
    },
  );

  assert.equal(harness.historyPoints, 0);
});

test("word bridge preflight normalizes and aggregates without creating history", () => {
  const harness = createWordBridgeHarness({ paragraphs: ["第一段"] });

  const result = harness.bridge.preflight([
    {
      name: "word_append_paragraph",
      arguments: { text: "兼容字段", spacingBefore: 6, spacingBeforePt: 8 },
    },
    {
      name: "word_insert_page_break",
      arguments: {},
    },
  ]);

  assert.equal(result.toolCalls[0].arguments.spacingBefore, undefined);
  assert.equal(result.toolCalls[0].arguments.spacingBeforePt, 8);
  assert.equal(result.argumentNormalizations[0].kind, "canonicalWins");
  assert.deepEqual(
    Array.from(result.validationErrors, item => [
      item.toolCallIndex,
      item.path,
      item.keyword,
    ]),
    [[1, "arguments.target", "semantic"]],
  );
  assert.equal(harness.historyPoints, 0);
});

test("word bridge uses ONLYOFFICE one-based placeholders for multilevel numbering", async () => {
  const customTypes = [];
  const levels = Array.from({ length: 9 }, (_, index) => ({
    index,
    SetCustomType(format, text, align) {
      customTypes.push({ index, format, text, align });
      return true;
    },
    SetStart() { return true; },
    SetRestart() { return true; },
  }));
  const harness = createWordBridgeHarness({
    paragraphs: ["一级", "二级"],
    createNumbering() {
      return { GetLevel(index) { return levels[index]; } };
    },
  });

  const result = await harness.bridge.execute([{
    name: "word_set_numbering",
    arguments: {
      paragraphIndexes: [1, 2],
      kind: "multilevel",
      level: 1,
      levels: [
        { level: 0, format: "decimal" },
        { level: 1, format: "lowerLetter" },
      ],
    },
  }]);

  assert.equal(result.changed, 2);
  assert.deepEqual(customTypes.map(entry => entry.text), ["%1.", "%1.%2."]);
  assert.ok(harness.paragraphObjects.every(paragraph => paragraph.numbering === levels[1]));
});

test("word bridge normalizes compatible numbering fields and applies level paragraph indents", async () => {
  const customTypes = [];
  const restarts = [];
  const indents = [];
  const levels = Array.from({ length: 9 }, (_, index) => ({
    index,
    SetCustomType(format, text, align) { customTypes.push({ index, format, text, align }); },
    SetStart() {},
    SetRestart(value) { restarts.push({ index, value }); },
    GetParaPr() {
      return {
        SetIndLeft(value) { indents.push({ index, field: "left", value }); },
        SetIndFirstLine(value) { indents.push({ index, field: "firstLine", value }); },
      };
    },
  }));
  const numbering = {
    GetInternalId() { return "normalized-list"; },
    GetLevel(index) { return levels[index]; },
  };
  const harness = createWordBridgeHarness({
    paragraphs: ["一级", "二级"],
    createNumbering() { return numbering; },
  });
  const call = {
    name: "word_set_numbering",
    arguments: {
      kind: "multilevel",
      levels: [
        { numberFormat: "lowerRoman", hangingIndentPt: 18, leftIndentPt: 36, restart: 1 },
        { level: 1, format: "decimal", numberFormat: "upperLetter", firstLineIndentPt: -12, hangingIndentPt: 20 },
      ],
      assignments: [
        { paragraphIndex: 1, level: 0 },
        { paragraphIndex: 2, level: 1 },
      ],
    },
  };

  const preflight = harness.bridge.preflight([call]);
  assert.equal(preflight.validationErrors.length, 0);
  assert.equal(preflight.toolCalls[0].arguments.levels[0].level, 0);
  assert.equal(preflight.toolCalls[0].arguments.levels[0].format, "lowerRoman");
  assert.equal(preflight.toolCalls[0].arguments.levels[0].firstLineIndentPt, -18);
  assert.equal(preflight.toolCalls[0].arguments.levels[0].restart, true);
  assert.equal(preflight.toolCalls[0].arguments.levels[1].format, "decimal");
  assert.equal(preflight.toolCalls[0].arguments.levels[1].numberFormat, undefined);
  assert.equal(preflight.toolCalls[0].arguments.levels[1].hangingIndentPt, undefined);
  assert.ok(preflight.argumentNormalizations.some(item => item.kind === "canonicalWins"));

  const result = await harness.bridge.execute([call]);
  assert.equal(result.results[0].listGroupId, "normalized-list");
  assert.deepEqual(customTypes.map(item => [item.format, item.text]), [
    ["lowerRoman", "%1."],
    ["decimal", "%1.%2."],
  ]);
  assert.deepEqual(restarts, [{ index: 0, value: true }]);
  assert.deepEqual(indents, [
    { index: 0, field: "firstLine", value: -360 },
    { index: 0, field: "left", value: 720 },
    { index: 1, field: "firstLine", value: -240 },
  ]);
});

test("word bridge rejects invalid numbering targets and duplicate levels before history", async () => {
  for (const argumentsValue of [
    {},
    {
      levels: [{ level: 0 }, { level: 0 }],
      assignments: [{ paragraphIndex: 1, level: 0 }],
    },
    {
      assignments: [{ paragraphIndexes: [1, 2], level: 0 }],
    },
    {
      assignments: [{ paragraphIndex: 1, search: "一级", level: 0 }],
    },
  ]) {
    const harness = createWordBridgeHarness({ paragraphs: ["一级", "二级"] });
    await assert.rejects(
      harness.bridge.execute([{ name: "word_set_numbering", arguments: argumentsValue }]),
      error => error.code === "INVALID_TOOL_ARGUMENTS" && error.details.partialMutationPossible === false,
    );
    assert.equal(harness.historyPoints, 0);
  }
});

test("word bridge applies multilevel assignments through one shared numbering instance and can continue it", async () => {
  let createCalls = 0;
  let numbering;
  const levels = Array.from({ length: 9 }, (_, index) => ({
    index,
    GetLevelIndex() { return index; },
    GetNumbering() { return numbering; },
    SetCustomType() { return true; },
    SetStart() { return true; },
    SetRestart() { return true; },
  }));
  numbering = {
    GetInternalId() { return "list-shared"; },
    GetLevel(index) { return levels[index]; },
  };
  const harness = createWordBridgeHarness({
    paragraphs: ["一级一", "二级一", "一级二", "二级二", "续项"],
    createNumbering() { createCalls += 1; return numbering; },
  });

  const created = await harness.bridge.execute([{
    name: "word_set_numbering",
    arguments: {
      kind: "multilevel",
      assignments: [
        { paragraphIndex: 1, level: 0 },
        { paragraphIndex: 2, level: 1 },
        { paragraphIndex: 3, level: 0 },
        { paragraphIndex: 4, level: 1 },
      ],
    },
  }]);
  const continued = await harness.bridge.execute([{
    name: "word_set_numbering",
    arguments: {
      continueFrom: { paragraphId: "para-4" },
      assignments: [{ paragraphIndex: 5, level: 0 }],
    },
  }]);

  assert.equal(createCalls, 1);
  assert.deepEqual(harness.paragraphObjects.slice(0, 5).map(paragraph => paragraph.numbering.index), [0, 1, 0, 1, 0]);
  assert.equal(created.results[0].listGroupId, "list-shared");
  assert.equal(continued.results[0].listGroupId, "list-shared");
  assert.equal(continued.results[0].continuedFromParagraphIndex, 4);
});

test("word advanced inspection groups fresh ApiNumbering wrappers by their shared internal numbering id", async () => {
  const internalNumbering = {
    Id: "num-shared-internal",
    GetId() { return this.Id; },
  };
  const levels = [0, 1].map(index => ({
    index,
    GetLevelIndex() { return this.index; },
    GetNumbering() { return { Num: internalNumbering }; },
  }));
  const numbering = {
    Num: internalNumbering,
    GetLevel(index) { return levels[index]; },
  };
  const harness = createWordBridgeHarness({
    paragraphs: ["父级一", "子级一", "父级二"],
    createNumbering() { return numbering; },
  });

  const updated = await harness.bridge.execute([{
    name: "word_set_numbering",
    arguments: {
      kind: "multilevel",
      assignments: [
        { paragraphIndex: 1, level: 0 },
        { paragraphIndex: 2, level: 1 },
        { paragraphIndex: 3, level: 0 },
      ],
    },
  }]);
  const inspected = await harness.bridge.execute([{
    name: "word_inspect_advanced",
    arguments: { includeNumbering: true },
  }]);

  assert.equal(updated.results[0].listGroupId, "num-shared-internal");
  assert.deepEqual(inspected.results[0].numbering.map(item => item.listGroup), [1, 1, 1]);
  assert.deepEqual(inspected.results[0].numbering.map(item => item.listGroupId), [
    "num-shared-internal", "num-shared-internal", "num-shared-internal",
  ]);
  assert.deepEqual(inspected.results[0].numbering.map(item => item.level), [0, 1, 0]);
});

test("word set_list reports a distinct listGroupId for every independent call", async () => {
  let created = 0;
  const harness = createWordBridgeHarness({
    paragraphs: ["第一项", "第二项"],
    createNumbering() {
      created += 1;
      const id = "set-list-" + created;
      const level = { index: 0 };
      return {
        GetInternalId() { return id; },
        GetLevel() { return level; },
      };
    },
  });

  const first = await harness.bridge.execute([{
    name: "word_set_list",
    arguments: { listType: "numbered", paragraphIndex: 1 },
  }]);
  const second = await harness.bridge.execute([{
    name: "word_set_list",
    arguments: { listType: "numbered", paragraphIndex: 2 },
  }]);

  assert.equal(first.results[0].listGroupId, "set-list-1");
  assert.equal(second.results[0].listGroupId, "set-list-2");
  assert.equal(first.results[0].independentPerCall, true);
  assert.notEqual(first.results[0].listGroupId, second.results[0].listGroupId);
});

test("word bridge applies cell paragraph styles to every table cell and rejects table-style type mixing", async () => {
  const harness = createWordBridgeHarness({ paragraphs: ["表格锚点"] });
  await harness.bridge.execute([{
    name: "word_manage_style",
    arguments: { action: "create", name: "Table Text", type: "paragraph" },
  }]);
  const historyBeforeInvalid = harness.historyPoints;
  await assert.rejects(
    harness.bridge.execute([{
      name: "word_add_table",
      arguments: { rows: 1, cols: 1, tableStyleName: "Table Text" },
    }]),
    error => /样式/.test(error.message) && error.details.partialMutationPossible === false,
  );
  assert.equal(harness.historyPoints, historyBeforeInvalid);

  const result = await harness.bridge.execute([{
    name: "word_add_table",
    arguments: {
      rows: 2,
      cols: 2,
      data: [["甲"]],
      cellParagraphStyleName: "Table Text",
    },
  }]);
  const table = harness.createdTables.at(-1);
  const cellStyles = [];
  for (let row = 0; row < 2; row += 1) {
    for (let column = 0; column < 2; column += 1) {
      cellStyles.push(table.GetRow(row).GetCell(column).GetContent().GetAllParagraphs()[0].style.GetName());
    }
  }
  assert.equal(result.changed, 4);
  assert.deepEqual(cellStyles, ["Table Text", "Table Text", "Table Text", "Table Text"]);
});

test("word editor preflight tracks same-batch style types and rejects invalid merges before history", async () => {
  const harness = createWordBridgeHarness({ paragraphs: ["锚点"] });
  const historyBefore = harness.historyPoints;
  await assert.rejects(
    harness.bridge.execute([
      { name: "word_manage_style", arguments: { action: "create", name: "Wrong Table Style", type: "paragraph" } },
      { name: "word_add_table", arguments: { rows: 1, cols: 1, tableStyleName: "Wrong Table Style" } },
    ]),
    error => /样式/.test(error.message)
      && error.details.completedToolCalls === 0
      && error.details.partialMutationPossible === false,
  );
  await assert.rejects(
    harness.bridge.execute([{
      name: "word_edit_table",
      arguments: { action: "mergeCells", tableIndex: 1, rowStart: 1, rowEnd: 2, columnStart: 1, columnEnd: 1 },
    }]),
    error => /表格范围/.test(error.message)
      && error.details.completedToolCalls === 0
      && error.details.partialMutationPossible === false,
  );
  assert.equal(harness.historyPoints, historyBefore);
});

test("word advanced inspection reports stable paragraph ids and actual style usage", async () => {
  const harness = createWordBridgeHarness({ paragraphs: ["三级标题", "正文"] });
  await harness.bridge.execute([
    { name: "word_manage_style", arguments: { action: "create", name: "Heading 3", type: "paragraph" } },
    { name: "word_format_paragraphs", arguments: { paragraphId: "para-1", styleName: "Heading 3" } },
  ]);
  const inspected = await harness.bridge.execute([
    { name: "word_inspect", arguments: { includeStructure: true } },
    { name: "word_inspect_advanced", arguments: { includeStyles: true } },
  ]);

  assert.equal(inspected.results[0].paragraphDetails[0].paragraphId, "para-1");
  assert.equal(inspected.results[0].paragraphDetails[0].internalId, "internal-1");
  const headingStyle = inspected.results[1].styles.find(style => style.name === "Heading 3");
  assert.equal(headingStyle.usageCount, 1);
  assert.equal(headingStyle.tableParagraphUsageCount, 0);
});

test("word bridge resolves before and after section boundaries without moving the heading into the old section", async () => {
  const afterHarness = createWordBridgeHarness({ paragraphs: ["旧节", "新节标题"] });
  const after = await afterHarness.bridge.execute([{
    name: "word_manage_section",
    arguments: { action: "create", boundary: { position: "after", paragraphId: "para-1" } },
  }]);
  assert.equal(afterHarness.createdSectionParagraphs[0], afterHarness.paragraphObjects[0]);
  assert.equal(after.results[0].boundary.endParagraphIndex, 1);

  const beforeHarness = createWordBridgeHarness({ paragraphs: ["旧节", "新节标题"] });
  const before = await beforeHarness.bridge.execute([{
    name: "word_manage_section",
    arguments: { action: "create", boundary: { position: "before", paragraphId: "para-2" } },
  }]);
  assert.equal(beforeHarness.createdSectionParagraphs[0], beforeHarness.paragraphObjects[0]);
  assert.equal(before.results[0].boundary.endParagraphIndex, 1);
});

test("word bridge emits a real SEQ field for captions and protects non-empty TOC targets", async () => {
  const captionHarness = createWordBridgeHarness({ paragraphs: ["图形对象"] });
  const caption = await captionHarness.bridge.execute([{
    name: "word_manage_long_document",
    arguments: { action: "addCaption", paragraphIndex: 1, label: "图", text: "系统架构", insertAt: "after" },
  }]);
  assert.match(captionHarness.fieldInstructions[0], /^SEQ /);
  assert.equal(caption.results[0].dynamic, true);

  const tocHarness = createWordBridgeHarness({ paragraphs: ["这里不是空段落"] });
  await assert.rejects(
    tocHarness.bridge.execute([{
      name: "word_manage_long_document",
      arguments: { action: "addToc", paragraphIndex: 1 },
    }]),
    error => /只允许替换空段落/.test(error.message) && error.details.partialMutationPossible === false,
  );
  assert.equal(tocHarness.historyPoints, 0);
});

test("word bridge supplies the required chart number formats", async () => {
  const harness = createWordBridgeHarness({ paragraphs: ["图表"] });

  const defaultFormats = await harness.bridge.execute([{
    name: "word_add_chart",
    arguments: {
      data: [[1, 2], [3, 4]],
      seriesNames: ["收入", "成本"],
      categories: ["一月", "二月"],
      paragraphIndex: 1,
    },
  }]);
  await harness.bridge.execute([{
    name: "word_add_chart",
    arguments: {
      data: [[0.1, 0.2]],
      numberFormats: ["0.0%"],
      paragraphIndex: 1,
    },
  }]);

  assert.equal(defaultFormats.changed, 1);
  assert.equal(harness.createdCharts.length, 2);
  assert.deepEqual(harness.createdCharts[0].arguments[7], ["General", "General"]);
  assert.deepEqual(harness.createdCharts[1].arguments[7], ["0.0%"]);
  assert.equal(harness.paragraphObjects[0].drawings.length, 2);
});

test("word bridge inspects and replaces the document macro collection through plugin methods", async () => {
  const macros = {
    macrosArray: [{ name: "Main", value: "(function(){ return 1; })();" }],
    current: 0,
  };
  const harness = createWordBridgeHarness({
    text: "宏文档",
    executeMethodResponses: { GetMacros: JSON.stringify(macros) },
  });

  const inspected = await harness.bridge.execute([{
    name: "word_inspect_macros",
    arguments: { kind: "office" },
  }]);
  const updated = await harness.bridge.execute([{
    name: "word_set_macros",
    arguments: { macros: macros.macrosArray, current: 0 },
  }]);

  assert.deepEqual(inspected.results[0].content, macros);
  assert.equal(inspected.results[0].kind, "office");
  assert.equal(inspected.needsSave, false);
  assert.equal(updated.needsSave, true);
  assert.equal(updated.changed, 1);
  assert.deepEqual(harness.executeMethodCalls.map(call => call.name), ["GetMacros", "SetMacros"]);
  assert.deepEqual(JSON.parse(harness.executeMethodCalls[1].args[0]), macros);
});

test("word bridge applies supported document editing restrictions", async () => {
  const harness = createWordBridgeHarness({ text: "受保护文档" });

  const protectedResult = await harness.bridge.execute([{
    name: "word_set_protection",
    arguments: { action: "protect", mode: "comments" },
  }]);
  const unprotectedResult = await harness.bridge.execute([{
    name: "word_set_protection",
    arguments: { action: "unprotect" },
  }]);

  assert.equal(protectedResult.results[0].restriction, "comments");
  assert.equal(unprotectedResult.results[0].restriction, "none");
  assert.deepEqual(harness.executeMethodCalls, [
    { name: "SetEditingRestrictions", args: ["comments"] },
    { name: "SetEditingRestrictions", args: ["none"] },
  ]);
});

test("word bridge manages and inspects advanced content control properties", async () => {
  const calls = [];
  const listItems = [
    { GetText() { return "甲"; }, GetValue() { return "a"; } },
    { GetText() { return "乙"; }, GetValue() { return "b"; } },
  ];
  const control = {
    tag: "old",
    GetClassType() { return "inlineLvlSdt"; },
    GetId() { return "control-1"; },
    GetTag() { return this.tag; },
    GetAlias() { return "选择项"; },
    GetLock() { return "unlocked"; },
    GetAppearance() { return "hidden"; },
    GetPlaceholderText() { return "请选择"; },
    IsCheckBox() { return false; },
    GetDate() { return new Date("2026-07-23T00:00:00.000Z"); },
    GetDataBinding() {
      return { prefixMapping: "xmlns:x='urn:test'", storeItemID: "part-1", xpath: "/x:data/x:value" };
    },
    GetDropdownList() {
      return { GetAllItems() { return listItems; } };
    },
    SetTag(value) { this.tag = value; calls.push(["tag", value]); return true; },
    SetAppearance(value) { calls.push(["appearance", value]); },
    SetDateFormat(value) { calls.push(["dateFormat", value]); return true; },
    SetDate(value) { calls.push(["date", value.toISOString()]); return true; },
    SetDataBinding(value) { calls.push(["dataBinding", value]); return true; },
    UpdateFromXmlMapping() { calls.push(["updateFromXml"]); return true; },
    SelectListItem(value) { calls.push(["selectedValue", value]); return true; },
  };
  const harness = createWordBridgeHarness({ text: "内容控件", contentControls: [control] });

  const updated = await harness.bridge.execute([{
    name: "word_manage_content_control",
    arguments: {
      action: "update",
      index: 1,
      tag: "choice",
      appearance: "hidden",
      date: "2026-07-23T00:00:00.000Z",
      dateFormat: "yyyy-MM-dd",
      selectedValue: "b",
      dataBinding: {
        prefixMapping: "xmlns:x='urn:test'",
        storeItemId: "part-1",
        xpath: "/x:data/x:value",
      },
      updateFromXml: true,
    },
  }]);
  const inspected = await harness.bridge.execute([{
    name: "word_inspect_advanced",
    arguments: { includeContentControls: true },
  }]);

  assert.equal(updated.changed, 1);
  assert.deepEqual(JSON.parse(JSON.stringify(calls)), [
    ["tag", "choice"],
    ["appearance", "hidden"],
    ["dateFormat", "yyyy-MM-dd"],
    ["date", "2026-07-23T00:00:00.000Z"],
    ["dataBinding", {
      prefixMapping: "xmlns:x='urn:test'",
      storeItemID: "part-1",
      xpath: "/x:data/x:value",
    }],
    ["updateFromXml"],
    ["selectedValue", "b"],
  ]);
  assert.deepEqual(inspected.results[0].contentControls[0], {
    index: 1,
    type: "inlineLvlSdt",
    tag: "choice",
    title: "选择项",
    lock: "unlocked",
    appearance: "hidden",
    placeholder: "请选择",
    checked: null,
    date: "2026-07-23T00:00:00.000Z",
    dataBinding: {
      prefixMapping: "xmlns:x='urn:test'",
      storeItemID: "part-1",
      xpath: "/x:data/x:value",
    },
    listItems: [
      { display: "甲", value: "a" },
      { display: "乙", value: "b" },
    ],
  });
});

test("word bridge passes a list value as the default selection when creating a content control", async () => {
  const created = [];
  const control = {
    GetId() { return "combo-1"; },
    GetTag() { return ""; },
  };
  const harness = createWordBridgeHarness({
    text: "选择",
    addComboBoxContentControl(list, selected) {
      created.push({ list, selected });
      return control;
    },
  });

  const result = await harness.bridge.execute([{
    name: "word_manage_content_control",
    arguments: {
      action: "add",
      kind: "comboBox",
      items: [
        { display: "甲", value: "a" },
        { display: "乙", value: "b" },
      ],
      selectedValue: "b",
    },
  }]);

  assert.equal(result.changed, 1);
  assert.deepEqual(JSON.parse(JSON.stringify(created)), [{
    list: [
      { display: "甲", value: "a" },
      { display: "乙", value: "b" },
    ],
    selected: "b",
  }]);
});

test("word bridge writes scalar and formatted cells in regular and nested tables", async () => {
  const harness = createWordBridgeHarness({ paragraphs: ["正文"] });
  const result = await harness.bridge.execute([
    {
      name: "word_add_table",
      arguments: {
        rows: 2,
        cols: 3,
        align: "Centre",
        fontFamily: "宋体",
        fontSize: 9,
        data: [
          [
            {
              text: "指标",
              fontFamily: "微软雅黑",
              bold: true,
              backgroundColor: "#E8EEF5",
              align: "center",
              verticalAlign: "center",
              widthPercent: 40,
            },
            42,
            true,
          ],
          [null, "嵌套目标"],
        ],
      },
    },
    {
      name: "word_add_nested_table",
      arguments: {
        tableIndex: 1,
        row: 2,
        column: 1,
        rows: 1,
        cols: 2,
        align: "RIGHT",
        fontFamily: "宋体",
        data: [[
          { text: "子标题", bold: true, textColor: "#2E74B5" },
          "子值",
        ]],
      },
    },
  ]);

  const table = harness.tables[0];
  const formattedCell = table.GetRow(0).GetCell(0);
  const formattedParagraph = formattedCell.GetContent().GetAllParagraphs()[0];
  const formattedRun = formattedParagraph.runs[0];
  const inheritedRun = table.GetRow(0).GetCell(1).GetContent().GetAllParagraphs()[0].runs[0];
  const nestedTable = table.GetRow(1).GetCell(0).nestedTables[0];
  const nestedRun = nestedTable.GetRow(0).GetCell(0).GetContent().GetAllParagraphs()[0].runs[0];

  assert.equal(result.changed, 7);
  assert.equal(table.align, "center");
  assert.equal(nestedTable.align, "right");
  assert.equal(formattedParagraph.text, "指标");
  assert.equal(formattedRun.fontFamily, "微软雅黑");
  assert.equal(formattedRun.bold, true);
  assert.equal(formattedCell.backgroundColor, "#E8EEF5");
  assert.equal(formattedCell.verticalAlign, "center");
  assert.deepEqual(formattedCell.width, { kind: "percent", value: 40 });
  assert.equal(formattedParagraph.align, "center");
  assert.equal(inheritedRun.fontFamily, "宋体");
  assert.equal(inheritedRun.fontSize, 18);
  assert.equal(table.GetRow(0).GetCell(2).GetContent().GetAllParagraphs()[0].text, "true");
  assert.equal(table.GetRow(1).GetCell(0).GetContent().GetAllParagraphs()[0].text, "");
  assert.equal(table.GetRow(1).GetCell(2).GetContent().GetAllParagraphs()[0].text, "");
  assert.equal(nestedRun.text, "子标题");
  assert.equal(nestedRun.bold, true);
  assert.equal(nestedRun.color, "#2E74B5");
  assert.doesNotMatch(harness.text, /\[object Object\]/);
});

test("word bridge rejects invalid formatted table cells before mutation", async () => {
  const harness = createWordBridgeHarness({ paragraphs: ["正文"] });

  await assert.rejects(
    harness.bridge.execute([{
      name: "word_add_table",
      arguments: {
        rows: 1,
        cols: 4,
        data: [[
          { text: "错误颜色字段", textColor: "#FFFFFF" },
          { bold: true },
          { text: { nested: true } },
          { text: "错误单位", leftIndent: 720 },
        ]],
      },
    }]),
    error => {
      assert.equal(error.code, "INVALID_TOOL_ARGUMENTS");
      assert.equal(error.details.completedToolCalls, 0);
      assert.equal(error.details.partialMutationPossible, false);
      assert.deepEqual(
        Array.from(error.details.validationErrors, item => item.path),
        [
          "arguments.data[0][1].text",
          "arguments.data[0][2].text",
          "arguments.data[0][3].leftIndent",
        ],
      );
      return true;
    },
  );
  assert.equal(harness.historyPoints, 0);
  assert.equal(harness.createdTables.length, 0);
});

test("word bridge inserts tables at top-level paragraph, search, table, cursor, and end positions", async () => {
  const relativeHarness = createWordBridgeHarness({
    paragraphs: ["甲锚点", "正文", "乙锚点"],
    getPosInParentSupported: false,
  });
  const relativeResult = await relativeHarness.bridge.execute([
    {
      name: "word_add_table",
      arguments: {
        rows: 1,
        cols: 1,
        data: [["搜索前"]],
        insertAt: "before",
        search: "锚点",
        occurrence: 2,
      },
    },
    {
      name: "word_add_table",
      arguments: {
        rows: 1,
        cols: 1,
        data: [["表前"]],
        insertAt: "before",
        tableIndex: 1,
      },
    },
    {
      name: "word_add_table",
      arguments: {
        rows: 1,
        cols: 1,
        data: [["段后"]],
        insertAt: "after",
        paragraphIndex: 1,
      },
    },
  ]);

  assert.deepEqual(relativeHarness.elements.map(element => (
    element.GetClassType() === "paragraph" ? element.GetText() : element.GetRow(0).GetCell(0).GetContent().GetElement(0).GetText()
  )), ["甲锚点", "段后", "正文", "表前", "搜索前", "乙锚点"]);
  assert.deepEqual(
    JSON.parse(JSON.stringify(relativeResult.results.map(entry => ({
      insertAt: entry.insertAt,
      tableIndex: entry.tableIndex,
      anchor: entry.anchor,
    })))),
    [
      {
        insertAt: "before",
        tableIndex: 1,
        anchor: {
          type: "search",
          search: "锚点",
          matchCase: false,
          matchMode: "contains",
          occurrence: 2,
          paragraphIndex: 3,
          position: 2,
        },
      },
      {
        insertAt: "before",
        tableIndex: 1,
        anchor: { type: "table", tableIndex: 1, position: 2 },
      },
      {
        insertAt: "after",
        tableIndex: 1,
        anchor: { type: "paragraph", paragraphIndex: 1, position: 0 },
      },
    ],
  );

  const cursorHarness = createWordBridgeHarness({ paragraphs: ["正文"] });
  const cursorResult = await cursorHarness.bridge.execute([
    {
      name: "word_add_table",
      arguments: { rows: 1, cols: 1, insertAt: "current" },
    },
    {
      name: "word_add_table",
      arguments: { rows: 1, cols: 1 },
    },
  ]);
  assert.equal(cursorHarness.insertContentCalls.length, 1);
  assert.deepEqual(Array.from(cursorResult.results, entry => entry.insertAt), ["current", "end"]);
  assert.deepEqual(Array.from(cursorResult.results, entry => entry.tableIndex), [1, 2]);
});

test("word bridge puts a page-break paragraph immediately before a positioned table", async () => {
  const harness = createWordBridgeHarness({ paragraphs: ["第一页", "第二页"] });

  const result = await harness.bridge.execute([{
    name: "word_add_table",
    arguments: {
      rows: 1,
      cols: 1,
      data: [["分页表"]],
      insertAt: "before",
      paragraphIndex: 2,
      pageBreakBefore: true,
    },
  }]);

  assert.equal(harness.elements.length, 4);
  assert.equal(harness.elements[1].GetClassType(), "paragraph");
  assert.equal(harness.elements[1].pageBreakBefore, true);
  assert.equal(harness.elements[2].GetClassType(), "table");
  assert.equal(result.results[0].tableIndex, 1);
  assert.equal(result.results[0].pageBreakBefore, true);
  assert.deepEqual(harness.addElementCalls.map(call => call.position), [1, 2]);
});

test("word bridge rejects conflicting, missing, out-of-range, and nested table anchors before insertion", async () => {
  const conflictHarness = createWordBridgeHarness({ paragraphs: ["锚点"] });
  await assert.rejects(
    conflictHarness.bridge.execute([{
      name: "word_add_table",
      arguments: {
        rows: 1,
        cols: 1,
        insertAt: "before",
        paragraphIndex: 1,
        search: "锚点",
      },
    }]),
    /必须且只能提供/,
  );
  assert.equal(conflictHarness.createdTables.length, 0);
  assert.equal(conflictHarness.tables.length, 0);

  await assert.rejects(
    conflictHarness.bridge.execute([{
      name: "word_add_table",
      arguments: { rows: 1, cols: 1, insertAt: "after", tableIndex: 1 },
    }]),
    /tableIndex 超出文档表格范围/,
  );
  assert.equal(conflictHarness.createdTables.length, 0);

  const nestedHarness = createWordBridgeHarness({ paragraphs: ["单元格内锚点"] });
  nestedHarness.paragraphObjects[0].parentTable = {};
  await assert.rejects(
    nestedHarness.bridge.execute([{
      name: "word_add_table",
      arguments: { rows: 1, cols: 1, insertAt: "before", paragraphIndex: 1 },
    }]),
    /嵌套表格/,
  );
  assert.equal(nestedHarness.createdTables.length, 0);
});

test("word bridge appends or replaces inline content controls in paragraphs and table cells", async () => {
  const paragraphHarness = createWordBridgeHarness({ paragraphs: ["前缀"] });
  const appendResult = await paragraphHarness.bridge.execute([{
    name: "word_manage_content_control",
    arguments: {
      action: "add",
      kind: "inline",
      paragraphIndex: 1,
      text: "追加",
      tag: "append-tag",
    },
  }]);
  assert.equal(paragraphHarness.text, "前缀追加");
  assert.equal(appendResult.results[0].contentMode, "append");

  const replaceResult = await paragraphHarness.bridge.execute([{
    name: "word_manage_content_control",
    arguments: {
      action: "add",
      kind: "inline",
      paragraphIndex: 1,
      contentMode: "replace",
      text: "",
      tag: "final-tag",
      title: "最终内容",
      placeholder: "请输入",
      lock: "both",
      appearance: "hidden",
    },
  }]);
  const paragraphControl = paragraphHarness.createdContentControls.at(-1);
  assert.equal(paragraphHarness.text, "");
  assert.equal(paragraphControl.text, "");
  assert.equal(paragraphControl.tag, "final-tag");
  assert.equal(paragraphControl.title, "最终内容");
  assert.equal(paragraphControl.placeholder, "请输入");
  assert.equal(paragraphControl.lock, "sdtContentLocked");
  assert.equal(paragraphControl.appearance, "hidden");
  assert.deepEqual(JSON.parse(JSON.stringify(replaceResult.results[0].target)), {
    type: "paragraph",
    paragraphIndex: 1,
  });

  const cellHarness = createWordBridgeHarness({ paragraphs: ["正文"] });
  await cellHarness.bridge.execute([{
    name: "word_add_table",
    arguments: { rows: 1, cols: 1, data: [["旧值"]] },
  }]);
  const cellResult = await cellHarness.bridge.execute([{
    name: "word_manage_content_control",
    arguments: {
      action: "add",
      kind: "inline",
      tableIndex: 1,
      row: 1,
      column: 1,
      contentMode: "replace",
      text: "最终值",
      tag: "cell-tag",
    },
  }]);
  const cellParagraph = cellHarness.tables[0].GetRow(0).GetCell(0).GetContent().GetElement(0);
  assert.equal(cellParagraph.GetText(), "最终值");
  assert.equal(cellParagraph.inlineControls.length, 1);
  assert.deepEqual(JSON.parse(JSON.stringify(cellResult.results[0].target)), {
    type: "cell",
    tableIndex: 1,
    row: 1,
    column: 1,
  });
});

test("word bridge validates inline content-control replacement targets before mutation", async () => {
  const harness = createWordBridgeHarness({ paragraphs: ["保留"] });

  await assert.rejects(
    harness.bridge.execute([{
      name: "word_manage_content_control",
      arguments: {
        action: "add",
        kind: "inline",
        tableIndex: 1,
        row: 1,
        contentMode: "replace",
        text: "新值",
      },
    }]),
    /tableIndex、row、column 必须同时提供/,
  );
  await assert.rejects(
    harness.bridge.execute([{
      name: "word_manage_content_control",
      arguments: {
        action: "add",
        kind: "inline",
        paragraphIndex: 1,
        current: true,
        contentMode: "replace",
        text: "新值",
      },
    }]),
    /目标互斥/,
  );
  await assert.rejects(
    harness.bridge.execute([{
      name: "word_manage_content_control",
      arguments: {
        action: "add",
        kind: "block",
        paragraphIndex: 1,
        contentMode: "replace",
        text: "新值",
      },
    }]),
    /仅支持 inline/,
  );

  assert.equal(harness.text, "保留");
  assert.equal(harness.createdContentControls.length, 0);
});

test("word bridge executes native replace, positioned table, and inline SDT in one history point", async () => {
  const harness = createWordBridgeHarness({
    paragraphs: ["A", "锚点"],
    nativeSearchAndReplace: true,
    staleTableCollections: true,
  });

  const result = await harness.bridge.execute([
    { name: "word_replace_text", arguments: { search: "A", replace: "B", matchCase: true } },
    {
      name: "word_add_table",
      arguments: {
        rows: 1,
        cols: 1,
        data: [["旧单元格"]],
        insertAt: "after",
        paragraphIndex: 1,
      },
    },
    {
      name: "word_manage_content_control",
      arguments: {
        action: "add",
        kind: "inline",
        tableIndex: 1,
        row: 1,
        column: 1,
        contentMode: "replace",
        text: "最终单元格",
        tag: "mixed",
      },
    },
  ]);

  assert.equal(result.changed, 3);
  assert.deepEqual(Array.from(result.results, entry => entry.name), [
    "word_replace_text",
    "word_add_table",
    "word_manage_content_control",
  ]);
  assert.equal(harness.historyPoints, 1);
  assert.equal(harness.callCommandCalls.filter(call => call.recalculate).length, 1);
  assert.equal(harness.executeMethodCalls.length, 0);
  assert.equal(harness.paragraphObjects[0].GetText(), "B");
  assert.equal(
    harness.tables[0].GetRow(0).GetCell(0).GetContent().GetElement(0).GetText(),
    "最终单元格",
  );
});

test("word bridge inserts a dynamic field into a zero-width placeholder range", async () => {
  const harness = createWordBridgeHarness({ paragraphs: ["标题", "正文"] });

  const result = await harness.bridge.execute([{
    name: "word_manage_fields",
    arguments: {
      action: "add",
      instruction: 'DATE \\\\@ "yyyy-MM-dd"',
      paragraphIndex: 2,
    },
  }]);

  assert.equal(result.changed, 1);
  assert.deepEqual(harness.fieldInstructions, ['DATE \\\\@ "yyyy-MM-dd"']);
  assert.equal(harness.paragraphs[0], "标题");
  assert.equal(harness.paragraphs[1], "正文\u200B");
});

test("word bridge rejects mutating format calls that contain no actual change", async () => {
  const harness = createWordBridgeHarness({ text: "主任委员：匿名" });

  await assert.rejects(
    harness.bridge.execute([{ name: "word_format_document", arguments: {} }]),
    /至少需要一个格式属性/,
  );
  await assert.rejects(
    harness.bridge.execute([{ name: "word_format_matches", arguments: { search: "匿名" } }]),
    /至少需要一个格式属性/,
  );
});

test("word bridge requires an explicit occurrence for a bookmark", async () => {
  const harness = createWordBridgeHarness({ text: "唯一目标" });

  await assert.rejects(
    harness.bridge.execute([{
      name: "word_add_bookmark",
      arguments: { search: "目标", name: "target" },
    }]),
    /occurrence 不能为空/,
  );
});

test("word bridge navigates to the document end without marking content dirty", async () => {
  const harness = createWordBridgeHarness({ text: "demo", pageCount: 7 });

  const result = await harness.bridge.execute([{
    name: "word_navigate",
    arguments: { target: "end" },
  }]);

  assert.equal(result.changed, 0);
  assert.equal(result.needsSave, false);
  assert.deepEqual(harness.navigatedPages, [6]);
  assert.equal(result.results[0].page, 7);
  assert.equal(result.results[0].pageCount, 7);
});

test("word bridge scrolls by relative pages without marking content dirty", async () => {
  const harness = createWordBridgeHarness({ text: "demo", pageCount: 10, currentPage: 4 });

  const down = await harness.bridge.execute([{
    name: "word_scroll",
    arguments: { direction: "down", pages: 3 },
  }]);
  const up = await harness.bridge.execute([{
    name: "word_scroll",
    arguments: { direction: "up", pages: 2 },
  }]);

  assert.equal(down.needsSave, false);
  assert.equal(up.needsSave, false);
  assert.deepEqual(harness.navigatedPages, [7, 5]);
  assert.equal(down.results[0].page, 8);
  assert.equal(up.results[0].page, 6);
});

test("word bridge selects a searched occurrence and reports its page", async () => {
  const harness = createWordBridgeHarness({ text: "甲乙甲乙", pageCount: 8, currentPage: 0, searchPage: 3 });

  const result = await harness.bridge.execute([{
    name: "word_navigate",
    arguments: { target: "search", search: "乙", occurrence: 2 },
  }]);

  assert.equal(result.needsSave, false);
  assert.equal(result.results[0].matches, 2);
  assert.equal(result.results[0].page, 4);
  assert.deepEqual(harness.rangeActions, [{ action: "select", start: 3 }]);
});

test("word bridge supports exact-range deletion, links, comments, and bookmarks", async () => {
  const harness = createWordBridgeHarness({ text: "目标 目标" });

  const result = await harness.bridge.execute([
    { name: "word_delete_matches", arguments: { search: "目标", occurrence: 1 } },
    { name: "word_add_hyperlink", arguments: { search: "目标", occurrence: 2, url: "https://example.test", screenTip: "示例" } },
    { name: "word_add_comment", arguments: { search: "目标", occurrence: 2, text: "请复核", author: "AI" } },
    { name: "word_add_bookmark", arguments: { search: "目标", occurrence: 2, name: "target_2" } },
  ]);

  assert.equal(result.changed, 4);
  assert.equal(result.needsSave, true);
  assert.deepEqual(harness.rangeActions, [
    { action: "delete", start: 0, end: 2 },
    { action: "hyperlink", start: 3, url: "https://example.test", screenTip: "示例", bookmarkName: "" },
    { action: "comment", start: 3, text: "请复核", author: "AI", userId: undefined },
    { action: "bookmark", start: 3, name: "target_2" },
  ]);
});

test("word bridge preserves replacement order in a batch", async () => {
  const harness = createWordBridgeHarness({ text: "A" });

  const result = await harness.bridge.execute([
    { name: "word_replace_text", arguments: { search: "A", replace: "B", matchCase: true } },
    { name: "word_replace_text", arguments: { search: "B", replace: "C", matchCase: true } },
  ]);

  assert.equal(harness.text, "C");
  assert.equal(result.changed, 2);
  assert.deepEqual(Array.from(result.results, entry => entry.search), ["A", "B"]);
});

test("word bridge rejects when ONLYOFFICE refuses SearchAndReplace", async () => {
  const harness = createWordBridgeHarness({
    text: "主任委员：姜涛",
    executeMethodAccepted: false,
  });

  await assert.rejects(
    harness.bridge.execute([{
      name: "word_replace_text",
      arguments: { search: "姜涛", replace: "敏感替换文本", matchCase: false },
    }]),
    error => {
      assert.equal(error.code, "EDITOR_COMMAND_REJECTED");
      assert.equal(error.details.editorType, "word");
      assert.equal(error.details.operation, "SearchAndReplace");
      assert.equal(error.details.reason, "execute_method_rejected");
      assert.equal(error.details.partialMutationPossible, false);
      assert.equal(error.details.requiredAction, "fix_or_use_supported_operation");
      assert.equal(error.details.reuseAllowed, false);
      assert.doesNotMatch(JSON.stringify(error.details), /敏感替换文本/);
      return true;
    },
  );
});

test("word bridge replaces the document body without ApiDocument.SetText", async () => {
  const harness = createWordBridgeHarness({ text: "旧正文" });
  assert.equal(typeof harness.document.SetText, "undefined");

  const result = await harness.bridge.execute([{
    name: "word_set_document_text",
    arguments: { text: "新正文" },
  }]);

  assert.equal(harness.text, "新正文");
  assert.deepEqual(Array.from(harness.paragraphs), ["新正文"]);
  assert.equal(result.changed, 1);
  assert.equal(result.needsSave, true);
  assert.equal(result.results[0].characters, 3);
  assert.equal(result.results[0].paragraphs, 1);
});

test("word bridge preserves mixed line endings, quotes, blank lines, and a trailing newline", async () => {
  const harness = createWordBridgeHarness({ text: "旧正文" });
  const source = "第一段 \"委员会\"\r\n\r\n第二段\r第三段\n";

  const result = await harness.bridge.execute([{
    name: "word_set_document_text",
    arguments: { text: source },
  }]);

  assert.equal(harness.text, "第一段 \"委员会\"\n\n第二段\n第三段\n");
  assert.deepEqual(
    Array.from(harness.paragraphs),
    ["第一段 \"委员会\"", "", "第二段", "第三段", ""],
  );
  assert.equal(result.results[0].characters, source.length);
  assert.equal(result.results[0].paragraphs, 5);
});

test("word bridge leaves one empty paragraph when replacing the body with an empty string", async () => {
  const harness = createWordBridgeHarness({ text: "旧正文" });

  const result = await harness.bridge.execute([{
    name: "word_set_document_text",
    arguments: { text: "" },
  }]);

  assert.equal(harness.text, "");
  assert.deepEqual(Array.from(harness.paragraphs), [""]);
  assert.equal(result.changed, 1);
  assert.equal(result.results[0].paragraphs, 1);
});

test("word bridge labels document command failures without exposing the document text", async () => {
  const harness = createWordBridgeHarness({
    text: "敏感正文",
    removeAllElementsAccepted: false,
  });

  await assert.rejects(
    harness.bridge.execute([{
      name: "word_set_document_text",
      arguments: { text: "新的敏感正文" },
    }]),
    error => {
      assert.match(error.message, /无法清空 Word 正文/);
      assert.equal(error.details.phase, "word-command");
      assert.doesNotMatch(error.message, /新的敏感正文/);
      return true;
    },
  );
});

test("word bridge inserts a proportional image at the current cursor", async () => {
  const harness = createWordBridgeHarness({ text: "正文" });
  const result = await harness.bridge.execute([{
    name: "word_add_image",
    arguments: {
      _image: {
        url: "https://app.test/api/v1/editor-relay/images/asset.png?token=signed",
        assetId: "asset.png",
        widthPx: 800,
        heightPx: 400,
      },
      widthMm: 60,
      wrapping: "inline",
      name: "封面图",
    },
  }]);

  assert.equal(result.changed, 1);
  assert.equal(result.needsSave, true);
  assert.equal(result.results[0].target, "current");
  assert.equal(result.results[0].widthMm, 60);
  assert.equal(result.results[0].heightMm, 30);
  assert.equal(harness.createdImages[0].width, 60 * 36000);
  assert.equal(harness.createdImages[0].height, 30 * 36000);
  assert.equal(harness.createdImages[0].name, "封面图");
  assert.equal(harness.paragraphObjects.at(-1).drawings[0], harness.createdImages[0]);
});

test("word bridge targets a paragraph by index or searched occurrence", async () => {
  const harness = createWordBridgeHarness({
    paragraphs: ["说明", "目标 A", "目标 B"],
  });
  const image = {
    url: "https://app.test/api/v1/editor-relay/images/asset.png?token=signed",
    assetId: "asset.png",
    widthPx: 300,
    heightPx: 600,
  };
  const result = await harness.bridge.execute([
    {
      name: "word_add_image",
      arguments: { _image: image, paragraphIndex: 1, heightMm: 40, wrapping: "square" },
    },
    {
      name: "word_add_image",
      arguments: { _image: image, search: "目标", occurrence: 2, matchCase: true },
    },
  ]);

  assert.equal(result.changed, 2);
  assert.equal(result.results[0].paragraphIndex, 1);
  assert.equal(result.results[0].widthMm, 20);
  assert.equal(result.results[0].heightMm, 40);
  assert.equal(result.results[1].paragraphIndex, 3);
  assert.equal(harness.paragraphObjects[0].drawings[0].wrapping, "square");
  assert.equal(harness.paragraphObjects[2].drawings.length, 1);
});

test("word bridge rejects conflicting or missing image targets without mutation", async () => {
  const harness = createWordBridgeHarness({ paragraphs: ["Alpha", "Beta"] });
  const image = {
    url: "https://app.test/api/v1/editor-relay/images/asset.png?token=signed",
    widthPx: 100,
    heightPx: 100,
  };

  await assert.rejects(
    harness.bridge.execute([{
      name: "word_add_image",
      arguments: { _image: image, current: true, paragraphIndex: 1 },
    }]),
    /最多只能指定一种/,
  );
  await assert.rejects(
    harness.bridge.execute([{
      name: "word_add_image",
      arguments: { _image: image, search: "Missing" },
    }]),
    /未找到指定的图片插入目标段落/,
  );

  assert.equal(harness.createdImages.length, 0);
  assert.ok(harness.paragraphObjects.every(paragraph => paragraph.drawings.length === 0));
});

test("word revision control keeps tracking and display mode independent in batch order", async () => {
  const harness = createWordBridgeHarness({ paragraphs: ["旧术语"] });
  const result = await harness.bridge.execute([
    {
      name: "word_manage_revisions",
      arguments: { action: "start", displayMode: "edit" },
    },
    {
      name: "word_set_paragraph_text",
      arguments: { paragraphIndex: 1, text: "新术语" },
    },
  ]);

  assert.equal(result.changed, 2);
  assert.equal(result.needsSave, true);
  assert.deepEqual(JSON.parse(JSON.stringify(result.results[0])), {
    name: "word_manage_revisions",
    action: "start",
    tracking: true,
    displayModeApplied: "edit",
  });
  assert.deepEqual(harness.revisionEvents, [
    { type: "tracking", value: true },
    { type: "display", value: "edit" },
  ]);
  assert.equal(harness.paragraphs[0], "新术语");
  assert.equal(harness.executeMethodCalls[0].name, "SetDisplayModeInReview");
  assert.deepEqual(harness.executeMethodCalls[0].args, ["edit"]);
  const tracking = await harness.bridge.execute([{
    name: "word_manage_revisions",
    arguments: { action: "setDisplay", displayMode: "simple" },
  }]);
  assert.equal(tracking.changed, 0);
  assert.equal(tracking.needsSave, false);
  assert.equal(tracking.results[0].tracking, true);
  assert.equal(tracking.results[0].displayModeApplied, "simple");
});

test("word revision display supports every mode and old start/stop calls stay compatible", async () => {
  const harness = createWordBridgeHarness();
  const started = await harness.bridge.execute([{
    name: "word_manage_revisions",
    arguments: { action: "start" },
  }]);
  assert.equal(started.results[0].tracking, true);
  assert.equal(started.results[0].displayModeApplied, undefined);

  for (const displayMode of ["edit", "simple", "final", "original"]) {
    const displayed = await harness.bridge.execute([{
      name: "word_manage_revisions",
      arguments: { action: "setDisplay", displayMode },
    }]);
    assert.equal(displayed.results[0].displayModeApplied, displayMode);
    assert.equal(displayed.results[0].tracking, true);
  }

  const stopped = await harness.bridge.execute([{
    name: "word_manage_revisions",
    arguments: { action: "stop" },
  }]);
  assert.equal(stopped.results[0].tracking, false);
  assert.equal(stopped.results[0].displayModeApplied, undefined);
});

test("word revision display validation and callback rejection stop following edits", async () => {
  const invalidHarness = createWordBridgeHarness({ paragraphs: ["未修改"] });
  await assert.rejects(
    invalidHarness.bridge.execute([{
      name: "word_manage_revisions",
      arguments: { action: "setDisplay" },
    }]),
    error => error.code === "INVALID_TOOL_ARGUMENTS"
      && error.details.partialMutationPossible === false,
  );
  await assert.rejects(
    invalidHarness.bridge.execute([{
      name: "word_manage_revisions",
      arguments: { action: "start", displayMode: "markup" },
    }]),
    error => error.code === "INVALID_TOOL_ARGUMENTS"
      && error.details.partialMutationPossible === false,
  );

  const rejectedHarness = createWordBridgeHarness({
    paragraphs: ["未修改"],
    executeMethodResponses: {
      SetDisplayModeInReview: { success: false, error: "display callback failed" },
    },
  });
  await assert.rejects(
    rejectedHarness.bridge.execute([
      {
        name: "word_manage_revisions",
        arguments: { action: "start", displayMode: "edit" },
      },
      {
        name: "word_set_paragraph_text",
        arguments: { paragraphIndex: 1, text: "不应修改" },
      },
    ]),
    error => /display callback failed/.test(error.message)
      && error.details.toolCallIndex === 0
      && error.details.partialMutationPossible === true,
  );
  assert.equal(rejectedHarness.paragraphs[0], "未修改");
});
