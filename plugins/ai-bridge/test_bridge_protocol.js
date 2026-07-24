const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");

const pluginSource = fs.readFileSync(path.join(__dirname, "plugin.js"), "utf8");
const hostSource = fs.readFileSync(path.join(__dirname, "host-bridge.js"), "utf8");
const clientSource = fs.readFileSync(path.join(__dirname, "client-sdk.js"), "utf8");
const wordBridgeSource = fs.readFileSync(path.join(__dirname, "bridges/word-bridge.js"), "utf8");
const publicContract = JSON.parse(fs.readFileSync(path.join(__dirname, "public-api.json"), "utf8"));
const docxCapabilityMatrix = fs.readFileSync(
  path.join(__dirname, "DOCX-CAPABILITIES.zh-CN.md"),
  "utf8",
);

function eventTarget(target) {
  const listeners = new Map();
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
  target.dispatchEvent = () => true;
  return target;
}

function fakeElement(tagName, id = "") {
  const attributes = new Map();
  return {
    tagName: String(tagName).toUpperCase(),
    id,
    className: "",
    dataset: {},
    children: [],
    parentElement: null,
    textContent: "",
    appendChild(child) {
      if (child.parentElement) {
        const previousIndex = child.parentElement.children.indexOf(child);
        if (previousIndex >= 0) child.parentElement.children.splice(previousIndex, 1);
      }
      this.children.push(child);
      child.parentElement = this;
      return child;
    },
    insertBefore(child, reference) {
      if (child.parentElement) {
        const previousIndex = child.parentElement.children.indexOf(child);
        if (previousIndex >= 0) child.parentElement.children.splice(previousIndex, 1);
      }
      const referenceIndex = this.children.indexOf(reference);
      assert.notEqual(referenceIndex, -1, "insertBefore reference must be a child");
      this.children.splice(referenceIndex, 0, child);
      child.parentElement = this;
      return child;
    },
    setAttribute(name, value) {
      attributes.set(String(name), String(value));
    },
    getAttribute(name) {
      return attributes.has(String(name)) ? attributes.get(String(name)) : null;
    },
  };
}

function createEditorUiHarness(options = {}) {
  const html = fakeElement("html");
  const head = fakeElement("head");
  const appTitle = fakeElement("section", "app-title");
  const logo = fakeElement("div", "header-logo");
  const save = fakeElement("div", "slot-btn-dt-save");
  const print = fakeElement("div", "slot-btn-dt-print");
  const undo = fakeElement("div", "slot-btn-dt-undo");
  const redo = fakeElement("div", "slot-btn-dt-redo");
  const quickAccessMenu = fakeElement("div", "slot-btn-dt-quick-access");
  const documentName = fakeElement("div", "id-box-doc-name");
  const userName = fakeElement("div");
  userName.setAttribute("data-layout-name", "header-user");
  const rightGroup = fakeElement("section", "box-right-btn-group");
  const editModeGroup = fakeElement("div");
  editModeGroup.setAttribute("data-layout-name", "header-editMode");
  const collaborators = fakeElement("div");
  collaborators.setAttribute("data-layout-name", "header-users");

  for (const element of [logo, save, print, undo, quickAccessMenu, documentName, userName]) {
    appTitle.appendChild(element);
  }
  if (!options.missingRedo) appTitle.insertBefore(redo, quickAccessMenu);
  rightGroup.appendChild(editModeGroup);
  rightGroup.appendChild(collaborators);

  const roots = [html, head, appTitle, rightGroup];
  const findById = id => {
    const pending = roots.slice();
    while (pending.length) {
      const element = pending.shift();
      if (element.id === id) return element;
      pending.push(...element.children);
    }
    return null;
  };
  const findByLayout = layoutName => {
    const pending = roots.slice();
    while (pending.length) {
      const element = pending.shift();
      if (element.getAttribute("data-layout-name") === layoutName) return element;
      pending.push(...element.children);
    }
    return null;
  };

  const resizeEvents = [];
  const editorDocument = {
    documentElement: html,
    head,
    defaultView: {
      Event: class Event {
        constructor(type) { this.type = type; }
      },
      requestAnimationFrame(callback) { callback(); },
      dispatchEvent(event) { resizeEvents.push(event.type); },
    },
    createElement(tagName) {
      return fakeElement(tagName);
    },
    querySelector(selector) {
      if (selector.startsWith("#")) return findById(selector.slice(1));
      if (selector === '[data-layout-name="header-editMode"]') {
        return findByLayout("header-editMode");
      }
      return null;
    },
    querySelectorAll() {
      return [];
    },
  };
  const frameListeners = new Map();
  const frame = {
    contentDocument: editorDocument,
    addEventListener(name, listener) {
      frameListeners.set(name, listener);
    },
  };
  const hostDocument = {
    documentElement: fakeElement("html"),
    currentScript: {
      src: "https://docs.test/sdkjs-plugins/ai-bridge/host-bridge.js",
    },
    querySelector(selector) {
      return selector === 'iframe[src*="/web-apps/apps/"]' ? frame : null;
    },
  };
  const observers = [];
  class FakeMutationObserver {
    constructor(callback) {
      this.callback = callback;
      observers.push(this);
    }
    observe() {}
    disconnect() {}
  }

  return {
    appTitle,
    collaborators,
    documentName,
    editModeGroup,
    editorDocument,
    frame,
    frameListeners,
    hostDocument,
    logo,
    observers,
    print,
    quickAccessMenu,
    redo,
    resizeEvents,
    rightGroup,
    save,
    undo,
    userName,
    MutationObserver: FakeMutationObserver,
  };
}

function createHarness(options = {}) {
  const editorType = options.editorType || "word";
  const editorConfig = {
    documentType: editorType,
    document: { key: "doc-key-v1", title: `demo.${editorType === "word" ? "docx" : editorType === "slide" ? "pptx" : "xlsx"}`, fileType: editorType === "word" ? "docx" : editorType === "slide" ? "pptx" : "xlsx" },
    editorConfig: {
      callbackUrl: "https://app.test/callback",
      lang: options.interfaceLanguage || "en",
      region: options.region || "en-US",
      user: { id: "user-1" },
    },
    token: options.editorToken || "editor-token",
  };
  const executedToolCalls = [];
  const initializationCommands = [];
  const servicePaths = [];
  const hostRelayPaths = [];
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
      hostname: options.hostname || "app.test",
      reload() { reloadCount += 1; },
    },
    crypto: { randomUUID: () => "00000000-0000-4000-8000-000000000001" },
    btoa: value => Buffer.from(value, "binary").toString("base64"),
    setTimeout,
    clearTimeout,
    history: { replaceState() {} },
    navigator: { sendBeacon() { return true; } },
    sessionStorage,
    localStorage,
    aiBridgeOptions: {
      getEditorConfig: () => editorConfig,
      clientOrigins: options.clientOrigins || [],
      httpRelay: options.httpRelay,
    },
  });
  if (typeof options.hostFetch === "function") {
    hostWindow.fetch = async (requestPath, requestOptions) => {
      hostRelayPaths.push(requestPath);
      return options.hostFetch(requestPath, requestOptions);
    };
  }

  const pluginWindow = eventTarget({
    location: {
      origin: "https://docs.test",
      href: "https://docs.test/sdkjs-plugins/ai-bridge/index.html",
      ancestorOrigins: ["https://app.test"],
    },
    setTimeout,
    clearTimeout,
    setInterval,
    clearInterval,
    fetch: async path => {
      servicePaths.push(path);
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

  const topProxy = {
    postMessage(message, targetOrigin) {
      assert.equal(targetOrigin, "https://app.test");
      queueMicrotask(() => hostWindow.dispatch("message", {
        origin: "https://docs.test",
        source: pluginHandle,
        data: message,
      }));
    },
  };

  const pluginHandle = {
    postMessage(message, targetOrigin) {
      assert.equal(targetOrigin, "https://docs.test");
      queueMicrotask(() => pluginWindow.dispatch("message", {
        origin: "https://app.test",
        source: topProxy,
        data: message,
      }));
    },
  };

  pluginWindow.top = topProxy;
  pluginWindow.AICopilotBridges = {
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

  const Asc = {
    scope: {},
    plugin: {
      info: { editorType },
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

  const hostDocument = options.hostDocument || {
    documentElement,
    querySelector() { return null; },
    currentScript: {
      src: "https://docs.test/sdkjs-plugins/ai-bridge/host-bridge.js",
    },
  };
  vm.runInNewContext(hostSource, {
    window: hostWindow,
    document: hostDocument,
    MutationObserver: options.MutationObserver || class MutationObserver {
      observe() {}
      disconnect() {}
    },
    URL,
    CustomEvent: class CustomEvent {
      constructor(name, options) { this.type = name; this.detail = options && options.detail; }
    },
    console,
  });

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
    hostRelayPaths,
    documentElement,
    sessionValues,
    localValues,
    pluginHandle,
    topProxy,
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
      path: `/copilot-api/images/${assetId}?token=signed-token`,
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
  let currentPage = Number(options.currentPage) || 0;
  const createdImages = [];
  const executeMethodCalls = [];
  const formattedRanges = [];
  const navigatedPages = [];
  const rangeActions = [];
  const createdStyles = [];
  const createdCharts = [];
  const appliedStyles = [];
  const fieldInstructions = [];
  const scope = {};
  const styles = new Map();

  function createStyle(name, type) {
    const style = {
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
    createdStyles.push(style);
    return style;
  }

  function syncDocumentText() {
    documentText = documentParagraphs.map(paragraph => paragraph.text).join("\n");
  }

  function createParagraph(text = "") {
    const paragraph = {
      text: String(text),
      drawings: [],
      runs: [],
      GetClassType() { return "paragraph"; },
      GetText() { return this.text; },
      AddText(value) {
        const run = {
          text: String(value),
          style: null,
          GetClassType() { return "run"; },
          SetStyle(style) { this.style = style; return true; },
          GetRange() {
            return {
              AddField(instruction) {
                fieldInstructions.push(String(instruction));
                return true;
              },
            };
          },
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
          this.runs.push({
            text: this.text,
            style: null,
            GetClassType() { return "run"; },
            SetStyle(style) { this.style = style; return true; },
          });
        }
        syncDocumentText();
        return this.runs[0] || {};
      },
      GetElementsCount() { return this.runs.length; },
      GetElement(index) { return this.runs[index] || null; },
      SetNumbering(level) { this.numbering = level; return true; },
      AddDrawing(drawing) {
        this.drawings.push(drawing);
        drawing.paragraph = this;
        return drawing;
      },
    };
    if (paragraph.text) {
      paragraph.runs.push({
        text: paragraph.text,
        style: null,
        GetClassType() { return "run"; },
        SetStyle(style) { this.style = style; return true; },
      });
    }
    return paragraph;
  }

  documentParagraphs = (
    Array.isArray(options.paragraphs)
      ? options.paragraphs
      : documentText.split("\n")
  ).map(createParagraph);
  if (!documentParagraphs.length) documentParagraphs = [createParagraph()];
  syncDocumentText();
  const document = {
    GetText() { return documentText; },
    GetElement(index) { return documentParagraphs[index] || null; },
    GetAllParagraphs() { return documentParagraphs; },
    GetAllTables() { return Array.isArray(options.tables) ? options.tables : []; },
    GetAllContentControls() {
      return Array.isArray(options.contentControls) ? options.contentControls : [];
    },
    GetContentControlsByTag(tag) {
      return this.GetAllContentControls().filter(control => (
        typeof control.GetTag === "function" && String(control.GetTag()) === String(tag)
      ));
    },
    GetCurrentContentControl() {
      return options.currentContentControl || null;
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
    RemoveAllElements() {
      if (options.removeAllElementsAccepted === false) return false;
      documentParagraphs = [createParagraph()];
      syncDocumentText();
      return true;
    },
    Push(paragraph) {
      if (options.pushAccepted === false) return false;
      documentParagraphs.push(paragraph);
      syncDocumentText();
      return true;
    },
    InsertContent(paragraphs) {
      if (options.insertContentAccepted === false) return false;
      documentParagraphs.push(...paragraphs);
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
            return { GetId() { return `comment-${format.start}`; } };
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
  };
  const officeContext = vm.createContext({
    Asc: { scope },
    Api: {
      GetDocument: () => document,
      CreateParagraph: () => createParagraph(),
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
        try {
          callback(vm.runInContext(`(${command.toString()})()`, officeContext));
        } catch (error) {
          callback(JSON.stringify({ ok: false, error: error.message }));
        }
      },
      executeMethod(name, args, callback) {
        executeMethodCalls.push({ name, args: JSON.parse(JSON.stringify(args)) });
        if (options.executeMethodAccepted === false) return false;
        if (name === "SearchAndReplace") documentText = replaceLiteral(documentText, args[0]);
        const responses = options.executeMethodResponses || {};
        queueMicrotask(() => callback(Object.prototype.hasOwnProperty.call(responses, name) ? responses[name] : undefined));
        return true;
      },
    },
  };
  pluginWindow.Asc = Asc;
  vm.runInNewContext(wordBridgeSource, { window: pluginWindow, Asc, console, Promise, JSON, String });
  return {
    bridge: pluginWindow.AICopilotBridges.word,
    document,
    executeMethodCalls,
    formattedRanges,
    navigatedPages,
    rangeActions,
    createdImages,
    createdCharts,
    createdStyles,
    appliedStyles,
    fieldInstructions,
    get text() { return documentText; },
    get paragraphs() { return documentParagraphs.map(paragraph => paragraph.text); },
    get paragraphObjects() { return documentParagraphs; },
  };
}

test("same-origin editor header moves the native quick access slots before edit mode", () => {
  const ui = createEditorUiHarness();
  createHarness({
    hostDocument: ui.hostDocument,
    MutationObserver: ui.MutationObserver,
  });

  const quickAccess = ui.editorDocument.querySelector("#ai-bridge-quick-access");
  assert.ok(quickAccess);
  assert.deepEqual(quickAccess.children, [ui.save, ui.undo, ui.redo]);
  assert.deepEqual(ui.rightGroup.children, [
    quickAccess,
    ui.editModeGroup,
    ui.collaborators,
  ]);
  assert.equal(quickAccess.getAttribute("role"), "menubar");
  assert.equal(quickAccess.getAttribute("aria-label"), "Quick access toolbar");
  assert.equal(ui.editorDocument.documentElement.dataset.aiBridgeCompactHeader, "true");
  assert.equal(ui.resizeEvents.length, 1);
  const uiStyle = ui.editorDocument.head.children.find(
    child => child.id === "ai-bridge-editor-ui-customization",
  );
  assert.ok(uiStyle);
  assert.match(uiStyle.textContent, /#btn-go-back,/);
  assert.match(uiStyle.textContent, /#id-btn-favorite,/);

  assert.equal(ui.print.parentElement, ui.appTitle);
  assert.equal(ui.quickAccessMenu.parentElement, ui.appTitle);
  assert.equal(ui.logo.parentElement, ui.appTitle);
  assert.equal(ui.documentName.parentElement, ui.appTitle);
  assert.equal(ui.userName.parentElement, ui.appTitle);
  assert.equal(ui.collaborators.parentElement, ui.rightGroup);
});

test("same-origin editor header customization is idempotent", () => {
  const ui = createEditorUiHarness();
  createHarness({
    hostDocument: ui.hostDocument,
    MutationObserver: ui.MutationObserver,
  });
  const quickAccess = ui.editorDocument.querySelector("#ai-bridge-quick-access");
  assert.equal(ui.observers.length, 1);

  ui.observers[0].callback([]);
  ui.observers[0].callback([]);

  assert.equal(ui.editorDocument.querySelector("#ai-bridge-quick-access"), quickAccess);
  assert.deepEqual(quickAccess.children, [ui.save, ui.undo, ui.redo]);
  assert.equal(ui.rightGroup.children.filter(child => child === quickAccess).length, 1);
  assert.equal(ui.resizeEvents.length, 1);
});

test("same-origin editor keeps the original title when a required slot is missing", () => {
  const ui = createEditorUiHarness({ missingRedo: true });
  createHarness({
    hostDocument: ui.hostDocument,
    MutationObserver: ui.MutationObserver,
  });

  assert.equal(ui.editorDocument.querySelector("#ai-bridge-quick-access"), null);
  assert.equal(ui.editorDocument.documentElement.dataset.aiBridgeCompactHeader, undefined);
  assert.equal(ui.save.parentElement, ui.appTitle);
  assert.equal(ui.undo.parentElement, ui.appTitle);
  assert.deepEqual(ui.rightGroup.children, [ui.editModeGroup, ui.collaborators]);
});

test("headless plugin handshakes with one host instance and executes a read-only tool", async () => {
  const { hostWindow } = createHarness();
  await hostWindow.aiBridge.ready({ timeoutMs: 1000 });

  assert.equal(hostWindow.aiBridge.version, "0.4.0");
  assert.equal(hostWindow.aiBridge.editorType, "word");
  assert.equal(hostWindow.aiBridge.context.documentKey, "doc-key-v1");
  assert.ok(hostWindow.aiBridge.capabilities.tools.includes("word_inspect"));

  const result = await hostWindow.aiBridge.word.inspect({}, { timeoutMs: 1000 });
  assert.equal(result.changed, 0);
  assert.equal(result.results[0].name, "word_inspect");
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

test("superseded HTTP Relay stops without re-registering or reloading", async () => {
  let registerCalls = 0;
  let pollCalls = 0;
  const harness = createHarness({
    hostname: "localhost",
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

test("first Relay credential failure schedules only one page reload", async () => {
  const harness = createHarness({
    hostname: "localhost",
    hostFetch: async requestPath => {
      assert.ok(requestPath.endsWith("/register"));
      return relayResponse(401, {
        ok: false,
        error: {
          code: "EDITOR_TOKEN_EXPIRED",
          message: "expired",
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

test("repeated Relay credential failure inside the guard window becomes terminal", async () => {
  const harness = createHarness({
    hostname: "localhost",
    localValues: {
      "aiBridgeCredentialReloadAt:demo.docx|docx|word|user-1": String(Date.now()),
    },
    hostFetch: async requestPath => {
      assert.ok(requestPath.endsWith("/register"));
      return relayResponse(401, {
        ok: false,
        error: {
          code: "EDITOR_TOKEN_EXPIRED",
          message: "expired again",
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
        assert.equal(requestPath, "/copilot-api/images/import");
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
    "/copilot-api/forcesave",
    "/copilot-api/history",
    "/copilot-api/undo",
    "/copilot-api/redo",
  ]);
});

test("reusing a completed requestId returns the cached response without applying an edit twice", async () => {
  const harness = createHarness();
  await harness.hostWindow.aiBridge.ready({ timeoutMs: 1000 });
  const options = { timeoutMs: 1000, requestId: "stable-agent-request-1" };

  const first = await harness.hostWindow.aiBridge.word.appendParagraph({ text: "once" }, options);
  const second = await harness.hostWindow.aiBridge.word.appendParagraph({ text: "once" }, options);

  assert.equal(first.results[0].name, "word_append_paragraph");
  assert.equal(second.results[0].name, "word_append_paragraph");
  assert.equal(harness.executedToolCalls.length, 1);
});

test("host imports an image before sending the internal source to the plugin", async () => {
  let imports = 0;
  const harness = createHarness({
    hostFetch: async (requestPath, requestOptions) => {
      imports += 1;
      assert.equal(requestPath, "/copilot-api/images/import");
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
  assert.equal(args._image.url, `https://app.test/copilot-api/images/${"a".repeat(64)}.png?token=signed-token`);
  assert.equal(args._image.widthPx, 400);
  assert.equal(args._image.heightPx, 200);
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
      if (requestPath === "/copilot-api/images/import") return importedImageResponse({
        widthPx: 1,
        heightPx: 1,
      });
      if (requestPath.startsWith("/copilot-api/images/")) {
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
      assert.doesNotMatch(error.message, /敏感正文/);
      return true;
    },
  );
});

test("force-save persists without reloading the live editor", async () => {
  const harness = createHarness({
    serviceResponse(path) {
      if (path === "/copilot-api/forcesave") {
        return {
          persisted: true,
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
  assert.equal(harness.reloadCount, 0);
  await new Promise(resolve => setTimeout(resolve, 80));
  assert.equal(harness.reloadCount, 0);
});

test("explicit save persists without reloading the live editor", async () => {
  const harness = createHarness({
    serviceResponse(path) {
      if (path === "/copilot-api/forcesave") {
        return {
          persisted: true,
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
  assert.equal(harness.reloadCount, 0);
  await new Promise(resolve => setTimeout(resolve, 80));
  assert.equal(harness.reloadCount, 0);
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

test("word bridge uses ONLYOFFICE zero-based placeholders for multilevel numbering", async () => {
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
  assert.deepEqual(customTypes.map(entry => entry.text), ["%0.", "%0.%1."]);
  assert.ok(harness.paragraphObjects.every(paragraph => paragraph.numbering === levels[1]));
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
    arguments: { kind: "onlyoffice" },
  }]);
  const updated = await harness.bridge.execute([{
    name: "word_set_macros",
    arguments: { macros: macros.macrosArray, current: 0 },
  }]);

  assert.deepEqual(inspected.results[0].content, macros);
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
      arguments: { search: "姜涛", replace: "张三", matchCase: false },
    }]),
    /ONLYOFFICE 拒绝执行 SearchAndReplace/,
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
        url: "https://app.test/copilot-api/images/asset.png?token=signed",
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
    url: "https://app.test/copilot-api/images/asset.png?token=signed",
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
    url: "https://app.test/copilot-api/images/asset.png?token=signed",
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
