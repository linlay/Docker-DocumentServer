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
    editorConfig: { callbackUrl: "https://app.test/callback", user: { id: "user-1" } },
    token: options.editorToken || "editor-token",
  };
  const executedToolCalls = [];
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
    setTimeout,
    clearTimeout,
    history: { replaceState() {} },
    navigator: { sendBeacon() { return true; } },
    sessionStorage,
    localStorage,
    aiBridgeOptions: {
      getEditorConfig: () => editorConfig,
      clientOrigins: options.clientOrigins || [],
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
    plugin: {
      info: { editorType },
      callCommand(command, close, calc, callback) { callback(JSON.stringify({ ok: true })); },
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
    console,
  });

  Asc.plugin.init();
  return {
    hostWindow,
    editorConfig,
    executedToolCalls,
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
  let currentPage = Number(options.currentPage) || 0;
  const executeMethodCalls = [];
  const formattedRanges = [];
  const navigatedPages = [];
  const rangeActions = [];
  const scope = {};
  const document = {
    GetText() { return documentText; },
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
    GoToPage(index) {
      currentPage = index;
      navigatedPages.push(index);
      return true;
    },
  };
  const officeContext = vm.createContext({
    Asc: { scope },
    Api: { GetDocument: () => document, Color: value => value },
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
        documentText = replaceLiteral(documentText, args[0]);
        queueMicrotask(() => callback(undefined));
        return true;
      },
    },
  };
  vm.runInNewContext(wordBridgeSource, { window: pluginWindow, Asc, console, Promise, JSON, String });
  return {
    bridge: pluginWindow.AICopilotBridges.word,
    document,
    executeMethodCalls,
    formattedRanges,
    navigatedPages,
    rangeActions,
    get text() { return documentText; },
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

  assert.equal(hostWindow.aiBridge.version, "0.2.0");
  assert.equal(hostWindow.aiBridge.editorType, "word");
  assert.equal(hostWindow.aiBridge.context.documentKey, "doc-key-v1");
  assert.ok(hostWindow.aiBridge.capabilities.tools.includes("word_inspect"));

  const result = await hostWindow.aiBridge.word.inspect({}, { timeoutMs: 1000 });
  assert.equal(result.changed, 0);
  assert.equal(result.results[0].name, "word_inspect");
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
      },
    },
  };

  for (const [editorType, entry] of Object.entries(cases)) {
    const harness = createHarness({ editorType });
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

test("persisting a new storage version reloads the page after returning the tool result", async () => {
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
  assert.equal(harness.reloadCount, 1);
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
    [
      "addBookmark", "addComment", "addHyperlink", "addTable", "appendParagraph",
      "deleteMatches", "deleteParagraphs", "editTable", "formatDocument",
      "formatMatches", "formatParagraphs", "formatSelection", "formatTable",
      "insertPageBreak", "insertParagraph", "inspect", "navigate", "replaceText",
      "scaleFont", "scroll", "setDocumentText", "setHeaderFooter", "setList",
      "setPageLayout", "setParagraphText", "setTableCell",
    ].sort(),
  );
  assert.equal(
    Object.keys(office.word).length +
    Object.keys(office.slides).length +
    Object.keys(office.sheets).length,
    56,
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
