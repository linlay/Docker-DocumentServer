const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");

const pluginSource = fs.readFileSync(path.join(__dirname, "plugin.js"), "utf8");
const hostSource = fs.readFileSync(path.join(__dirname, "host-bridge.js"), "utf8");
const clientSource = fs.readFileSync(path.join(__dirname, "client-sdk.js"), "utf8");
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

function createHarness(options = {}) {
  const editorType = options.editorType || "word";
  const editorConfig = {
    documentType: editorType,
    document: { key: "doc-key-v1", title: `demo.${editorType === "word" ? "docx" : editorType === "slide" ? "pptx" : "xlsx"}`, fileType: editorType === "word" ? "docx" : editorType === "slide" ? "pptx" : "xlsx" },
    editorConfig: { callbackUrl: "https://app.test/callback", user: { id: "user-1" } },
  };
  const executedToolCalls = [];
  const servicePaths = [];
  let reloadCount = 0;

  const hostWindow = eventTarget({
    location: {
      origin: "https://app.test",
      href: "https://app.test/editor",
      reload() { reloadCount += 1; },
    },
    crypto: { randomUUID: () => "00000000-0000-4000-8000-000000000001" },
    setTimeout,
    clearTimeout,
    history: { replaceState() {} },
    aiBridgeOptions: {
      getEditorConfig: () => editorConfig,
      clientOrigins: options.clientOrigins || [],
    },
  });

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
      return { ok: true, status: 200, json: async () => ({ persisted: true, path }) };
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
          changed: toolCalls.filter(call => !call.name.endsWith("_inspect")).length,
          results: toolCalls.map(call => ({ name: call.name, document: "demo" })),
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

  vm.runInNewContext(hostSource, {
    window: hostWindow,
    document: {
      documentElement: { dataset: {} },
      querySelector() { return null; },
      currentScript: {
        src: "https://docs.test/sdkjs-plugins/ai-bridge/host-bridge.js",
      },
    },
    MutationObserver: class MutationObserver {
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
    pluginHandle,
    topProxy,
    get reloadCount() { return reloadCount; },
  };
}

test("headless plugin handshakes with one host instance and executes a read-only tool", async () => {
  const { hostWindow } = createHarness();
  await hostWindow.aiBridge.ready({ timeoutMs: 1000 });

  assert.equal(hostWindow.aiBridge.version, "0.1.0");
  assert.equal(hostWindow.aiBridge.editorType, "word");
  assert.equal(hostWindow.aiBridge.context.documentKey, "doc-key-v1");
  assert.ok(hostWindow.aiBridge.capabilities.tools.includes("word_inspect"));

  const result = await hostWindow.aiBridge.word.inspect({}, { timeoutMs: 1000 });
  assert.equal(result.changed, 0);
  assert.equal(result.results[0].name, "word_inspect");
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

test("public contract, plugin allow-lists, and all 27 convenience methods stay aligned", async () => {
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
        scaleFont: ["word_scale_font", { scale: 0.9 }],
        addTable: ["word_add_table", { rows: 1, cols: 1 }],
        setDocumentText: ["word_set_document_text", { text: "a" }],
      },
    },
    slide: {
      group: "slides",
      methods: {
        inspect: ["slides_inspect", {}],
        replaceText: ["slides_replace_text", { search: "a", replace: "b" }],
        scaleFont: ["slides_scale_font", { scale: 0.9 }],
        formatText: ["slides_format_text", { bold: true }],
        formatSelection: ["slides_format_selection", { italic: true }],
        addSlide: ["slides_add_slide", { title: "a" }],
        duplicateSlide: ["slides_duplicate_slide", { slide: 1 }],
        deleteSlide: ["slides_delete_slide", { slide: 1 }],
        addTextBox: ["slides_add_textbox", { slide: 1, text: "a" }],
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
