(function () {
  "use strict";

  const VERSION = "0.1.0";
  const PROTOCOL_VERSION = 1;
  const PLUGIN_GUID = "asc.{A17E5F31-64AA-4E37-9A42-8D430814C2F6}";
  const REQUEST_ID_PATTERN = /^[A-Za-z0-9._:-]{1,200}$/;
  const EXTERNAL_CLIENT_SOURCE = "ai-bridge-client";
  const EXTERNAL_RELAY_SOURCE = "ai-bridge-relay";
  const currentScript = document.currentScript;
  const pluginOrigin = currentScript
    ? new URL(currentScript.src, window.location.href).origin
    : window.location.origin;
  let sequence = 0;

  if (window.aiBridge && window.aiBridge.version === VERSION) return;
  document.documentElement.dataset.aiBridgeState = "loading";

  const pending = new Map();
  const readyWaiters = new Set();
  const listeners = new Map();
  const relayClients = new Map();
  const sessionId = createRequestId("session");
  let pluginWindow = null;
  let ready = false;
  let editorType = null;
  let bridgeCapabilities = null;

  const editorUiStyleId = "ai-bridge-editor-ui-customization";
  const editorUiStyle = [
    "#left-menu,",
    "#view-left-menu,",
    'li[data-layout-name="toolbar-collaboration"],',
    'li[data-layout-name="toolbar-plugins"]',
    "{ display: none !important; }",
    "#left-menu {",
    "  width: 0 !important;",
    "  min-width: 0 !important;",
    "  max-width: 0 !important;",
    "  flex: 0 0 0 !important;",
    "}",
  ].join("\n");

  function hideDynamicEditorUi(editorDocument) {
    const aiLinks = editorDocument.querySelectorAll('li.ribtab > a[data-title="AI"]');
    for (const link of aiLinks) {
      if (link.parentElement) link.parentElement.style.setProperty("display", "none", "important");
    }
  }

  function customizeEditorFrame(frame) {
    try {
      const editorDocument = frame.contentDocument;
      if (!editorDocument || !editorDocument.documentElement) return;
      if (editorDocument.documentElement.dataset.aiBridgeUiCustomized === "true") return;
      editorDocument.documentElement.dataset.aiBridgeUiCustomized = "true";

      const style = editorDocument.createElement("style");
      style.id = editorUiStyleId;
      style.textContent = editorUiStyle;
      (editorDocument.head || editorDocument.documentElement).appendChild(style);

      hideDynamicEditorUi(editorDocument);
      const editorWindow = editorDocument.defaultView;
      if (editorWindow) {
        editorWindow.requestAnimationFrame(function () {
          editorWindow.dispatchEvent(new editorWindow.Event("resize"));
        });
      }
      const observer = new MutationObserver(function () { hideDynamicEditorUi(editorDocument); });
      observer.observe(editorDocument.documentElement, { childList: true, subtree: true });
    } catch (error) {
      // The bundled example is same-origin. Cross-origin integrations use the
      // supported customization.layout settings injected into editorConfig.
    }
  }

  function installEditorUiCustomization() {
    const frame = document.querySelector('iframe[src*="/web-apps/apps/"]');
    if (!frame) return false;
    frame.addEventListener("load", function () { customizeEditorFrame(frame); });
    customizeEditorFrame(frame);
    return true;
  }

  if (!installEditorUiCustomization()) {
    const frameObserver = new MutationObserver(function () {
      if (installEditorUiCustomization()) frameObserver.disconnect();
    });
    frameObserver.observe(document.documentElement, { childList: true, subtree: true });
  }

  class AiBridgeError extends Error {
    constructor(code, message, options) {
      super(message || "ai-bridge 调用失败");
      this.name = "AiBridgeError";
      this.code = code || "AI_BRIDGE_ERROR";
      this.requestId = options && options.requestId || null;
      this.details = options && options.details;
    }
  }

  function bridgeError(code, message, options) {
    return new AiBridgeError(code, message, options);
  }

  function wireError(error, requestId) {
    if (error && typeof error === "object" && !(error instanceof Error)) {
      return bridgeError(error.code, error.message, {
        requestId: error.requestId || requestId,
        details: error.details,
      });
    }
    if (error instanceof AiBridgeError) return error;
    return bridgeError("EXECUTION_FAILED", error && error.message ? error.message : String(error), { requestId });
  }

  function serializedError(error, requestId) {
    const normalized = wireError(error, requestId);
    return {
      code: normalized.code,
      message: normalized.message,
      requestId: normalized.requestId || requestId || null,
      details: normalized.details,
    };
  }

  function emit(name, detail) {
    for (const listener of listeners.get(name) || []) {
      try { listener(detail); } catch (error) { window.setTimeout(function () { throw error; }, 0); }
    }
    window.dispatchEvent(new CustomEvent(`ai-bridge-${name}`, { detail }));
  }

  function createRequestId(prefix) {
    if (window.crypto && typeof window.crypto.randomUUID === "function") {
      return `${prefix || "request"}:${window.crypto.randomUUID()}`;
    }
    sequence += 1;
    return `${prefix || "request"}:${Date.now()}:${sequence}`;
  }

  function editorConfiguration() {
    if (window.aiBridgeOptions && typeof window.aiBridgeOptions.getEditorConfig === "function") {
      return window.aiBridgeOptions.getEditorConfig() || {};
    }
    return typeof config === "object" && config ? config : {};
  }

  function editorToken() {
    const editorConfig = editorConfiguration();
    return typeof editorConfig.token === "string" ? editorConfig.token : "";
  }

  function currentConfig() {
    const editorConfig = editorConfiguration();
    return {
      documentKey: editorConfig.document && editorConfig.document.key || "",
      fileName: editorConfig.document && editorConfig.document.title || "",
      fileType: editorConfig.document && editorConfig.document.fileType || "",
      editorType: editorConfig.documentType || "",
      callbackUrl: editorConfig.editorConfig && editorConfig.editorConfig.callbackUrl || "",
      userId: editorConfig.editorConfig && editorConfig.editorConfig.user && editorConfig.editorConfig.user.id || "",
    };
  }

  function publicContext() {
    const context = currentConfig();
    return {
      documentKey: context.documentKey,
      fileName: context.fileName,
      fileType: context.fileType,
      editorType: context.editorType,
      userId: context.userId,
    };
  }

  function stateSnapshot() {
    return {
      version: VERSION,
      protocolVersion: PROTOCOL_VERSION,
      pluginGuid: PLUGIN_GUID,
      ready,
      editorType,
      context: publicContext(),
      capabilities: bridgeCapabilities,
    };
  }

  function basePayload(message) {
    return {
      source: "ai-bridge-host",
      protocolVersion: PROTOCOL_VERSION,
      pluginGuid: PLUGIN_GUID,
      sessionId,
      ...message,
    };
  }

  function send(message) {
    if (!pluginWindow) return false;
    pluginWindow.postMessage(basePayload(message), pluginOrigin);
    return true;
  }

  function sendConfiguration() {
    send({ type: "configure", config: currentConfig() });
  }

  function waitUntilReady(timeoutMs) {
    if (ready) return Promise.resolve();
    return new Promise(function (resolve, reject) {
      const waiter = { resolve, reject };
      const timeout = window.setTimeout(function () {
        readyWaiters.delete(waiter);
        reject(bridgeError("NOT_READY", "ai-bridge 插件尚未就绪"));
      }, timeoutMs);
      waiter.resolve = function () {
        window.clearTimeout(timeout);
        resolve();
      };
      readyWaiters.add(waiter);
      send({ type: "host-ready" });
    });
  }

  function markReady(message) {
    ready = true;
    editorType = message.editorType || editorType;
    bridgeCapabilities = message.capabilities || bridgeCapabilities;
    document.documentElement.dataset.aiBridgeState = "ready";
    document.documentElement.dataset.aiBridgeVersion = VERSION;
    for (const waiter of readyWaiters) waiter.resolve();
    readyWaiters.clear();
    emit("ready", { editorType, capabilities: bridgeCapabilities, context: publicContext() });
  }

  function handlePluginMessage(event) {
    if (event.origin !== pluginOrigin) return;
    const message = event.data;
    if (!message || message.source !== "ai-bridge-plugin") return;
    if (message.pluginGuid !== PLUGIN_GUID || message.protocolVersion !== PROTOCOL_VERSION) return;

    if (message.type === "hello") {
      if (pluginWindow && pluginWindow !== event.source) return;
      pluginWindow = event.source;
      sendConfiguration();
      return;
    }
    if (event.source !== pluginWindow || message.sessionId !== sessionId) return;
    if (message.type === "ready") {
      markReady(message);
      return;
    }
    if (message.type === "reload") {
      emit("reload", { reason: message.reason || "unknown" });
      notifyRelayClients("reload", { reason: message.reason || "unknown" });
      window.location.reload();
      return;
    }
    if (message.type !== "result" || !message.requestId) return;

    const request = pending.get(message.requestId);
    if (!request) return;
    pending.delete(message.requestId);
    window.clearTimeout(request.timeout);
    if (message.error) {
      const error = wireError(message.error, message.requestId);
      emit("error", { error, requestId: message.requestId });
      request.reject(error);
    }
    else request.resolve(message.result);
  }

  window.addEventListener("message", handlePluginMessage);

  async function request(type, payload, options) {
    const timeoutMs = normalizeTimeout(options && options.timeoutMs, 90000);
    await waitUntilReady(timeoutMs);
    const requestId = options && options.requestId || createRequestId("command");
    if (typeof requestId !== "string" || !REQUEST_ID_PATTERN.test(requestId)) {
      throw bridgeError("INVALID_REQUEST_ID", "requestId 只能包含字母、数字、点、下划线、冒号或连字符，且长度不超过 200");
    }
    if (pending.has(requestId)) {
      throw bridgeError("REQUEST_IN_FLIGHT", `requestId 正在执行：${requestId}`, { requestId });
    }
    const context = currentConfig();
    if (!context.documentKey) throw bridgeError("DOCUMENT_NOT_CONFIGURED", "外部页面没有提供 document.key");
    // Refresh mutable metadata (notably document.key after async host setup)
    // before every command. postMessage ordering keeps configure ahead of it.
    sendConfiguration();

    return new Promise(function (resolve, reject) {
      const timeout = window.setTimeout(function () {
        pending.delete(requestId);
        const error = bridgeError("TIMEOUT", "ai-bridge 命令执行超时", { requestId });
        emit("error", { error, requestId });
        reject(error);
      }, timeoutMs);
      pending.set(requestId, { resolve, reject, timeout });
      send({
        type,
        requestId,
        target: { documentKey: context.documentKey, editorType: context.editorType },
        ...payload,
      });
    });
  }

  function execute(command, options) {
    return request("execute", { command }, options);
  }

  function executeTool(name, args, options) {
    return execute({ toolCalls: [{ name, arguments: args || {} }] }, options);
  }

  function control(action, options) {
    return request("control", { action }, options);
  }

  function normalizeTimeout(value, fallback) {
    const timeout = Number(value === undefined ? fallback : value);
    if (!Number.isFinite(timeout)) return fallback;
    return Math.min(300000, Math.max(100, timeout));
  }

  function configuredClientOrigins() {
    const values = window.aiBridgeOptions && window.aiBridgeOptions.clientOrigins;
    if (!Array.isArray(values)) return [];
    return values.map(function (origin) {
      if (typeof origin !== "string" || !origin || origin === "*") return null;
      try {
        const parsed = new URL(origin);
        return /^https?:$/.test(parsed.protocol) ? parsed.origin : null;
      } catch (error) {
        return null;
      }
    }).filter(Boolean);
  }

  function relayPost(client, message) {
    client.window.postMessage({
      source: EXTERNAL_RELAY_SOURCE,
      protocolVersion: PROTOCOL_VERSION,
      clientId: client.clientId,
      ...message,
    }, client.origin);
  }

  function stableHash(value) {
    let first = 0xdeadbeef;
    let second = 0x41c6ce57;
    const text = String(value);
    for (let index = 0; index < text.length; index += 1) {
      const code = text.charCodeAt(index);
      first = Math.imul(first ^ code, 2654435761);
      second = Math.imul(second ^ code, 1597334677);
    }
    first = Math.imul(first ^ (first >>> 16), 2246822507) ^ Math.imul(second ^ (second >>> 13), 3266489909);
    second = Math.imul(second ^ (second >>> 16), 2246822507) ^ Math.imul(first ^ (first >>> 13), 3266489909);
    return `${(second >>> 0).toString(36)}${(first >>> 0).toString(36)}`;
  }

  function relayRequestId(clientId, requestId) {
    return `relay:${stableHash(clientId)}:${stableHash(requestId)}`;
  }

  function notifyRelayClients(eventName, detail) {
    for (const client of relayClients.values()) {
      relayPost(client, { type: "event", event: eventName, detail });
    }
  }

  async function executeRelayMethod(method, params, requestId) {
    const requestOptions = { ...(params.options || {}), requestId };
    switch (method) {
      case "execute": return execute(params.command, requestOptions);
      case "executeTool": return executeTool(params.name, params.arguments || {}, requestOptions);
      case "executeBatch": return execute({ toolCalls: params.toolCalls }, requestOptions);
      case "save": return control("save", requestOptions);
      case "history": return control("history", requestOptions);
      case "undo": return control("undo", requestOptions);
      case "redo": return control("redo", requestOptions);
      case "getState": return stateSnapshot();
      default: throw bridgeError("METHOD_NOT_ALLOWED", `不允许外部方法：${method || "unknown"}`, { requestId });
    }
  }

  function handleExternalMessage(event) {
    const message = event.data;
    if (!message || message.source !== EXTERNAL_CLIENT_SOURCE) return;
    if (!configuredClientOrigins().includes(event.origin)) return;
    if (message.protocolVersion !== PROTOCOL_VERSION) return;
    if (typeof message.clientId !== "string" || !REQUEST_ID_PATTERN.test(message.clientId)) return;

    if (message.type === "connect") {
      const client = { clientId: message.clientId, origin: event.origin, window: event.source };
      relayClients.set(message.clientId, client);
      waitUntilReady(normalizeTimeout(message.timeoutMs, 30000)).then(function () {
        if (relayClients.get(message.clientId) !== client) return;
        relayPost(client, { type: "connected", state: stateSnapshot() });
      }, function (error) {
        relayPost(client, { type: "connection-error", error: serializedError(error) });
      });
      return;
    }

    const client = relayClients.get(message.clientId);
    if (!client || client.origin !== event.origin || client.window !== event.source) return;
    if (message.type === "disconnect") {
      relayClients.delete(message.clientId);
      return;
    }
    if (message.type !== "request" || typeof message.requestId !== "string" || !REQUEST_ID_PATTERN.test(message.requestId)) return;

    const params = message.params && typeof message.params === "object" ? message.params : {};
    const internalRequestId = relayRequestId(message.clientId, message.requestId);
    executeRelayMethod(message.method, params, internalRequestId).then(function (result) {
      relayPost(client, { type: "result", requestId: message.requestId, result });
    }, function (error) {
      const payload = serializedError(error, message.requestId);
      payload.requestId = message.requestId;
      relayPost(client, { type: "result", requestId: message.requestId, error: payload });
    });
  }

  window.addEventListener("message", handleExternalMessage);

  const api = {
    version: VERSION,
    protocolVersion: PROTOCOL_VERSION,
    pluginGuid: PLUGIN_GUID,
    get isReady() { return ready; },
    get editorType() { return editorType; },
    get context() { return currentConfig(); },
    get capabilities() { return bridgeCapabilities; },
    AiBridgeError,
    getState: stateSnapshot,
    on: function (name, listener) {
      if (typeof listener !== "function") throw bridgeError("INVALID_LISTENER", "listener 必须是函数");
      const entries = listeners.get(name) || new Set();
      entries.add(listener);
      listeners.set(name, entries);
      return function () { entries.delete(listener); };
    },
    off: function (name, listener) {
      const entries = listeners.get(name);
      if (entries) entries.delete(listener);
    },
    ready: function (options) { return waitUntilReady(options && options.timeoutMs || 30000); },
    execute,
    executeTool,
    executeBatch: function (toolCalls, options) { return execute({ toolCalls }, options); },
    save: function (options) { return control("save", options); },
    history: function (options) { return control("history", options); },
    undo: function (options) { return control("undo", options); },
    redo: function (options) { return control("redo", options); },
    word: {
      inspect: function (args, options) { return executeTool("word_inspect", args, options); },
      replaceText: function (args, options) { return executeTool("word_replace_text", args, options); },
      appendParagraph: function (args, options) { return executeTool("word_append_paragraph", args, options); },
      insertParagraph: function (args, options) { return executeTool("word_insert_paragraph", args, options); },
      formatDocument: function (args, options) { return executeTool("word_format_document", args, options); },
      formatSelection: function (args, options) { return executeTool("word_format_selection", args, options); },
      formatMatches: function (args, options) { return executeTool("word_format_matches", args, options); },
      deleteMatches: function (args, options) { return executeTool("word_delete_matches", args, options); },
      addHyperlink: function (args, options) { return executeTool("word_add_hyperlink", args, options); },
      addComment: function (args, options) { return executeTool("word_add_comment", args, options); },
      addBookmark: function (args, options) { return executeTool("word_add_bookmark", args, options); },
      formatParagraphs: function (args, options) { return executeTool("word_format_paragraphs", args, options); },
      setParagraphText: function (args, options) { return executeTool("word_set_paragraph_text", args, options); },
      deleteParagraphs: function (args, options) { return executeTool("word_delete_paragraphs", args, options); },
      setList: function (args, options) { return executeTool("word_set_list", args, options); },
      insertPageBreak: function (args, options) { return executeTool("word_insert_page_break", args, options); },
      navigate: function (args, options) { return executeTool("word_navigate", args, options); },
      scroll: function (args, options) { return executeTool("word_scroll", args, options); },
      scaleFont: function (args, options) { return executeTool("word_scale_font", args, options); },
      addTable: function (args, options) { return executeTool("word_add_table", args, options); },
      setTableCell: function (args, options) { return executeTool("word_set_table_cell", args, options); },
      formatTable: function (args, options) { return executeTool("word_format_table", args, options); },
      editTable: function (args, options) { return executeTool("word_edit_table", args, options); },
      setPageLayout: function (args, options) { return executeTool("word_set_page_layout", args, options); },
      setHeaderFooter: function (args, options) { return executeTool("word_set_header_footer", args, options); },
      setDocumentText: function (args, options) { return executeTool("word_set_document_text", args, options); },
    },
    slides: {
      inspect: function (args, options) { return executeTool("slides_inspect", args, options); },
      replaceText: function (args, options) { return executeTool("slides_replace_text", args, options); },
      scaleFont: function (args, options) { return executeTool("slides_scale_font", args, options); },
      formatText: function (args, options) { return executeTool("slides_format_text", args, options); },
      formatSelection: function (args, options) { return executeTool("slides_format_selection", args, options); },
      addSlide: function (args, options) { return executeTool("slides_add_slide", args, options); },
      duplicateSlide: function (args, options) { return executeTool("slides_duplicate_slide", args, options); },
      deleteSlide: function (args, options) { return executeTool("slides_delete_slide", args, options); },
      addTextBox: function (args, options) { return executeTool("slides_add_textbox", args, options); },
    },
    sheets: {
      inspect: function (args, options) { return executeTool("sheets_inspect", args, options); },
      setValues: function (args, options) { return executeTool("sheets_set_values", args, options); },
      setFormula: function (args, options) { return executeTool("sheets_set_formula", args, options); },
      replaceText: function (args, options) { return executeTool("sheets_replace_text", args, options); },
      formatRange: function (args, options) { return executeTool("sheets_format_range", args, options); },
      addSheet: function (args, options) { return executeTool("sheets_add_sheet", args, options); },
      renameSheet: function (args, options) { return executeTool("sheets_rename_sheet", args, options); },
      deleteSheet: function (args, options) { return executeTool("sheets_delete_sheet", args, options); },
      addChart: function (args, options) { return executeTool("sheets_add_chart", args, options); },
    },
  };

  window.aiBridge = api;
  window.onlyofficeAI = api;
  window.AiBridgeError = window.AiBridgeError || AiBridgeError;

  function localHttpRelayEnabled() {
    if (window.aiBridgeOptions && window.aiBridgeOptions.httpRelay === false) return false;
    return window.location.hostname === "localhost"
      || window.location.hostname === "127.0.0.1"
      || window.location.hostname === "::1";
  }

  async function httpRelayPost(path, payload) {
    const response = await window.fetch(`/copilot-api/bridge/${path}`, {
      method: "POST",
      credentials: "same-origin",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    let body;
    try {
      body = await response.json();
    } catch (error) {
      throw bridgeError("HTTP_RELAY_INVALID_RESPONSE", `HTTP Relay 返回了无效响应：${response.status}`);
    }
    if (!response.ok || !body || body.ok === false) {
      const error = body && body.error || {};
      throw bridgeError(
        error.code || "HTTP_RELAY_FAILED",
        error.message || `HTTP Relay 请求失败：${response.status}`,
        { details: error.details },
      );
    }
    return body;
  }

  function startLocalHttpRelay() {
    if (!localHttpRelayEnabled() || typeof window.fetch !== "function") return;
    const httpSessionId = createRequestId("http-session");
    const credentialReloadWindowMs = 60000;
    let relayKey = null;
    let resumeToken = null;
    let stopped = false;
    let retryDelayMs = 1000;

    function setRelayState(value) {
      if (document.documentElement && document.documentElement.dataset) {
        document.documentElement.dataset.aiBridgeRelayState = value;
      }
    }

    function credentialReloadKey() {
      const snapshot = stateSnapshot();
      const context = snapshot && snapshot.context || {};
      return "aiBridgeCredentialReloadAt:" + [
        context.fileName,
        context.fileType,
        context.editorType,
        context.userId,
      ].map(function (value) { return String(value || ""); }).join("|");
    }

    function credentialReloadAt() {
      const key = credentialReloadKey();
      let latest = 0;
      try {
        const stored = Number(window.localStorage && window.localStorage.getItem(key));
        if (Number.isFinite(stored) && stored > latest) latest = stored;
      } catch (error) {
        // Fall back to per-tab storage when persistent storage is unavailable.
      }
      try {
        const stored = Number(window.sessionStorage && window.sessionStorage.getItem(key));
        if (Number.isFinite(stored) && stored > latest) latest = stored;
      } catch (error) {
        // Fall back to a URL marker when sessionStorage is unavailable.
      }
      if (latest > 0) return latest;
      try {
        return Number(new URL(window.location.href).searchParams.get("aiBridgeCredentialReloadAt")) || 0;
      } catch (error) {
        return 0;
      }
    }

    function rememberCredentialReload(value) {
      const key = credentialReloadKey();
      let stored = false;
      try {
        if (window.localStorage) {
          window.localStorage.setItem(key, String(value));
          stored = true;
        }
      } catch (error) {
        // Fall back to per-tab storage when persistent storage is unavailable.
      }
      try {
        if (window.sessionStorage) {
          window.sessionStorage.setItem(key, String(value));
          stored = true;
        }
      } catch (error) {
        // Fall back to a URL marker when sessionStorage is unavailable.
      }
      if (stored) return;
      try {
        const currentUrl = new URL(window.location.href);
        currentUrl.searchParams.set("aiBridgeCredentialReloadAt", String(value));
        window.history.replaceState(null, "", currentUrl.toString());
      } catch (error) {
        // The in-page terminal state still prevents another request loop.
      }
    }

    function stopRelay(reason, code) {
      stopped = true;
      relayKey = null;
      setRelayState(reason);
      emit("relayError", { reason, code: code || "HTTP_RELAY_STOPPED" });
    }

    function reloadForCredentialError(code) {
      const now = Date.now();
      if (now - credentialReloadAt() < credentialReloadWindowMs) {
        stopRelay("credential-error", code);
        return;
      }
      rememberCredentialReload(now);
      stopped = true;
      setRelayState("credential-reload");
      emit("reload", { reason: "relay-credentials-expired", code });
      window.setTimeout(function () { window.location.reload(); }, 50);
    }

    async function register() {
      await waitUntilReady(30000);
      const payload = {
        sessionId: httpSessionId,
        state: stateSnapshot(),
      };
      if (resumeToken) {
        payload.resumeToken = resumeToken;
      } else {
        const token = editorToken();
        if (!token) throw bridgeError("EDITOR_TOKEN_REQUIRED", "当前编辑器没有可用于 HTTP Relay 的绑定凭证");
        payload.editorToken = token;
      }
      const response = await httpRelayPost("register", payload);
      relayKey = response.relayKey;
      resumeToken = response.resumeToken || resumeToken;
      setRelayState("ready");
    }

    async function report(command, outcome) {
      await httpRelayPost("result", {
        sessionId: httpSessionId,
        relayKey,
        commandId: command.commandId,
        state: stateSnapshot(),
        ...outcome,
      });
    }

    async function loop() {
      while (!stopped) {
        try {
          if (!relayKey) await register();
          const response = await httpRelayPost("poll", {
            sessionId: httpSessionId,
            relayKey,
            timeoutMs: 25000,
            state: stateSnapshot(),
          });
          retryDelayMs = 1000;
          const command = response.command;
          if (!command) continue;
          const internalRequestId = relayRequestId(httpSessionId, command.requestId);
          try {
            const result = await executeRelayMethod(command.method, command.params || {}, internalRequestId);
            await report(command, { ok: true, result });
          } catch (error) {
            await report(command, { ok: false, error: serializedError(error, command.requestId) });
          }
        } catch (error) {
          const code = error && error.code;
          if (code === "INVALID_RELAY_SESSION") {
            relayKey = null;
            continue;
          }
          if (code === "INVALID_BRIDGE_RESUME_TOKEN") {
            relayKey = null;
            resumeToken = null;
            continue;
          }
          if (code === "SESSION_SUPERSEDED") {
            stopRelay("superseded", code);
            return;
          }
          if (
            code === "EDITOR_TOKEN_REQUIRED" ||
            code === "INVALID_EDITOR_TOKEN" ||
            code === "EDITOR_TOKEN_EXPIRED" ||
            code === "BRIDGE_RESUME_TOKEN_EXPIRED" ||
            code === "DOCUMENT_MISMATCH" ||
            code === "EDITOR_MISMATCH" ||
            code === "DOCUMENT_IDENTITY_MISMATCH" ||
            code === "INCOMPLETE_BRIDGE_IDENTITY"
          ) {
            reloadForCredentialError(code);
            return;
          }
          setRelayState("retrying");
          await new Promise(function (resolve) { window.setTimeout(resolve, retryDelayMs); });
          retryDelayMs = Math.min(10000, retryDelayMs * 2);
        }
      }
    }

    window.addEventListener("beforeunload", function () {
      stopped = true;
      if (!relayKey || !window.navigator || typeof window.navigator.sendBeacon !== "function") return;
      const body = JSON.stringify({ sessionId: httpSessionId, relayKey });
      window.navigator.sendBeacon("/copilot-api/bridge/unregister", body);
    });
    loop();
  }

  startLocalHttpRelay();

  const currentUrl = new URL(window.location.href);
  const initialCommand = currentUrl.searchParams.get("aiCommand");
  if (initialCommand) {
    currentUrl.searchParams.delete("aiCommand");
    window.history.replaceState(null, "", currentUrl.toString());
    try {
      api.lastExecution = execute(JSON.parse(initialCommand));
      api.lastExecution.then(
        function (result) { api.lastResult = result; },
        function (error) { api.lastError = error.message; },
      );
    } catch (error) {
      api.lastError = "aiCommand 必须是 JSON";
    }
  }
})();
