(function () {
  "use strict";

  const PLUGIN_VERSION = "0.1.0";
  const PROTOCOL_VERSION = 1;
  const PLUGIN_GUID = "asc.{A17E5F31-64AA-4E37-9A42-8D430814C2F6}";
  const MAX_CALLS = 20;
  const MAX_CACHED_REQUESTS = 100;
  const REQUEST_ID_PATTERN = /^[A-Za-z0-9._:-]{1,200}$/;

  const ALLOWED_TOOLS = {
    word: new Set([
      "word_inspect",
      "word_replace_text",
      "word_append_paragraph",
      "word_insert_paragraph",
      "word_format_document",
      "word_format_selection",
      "word_scale_font",
      "word_add_table",
      "word_set_document_text",
    ]),
    slide: new Set([
      "slides_inspect",
      "slides_replace_text",
      "slides_scale_font",
      "slides_format_text",
      "slides_format_selection",
      "slides_add_slide",
      "slides_duplicate_slide",
      "slides_delete_slide",
      "slides_add_textbox",
    ]),
    cell: new Set([
      "sheets_inspect",
      "sheets_set_values",
      "sheets_set_formula",
      "sheets_replace_text",
      "sheets_format_range",
      "sheets_add_sheet",
      "sheets_rename_sheet",
      "sheets_delete_sheet",
      "sheets_add_chart",
    ]),
  };

  const state = {
    initialized: false,
    configured: false,
    editorType: null,
    bridge: null,
    hostOrigin: null,
    hostWindow: null,
    sessionId: null,
    config: {},
    queue: Promise.resolve(),
    inFlight: new Set(),
    responseCache: new Map(),
    announceTimer: null,
  };

  function bridgeError(code, message, details) {
    const error = new Error(message);
    error.name = "AiBridgeError";
    error.code = code;
    error.details = details;
    return error;
  }

  function errorPayload(error, requestId) {
    return {
      code: error && error.code || "EXECUTION_FAILED",
      message: error && error.message ? error.message : String(error),
      requestId: requestId || null,
      details: error && error.details,
    };
  }

  function resolveTopOrigin() {
    const ancestorOrigins = window.location.ancestorOrigins;
    if (ancestorOrigins && ancestorOrigins.length) {
      return ancestorOrigins[ancestorOrigins.length - 1];
    }
    try {
      return window.top.location.origin;
    } catch (error) {
      return null;
    }
  }

  function basePayload(message) {
    return {
      source: "ai-bridge-plugin",
      protocolVersion: PROTOCOL_VERSION,
      pluginVersion: PLUGIN_VERSION,
      pluginGuid: PLUGIN_GUID,
      ...message,
    };
  }

  function announce() {
    if (!state.initialized || state.configured) return;
    const hostOrigin = resolveTopOrigin();
    window.top.postMessage(basePayload({
      type: "hello",
      editorType: state.editorType,
    }), hostOrigin || "*");
  }

  function publish(message) {
    if (!state.configured || !state.hostWindow || !state.hostOrigin) return;
    state.hostWindow.postMessage(basePayload({
      sessionId: state.sessionId,
      ...message,
    }), state.hostOrigin);
  }

  function normalizeConfig(rawConfig) {
    const config = rawConfig && typeof rawConfig === "object" ? rawConfig : {};
    return {
      documentKey: typeof config.documentKey === "string" ? config.documentKey : "",
      fileName: typeof config.fileName === "string" ? config.fileName : "",
      fileType: typeof config.fileType === "string" ? config.fileType : "",
      editorType: typeof config.editorType === "string" ? config.editorType : "",
      callbackUrl: typeof config.callbackUrl === "string" ? config.callbackUrl : "",
      userId: typeof config.userId === "string" ? config.userId : "",
    };
  }

  function capabilities() {
    return {
      editorType: state.editorType,
      tools: Array.from(ALLOWED_TOOLS[state.editorType] || []),
      controls: ["save", "history", "undo", "redo"],
    };
  }

  function publishReady() {
    if (!state.initialized || !state.configured) return;
    publish({
      type: "ready",
      editorType: state.editorType,
      documentKey: state.config.documentKey,
      capabilities: capabilities(),
    });
  }

  function normalizeToolCalls(rawCalls) {
    if (!Array.isArray(rawCalls)) throw bridgeError("INVALID_COMMAND", "toolCalls 必须是数组");
    if (!rawCalls.length) return [];
    if (rawCalls.length > MAX_CALLS) throw bridgeError("TOO_MANY_CALLS", `单次最多执行 ${MAX_CALLS} 个工具`);

    const allowed = ALLOWED_TOOLS[state.editorType];
    return rawCalls.map(function (rawCall) {
      if (!rawCall || typeof rawCall !== "object") throw bridgeError("INVALID_TOOL_CALL", "工具调用必须是对象");
      const name = rawCall.name;
      if (typeof name !== "string" || !allowed || !allowed.has(name)) {
        throw bridgeError("TOOL_NOT_ALLOWED", `当前编辑器不允许工具：${name || "unknown"}`, {
          editorType: state.editorType,
          tool: name || null,
        });
      }

      let args = rawCall.arguments === undefined ? rawCall.args : rawCall.arguments;
      if (args === undefined || args === null) args = {};
      if (typeof args === "string") args = JSON.parse(args || "{}");
      if (!args || typeof args !== "object" || Array.isArray(args)) {
        throw bridgeError("INVALID_ARGUMENTS", `${name}.arguments 必须是对象`);
      }
      if (JSON.stringify(args).length > 250000) throw bridgeError("ARGUMENTS_TOO_LARGE", `${name}.arguments 过大`);
      return { id: rawCall.id || null, name, arguments: args };
    });
  }

  function validateTarget(target) {
    if (!target || typeof target !== "object") throw bridgeError("INVALID_TARGET", "命令缺少 target");
    if (!state.config.documentKey) throw bridgeError("DOCUMENT_NOT_CONFIGURED", "ai-bridge 尚未取得 document.key");
    if (target.documentKey !== state.config.documentKey) {
      throw bridgeError("DOCUMENT_MISMATCH", "命令目标不是当前文档", {
        expectedDocumentKey: state.config.documentKey,
        receivedDocumentKey: target.documentKey || null,
      });
    }
    if (target.editorType && target.editorType !== state.editorType) {
      throw bridgeError("EDITOR_MISMATCH", "命令目标编辑器类型不匹配", {
        expectedEditorType: state.editorType,
        receivedEditorType: target.editorType,
      });
    }
  }

  function saveInsideEditor() {
    return new Promise(function (resolve, reject) {
      Asc.plugin.callCommand(
        function () {
          try {
            Api.Save();
            return JSON.stringify({ ok: true });
          } catch (error) {
            return JSON.stringify({ ok: false, error: error && error.message ? error.message : String(error) });
          }
        },
        false,
        true,
        function (rawResult) {
          try {
            if (rawResult === undefined || rawResult === null || rawResult === "") {
              resolve({ ok: true });
              return;
            }
            const result = typeof rawResult === "string" ? JSON.parse(rawResult) : rawResult;
            if (result && result.ok === false) throw new Error(result.error || "编辑器保存失败");
            resolve(result || { ok: true });
          } catch (error) {
            reject(error);
          }
        },
      );
    });
  }

  async function apiRequest(path, payload) {
    const response = await fetch(`/copilot-api/${path}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    const result = await response.json().catch(function () { return {}; });
    if (!response.ok || result.error) {
      throw bridgeError("PERSISTENCE_FAILED", result.error || `ai-bridge 服务返回 ${response.status}`, {
        path,
        status: response.status,
      });
    }
    return result;
  }

  function versionRequest(action) {
    return apiRequest(action, {
      key: state.config.documentKey || null,
      fileName: state.config.fileName || null,
      userId: state.config.userId || null,
    });
  }

  function forceSave(options) {
    return apiRequest("forcesave", {
      key: state.config.documentKey,
      fileName: state.config.fileName || null,
      allowNoChanges: Boolean(options && options.allowNoChanges),
    });
  }

  async function executeTools(toolCalls) {
    const mutating = toolCalls.some(function (call) { return !call.name.endsWith("_inspect"); });
    if (mutating) await versionRequest("checkpoint");

    const result = toolCalls.length
      ? await state.bridge.execute(toolCalls)
      : { changed: 0, results: [] };

    if (!mutating) return { ...result, persisted: false };
    await saveInsideEditor();
    const saved = await forceSave();
    return { ...result, forceSave: saved, persisted: Boolean(saved.persisted) };
  }

  async function executeCommand(command) {
    if (!command || typeof command !== "object") throw bridgeError("INVALID_COMMAND", "command 必须是对象");
    let toolCalls = command.toolCalls;
    if (!toolCalls && command.name) toolCalls = [command];
    return executeTools(normalizeToolCalls(toolCalls));
  }

  async function executeControl(action) {
    switch (action) {
      case "save": {
        await saveInsideEditor();
        return { result: await forceSave({ allowNoChanges: true }), reload: false };
      }
      case "history":
        return { result: await versionRequest("history"), reload: false };
      case "undo":
      case "redo":
        return { result: await versionRequest(action), reload: true };
      default:
        throw bridgeError("CONTROL_NOT_ALLOWED", `不允许控制操作：${action || "unknown"}`);
    }
  }

  function rememberResponse(requestId, payload) {
    state.responseCache.set(requestId, payload);
    while (state.responseCache.size > MAX_CACHED_REQUESTS) {
      state.responseCache.delete(state.responseCache.keys().next().value);
    }
  }

  function processRequest(message) {
    const requestId = message.requestId;
    if (typeof requestId !== "string" || !REQUEST_ID_PATTERN.test(requestId)) return;

    const cached = state.responseCache.get(requestId);
    if (cached) {
      publish(cached);
      return;
    }
    if (state.inFlight.has(requestId)) return;
    state.inFlight.add(requestId);

    state.queue = state.queue.then(async function () {
      let response;
      let reload = false;
      try {
        validateTarget(message.target);
        if (message.type === "execute") {
          response = { type: "result", requestId, result: await executeCommand(message.command) };
        } else if (message.type === "control") {
          const controlled = await executeControl(message.action);
          response = { type: "result", requestId, result: controlled.result };
          reload = controlled.reload;
        } else {
          throw bridgeError("MESSAGE_NOT_SUPPORTED", `不支持消息类型：${message.type}`);
        }
      } catch (error) {
        response = { type: "result", requestId, error: errorPayload(error, requestId) };
      } finally {
        state.inFlight.delete(requestId);
      }

      rememberResponse(requestId, response);
      publish(response);
      if (reload) window.setTimeout(function () { publish({ type: "reload", reason: message.action }); }, 50);
    }).catch(function () {
      state.inFlight.delete(requestId);
    });
  }

  function configure(event, message) {
    if (message.protocolVersion !== PROTOCOL_VERSION) return;
    if (typeof message.sessionId !== "string" || !REQUEST_ID_PATTERN.test(message.sessionId)) return;
    if (state.configured && message.sessionId !== state.sessionId) return;

    const config = normalizeConfig(message.config);
    if (config.editorType && state.editorType && config.editorType !== state.editorType) return;
    if (state.configured && state.config.documentKey && config.documentKey !== state.config.documentKey) return;

    state.hostWindow = event.source;
    state.hostOrigin = event.origin;
    state.sessionId = message.sessionId;
    state.config = config;
    state.configured = true;
    if (state.announceTimer) {
      window.clearInterval(state.announceTimer);
      state.announceTimer = null;
    }
    publishReady();
  }

  window.addEventListener("message", function (event) {
    if (event.source !== window.top) return;
    const expectedOrigin = resolveTopOrigin();
    if (expectedOrigin && event.origin !== expectedOrigin) return;

    const message = event.data;
    if (!message || message.source !== "ai-bridge-host") return;
    if (message.pluginGuid !== PLUGIN_GUID) return;

    if (message.type === "configure") {
      configure(event, message);
      return;
    }
    if (!state.configured || event.origin !== state.hostOrigin || message.sessionId !== state.sessionId) return;
    if (message.type === "host-ready") {
      publishReady();
      return;
    }
    if (message.type === "execute" || message.type === "control") processRequest(message);
  });

  Asc.plugin.init = function () {
    const editorType = Asc.plugin.info && Asc.plugin.info.editorType;
    const bridge = editorType && window.AICopilotBridges && window.AICopilotBridges[editorType];
    if (!bridge || !ALLOWED_TOOLS[editorType]) return;

    state.editorType = editorType;
    state.bridge = bridge;
    state.initialized = true;
    state.announceTimer = window.setInterval(announce, 500);
    announce();
  };

  Asc.plugin.button = function () {};
})();
