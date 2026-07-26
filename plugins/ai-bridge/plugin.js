(function () {
  "use strict";

  const PLUGIN_VERSION = "0.4.0";
  const PROTOCOL_VERSION = 1;
  const PLUGIN_GUID = "asc.{A17E5F31-64AA-4E37-9A42-8D430814C2F6}";
  const MAX_CALLS = 20;
  const MAX_CACHED_REQUESTS = 100;
  const REQUEST_ID_PATTERN = /^[A-Za-z0-9._:-]{1,200}$/;
  const READ_ONLY_TOOLS = new Set([
    "word_inspect",
    "word_inspect_advanced",
    "word_inspect_macros",
    "word_navigate",
    "word_scroll",
    "slides_inspect",
    "slides_inspect_layouts",
    "slides_inspect_themes",
    "slides_inspect_builtin_themes",
    "slides_inspect_objects",
    "slides_inspect_charts",
    "slides_inspect_animations",
    "slides_inspect_comments",
    "slides_inspect_macros",
    "slides_control_slideshow",
    "sheets_inspect",
    "sheets_inspect_charts",
    "sheets_inspect_range",
    "sheets_inspect_names",
    "sheets_inspect_tables",
    "sheets_inspect_pivots",
    "sheets_inspect_drawings",
    "sheets_inspect_comments",
    "sheets_inspect_freeze_panes",
    "sheets_inspect_properties",
    "sheets_inspect_protected_ranges",
    "sheets_inspect_page_layout",
    "sheets_inspect_macros",
  ]);

  const ALLOWED_TOOLS = {
    word: new Set([
      "word_inspect",
      "word_replace_text",
      "word_append_paragraph",
      "word_insert_paragraph",
      "word_format_document",
      "word_format_selection",
      "word_format_matches",
      "word_delete_matches",
      "word_add_hyperlink",
      "word_add_comment",
      "word_add_bookmark",
      "word_add_image",
      "word_inspect_advanced",
      "word_set_document_properties",
      "word_manage_section",
      "word_manage_style",
      "word_set_tabs",
      "word_set_numbering",
      "word_format_table_advanced",
      "word_add_nested_table",
      "word_manage_drawing",
      "word_add_shape",
      "word_add_chart",
      "word_add_math",
      "word_add_ole_object",
      "word_manage_fields",
      "word_manage_long_document",
      "word_manage_comments",
      "word_manage_revisions",
      "word_set_protection",
      "word_manage_content_control",
      "word_manage_custom_xml",
      "word_inspect_macros",
      "word_set_macros",
      "word_set_watermark",
      "word_format_paragraphs",
      "word_set_paragraph_text",
      "word_delete_paragraphs",
      "word_set_list",
      "word_insert_page_break",
      "word_navigate",
      "word_scroll",
      "word_scale_font",
      "word_add_table",
      "word_set_table_cell",
      "word_format_table",
      "word_edit_table",
      "word_set_page_layout",
      "word_set_header_footer",
      "word_set_document_text",
    ]),
    slide: new Set([
      "slides_inspect",
      "slides_inspect_layouts",
      "slides_inspect_themes",
      "slides_inspect_builtin_themes",
      "slides_replace_text",
      "slides_scale_font",
      "slides_format_text",
      "slides_format_selection",
      "slides_add_slide",
      "slides_duplicate_slide",
      "slides_delete_slide",
      "slides_move_slide",
      "slides_set_visibility",
      "slides_set_size",
      "slides_apply_layout",
      "slides_set_show_settings",
      "slides_apply_theme",
      "slides_apply_builtin_theme",
      "slides_set_theme",
      "slides_create_layout",
      "slides_add_template_shape",
      "slides_manage_template_object",
      "slides_set_template_background",
      "slides_set_text_content",
      "slides_format_paragraphs",
      "slides_update_object",
      "slides_set_hyperlink",
      "slides_set_notes",
      "slides_add_comment",
      "slides_inspect_comments",
      "slides_manage_comment",
      "slides_set_transition",
      "slides_inspect_animations",
      "slides_manage_animation",
      "slides_add_table",
      "slides_set_table_cell",
      "slides_edit_table",
      "slides_format_table",
      "slides_align_objects",
      "slides_group_objects",
      "slides_reorder_object",
      "slides_add_connector",
      "slides_add_freeform",
      "slides_add_textbox",
      "slides_add_word_art",
      "slides_add_math",
      "slides_add_image",
      "slides_add_image_shape",
      "slides_add_ole_object",
      "slides_inspect_objects",
      "slides_set_background",
      "slides_add_shape",
      "slides_update_shape",
      "slides_delete_object",
      "slides_inspect_charts",
      "slides_add_chart",
      "slides_update_chart",
      "slides_delete_chart",
      "slides_inspect_macros",
      "slides_set_macros",
      "slides_control_slideshow",
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
      "sheets_inspect_charts",
      "sheets_update_chart",
      "sheets_delete_chart",
      "sheets_inspect_range",
      "sheets_set_array_formula",
      "sheets_manage_sheet",
      "sheets_manage_range",
      "sheets_set_rich_text",
      "sheets_inspect_names",
      "sheets_manage_names",
      "sheets_recalculate",
      "sheets_sort",
      "sheets_filter",
      "sheets_inspect_tables",
      "sheets_manage_table",
      "sheets_manage_conditional_format",
      "sheets_manage_validation",
      "sheets_inspect_pivots",
      "sheets_manage_pivot",
      "sheets_inspect_drawings",
      "sheets_manage_drawing",
      "sheets_manage_hyperlink",
      "sheets_inspect_comments",
      "sheets_manage_comments",
      "sheets_inspect_freeze_panes",
      "sheets_manage_freeze_panes",
      "sheets_inspect_properties",
      "sheets_manage_properties",
      "sheets_inspect_protected_ranges",
      "sheets_manage_protected_ranges",
      "sheets_inspect_page_layout",
      "sheets_manage_page_layout",
      "sheets_inspect_macros",
      "sheets_set_macros",
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
    defaultDocumentLanguageApplied: false,
  };

  function bridgeError(code, message, details) {
    const error = new Error(message);
    error.name = "AiBridgeError";
    error.code = code;
    error.details = details;
    return error;
  }

  function withErrorPhase(error, phase, extraDetails) {
    const phasedError = error instanceof Error ? error : new Error(String(error));
    const details = phasedError.details && typeof phasedError.details === "object"
      ? { ...phasedError.details }
      : {};
    if (!details.phase) details.phase = phase;
    if (extraDetails && typeof extraDetails === "object") {
      Object.assign(details, extraDetails);
    }
    phasedError.details = details;
    return phasedError;
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
      interfaceLanguage: typeof config.interfaceLanguage === "string" ? config.interfaceLanguage : "",
      region: typeof config.region === "string" ? config.region : "",
    };
  }

  function applyDefaultChineseDocumentLanguage() {
    if (
      state.editorType !== "word"
      || !/^zh(?:-|$)/i.test(state.config.interfaceLanguage)
      || state.defaultDocumentLanguageApplied
    ) {
      return;
    }

    state.defaultDocumentLanguageApplied = true;
    try {
      Asc.scope.aiBridgeDefaultDocumentLanguage = 0x0004;
      Asc.plugin.callCommand(
        function () {
          try {
            if (
              typeof Asc === "undefined"
              || !Asc.editor
              || typeof Asc.editor.asc_setDefaultLanguage !== "function"
            ) {
              return JSON.stringify({
                ok: false,
                error: "ONLYOFFICE 默认文档语言 API 不可用",
              });
            }
            Asc.editor.asc_setDefaultLanguage(Asc.scope.aiBridgeDefaultDocumentLanguage);
            return JSON.stringify({ ok: true });
          } catch (error) {
            return JSON.stringify({
              ok: false,
              error: error && error.message ? error.message : String(error),
            });
          }
        },
        false,
        true,
        function () {},
      );
    } catch (error) {
      state.defaultDocumentLanguageApplied = false;
    }
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
      if (
        name === "word_add_image"
        || name === "slides_add_image"
        || name === "slides_add_image_shape"
        || name === "word_add_ole_object"
        || name === "slides_add_ole_object"
        || (name === "sheets_manage_drawing" && args._image !== undefined)
        || (name === "word_set_watermark" && args._image !== undefined)
      ) {
        if (args.source !== undefined) {
          throw bridgeError("INVALID_IMAGE_SOURCE", "图片来源尚未经过宿主页安全导入");
        }
        const image = args._image;
        let parsedImageUrl;
        try {
          parsedImageUrl = image && typeof image.url === "string" ? new URL(image.url) : null;
        } catch (error) {
          parsedImageUrl = null;
        }
        const signedSameOriginUrl = Boolean(
          parsedImageUrl
          && parsedImageUrl.origin !== "null"
          && parsedImageUrl.origin === state.hostOrigin
          && /^\/copilot-api\/images\/[0-9a-f]{64}\.(?:png|jpg|gif|webp|svg)$/.test(parsedImageUrl.pathname)
          && parsedImageUrl.searchParams.get("token")
        );
        const internalDataUrl = Boolean(
          image
          && image.transport === "dataUrl"
          && typeof image.url === "string"
          && /^data:(?:image\/(?:png|jpeg|gif|webp)|image\/svg\+xml);base64,[A-Za-z0-9+/]+={0,2}$/.test(image.url)
          && /^[0-9a-f]{64}\.(?:png|jpg|gif|webp|svg)$/.test(String(image.assetId || ""))
        );
        if (
          (!signedSameOriginUrl && !internalDataUrl)
          || !Number.isFinite(Number(image.widthPx))
          || !Number.isFinite(Number(image.heightPx))
          || Number(image.widthPx) <= 0
          || Number(image.heightPx) <= 0
        ) {
          throw bridgeError("INVALID_IMAGE_SOURCE", "内部图片资源无效或来源不受信任");
        }
      }
      const argumentLimit = (
        args._image && args._image.transport === "dataUrl"
          ? 12000000
          : 250000
      );
      if (JSON.stringify(args).length > argumentLimit) {
        throw bridgeError("ARGUMENTS_TOO_LARGE", `${name}.arguments 过大`);
      }
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
    }).catch(function (error) {
      throw withErrorPhase(error, "editor-save", {
        partialMutationPossible: true,
      });
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

  async function forceSave(options) {
    try {
      return await apiRequest("forcesave", {
        key: state.config.documentKey,
        fileName: state.config.fileName || null,
        allowNoChanges: Boolean(options && options.allowNoChanges),
      });
    } catch (error) {
      throw withErrorPhase(error, "persistence", {
        partialMutationPossible: true,
      });
    }
  }

  async function executeTools(toolCalls) {
    const mutating = toolCalls.some(function (call) { return !READ_ONLY_TOOLS.has(call.name); });
    if (mutating) await versionRequest("checkpoint");

    const result = toolCalls.length
      ? await state.bridge.execute(toolCalls)
      : { changed: 0, results: [] };

    if (!mutating || !result.needsSave) {
      return { result: { ...result, persisted: false }, reload: false };
    }
    await saveInsideEditor();
    const saved = await forceSave();
    return {
      result: { ...result, forceSave: saved, persisted: Boolean(saved.persisted) },
      // The live editor already contains this version. Force-save only persists
      // it through CommandService/callback; reloading would discard the user's
      // current page, selection, and scroll position.
      reload: false,
    };
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
        const saved = await forceSave({ allowNoChanges: true });
        return { result: saved, reload: false };
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
          const executed = await executeCommand(message.command);
          response = { type: "result", requestId, result: executed.result };
          reload = executed.reload;
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
      if (reload) {
        const reason = message.type === "execute" ? "document-version-changed" : message.action;
        window.setTimeout(function () { publish({ type: "reload", reason }); }, 50);
      }
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
    applyDefaultChineseDocumentLanguage();
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
