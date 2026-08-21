(function () {
  "use strict";

  const VERSION = "0.2.6";
  const PROTOCOL_VERSION = 1;
  const REQUEST_ID_PATTERN = /^[A-Za-z0-9._:-]{1,200}$/;
  const CLIENT_SOURCE = "ai-bridge-client";
  const RELAY_SOURCE = "ai-bridge-relay";
  let sequence = 0;

  class AiBridgeError extends Error {
    constructor(code, message, options) {
      super(message || "ai-bridge 调用失败");
      this.name = "AiBridgeError";
      this.code = code || "AI_BRIDGE_ERROR";
      this.requestId = options && options.requestId || null;
      this.details = options && options.details;
    }
  }

  function createId(prefix) {
    if (window.crypto && typeof window.crypto.randomUUID === "function") {
      return `${prefix}:${window.crypto.randomUUID()}`;
    }
    sequence += 1;
    return `${prefix}:${Date.now()}:${sequence}`;
  }

  function timeoutValue(value, fallback) {
    const timeout = Number(value === undefined ? fallback : value);
    if (!Number.isFinite(timeout)) return fallback;
    return Math.min(300000, Math.max(100, timeout));
  }

  function fromWireError(value, requestId) {
    if (value && typeof value === "object") {
      return new AiBridgeError(value.code, value.message, {
        requestId: value.requestId || requestId,
        details: value.details,
      });
    }
    return new AiBridgeError("AI_BRIDGE_ERROR", String(value || "ai-bridge 调用失败"), { requestId });
  }

  class AiBridgeClient {
    constructor(options) {
      const settings = options || {};
      if (!settings.targetWindow || typeof settings.targetWindow.postMessage !== "function") {
        throw new AiBridgeError("INVALID_TARGET_WINDOW", "targetWindow 必须是包含 ai-bridge 的父窗口或 opener");
      }
      if (typeof settings.targetOrigin !== "string" || !settings.targetOrigin || settings.targetOrigin === "*") {
        throw new AiBridgeError("INVALID_TARGET_ORIGIN", "targetOrigin 必须是明确的源，不能使用 *");
      }

      let parsedUrl;
      try { parsedUrl = new URL(settings.targetOrigin); } catch (error) {
        throw new AiBridgeError("INVALID_TARGET_ORIGIN", "targetOrigin 不是有效 URL 源");
      }
      if (!/^https?:$/.test(parsedUrl.protocol) || parsedUrl.username || parsedUrl.password ||
          (parsedUrl.pathname && parsedUrl.pathname !== "/") || parsedUrl.search || parsedUrl.hash) {
        throw new AiBridgeError("INVALID_TARGET_ORIGIN", "targetOrigin 只能包含协议、主机和端口");
      }
      const parsedOrigin = parsedUrl.origin;

      this.targetWindow = settings.targetWindow;
      this.targetOrigin = parsedOrigin;
      this.timeoutMs = timeoutValue(settings.timeoutMs, 90000);
      this.clientId = settings.clientId || createId("client");
      if (!REQUEST_ID_PATTERN.test(this.clientId)) {
        throw new AiBridgeError("INVALID_CLIENT_ID", "clientId 格式不合法");
      }

      this.version = VERSION;
      this.protocolVersion = PROTOCOL_VERSION;
      this.state = null;
      this.pending = new Map();
      this.listeners = new Map();
      this.connectPromise = null;
      this.destroyed = false;
      this.handleMessage = this.handleMessage.bind(this);
      window.addEventListener("message", this.handleMessage);

      this.word = this.createToolGroup({
        inspect: "word_inspect",
        replaceText: "word_replace_text",
        appendParagraph: "word_append_paragraph",
        insertParagraph: "word_insert_paragraph",
        formatDocument: "word_format_document",
        formatSelection: "word_format_selection",
        formatMatches: "word_format_matches",
        deleteMatches: "word_delete_matches",
        addHyperlink: "word_add_hyperlink",
        addComment: "word_add_comment",
        addBookmark: "word_add_bookmark",
        addImage: "word_add_image",
        inspectAdvanced: "word_inspect_advanced",
        setDocumentProperties: "word_set_document_properties",
        manageSection: "word_manage_section",
        manageStyle: "word_manage_style",
        setTabs: "word_set_tabs",
        setNumbering: "word_set_numbering",
        formatTableAdvanced: "word_format_table_advanced",
        addNestedTable: "word_add_nested_table",
        manageDrawing: "word_manage_drawing",
        addShape: "word_add_shape",
        addChart: "word_add_chart",
        addMath: "word_add_math",
        addOleObject: "word_add_ole_object",
        manageFields: "word_manage_fields",
        manageLongDocument: "word_manage_long_document",
        manageComments: "word_manage_comments",
        manageRevisions: "word_manage_revisions",
        setProtection: "word_set_protection",
        manageContentControl: "word_manage_content_control",
        manageCustomXml: "word_manage_custom_xml",
        inspectMacros: "word_inspect_macros",
        setMacros: "word_set_macros",
        setWatermark: "word_set_watermark",
        formatParagraphs: "word_format_paragraphs",
        setParagraphText: "word_set_paragraph_text",
        deleteParagraphs: "word_delete_paragraphs",
        setList: "word_set_list",
        insertPageBreak: "word_insert_page_break",
        navigate: "word_navigate",
        scroll: "word_scroll",
        scaleFont: "word_scale_font",
        addTable: "word_add_table",
        setTableCell: "word_set_table_cell",
        formatTable: "word_format_table",
        editTable: "word_edit_table",
        setPageLayout: "word_set_page_layout",
        setHeaderFooter: "word_set_header_footer",
        setDocumentText: "word_set_document_text",
      });
      this.slides = this.createToolGroup({
        inspect: "slides_inspect",
        inspectLayouts: "slides_inspect_layouts",
        inspectThemes: "slides_inspect_themes",
        inspectBuiltinThemes: "slides_inspect_builtin_themes",
        inspectObjects: "slides_inspect_objects",
        validateLayout: "slides_validate_layout",
        replaceText: "slides_replace_text",
        scaleFont: "slides_scale_font",
        formatText: "slides_format_text",
        formatSelection: "slides_format_selection",
        addSlide: "slides_add_slide",
        duplicateSlide: "slides_duplicate_slide",
        deleteSlide: "slides_delete_slide",
        moveSlide: "slides_move_slide",
        setVisibility: "slides_set_visibility",
        setSize: "slides_set_size",
        applyLayout: "slides_apply_layout",
        setShowSettings: "slides_set_show_settings",
        applyTheme: "slides_apply_theme",
        applyBuiltinTheme: "slides_apply_builtin_theme",
        setTheme: "slides_set_theme",
        createLayout: "slides_create_layout",
        addTemplateShape: "slides_add_template_shape",
        manageTemplateObject: "slides_manage_template_object",
        setTemplateBackground: "slides_set_template_background",
        setTextContent: "slides_set_text_content",
        formatParagraphs: "slides_format_paragraphs",
        updateObject: "slides_update_object",
        setHyperlink: "slides_set_hyperlink",
        setNotes: "slides_set_notes",
        addComment: "slides_add_comment",
        inspectComments: "slides_inspect_comments",
        manageComment: "slides_manage_comment",
        setTransition: "slides_set_transition",
        inspectAnimations: "slides_inspect_animations",
        manageAnimation: "slides_manage_animation",
        addTable: "slides_add_table",
        setTableCell: "slides_set_table_cell",
        editTable: "slides_edit_table",
        formatTable: "slides_format_table",
        alignObjects: "slides_align_objects",
        groupObjects: "slides_group_objects",
        reorderObject: "slides_reorder_object",
        addConnector: "slides_add_connector",
        addFreeform: "slides_add_freeform",
        addTextBox: "slides_add_textbox",
        addWordArt: "slides_add_word_art",
        addMath: "slides_add_math",
        addImage: "slides_add_image",
        addImageShape: "slides_add_image_shape",
        addOleObject: "slides_add_ole_object",
        setBackground: "slides_set_background",
        addShape: "slides_add_shape",
        updateShape: "slides_update_shape",
        deleteObject: "slides_delete_object",
        inspectSmartArts: "slides_inspect_smartarts",
        addSmartArt: "slides_add_smartart",
        updateSmartArt: "slides_update_smartart",
        deleteSmartArt: "slides_delete_smartart",
        inspectCharts: "slides_inspect_charts",
        addChart: "slides_add_chart",
        updateChart: "slides_update_chart",
        deleteChart: "slides_delete_chart",
        inspectMacros: "slides_inspect_macros",
        setMacros: "slides_set_macros",
        controlSlideshow: "slides_control_slideshow",
      });
      this.sheets = this.createToolGroup({
        inspect: "sheets_inspect",
        setValues: "sheets_set_values",
        setFormula: "sheets_set_formula",
        replaceText: "sheets_replace_text",
        formatRange: "sheets_format_range",
        addSheet: "sheets_add_sheet",
        renameSheet: "sheets_rename_sheet",
        deleteSheet: "sheets_delete_sheet",
        addChart: "sheets_add_chart",
        inspectCharts: "sheets_inspect_charts",
        updateChart: "sheets_update_chart",
        deleteChart: "sheets_delete_chart",
        inspectRange: "sheets_inspect_range",
        setArrayFormula: "sheets_set_array_formula",
        manageSheet: "sheets_manage_sheet",
        manageRange: "sheets_manage_range",
        setRichText: "sheets_set_rich_text",
        inspectNames: "sheets_inspect_names",
        manageNames: "sheets_manage_names",
        recalculate: "sheets_recalculate",
        sort: "sheets_sort",
        filter: "sheets_filter",
        inspectTables: "sheets_inspect_tables",
        manageTable: "sheets_manage_table",
        manageConditionalFormat: "sheets_manage_conditional_format",
        manageValidation: "sheets_manage_validation",
        inspectPivots: "sheets_inspect_pivots",
        managePivot: "sheets_manage_pivot",
        inspectDrawings: "sheets_inspect_drawings",
        manageDrawing: "sheets_manage_drawing",
        manageHyperlink: "sheets_manage_hyperlink",
        inspectComments: "sheets_inspect_comments",
        manageComments: "sheets_manage_comments",
        inspectFreezePanes: "sheets_inspect_freeze_panes",
        manageFreezePanes: "sheets_manage_freeze_panes",
        inspectProperties: "sheets_inspect_properties",
        manageProperties: "sheets_manage_properties",
        inspectProtectedRanges: "sheets_inspect_protected_ranges",
        manageProtectedRanges: "sheets_manage_protected_ranges",
        inspectPageLayout: "sheets_inspect_page_layout",
        managePageLayout: "sheets_manage_page_layout",
        inspectMacros: "sheets_inspect_macros",
        setMacros: "sheets_set_macros",
      });
    }

    get isReady() { return Boolean(this.state && this.state.ready); }
    get editorType() { return this.state && this.state.editorType || null; }
    get context() { return this.state && this.state.context || null; }
    get capabilities() { return this.state && this.state.capabilities || null; }
    get pluginGuid() { return this.state && this.state.pluginGuid || null; }

    basePayload(message) {
      return {
        source: CLIENT_SOURCE,
        protocolVersion: PROTOCOL_VERSION,
        clientId: this.clientId,
        ...message,
      };
    }

    post(message) {
      if (this.destroyed) throw new AiBridgeError("CLIENT_DESTROYED", "AiBridgeClient 已销毁");
      this.targetWindow.postMessage(this.basePayload(message), this.targetOrigin);
    }

    connect(options) {
      if (this.isReady) return Promise.resolve(this.state);
      if (this.connectPromise) return this.connectPromise;
      const timeoutMs = timeoutValue(options && options.timeoutMs, 30000);

      this.connectPromise = new Promise((resolve, reject) => {
        const timeout = window.setTimeout(() => {
          this.pending.delete("__connect__");
          this.connectPromise = null;
          reject(new AiBridgeError("CONNECTION_TIMEOUT", "连接 ai-bridge Relay 超时；请检查 clientOrigins 和 targetOrigin"));
        }, timeoutMs);
        this.pending.set("__connect__", { resolve, reject, timeout });
        this.post({ type: "connect", timeoutMs });
      });
      return this.connectPromise;
    }

    ready(options) { return this.connect(options); }

    handleMessage(event) {
      if (event.source !== this.targetWindow || event.origin !== this.targetOrigin) return;
      const message = event.data;
      if (!message || message.source !== RELAY_SOURCE || message.protocolVersion !== PROTOCOL_VERSION) return;
      if (message.clientId !== this.clientId) return;

      if (message.type === "connected" || message.type === "connection-error") {
        const connection = this.pending.get("__connect__");
        if (!connection) return;
        this.pending.delete("__connect__");
        window.clearTimeout(connection.timeout);
        if (message.error) {
          this.connectPromise = null;
          connection.reject(fromWireError(message.error));
        } else {
          this.state = message.state;
          connection.resolve(this.state);
          this.emit("ready", this.state);
        }
        return;
      }

      if (message.type === "event") {
        this.emit(message.event, message.detail);
        return;
      }
      if (message.type !== "result" || !message.requestId) return;

      const request = this.pending.get(message.requestId);
      if (!request) return;
      this.pending.delete(message.requestId);
      window.clearTimeout(request.timeout);
      if (message.error) {
        const error = fromWireError(message.error, message.requestId);
        this.emit("error", { error, requestId: message.requestId });
        request.reject(error);
      } else {
        if (request.method === "getState") this.state = message.result;
        request.resolve(message.result);
      }
    }

    async request(method, params, options) {
      const settings = options || {};
      await this.connect({ timeoutMs: settings.timeoutMs });
      const requestId = settings.requestId || createId("command");
      if (typeof requestId !== "string" || !REQUEST_ID_PATTERN.test(requestId)) {
        throw new AiBridgeError("INVALID_REQUEST_ID", "requestId 格式不合法");
      }
      if (this.pending.has(requestId)) {
        throw new AiBridgeError("REQUEST_IN_FLIGHT", `requestId 正在执行：${requestId}`, { requestId });
      }
      const timeoutMs = timeoutValue(settings.timeoutMs, this.timeoutMs);

      return new Promise((resolve, reject) => {
        const timeout = window.setTimeout(() => {
          this.pending.delete(requestId);
          const error = new AiBridgeError("TIMEOUT", "ai-bridge Relay 命令执行超时", { requestId });
          this.emit("error", { error, requestId });
          reject(error);
        }, timeoutMs);
        this.pending.set(requestId, { resolve, reject, timeout, method });
        this.post({
          type: "request",
          requestId,
          method,
          params: { ...(params || {}), options: { timeoutMs } },
        });
      });
    }

    execute(command, options) { return this.request("execute", { command }, options); }
    executeTool(name, args, options) { return this.request("executeTool", { name, arguments: args || {} }, options); }
    executeBatch(toolCalls, options) { return this.request("executeBatch", { toolCalls }, options); }
    save(options) { return this.request("save", {}, options); }
    history(options) { return this.request("history", {}, options); }
    undo(options) { return this.request("undo", {}, options); }
    redo(options) { return this.request("redo", {}, options); }
    refreshState(options) { return this.request("getState", {}, options); }

    createToolGroup(mapping) {
      const group = {};
      for (const method of Object.keys(mapping)) {
        group[method] = (args, options) => this.executeTool(mapping[method], args, options);
      }
      return Object.freeze(group);
    }

    on(name, listener) {
      if (typeof listener !== "function") throw new AiBridgeError("INVALID_LISTENER", "listener 必须是函数");
      const entries = this.listeners.get(name) || new Set();
      entries.add(listener);
      this.listeners.set(name, entries);
      return () => entries.delete(listener);
    }

    off(name, listener) {
      const entries = this.listeners.get(name);
      if (entries) entries.delete(listener);
    }

    emit(name, detail) {
      for (const listener of this.listeners.get(name) || []) {
        try { listener(detail); } catch (error) { window.setTimeout(function () { throw error; }, 0); }
      }
    }

    destroy() {
      if (this.destroyed) return;
      try { this.post({ type: "disconnect" }); } catch (error) {}
      this.destroyed = true;
      window.removeEventListener("message", this.handleMessage);
      for (const [requestId, request] of this.pending) {
        window.clearTimeout(request.timeout);
        request.reject(new AiBridgeError("CLIENT_DESTROYED", "AiBridgeClient 已销毁", {
          requestId: requestId === "__connect__" ? null : requestId,
        }));
      }
      this.pending.clear();
      this.listeners.clear();
      this.connectPromise = null;
    }
  }

  window.AiBridgeClient = AiBridgeClient;
  window.AiBridgeError = window.AiBridgeError || AiBridgeError;
})();
