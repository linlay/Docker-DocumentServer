(function () {
  "use strict";

  const VERSION = "0.2.5";
  const PROTOCOL_VERSION = 1;
  const CONTRACT_SHA256 = "3a8940e1e7ba8b042d65871eabb23d7ee1cc3f1a5e7f1854ac82081490e05d3e";
  const PLUGIN_GUID = "asc.{A17E5F31-64AA-4E37-9A42-8D430814C2F6}";
  const REQUEST_ID_PATTERN = /^[A-Za-z0-9._:-]{1,200}$/;
  const CHANNEL_ID_PATTERN = /^[A-Za-z0-9._:-]{16,200}$/;
  const DOCUMENT_ID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/;
  const EXTERNAL_CLIENT_SOURCE = "ai-bridge-client";
  const EXTERNAL_RELAY_SOURCE = "ai-bridge-relay";
  const currentScript = document.currentScript;
  const configuredOptions = window.aiBridgeOptions && typeof window.aiBridgeOptions === "object"
    ? window.aiBridgeOptions
    : {};
  const pluginOrigin = currentScript
    ? new URL(currentScript.src, window.location.href).origin
    : window.location.origin;
  let sequence = 0;

  if (window.aiBridge && window.aiBridge.version === VERSION) return;
  document.documentElement.dataset.aiBridgeState = "loading";

  const pluginHandshake = configuredPluginHandshake();
  const pending = new Map();
  const readyWaiters = new Set();
  const listeners = new Map();
  const relayClients = new Map();
  const preparedImageCommands = new Map();
  const IMAGE_TOOL_NAMES = new Set(["word_add_image","word_add_ole_object","slides_add_image","slides_add_image_shape","slides_add_ole_object"]);
  const OPTIONAL_IMAGE_TOOL_NAMES = new Set(["word_set_watermark","slides_set_template_background","slides_set_background","sheets_manage_drawing"]);
  const sessionId = createRequestId("session");
  let pluginWindow = null;
  let ready = false;
  let editorType = null;
  let bridgeCapabilities = null;
  let bridgeContractVersion = VERSION;
  let bridgeContractSha256 = CONTRACT_SHA256;
  let capabilityProbedAt = null;
  let saveStatus = "idle";
  let saveStatusAt = Date.now();
  let bridgeTerminalError = null;
  let relayStateValue = null;
  let relayStateCode = null;
  let relayStateSignature = null;
  const startupNavigationStartedAt = (function () {
    const performance = window.performance;
    if (performance && Number.isFinite(Number(performance.timeOrigin))) {
      return Number(performance.timeOrigin);
    }
    if (
      performance
      && performance.timing
      && Number.isFinite(Number(performance.timing.navigationStart))
    ) {
      return Number(performance.timing.navigationStart);
    }
    return Date.now();
  }());
  const startupMarks = {};

  function startupElapsedMs() {
    return Math.max(0, Math.round(Date.now() - startupNavigationStartedAt));
  }

  function markStartup(field) {
    if (startupMarks[field] === undefined) startupMarks[field] = startupElapsedMs();
  }

  markStartup("hostScriptMs");
  window.addEventListener("load", function () { markStartup("windowLoadMs"); });

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

  function relayStateDetail(value, code, error) {
    const diagnostic = error ? serializedError(error) : null;
    const detail = { state: value };
    const normalizedCode = code || diagnostic && diagnostic.code || null;
    if (normalizedCode) detail.code = normalizedCode;
    if (diagnostic && diagnostic.message) detail.message = diagnostic.message;
    if (diagnostic && diagnostic.requestId) detail.requestId = diagnostic.requestId;
    if (diagnostic && diagnostic.details !== undefined) detail.details = diagnostic.details;
    return detail;
  }

  function setRelayState(value, code, error) {
    if (document.documentElement && document.documentElement.dataset) {
      document.documentElement.dataset.aiBridgeRelayState = value;
    }
    const detail = relayStateDetail(value, code, error);
    const normalizedCode = detail.code || null;
    let signature;
    try {
      signature = JSON.stringify(detail);
    } catch (signatureError) {
      signature = `${value}:${normalizedCode || ""}:${detail.message || ""}`;
    }
    if (relayStateSignature === signature) return;
    relayStateValue = value;
    relayStateCode = normalizedCode;
    relayStateSignature = signature;
    window.dispatchEvent(new CustomEvent("ai-bridge-relay-state", { detail }));
  }

  function createRequestId(prefix) {
    if (window.crypto && typeof window.crypto.randomUUID === "function") {
      return `${prefix || "request"}:${window.crypto.randomUUID()}`;
    }
    sequence += 1;
    return `${prefix || "request"}:${Date.now()}:${sequence}`;
  }

  function editorConfiguration() {
    if (typeof configuredOptions.getEditorConfig === "function") {
      return configuredOptions.getEditorConfig() || {};
    }
    return typeof config === "object" && config ? config : {};
  }

  function configuredPluginHandshake() {
    const editorConfig = editorConfiguration();
    const plugins = editorConfig.editorConfig
      && editorConfig.editorConfig.plugins
      && typeof editorConfig.editorConfig.plugins === "object"
      ? editorConfig.editorConfig.plugins
      : {};
    const options = plugins.options && typeof plugins.options === "object"
      ? plugins.options
      : {};
    const raw = options[PLUGIN_GUID];
    if (!raw || typeof raw !== "object") {
      return { strict: false, valid: true, hostOrigin: null, channelId: null };
    }

    const hostOrigin = typeof raw.hostOrigin === "string" ? raw.hostOrigin : "";
    const channelId = typeof raw.channelId === "string" ? raw.channelId : "";
    let normalizedOrigin = "";
    try {
      normalizedOrigin = new URL(hostOrigin).origin;
    } catch (error) {
      // Strict mode remains enabled, but no message is accepted.
    }
    return {
      strict: true,
      valid: (
        normalizedOrigin === hostOrigin
        && hostOrigin === window.location.origin
        && CHANNEL_ID_PATTERN.test(channelId)
      ),
      hostOrigin,
      channelId,
    };
  }

  function editorToken() {
    const editorConfig = editorConfiguration();
    return typeof editorConfig.token === "string" ? editorConfig.token : "";
  }

  function configuredDocumentId(editorConfig) {
    const configured = [editorConfig.documentId, configuredOptions.documentId].find(function (value) {
      return typeof value === "string" && DOCUMENT_ID_PATTERN.test(value);
    });
    if (configured) return configured;
    const match = String(window.location.pathname || "").match(
      /^\/documents\/([0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12})\/?$/,
    );
    return match ? match[1] : "";
  }

  function currentConfig() {
    const editorConfig = editorConfiguration();
    return {
      documentId: configuredDocumentId(editorConfig),
      documentKey: editorConfig.document && editorConfig.document.key || "",
      fileName: editorConfig.document && editorConfig.document.title || "",
      fileType: editorConfig.document && editorConfig.document.fileType || "",
      editorType: editorConfig.documentType || "",
      callbackUrl: editorConfig.editorConfig && editorConfig.editorConfig.callbackUrl || "",
      userId: editorConfig.editorConfig && editorConfig.editorConfig.user && editorConfig.editorConfig.user.id || "",
      interfaceLanguage: editorConfig.editorConfig && editorConfig.editorConfig.lang || "",
      region: editorConfig.editorConfig && editorConfig.editorConfig.region || "",
      persistenceViaHost: Boolean(persistenceBaseURL()),
    };
  }

  function publicContext() {
    const context = currentConfig();
    return {
      documentId: context.documentId,
      documentKey: context.documentKey,
      fileName: context.fileName,
      fileType: context.fileType,
      editorType: context.editorType,
      userId: context.userId,
    };
  }

  function installEditorStartupInstrumentation() {
    const editorConfig = editorConfiguration();
    if (!editorConfig || typeof editorConfig !== "object") return;
    const events = editorConfig.events && typeof editorConfig.events === "object"
      ? editorConfig.events
      : (editorConfig.events = {});
    [
      ["onAppReady", "appReadyMs"],
      ["onDocumentReady", "documentReadyMs"],
    ].forEach(function (entry) {
      const eventName = entry[0];
      const timingField = entry[1];
      const original = events[eventName];
      if (original && original.__aiBridgeStartupWrapped === true) return;
      const wrapped = function () {
        markStartup(timingField);
        if (typeof original === "function") return original.apply(this, arguments);
        return undefined;
      };
      wrapped.__aiBridgeStartupWrapped = true;
      events[eventName] = wrapped;
    });
  }

  function stateSnapshot() {
    const documentReady = startupMarks.documentReadyMs !== undefined;
    const capabilityProbeComplete = Boolean(
      ready
      && bridgeCapabilities
      && (
        editorType !== "cell"
        || (
          bridgeCapabilities.features
          && bridgeCapabilities.features.sheets
        )
      )
    );
    return {
      version: VERSION,
      protocolVersion: PROTOCOL_VERSION,
      pluginGuid: PLUGIN_GUID,
      contractVersion: bridgeContractVersion,
      contractSha256: bridgeContractSha256,
      ready,
      documentReady,
      capabilityProbeComplete,
      capabilityProbedAt,
      saveReady: Boolean(ready && documentReady && capabilityProbeComplete && saveStatus !== "saving"),
      saveStatus,
      saveStatusAt,
      editorType,
      context: publicContext(),
      capabilities: bridgeCapabilities,
    };
  }

  function relayStateSnapshot() {
    return {
      ...stateSnapshot(),
      _diagnostics: {
        startup: { ...startupMarks },
      },
    };
  }

  function basePayload(message) {
    const payload = {
      source: "ai-bridge-host",
      protocolVersion: PROTOCOL_VERSION,
      pluginGuid: PLUGIN_GUID,
      sessionId,
      ...message,
    };
    if (pluginHandshake.strict) payload.channelId = pluginHandshake.channelId;
    return payload;
  }

  function send(message) {
    if (!pluginWindow || (pluginHandshake.strict && !pluginHandshake.valid)) return false;
    pluginWindow.postMessage(basePayload(message), pluginOrigin);
    return true;
  }

  function sendConfiguration() {
    send({ type: "configure", config: currentConfig() });
  }

  function waitUntilReady(timeoutMs) {
    if (ready) return Promise.resolve();
    if (bridgeTerminalError) return Promise.reject(bridgeTerminalError);
    return new Promise(function (resolve, reject) {
      const waiter = {};
      const timeout = window.setTimeout(function () {
        readyWaiters.delete(waiter);
        reject(bridgeError("NOT_READY", "ai-bridge 插件尚未就绪"));
      }, timeoutMs);
      waiter.resolve = function () {
        window.clearTimeout(timeout);
        resolve();
      };
      waiter.reject = function (error) {
        window.clearTimeout(timeout);
        reject(error);
      };
      readyWaiters.add(waiter);
      send({ type: "host-ready" });
    });
  }

  function markContractMismatch(actualSha256) {
    if (bridgeTerminalError) return;
    ready = false;
    bridgeTerminalError = bridgeError(
      "CONTRACT_MISMATCH",
      "ai-bridge 插件资源版本与页面不一致，请关闭并重新打开文档",
      {
        details: {
          expectedContractSha256: CONTRACT_SHA256,
          actualContractSha256: actualSha256 || null,
        },
      },
    );
    document.documentElement.dataset.aiBridgeState = "contract-mismatch";
    setRelayState("contract-mismatch", "CONTRACT_MISMATCH", bridgeTerminalError);
    for (const waiter of readyWaiters) waiter.reject(bridgeTerminalError);
    readyWaiters.clear();
  }

  function markReady(message) {
    markStartup("bridgeReadyMs");
    bridgeContractVersion = message.contractVersion || message.pluginVersion || VERSION;
    bridgeContractSha256 = message.contractSha256 || CONTRACT_SHA256;
    if (bridgeContractSha256 !== CONTRACT_SHA256) {
      markContractMismatch(bridgeContractSha256);
      return;
    }
    if (bridgeTerminalError) return;
    ready = true;
    editorType = message.editorType || editorType;
    bridgeCapabilities = message.capabilities || bridgeCapabilities;
    if (bridgeCapabilities && capabilityProbedAt === null) capabilityProbedAt = Date.now();
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
    if (message.contractSha256 && message.contractSha256 !== CONTRACT_SHA256) {
      markContractMismatch(message.contractSha256);
      return;
    }
    if (pluginHandshake.strict) {
      if (!pluginHandshake.valid || message.channelId !== pluginHandshake.channelId) return;
    } else if (message.channelId !== undefined) {
      return;
    }

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
    if (message.type === "service-request") {
      handlePluginService(message);
      return;
    }
    if (message.type === "persistence-state") {
      saveStatus = String(message.saveStatus || "unknown");
      saveStatusAt = Number(message.saveStatusAt) || Date.now();
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

  async function importImageSource(source) {
    const token = editorToken();
    if (!token) throw bridgeError("EDITOR_TOKEN_REQUIRED", "当前编辑器没有可用于导入图片的凭证");
    const imageBaseURL = imageRelayBaseURL();
    if (!imageBaseURL) throw bridgeError("IMAGE_FETCH_FAILED", "当前页面没有配置图片导入服务");
    if (!source || typeof source !== "object" || Array.isArray(source)) {
      throw bridgeError("INVALID_IMAGE_SOURCE", "source 必须是图片来源对象");
    }
    let asset;
    if (source.type === "url") {
      if (typeof source.url !== "string" || !/^https:\/\//i.test(source.url)) {
        throw bridgeError("INVALID_IMAGE_SOURCE", "外部图片只允许 HTTPS URL");
      }
    } else if (source.type === "dataUrl") {
      if (typeof source.dataUrl !== "string" || !/^data:(?:image\/(?:png|jpeg|gif|webp)|image\/svg\+xml);base64,/i.test(source.dataUrl)) {
        throw bridgeError("INVALID_IMAGE_SOURCE", "Data URL 必须是受支持图片的 Base64 编码");
      }
    } else if (source.type === "relayAsset") {
      asset = {
        path: source.path,
        assetToken: source.assetToken,
        assetId: source.assetId,
        mimeType: source.mimeType,
        widthPx: source.widthPx,
        heightPx: source.heightPx,
      };
    } else {
      throw bridgeError("INVALID_IMAGE_SOURCE", "source.type 必须是 url、dataUrl 或 relayAsset");
    }
    if (!asset) {
      if (typeof window.fetch !== "function") {
        throw bridgeError("IMAGE_FETCH_FAILED", "当前宿主页不支持图片导入请求");
      }
      let response;
      try {
        response = await window.fetch(`${imageBaseURL}/import`, {
          method: "POST",
          credentials: "same-origin",
          headers: {
            ...httpRelayHeaders(),
            "Authorization": `Bearer ${token}`,
          },
          body: JSON.stringify({ source }),
        });
      } catch (error) {
        throw bridgeError("IMAGE_FETCH_FAILED", "图片导入服务不可用");
      }
      let body;
      try {
        body = await response.json();
      } catch (error) {
        throw bridgeError("IMAGE_FETCH_FAILED", `图片导入服务返回了无效响应：${response.status}`);
      }
      if (!response.ok || !body || body.ok === false) {
        const serviceError = body && body.error;
        throw bridgeError(
          serviceError && serviceError.code || "IMAGE_FETCH_FAILED",
          serviceError && serviceError.message || `图片导入失败：${response.status}`,
          { details: serviceError && serviceError.details },
        );
      }
      asset = body.asset;
    }
    const configuredAssetPrefix = `${imageBaseURL}/`;
    if (
      !asset
      || typeof asset.path !== "string"
      || !asset.path.startsWith(configuredAssetPrefix)
      || !/\/[0-9a-f]{64}\.(?:png|jpg|gif|webp|svg)(?:\?token=|$)/.test(asset.path)
      || (source.type === "relayAsset" && (typeof asset.assetToken !== "string" || !asset.assetToken))
      || (source.type === "relayAsset" && (typeof asset.assetId !== "string" || !/^[0-9a-f]{64}\.(?:png|jpg|gif|webp|svg)$/.test(asset.assetId) || !asset.path.endsWith(`/${asset.assetId}`)))
      || (asset.assetToken !== undefined && (typeof asset.assetToken !== "string" || !asset.assetToken || asset.assetToken.length > 4096))
      || typeof asset.mimeType !== "string"
      || !/^(?:image\/(?:png|jpeg|gif|webp)|image\/svg\+xml)$/.test(asset.mimeType)
      || !Number.isFinite(Number(asset.widthPx))
      || !Number.isFinite(Number(asset.heightPx))
      || Number(asset.widthPx) <= 0
      || Number(asset.heightPx) <= 0
    ) {
      throw bridgeError("IMAGE_FETCH_FAILED", "图片导入服务缺少有效资源信息");
    }
    let resolvedUrl = new URL(asset.path, window.location.origin).toString();
    let transport = "url";
    const loopbackHost = (
      window.location.hostname === "localhost"
      || window.location.hostname === "127.0.0.1"
      || window.location.hostname === "::1"
      || window.location.hostname === "[::1]"
    );
    if (loopbackHost || asset.assetToken) {
      let assetResponse;
      try {
        assetResponse = await window.fetch(asset.path, {
          method: "GET",
          credentials: "same-origin",
          headers: {
            ...httpRelayHeaders(),
            ...(asset.assetToken ? {"X-AI-Asset-Token": String(asset.assetToken)} : {}),
          },
        });
      } catch (error) {
        throw bridgeError("IMAGE_FETCH_FAILED", "无法读取已导入的图片资源");
      }
      if (!assetResponse.ok || typeof assetResponse.arrayBuffer !== "function") {
        throw bridgeError(
          assetResponse.status === 410 ? "IMAGE_ASSET_EXPIRED" : "IMAGE_FETCH_FAILED",
          "无法读取已导入的图片资源",
        );
      }
      let contentType;
      let bytes;
      try {
        contentType = String(assetResponse.headers && assetResponse.headers.get("Content-Type") || "")
          .split(";", 1)[0]
          .toLowerCase();
        bytes = new Uint8Array(await assetResponse.arrayBuffer());
      } catch (error) {
        throw bridgeError("IMAGE_FETCH_FAILED", "无法读取已导入的图片资源");
      }
      if (contentType !== String(asset.mimeType || "").toLowerCase()) {
        throw bridgeError("UNSUPPORTED_IMAGE_FORMAT", "已导入图片的响应类型不一致");
      }
      if (!bytes.length || bytes.length > 8 * 1024 * 1024) {
        throw bridgeError("IMAGE_TOO_LARGE", "已导入图片超过 8 MiB 限制");
      }
      if (typeof window.btoa !== "function") {
        throw bridgeError("IMAGE_API_UNSUPPORTED", "当前宿主页不能编码内部图片资源");
      }
      let binary = "";
      for (let offset = 0; offset < bytes.length; offset += 32768) {
        binary += String.fromCharCode.apply(null, bytes.subarray(offset, offset + 32768));
      }
      resolvedUrl = `data:${contentType};base64,${window.btoa(binary)}`;
      transport = "dataUrl";
    }
    return {
      url: resolvedUrl,
      transport,
      assetId: String(asset.assetId || ""),
      mimeType: String(asset.mimeType || ""),
      widthPx: Number(asset.widthPx),
      heightPx: Number(asset.heightPx),
    };
  }

  async function prepareImageCommand(command) {
    if (!command || typeof command !== "object") return command;
    const rawCalls = Array.isArray(command.toolCalls)
      ? command.toolCalls
      : command.name ? [command] : null;
    if (!rawCalls || !rawCalls.some(function (call) {
      return call && (IMAGE_TOOL_NAMES.has(call.name) || OPTIONAL_IMAGE_TOOL_NAMES.has(call.name));
    })) {
      return command;
    }
    const preparedCalls = await Promise.all(rawCalls.map(async function (call) {
      if (!call || (!IMAGE_TOOL_NAMES.has(call.name) && !OPTIONAL_IMAGE_TOOL_NAMES.has(call.name))) return call;
      let args = call.arguments === undefined ? call.args : call.arguments;
      if (typeof args === "string") {
        try {
          args = JSON.parse(args || "{}");
        } catch (error) {
          throw bridgeError("INVALID_ARGUMENTS", `${call.name}.arguments 不是有效 JSON`);
        }
      }
      if (!args || typeof args !== "object" || Array.isArray(args)) {
        throw bridgeError("INVALID_ARGUMENTS", `${call.name}.arguments 必须是对象`);
      }
      if (args._image !== undefined) {
        throw bridgeError("INVALID_IMAGE_SOURCE", "外部调用不能提供内部图片资源");
      }
      if (OPTIONAL_IMAGE_TOOL_NAMES.has(call.name) && args.source === undefined) return call;
      const imported = await importImageSource(args.source);
      const preparedArgs = { ...args, _image: imported };
      delete preparedArgs.source;
      return { ...call, arguments: preparedArgs };
    }));
    return { toolCalls: preparedCalls };
  }

  function execute(command, options) {
    const requestOptions = { ...(options || {}) };
    if (!requestOptions.requestId) requestOptions.requestId = createRequestId("command");
    if (typeof requestOptions.requestId !== "string" || !REQUEST_ID_PATTERN.test(requestOptions.requestId)) {
      return Promise.reject(bridgeError(
        "INVALID_REQUEST_ID",
        "requestId 只能包含字母、数字、点、下划线、冒号或连字符，且长度不超过 200",
      ));
    }
    const containsImage = Boolean(
      command
      && typeof command === "object"
      && (
        IMAGE_TOOL_NAMES.has(command.name)
        || OPTIONAL_IMAGE_TOOL_NAMES.has(command.name)
        || (Array.isArray(command.toolCalls)
          && command.toolCalls.some(function (call) {
            return call && (IMAGE_TOOL_NAMES.has(call.name) || OPTIONAL_IMAGE_TOOL_NAMES.has(call.name));
          }))
      )
    );
    if (!containsImage) return request("execute", { command }, requestOptions);

    const fingerprint = stableHash(JSON.stringify(command));
    const cached = preparedImageCommands.get(requestOptions.requestId);
    if (cached && cached.fingerprint !== fingerprint) {
      return Promise.reject(bridgeError(
        "REQUEST_ID_CONFLICT",
        "相同 requestId 不能用于不同的图片命令",
        {
          requestId: requestOptions.requestId,
          details: {
            requestId: requestOptions.requestId,
            retryable: false,
            reuseAllowed: false,
            requiredAction: "use_new_request_id",
            partialMutationPossible: false,
          },
        },
      ));
    }
    const prepared = cached || {
      fingerprint,
      promise: prepareImageCommand(command),
    };
    if (!cached) {
      preparedImageCommands.set(requestOptions.requestId, prepared);
      while (preparedImageCommands.size > 100) {
        preparedImageCommands.delete(preparedImageCommands.keys().next().value);
      }
      prepared.promise.catch(function () {
        if (preparedImageCommands.get(requestOptions.requestId) === prepared) {
          preparedImageCommands.delete(requestOptions.requestId);
        }
      });
    }
    return prepared.promise.then(function (preparedCommand) {
      return request("execute", { command: preparedCommand }, requestOptions);
    });
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
    const values = configuredOptions.clientOrigins;
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
      addImage: function (args, options) { return executeTool("word_add_image", args, options); },
      inspectAdvanced: function (args, options) { return executeTool("word_inspect_advanced", args, options); },
      setDocumentProperties: function (args, options) { return executeTool("word_set_document_properties", args, options); },
      manageSection: function (args, options) { return executeTool("word_manage_section", args, options); },
      manageStyle: function (args, options) { return executeTool("word_manage_style", args, options); },
      setTabs: function (args, options) { return executeTool("word_set_tabs", args, options); },
      setNumbering: function (args, options) { return executeTool("word_set_numbering", args, options); },
      formatTableAdvanced: function (args, options) { return executeTool("word_format_table_advanced", args, options); },
      addNestedTable: function (args, options) { return executeTool("word_add_nested_table", args, options); },
      manageDrawing: function (args, options) { return executeTool("word_manage_drawing", args, options); },
      addShape: function (args, options) { return executeTool("word_add_shape", args, options); },
      addChart: function (args, options) { return executeTool("word_add_chart", args, options); },
      addMath: function (args, options) { return executeTool("word_add_math", args, options); },
      addOleObject: function (args, options) { return executeTool("word_add_ole_object", args, options); },
      manageFields: function (args, options) { return executeTool("word_manage_fields", args, options); },
      manageLongDocument: function (args, options) { return executeTool("word_manage_long_document", args, options); },
      manageComments: function (args, options) { return executeTool("word_manage_comments", args, options); },
      manageRevisions: function (args, options) { return executeTool("word_manage_revisions", args, options); },
      setProtection: function (args, options) { return executeTool("word_set_protection", args, options); },
      manageContentControl: function (args, options) { return executeTool("word_manage_content_control", args, options); },
      manageCustomXml: function (args, options) { return executeTool("word_manage_custom_xml", args, options); },
      inspectMacros: function (args, options) { return executeTool("word_inspect_macros", args, options); },
      setMacros: function (args, options) { return executeTool("word_set_macros", args, options); },
      setWatermark: function (args, options) { return executeTool("word_set_watermark", args, options); },
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
      inspectLayouts: function (args, options) { return executeTool("slides_inspect_layouts", args, options); },
      inspectThemes: function (args, options) { return executeTool("slides_inspect_themes", args, options); },
      inspectBuiltinThemes: function (args, options) { return executeTool("slides_inspect_builtin_themes", args, options); },
      inspectObjects: function (args, options) { return executeTool("slides_inspect_objects", args, options); },
      validateLayout: function (args, options) { return executeTool("slides_validate_layout", args, options); },
      replaceText: function (args, options) { return executeTool("slides_replace_text", args, options); },
      scaleFont: function (args, options) { return executeTool("slides_scale_font", args, options); },
      formatText: function (args, options) { return executeTool("slides_format_text", args, options); },
      formatSelection: function (args, options) { return executeTool("slides_format_selection", args, options); },
      addSlide: function (args, options) { return executeTool("slides_add_slide", args, options); },
      duplicateSlide: function (args, options) { return executeTool("slides_duplicate_slide", args, options); },
      deleteSlide: function (args, options) { return executeTool("slides_delete_slide", args, options); },
      moveSlide: function (args, options) { return executeTool("slides_move_slide", args, options); },
      setVisibility: function (args, options) { return executeTool("slides_set_visibility", args, options); },
      setSize: function (args, options) { return executeTool("slides_set_size", args, options); },
      applyLayout: function (args, options) { return executeTool("slides_apply_layout", args, options); },
      setShowSettings: function (args, options) { return executeTool("slides_set_show_settings", args, options); },
      applyTheme: function (args, options) { return executeTool("slides_apply_theme", args, options); },
      applyBuiltinTheme: function (args, options) { return executeTool("slides_apply_builtin_theme", args, options); },
      setTheme: function (args, options) { return executeTool("slides_set_theme", args, options); },
      createLayout: function (args, options) { return executeTool("slides_create_layout", args, options); },
      addTemplateShape: function (args, options) { return executeTool("slides_add_template_shape", args, options); },
      manageTemplateObject: function (args, options) { return executeTool("slides_manage_template_object", args, options); },
      setTemplateBackground: function (args, options) { return executeTool("slides_set_template_background", args, options); },
      setTextContent: function (args, options) { return executeTool("slides_set_text_content", args, options); },
      formatParagraphs: function (args, options) { return executeTool("slides_format_paragraphs", args, options); },
      updateObject: function (args, options) { return executeTool("slides_update_object", args, options); },
      setHyperlink: function (args, options) { return executeTool("slides_set_hyperlink", args, options); },
      setNotes: function (args, options) { return executeTool("slides_set_notes", args, options); },
      addComment: function (args, options) { return executeTool("slides_add_comment", args, options); },
      inspectComments: function (args, options) { return executeTool("slides_inspect_comments", args, options); },
      manageComment: function (args, options) { return executeTool("slides_manage_comment", args, options); },
      setTransition: function (args, options) { return executeTool("slides_set_transition", args, options); },
      inspectAnimations: function (args, options) { return executeTool("slides_inspect_animations", args, options); },
      manageAnimation: function (args, options) { return executeTool("slides_manage_animation", args, options); },
      addTable: function (args, options) { return executeTool("slides_add_table", args, options); },
      setTableCell: function (args, options) { return executeTool("slides_set_table_cell", args, options); },
      editTable: function (args, options) { return executeTool("slides_edit_table", args, options); },
      formatTable: function (args, options) { return executeTool("slides_format_table", args, options); },
      alignObjects: function (args, options) { return executeTool("slides_align_objects", args, options); },
      groupObjects: function (args, options) { return executeTool("slides_group_objects", args, options); },
      reorderObject: function (args, options) { return executeTool("slides_reorder_object", args, options); },
      addConnector: function (args, options) { return executeTool("slides_add_connector", args, options); },
      addFreeform: function (args, options) { return executeTool("slides_add_freeform", args, options); },
      addTextBox: function (args, options) { return executeTool("slides_add_textbox", args, options); },
      addWordArt: function (args, options) { return executeTool("slides_add_word_art", args, options); },
      addMath: function (args, options) { return executeTool("slides_add_math", args, options); },
      addImage: function (args, options) { return executeTool("slides_add_image", args, options); },
      addImageShape: function (args, options) { return executeTool("slides_add_image_shape", args, options); },
      addOleObject: function (args, options) { return executeTool("slides_add_ole_object", args, options); },
      setBackground: function (args, options) { return executeTool("slides_set_background", args, options); },
      addShape: function (args, options) { return executeTool("slides_add_shape", args, options); },
      updateShape: function (args, options) { return executeTool("slides_update_shape", args, options); },
      deleteObject: function (args, options) { return executeTool("slides_delete_object", args, options); },
      inspectSmartArts: function (args, options) { return executeTool("slides_inspect_smartarts", args, options); },
      addSmartArt: function (args, options) { return executeTool("slides_add_smartart", args, options); },
      updateSmartArt: function (args, options) { return executeTool("slides_update_smartart", args, options); },
      deleteSmartArt: function (args, options) { return executeTool("slides_delete_smartart", args, options); },
      inspectCharts: function (args, options) { return executeTool("slides_inspect_charts", args, options); },
      addChart: function (args, options) { return executeTool("slides_add_chart", args, options); },
      updateChart: function (args, options) { return executeTool("slides_update_chart", args, options); },
      deleteChart: function (args, options) { return executeTool("slides_delete_chart", args, options); },
      inspectMacros: function (args, options) { return executeTool("slides_inspect_macros", args, options); },
      setMacros: function (args, options) { return executeTool("slides_set_macros", args, options); },
      controlSlideshow: function (args, options) { return executeTool("slides_control_slideshow", args, options); },
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
      inspectCharts: function (args, options) { return executeTool("sheets_inspect_charts", args, options); },
      updateChart: function (args, options) { return executeTool("sheets_update_chart", args, options); },
      deleteChart: function (args, options) { return executeTool("sheets_delete_chart", args, options); },
      inspectRange: function (args, options) { return executeTool("sheets_inspect_range", args, options); },
      setArrayFormula: function (args, options) { return executeTool("sheets_set_array_formula", args, options); },
      manageSheet: function (args, options) { return executeTool("sheets_manage_sheet", args, options); },
      manageRange: function (args, options) { return executeTool("sheets_manage_range", args, options); },
      setRichText: function (args, options) { return executeTool("sheets_set_rich_text", args, options); },
      inspectNames: function (args, options) { return executeTool("sheets_inspect_names", args, options); },
      manageNames: function (args, options) { return executeTool("sheets_manage_names", args, options); },
      recalculate: function (args, options) { return executeTool("sheets_recalculate", args, options); },
      sort: function (args, options) { return executeTool("sheets_sort", args, options); },
      filter: function (args, options) { return executeTool("sheets_filter", args, options); },
      inspectTables: function (args, options) { return executeTool("sheets_inspect_tables", args, options); },
      manageTable: function (args, options) { return executeTool("sheets_manage_table", args, options); },
      manageConditionalFormat: function (args, options) { return executeTool("sheets_manage_conditional_format", args, options); },
      manageValidation: function (args, options) { return executeTool("sheets_manage_validation", args, options); },
      inspectPivots: function (args, options) { return executeTool("sheets_inspect_pivots", args, options); },
      managePivot: function (args, options) { return executeTool("sheets_manage_pivot", args, options); },
      inspectDrawings: function (args, options) { return executeTool("sheets_inspect_drawings", args, options); },
      manageDrawing: function (args, options) { return executeTool("sheets_manage_drawing", args, options); },
      manageHyperlink: function (args, options) { return executeTool("sheets_manage_hyperlink", args, options); },
      inspectComments: function (args, options) { return executeTool("sheets_inspect_comments", args, options); },
      manageComments: function (args, options) { return executeTool("sheets_manage_comments", args, options); },
      inspectFreezePanes: function (args, options) { return executeTool("sheets_inspect_freeze_panes", args, options); },
      manageFreezePanes: function (args, options) { return executeTool("sheets_manage_freeze_panes", args, options); },
      inspectProperties: function (args, options) { return executeTool("sheets_inspect_properties", args, options); },
      manageProperties: function (args, options) { return executeTool("sheets_manage_properties", args, options); },
      inspectProtectedRanges: function (args, options) { return executeTool("sheets_inspect_protected_ranges", args, options); },
      manageProtectedRanges: function (args, options) { return executeTool("sheets_manage_protected_ranges", args, options); },
      inspectPageLayout: function (args, options) { return executeTool("sheets_inspect_page_layout", args, options); },
      managePageLayout: function (args, options) { return executeTool("sheets_manage_page_layout", args, options); },
      inspectMacros: function (args, options) { return executeTool("sheets_inspect_macros", args, options); },
      setMacros: function (args, options) { return executeTool("sheets_set_macros", args, options); },
    },
  };

  installEditorStartupInstrumentation();

  window.aiBridge = api;
  window.onlyofficeAI = api;
  window.AiBridgeError = window.AiBridgeError || AiBridgeError;

  function httpRelayEnabled() {
    return configuredOptions.httpRelay === true && Boolean(httpRelayBaseURL());
  }

  function configuredServiceBaseURL(value) {
    if (typeof value !== "string" || !value) return "";
    try {
      const resolved = new URL(value, window.location.href);
      if (resolved.origin !== window.location.origin || resolved.search || resolved.hash) return "";
      return resolved.pathname.replace(/\/$/, "");
    } catch (error) {
      return "";
    }
  }

  function httpRelayBaseURL() {
    return configuredServiceBaseURL(configuredOptions.relayBaseUrl);
  }

  function imageRelayBaseURL() {
    return configuredServiceBaseURL(configuredOptions.imageBaseUrl);
  }

  function persistenceBaseURL() {
    return configuredServiceBaseURL(configuredOptions.persistenceBaseUrl);
  }

  function httpRelayHeaders() {
    const headers = { "Content-Type": "application/json" };
    if (configuredOptions.editorSessionId) {
      headers["X-Editor-Session-ID"] = String(configuredOptions.editorSessionId);
    }
    const csrf = String(document.cookie || "").split("; ").find(function (item) {
      return item.indexOf("document_hub_csrf=") === 0;
    });
    if (csrf) {
      headers["X-CSRF-Token"] = decodeURIComponent(csrf.split("=").slice(1).join("="));
    }
    return headers;
  }

  async function persistenceRequest(action, requestId, payload) {
    const baseURL = persistenceBaseURL();
    const allowed = new Set(["checkpoint", "forcesave", "history", "undo", "redo"]);
    if (!baseURL || !allowed.has(action)) {
      throw bridgeError("PERSISTENCE_NOT_AVAILABLE", "当前页面没有可用的文档持久化服务");
    }
    const response = await window.fetch(`${baseURL}/${encodeURIComponent(action)}`, {
      method: "POST",
      credentials: "same-origin",
      headers: httpRelayHeaders(),
      body: JSON.stringify({ requestId, payload: payload || {} }),
    });
    let body;
    try {
      body = await response.json();
    } catch (error) {
      throw bridgeError("PERSISTENCE_INVALID_RESPONSE", `文档持久化服务返回了无效响应：${response.status}`);
    }
    if (!response.ok || !body || body.ok === false || body.error) {
      const error = body && body.error || {};
      throw bridgeError(
        error.code || "PERSISTENCE_FAILED",
        error.message || `文档持久化请求失败：${response.status}`,
        {
          details: {
            ...(error.details && typeof error.details === "object" ? error.details : {}),
            path: action,
          },
        },
      );
    }
    return body;
  }

  async function handlePluginService(message) {
    const requestId = message.serviceRequestId;
    if (
      message.service !== "persistence"
      || typeof requestId !== "string"
      || !REQUEST_ID_PATTERN.test(requestId)
    ) {
      return;
    }
    try {
      if (message.action === "forcesave") {
        saveStatus = "saving";
        saveStatusAt = Date.now();
      }
      const result = await persistenceRequest(message.action, requestId, message.payload);
      if (message.action === "forcesave") {
        saveStatus = "saved";
        saveStatusAt = Date.now();
      }
      send({ type: "service-response", serviceRequestId: requestId, result });
    } catch (error) {
      if (message.action === "forcesave") {
        saveStatus = "failed";
        saveStatusAt = Date.now();
      }
      send({
        type: "service-response",
        serviceRequestId: requestId,
        error: serializedError(error, requestId),
      });
    }
  }

  async function httpRelayPost(path, payload) {
    let response;
    try {
      response = await window.fetch(`${httpRelayBaseURL()}/${path}`, {
        method: "POST",
        credentials: "same-origin",
        headers: httpRelayHeaders(),
        body: JSON.stringify(payload),
      });
    } catch (error) {
      throw bridgeError(
        "HTTP_RELAY_NETWORK_ERROR",
        `HTTP Relay 网络请求失败：${error && error.message ? error.message : String(error)}`,
        { details: { path } },
      );
    }
    let body;
    try {
      body = await response.json();
    } catch (error) {
      throw bridgeError(
        "HTTP_RELAY_INVALID_RESPONSE",
        `HTTP Relay 返回了无效响应：${response.status}`,
        { details: { httpStatus: response.status, path } },
      );
    }
    if (!response.ok || !body || body.ok === false) {
      const error = body && body.error || {};
      const details = error.details && typeof error.details === "object"
        ? { ...error.details }
        : {};
      details.httpStatus = response.status;
      details.path = path;
      throw bridgeError(
        error.code || "HTTP_RELAY_FAILED",
        error.message || `HTTP Relay 请求失败：${response.status}`,
        { details },
      );
    }
    return body;
  }

  function startHttpRelay() {
    if (!httpRelayEnabled() || typeof window.fetch !== "function") return;
    const httpSessionId = createRequestId("http-session");
    const relayStartupTimeoutMs = 180000;
    const credentialReloadWindowMs = 600000;
    let relayKey = null;
    let resumeToken = null;
    let stopped = false;
    let retryDelayMs = 1000;

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

    function clearCredentialReload() {
      const key = credentialReloadKey();
      try {
        if (window.localStorage) window.localStorage.removeItem(key);
      } catch (error) {
        // A stale guard is harmless when persistent storage is unavailable.
      }
      try {
        if (window.sessionStorage) window.sessionStorage.removeItem(key);
      } catch (error) {
        // A stale guard is harmless when per-tab storage is unavailable.
      }
      try {
        const currentUrl = new URL(window.location.href);
        if (currentUrl.searchParams.has("aiBridgeCredentialReloadAt")) {
          currentUrl.searchParams.delete("aiBridgeCredentialReloadAt");
          window.history.replaceState(null, "", currentUrl.toString());
        }
      } catch (error) {
        // Storage-backed guards were already cleared when available.
      }
    }

    function stopRelay(reason, code, error) {
      stopped = true;
      relayKey = null;
      setRelayState(reason, code, error);
      emit("relayError", {
        reason,
        ...serializedError(error || bridgeError(code || "HTTP_RELAY_STOPPED", reason)),
      });
    }

    function reloadForCredentialError(code, error) {
      const now = Date.now();
      if (now - credentialReloadAt() < credentialReloadWindowMs) {
        stopRelay("credential-error", code, error);
        return;
      }
      rememberCredentialReload(now);
      stopped = true;
      setRelayState("credential-reload");
      emit("reload", { reason: "relay-credentials-expired", code });
      window.setTimeout(function () { window.location.reload(); }, 50);
    }

    async function register() {
      setRelayState("registering");
      try {
        await waitUntilReady(relayStartupTimeoutMs);
      } catch (error) {
        if (!error || error.code !== "NOT_READY") throw error;
        throw bridgeError(
          "PLUGIN_STARTUP_TIMEOUT",
          "ai-bridge 插件启动超时",
          {
            details: {
              timeoutMs: relayStartupTimeoutMs,
              startup: { ...startupMarks },
            },
          },
        );
      }
      const payload = {
        sessionId: httpSessionId,
        state: relayStateSnapshot(),
      };
      if (resumeToken) {
        payload.resumeToken = resumeToken;
      } else {
        const token = editorToken();
        if (!token) throw bridgeError("EDITOR_TOKEN_REQUIRED", "当前编辑器没有可用于 HTTP Relay 注册的 editor JWT");
        payload.editorToken = token;
      }
      const response = await httpRelayPost("register", payload);
      relayKey = response.relayKey;
      resumeToken = response.resumeToken || resumeToken;
      clearCredentialReload();
      setRelayState("ready");
    }

    async function report(command, outcome) {
      await httpRelayPost("result", {
        sessionId: httpSessionId,
        relayKey,
        commandId: command.commandId,
        state: relayStateSnapshot(),
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
            state: relayStateSnapshot(),
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
            stopRelay("superseded", code, error);
            return;
          }
          if (code === "CONTRACT_MISMATCH" || code === "CONTRACT_VERSION_MISMATCH") {
            stopRelay("contract-mismatch", code, error);
            return;
          }
          if (code === "PLUGIN_STARTUP_TIMEOUT") {
            stopRelay("startup-timeout", code, error);
            return;
          }
          if (
            code === "EDITOR_TOKEN_REQUIRED" ||
            code === "INVALID_EDITOR_TOKEN" ||
            code === "EDITOR_TOKEN_EXPIRED" ||
            code === "BRIDGE_RESUME_TOKEN_EXPIRED" ||
            code === "editor_session_invalid" ||
            code === "DOCUMENT_MISMATCH" ||
            code === "EDITOR_MISMATCH" ||
            code === "DOCUMENT_IDENTITY_MISMATCH" ||
            code === "INCOMPLETE_BRIDGE_IDENTITY"
          ) {
            reloadForCredentialError(code, error);
            return;
          }
          setRelayState("retrying", code, error);
          await new Promise(function (resolve) { window.setTimeout(resolve, retryDelayMs); });
          retryDelayMs = Math.min(10000, retryDelayMs * 2);
        }
      }
    }

    window.addEventListener("beforeunload", function () {
      stopped = true;
      if (!relayKey || typeof window.fetch !== "function") return;
      const body = JSON.stringify({ sessionId: httpSessionId, relayKey });
      window.fetch(`${httpRelayBaseURL()}/unregister`, {
        method: "POST",
        credentials: "same-origin",
        headers: httpRelayHeaders(),
        body,
        keepalive: true,
      }).catch(function () {});
    });
    loop();
  }

  startHttpRelay();

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
