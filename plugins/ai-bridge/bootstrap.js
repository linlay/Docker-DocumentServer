(function () {
  "use strict";

  const PLUGIN_GUID = "asc.{A17E5F31-64AA-4E37-9A42-8D430814C2F6}";
  const PROTOCOL_VERSION = 1;
  const CHANNEL_ID_PATTERN = /^[A-Za-z0-9._:-]{16,200}$/;
  const currentScript = document.currentScript;
  const revision = currentScript
    ? new URL(currentScript.src, window.location.href).searchParams.get("v") || ""
    : "";
  const state = window.__aiBridgeBootstrapState = {
    loaded: [],
    failures: [],
    initRequested: false,
    initThis: null,
    initArguments: null,
  };

  function pluginOptions() {
    return window.Asc
      && window.Asc.plugin
      && window.Asc.plugin.info
      && window.Asc.plugin.info.options
      && typeof window.Asc.plugin.info.options === "object"
      ? window.Asc.plugin.info.options
      : {};
  }

  function ancestorWindows() {
    const ancestors = [];
    let current = window.parent;
    while (current && current !== window && !ancestors.includes(current)) {
      ancestors.push(current);
      if (current === window.top) break;
      try {
        current = current.parent;
      } catch (error) {
        break;
      }
    }
    if (!ancestors.length && window.top && window.top !== window) ancestors.push(window.top);
    return ancestors;
  }

  function resolveTopOrigin() {
    const origins = window.location.ancestorOrigins;
    if (origins && origins.length) return origins[origins.length - 1];
    try {
      return window.top.location.origin;
    } catch (error) {
      return null;
    }
  }

  function announceAssetFailure(failure) {
    const options = pluginOptions();
    const strict = (
      Object.prototype.hasOwnProperty.call(options, "hostOrigin")
      || Object.prototype.hasOwnProperty.call(options, "channelId")
    );
    const hostOrigin = typeof options.hostOrigin === "string" ? options.hostOrigin : "";
    const channelId = typeof options.channelId === "string" ? options.channelId : "";
    let normalizedOrigin = "";
    try {
      normalizedOrigin = new URL(hostOrigin).origin;
    } catch (error) {
      // The host bridge reports invalid strict configuration itself.
    }
    const payload = {
      source: "ai-bridge-plugin",
      protocolVersion: PROTOCOL_VERSION,
      pluginGuid: PLUGIN_GUID,
      type: "startup-error",
      error: {
        code: "PLUGIN_ASSET_LOAD_FAILED",
        message: "ai-bridge 浏览器资源加载失败",
        details: { asset: failure.asset, reason: failure.reason },
      },
    };
    if (strict) {
      if (normalizedOrigin !== hostOrigin || !CHANNEL_ID_PATTERN.test(channelId)) return;
      payload.channelId = channelId;
      for (const ancestor of ancestorWindows()) ancestor.postMessage(payload, hostOrigin);
      return;
    }
    window.top.postMessage(payload, resolveTopOrigin() || "*");
  }

  function installPluginFailureReporter(failure) {
    if (!window.Asc || !window.Asc.plugin) return;
    window.Asc.plugin.init = function () {
      announceAssetFailure(failure);
      window.setInterval(function () { announceAssetFailure(failure); }, 500);
    };
    if (state.initRequested) {
      window.Asc.plugin.init.apply(state.initThis, state.initArguments || []);
    }
  }

  function loadAsset(asset) {
    return new Promise(function (resolve) {
      const script = document.createElement("script");
      const url = new URL(asset, currentScript ? currentScript.src : window.location.href);
      if (revision) url.searchParams.set("v", revision);
      let evaluationFailed = false;
      function recordFailure(reason) {
        const existing = state.failures.find(function (failure) { return failure.asset === asset; });
        const failure = existing || { asset, reason };
        if (!existing) state.failures.push(failure);
        if (asset === "plugin.js") installPluginFailureReporter(failure);
        return failure;
      }
      function onRuntimeError(event) {
        if (!event || !event.filename) return;
        try {
          if (new URL(event.filename, window.location.href).pathname !== url.pathname) return;
        } catch (error) {
          return;
        }
        evaluationFailed = true;
        recordFailure("evaluation-error");
      }
      function stopObservingRuntimeErrors() {
        if (typeof window.removeEventListener === "function") {
          window.removeEventListener("error", onRuntimeError);
        }
      }
      if (typeof window.addEventListener === "function") {
        window.addEventListener("error", onRuntimeError);
      }
      script.src = url.toString();
      script.async = false;
      script.onload = function () {
        stopObservingRuntimeErrors();
        if (!evaluationFailed) state.loaded.push(asset);
        resolve(!evaluationFailed);
      };
      script.onerror = function () {
        stopObservingRuntimeErrors();
        recordFailure("network-or-policy-error");
        resolve(false);
      };
      try {
        document.head.appendChild(script);
      } catch (error) {
        stopObservingRuntimeErrors();
        recordFailure("dom-insertion-error");
        resolve(false);
      }
    });
  }

  const queuedInit = function () {
    state.initRequested = true;
    state.initThis = this;
    state.initArguments = Array.prototype.slice.call(arguments);
  };
  if (window.Asc && window.Asc.plugin) window.Asc.plugin.init = queuedInit;

  (async function loadRuntime() {
    for (const asset of [
      "bridges/word-bridge.js",
      "bridges/slides-bridge.js",
      "bridges/sheets-bridge.js",
    ]) {
      await loadAsset(asset);
    }
    const pluginLoaded = await loadAsset("plugin.js");
    if (
      pluginLoaded
      && state.initRequested
      && window.Asc
      && window.Asc.plugin
      && typeof window.Asc.plugin.init === "function"
      && window.Asc.plugin.init !== queuedInit
    ) {
      window.Asc.plugin.init.apply(state.initThis, state.initArguments || []);
    }
  }());
}());
