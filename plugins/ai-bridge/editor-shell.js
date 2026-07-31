(function (root) {
  "use strict";

  var currentScript = root.document.currentScript;
  var container = root.document.getElementById("iframeEditor");

  function renderFailure(error) {
    if (!container) return;
    while (container.firstChild) container.removeChild(container.firstChild);
    var panel = root.document.createElement("div");
    var title = root.document.createElement("strong");
    var detail = root.document.createElement("p");
    var retry = root.document.createElement("button");
    panel.style.cssText = "box-sizing:border-box;max-width:560px;margin:12vh auto;padding:24px;border:1px solid #ddd;border-radius:8px;font:14px/1.5 Arial,sans-serif;color:#333;background:#fff;";
    title.textContent = "文档编辑器初始化失败";
    detail.textContent = error && error.message ? error.message : "请重试。";
    retry.type = "button";
    retry.textContent = "重试";
    retry.addEventListener("click", function () {
      root.location.reload();
    });
    panel.appendChild(title);
    panel.appendChild(detail);
    panel.appendChild(retry);
    container.appendChild(panel);
  }

  function decodeConfiguration(encoded) {
    var normalized = String(encoded || "").replace(/-/g, "+").replace(/_/g, "/");
    normalized += "=".repeat((4 - normalized.length % 4) % 4);
    var binary = root.atob(normalized);
    var bytes = new Uint8Array(binary.length);
    for (var index = 0; index < binary.length; index += 1) {
      bytes[index] = binary.charCodeAt(index);
    }
    return JSON.parse(new TextDecoder("utf-8", {fatal: true}).decode(bytes));
  }

  function loadHostBridge() {
    return new Promise(function (resolve, reject) {
      var script = root.document.createElement("script");
      var revision = currentScript ? new URL(currentScript.src).search : "";
      script.src = new URL("host-bridge.js" + revision, currentScript.src).toString();
      script.onload = resolve;
      script.onerror = function () {
        reject(new Error("无法加载文档桥接脚本"));
      };
      root.document.head.appendChild(script);
    });
  }

  function enableHttpRelay() {
    var options = root.aiBridgeOptions;
    var enabled = !(options && options.httpRelay === false);
    root.aiBridgeOptions = Object.assign(
      {},
      options && typeof options === "object" ? options : {},
      {httpRelay: enabled}
    );
  }

  async function start() {
    if (!container) throw new Error("缺少文档编辑器容器");
    if (!currentScript) throw new Error("无法定位编辑器启动脚本");
    if (!root.DocsAPI || typeof root.DocsAPI.DocEditor !== "function") {
      throw new Error("ONLYOFFICE Docs API 未加载");
    }
    var config = decodeConfiguration(container.dataset.editorConfig);
    delete container.dataset.editorConfig;

    // host-bridge must observe the exact signed config before DocEditor starts.
    root.config = config;
    enableHttpRelay();
    await loadHostBridge();
    if (
      !root.OnlyOfficeLocalGuest
      || typeof root.OnlyOfficeLocalGuest.prepareEditor !== "function"
    ) {
      throw new Error("访客身份服务未加载");
    }
    if (!(await root.OnlyOfficeLocalGuest.prepareEditor(config))) return;
    root.docEditor = new root.DocsAPI.DocEditor("iframeEditor", config);
  }

  start().catch(function (error) {
    if (root.console && typeof root.console.error === "function") {
      root.console.error("[editor-shell] startup failed", error);
    }
    renderFailure(error);
  });
})(window);
