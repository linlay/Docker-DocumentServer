(function (root) {
  "use strict";

  var STORAGE_KEY = "onlyoffice.localGuestId.v1";
  var DEFAULT_ENDPOINT = "/copilot-api/editor-config/anonymous";
  var DISPLAY_NAME = "访客";
  var USER_ID_PREFIX = "local-guest:";
  var UUID_V4_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/;

  function localGuestError(code, message) {
    var error = new Error(message);
    error.code = code;
    return error;
  }

  function formatUuid(bytes) {
    bytes[6] = (bytes[6] & 15) | 64;
    bytes[8] = (bytes[8] & 63) | 128;
    var hex = Array.prototype.map.call(bytes, function (value) {
      return value.toString(16).padStart(2, "0");
    }).join("");
    return [
      hex.slice(0, 8),
      hex.slice(8, 12),
      hex.slice(12, 16),
      hex.slice(16, 20),
      hex.slice(20),
    ].join("-");
  }

  function generateUuid() {
    if (!root.crypto) {
      throw localGuestError("LOCAL_GUEST_CRYPTO_UNAVAILABLE", "浏览器不支持安全随机访客身份");
    }
    if (typeof root.crypto.randomUUID === "function") {
      return root.crypto.randomUUID().toLowerCase();
    }
    if (typeof root.crypto.getRandomValues === "function") {
      return formatUuid(root.crypto.getRandomValues(new Uint8Array(16)));
    }
    throw localGuestError("LOCAL_GUEST_CRYPTO_UNAVAILABLE", "浏览器不支持安全随机访客身份");
  }

  function getOrCreateAnonymousId(options) {
    options = options || {};
    var storage = options.storage || root.localStorage;
    if (!storage || typeof storage.getItem !== "function" || typeof storage.setItem !== "function") {
      throw localGuestError("LOCAL_GUEST_STORAGE_UNAVAILABLE", "浏览器无法保存稳定访客身份");
    }
    var storageKey = options.storageKey || STORAGE_KEY;
    var existing;
    try {
      existing = String(storage.getItem(storageKey) || "").trim().toLowerCase();
    } catch (error) {
      throw localGuestError("LOCAL_GUEST_STORAGE_UNAVAILABLE", "浏览器无法读取稳定访客身份");
    }
    if (UUID_V4_PATTERN.test(existing)) return existing;
    var generated = generateUuid();
    if (!UUID_V4_PATTERN.test(generated)) {
      throw localGuestError("LOCAL_GUEST_UUID_INVALID", "浏览器生成了无效的访客身份");
    }
    try {
      storage.setItem(storageKey, generated);
    } catch (error) {
      throw localGuestError("LOCAL_GUEST_STORAGE_UNAVAILABLE", "浏览器无法保存稳定访客身份");
    }
    return generated;
  }

  function responseError(payload, status) {
    var error = payload && payload.error;
    var code = error && typeof error === "object" ? error.code : "LOCAL_GUEST_REISSUE_FAILED";
    var message = error && typeof error === "object" && error.message
      ? error.message
      : "无法初始化本地访客身份（HTTP " + status + "）";
    return localGuestError(code, message);
  }

  async function prepare(config, options) {
    options = options || {};
    if (!config || typeof config !== "object") {
      throw localGuestError("LOCAL_GUEST_CONFIG_REQUIRED", "缺少 ONLYOFFICE 编辑器配置");
    }
    if (typeof config.token !== "string" || !config.token) {
      throw localGuestError("LOCAL_GUEST_TOKEN_REQUIRED", "缺少 ONLYOFFICE 编辑器签名");
    }
    var anonymousId = options.anonymousId || getOrCreateAnonymousId(options);
    if (!UUID_V4_PATTERN.test(String(anonymousId).toLowerCase())) {
      throw localGuestError("LOCAL_GUEST_UUID_INVALID", "anonymousId 必须是 UUID v4");
    }
    anonymousId = String(anonymousId).toLowerCase();
    var request = options.fetch || root.fetch;
    if (typeof request !== "function") {
      throw localGuestError("LOCAL_GUEST_FETCH_UNAVAILABLE", "浏览器无法请求访客签名");
    }
    var response;
    try {
      response = await request(options.endpoint || DEFAULT_ENDPOINT, {
        method: "POST",
        credentials: "same-origin",
        cache: "no-store",
        headers: {"Content-Type": "application/json"},
        body: JSON.stringify({
          editorToken: config.token,
          anonymousId: anonymousId,
        }),
      });
    } catch (error) {
      throw localGuestError("LOCAL_GUEST_REISSUE_FAILED", "无法连接本地访客签名服务");
    }
    var payload;
    try {
      payload = await response.json();
    } catch (error) {
      throw localGuestError("LOCAL_GUEST_RESPONSE_INVALID", "本地访客签名服务返回了无效响应");
    }
    if (!response.ok || !payload || payload.ok !== true) {
      throw responseError(payload, response.status);
    }
    var expectedUserId = USER_ID_PREFIX + anonymousId;
    if (
      !payload.user
      || payload.user.id !== expectedUserId
      || payload.user.name !== DISPLAY_NAME
      || typeof payload.token !== "string"
      || !payload.token
    ) {
      throw localGuestError("LOCAL_GUEST_RESPONSE_INVALID", "本地访客签名响应与请求身份不一致");
    }
    config.editorConfig = config.editorConfig || {};
    config.editorConfig.user = payload.user;
    config.editorConfig.customization = config.editorConfig.customization || {};
    config.editorConfig.customization.anonymous = {
      request: false,
      label: DISPLAY_NAME,
    };
    config.token = payload.token;
    return {
      anonymousId: anonymousId,
      user: payload.user,
      expiresAt: payload.expiresAt,
    };
  }

  function isDefaultExampleGuest(locationValue) {
    var href = locationValue && locationValue.href
      ? locationValue.href
      : String(locationValue || root.location.href);
    var userid = new URL(href).searchParams.get("userid");
    return !userid || userid === "uid-0";
  }

  function renderFailure(error) {
    if (!root.document) return;
    var container = root.document.getElementById("iframeEditor") || root.document.body;
    if (!container || typeof root.document.createElement !== "function") return;
    while (container.firstChild) container.removeChild(container.firstChild);
    var panel = root.document.createElement("div");
    var title = root.document.createElement("strong");
    var detail = root.document.createElement("p");
    var retry = root.document.createElement("button");
    panel.style.cssText = "box-sizing:border-box;max-width:520px;margin:12vh auto;padding:24px;border:1px solid #ddd;border-radius:8px;font:14px/1.5 Arial,sans-serif;color:#333;background:#fff;";
    title.textContent = "访客身份初始化失败";
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

  async function prepareExample(config, options) {
    if (!isDefaultExampleGuest(root.location)) return true;
    try {
      await prepare(config, options);
      return true;
    } catch (error) {
      if (root.console && typeof root.console.error === "function") {
        root.console.error("[local-guest] initialization failed", error);
      }
      renderFailure(error);
      return false;
    }
  }

  root.OnlyOfficeLocalGuest = Object.freeze({
    STORAGE_KEY: STORAGE_KEY,
    getOrCreateAnonymousId: getOrCreateAnonymousId,
    isDefaultExampleGuest: isDefaultExampleGuest,
    prepare: prepare,
    prepareExample: prepareExample,
  });
})(window);
