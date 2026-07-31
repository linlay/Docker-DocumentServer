const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");

const source = fs.readFileSync(path.join(__dirname, "local-guest.js"), "utf8");
const firstUuid = "123e4567-e89b-42d3-a456-426614174000";
const secondUuid = "123e4567-e89b-42d3-b456-426614174001";

function storageHarness() {
  const values = new Map();
  return {
    values,
    getItem(key) {
      return values.has(key) ? values.get(key) : null;
    },
    setItem(key, value) {
      values.set(key, String(value));
    },
    clear() {
      values.clear();
    },
  };
}

function documentHarness() {
  function element(tagName) {
    return {
      tagName: tagName.toUpperCase(),
      children: [],
      firstChild: null,
      style: {},
      textContent: "",
      appendChild(child) {
        this.children.push(child);
        this.firstChild = this.children[0] || null;
      },
      removeChild(child) {
        this.children.splice(this.children.indexOf(child), 1);
        this.firstChild = this.children[0] || null;
      },
      addEventListener(name, listener) {
        this.listener = {name, listener};
      },
    };
  }
  const editor = element("div");
  return {
    editor,
    body: element("body"),
    getElementById(id) {
      return id === "iframeEditor" ? editor : null;
    },
    createElement: element,
  };
}

function loadGuest(options = {}) {
  const storage = options.storage || storageHarness();
  const document = options.document || documentHarness();
  let uuidIndex = 0;
  const uuids = options.uuids || [firstUuid];
  const window = {
    console: options.console || {error() {}},
    crypto: {
      randomUUID() {
        return uuids[Math.min(uuidIndex++, uuids.length - 1)];
      },
    },
    document,
    fetch: options.fetch,
    localStorage: storage,
    location: {
      href: options.href || "http://localhost:8088/docx/123e4567-e89b-42d3-a456-426614174000",
      reload() {},
    },
  };
  vm.runInNewContext(source, {window, URL, Uint8Array, Error, JSON, Object, String});
  return {api: window.OnlyOfficeLocalGuest, document, storage, window};
}

test("stable guest UUID persists for the browser profile", () => {
  const harness = loadGuest({uuids: [firstUuid, secondUuid]});

  assert.equal(harness.api.getOrCreateAnonymousId(), firstUuid);
  assert.equal(harness.api.getOrCreateAnonymousId(), firstUuid);
  harness.storage.clear();
  assert.equal(harness.api.getOrCreateAnonymousId(), secondUuid);
});

test("prepare reissues the token and applies the fixed guest identity", async () => {
  let captured;
  const harness = loadGuest({
    fetch: async (url, request) => {
      captured = {url, request};
      return {
        ok: true,
        status: 200,
        async json() {
          return {
            ok: true,
            user: {
              group: "",
              id: `local-guest:${firstUuid}`,
              image: "",
              name: "访客",
              roles: [],
            },
            token: "reissued-token",
            expiresAt: 2000,
          };
        },
      };
    },
  });
  const config = {
    token: "original-token",
    editorConfig: {
      user: {id: "uid-1", name: "John Smith"},
      customization: {compactToolbar: true},
    },
  };

  const result = await harness.api.prepare(config);
  const requestBody = JSON.parse(captured.request.body);

  assert.equal(captured.url, "/copilot-api/editor-config/anonymous");
  assert.deepEqual(requestBody, {
    editorToken: "original-token",
    anonymousId: firstUuid,
  });
  assert.equal(config.editorConfig.user.id, `local-guest:${firstUuid}`);
  assert.equal(config.editorConfig.user.name, "访客");
  assert.equal(config.editorConfig.customization.compactToolbar, true);
  assert.equal(config.editorConfig.customization.anonymous.request, false);
  assert.equal(config.token, "reissued-token");
  assert.equal(result.expiresAt, 2000);
});

test("editor startup always replaces the placeholder identity", async () => {
  let fetchCalls = 0;
  const harness = loadGuest({
    fetch: async () => {
      fetchCalls += 1;
      return {
        ok: true,
        status: 200,
        async json() {
          return {
            ok: true,
            user: {
              group: "",
              id: `local-guest:${firstUuid}`,
              image: "",
              name: "访客",
              roles: [],
            },
            token: "reissued-token",
            expiresAt: 2000,
          };
        },
      };
    },
  });
  const config = {
    token: "original-token",
    editorConfig: {user: {id: "pending-local-guest", name: "访客"}},
  };

  assert.equal(await harness.api.prepareEditor(config), true);
  assert.equal(fetchCalls, 1);
  assert.equal(config.editorConfig.user.id, `local-guest:${firstUuid}`);
});

test("failed guest reissue blocks editor startup and renders retry UI", async () => {
  const harness = loadGuest({
    fetch: async () => ({
      ok: false,
      status: 401,
      async json() {
        return {
          ok: false,
          error: {code: "INVALID_EDITOR_TOKEN", message: "凭证无效"},
        };
      },
    }),
  });
  const config = {
    token: "original-token",
    editorConfig: {user: {id: "uid-1", name: "John Smith"}},
  };

  assert.equal(await harness.api.prepareEditor(config), false);
  assert.equal(config.token, "original-token");
  assert.equal(config.editorConfig.user.name, "John Smith");
  assert.equal(harness.document.editor.children.length, 1);
  assert.equal(harness.document.editor.children[0].children[0].textContent, "访客身份初始化失败");
  assert.equal(harness.document.editor.children[0].children[2].textContent, "重试");
});
