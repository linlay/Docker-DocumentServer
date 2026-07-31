const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");

const source = fs.readFileSync(path.join(__dirname, "editor-shell.js"), "utf8");

function encodedConfig(config) {
  return Buffer.from(JSON.stringify(config), "utf8").toString("base64url");
}

test("first-party shell prepares identity before constructing DocEditor", async () => {
  const calls = [];
  const config = {
    document: {key: "document-key"},
    editorConfig: {plugins: {options: {}}},
    token: "signed-token",
  };
  const container = {
    dataset: {editorConfig: encodedConfig(config)},
    children: [],
    firstChild: null,
    appendChild(child) {
      this.children.push(child);
      this.firstChild = this.children[0] || null;
    },
    removeChild(child) {
      this.children.splice(this.children.indexOf(child), 1);
      this.firstChild = this.children[0] || null;
    },
  };
  const currentScript = {
    src: "https://office.test/sdkjs-plugins/plugin/editor-shell.js?v=rev",
  };
  const document = {
    currentScript,
    head: {
      appendChild(script) {
        calls.push(["host", script.src]);
        queueMicrotask(script.onload);
      },
    },
    getElementById(id) {
      return id === "iframeEditor" ? container : null;
    },
    createElement(tagName) {
      return {
        tagName,
        children: [],
        style: {},
        appendChild(child) {
          this.children.push(child);
        },
        addEventListener() {},
      };
    },
  };
  const window = {
    atob(value) {
      return Buffer.from(value, "base64").toString("binary");
    },
    console: {error() {}},
    document,
    location: {reload() {}},
    aiBridgeOptions: {clientOrigins: ["https://client.test"]},
    DocsAPI: {
      DocEditor: function (id, prepared) {
        calls.push(["editor", id, prepared.editorConfig.user.id]);
      },
    },
    OnlyOfficeLocalGuest: {
      async prepareEditor(prepared) {
        calls.push(["guest", prepared.document.key]);
        prepared.editorConfig.user = {id: "local-guest:test"};
        return true;
      },
    },
  };

  vm.runInNewContext(source, {
    Buffer,
    Error,
    JSON,
    Promise,
    TextDecoder,
    Uint8Array,
    URL,
    window,
  });
  await new Promise((resolve) => setImmediate(resolve));

  assert.equal(window.config.document.key, "document-key");
  assert.equal(window.aiBridgeOptions.httpRelay, true);
  assert.equal(window.aiBridgeOptions.clientOrigins[0], "https://client.test");
  assert.match(calls[0][1], /host-bridge\.js\?v=rev$/);
  assert.deepEqual(calls.slice(1), [
    ["guest", "document-key"],
    ["editor", "iframeEditor", "local-guest:test"],
  ]);
});
