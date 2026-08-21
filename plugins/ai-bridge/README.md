# ONLYOFFICE ai-bridge

ai-bridge is a headless ONLYOFFICE plugin that exposes validated Word, Slides,
and Sheets tools. In production, document-hub v1 is the only document and HTTP
business boundary.

## Architecture

- `bootstrap.js` loads the three editor bridges and `plugin.js` in order while
  recording network, policy, DOM insertion, and script evaluation failures.
- `plugin.js` runs inside the ONLYOFFICE plugin frame and dispatches validated
  commands to the editor-specific bridge.
- `host-bridge.js` exposes `window.aiBridge`, supports the postMessage SDK,
  and connects an open editor to document-hub's long-poll Relay.
- `copilot_server.py` is a private Relay and validation service. It does not
  create, store, list, administer, or version documents.
- document-hub owns authentication, documents, callbacks, permissions,
  persistence, image ingress, and all public APIs.

Legacy UUID document routes, public attach/binding endpoints, `/chat`, and
Python persistence endpoints are intentionally disabled.

## Host configuration

Set the host options after the ONLYOFFICE editor configuration is available:

```js
window.aiBridgeOptions = {
  httpRelay: true,
  relayBaseUrl: "/api/v1/editor-relay",
  imageBaseUrl: "/api/v1/editor-relay/images",
  persistenceBaseUrl: "/api/v1/editor-relay/persistence",
  editorSessionId: editorConfig.sessionId,
  documentId: editorConfig.documentId,
  getEditorConfig: () => editorConfig,
  clientOrigins: ["https://app.example.com"],
};
```

Then load the host bridge:

```html
<script src="https://docs.example.com/sdkjs-plugins/{A17E5F31-64AA-4E37-9A42-8D430814C2F6}/host-bridge.js?v=<asset-revision>"></script>
```

HTTP Relay is never enabled implicitly. Relay, image, and persistence base URLs
must be explicit same-origin paths.

## JavaScript API

```js
await window.aiBridge.ready();
const result = await window.aiBridge.word.inspect({ maxChars: 5000 });
await window.aiBridge.executeBatch([
  { name: "manage_revisions", arguments: { action: "start", displayMode: "edit" } },
  { name: "replace_text", arguments: { search: "old", replace: "new" } },
]);
```

`word.manageRevisions` / public `manage_revisions` keeps revision tracking and
review display independent. `start` and `stop` toggle tracking; `setDisplay`
changes only `edit`, `simple`, `final`, or `original` display. When a display
mode is supplied, the bridge waits for `SetDisplayModeInReview` before running
the next tool call. Existing calls without `displayMode` keep their old behavior.

Tool names are unprefixed on public Relay and batch envelopes. Direct
editor-specific methods remain grouped under `word`, `slides`, and `sheets`.
The compatibility alias `window.onlyofficeAI` remains available.

Save, history, undo, and redo require a configured document-hub persistence
service. Without it, the bridge returns `PERSISTENCE_NOT_AVAILABLE`.

## postMessage client

```html
<script src="https://docs.example.com/sdkjs-plugins/{A17E5F31-64AA-4E37-9A42-8D430814C2F6}/client-sdk.js?v=<asset-revision>"></script>
```

```js
const client = new AiBridgeClient({
  targetWindow: editorFrame.contentWindow,
  targetOrigin: "https://docs.example.com",
});
await client.ready();
const state = await client.refreshState();
```

`clientOrigins` must contain exact origins. Wildcards are rejected.

## Contract

`public-api.json` is the source of truth for tool schemas, limits, controls,
errors, transport metadata, and the contract hash. Run the generator after
changing it:

```bash
python3 tools/sync_contract.py --write
python3 tools/sync_contract.py --check
```

Do not edit generated regions in `plugin.js` or `public-api.d.ts` manually.

## Private Relay endpoints

document-hub is the only caller of the internal endpoints:

- browser Relay: `/bridge/register`, `/bridge/poll`, `/bridge/result`,
  `/bridge/unregister`
- document-hub Relay: `/bridge/internal/execute`,
  `/bridge/internal/validate`, `/bridge/internal/session`,
  `/bridge/internal/drop`, `/bridge/internal/images/import`
- diagnostics/assets: `/health`, `/bridge/contract/{editor}`, `/images/*`

Every request requires the private Relay secret. Port 3001 must not be
published.

## Startup diagnostics

The browser host records milestones from host script load through editor app,
document, plugin handshake, configuration, and bridge readiness. A terminal
startup failure includes a `diagnosticId`, `stalledPhase`, the milestone
timeline, plugin registration checks, browser-visible host/plugin origins, the
asset revision, and any failed plugin asset.

Startup failures use specific codes instead of one generic timeout:

- `EDITOR_APP_STARTUP_TIMEOUT` and `DOCUMENT_LOAD_TIMEOUT` identify the
  ONLYOFFICE application or document stage.
- `PLUGIN_NOT_LOADED` and `PLUGIN_INITIALIZATION_FAILED` distinguish an absent
  plugin from a plugin that connected but never became ready.
- `PLUGIN_HANDSHAKE_CONFIG_INVALID` and `PLUGIN_HANDSHAKE_REJECTED` expose
  origin/channel configuration failures, including on remote HTTPS hosts.
- `PLUGIN_ASSET_LOAD_FAILED`, `PLUGIN_BRIDGE_MISSING`, and
  `UNSUPPORTED_EDITOR_TYPE` identify browser resource or editor adapter issues.
- `CONTRACT_MISMATCH` and `CONTRACT_VERSION_MISMATCH` identify mixed browser
  or Relay resource revisions.

When document-hub is the host, the browser posts the same bounded diagnostic to
`/api/v1/editor-relay/startup-failure`. document-hub logs the diagnostic ID and
safe structured fields and increments
`document_hub_ai_startup_failures_total`; it does not log document content.

## Tests

```bash
PYTHONDONTWRITEBYTECODE=1 python3 -m unittest -v \
  test_copilot_server.py test_contract_generation.py
node --test \
  test_bridge_protocol.js \
  test_slides_sheets_bridges.js \
  test_sheets_advanced_bridge.js
```

Run document-hub's Go tests and frontend build for cross-repository changes.
