# ONLYOFFICE ai-bridge production integration

The production boundary is document-hub v1. ONLYOFFICE DocumentServer hosts the
headless ai-bridge browser plugin and the private `copilot_server.py` Relay;
document-hub owns document creation, storage, callbacks, versions, permissions,
anonymous links, and all public HTTP APIs.

## Request flow

1. document-hub creates or opens the document and returns the signed ONLYOFFICE
   editor configuration.
2. The host page explicitly configures `window.aiBridgeOptions` with the
   document-hub Relay, image, persistence, document, and editor-session paths.
3. `host-bridge.js` registers the open editor with the private Relay on port
   3001. Port 3001 is available only on the Compose private network.
4. Authenticated document-hub APIs validate and forward tool calls to
   `/bridge/internal/*`; the browser polls, executes the command, and returns
   the result.
5. Saves, history, undo, redo, image imports, and document lifecycle operations
   remain document-hub responsibilities.

The old UUID document application, public attach/binding Relay, `/chat`, and
Python persistence endpoints are intentionally disabled and return
`LEGACY_BUSINESS_DISABLED`.

## Deployment

Use `document-hub/deploy/compose.yml` and follow
`document-hub/README.md`. Required deployment inputs include the pinned
DocumentServer image, the sibling Docker-DocumentServer build context, public
origins, authentication configuration, and the four file-based secrets
documented there.

The source-ui image bakes the browser plugin and Python Relay into one
immutable artifact. Do not bind-mount `plugins/ai-bridge` over the running
container: static assets would change immediately while the Relay keeps the
contract loaded at process startup, allowing one container to expose two
contract versions.

The same `AI_RELAY_INTERNAL_SECRET_FILE` must be mounted into document-hub and
DocumentServer. Keep DocumentServer port 3001 unpublished. The empty
`nginx-document-app.conf` include is intentional: it prevents legacy public
DocumentServer routes from being restored.

## Verification

```bash
docker compose --env-file .env -f document-hub/deploy/compose.yml config
docker compose --env-file .env -f document-hub/deploy/compose.yml up -d --build
curl -fsS http://127.0.0.1:8090/health
```

Open an editable document through document-hub and verify Relay readiness,
tool execution, image import, save, history, undo, and redo. Direct requests to
legacy `/copilot-api/bridge/attach`, `/copilot-api/chat`, UUID document, and
old persistence paths must remain unavailable.
