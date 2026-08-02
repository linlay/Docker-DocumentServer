# ONLYOFFICE Copilot production integration

This deployment uses ONLYOFFICE DocumentServer as the editing engine and the
first-party `copilot_server.py` process as the document manager and storage
service. The bundled DocumentServer sample application is disabled with
`EXAMPLE_ENABLED=false` and is not part of the request path.

## Request flow

1. `POST /new-docx`, `/new-xlsx`, or `/new-pptx` creates an OOXML file under
   `/var/lib/onlyoffice/copilot/documents`.
2. `GET /<type>/<uuid>` returns a first-party HTML shell, a signed editor
   configuration, and the ai-bridge plugin configuration.
3. DocumentServer downloads the source through the loopback-only
   `/__document_storage/download/<type>/<uuid>` endpoint.
4. DocumentServer posts status changes to the loopback-only
   `/__document_storage/callback/<type>/<uuid>` endpoint.
5. Status `2` and `6` callbacks download the authenticated DocumentServer
   output into a same-directory temporary file, validate the OOXML signature,
   fsync it, and atomically replace the canonical file.

The public reverse proxy never needs to allow the storage endpoints. Container
Nginx accepts those endpoints only from `127.0.0.1` or `::1`, and the Python
service additionally verifies the shared HS256 JWT.

## Persistence

Compose bind-mounts:

```text
./data/documents -> /var/lib/onlyoffice/copilot/documents
```

On `singapore02`, the host path is:

```text
/docker/Docker-DocumentServer/data/documents
```

The mounted directory must be writable by container user `ds` (`101:102`).
Files under `data/` are runtime data and are excluded from Git.

## Required configuration

Copy `.env.copilot.example` to `.env` and set:

- `ONLYOFFICE_JWT_SECRET` to a long, stable secret.
- `DOCUMENT_PUBLIC_ORIGIN` to the browser-facing origin.
- `DOCUMENT_FRAME_ANCESTORS` to `'self'` plus the exact parent origins allowed
  to embed document pages. `http://localhost:*` and
  `http://127.0.0.1:*` are the only supported wildcard-port forms; wildcard
  domains are rejected. When a cross-origin parent is configured, CSP replaces
  the conflicting `X-Frame-Options: SAMEORIGIN` restriction.
- `DOCUMENT_PPTX_TEMPLATE_PATH` optionally selects a different valid PPTX
  template. The Compose default is the generic blank template mounted with
  ai-bridge; this setting does not affect DOCX or XLSX templates.
- `DOCUMENT_ADMIN_USERNAME` and `DOCUMENT_ADMIN_PASSWORD`.
- `DOCUMENT_MAX_SAVE_BYTES` to the maximum accepted callback output size.

The same JWT secret protects browser editor configuration, DocumentServer
commands, source downloads, callback requests, and ai-bridge credentials.
Never commit `.env`.

## Operational verification

After deployment:

```bash
docker compose -f docker-compose.copilot.yml config --quiet
docker compose -f docker-compose.copilot.yml up -d
docker inspect onlyoffice-documentserver --format '{{.State.Health.Status}}'
curl -fsS http://127.0.0.1:11981/copilot-api/health
```

Create a document, open the returned URL, change it, trigger Save, close the
editor, and confirm the host file modification time and size change. Restart
the container and reopen the same UUID URL to verify persistence.

The public `/__document_storage/` path must remain inaccessible. A direct
public request should return `403`, while authenticated DocumentServer
loopback requests should return `200`.
