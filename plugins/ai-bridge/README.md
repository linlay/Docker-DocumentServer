# ai-bridge

`ai-bridge` is a headless ONLYOFFICE plugin. It has no panel, toolbar, chat UI,
or model client. An external Copilot page sends allow-listed JSON tool calls to
the plugin, and the plugin executes them in the current document with the
ONLYOFFICE Office JavaScript API.

Plugin identity:

- Name: `ai-bridge`
- Version: `0.2.0`
- GUID: `asc.{A17E5F31-64AA-4E37-9A42-8D430814C2F6}`
- Editors: Word, Presentation, Spreadsheet

## Architecture

```text
External Copilot UI
  -> window.aiBridge (same-page API)
     OR client-sdk.js -> allow-listed postMessage relay
     OR local HTTPX -> editor-JWT-bound /copilot-api/bridge relay
  -> host-bridge.js
  -> instance-bound postMessage protocol
  -> headless ai-bridge plugin (plugin.js)
  -> allow-listed Word / Slides / Sheets bridge
  -> Asc.plugin.callCommand()
  -> ONLYOFFICE Office JavaScript API
  -> Api.Save()
  -> POST /copilot-api/forcesave
  -> callback persistence
```

The host and plugin perform a per-editor handshake. Commands contain the
current `document.key` and editor type, and the plugin rejects a command if its
target does not match the open document. The protocol does not use
`BroadcastChannel`, so opening multiple documents in same-origin browser tabs
does not broadcast a command to every plugin instance.

## Add the plugin to an editor

Add the plugin config to the editor configuration and autostart it:

```js
const editorConfig = {
  documentType: "word",
  document: {
    key: "document-123-version-7",
    fileType: "docx",
    title: "report.docx",
    url: "https://storage.example.com/report.docx",
  },
  editorConfig: {
    callbackUrl: "https://app.example.com/onlyoffice/callback",
    user: { id: "user-42", name: "Demo User" },
    customization: {
      compactToolbar: true,
    },
    plugins: {
      pluginsData: [
        "https://docs.example.com/sdkjs-plugins/{A17E5F31-64AA-4E37-9A42-8D430814C2F6}/config.json?v=0.2.0-rev19",
      ],
      autostart: ["asc.{A17E5F31-64AA-4E37-9A42-8D430814C2F6}"],
    },
  },
};

const docEditor = new DocsAPI.DocEditor("editor", editorConfig);
```

The bundled example injects this configuration through
`nginx-ds-example.conf`.

The bundled Nginx injection does not select or rewrite an editor user. The
example application's `userid` input remains authoritative, and ai-bridge reads
the resulting signed `editorConfig.user.id`. The injection removes the complete
left-menu layout container, including its reserved width, and hides the unused
Collaboration, Plugins, and default AI toolbar tabs in the same-origin demo
without changing document permissions. It also enables ONLYOFFICE's native
compact toolbar: the ribbon starts folded, a normal tab click expands it, and
clicking the active tab folds it again and clears the selected tab. Existing
ONLYOFFICE toolbar preferences still take precedence. External integrations
that want the same behavior must set
`editorConfig.customization.compactToolbar` to `true` before constructing
`DocsAPI.DocEditor`.

## Load the external-page API

Load `host-bridge.js` in the page that owns the ONLYOFFICE editor. If the
editor config is not available as a global `config` variable, provide it before
loading the script:

```html
<script>
  window.aiBridgeOptions = {
    getEditorConfig: () => editorConfig,
  };
</script>
<script src="https://docs.example.com/sdkjs-plugins/{A17E5F31-64AA-4E37-9A42-8D430814C2F6}/host-bridge.js?v=0.2.0-rev19"></script>
```

`host-bridge.js` must run in the page that contains the editor. A cross-origin
Copilot iframe or popup can use the allow-listed `client-sdk.js` relay described
in `INTEGRATION.zh-CN.md`. A completely unrelated tab or process with no window
reference must relay its command through the application's authenticated
WebSocket or HTTP channel. The bundled localhost demo now includes such an HTTP
Relay for HTTPX: the browser registers with its signed ONLYOFFICE editor JWT,
while HTTPX exchanges the same short-lived JWT once for a scoped 12-hour bridge
binding token. The binding follows the authoritative browser session for the
same file, editor type, and user even when a save rotates `document.key`. The
compose ports are bound to `127.0.0.1`; this demo Relay is not a production
authentication design.

The bundled Nginx example keeps the pre-existing demo storage identity
`185.199.108.133` stable after the loopback bind, so previously uploaded local
example files continue resolving to the same storage directory.

## External Copilot API

Wait for the hidden plugin to be ready:

```js
await window.aiBridge.ready();

console.log(window.aiBridge.context);
console.log(window.aiBridge.capabilities);
```

Execute one typed operation:

```js
await window.aiBridge.word.replaceText({
  search: "Old company",
  replace: "New company",
});
```

Execute an agent-produced batch:

```js
const result = await window.aiBridge.executeBatch([
  {
    name: "word_replace_text",
    arguments: { search: "Draft", replace: "Final" },
  },
  {
    name: "word_append_paragraph",
    arguments: { text: "Generated by the external Copilot." },
  },
]);
```

Common controls:

```js
await window.aiBridge.save();
await window.aiBridge.undo();
await window.aiBridge.redo();
const history = await window.aiBridge.history();
```

`window.onlyofficeAI` is kept as a compatibility alias for
`window.aiBridge`.

For a complete Chinese integration guide covering a direct host page,
cross-origin iframe/popup relay, document identity, idempotency, errors, and
production persistence, see [INTEGRATION.zh-CN.md](INTEGRATION.zh-CN.md).

The versioned public contract is available as:

- `public-api.d.ts`: TypeScript declarations for the full API and all tool arguments.
- `public-api.json`: machine-readable tool schemas, limits, errors, and transport metadata.
- `client-sdk.js`: optional cross-origin iframe/popup client for the allow-listed relay.

## Tool names

- Word: `word_inspect`, `word_replace_text`, `word_append_paragraph`,
  `word_insert_paragraph`, `word_format_document`, `word_format_selection`,
  `word_format_matches`, `word_delete_matches`, `word_add_hyperlink`,
  `word_add_comment`, `word_add_bookmark`, `word_format_paragraphs`,
  `word_set_paragraph_text`, `word_delete_paragraphs`, `word_set_list`,
  `word_insert_page_break`, `word_navigate`, `word_scroll`, `word_scale_font`,
  `word_add_table`, `word_set_table_cell`, `word_format_table`,
  `word_edit_table`, `word_set_page_layout`, `word_set_header_footer`,
  `word_set_document_text`.
- Slides: `slides_inspect`, `slides_inspect_objects`, `slides_replace_text`,
  `slides_scale_font`, `slides_format_text`, `slides_format_selection`,
  `slides_add_slide`, `slides_duplicate_slide`, `slides_delete_slide`,
  `slides_add_textbox`, `slides_set_background`, `slides_add_shape`,
  `slides_update_shape`, `slides_delete_object`, `slides_inspect_charts`,
  `slides_add_chart`, `slides_update_chart`, `slides_delete_chart`.
- Sheets: `sheets_inspect`, `sheets_set_values`, `sheets_set_formula`,
  `sheets_replace_text`, `sheets_format_range`, `sheets_add_sheet`,
  `sheets_rename_sheet`, `sheets_delete_sheet`, `sheets_add_chart`,
  `sheets_inspect_charts`, `sheets_update_chart`, `sheets_delete_chart`.

The external agent must send tool names and JSON arguments, never JavaScript
source. `plugin.js` checks the editor-specific allow list before calling a
bridge. Requests are serialized, deduplicated by request ID, and cached briefly
so a retried transport message cannot apply the same edit twice.

### Word capability coverage

| Area | Supported operations |
| --- | --- |
| Read/context | Full text, selection, pages, visible pages, paragraph indexes/styles, table dimensions, optional comments |
| Text | Replace, delete exact matches, overwrite a paragraph or the whole document |
| Character style | Font, size, bold, italic, underline, strikeout, color, highlight, caps, spacing, sub/superscript |
| Paragraph style | Named/heading styles, outline level, alignment, spacing, line spacing, indents, keep/widow/page-break rules |
| Structure | Append/insert/delete paragraphs, bullet/numbered lists, page breaks, bookmarks, links, comments |
| Tables | Create/fill, style, format cells, add/remove rows and columns, merge/split, clear/delete |
| Page layout | A4/Letter/Legal/custom size, portrait/landscape, margins, header/footer distances |
| Header/footer | Default/first/even header or footer, text, dynamic page number and page count |
| View navigation | Start/end/page, next/previous/relative page, search-to-selection, page-granular up/down scrolling |

`word_scroll` deliberately uses document pages rather than synthetic browser
mouse events. The Office API exposes stable page navigation and selection
movement, but not a cross-version pixel-wheel contract. Page-granular scrolling
therefore remains deterministic and does not create an undo entry.

### Slides and Sheets capability coverage

| Area | Supported operations |
| --- | --- |
| PPT read | Slide text plus all drawings; object ID/index/name, kind, position, size, rotation, flips, shape geometry/text/fill/line, chart summary, and optional raw Office JSON |
| PPT slides | Add, duplicate, delete, and set custom/clear/layout/master background |
| PPT shapes | Add any preset geometry; update text, geometry, name, position, size, rotation, flips, padding, text style, fill, and line; delete any drawing |
| PPT charts | Inspect, add, update, and delete; data/categories, series and points, axes, legend, labels, gridlines, number formats, fills, lines, style, position, and size |
| XLSX cells | Inspect workbook, set values/formulas, replace text, format ranges, and add/rename/delete sheets |
| XLSX charts | Inspect, add, update, and delete; source/category/series ranges, series and points, axes, legend, labels, gridlines, number formats, fills, lines, style, position, and size |
| Fill model | None, solid, linear gradient, radial gradient, pattern, and lossless `raw` Office JSON replay |

Gradient stops use `position: 0..100`; `angleDeg` is in degrees. Object and
chart indexes are zero-based. All physical dimensions use millimetres and line
widths use points. `inspectObjects` and `inspectCharts` can return the exact
`raw` JSON accepted by later `fill: { raw }` or `line: { raw }` updates.
Spreadsheet chart deletion delegates to `ApiDrawing.Delete()`, which is a paid
capability in some ONLYOFFICE Docs editions; unsupported editions return an
explicit error without reporting a false deletion.

## Persistence service

Mutating commands create a checkpoint, edit the live document, call
`Api.Save()`, and then force-save the current document back to the bundled
example storage. These same-origin endpoints are used:

- `POST /copilot-api/checkpoint`
- `POST /copilot-api/forcesave`
- `POST /copilot-api/history`
- `POST /copilot-api/undo`
- `POST /copilot-api/redo`

`copilot_server.py` implements the endpoints for the example application. In a
production integration, persist ONLYOFFICE callback statuses `6` and `2` in
your own storage service.

## Local HTTPX Relay

The localhost example exposes these additional endpoints:

- `POST /copilot-api/bridge/register`: the open editor page registers with its
  signed editor JWT and receives an opaque page relay key.
- `POST /copilot-api/bridge/poll` and `POST /copilot-api/bridge/result`: the
  registered page receives commands and returns `window.aiBridge` results.
- `GET /copilot-api/bridge/sessions`: accepts an editor JWT or bridge binding
  token, returns the single authoritative page, and renews the scoped binding
  token.
- `POST /copilot-api/bridge/execute`: accepts the binding token, validates the
  requested tool against the authoritative page capabilities, and waits for the
  live result.

The caller first fetches the exact local `/example/editor?...` page and stores
its short-lived editor JWT without printing it. The first `sessions` request
exchanges that JWT for a 12-hour `ai-bridge-binding` token scoped to the exact
file name, file type, editor type, and user. A newly registered page for that
stable identity supersedes the previous Relay session, so a storage-version
change does not require another bind. Stable `requestId` values deduplicate
retries across page handoff. The editor page must remain open.

## Local run

```bash
docker compose -f docker-compose.copilot.yml up -d
```

Open `http://localhost:8088/example/`. The plugin is loaded automatically but
does not render a user interface.

## Files

- `config.json`: headless plugin registration.
- `index.html`: script-only plugin entry point.
- `plugin.js`: handshake, validation, command queue, save controls.
- `host-bridge.js`: API exposed to the external Copilot page.
- `client-sdk.js`: optional cross-origin iframe/popup client.
- `public-api.d.ts`: TypeScript API declarations.
- `public-api.json`: machine-readable API and tool contract.
- `INTEGRATION.zh-CN.md`: complete external-project integration guide.
- `bridges/*.js`: editor-specific Office API implementations.
- `copilot_server.py`: optional planning API plus example save/version service.
- `nginx-ds-example.conf`: example plugin and host API injection.
- `../../docker-compose.copilot.yml`: local deployment.
