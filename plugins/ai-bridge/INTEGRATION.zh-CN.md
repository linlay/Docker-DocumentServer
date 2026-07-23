# ai-bridge 0.2.0 外部工程接入说明

`ai-bridge` 是无界面 ONLYOFFICE 插件。外部 Copilot 负责理解自然语言并生成受控 JSON 工具调用；插件只负责在当前文档中执行白名单操作、保存和版本回退，不执行模型生成的 JavaScript。

## 先选择接入方式

### 方式 A：Copilot 工程就是 ONLYOFFICE 编辑器宿主页

这是最简单的生产形态。你的页面创建 `DocsAPI.DocEditor`，同时加载 `host-bridge.js`，然后直接调用 `window.aiBridge`。

```html
<div id="editor"></div>
<script src="https://docs.example.com/web-apps/apps/api/documents/api.js"></script>
<script>
  const editorConfig = {
    documentType: "word",
    document: {
      key: "tenant-7:document-42:version-18",
      title: "合同.docx",
      fileType: "docx",
      url: "https://app.example.com/files/42/download",
    },
    editorConfig: {
      callbackUrl: "https://app.example.com/onlyoffice/callback",
      user: { id: "user-9", name: "张三" },
      plugins: {
        pluginsData: [
          "https://docs.example.com/sdkjs-plugins/{A17E5F31-64AA-4E37-9A42-8D430814C2F6}/config.json?v=0.2.0-rev19"
        ],
        autostart: ["asc.{A17E5F31-64AA-4E37-9A42-8D430814C2F6}"],
      },
    },
  };

  window.aiBridgeOptions = {
    getEditorConfig: () => editorConfig,
  };
  new DocsAPI.DocEditor("editor", editorConfig);
</script>
<script src="https://docs.example.com/sdkjs-plugins/{A17E5F31-64AA-4E37-9A42-8D430814C2F6}/host-bridge.js?v=0.2.0-rev19"></script>
```

页面加载后：

```js
await window.aiBridge.ready({ timeoutMs: 30_000 });

const state = window.aiBridge.getState();
console.log(state.context.documentKey); // 当前文档唯一键
console.log(state.context.fileName);    // 当前文件名
console.log(state.editorType);          // word / slide / cell
console.log(state.capabilities.tools); // 当前编辑器允许的工具

await window.aiBridge.word.replaceText(
  { search: "甲方旧名称", replace: "甲方新名称" },
  { requestId: "copilot-turn-839-tool-1", timeoutMs: 90_000 },
);
```

### 方式 B：Copilot 是跨域 iframe 或 popup

编辑器宿主页必须明确允许 Copilot 的源。没有默认白名单，也不能使用 `*`：

```html
<script>
  window.aiBridgeOptions = {
    getEditorConfig: () => editorConfig,
    clientOrigins: ["https://copilot.example.com"],
  };
</script>
<script src="https://docs.example.com/sdkjs-plugins/{A17E5F31-64AA-4E37-9A42-8D430814C2F6}/host-bridge.js?v=0.2.0-rev19"></script>
```

Copilot iframe 页面加载 SDK，并把父窗口和父窗口的准确源交给客户端：

```html
<script src="https://docs.example.com/sdkjs-plugins/{A17E5F31-64AA-4E37-9A42-8D430814C2F6}/client-sdk.js?v=0.2.0-rev19"></script>
<script>
  const office = new AiBridgeClient({
    targetWindow: window.parent,
    targetOrigin: "https://editor.example.com",
    timeoutMs: 90_000,
  });

  const state = await office.connect();
  console.log(state.context.documentKey, state.context.fileName);

  await office.word.appendParagraph({
    text: "本段由外部 Copilot 生成。",
    fontSize: 12,
  });
</script>
```

popup 把 `targetWindow` 改成 `window.opener`。Copilot 与编辑器如果是两个互不引用的独立标签页、两台设备或两个后端进程，浏览器 `postMessage` 无法直接连接；此时由你的业务后端通过已认证 HTTP/WebSocket 把命令投递到仍打开编辑器的宿主页，再调用 `window.aiBridge`。本插件有意不使用 `BroadcastChannel`，以免同源打开多个文档时把一条编辑命令广播到错误文档。

### 本机示例：HTTPX 调用当前编辑器

仓库自带的 `localhost` demo 实现了上述 HTTP 后端投递通道，供 Agent
Platform builtin HTTPX 联调。它不是面向公网的生产鉴权方案：

1. 浏览器中的 `host-bridge.js` 使用当前 ONLYOFFICE editor JWT 注册长轮询会话。
2. HTTPX 先获取用户明确给出的 `/example/editor?...` 页面，把其中短期 editor
   JWT 保存进自身 state，不能输出或记录该 JWT。
3. 首次 `GET /copilot-api/bridge/sessions` 用短期 editor JWT 换取 12 小时
   `ai-bridge-binding` token；HTTPX 立即用新 token 覆盖临时凭证。
4. Bridge token 绑定准确的文件名、文件类型、编辑器类型与用户，但不绑定保存后会
   变化的 `document.key`。服务端只把命令投递给该稳定身份最后注册的权威页面。
5. 新页面注册后会接管旧 Relay；旧页面停止 Relay，但不循环刷新，也不影响手工编辑。
6. 浏览器页面通过 `/bridge/poll` 取得命令，调用 `window.aiBridge`，再通过
   `/bridge/result` 返回真实执行结果。

`bridge/execute` 支持 `executeTool`、`executeBatch`、`save`、`history`、
`undo`、`redo`、`getState`。写操作仍应先 inspect，并使用稳定 `requestId`；
超时重试必须复用同一 ID。compose 将 `8088/8443` 绑定到 `127.0.0.1`，浏览器
页面必须保持打开。生产系统应换成自身的用户鉴权、权限校验、审计与 WebSocket/HTTP
投递服务。

为避免回环端口绑定改变 DocumentServer example 按访问地址划分的本地存储目录，
示例 Nginx 配置固定沿用既有 demo 存储身份 `185.199.108.133`。这只属于本机示例
数据兼容逻辑，生产系统不得用来源地址代替真实用户或租户身份。

## ai-bridge 怎样知道操作哪个文档

文档身份来自创建 ONLYOFFICE 编辑器时的 `editorConfig.document.key`。宿主页每次发命令时自动附带当前 `documentKey` 和 `editorType`，插件会与它初始化时绑定的文档再次比对。外部 Client SDK 不允许自行指定另一个文档键。

建议业务系统生成稳定且不可猜错的键，例如：

```text
tenantId:documentId:storageVersion
```

注意：`document.key` 既是路由身份，也是 ONLYOFFICE 协作会话/缓存身份。文件内容产生新存储版本并重新打开时，应按你的集成规则生成新 key；不要用文件名代替唯一键。

本机 HTTPX demo 另外维护 `(fileName, fileType, editorType, userId)` 稳定绑定，
只用于把版本轮换前后的浏览器会话串起来。页面注册时仍必须用 editor JWT 校验准确
`document.key`；稳定绑定不能绕过页面级文档校验。生产系统应把这里的文件名替换为
真实租户 ID 与文档 ID。

`userId` 必须来自当前编辑器签名配置中的 `editorConfig.user.id`。本机示例应用的
`userid` URL 参数或页面选择是其来源；Nginx 注入和 ai-bridge 均不得固定、覆盖或
猜测具体用户 ID。

## Copilot 怎样发命令

推荐流程：

1. `ready()` / `connect()`，读取 `context` 和 `capabilities`。
2. 先调用对应编辑器的 `inspect()` 获取当前内容、选区、页或工作表信息。
3. 把自然语言、inspect 结果和 `public-api.json` 中当前编辑器的工具 schema 交给模型。
4. 只接收模型返回的工具名和 JSON 参数，调用 `executeBatch()`。
5. 使用本次 Copilot turn 和工具序号生成稳定 `requestId`，网络重试时复用同一个值。
6. 展示结果；失败时读取 `error.code` 决定重试、提示用户或重新 inspect。

```js
const context = await window.aiBridge.word.inspect({ maxChars: 20_000 });

// toolCalls 是模型结构化输出经过业务层校验后的结果。
const toolCalls = [
  {
    name: "word_replace_text",
    arguments: { search: "草稿", replace: "定稿", matchCase: false },
  },
  {
    name: "word_append_paragraph",
    arguments: { text: "审批结论：同意。", bold: true },
  },
];

const result = await window.aiBridge.executeBatch(toolCalls, {
  requestId: "turn-20260723-001",
});

console.log(result.changed, result.persisted, result.results);
```

一批最多 20 个工具，参数 JSON 最大 250000 字符，超时上限 300000 ms。修改类调用会串行执行：创建 checkpoint、调用 Office API、`Api.Save()`、force-save。`inspect` 只读调用不会写存储。

## 完整 JavaScript API

宿主页使用 `window.aiBridge`，Client SDK 实例使用相同的执行和快捷方法：

```ts
ready(options?)
getState()                  // 宿主页同步读取
refreshState(options?)      // Client SDK 远程刷新
execute(command, options?)
executeTool(name, args, options?)
executeBatch(toolCalls, options?)
save(options?)
history(options?)
undo(options?)
redo(options?)
on("ready" | "reload" | "error", listener)
off(eventName, listener)
```

所有 56 个快捷方法：

```text
word.inspect                 word.replaceText          word.appendParagraph
word.insertParagraph         word.formatDocument       word.formatSelection
word.formatMatches           word.deleteMatches        word.addHyperlink
word.addComment              word.addBookmark          word.formatParagraphs
word.setParagraphText        word.deleteParagraphs     word.setList
word.insertPageBreak         word.navigate             word.scroll
word.scaleFont               word.addTable             word.setTableCell
word.formatTable             word.editTable            word.setPageLayout
word.setHeaderFooter         word.setDocumentText

slides.inspect               slides.inspectObjects     slides.replaceText
slides.scaleFont             slides.formatText         slides.formatSelection
slides.addSlide              slides.duplicateSlide     slides.deleteSlide
slides.addTextBox            slides.setBackground      slides.addShape
slides.updateShape           slides.deleteObject       slides.inspectCharts
slides.addChart              slides.updateChart        slides.deleteChart

sheets.inspect               sheets.setValues          sheets.setFormula
sheets.replaceText           sheets.formatRange        sheets.addSheet
sheets.renameSheet           sheets.deleteSheet        sheets.addChart
sheets.inspectCharts         sheets.updateChart        sheets.deleteChart
```

完整参数类型在 `public-api.d.ts`；机器可读 schema 在 `public-api.json`。另一个 TypeScript 工程可以复制这两个文件，或从 DocumentServer 静态地址下载并固定到版本 `0.2.0`。

### Word 完整操作边界

- `word_inspect` 返回正文、选区、当前页/可见页、带一基序号的段落结构、表格尺寸，并可按需返回批注。
- 精确文本操作包括替换、删除、字符样式、超链接、批注和书签；`occurrence` 用于只处理第 N 次命中。
- 段落操作支持标题/命名样式、字体、对齐、段前段后、行距、缩进、同页/孤行控制、分页符和项目符号/编号。
- 表格操作支持创建填充、样式、单元格格式、增删行列、合并/拆分、清空和删除。
- 页面操作支持纸张大小、横竖向、页边距、页眉页脚及动态页码。
- `word_scroll` 是稳定的按页上滚/下滚；`word_navigate` 还支持首页、末页、指定页、上一页、下一页、相对页和搜索定位。ONLYOFFICE Office API 没有跨版本稳定的像素级鼠标滚轮契约，因此 Bridge 不注入浏览器鼠标事件，也不会把按页导航误报成像素滚动。

### PPTX 图形、渐变和图表完整边界

- `slides_inspect_objects` 读取图形、图表、图片、表格、OLE、组合等绘图对象，返回零基 `objectIndex`、内部 ID、名称、类型、位置、尺寸、旋转、翻转；图形还返回几何类型、文本、填充和线条。
- `slides_add_shape` / `slides_update_shape` 支持任意 ONLYOFFICE 预设 `shapeType`，以及 `text`、`xMm`、`yMm`、`widthMm`、`heightMm`、`rotationDeg`、`flipH`、`flipV`、`name`、字体、对齐、`paddingMm`、`fill` 和 `line`。
- `slides_delete_object` 可用 `objectId`、零基 `objectIndex` 或 `name` 删除任意绘图对象。
- `slides_set_background` 支持 `custom`、`clear`、`layout`、`master`；`custom` 使用同一套 `fill` 模型。
- `slides_inspect_charts` / `slides_add_chart` / `slides_update_chart` / `slides_delete_chart` 覆盖图表类型、二维数值系列、系列名、分类、数值格式、标题、样式、位置、尺寸、旋转、图表区/绘图区/标题填充与线条、图例、横纵轴、标签、网格线、系列和数据点格式、系列增删。
- `seriesUpdates[].type` 可修改组合图中的单个系列类型；数据点还支持
  `markerFill`、`markerLine` 和 `allMarkers`。`line`、`lineMarker`、`column`
  等友好名称会转换为 ONLYOFFICE 的 `lineNormal`、`lineNormalMarker`、`bar`。

填充模型：

```ts
type Fill =
  | { type: "none" }
  | { type: "solid"; color: "#RRGGBB" }
  | {
      type: "linearGradient";
      angleDeg?: number;
      stops: Array<{ position: number /* 0..100 */; color: "#RRGGBB" }>;
    }
  | {
      type: "radialGradient";
      stops: Array<{ position: number /* 0..100 */; color: "#RRGGBB" }>;
    }
  | {
      type: "pattern";
      pattern: string;
      backgroundColor: "#RRGGBB";
      foregroundColor: "#RRGGBB";
    }
  | { type?: "raw"; raw: unknown };
```

`includeRaw: true` 会返回 ONLYOFFICE `ToJSON()` 的无损对象；把
`inspectObjects` 得到的 `fill.raw` 或 `line.raw` 原样传回
`fill: { raw }` / `line: { raw }`，即可保留 Bridge 尚未单独建模的高级
Office 属性。

创建带线性渐变的图形：

```js
await window.aiBridge.slides.addShape({
  slide: 1,
  shapeType: "roundRect",
  name: "KPI",
  text: "42%",
  xMm: 20,
  yMm: 25,
  widthMm: 70,
  heightMm: 32,
  fill: {
    type: "linearGradient",
    angleDeg: 45,
    stops: [
      { position: 0, color: "#2F80ED" },
      { position: 100, color: "#56CCF2" },
    ],
  },
  line: { widthPt: 1.5, color: "#1B4F9C" },
});
```

创建和更新 PPT 图表：

```js
await window.aiBridge.slides.addChart({
  slide: 1,
  type: "bar",
  series: [[120, 180, 240], [90, 150, 210]],
  seriesNames: ["收入", "成本"],
  categories: ["Q1", "Q2", "Q3"],
  title: "季度趋势",
  legend: { position: "bottom" },
  dataLabels: { showValue: true },
  xMm: 20,
  yMm: 55,
  widthMm: 190,
  heightMm: 100,
});

await window.aiBridge.slides.updateChart({
  slide: 1,
  chartIndex: 0,
  horizontalAxis: { title: "季度" },
  verticalAxis: { title: "金额", numberFormat: "#,##0" },
  seriesUpdates: [{
    index: 0,
    name: "净收入",
    values: [125, 190, 260],
    fill: { type: "solid", color: "#2F80ED" },
  }],
});
```

### XLSX 图表完整边界

`sheets_add_chart`、`sheets_inspect_charts`、`sheets_update_chart` 和
`sheets_delete_chart` 使用零基 `chartIndex` 或名称定位。除 PPT 图表共有的
标题、样式、填充、线条、图例、坐标轴、标签、网格线、系列和数据点格式外，
还支持：

- `range`：新图表的 A1 数据源，可使用 `selection`。
- `categoryRange`：更新分类来源区域。
- `addSeries`：用 `name`、`valuesRange`、可选 `xValuesRange` 新增系列。
- `seriesUpdates[].valuesRange` / `xValuesRange`：重定向已有系列。
- `fromColumn` / `fromRow`、`columnOffsetMm` / `rowOffsetMm`、`widthMm` /
  `heightMm`：设置单元格锚点和尺寸。

所有尺寸参数均为毫米，线宽为磅；渐变节点位置为 `0..100`，内部会转换为
ONLYOFFICE 的 `0..100000` 单位，角度会转换为 `1/60000` 度单位。

`sheets_delete_chart` 调用官方 `ApiDrawing.Delete()`。该方法在部分
ONLYOFFICE Docs 版本/许可中属于付费能力；方法不可用时 Bridge 会明确返回
“当前 ONLYOFFICE 版本不支持删除图表”，不会伪装删除成功。本仓库默认 Community
DocumentServer 9.4 已实测支持图表创建、读取和更新，但不暴露该删除方法。

## 不使用 Client SDK 时的底层 postMessage 协议

建议优先使用 `client-sdk.js`。如果另一个工程要自己实现客户端，Relay 的公开消息包如下。所有 `postMessage` 都必须使用准确的 `targetOrigin`。

连接：

```js
parent.postMessage({
  source: "ai-bridge-client",
  protocolVersion: 1,
  clientId: "client:550e8400-e29b-41d4-a716-446655440000",
  type: "connect",
  timeoutMs: 30000,
}, "https://editor.example.com");
```

连接成功响应：

```json
{
  "source": "ai-bridge-relay",
  "protocolVersion": 1,
  "clientId": "client:550e8400-e29b-41d4-a716-446655440000",
  "type": "connected",
  "state": {
    "version": "0.2.0",
    "ready": true,
    "editorType": "word",
    "context": { "documentKey": "document-42:v18", "fileName": "合同.docx" },
    "capabilities": { "tools": ["word_inspect", "word_replace_text"], "controls": ["save", "history", "undo", "redo"] }
  }
}
```

调用：

```js
parent.postMessage({
  source: "ai-bridge-client",
  protocolVersion: 1,
  clientId: "client:550e8400-e29b-41d4-a716-446655440000",
  type: "request",
  requestId: "turn-55-tool-1",
  method: "executeTool",
  params: {
    name: "word_replace_text",
    arguments: { search: "草稿", replace: "定稿" },
    options: { timeoutMs: 90000 }
  }
}, "https://editor.example.com");
```

`method` 只允许 `execute`、`executeTool`、`executeBatch`、`save`、`history`、`undo`、`redo`、`getState`。响应为同一 `requestId` 的 `type: "result"`，成功时包含 `result`，失败时包含 `{ error: { code, message, requestId, details } }`。Relay 还可能发送 `type: "event"`，事件名为 `reload`。收到消息时客户端必须同时验证 `event.source === parent/opener`、`event.origin`、`source`、`protocolVersion` 和 `clientId`。

这套 Relay 协议是外部工程的公共协议。`host-bridge.js` 与隐藏插件 iframe 之间还有一层实例绑定消息，但它属于内部实现，不应由外部工程直接伪造或依赖。

## 错误处理与幂等重试

```js
try {
  await office.sheets.setValues(
    { sheet: "预算", range: "B2:B4", values: [[10], [20], [30]] },
    { requestId: "turn-55-write-budget" },
  );
} catch (error) {
  if (error instanceof AiBridgeError) {
    console.error(error.code, error.message, error.requestId, error.details);
  }
}
```

常见错误：

- `DOCUMENT_MISMATCH`：宿主页的当前 key 已改变，命令被拒绝。
- `EDITOR_MISMATCH`：把 Word 命令发给了 PPT/表格实例。
- `TOOL_NOT_ALLOWED`：当前编辑器不允许该工具。
- `INVALID_ARGUMENTS` / `INVALID_COMMAND`：模型结构化输出不符合契约。
- `PERSISTENCE_FAILED`：Office API 已执行，但示例持久化服务保存失败，应提示用户并核对 callback/force-save。
- `TIMEOUT` / `CONNECTION_TIMEOUT`：调用或 Relay 握手超时。

插件按 `requestId` 缓存最近 100 个响应。对可能修改文档的超时请求，不要换一个新 ID 盲目重试；应复用原 ID，使插件返回已缓存结果而不是重复编辑。

## 保存、撤销和生产存储

```js
await office.save();
const history = await office.history();
await office.undo(); // 回到上一 checkpoint，页面会 reload
await office.redo(); // 回到下一 checkpoint，页面会 reload
```

仓库内 `/copilot-api/*` 只实现 DocumentServer 示例应用的保存/版本服务。生产环境仍应由你的文档存储服务正确处理 ONLYOFFICE callback，尤其是 force-save 状态 6 和最终保存状态 2，并做用户鉴权、文档权限校验和审计。

## 安全约束

- Relay 只接受 `clientOrigins` 中的精确源；不要使用通配符。
- Copilot iframe 建议同时设置 CSP `frame-ancestors`，宿主页设置合适的 `frame-src`。
- 模型只能输出 `public-api.json` 中的工具调用，不能把任意 JavaScript 交给插件执行。
- 在业务后端验证当前用户对 `documentKey` 的读写权限；浏览器源校验不能代替用户授权。
- 写操作建议记录 `requestId`、用户、文档键、工具名、参数摘要和结果，敏感正文不要直接写入日志。

## 版本兼容

当前插件版本为 `0.2.0`，消息协议版本为 `1`，本次构建缓存键为
`?v=0.2.0-rev19`。生产页面应固定到实际发布的不可变缓存键，升级前先比较
`public-api.json`。协议版本不一致时 Client 和 Relay 不建立连接。
