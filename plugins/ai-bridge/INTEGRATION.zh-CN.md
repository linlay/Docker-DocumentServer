# ai-bridge 0.6.0 外部工程接入说明

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
      customization: {
        compactToolbar: true,
      },
      plugins: {
        pluginsData: [
          "https://docs.example.com/sdkjs-plugins/{A17E5F31-64AA-4E37-9A42-8D430814C2F6}/config.json?v=0.6.0-rev4"
        ],
        autostart: ["asc.{A17E5F31-64AA-4E37-9A42-8D430814C2F6}"],
        options: {
          "asc.{A17E5F31-64AA-4E37-9A42-8D430814C2F6}": {
            hostOrigin: window.location.origin,
            channelId: crypto.randomUUID(),
          },
        },
      },
    },
  };

  window.aiBridgeOptions = {
    getEditorConfig: () => editorConfig,
  };
  new DocsAPI.DocEditor("editor", editorConfig);
</script>
<script src="https://docs.example.com/sdkjs-plugins/{A17E5F31-64AA-4E37-9A42-8D430814C2F6}/host-bridge.js?v=0.6.0-rev4"></script>
```

`plugins.options` 必须在创建 `DocsAPI.DocEditor` 前写入。同一个配置对象会把准确的
宿主页源和本次页面加载生成的随机 channel 同时交给 `host-bridge.js` 与隐藏插件。
当整个文档宿主页又被嵌入跨域外层 iframe 时，插件会逐层检查祖先窗口，只与
`hostOrigin` 匹配的宿主页握手；外层页面不需要加载任何 Bridge 脚本。sandbox
iframe 必须至少包含 `allow-scripts allow-same-origin`；来源为 `"null"` 的严格
sandbox 不支持。缺少 options 时只保留原生顶层页签的旧握手方式。

#### 固定显示“访客”，不询问昵称

本地版可以让浏览器保存一个不含真实用户信息的 UUID，并在创建编辑器之前把已有
签名配置重签成稳定访客。`local-guest.js` 默认使用
`localStorage["onlyoffice.localGuestId.v1"]`，显示名固定为 `访客`：

```html
<script src="https://docs.example.com/sdkjs-plugins/{A17E5F31-64AA-4E37-9A42-8D430814C2F6}/local-guest.js?v=0.6.0-rev4"></script>
<script>
  // editorConfig 必须已经包含业务后端签发的短期 HS256 token。
  await window.OnlyOfficeLocalGuest.prepare(editorConfig, {
    // 外部宿主页应把此接口反向代理到自己的同源路径，不开放通配 CORS。
    endpoint: "/api/onlyoffice/editor-config/anonymous",
  });

  new DocsAPI.DocEditor("editor", editorConfig);
</script>
```

接口契约为
`POST { editorToken, anonymousId } -> { ok, user, token, expiresAt }`。仓库内置
实现位于 `/copilot-api/editor-config/anonymous`：它先验证原 editor JWT，再只把
用户改成 `{ id: "local-guest:<uuid>", name: "访客" }`，设置
`customization.anonymous.request = false`，保留原文档、权限、回调地址和到期时间
后重新签名。接口不接受客户端提供的显示名、完整用户 ID、文档或权限；前端也不得
持有 `JWT_SECRET`。清除浏览器站点数据后会生成新的匿名身份。

`customization.compactToolbar: true` 启用 ONLYOFFICE 原生紧凑功能区：首次进入时
菜单默认收起且没有选中的普通页签；点击 Home、Insert、Layout 等页签会展开完整
菜单，再点当前页签会收起并取消选中。已有的 ONLYOFFICE 本地工具栏偏好优先于该
默认值。外部集成必须在创建 `DocsAPI.DocEditor` 前设置此项；无需监听 Home 点击，
也不要依赖编辑器内部 `section` 的绝对 XPath。

当 `host-bridge.js` 与编辑器 iframe 同源时，还会把原生 Save、Undo、Redo 按钮
移动到 Editing 控件左侧，并隐藏完整的 28px 文档标题栏。Logo、文档名称、标题栏
用户名、Print 快捷按钮和快速访问下拉不再占用空间；Print 仍保留在 File 菜单，
Open file location 与 Mark as favorite 两个顶栏按钮也会隐藏，主工具栏的协作者
状态不受影响。迁移直接复用 ONLYOFFICE 原生按钮节点，保留按钮状态、快捷键和
事件；如果未来版本缺少所需语义槽位，则保留原标题栏作为安全降级。跨域编辑器
iframe 无法使用这项 DOM 收缩。

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
<script src="https://docs.example.com/sdkjs-plugins/{A17E5F31-64AA-4E37-9A42-8D430814C2F6}/host-bridge.js?v=0.6.0-rev4"></script>
```

Copilot iframe 页面加载 SDK，并把父窗口和父窗口的准确源交给客户端：

```html
<script src="https://docs.example.com/sdkjs-plugins/{A17E5F31-64AA-4E37-9A42-8D430814C2F6}/client-sdk.js?v=0.6.0-rev4"></script>
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
2. 真实 `/docx|xlsx|pptx/<uuid>` 页面在浏览器中打开并保持 Relay ready 后，
   HTTPX 只把准确 UUID 文件名和配置中固定的 editor type 传给
   `POST /copilot-api/bridge/attach`。
3. attach 最多等待 10 秒，选择相同文件名、文件类型和 editor type 的最新 ready
   权威 Relay，并签发 12 小时 `ai-bridge-binding` token。HTTPX 把它保存到当前
   Chat state；Agent 不接触 editor JWT、binding token、sessionId 或 documentKey。
4. Bridge token 绑定准确的文件名、文件类型、编辑器类型与用户，但不绑定保存后会
   变化的 `document.key`。服务端只把命令投递给该稳定身份最后注册的权威页面。
5. 新页面注册后会接管旧 Relay；旧页面停止 Relay，但不循环刷新，也不影响手工编辑。
6. 浏览器页面通过 `/bridge/poll` 取得命令，调用 `window.aiBridge`，再通过
   `/bridge/result` 返回真实执行结果。

同一稳定文档身份始终只有最后注册的一个活跃 Relay。HTTPX 请求携带的
`sessionId` 只是兼容性提示，服务端以 binding token 为边界，把命令路由到当前
权威 session；旧 ID 不会导致 `SESSION_NOT_AUTHORITATIVE`。最新页面关闭后不会
自动恢复旧页面，AI 返回 `NO_ACTIVE_EDITOR`；刷新旧页面会生成新的 session 并
重新接管。

旧的 `/documents/editor` 与 `/bridge/sessions` 继续兼容已有客户端，但不再暴露
在 Agent HTTPX 配置和技能中。刷新页面或 force-save 后继续复用现有 binding
token；只有 token 明确过期、无效或目标文档改变时才重新 attach。

`bridge/execute` 支持 `executeTool`、`executeBatch`、`save`、`history`、
`undo`、`redo`、`getState`。`inspect` 只检查调用包装；写操作应先调用真实
`POST /copilot-api/bridge/validate` 预检，并使用稳定 `requestId`；
超时重试必须复用同一 ID。compose 将 `8088/8443` 绑定到 `127.0.0.1`，浏览器
页面必须保持打开。生产系统应换成自身的用户鉴权、权限校验、审计与 WebSocket/HTTP
投递服务。

Word、Slides 和 Sheets 批次可先调用只读 `POST /copilot-api/bridge/validate`。
它与执行共用参数归一化、schema 和语义校验，但不向编辑器投递命令，也不创建
history point 或保存文件。
机器可读策略位于 `public-api.json` 的 `inputNormalization.word`：标准字段优先于
deprecated 别名，已知枚举拼写、大小写与分隔符差异会转换成标准值；返回值可带仅含
路径和转换类型的 `argumentNormalizations`，不会回显参数值。未知字段、字符串到
数字/布尔值的隐式转换、疑似把 twips 当 pt，以及缺少目标的分页或破坏性操作仍会
拒绝。

公开网关通过 `POST /new-docx`、`POST /new-xlsx` 和 `POST /new-pptx`
原子复制官方空白模板，返回 UUID v4 能力链接。只有知道完整 URL 的调用者才能打开
对应文档；普通用户没有列举或找回接口。`GET /admin/` 使用 Basic Auth 展示 UUID
文件，管理员账号来自 `DOCUMENT_ADMIN_USERNAME` 和 `DOCUMENT_ADMIN_PASSWORD`。
原 `/example*` 路径对外返回 404。

为避免回环端口绑定改变 DocumentServer example 按访问地址划分的内部存储目录，
Nginx 在内部仍固定沿用 demo 存储身份 `185.199.108.133`，下载和 callback 仅允许
容器回环访问。生产系统不得用来源地址代替真实用户或租户身份。

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

`userId` 必须来自当前编辑器签名配置中的 `editorConfig.user.id`。本机 UUID
能力链接不接收 `userid`，而是通过受限重签接口使用浏览器稳定的
`local-guest:<uuid>`。外部业务系统应使用自己的认证用户，或采用上面的稳定访客
流程，不能在 JWT 签发之后仅修改未签名的 `editorConfig.user`。

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

一批最多 20 个工具，普通参数 JSON 最大 250000 字符，超时上限 300000 ms。图片导入和 HTTP Relay execute 路径单独允许 12 MiB 请求体；其他聊天、保存和控制接口仍使用原有限制。修改类调用会串行执行：创建 checkpoint、调用 Office API、`Api.Save()`、force-save。`inspect` 只读调用不会写存储。

### Word 大文档的定点表格、inline 内容控件与替换

`word_add_table` 默认仍以 `insertAt: "end"` 追加到文末。`current` 在当前光标处插入；
`before` / `after` 必须且只能使用 `paragraphIndex`、`tableIndex` 或 `search` 中的一个
作为锚点。锚点只支持 Word 正文的顶层段落或顶层表格；表格单元格、嵌套表格和内容
控件内部的段落会被拒绝。搜索锚点可附带 `matchCase`、`matchMode` 和从 1 开始的
`occurrence`。`pageBreakBefore: true` 会在表格紧前方插入一个分页段落。

```js
await window.aiBridge.executeBatch([
  {
    name: "word_add_table",
    arguments: {
      rows: 2,
      cols: 2,
      data: [["项目", "结论"], ["安全复核", "通过"]],
      insertAt: "before",
      search: "附件",
      matchMode: "exact",
      occurrence: 1,
      pageBreakBefore: true,
    },
  },
  {
    name: "word_manage_content_control",
    arguments: {
      action: "add",
      kind: "inline",
      tableIndex: 1,
      row: 2,
      column: 2,
      contentMode: "replace",
      text: "通过",
      tag: "review-result",
      title: "复核结论",
    },
  },
]);
```

inline 内容控件的 `contentMode` 默认为 `append`，与旧调用一致。显式使用 `replace`
时必须指定 `paragraphIndex`、`current: true` 或完整的 `tableIndex + row + column`
目标；目标原内容会被清空，再直接写入携带最终 `text` 和属性的 inline SDT。它不承诺
保留被替换区域的富文本格式，也不提供任意文本范围的原样包裹。

### Word / PPT 表格 `data` 单元格契约

`word_add_table`、`word_add_nested_table` 和 `slides_add_table` 的 `data` 都接受二维
数组。每个单元格只能是标量 `string | number | boolean | null`，或一个必须包含
字符串 `text` 的格式对象。标量会统一转成文本，`null` 和行尾缺失的单元格写为空白。

Word 普通表格与嵌套表格使用同一规则。顶层 `fontFamily`、`fontSize`、`bold`、
`italic`、`color` 是整表文字默认值，单元格对象中的同名字段优先；对象还可使用
`word_set_table_cell` 的文字、背景、水平/垂直对齐、段落格式和 `widthPercent`
字段。例如：

```js
await window.aiBridge.word.addTable({
  rows: 3,
  cols: 2,
  fontFamily: "宋体",
  fontSize: 10.5,
  data: [
    ["项目", "数值"],
    [
      {
        text: "营业收入",
        fontFamily: "微软雅黑",
        bold: true,
        color: "#FFFFFF",
        backgroundColor: "#2E74B5",
        align: "center",
        verticalAlign: "center",
      },
      128.6,
    ],
    [true, null],
  ],
});
```

PPT 单元格格式对象复用 `SlidesTableCellFormat`。创建时先应用各单元格对象，再用
`header` 覆盖首行格式；只有 `header.text` 明确存在时才会覆盖首行单元格文字：

```js
await window.aiBridge.slides.addTable({
  slide: 1,
  rows: 2,
  columns: 2,
  data: [
    [{ text: "指标", fontFamily: "微软雅黑", color: "#112233" }, "数值"],
    [{ text: "毛利率", backgroundColor: "#E8EEF5", align: "right" }, 0.376],
  ],
  header: { bold: true, backgroundColor: "#DDEEFF", align: "center" },
});
```

文字颜色字段只允许 `color`，不支持 `textColor`。格式对象缺少 `text`、出现未知
字段、颜色/枚举/嵌套 `fill` 或 `border` 不符合 schema 时，整批会在任何文档变更前
返回 `INVALID_TOOL_ARGUMENTS`；错误路径精确到对应单元格或字段，例如
`arguments.data[1][0].textColor`。对象不会再被隐式写成 `[object Object]`。已经写入
该字符串的旧文档不会自动迁移，需要重写相关单元格或重新生成文档。

`word_replace_text` 的参数与结果没有变化。桥接首次替换时会探测运行时是否支持
文档级 `SearchAndReplace()`：支持时，多个替换和同批表格、内容控件会按调用顺序留在
同一个 mutation `callCommand` 与 history point 中；旧运行时自动使用兼容插件方法。
原生方法已经开始执行后若返回失败，整批直接报错，不会再次走兼容路径。

force-save 成功后不会刷新编辑器页面。当前 WebSocket 会话已经包含刚保存的版本，
继续使用现有会话可以完整保留当前页、页内滚动位置和选区。只有 `undo` / `redo`
这类用历史文件替换当前存储版本的操作才会 reload。

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

快捷方法以 `public-api.d.ts` 和 `public-api.json` 为准，以下列出 0.6.0 的主要方法：

```text
word.inspect                 word.replaceText          word.appendParagraph
word.insertParagraph         word.formatDocument       word.formatSelection
word.formatMatches           word.deleteMatches        word.addHyperlink
word.addComment              word.addBookmark          word.addImage
word.inspectAdvanced         word.setDocumentProperties
word.manageSection           word.manageStyle          word.setTabs
word.setNumbering            word.formatTableAdvanced  word.addNestedTable
word.manageDrawing           word.addShape             word.addChart
word.addMath                 word.addOleObject          word.manageFields
word.manageLongDocument      word.manageComments       word.manageRevisions
word.setProtection           word.manageContentControl word.manageCustomXml
word.inspectMacros           word.setMacros             word.setWatermark
word.formatParagraphs
word.setParagraphText        word.deleteParagraphs     word.setList
word.insertPageBreak         word.navigate             word.scroll
word.scaleFont               word.addTable             word.setTableCell
word.formatTable             word.editTable            word.setPageLayout
word.setHeaderFooter         word.setDocumentText

slides.inspect               slides.inspectObjects     slides.replaceText
slides.scaleFont             slides.formatText         slides.formatSelection
slides.addSlide              slides.duplicateSlide     slides.deleteSlide
slides.addTextBox            slides.addImage           slides.setBackground
slides.addShape
slides.updateShape           slides.deleteObject       slides.inspectCharts
slides.addChart              slides.updateChart        slides.deleteChart

sheets.inspect               sheets.setValues          sheets.setFormula
sheets.inspectRange          sheets.setArrayFormula    sheets.replaceText
sheets.formatRange           sheets.manageRange        sheets.setRichText
sheets.addSheet              sheets.renameSheet        sheets.deleteSheet
sheets.manageSheet           sheets.inspectNames       sheets.manageNames
sheets.recalculate           sheets.sort               sheets.filter
sheets.inspectTables         sheets.manageTable
sheets.manageConditionalFormat
sheets.manageValidation      sheets.inspectPivots      sheets.managePivot
sheets.addChart              sheets.inspectCharts      sheets.updateChart
sheets.deleteChart           sheets.inspectDrawings    sheets.manageDrawing
sheets.manageHyperlink       sheets.inspectComments    sheets.manageComments
sheets.inspectFreezePanes    sheets.manageFreezePanes  sheets.inspectProperties
sheets.manageProperties      sheets.inspectProtectedRanges
sheets.manageProtectedRanges sheets.inspectPageLayout  sheets.managePageLayout
sheets.inspectMacros         sheets.setMacros
```

完整参数类型在 `public-api.d.ts`；机器可读 schema 在 `public-api.json`。另一个 TypeScript 工程可以复制这两个文件，或从 DocumentServer 静态地址下载并固定到版本 `0.6.0`。DOCX D01-D70 的逐项状态与公开 API 边界见 `DOCX-CAPABILITIES.zh-CN.md`；PPTX P01-P77 见 `PPTX-CAPABILITIES.zh-CN.md`；XLSX X01-X78 见 `XLSX-CAPABILITIES.zh-CN.md`。

### Word 完整操作边界

- `word_inspect` 返回正文、选区、当前页/可见页、带一基序号的段落结构、表格尺寸，并可按需返回批注。
- 精确文本操作包括替换、删除、字符样式、超链接、批注和书签；`occurrence` 用于只处理第 N 次命中。
- 段落操作支持标题/命名样式、字体、对齐、段前段后、行距、缩进、同页/孤行控制、分页符和项目符号/编号。
- 表格操作支持创建填充、样式、单元格格式、增删行列、合并/拆分、清空和删除。
- 页面操作支持纸张大小、横竖向、页边距、页眉页脚及动态页码。
- 高级操作覆盖分节/分栏、字符与段落样式、自定义编号、重复表头和嵌套表格、图形/图表/公式/OLE、目录/题注/交叉引用/脚注尾注、修订、保护、内容控件、自定义 XML、宏和水印。完整边界见 `DOCX-CAPABILITIES.zh-CN.md`。
- `word_scroll` 是稳定的按页上滚/下滚；`word_navigate` 还支持首页、末页、指定页、上一页、下一页、相对页和搜索定位。ONLYOFFICE Office API 没有跨版本稳定的像素级鼠标滚轮契约，因此 Bridge 不注入浏览器鼠标事件，也不会把按页导航误报成像素滚动。

### DOCX / PPTX / XLSX 外部图片

图片来源类型：

```ts
type AiBridgeImageSource =
  | { type: "url"; url: string }
  | { type: "dataUrl"; dataUrl: string };
```

在 Word 当前光标插图，或通过一基段落序号/文本命中定位：

```js
await window.aiBridge.word.addImage({
  source: { type: "url", url: "https://cdn.example.com/architecture.png" },
  widthMm: 120,
  search: "系统架构",
  occurrence: 1,
  wrapping: "square",
  name: "系统架构图",
});
```

`current`、`paragraphIndex`、`search` 最多指定一种；未指定时使用当前光标。
`wrapping` 默认为 `inline`，还可使用 `square`、`tight`、`through`、
`topAndBottom`、`behind`、`inFront`。只给宽或高时会按原始比例推导另一边；
同时给出宽高且 `preserveAspectRatio` 未关闭时，会在该边界框内 contain。

向指定 PPT 页插入从剪贴板 Blob 转出的 Data URL：

```js
await window.aiBridge.slides.addImage({
  slide: 2,
  source: { type: "dataUrl", dataUrl: clipboardImageDataUrl },
  widthMm: 150,
  rotationDeg: 2,
  flipH: false,
  name: "产品截图",
});

await window.aiBridge.slides.addImageShape({
  slide: 2,
  source: { type: "url", url: "https://cdn.example.com/avatar.svg" },
  shapeType: "ellipse",
  widthMm: 45,
  heightMm: 45,
  preserveAspectRatio: false,
  line: { widthPt: 1.5, color: "#336699" },
  name: "圆形头像",
});
```

PPT 未给尺寸时按原比例放进 `160 × 90 mm` 边界框；未给坐标时在页面居中。
返回值是标准 drawing 描述，后续可用 `slides.inspectObjects()` 定位，并用
`slides.deleteObject()` 删除。

Host 在发送编辑命令前，会把批次中所有 URL/Data URL 一次性预导入
`POST /copilot-api/images/import`。全部成功后才创建 checkpoint 并调用插件；
任一图片失败时整个批次不会修改文档。插件仅接受 Host 从服务端签名资源解析出的
内部图片来源，外部调用者不能直接传 `_image`。正常部署把签名同源 URL 交给
编辑器；localhost demo 为保持 ONLYOFFICE 的私网请求过滤器开启，会由可信 Host
通过签名 URL 读取资源后转换成内部 Data URL。相同内容按 SHA-256 去重；相同
`requestId` 的重试不会重复插图。

导入服务仅接受 PNG、JPEG、GIF、WebP、SVG，原图最大 8 MiB、最大边长 12,000 px、
最大 40 MP。SVG 会先解析并规范化，拒绝脚本、事件处理器、外部资源、DTD 和实体。
外部 URL 必须为 HTTPS，不能包含用户名密码，最多跟随 3 次重定向，
每次解析都拒绝回环、私网、链路本地、保留、多播和未指定地址，总超时 20 秒。
可用 `COPILOT_IMAGE_ALLOWED_HOSTS=cdn.example.com,images.example.com`
进一步限制域名。下载签名绑定资源、文档身份和过期时间，默认 15 分钟有效；临时
文件 24 小时后清理。图片写入 DOCX/PPTX 后由 ONLYOFFICE 内嵌到文件包，不依赖
临时 URL。

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
    "version": "0.6.0",
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
- `INVALID_TOOL_ARGUMENTS`：整批 `word_*` 参数或 `slides_add_table.data` 在修改前校验失败；读取 `details.validationErrors[]` 一次修正全部错误，此时 `completedToolCalls=0` 且 `partialMutationPossible=false`。
- `INVALID_ARGUMENTS` / `INVALID_COMMAND`：运行期参数语义或命令结构不符合契约。
- `INVALID_IMAGE_SOURCE`：图片来源类型、Data URL 或内部资源不合法。
- `IMAGE_FETCH_BLOCKED` / `IMAGE_FETCH_FAILED`：URL 被网络策略拦截或下载失败。
- `IMAGE_TOO_LARGE` / `UNSUPPORTED_IMAGE_FORMAT`：图片字节、像素或格式不符合限制。
- `IMAGE_ASSET_EXPIRED`：签名下载地址已过期。
- `IMAGE_API_UNSUPPORTED`：当前 ONLYOFFICE 运行时缺少必需图片 API。
- `WORD_API_UNSUPPORTED`：当前 ONLYOFFICE 运行时缺少对应的 Word API。
- `PERSISTENCE_FAILED`：Office API 已执行，但示例持久化服务保存失败，应提示用户并核对 callback/force-save。
- `TIMEOUT` / `CONNECTION_TIMEOUT`：调用或 Relay 握手超时。

HTTP Relay 不再把所有编辑器失败统一映射为 `409`。请求或协议错误使用
`400`，认证错误使用 `401`，安全策略拒绝使用 `403`，资源不存在或过期使用
`404/410`，只有文档、编辑器、会话或 `requestId` 的状态冲突使用 `409`。
参数语义或目标错误使用 `422`，未知编辑器异常使用 `500`，ONLYOFFICE
能力缺失使用 `501`，下载、Relay 或持久化等下游失败使用 `502`，编辑器未就绪
使用 `503`，超时使用 `504`。图片或参数体积过大使用 `413`，图片格式不支持
使用 `415`。无论 HTTP 状态如何，调用方都应优先读取响应
`error.code`、`error.message` 和经过过滤的 `error.details`。

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
- 图片错误不得记录或回显 Base64、签名 URL、远程响应正文；生产反向代理仅对
  `/images/import` 和 `/bridge/execute` 放宽到 12 MiB。

## 本地 Relay 观测日志

本地 Relay 输出两类单行 JSON 日志：

- `[bridge-command]`：每个首次完成、失败或超时的请求只记录一次，包含
  `requestId`、`sessionId`、文档身份摘要、`method`、`toolCount`、
  `queueWaitMs`、`editorRoundTripMs` 和 `totalMs`。幂等缓存命中不会重复记录。
- `[bridge-startup]`：编辑器页面注册时记录 host script、window load、编辑器
  iframe、`onAppReady`、`onDocumentReady` 和 Bridge ready 相对导航开始的毫秒数。

这两类日志不记录工具参数、正文、返回内容、JWT、Relay key 或错误正文。启动数据
只在页面与本地 Relay 之间传输，不加入公开 `getState()` 或 sessions 响应。

## 版本兼容

当前插件版本为 `0.6.0`，消息协议版本为 `1`，本次构建缓存键为
`?v=0.6.0-rev4`。生产页面应固定到实际发布的不可变缓存键，升级前先比较
`public-api.json`。协议版本不一致时 Client 和 Relay 不建立连接。
