# ai-bridge v1 接入说明

生产环境只支持 document-hub v1。document-hub 负责身份、文档、权限、回调、
版本和公开 API；ai-bridge Python 服务只负责私网 Relay、参数预检和图片资源。

旧 UUID 文档应用、公开 attach/binding、`/chat`、匿名 editor-config 和 Python
持久化接口均已停用，不应由反向代理重新暴露。

## 1. 宿主页配置

创建 ONLYOFFICE 编辑器后，宿主页必须显式提供以下配置：

```js
window.aiBridgeOptions = {
  httpRelay: true,
  relayBaseUrl: "/api/v1/editor-relay",
  imageBaseUrl: "/api/v1/editor-relay/images",
  persistenceBaseUrl: "/api/v1/editor-relay/persistence",
  editorSessionId: config.sessionId,
  documentId: config.documentId,
  getEditorConfig: () => config,
  clientOrigins: ["https://app.example.com"],
};
```

然后加载：

```html
<script src="https://docs.example.com/sdkjs-plugins/{A17E5F31-64AA-4E37-9A42-8D430814C2F6}/host-bridge.js?v=<asset-revision>"></script>
```

三个 base URL 必须是同源路径；HTTP Relay 不再根据 localhost 自动启用。
`clientOrigins` 只能填写完整 origin，不能使用通配符。

## 2. 直接调用

```js
await window.aiBridge.ready();

const state = await window.aiBridge.getState();
const inspected = await window.aiBridge.word.inspect({ maxChars: 5000 });

await window.aiBridge.executeBatch([
  { name: "manage_revisions", arguments: { action: "start", displayMode: "edit" } },
  { name: "replace_text", arguments: { search: "旧文本", replace: "新文本" } },
]);
```

审阅工作流默认先调用 `manage_revisions({action:"start", displayMode:"edit"})`：
`start/stop` 只控制是否记录修订，`setDisplay` 只控制 `edit/simple/final/original`
显示。Bridge 会等待 ONLYOFFICE 显示模式回调成功后才执行同批后续修改；显示失败时
整批立即停止。AI 修改完成后不自动调用 `stop`。

公开批量工具名称不带 `word_`、`slides_`、`sheets_` 前缀。兼容字段
`args`、`action` 和别名 `window.onlyofficeAI` 继续保留。

## 3. postMessage SDK

```html
<script src="https://docs.example.com/sdkjs-plugins/{A17E5F31-64AA-4E37-9A42-8D430814C2F6}/client-sdk.js?v=<asset-revision>"></script>
```

```js
const client = new AiBridgeClient({
  targetWindow: editorFrame.contentWindow,
  targetOrigin: "https://docs.example.com",
});

await client.ready();
const result = await client.word.inspect({ maxChars: 5000 });
```

## 4. Relay 与持久化

浏览器页面使用 `register/poll/result/unregister` 维持 Relay 会话，并把启动故障和运行期
`fetch` 拒绝分别上报到 `startup-failure`、`runtime-failure`。运行期诊断包含请求序号、
耗时、连续失败次数、在线状态、页面可见性和最近的页面生命周期事件，不包含正文。
document-hub 通过私网 secret 调用：

- `/bridge/internal/execute`
- `/bridge/internal/validate`
- `/bridge/internal/session`
- `/bridge/internal/drop`
- `/bridge/internal/images/import`

document-hub 根据已认证的稳定文档身份确定 `relaySessionId`。不存在公开
attach 或 binding token 交换。

`save/history/undo/redo` 必须经宿主页转发到 document-hub persistence
服务；未配置时返回 `PERSISTENCE_NOT_AVAILABLE`，不会回退到旧 Python 路径。

## 5. 契约与生成

`public-api.json` 是工具 Schema、限制、错误码、能力和 transport 元数据的
唯一来源。修改后执行：

```bash
python3 tools/sync_contract.py --write
python3 tools/sync_contract.py --check
```

生成器会同步 `plugin.js`、`public-api.d.ts`、插件版本和浏览器资源 revision。
不要手工修改生成区。

## 6. 验证

```bash
PYTHONDONTWRITEBYTECODE=1 python3 -m unittest -v \
  test_copilot_server.py test_contract_generation.py
node --test \
  test_bridge_protocol.js \
  test_slides_sheets_bridges.js \
  test_sheets_advanced_bridge.js
```

跨仓变更还必须运行 document-hub Go 测试和前端构建。部署时端口 3001 只允许
Compose 私网访问，空的 `nginx-document-app.conf` 必须继续挂载以阻止旧路由。
