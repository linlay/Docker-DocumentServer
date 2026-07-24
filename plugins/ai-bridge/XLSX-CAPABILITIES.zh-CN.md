# XLSX X01–X78 能力矩阵

适用版本：

- ONLYOFFICE DocumentServer `9.4.0-129`
- ai-bridge `0.4.0`
- Spreadsheet API 以 ONLYOFFICE 9.4 官方 Office JavaScript API 为稳定执行边界

状态说明：

- `✅`：可直接实现，ai-bridge 已提供第一方工具。
- `🟨`：部分实现，或需要工具组合、宿主文档服务、特定许可/对象支持。
- `⚪`：ONLYOFFICE 编辑器 UI 有此能力，但 9.4 文档化 Spreadsheet/Plugin API 没有稳定入口；ai-bridge 不伪造成功。
- `⛔`：ONLYOFFICE 9.4 本身不提供等价 Excel 能力。

“Bridge 入口”列使用外部快捷方法名；对应工具名可由驼峰转换为
`sheets_*`，完整参数以 `public-api.d.ts` 和 `public-api.json` 为准。

## 工作簿与工作表

| ID | 能力 | ONLYOFFICE 9.4 | ai-bridge 0.4.0 | Bridge 入口与说明 |
| --- | --- | --- | --- | --- |
| X01 | 工作簿创建、打开、保存 | ✅ | 🟨 | `save()` 保存当前工作簿；创建、打开和切换文件属于宿主文档服务，不由当前文档内插件完成 |
| X02 | 工作表创建、删除、复制、重命名 | ✅ | ✅ | `addSheet`、`deleteSheet`、`renameSheet`、`manageSheet({action:"copy"})` |
| X03 | 调整工作表顺序 | ✅ | ✅ | `manageSheet({action:"moveBefore"})` |
| X04 | 隐藏工作表 | ✅ | ✅ | `manageSheet({action:"setVisibility"})`，并保护“至少一个可见表”约束 |
| X05 | 工作表标签颜色 | ✅ | ⚪ | 9.4 文档化 Spreadsheet API 无标签颜色方法 |
| X06 | 工作簿和工作表属性 | ✅ | 🟨 | `inspectProperties`、`manageProperties` 支持核心/自定义工作簿属性；工作表名称、序号、可见性可读写；不含标签颜色 |

## 单元格与区域

| ID | 能力 | ONLYOFFICE 9.4 | ai-bridge 0.4.0 | Bridge 入口与说明 |
| --- | --- | --- | --- | --- |
| X07 | 单元格文本 | ✅ | ✅ | `setValues`、`inspectRange` |
| X08 | 整数和小数 | ✅ | ✅ | `setValues` |
| X09 | 日期和时间 | ✅ | ✅ | `setValues` 支持 `{type:"date"|"datetime"|"time",value:...}`；配合 `formatRange.numberFormat` |
| X10 | 百分比 | ✅ | ✅ | `setValues` 支持 `{type:"percent",value:"25%"}` 或小数；配合百分比格式 |
| X11 | 布尔值 | ✅ | ✅ | JSON 布尔值或 `{type:"boolean",value:...}` |
| X12 | 单元格区域 | ✅ | ✅ | 所有区域工具支持 A1 地址与 `selection` |
| X13 | 合并单元格 | ✅ | ✅ | `manageRange({action:"merge"|"unmerge"})` |
| X14 | 插入、删除、移动行列 | ✅ | ✅ | `manageRange` 的 `insert`、`delete`、`cut`；整行/整列使用行列地址 |
| X15 | 行高和列宽 | ✅ | ✅ | `formatRange.rowHeight`、`columnWidth` |
| X16 | 自动调整行高、列宽 | ✅ | ✅ | `manageRange({action:"autoFit",rows,columns})` |
| X17 | 隐藏行列 | ✅ | ✅ | `manageRange({action:"setHidden"})` |
| X18 | 行列分组与折叠 | ✅ | ⚪ | 9.4 文档化 Spreadsheet API 无 outline/group/ungroup 方法 |

## 单元格样式

| ID | 能力 | ONLYOFFICE 9.4 | ai-bridge 0.4.0 | Bridge 入口与说明 |
| --- | --- | --- | --- | --- |
| X19 | 字体、字号、颜色 | ✅ | ✅ | `formatRange` |
| X20 | 粗体、斜体、下划线 | ✅ | ✅ | `formatRange`；同时支持删除线 |
| X21 | 单元格背景填充 | ✅ | ✅ | `formatRange.fillColor` |
| X22 | 单元格边框 | ✅ | ✅ | `formatRange.borders[]` |
| X23 | 水平和垂直对齐 | ✅ | ✅ | `formatRange.horizontalAlign`、`verticalAlign` |
| X24 | 自动换行 | ✅ | ✅ | `formatRange.wrap` |
| X25 | 文本旋转 | ✅ | ✅ | `formatRange.orientation` |
| X26 | 数字格式 | ✅ | ✅ | `formatRange.numberFormat` |
| X27 | 日期、货币、百分比格式 | ✅ | ✅ | 内置或自定义 `numberFormat` |
| X28 | 自定义数字格式 | ✅ | ✅ | 任意 ONLYOFFICE 接受的格式代码 |
| X29 | 单元格内局部富文本 | ✅ | ✅ | `setRichText` 按字符位置设置内容和字体 |

## 计算功能

| ID | 能力 | ONLYOFFICE 9.4 | ai-bridge 0.4.0 | Bridge 入口与说明 |
| --- | --- | --- | --- | --- |
| X30 | 公式 | ✅ | ✅ | `setFormula` |
| X31 | 函数 | ✅ | ✅ | 通过公式调用 ONLYOFFICE 支持的工作表函数 |
| X32 | 跨工作表引用 | ✅ | ✅ | 公式和定义名称可使用带工作表名的引用 |
| X33 | 外部工作簿引用 | 🟨 | 🟨 | `setFormula` 可写外部引用语法；是否解析/刷新取决于文件、来源和 ONLYOFFICE 支持，Bridge 不管理外部链接 |
| X34 | 绝对、相对和混合引用 | ✅ | ✅ | 原样写入 `$A$1`、`A$1`、`$A1` |
| X35 | 数组公式 | ✅ | ✅ | `setArrayFormula` |
| X36 | 动态数组 | ✅ | ✅ | `setFormula` 写入 9.4 支持的动态数组公式；溢出结果由引擎计算 |
| X37 | 命名区域 | ✅ | ✅ | `inspectNames`、`manageNames` |
| X38 | 公式计算及缓存值 | ✅ | ✅ | `recalculate`；`inspectRange` 同时返回公式和计算值 |
| X39 | 自动、手动计算模式 | ✅ | ⚪ | 9.4 文档化 Spreadsheet API 可触发重算，但无稳定的计算模式读写方法 |

## 数据组织

| ID | 能力 | ONLYOFFICE 9.4 | ai-bridge 0.4.0 | Bridge 入口与说明 |
| --- | --- | --- | --- | --- |
| X40 | 排序 | ✅ | ✅ | `sort`，最多三个排序键 |
| X41 | 自动筛选 | ✅ | ✅ | `filter` 的 `set`、`showAll`、`reapply` |
| X42 | Excel 格式化表格 | ✅ | ✅ | `inspectTables`、`manageTable` 支持创建、格式化、更新、缩放、删除及转普通区域 |
| X43 | 表格样式 | ✅ | ✅ | 可检查并更新名称、样式、标题/汇总行、首末列、条纹、筛选按钮和替代文本；`ApiListObject` 高级操作依许可可用 |
| X44 | 高级筛选 | 🟨 | 🟨 | 支持多条件自动筛选；不提供 Excel“高级筛选并复制到另一位置”的独立命令 |
| X45 | 分类汇总 | ✅ | 🟨 | 可用 `SUBTOTAL` 公式、表格汇总行实现计算；无 UI 分类汇总/大纲生成命令 |
| X46 | 数据分组 | ✅ | ⚪ | 与 X18 相同，公开 API 无行列 outline/group 入口 |

## 数据规则

| ID | 能力 | ONLYOFFICE 9.4 | ai-bridge 0.4.0 | Bridge 入口与说明 |
| --- | --- | --- | --- | --- |
| X47 | 条件格式 | ✅ | ✅ | `manageConditionalFormat` |
| X48 | 数据验证 | ✅ | ✅ | `manageValidation` |
| X49 | 下拉列表 | ✅ | ✅ | `manageValidation` 使用 `xlValidateList` |
| X50 | 输入提示和错误提示 | ✅ | ✅ | 输入/错误标题、消息和显示开关 |
| X51 | 重复值和自定义公式规则 | ✅ | ✅ | `duplicateValues`、`uniqueValues`、`expression` |

## 数据分析

| ID | 能力 | ONLYOFFICE 9.4 | ai-bridge 0.4.0 | Bridge 入口与说明 |
| --- | --- | --- | --- | --- |
| X52 | 数据透视表 | ✅ | ✅ | `inspectPivots`、`managePivot` 支持创建、字段、值字段、样式、刷新和清空 |
| X53 | 数据透视图 | 🟨 | 🟨 | 可对透视结果区域创建普通图表；无独立 PivotChart 对象 API |
| X54 | 迷你图 | ✅ | ⚪ | 9.4 文档化 Spreadsheet API 无 sparkline 方法 |
| X55 | Power Query | ⛔ | ⛔ | ONLYOFFICE 9.4 无 Excel Power Query 引擎 |
| X56 | 外部数据连接 | 🟨 | ⚪ | ONLYOFFICE 可保留部分外链/连接信息，但 9.4 公开 API 无连接枚举、凭证和刷新入口 |
| X57 | 数据模型 | ⛔ | ⛔ | 无 Excel Power Pivot/Data Model 等价引擎 |

## 图形与对象

| ID | 能力 | ONLYOFFICE 9.4 | ai-bridge 0.4.0 | Bridge 入口与说明 |
| --- | --- | --- | --- | --- |
| X58 | 图片 | ✅ | ✅ | `manageDrawing({action:"addImage"})`；HTTPS/Data URL 先由宿主页安全导入 |
| X59 | 图表 | ✅ | ✅ | `inspectCharts`、`addChart`、`updateChart`、`deleteChart` |
| X60 | 形状 | ✅ | ✅ | `manageDrawing({action:"addShape"})` |
| X61 | 文本框 | ✅ | ✅ | `manageDrawing({action:"addTextBox"})` |
| X62 | 超链接 | ✅ | ✅ | `manageHyperlink` |
| X63 | 批注、备注 | ✅ | ✅ | `inspectComments`、`manageComments`，含回复与 solved 状态 |
| X64 | 复选框和表单控件 | ⛔ | ⛔ | ONLYOFFICE 9.4 Spreadsheet 不提供 Excel 表单控件等价能力 |
| X65 | 嵌入对象 | ✅ | ✅ | `manageDrawing({action:"addOleObject"})`；要求目标 OLE 插件 `appId`、数据和预览图 |

## 窗口与打印

| ID | 能力 | ONLYOFFICE 9.4 | ai-bridge 0.4.0 | Bridge 入口与说明 |
| --- | --- | --- | --- | --- |
| X66 | 冻结窗格 | ✅ | ✅ | `inspectFreezePanes`、`manageFreezePanes` |
| X67 | 拆分窗口 | ⛔ | ⛔ | ONLYOFFICE 9.4 无 Excel 拆分窗口等价能力 |
| X68 | 页面尺寸和方向 | ✅ | 🟨 | `managePageLayout.orientation` 已实现；9.4 Spreadsheet API 无纸张尺寸方法 |
| X69 | 页边距 | ✅ | ✅ | `managePageLayout.margins` |
| X70 | 打印区域 | ✅ | ⚪ | UI 可设置，公开 API 无 print-area 方法 |
| X71 | 打印标题行列 | ✅ | ⚪ | `printHeadings` 仅表示打印行号/列标，不冒充重复标题行列；公开 API 无 print-titles 方法 |
| X72 | 分页符 | ✅ | ⚪ | 公开 API 无分页符集合/插入方法 |
| X73 | 打印缩放 | ✅ | ⚪ | 公开 API 无 fit-to-page/scale 方法 |
| X74 | 页眉、页脚 | ✅ | ⚪ | 9.4 Spreadsheet API 无页眉页脚方法 |
| X75 | 页码 | ✅ | ⚪ | 依赖页眉页脚，公开 API 无稳定入口 |

## 安全与扩展

| ID | 能力 | ONLYOFFICE 9.4 | ai-bridge 0.4.0 | Bridge 入口与说明 |
| --- | --- | --- | --- | --- |
| X76 | 工作表保护 | ✅ | 🟨 | `inspectProtectedRanges`、`manageProtectedRanges` 支持协作受保护区域和用户权限；不等同于带密码的整表保护 |
| X77 | 工作簿保护 | 🟨 | ⚪ | 9.4 文档化 Spreadsheet API 无工作簿结构/窗口保护方法；文件级访问控制应由宿主服务负责 |
| X78 | 宏 | 🟨 | 🟨 | `inspectMacros`、`setMacros` 管理 ONLYOFFICE JavaScript 宏；可读取 VBA 源信息，但 ONLYOFFICE 不执行 VBA |

## ai-bridge 0.4.0 新增的 Sheets 能力域

除原有值、公式、格式、工作表和图表工具外，0.4.0 新增：

- 区域深度检查、数组公式、行列/区域操作、富文本。
- 工作表复制/顺序/可见性、定义名称、公式与透视表重算。
- 排序、筛选、结构化表格完整生命周期、条件格式、数据验证。
- 数据透视表检查与管理。
- 图片、形状、文本框、OLE、超链接和批注。
- 冻结窗格、工作簿属性、受保护区域和公开的页面布局属性。
- ONLYOFFICE JavaScript 宏读写和 VBA 宏源信息检查。

运行时会检查目标方法是否真实存在。社区版或特定许可缺少的方法会返回明确错误，
不会把“方法不存在”“返回 `false`”或“对象未创建”报告为成功。
