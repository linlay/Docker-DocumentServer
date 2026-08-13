# PPTX P01–P77 能力矩阵

适用版本：

- ONLYOFFICE DocumentServer `9.4.0-129`
- ai-bridge `0.2.3`
- Presentation API 以 ONLYOFFICE 9.4 文档化 Office JavaScript API 和 Plugin API 为稳定执行边界

状态说明：

- `✅`：可直接实现，ai-bridge 已提供第一方工具或安全的工具组合；P50 SmartArt 使用锁定到 ONLYOFFICE 9.4 的原生 UI/数据模型扩展入口。
- `🟨`：部分实现，或创建/打开文件属于宿主服务，或 ONLYOFFICE 只提供部分等价能力。
- `⚪`：ONLYOFFICE 编辑器 UI 或文件引擎有此能力，但 9.4 文档化 Presentation/Plugin API 没有可验证入口；ai-bridge 不伪造成功。
- `⛔`：ONLYOFFICE 9.4 Presentation Editor 本身没有对应的 PowerPoint 能力。

“Bridge 入口”列使用 `window.aiBridge.slides` / Client SDK 的快捷方法名；底层工具名为
`slides_*`，完整参数以 `public-api.d.ts` 和 `public-api.json` 为准。

## 演示文稿结构

| ID | 能力 | ONLYOFFICE 9.4 | ai-bridge 0.2.3 | Bridge 入口与说明 |
| --- | --- | --- | --- | --- |
| P01 | 演示文稿创建、打开、保存 | ✅ | 🟨 | `save()` 保存当前演示文稿；创建、打开和切换文件属于宿主文档服务，不由当前文档内插件完成 |
| P02 | 新建、复制、删除幻灯片 | ✅ | ✅ | `addSlide` 可在创建时指定 `masterIndex/layoutIndex` 并返回实际版式；另有 `duplicateSlide`、`deleteSlide` |
| P03 | 调整幻灯片顺序 | ✅ | ✅ | `moveSlide` |
| P04 | 隐藏幻灯片 | ✅ | ✅ | `setVisibility` |
| P05 | 幻灯片尺寸 | ✅ | ✅ | `setSize`，支持宽屏、标准、自定义尺寸及横竖方向 |
| P06 | 幻灯片背景 | ✅ | ✅ | `setBackground`，支持纯色、渐变、图案、原生 stretch/tile 图片背景、清除及跟随版式/母版 |
| P07 | 幻灯片分节（高级） | 🟨 | ⚪ | 9.4 文件引擎可保留 PPTX section 数据，但编辑器没有稳定的分节管理 Presentation/Plugin API |

## 主题与版式

| ID | 能力 | ONLYOFFICE 9.4 | ai-bridge 0.2.3 | Bridge 入口与说明 |
| --- | --- | --- | --- | --- |
| P08 | 使用现有主题 | ✅ | ✅ | `inspectBuiltinThemes`、`applyBuiltinTheme`；文稿内主题使用 `inspectThemes`、`applyTheme` |
| P09 | 使用现有幻灯片版式 | ✅ | ✅ | `inspectLayouts`、`applyLayout` |
| P10 | 标题和内容占位符 | ✅ | ✅ | `createLayout.placeholders`、`addTemplateShape.placeholderType`、`setTextContent` |
| P11 | 自定义主题颜色和字体（高级） | ✅ | ✅ | `setTheme`，支持 12 色主题方案和主/次字体方案 |
| P12 | 编辑幻灯片母版（高级） | ✅ | ✅ | `inspectLayouts({includeObjects:true})`、`addTemplateShape`、`manageTemplateObject`、`setTemplateBackground`；母版和版式均支持原生图片背景 |
| P13 | 创建自定义版式（高级） | ✅ | ✅ | `createLayout`，可创建占位符并立即应用到指定幻灯片 |
| P14 | 备注母版、讲义母版（高级） | ⛔ | ⛔ | 9.4 可设置“备注和讲义”的页眉页脚打印选项，但没有 PowerPoint 等价的备注母版/讲义母版编辑器和公开对象模型 |

## 文本内容

| ID | 能力 | ONLYOFFICE 9.4 | ai-bridge 0.2.3 | Bridge 入口与说明 |
| --- | --- | --- | --- | --- |
| P15 | 标题 | ✅ | ✅ | `addSlide.title`、`addTextBox`、标题占位符 |
| P16 | 正文文本 | ✅ | ✅ | `setTextContent`、`replaceText` |
| P17 | 文本框 | ✅ | ✅ | `addTextBox` 支持普通 `text` 或原生富文本 `paragraphs`，两者互斥 |
| P18 | 字体、字号、颜色 | ✅ | ✅ | `formatText`、`formatSelection`、`setTextContent` |
| P19 | 粗体、斜体、下划线 | ✅ | ✅ | `formatText`、`formatSelection`、富文本段落 |
| P20 | 段落对齐、缩进和行距 | ✅ | ✅ | `formatParagraphs` |
| P21 | 项目符号列表 | ✅ | ✅ | `addTextBox/addShape.paragraphs` 或 `formatParagraphs({listType:"bullet"})` |
| P22 | 编号列表 | ✅ | ✅ | `addTextBox/addShape.paragraphs` 或 `formatParagraphs({listType:"number"})` |
| P23 | 多级列表（高级） | ✅ | ✅ | `formatParagraphs.level` 支持 0–8 级，并可分别设置项目符号或编号 |
| P24 | 艺术字和文本特效（高级） | ✅ | ✅ | `addWordArt`，支持文字变换、字体、填充、描边、旋转和尺寸 |
| P25 | 文本自动缩放和溢出控制（高级） | ✅ | ⚪ | UI 可设置文本自动适应；9.4 文档化 `ApiShape`/`ApiDrawing` 无 text-fit/overflow 读写方法 |

## 形状与布局

| ID | 能力 | ONLYOFFICE 9.4 | ai-bridge 0.2.3 | Bridge 入口与说明 |
| --- | --- | --- | --- | --- |
| P26 | 基本形状 | ✅ | ✅ | `addShape`，接受 ONLYOFFICE 预设几何类型 |
| P27 | 线条和箭头 | ✅ | ✅ | `addConnector` 或 `addShape`，线条样式由 `line` 设置 |
| P28 | 形状填充、边框 | ✅ | ✅ | `addShape`、`updateShape`，支持纯色、线性/径向渐变、图案和描边 |
| P29 | 对象位置和尺寸 | ✅ | ✅ | `updateObject`、各创建工具的 `xMm/yMm/widthMm/heightMm` |
| P30 | 对象旋转 | ✅ | ✅ | `updateObject.rotationDeg` |
| P31 | 对象层级 | ✅ | ✅ | `reorderObject` 的置顶、置底、上移、下移 |
| P32 | 对象对齐和等间距分布 | ✅ | ✅ | `alignObjects`，可相对选区或幻灯片对齐，并支持水平/垂直分布 |
| P33 | 对象组合和取消组合 | ✅ | ✅ | `groupObjects` |
| P34 | 连接线和连接点（高级） | ✅ | 🟨 | `addConnector` 支持直线、折线、曲线连接线及端点坐标；9.4 公开 API 不提供把端点绑定到图形连接点的稳定方法 |
| P35 | 自由形状和编辑顶点（高级） | ✅ | ✅ | `addFreeform` 支持直线、二次/三次贝塞尔曲线、圆弧和闭合路径 |
| P36 | 阴影、发光、三维效果（高级） | 🟨 | ⚪ | 9.4 UI 提供基础阴影；发光/三维支持不完整，文档化 Presentation API 也无效果对象读写入口 |

### 结构化检查边界

- `inspectObjects({includeTextStyles:true})` 返回可读取的段落与 run 样式；当前 ONLYOFFICE 版本没有 getter 的属性明确标记为不可用。
- `validateLayout` 只读检查版式、对象预算、命名、边界、安全边距、几何交叉、字号和手写列表前缀。
- `validateLayout.visualVerified` 与 `textOverflowVerified` 固定为 `false`。它不能替代截图、渲染或人工视觉验收，也不能证明自动换行和文字溢出正确。

## 图片

| ID | 能力 | ONLYOFFICE 9.4 | ai-bridge 0.2.3 | Bridge 入口与说明 |
| --- | --- | --- | --- | --- |
| P37 | 图片 | ✅ | ✅ | `addImage`；HTTPS/Data URL 均先经过宿主页受控导入 |
| P38 | 图片缩放、裁剪、旋转 | ✅ | 🟨 | `addImage`、`updateObject` 完整支持缩放和旋转；`addImageShape({shapeType:"rect"})` 可做矩形蒙版，但 9.4 公开 API 无原生 crop-offset 读写 |
| P39 | 图片透明度和边框 | ✅ | 🟨 | `updateObject.line` 支持图片边框；UI 有透明度，但公开 `ApiImage`/`ApiBlipFill` 无透明度设置方法 |
| P40 | 图片按形状裁剪 | ✅ | ✅ | `addImageShape`，支持圆形及任意预设几何，可同时设置边框 |
| P41 | SVG 和图标 | ✅ | ✅ | `addImage`、`addImageShape` 支持 SVG；SVG 在服务端拒绝脚本、事件、外部资源和 DTD/实体后再导入 |
| P42 | 删除背景（高级） | ⛔ | ⛔ | 9.4 Presentation Editor 无 PowerPoint“删除背景”的核心等价功能 |
| P43 | 图片艺术效果和重新着色（高级） | ⛔ | ⛔ | 9.4 没有 PowerPoint 图片艺术滤镜/重新着色的等价对象模型 |

## 表格与数据

| ID | 能力 | ONLYOFFICE 9.4 | ai-bridge 0.2.3 | Bridge 入口与说明 |
| --- | --- | --- | --- | --- |
| P44 | 表格 | ✅ | ✅ | `addTable` |
| P45 | 插入、删除表格行列 | ✅ | ✅ | `editTable` 的 `addRow`、`addColumn`、`removeRow`、`removeColumn` |
| P46 | 合并、拆分单元格 | ✅ | ✅ | `editTable` 的 `mergeCells`、`splitCell` |
| P47 | 表格边框、填充和对齐 | ✅ | ✅ | `setTableCell`、`formatTable` |
| P48 | 图表（高级） | ✅ | ✅ | `inspectCharts`、`addChart`、`updateChart`、`deleteChart`；覆盖 45 个规范类型/别名、组合图、系列与数据点格式、标题/坐标轴标题粗体、填充、线条、图例、标签和网格线 |
| P49 | 图表内嵌 Excel 数据（高级） | ✅ | ✅ | `addChart.series/categories` 和 `updateChart.seriesUpdates` 写入图表内嵌工作簿数据，并支持系列/数据点数值格式及 `allSeries` 批量点格式 |
| P50 | SmartArt（高级） | ✅ | ✅ | `inspectSmartarts`、`addSmartart`、`updateSmartart`、`deleteSmartart`；创建 151 种 9.4 原生预设，读写逻辑节点文本/样式并同步回 SmartArt 数据模型。升级 DocumentServer 时必须重新回归版本锁定的原生扩展入口 |
| P51 | 数学公式（高级） | ✅ | ✅ | `addMath` 支持 LaTeX、Unicode 和 MathML |

## 动画与切换

| ID | 能力 | ONLYOFFICE 9.4 | ai-bridge 0.2.3 | Bridge 入口与说明 |
| --- | --- | --- | --- | --- |
| P52 | 幻灯片切换（高级） | ✅ | ✅ | `setTransition`，支持效果、速度、时长和换片条件 |
| P53 | 对象进入、强调、退出动画（高级） | ✅ | ✅ | `manageAnimation({action:"add"})`，`effectType` 使用 ONLYOFFICE 9.4 动画效果名 |
| P54 | 动画顺序和计时（高级） | ✅ | ✅ | `inspectAnimations`、`manageAnimation` 支持触发方式、时长、延迟、重复次数和顺序 |
| P55 | 路径动画（高级） | ✅ | ✅ | `manageAnimation.effectType` 支持 9.4 公布的圆、方形、菱形、心形、星形、方向路径等 |
| P56 | 动画触发器（高级） | ✅ | ✅ | `sequence:"interactive"` 与 `triggerObject` 创建对象点击触发序列 |
| P57 | Morph 平滑切换（高级） | ✅ | ✅ | `setTransition.effect` 使用 `effectMorphByObject`、`effectMorphByWord` 或 `effectMorphByChar` |

## 多媒体

| ID | 能力 | ONLYOFFICE 9.4 | ai-bridge 0.2.3 | Bridge 入口与说明 |
| --- | --- | --- | --- | --- |
| P58 | 音频（高级） | ✅ | ⚪ | UI 可插入音频；9.4 文档化 Presentation/Plugin API 没有 `CreateAudio`/`AddAudio` |
| P59 | 视频（高级） | ✅ | ⚪ | UI 可插入视频；9.4 文档化 Presentation/Plugin API 没有 `CreateVideo`/`AddVideo` |
| P60 | 动态 GIF（高级） | ✅ | ✅ | `addImage` 安全导入 GIF；动画由 ONLYOFFICE 放映引擎处理 |
| P61 | 音视频裁剪及播放设置（高级） | ✅ | ⚪ | UI 有媒体播放设置，但没有稳定的媒体对象和播放参数 Plugin API |
| P62 | 字幕（高级） | ⛔ | ⛔ | 9.4 无 PowerPoint 媒体字幕轨道编辑能力 |
| P63 | 屏幕录制内容（高级） | ⛔ | ⛔ | 9.4 无 PowerPoint 屏幕录制工具；已有录屏文件可在 UI 中作为普通视频插入 |

## 导航与放映

| ID | 能力 | ONLYOFFICE 9.4 | ai-bridge 0.2.3 | Bridge 入口与说明 |
| --- | --- | --- | --- | --- |
| P64 | 超链接 | ✅ | ✅ | `setHyperlink` 支持 HTTPS/HTTP、邮件、FTP 和放映导航链接 |
| P65 | 动作按钮（高级） | ✅ | ✅ | `addShape` 创建 actionButton 预设形状，再用 `setHyperlink` 设置动作；可在同一批次原子执行 |
| P66 | 跳转到指定幻灯片（高级） | ✅ | ✅ | `setHyperlink({action:"slide",targetSlide:n})`；实时放映可用 `controlSlideshow({action:"goto"})` |
| P67 | 自定义放映（高级） | ⛔ | ⛔ | 9.4 无 PowerPoint“自定义放映”集合的 UI/API 等价能力 |
| P68 | 自动换片和循环播放（高级） | ✅ | ✅ | `setTransition.advanceOnTime/advanceTimeMs`、`setShowSettings.loop` |
| P69 | 放映计时（高级） | ✅ | 🟨 | 可读写每页自动换片时间和动画时间线；没有 PowerPoint“排练计时”录制会话 API |
| P70 | 旁白（高级） | ⛔ | ⛔ | 9.4 无 PowerPoint 旁白录制/轨道管理等价能力 |

## 页脚与协作

| ID | 能力 | ONLYOFFICE 9.4 | ai-bridge 0.2.3 | Bridge 入口与说明 |
| --- | --- | --- | --- | --- |
| P71 | 幻灯片编号 | ✅ | ⚪ | UI 可设置；公开 API 可创建 `sldNum` 占位符，但没有稳定的动态页码/页眉页脚属性入口，Bridge 不把静态文本冒充页码 |
| P72 | 日期和时间 | ✅ | ⚪ | UI 可设置固定/自动日期；9.4 文档化 Presentation/Plugin API 无页眉页脚日期属性入口 |
| P73 | 页脚文字 | ✅ | ⚪ | UI 可设置；公开 API 可创建 `ftr` 占位符，但不能稳定设置演示文稿页脚属性 |
| P74 | 演讲者备注 | ✅ | ✅ | `setNotes`；`inspect` 返回备注文本 |
| P75 | 批注（高级） | ✅ | ✅ | `addComment`、`inspectComments`、`manageComment`，支持回复、solved 状态、位置、作者及删除 |
| P76 | 嵌入文件或 OLE 对象（高级） | ✅ | ✅ | `addOleObject`；要求 OLE `appId`、对象数据和经安全导入的预览图 |
| P77 | 宏（高级） | 🟨 | 🟨 | `inspectMacros`、`setMacros` 管理 ONLYOFFICE JavaScript 宏；可读取 VBA 宏源信息，但 ONLYOFFICE 不执行 VBA |

## 结论

在 P01–P77 中：

- `55` 项已由 ai-bridge 直接实现。
- `6` 项因宿主边界或 ONLYOFFICE 公开 API 只覆盖部分能力而部分实现。
- `9` 项在 ONLYOFFICE UI/文件引擎中存在，但 9.4 没有可验证的自动化入口。
- `7` 项在 ONLYOFFICE 9.4 中没有 PowerPoint 等价能力。

因此，ONLYOFFICE 9.4 **通过稳定 Presentation/Plugin API 可实现的能力均已接入 ai-bridge**，
并额外提供版本锁定的原生 SmartArt 扩展。
运行时会继续检查目标方法是否真实存在；方法缺失、返回 `false` 或对象未创建都会返回明确错误，
不会把“UI 能做但 API 不开放”报告为自动化成功。

本轮补齐的主要能力包括：

- 主题库、主题颜色/字体、母版、版式、占位符及母版对象完整增改删查。
- 富文本段落、多级列表、艺术字、公式、自由形状、组合/层级/对齐。
- 表格行列、合并/拆分、完整图表模型及内嵌数据。
- 151 种原生 SmartArt 的创建、检查、节点文本/样式更新和删除。
- SVG 安全导入、图片按形状裁剪、图片边框、OLE。
- 切换、Morph、对象动画时间线、交互触发器。
- 批注回复/状态管理、宏读写及放映控制。
