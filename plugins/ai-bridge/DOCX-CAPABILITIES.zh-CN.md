# ai-bridge DOCX 能力清单（D01–D70）

本文按 ONLYOFFICE Docs 9.4 的公开 Docs API、Office API 和插件方法核对
`ai-bridge` 的 DOCX 能力。机器可读参数以 `public-api.json` 为准，
TypeScript 参数以 `public-api.d.ts` 为准。

状态说明：

- **已实现**：`ai-bridge` 已有可调用工具或宿主控制接口。
- **有限实现**：公开 API 只能完成该功能的一部分，Bridge 已覆盖可可靠实现的部分。
- **待导出验收**：Bridge 路径和单元回归已接通，但尚未由真实 DocumentServer
  导出的 OOXML/逐页渲染样本证明，不按“已实现”对外承诺。
- **宿主负责**：属于编辑器实例或文件生命周期，不是在已打开文档中执行的插件命令。
- **公开 API 不支持**：ONLYOFFICE 编辑器界面可能具备该功能，但 9.4 的公开插件/Office API
  没有可靠的创建或修改方法，因此 Bridge 不伪造支持。

## 文档结构

| 编号 | 功能 | 状态 | ai-bridge 实现与边界 |
|---|---|---|---|
| D01 | 文档创建、打开、保存 | 宿主负责 / 已实现 | 创建、打开由 Docs API 的编辑器配置和业务文件服务负责；当前文档保存、历史、撤销、重做由 `save`、`history`、`undo`、`redo` 提供。 |
| D02 | 文档基本属性 | 已实现 | `word_inspect`、`word_inspect_advanced`、`word_set_document_properties`；支持核心属性、自定义属性写入，以及按名称读取自定义属性。 |
| D03 | 页面尺寸与方向 | 已实现 | `word_set_page_layout`；支持 A4、Letter、Legal、自定义宽高及横竖向。 |
| D04 | 页边距 | 已实现 | `word_set_page_layout`；支持上下左右页边距及页眉、页脚距离。 |
| D05 | 分页 | 已实现 | `word_insert_page_break`、段落 `pageBreakBefore`，以及页导航和页数检查。 |
| D06 | 分节 | 已实现 | `word_manage_section` 创建或配置分节，支持 continuous、nextPage、evenPage、oddPage；`CreateSection` 的目标是旧节结束段落，Bridge 通过 `boundary.position=before/after` 显式解析边界。 |
| D07 | 分栏（高级） | 已实现 | `word_manage_section.columns`；支持等宽和自定义非等宽分栏。公开 API 不提供分隔线参数。 |
| D08 | 不同节使用不同页面设置（高级） | 已实现 | `word_set_page_layout.sectionIndex` 与 `word_manage_section.sectionIndex`。 |

## 文本与段落

| 编号 | 功能 | 状态 | ai-bridge 实现与边界 |
|---|---|---|---|
| D09 | 正文文本 | 已实现 | `word_set_document_text`、`word_append_paragraph`、`word_insert_paragraph`、`word_replace_text`、`word_set_paragraph_text`。 |
| D10 | 段落 | 已实现 | `word_inspect`、`word_format_paragraphs`、`word_delete_paragraphs` 及段落插入、替换工具。 |
| D11 | 标题 | 已实现 | `headingLevel`、`outlineLevel` 和标题样式。 |
| D12 | 字体、字号、颜色 | 已实现 | 文档、段落、选区、文本命中和表格单元格均可设置。 |
| D13 | 粗体、斜体、下划线、删除线 | 已实现 | 支持粗体、斜体、下划线、单删除线、双删除线。 |
| D14 | 上标、下标 | 已实现 | `vertAlign` 支持 baseline、superscript、subscript。 |
| D15 | 段落对齐 | 已实现 | left、center、right、both。 |
| D16 | 缩进、行距、段前段后距 | 已实现 | 首行、左右缩进，自动/固定/最小行距，段前段后距。 |
| D17 | 字符样式和段落样式 | 已实现 | `word_manage_style` 创建 paragraph/character 样式；正文用 `styleName`，表格单元格段落用 `cellParagraphStyleName` / `paragraphStyleName`，字符用 `characterStyleName`。Bridge 校验真实样式类型，不能再把 paragraph 样式交给 `table.SetStyle()`。 |
| D18 | 自定义样式及样式继承（高级） | 已实现 | `word_manage_style.basedOn`；支持创建和更新公开 API 可编辑的样式属性。 |
| D19 | 制表位（高级） | 已实现 | `word_set_tabs`；支持位置及 left、center、right、decimal、bar、clear 对齐。公开 API 不提供前导符设置。 |
| D20 | 首字下沉（高级） | 公开 API 不支持 | ONLYOFFICE 9.4 的公开文档 API 没有可靠的 drop-cap 创建/修改方法。 |

## 列表

| 编号 | 功能 | 状态 | ai-bridge 实现与边界 |
|---|---|---|---|
| D21 | 项目符号列表 | 已实现 | `word_set_list`、`word_set_numbering`。 |
| D22 | 编号列表 | 已实现 | `word_set_list`、`word_set_numbering`。 |
| D23 | 多级列表（高级） | 已实现 | `word_set_numbering.levels` 支持 0–8 级；一次调用的 `assignments[]` 共用同一编号实例，`continueFrom` 复用已有定义。`inspect_advanced` 按 ONLYOFFICE 内部编号 ID 返回稳定 `listGroup`。旧 `paragraphIndexes + level` 每次调用仍是独立列表，不可用于跨层级连续编号。已通过 6 个父级、24 个子级的 ONLYOFFICE 9.4 导出验收，目标段落共享一个 OOXML `numId`。 |
| D24 | 自定义编号规则（高级） | 已实现 | 支持每级编号格式、格式文本、起始值、对齐、重启规则。 |

## 表格

| 编号 | 功能 | 状态 | ai-bridge 实现与边界 |
|---|---|---|---|
| D25 | 表格 | 已实现 | `word_add_table`、`word_inspect`。 |
| D26 | 插入、删除行列 | 已实现 | `word_edit_table` 的 addRow、addColumn、removeRow、removeColumn。 |
| D27 | 合并、拆分单元格 | 已实现 | `word_edit_table` 的 mergeCells、splitCell。 |
| D28 | 行高、列宽 | 已实现 | `word_format_table_advanced`；行高规则支持公开 API 的 auto、atLeast。 |
| D29 | 单元格边框、底色、对齐 | 已实现 | `word_set_table_cell` 和 `word_format_table_advanced`；指定 row/column 时边距和外边框只作用于目标单元格。 |
| D30 | 表格整体样式 | 已实现 | `word_format_table.tableStyleName` 仅接受 table 样式；`word_add_table.cellParagraphStyleName` 可把 paragraph 样式应用到全部单元格段落，单元格的 `paragraphStyleName` 可覆盖。 |
| D31 | 表头跨页重复（高级） | 已实现 | `word_format_table_advanced.repeatHeader`。 |
| D32 | 表格跨页控制（高级） | 公开 API 不支持 | 公开 `ApiTableRow` 没有“允许跨页断行”开关；可用段落 keepLines/keepNext 做有限的内容级控制，但不等价。 |
| D33 | 嵌套表格（高级） | 已实现 | `word_add_nested_table`。 |

## 图片与图形

| 编号 | 功能 | 状态 | ai-bridge 实现与边界 |
|---|---|---|---|
| D34 | 图片 | 已实现 | `word_add_image`；URL/Data URL 先经 Host 安全导入，再内嵌到文档。 |
| D35 | 图片缩放、裁剪、旋转 | 有限实现 | `word_add_image` 与 `word_manage_drawing` 支持等比缩放、指定尺寸和旋转；公开 API 不提供可靠裁剪接口。 |
| D36 | 图片边框、透明度 | 有限实现 | `word_manage_drawing.border` 支持边框；公开图片 API 不提供可靠透明度设置。 |
| D37 | 文字环绕和浮动定位（高级） | 已实现 | 支持 inline、square、tight、through、topAndBottom、behind、inFront，以及绝对/对齐定位。 |
| D38 | 形状（高级） | 已实现 | `word_add_shape`、`word_manage_drawing`。 |
| D39 | 文本框（高级） | 已实现 | 使用带文本内容的 `word_add_shape` 创建和格式化文本框型形状。 |
| D40 | 图表（高级） | 已实现 | `word_add_chart`；支持数据、系列、分类、数字格式、标题、图例、数据标签和环绕。 |
| D41 | SmartArt（高级） | 有限实现 | `word_inspect_advanced` 可识别，`word_manage_drawing` 可移动、缩放、旋转、改名或删除已有 SmartArt；公开 API 没有 SmartArt 创建工厂。 |
| D42 | 数学公式（高级） | 已实现 | `word_add_math`；支持 unicode、latex、mathml。 |
| D43 | 嵌入文件或 OLE 对象（高级） | 有限实现 | `word_add_ole_object`；需要当前 ONLYOFFICE 版本/许可证提供 `Api.CreateOleObject`，并由调用方提供应用数据及预览图。 |

## 页眉页脚

| 编号 | 功能 | 状态 | ai-bridge 实现与边界 |
|---|---|---|---|
| D44 | 页眉 | 已实现 | `word_set_header_footer(kind="header")`。 |
| D45 | 页脚 | 已实现 | `word_set_header_footer(kind="footer")`。 |
| D46 | 页码 | 已实现 | `pageNumber`、`pagesCount` 或通用字段指令。 |
| D47 | 首页不同 | 已实现 | `word_set_page_layout.titlePage` / `word_manage_section.titlePage`，配合 type=`first`。 |
| D48 | 奇偶页不同 | 已实现 | `word_manage_section.evenAndOddHeaders`，配合 type=`even`。 |
| D49 | 分节独立页眉页脚（高级） | 已实现 | `word_set_header_footer.sectionIndex` 为指定节创建、替换或删除 default/first/even 内容。 |
| D50 | 页码格式及指定起始页码 | 有限实现 | `word_manage_section.startPageNumber` 支持指定起始页码；公开 `ApiSection` 不提供页码数字格式设置。 |
| D51 | 页眉页脚中的动态字段（高级） | 已实现 | `fields[]` 是标准方式并保持显式顺序；PAGE/PAGES 布尔快捷参数仅兼容使用，Bridge 会插入分隔文本，避免字段粘连成“第 210”。 |

## 长文档功能

| 编号 | 功能 | 状态 | ai-bridge 实现与边界 |
|---|---|---|---|
| D52 | 超链接 | 已实现 | `word_add_hyperlink`，支持外部 URL 和书签目标。 |
| D53 | 书签（高级） | 已实现 | `word_add_bookmark`、`word_inspect_advanced`、`word_manage_long_document.deleteBookmark`。 |
| D54 | 自动目录（高级） | 已实现 | `word_manage_long_document` 的 addToc、updateToc。 |
| D55 | 题注（高级） | 待导出验收 | addCaption 直接插入真实 `SEQ` 域，支持样式、书签和表格/段落目标；只有导出 OOXML 中存在 `SEQ` 且 `ai/qa` 通过后才算能力通过。 |
| D56 | 图表目录（高级） | 已实现 | addTableOfFigures、updateTableOfFigures。 |
| D57 | 交叉引用（高级） | 待导出验收 | 书签引用可插入真实 `REF` 域并替换占位文本；其他引用仍覆盖 caption、heading、numbered、footnote、endnote 目标。只有导出 OOXML 的 `REF`/目标关系通过 QA 后才对外承诺。 |
| D58 | 脚注、尾注（高级） | 已实现 | addFootnote、addEndnote，并可检查已有注释文本。 |
| D59 | 引文和参考文献（高级） | 有限实现 | 可用 `word_manage_fields` 写入 CITATION/BIBLIOGRAPHY 等字段指令并更新字段；公开 API 不提供结构化文献源库的增删改查。 |
| D60 | 域代码及动态字段（高级） | 已实现 | `word_manage_fields` 支持按文本范围或段落插入字段、更新全部字段；页眉页脚也支持任意字段指令。 |

## 审阅与控制

| 编号 | 功能 | 状态 | ai-bridge 实现与边界 |
|---|---|---|---|
| D61 | 批注（高级） | 待导出验收 | 添加后即时检查 `GetQuoteText()` 并返回 `anchored/quoteText`；管理操作支持回复、编辑、删除、全部删除、解决和重新打开。最终还需导出 OOXML 同时存在 range start/end/reference。 |
| D62 | 修订记录（高级） | 已实现 | `word_manage_revisions` 的 `start`/`stop` 独立控制是否记录修订；可选 `displayMode=edit/simple/final/original` 通过稳定插件方法等待应用审阅显示，`setDisplay` 可只改变显示而不改变跟踪状态。默认 AI 审阅工作流使用 `start + edit`，修改后保持跟踪开启，并通过 `word_inspect_advanced` 读取 review report。 |
| D63 | 接受、拒绝修订（高级） | 有限实现 | 支持接受全部或拒绝全部；公开 API 没有按单条 revision ID 接受/拒绝的方法。 |
| D64 | 文档保护（高级） | 有限实现 | `word_set_protection` 支持 readOnly、comments、forms 和解除限制；公开插件方法不支持设置密码。 |
| D65 | 内容控件、复选框（高级） | 已实现 | 支持块级、行内、复选框、组合框、下拉列表、日期、图片内容控件，以及列表默认值、日期/格式、外观、更新、锁定、勾选、清空和删除。 |
| D66 | 宏（高级） | 已实现 | `word_inspect_macros` 读取 ONLYOFFICE/VBA 宏信息；`word_set_macros` 替换文档保存的 ONLYOFFICE 宏集合。Bridge 不自动执行新写入的宏。 |
| D67 | 自定义 XML（高级） | 已实现 | 支持 part 添加、替换、删除，XPath 元素/属性插入、更新、删除，以及内容控件 XML 数据绑定和重新同步。 |

## 页面装饰

| 编号 | 功能 | 状态 | ai-bridge 实现与边界 |
|---|---|---|---|
| D68 | 水印 | 已实现 | `word_set_watermark`；支持文字、图片、透明度、方向、比例和删除。 |
| D69 | 页面背景 | 公开 API 不支持 | ONLYOFFICE 9.4 的公开 `ApiDocument` / `ApiSection` 没有页面背景色设置方法。 |
| D70 | 页面边框（高级） | 公开 API 不支持 | 公开 `ApiSection` 没有 Word 页面边框设置方法；段落、表格、形状边框不能冒充页面边框。 |

## 对外 Word 工具

0.2.3 的 Bridge 内部 Word 工具共 50 个（Skill/HTTPX 公共 action 使用无前缀名称）：

```text
word_inspect                    word_replace_text
word_append_paragraph           word_insert_paragraph
word_format_document            word_format_selection
word_format_matches             word_delete_matches
word_add_hyperlink              word_add_comment
word_add_bookmark               word_add_image
word_inspect_advanced           word_set_document_properties
word_manage_section             word_manage_style
word_set_tabs                   word_set_numbering
word_format_table_advanced      word_add_nested_table
word_manage_drawing             word_add_shape
word_add_chart                  word_add_math
word_add_ole_object             word_manage_fields
word_manage_long_document       word_manage_comments
word_manage_revisions           word_set_protection
word_manage_content_control     word_manage_custom_xml
word_inspect_macros             word_set_macros
word_set_watermark              word_format_paragraphs
word_set_paragraph_text         word_delete_paragraphs
word_set_list                   word_insert_page_break
word_navigate                   word_scroll
word_scale_font                 word_add_table
word_set_table_cell             word_format_table
word_edit_table                 word_set_page_layout
word_set_header_footer          word_set_document_text
```

## 核对依据

- [ONLYOFFICE Document API](https://api.onlyoffice.com/docs/office-api/usage-api/document-api/)
- [ApiDocument](https://api.onlyoffice.com/docs/office-api/usage-api/document-api/ApiDocument/)
- [ApiSection](https://api.onlyoffice.com/docs/office-api/usage-api/document-api/ApiSection/)
- [ApiTable](https://api.onlyoffice.com/docs/office-api/usage-api/document-api/ApiTable/)
- [ApiDrawing](https://api.onlyoffice.com/docs/office-api/usage-api/document-api/ApiDrawing/)
- [ApiWatermarkSettings](https://api.onlyoffice.com/docs/office-api/usage-api/document-api/ApiWatermarkSettings/)
