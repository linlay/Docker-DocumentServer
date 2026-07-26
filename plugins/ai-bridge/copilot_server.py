#!/usr/bin/env python3
"""Planning and persistence API for the headless ONLYOFFICE ai-bridge plugin."""

from __future__ import annotations

import base64
import binascii
import glob
import hashlib
import hmac
import html
import ipaddress
import json
import math
import os
import re
import secrets
import shutil
import socket
import struct
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid
import xml.etree.ElementTree as ET
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from typing import Any


PORT = int(os.environ.get("COPILOT_PORT", "3001"))
LOCAL_CONFIG = "/etc/onlyoffice/documentserver/local.json"
EXAMPLE_FILES = "/var/lib/onlyoffice/documentserver-example/files"
EXAMPLE_STORAGE_ID = "185.199.108.133"
DOCUMENT_TEMPLATE_ROOT = (
    "/var/www/onlyoffice/documentserver/document-templates/new/zh-CN"
)
DOCUMENT_TYPES = {
    "docx": "word",
    "xlsx": "cell",
    "pptx": "slide",
}
DOCUMENT_UUID_PATTERN = re.compile(
    r"^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$"
)
DOCUMENT_FILE_PATTERN = re.compile(
    r"^(?P<documentId>[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-"
    r"[89ab][0-9a-f]{3}-[0-9a-f]{12})\.(?P<fileType>docx|xlsx|pptx)$"
)
PUBLIC_DOCUMENT_PATH_PATTERN = re.compile(
    r"^/(?P<fileType>docx|xlsx|pptx)/"
    r"(?P<documentId>[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-"
    r"[89ab][0-9a-f]{3}-[0-9a-f]{12})$"
)
NEW_DOCUMENT_PATH_PATTERN = re.compile(r"^/new-(?P<fileType>docx|xlsx|pptx)$")
DOCUMENT_CREATE_RATE_PER_SECOND = 10 / 60
DOCUMENT_CREATE_BURST = 20
DOCUMENT_RATE_LIMIT_LOCK = threading.Lock()
DOCUMENT_RATE_LIMITS: dict[str, tuple[float, float]] = {}
MAX_VERSION_SNAPSHOTS = 20
BRIDGE_SESSION_TTL_SECONDS = 180
BRIDGE_RESUME_TOKEN_TTL_SECONDS = 12 * 60 * 60
BRIDGE_BINDING_TOKEN_TTL_SECONDS = 12 * 60 * 60
BRIDGE_DISCOVERY_WAIT_SECONDS = 10
BRIDGE_RESULT_TTL_SECONDS = 300
BRIDGE_PENDING_TTL_SECONDS = 660
BRIDGE_MAX_POLL_SECONDS = 30
BRIDGE_REJECTED_CREDENTIAL_TTL_SECONDS = 12 * 60 * 60
BRIDGE_REGISTRATION_CREDENTIAL_ERRORS = {
    "EDITOR_TOKEN_REQUIRED",
    "INVALID_EDITOR_TOKEN",
    "EDITOR_TOKEN_EXPIRED",
    "INVALID_BRIDGE_RESUME_TOKEN",
    "BRIDGE_RESUME_TOKEN_EXPIRED",
    "DOCUMENT_MISMATCH",
    "EDITOR_MISMATCH",
    "DOCUMENT_IDENTITY_MISMATCH",
    "INCOMPLETE_BRIDGE_IDENTITY",
}
BRIDGE_ALLOWED_METHODS = {
    "executeTool",
    "executeBatch",
    "save",
    "history",
    "undo",
    "redo",
    "getState",
}
BRIDGE_ID_PATTERN = re.compile(r"^[A-Za-z0-9._:-]{1,200}$")
BRIDGE_LOG_PHASES = {"word-command", "editor-save"}
BRIDGE_STARTUP_TIMING_FIELDS = (
    "hostScriptMs",
    "windowLoadMs",
    "editorFrameSeenMs",
    "editorFrameLoadMs",
    "appReadyMs",
    "documentReadyMs",
    "bridgeReadyMs",
)
LOCAL_GUEST_DISPLAY_NAME = "访客"
LOCAL_GUEST_USER_PREFIX = "local-guest:"
IMAGE_MAX_BYTES = 8 * 1024 * 1024
IMAGE_MAX_REQUEST_BYTES = 12_000_000
IMAGE_MAX_EDGE_PX = 12_000
IMAGE_MAX_PIXELS = 40_000_000
IMAGE_FETCH_TIMEOUT_SECONDS = 20
IMAGE_MAX_REDIRECTS = 3
IMAGE_URL_TTL_SECONDS = 15 * 60
IMAGE_RETENTION_SECONDS = 24 * 60 * 60
IMAGE_ASSET_PATTERN = re.compile(r"^[0-9a-f]{64}\.(png|jpg|gif|webp|svg)$")
IMAGE_MIME_EXTENSIONS = {
    "image/png": "png",
    "image/jpeg": "jpg",
    "image/gif": "gif",
    "image/webp": "webp",
    "image/svg+xml": "svg",
}
IMAGE_SOURCE_TOOLS = {
    "word_add_image",
    "slides_add_image",
    "slides_add_image_shape",
    "slides_add_ole_object",
    "word_add_ole_object",
    "word_set_watermark",
    "sheets_manage_drawing",
}
BRIDGE_CONDITION = threading.Condition()
BRIDGE_SESSIONS: dict[str, dict[str, Any]] = {}
BRIDGE_AUTHORITATIVE: dict[tuple[str, str, str, str], str] = {}
BRIDGE_SUPERSEDED: dict[str, dict[str, Any]] = {}
BRIDGE_COMMANDS: dict[str, dict[str, Any]] = {}
BRIDGE_REQUESTS: dict[tuple[tuple[str, str, str, str], str], str] = {}
BRIDGE_REJECTED_CREDENTIALS: dict[str, dict[str, Any]] = {}
BRIDGE_GENERATION = 0


def tool(name: str, description: str, properties: dict[str, Any], required: list[str] | None = None) -> dict[str, Any]:
    return {
        "type": "function",
        "function": {
            "name": name,
            "description": description,
            "parameters": {
                "type": "object",
                "properties": properties,
                "required": required or [],
                "additionalProperties": False,
            },
        },
    }


STRING = {"type": "string"}
NUMBER = {"type": "number"}
BOOLEAN = {"type": "boolean"}
SHEET = {"type": "string", "description": "工作表名称；省略时使用活动工作表"}
RANGE = {"type": "string", "description": "A1 表示法，例如 A1:D10；使用 selection 表示当前选区"}


WORD_TOOL_DESCRIPTIONS = {
    "word_inspect": "读取 Word 正文、选区、页码、段落结构、表格和可选批注；不修改文档",
    "word_replace_text": "在 Word 全文中查找并替换文本",
    "word_append_paragraph": "在 Word 文末追加带字符、段落、标题或列表样式的段落",
    "word_insert_paragraph": "在 Word 当前光标处插入带样式的段落",
    "word_format_document": "统一设置 Word 全文的字符与段落格式",
    "word_format_selection": "只格式化 Word 当前选中文本",
    "word_format_matches": "搜索 Word 文本并只格式化精确命中范围",
    "word_delete_matches": "删除 Word 中指定的文本命中范围",
    "word_add_hyperlink": "给 Word 中指定的文本命中范围添加超链接",
    "word_add_comment": "给 Word 中指定的文本命中范围添加批注",
    "word_add_bookmark": "给 Word 中精确指定的一处文本添加书签",
    "word_add_image": "把 HTTPS URL 或 Base64 Data URL 图片安全导入 Word，可按光标、段落序号或文本命中定位并设置尺寸与环绕",
    "word_inspect_advanced": "只读检查 Word 文档属性、节、样式、编号、绘图、书签、脚注尾注、批注、修订、内容控件和自定义 XML",
    "word_set_document_properties": "设置 Word 文档标题、主题、作者等文档属性",
    "word_manage_section": "创建或配置 Word 文档节、分节类型、起始页码和分栏",
    "word_manage_style": "创建或更新 Word 段落、字符、表格或编号样式及其继承关系",
    "word_set_tabs": "设置 Word 段落制表位",
    "word_set_numbering": "设置 Word 段落的高级编号属性",
    "word_format_table_advanced": "设置 Word 表格的高级布局、边框、间距和单元格属性",
    "word_add_nested_table": "在 Word 表格单元格中添加嵌套表格",
    "word_manage_drawing": "定位并更新或删除 Word 绘图对象",
    "word_add_shape": "在 Word 中添加带可选文本的形状并设置尺寸、样式、旋转和环绕",
    "word_add_chart": "在 Word 中添加图表并设置数据和样式",
    "word_add_math": "在 Word 中添加数学公式对象",
    "word_add_ole_object": "在 Word 中添加受控 OLE 对象",
    "word_manage_fields": "添加 Word 动态字段、更新全部字段或清除表单字段",
    "word_manage_long_document": "管理 Word 目录、题注、图表目录、交叉引用、脚注尾注和书签",
    "word_manage_comments": "管理 Word 批注线程和状态",
    "word_manage_revisions": "启停修订模式，或接受、拒绝全部 Word 修订",
    "word_set_protection": "设置或移除 Word 文档保护",
    "word_manage_content_control": "添加、更新、检查或删除 Word 内容控件，支持列表、日期、外观和自定义 XML 数据绑定",
    "word_manage_custom_xml": "管理 Word 自定义 XML 部件及 XPath 元素、属性",
    "word_inspect_macros": "只读检查 Word 的 ONLYOFFICE 或 VBA 宏",
    "word_set_macros": "替换 Word 文档中保存的 ONLYOFFICE 宏集合",
    "word_set_watermark": "设置或移除 Word 水印",
    "word_format_paragraphs": "按段落序号、文本或当前光标设置完整段落样式",
    "word_set_paragraph_text": "替换指定段落的文本并可同时设置样式",
    "word_delete_paragraphs": "删除按序号、文本或当前光标选中的段落",
    "word_set_list": "把指定段落设置为项目符号或编号列表",
    "word_insert_page_break": "在指定段落之前或之后插入分页符",
    "word_navigate": "只读跳转到首页、末页、指定页、相对页或搜索命中位置",
    "word_scroll": "只读按页向上或向下滚动 Word 视图",
    "word_scale_font": "按比例缩放 Word 全文字号；0.8 表示缩小 20%",
    "word_add_table": "在 Word 文末创建并填充表格",
    "word_set_table_cell": "设置指定 Word 表格单元格的文本和格式",
    "word_format_table": "设置指定 Word 表格的宽度、样式、外观和单元格格式",
    "word_edit_table": "增删 Word 表格行列，或合并、拆分、清空、删除表格",
    "word_set_page_layout": "设置 Word 节的纸张、方向、页边距和页眉页脚距离",
    "word_set_header_footer": "设置或删除 Word 默认、首页或偶数页的页眉页脚",
    "word_set_document_text": "用纯文本重写整个 Word 文档；仅在用户明确要求时使用",
}


def resolve_contract_schema(value: Any, definitions: dict[str, Any]) -> Any:
    if isinstance(value, list):
        return [resolve_contract_schema(item, definitions) for item in value]
    if not isinstance(value, dict):
        return value
    reference = value.get("$ref")
    if isinstance(reference, str) and reference.startswith("#/$defs/"):
        definition_name = reference.removeprefix("#/$defs/")
        if definition_name not in definitions:
            raise RuntimeError(f"public-api.json 引用了不存在的定义：{reference}")
        merged = dict(definitions[definition_name])
        merged.update({key: item for key, item in value.items() if key != "$ref"})
        return resolve_contract_schema(merged, definitions)
    return {
        key: resolve_contract_schema(item, definitions)
        for key, item in value.items()
    }


def load_contract_tools(editor: str, descriptions: dict[str, str]) -> list[dict[str, Any]]:
    configured_path = os.environ.get("COPILOT_PUBLIC_API")
    candidates = [
        configured_path,
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "public-api.json"),
        "/var/www/onlyoffice/documentserver/sdkjs-plugins/"
        "{A17E5F31-64AA-4E37-9A42-8D430814C2F6}/public-api.json",
    ]
    contract_path = next((path for path in candidates if path and os.path.isfile(path)), None)
    if not contract_path:
        raise RuntimeError("找不到 ai-bridge public-api.json")
    with open(contract_path, encoding="utf-8") as stream:
        contract = json.load(stream)
    schemas = contract.get("tools", {}).get(editor)
    if not isinstance(schemas, dict) or not schemas:
        raise RuntimeError(f"public-api.json 缺少 {editor} 工具契约")
    definitions = contract.get("$defs", {})
    missing_descriptions = sorted(set(schemas) - set(descriptions))
    if missing_descriptions:
        raise RuntimeError(f"{editor} 工具缺少模型描述：{', '.join(missing_descriptions)}")
    return [
        {
            "type": "function",
            "function": {
                "name": name,
                "description": descriptions[name],
                "parameters": resolve_contract_schema(schema, definitions),
            },
        }
        for name, schema in schemas.items()
    ]


WORD_TOOLS = load_contract_tools("word", WORD_TOOL_DESCRIPTIONS)


SLIDE_TOOL_DESCRIPTIONS = {
    "slides_inspect": "只读检查 PPT 页数与文本摘要；修改前优先调用",
    "slides_inspect_layouts": "只读检查 PPT 母版和版式",
    "slides_inspect_themes": "只读检查 PPT 主题、主题颜色和主题字体",
    "slides_inspect_builtin_themes": "只读检查 ONLYOFFICE 编辑器内置主题库",
    "slides_inspect_objects": "只读检查 PPT 图形、图表、图片、表格、组合等对象及填充、线条、位置和可选原始 JSON",
    "slides_replace_text": "在 PPT 全部或指定幻灯片中查找替换",
    "slides_scale_font": "按比例缩放 PPT 全部或指定页字号",
    "slides_format_text": "格式化 PPT 全部或指定页文本",
    "slides_format_selection": "格式化 PPT 当前选中的文本框或形状",
    "slides_add_slide": "新建幻灯片",
    "slides_duplicate_slide": "复制指定幻灯片",
    "slides_delete_slide": "删除指定幻灯片",
    "slides_move_slide": "移动指定幻灯片到新的页序",
    "slides_set_visibility": "显示或隐藏指定幻灯片",
    "slides_set_size": "设置演示文稿页面尺寸和方向",
    "slides_apply_layout": "给指定幻灯片应用母版版式",
    "slides_set_show_settings": "设置幻灯片放映循环选项",
    "slides_apply_theme": "把现有 PPT 主题应用到指定幻灯片",
    "slides_apply_builtin_theme": "把 ONLYOFFICE 编辑器内置主题应用到当前演示文稿",
    "slides_set_theme": "创建或更新 PPT 主题及其颜色、字体方案",
    "slides_create_layout": "在 PPT 母版中创建自定义幻灯片版式",
    "slides_add_template_shape": "在 PPT 母版或版式中添加模板图形或占位符",
    "slides_manage_template_object": "更新或删除 PPT 母版、版式中的模板对象",
    "slides_set_template_background": "设置 PPT 母版或版式的背景填充",
    "slides_add_textbox": "在指定 PPT 页添加可设置位置、尺寸、文本、渐变填充和线条的文本框；尺寸单位为毫米",
    "slides_add_word_art": "在指定 PPT 页添加艺术字并设置变换、字体、填充、线条和位置",
    "slides_add_math": "在指定 PPT 页插入 LaTeX、Unicode 或 MathML 数学公式",
    "slides_add_image": "把 HTTPS URL 或 Base64 Data URL 图片安全导入指定 PPT 页，可设置位置、尺寸、旋转、翻转和名称",
    "slides_add_image_shape": "把安全导入的图片填充到圆形或其他预设形状中，实现按形状裁剪并可设置边框",
    "slides_add_ole_object": "在指定 PPT 页添加带安全预览图的 OLE 对象",
    "slides_add_connector": "在指定 PPT 页添加连接线",
    "slides_add_freeform": "在指定 PPT 页添加自由形状",
    "slides_group_objects": "组合或取消组合指定 PPT 绘图对象",
    "slides_align_objects": "对齐或分布指定 PPT 绘图对象",
    "slides_reorder_object": "调整指定 PPT 绘图对象的层级顺序",
    "slides_set_text_content": "替换指定 PPT 对象的富文本段落内容",
    "slides_format_paragraphs": "设置指定 PPT 文本对象的段落格式、缩进、间距和列表",
    "slides_update_object": "按对象 ID、序号或名称更新 PPT 绘图对象的通用属性",
    "slides_set_hyperlink": "给 PPT 文本或对象设置或移除超链接",
    "slides_set_notes": "设置或清除指定幻灯片的演讲者备注",
    "slides_add_comment": "在指定幻灯片添加批注",
    "slides_inspect_comments": "只读检查 PPT 批注及其回复",
    "slides_manage_comment": "更新、回复、删除 PPT 批注或批注回复",
    "slides_set_transition": "设置或清除指定幻灯片的切换效果和计时",
    "slides_inspect_animations": "只读检查 PPT 对象动画时间线、序列和计时",
    "slides_manage_animation": "添加、更新、排序、删除或清空 PPT 对象动画",
    "slides_add_table": "在指定幻灯片添加表格",
    "slides_set_table_cell": "设置 PPT 表格单元格内容和格式",
    "slides_format_table": "设置 PPT 表格整体及单元格格式",
    "slides_edit_table": "增删 PPT 表格行列，或合并、拆分和删除表格",
    "slides_set_background": "设置、清除或恢复 PPT 页背景；自定义背景支持纯色、线性渐变、径向渐变、图案和 raw 填充",
    "slides_add_shape": "在指定 PPT 页添加任意预设图形，并设置文本、几何、渐变填充、线条、旋转和内边距",
    "slides_update_shape": "按 objectId、objectIndex 或 name 更新 PPT 图形的文本、类型、位置、尺寸、旋转、填充和线条",
    "slides_delete_object": "按 objectId、objectIndex 或 name 删除 PPT 页中的任意对象",
    "slides_inspect_charts": "只读检查 PPT 图表类型、标题、系列、位置、样式和可选原始 JSON",
    "slides_add_chart": "用数值系列和分类在 PPT 页创建图表，并设置系列、坐标轴、图例、标签、渐变填充和位置",
    "slides_update_chart": "按 chartId、chartIndex 或 name 更新 PPT 图表系列、分类、坐标轴、图例、标签、颜色、位置和尺寸",
    "slides_delete_chart": "按 chartId、chartIndex 或 name 删除 PPT 图表",
    "slides_inspect_macros": "只读检查 PPT 的 ONLYOFFICE 宏或 VBA 宏",
    "slides_set_macros": "设置或清除 PPT 的 ONLYOFFICE 宏内容",
    "slides_control_slideshow": "启动、结束、暂停、继续或导航当前 PPT 放映",
}

SHEET_TOOL_DESCRIPTIONS = {
    "sheets_inspect": "只读检查 XLSX 工作表、使用区域和值；修改前优先调用",
    "sheets_set_values": "向 XLSX 单元格区域写入单个值或二维数组",
    "sheets_set_formula": "向 XLSX 单元格写入公式",
    "sheets_replace_text": "在 XLSX 使用区域或指定区域查找替换",
    "sheets_format_range": "设置 XLSX 区域字体、填充、对齐、数字格式和行列尺寸",
    "sheets_add_sheet": "新建工作表",
    "sheets_rename_sheet": "重命名工作表",
    "sheets_delete_sheet": "删除工作表",
    "sheets_add_chart": "基于区域数据创建 XLSX 图表，并设置系列、坐标轴、图例、标签、渐变填充和位置",
    "sheets_inspect_charts": "只读检查 XLSX 图表类型、标题、系列、位置、样式和可选原始 JSON",
    "sheets_update_chart": "按 chartIndex 或 name 更新 XLSX 图表系列、数据区域、坐标轴、图例、标签、颜色、位置和尺寸",
    "sheets_delete_chart": "按 chartIndex 或 name 删除 XLSX 图表；部分 ONLYOFFICE 版本/许可不提供 ApiDrawing.Delete，此时会明确失败",
    "sheets_inspect_range": "只读检查 XLSX 指定区域的值、公式和格式",
    "sheets_manage_range": "管理 XLSX 区域的合并、行列插删、复制、剪切、清除、填充、隐藏和自动调整",
    "sheets_set_array_formula": "给 XLSX 区域设置数组公式",
    "sheets_set_rich_text": "给 XLSX 单元格设置富文本内容",
    "sheets_sort": "对 XLSX 区域按指定字段排序",
    "sheets_filter": "设置、更新或清除 XLSX 区域筛选",
    "sheets_inspect_tables": "只读检查 XLSX 工作表中的结构化表格、范围、样式和显示属性",
    "sheets_manage_validation": "管理 XLSX 单元格数据验证",
    "sheets_manage_conditional_format": "管理 XLSX 条件格式规则",
    "sheets_manage_table": "创建、格式化、更新、缩放、删除结构化表格，或把表格转换为普通区域",
    "sheets_manage_hyperlink": "设置或移除 XLSX 单元格超链接",
    "sheets_manage_comments": "添加、更新或删除 XLSX 单元格批注",
    "sheets_inspect_comments": "只读检查 XLSX 单元格批注",
    "sheets_manage_names": "创建、更新或删除 XLSX 定义名称",
    "sheets_inspect_names": "只读检查 XLSX 定义名称",
    "sheets_manage_freeze_panes": "冻结或取消冻结 XLSX 窗格",
    "sheets_inspect_freeze_panes": "只读检查 XLSX 冻结窗格状态",
    "sheets_manage_page_layout": "设置 XLSX 打印页面布局",
    "sheets_inspect_page_layout": "只读检查 XLSX 打印页面布局",
    "sheets_manage_properties": "设置 XLSX 工作簿属性",
    "sheets_inspect_properties": "只读检查 XLSX 工作簿属性",
    "sheets_manage_sheet": "管理 XLSX 工作表的可见性、顺序和其他属性",
    "sheets_manage_protected_ranges": "创建或更新 XLSX 受保护区域，并管理其可编辑用户",
    "sheets_inspect_protected_ranges": "只读检查 XLSX 受保护区域",
    "sheets_manage_drawing": "添加、更新或删除 XLSX 绘图对象",
    "sheets_inspect_drawings": "只读检查 XLSX 绘图对象",
    "sheets_manage_pivot": "创建、配置、刷新或清空 XLSX 数据透视表",
    "sheets_inspect_pivots": "只读检查 XLSX 数据透视表",
    "sheets_set_macros": "设置或清除 XLSX 宏内容",
    "sheets_inspect_macros": "只读检查 XLSX 宏信息",
    "sheets_recalculate": "触发 XLSX 工作簿重新计算",
}

SLIDE_TOOLS = load_contract_tools("slide", SLIDE_TOOL_DESCRIPTIONS)
SHEET_TOOLS = load_contract_tools("cell", SHEET_TOOL_DESCRIPTIONS)


TOOLS_BY_EDITOR = {"word": WORD_TOOLS, "slide": SLIDE_TOOLS, "cell": SHEET_TOOLS}
PREFIX_BY_EDITOR = {"word": "word_", "slide": "slides_", "cell": "sheets_"}
ALLOWED_BY_EDITOR = {
    editor: {entry["function"]["name"] for entry in entries}
    for editor, entries in TOOLS_BY_EDITOR.items()
}


def compact_json(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


class BridgeError(Exception):
    def __init__(self, status: int, code: str, message: str, details: Any = None):
        super().__init__(message)
        self.status = status
        self.code = code
        self.message = message
        self.details = details
        self.suppress_log = False

    def payload(self) -> dict[str, Any]:
        error: dict[str, Any] = {"code": self.code, "message": self.message}
        if self.details is not None:
            error["details"] = self.details
        return {"ok": False, "error": error}


EDITOR_ERROR_HTTP_STATUS = {
    "INVALID_COMMAND": 400,
    "INVALID_TOOL_CALL": 400,
    "INVALID_REQUEST_ID": 400,
    "TOO_MANY_CALLS": 400,
    "TOOL_NOT_ALLOWED": 400,
    "METHOD_NOT_ALLOWED": 400,
    "MESSAGE_NOT_SUPPORTED": 400,
    "EDITOR_TOKEN_REQUIRED": 401,
    "INVALID_EDITOR_TOKEN": 401,
    "EDITOR_TOKEN_EXPIRED": 401,
    "INVALID_BRIDGE_BINDING_TOKEN": 401,
    "BRIDGE_BINDING_TOKEN_EXPIRED": 401,
    "INVALID_BRIDGE_RESUME_TOKEN": 401,
    "BRIDGE_RESUME_TOKEN_EXPIRED": 401,
    "INVALID_RELAY_SESSION": 401,
    "CONTROL_NOT_ALLOWED": 403,
    "IMAGE_FETCH_BLOCKED": 403,
    "SESSION_NOT_FOUND": 404,
    "COMMAND_NOT_FOUND": 404,
    "IMAGE_ASSET_EXPIRED": 410,
    "DOCUMENT_MISMATCH": 409,
    "DOCUMENT_IDENTITY_MISMATCH": 409,
    "EDITOR_MISMATCH": 409,
    "SESSION_ID_CONFLICT": 409,
    "SESSION_NOT_AUTHORITATIVE": 409,
    "SESSION_SUPERSEDED": 409,
    "REQUEST_ID_CONFLICT": 409,
    "REQUEST_IN_FLIGHT": 409,
    "MULTIPLE_ACTIVE_EDITORS": 409,
    "ARGUMENTS_TOO_LARGE": 413,
    "IMAGE_TOO_LARGE": 413,
    "UNSUPPORTED_IMAGE_FORMAT": 415,
    "INVALID_ARGUMENTS": 422,
    "INVALID_TARGET": 422,
    "INVALID_IMAGE_SOURCE": 422,
    "EXECUTION_FAILED": 500,
    "WORD_API_UNSUPPORTED": 501,
    "IMAGE_API_UNSUPPORTED": 501,
    "PERSISTENCE_FAILED": 502,
    "IMAGE_FETCH_FAILED": 502,
    "HTTP_RELAY_INVALID_RESPONSE": 502,
    "NOT_READY": 503,
    "DOCUMENT_NOT_CONFIGURED": 503,
    "NO_ACTIVE_EDITOR": 503,
    "TIMEOUT": 504,
    "CONNECTION_TIMEOUT": 504,
    "BRIDGE_TIMEOUT": 504,
}


def editor_error_http_status(code: str) -> int:
    return EDITOR_ERROR_HTTP_STATUS.get(code, 500)


def verify_editor_jwt_payload(token: str) -> dict[str, Any]:
    try:
        encoded_header, encoded_payload, encoded_signature = token.split(".")
        unsigned = f"{encoded_header}.{encoded_payload}"
        expected = b64url(hmac.new(get_jwt_secret().encode(), unsigned.encode(), hashlib.sha256).digest())
        if not hmac.compare_digest(encoded_signature, expected):
            raise ValueError("signature")
        header = json.loads(base64.urlsafe_b64decode(encoded_header + "=" * (-len(encoded_header) % 4)))
        payload = json.loads(base64.urlsafe_b64decode(encoded_payload + "=" * (-len(encoded_payload) % 4)))
        if header.get("alg") != "HS256" or not isinstance(payload, dict):
            raise ValueError("claims")
        if float(payload.get("exp", 0)) < time.time():
            raise BridgeError(401, "EDITOR_TOKEN_EXPIRED", "当前编辑器凭证已过期，请重新绑定当前编辑器")
        document = payload.get("document") if isinstance(payload.get("document"), dict) else {}
        document_key = str(document.get("key", "")).strip()
        if not document_key:
            raise ValueError("document.key")
        return payload
    except BridgeError:
        raise
    except Exception as error:
        raise BridgeError(401, "INVALID_EDITOR_TOKEN", "无效的 ONLYOFFICE 当前编辑器凭证") from error


def verify_editor_jwt(token: str) -> dict[str, Any]:
    payload = verify_editor_jwt_payload(token)
    document = payload.get("document") if isinstance(payload.get("document"), dict) else {}
    editor_config = payload.get("editorConfig") if isinstance(payload.get("editorConfig"), dict) else {}
    user = editor_config.get("user") if isinstance(editor_config.get("user"), dict) else {}
    return {
        "documentKey": str(document.get("key", "")).strip(),
        "fileName": str(document.get("title", "")),
        "fileType": str(document.get("fileType", "")),
        "editorType": str(payload.get("documentType", "")),
        "userId": str(user.get("id", "")),
        "authKind": "editor",
    }


def normalize_local_guest_uuid(value: Any) -> str:
    candidate = str(value or "").strip().lower()
    try:
        parsed = uuid.UUID(candidate)
    except (AttributeError, ValueError) as error:
        raise BridgeError(
            400,
            "INVALID_ANONYMOUS_ID",
            "anonymousId 必须是规范的 UUID v4",
        ) from error
    if parsed.version != 4 or str(parsed) != candidate:
        raise BridgeError(
            400,
            "INVALID_ANONYMOUS_ID",
            "anonymousId 必须是规范的 UUID v4",
        )
    return candidate


def issue_anonymous_editor_config(payload: dict[str, Any]) -> dict[str, Any]:
    token = str(payload.get("editorToken", "")).strip()
    if not token:
        raise BridgeError(
            400,
            "EDITOR_TOKEN_REQUIRED",
            "缺少待重签的 ONLYOFFICE 编辑器凭证",
        )
    anonymous_id = normalize_local_guest_uuid(payload.get("anonymousId"))
    editor_payload = verify_editor_jwt_payload(token)
    editor_config = editor_payload.get("editorConfig")
    if not isinstance(editor_config, dict):
        raise BridgeError(
            401,
            "INVALID_EDITOR_TOKEN",
            "ONLYOFFICE 编辑器凭证缺少 editorConfig",
        )
    customization = editor_config.get("customization")
    if not isinstance(customization, dict):
        customization = {}
        editor_config["customization"] = customization
    anonymous = customization.get("anonymous")
    if not isinstance(anonymous, dict):
        anonymous = {}
        customization["anonymous"] = anonymous
    anonymous.update({"request": False, "label": LOCAL_GUEST_DISPLAY_NAME})
    user = {
        "group": "",
        "id": f"{LOCAL_GUEST_USER_PREFIX}{anonymous_id}",
        "image": "",
        "name": LOCAL_GUEST_DISPLAY_NAME,
        "roles": [],
    }
    editor_config["user"] = user
    expires_at = int(float(editor_payload["exp"]))
    return {
        "ok": True,
        "user": user,
        "token": sign_jwt(editor_payload, get_jwt_secret()),
        "expiresAt": expires_at,
    }


def bridge_identity(claims: dict[str, Any]) -> tuple[str, str, str, str]:
    identity = (
        str(claims.get("fileName", "")).strip(),
        str(claims.get("fileType", "")).strip(),
        str(claims.get("editorType", "")).strip(),
        str(claims.get("userId", "")).strip(),
    )
    if not all(identity):
        raise BridgeError(
            401,
            "INCOMPLETE_BRIDGE_IDENTITY",
            "当前编辑器凭证缺少稳定文档身份，无法建立 ai-bridge 绑定",
        )
    return identity


def bridge_identity_digest(identity: tuple[str, str, str, str]) -> str:
    return hashlib.sha256("\0".join(identity).encode()).hexdigest()[:12]


def bridge_startup_timings(state: Any) -> dict[str, int]:
    if not isinstance(state, dict):
        return {}
    diagnostics = state.get("_diagnostics")
    if not isinstance(diagnostics, dict):
        return {}
    startup = diagnostics.get("startup")
    if not isinstance(startup, dict):
        return {}
    timings: dict[str, int] = {}
    for field in BRIDGE_STARTUP_TIMING_FIELDS:
        value = startup.get(field)
        if not isinstance(value, (int, float)) or isinstance(value, bool):
            continue
        numeric = float(value)
        if not math.isfinite(numeric) or numeric < 0 or numeric > 3_600_000:
            continue
        timings[field] = round(numeric)
    return timings


def bridge_command_tool_count(command: dict[str, Any]) -> int:
    method = command.get("method")
    params = command.get("params") if isinstance(command.get("params"), dict) else {}
    if method == "executeTool":
        return 1
    if method == "executeBatch":
        tool_calls = params.get("toolCalls")
        return len(tool_calls) if isinstance(tool_calls, list) else 0
    return 0


def bridge_elapsed_ms(start: Any, end: Any) -> int | None:
    if not isinstance(start, (int, float)) or isinstance(start, bool):
        return None
    if not isinstance(end, (int, float)) or isinstance(end, bool):
        return None
    elapsed = (float(end) - float(start)) * 1000
    if not math.isfinite(elapsed):
        return None
    return max(0, round(elapsed))


def log_bridge_command(command: dict[str, Any], event: str, completed_monotonic: float) -> None:
    if command.get("timingLogged"):
        return
    command["timingLogged"] = True
    created_monotonic = command.get("createdMonotonic")
    first_delivered_monotonic = command.get("firstDeliveredMonotonic")
    event_payload = {
        "event": event,
        "requestId": command.get("requestId"),
        "sessionId": command.get("sessionId"),
        "identity": bridge_identity_digest(command["identity"]),
        "method": command.get("method"),
        "toolCount": bridge_command_tool_count(command),
        "queueWaitMs": bridge_elapsed_ms(created_monotonic, first_delivered_monotonic),
        "editorRoundTripMs": bridge_elapsed_ms(first_delivered_monotonic, completed_monotonic),
        "totalMs": bridge_elapsed_ms(created_monotonic, completed_monotonic),
    }
    print("[bridge-command] " + compact_json(event_payload), flush=True)


def image_asset_root() -> str:
    return os.environ.get(
        "COPILOT_IMAGE_DIR",
        os.path.join(EXAMPLE_FILES, ".ai-bridge-images"),
    )


def image_allowed_hosts() -> list[str]:
    return [
        value.strip().lower().rstrip(".")
        for value in os.environ.get("COPILOT_IMAGE_ALLOWED_HOSTS", "").split(",")
        if value.strip()
    ]


def image_error(status: int, code: str, message: str) -> BridgeError:
    return BridgeError(status, code, message)


def validate_image_url(value: str) -> str:
    try:
        parsed = urllib.parse.urlsplit(str(value))
    except ValueError as error:
        raise image_error(400, "INVALID_IMAGE_SOURCE", "图片 URL 无效") from error
    if parsed.scheme.lower() != "https" or not parsed.hostname:
        raise image_error(400, "INVALID_IMAGE_SOURCE", "外部图片只允许 HTTPS URL")
    if parsed.username or parsed.password:
        raise image_error(400, "INVALID_IMAGE_SOURCE", "图片 URL 不允许包含用户名或密码")
    hostname = parsed.hostname.lower().rstrip(".")
    allowed_hosts = image_allowed_hosts()
    if allowed_hosts and not any(
        hostname == allowed or hostname.endswith("." + allowed)
        for allowed in allowed_hosts
    ):
        raise image_error(403, "IMAGE_FETCH_BLOCKED", "图片来源域名不在允许列表中")
    try:
        addresses = socket.getaddrinfo(
            hostname,
            parsed.port or 443,
            type=socket.SOCK_STREAM,
        )
    except OSError as error:
        raise image_error(502, "IMAGE_FETCH_FAILED", "无法解析图片来源地址") from error
    if not addresses:
        raise image_error(502, "IMAGE_FETCH_FAILED", "无法解析图片来源地址")
    for address in addresses:
        try:
            ip = ipaddress.ip_address(address[4][0].split("%", 1)[0])
        except ValueError as error:
            raise image_error(403, "IMAGE_FETCH_BLOCKED", "图片来源地址不安全") from error
        if not ip.is_global:
            raise image_error(403, "IMAGE_FETCH_BLOCKED", "禁止从本机、私网或保留地址导入图片")
    return urllib.parse.urlunsplit(parsed)


class SafeImageRedirectHandler(urllib.request.HTTPRedirectHandler):
    def redirect_request(
        self,
        request: urllib.request.Request,
        fp: Any,
        code: int,
        msg: str,
        headers: Any,
        new_url: str,
    ) -> urllib.request.Request | None:
        redirect_count = int(getattr(request, "_ai_bridge_redirect_count", 0)) + 1
        if redirect_count > IMAGE_MAX_REDIRECTS:
            raise image_error(502, "IMAGE_FETCH_FAILED", "图片 URL 重定向次数过多")
        validated_url = validate_image_url(urllib.parse.urljoin(request.full_url, new_url))
        redirected = super().redirect_request(request, fp, code, msg, headers, validated_url)
        if redirected is not None:
            setattr(redirected, "_ai_bridge_redirect_count", redirect_count)
        return redirected


def sniff_image(data: bytes) -> tuple[str, str, int, int]:
    mime_type = ""
    extension = ""
    width = 0
    height = 0
    if data.startswith(b"\x89PNG\r\n\x1a\n") and len(data) >= 24:
        mime_type, extension = "image/png", "png"
        width, height = struct.unpack(">II", data[16:24])
    elif data[:6] in (b"GIF87a", b"GIF89a") and len(data) >= 10:
        mime_type, extension = "image/gif", "gif"
        width, height = struct.unpack("<HH", data[6:10])
    elif data.startswith(b"\xff\xd8"):
        mime_type, extension = "image/jpeg", "jpg"
        offset = 2
        start_of_frame = {
            0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7,
            0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF,
        }
        while offset + 4 <= len(data):
            if data[offset] != 0xFF:
                offset += 1
                continue
            while offset < len(data) and data[offset] == 0xFF:
                offset += 1
            if offset >= len(data):
                break
            marker = data[offset]
            offset += 1
            if marker in (0xD8, 0xD9) or 0xD0 <= marker <= 0xD7:
                continue
            if offset + 2 > len(data):
                break
            segment_length = struct.unpack(">H", data[offset:offset + 2])[0]
            if segment_length < 2 or offset + segment_length > len(data):
                break
            if marker in start_of_frame and segment_length >= 7:
                height, width = struct.unpack(">HH", data[offset + 3:offset + 7])
                break
            offset += segment_length
    elif len(data) >= 30 and data.startswith(b"RIFF") and data[8:12] == b"WEBP":
        mime_type, extension = "image/webp", "webp"
        chunk_type = data[12:16]
        if chunk_type == b"VP8X":
            width = int.from_bytes(data[24:27], "little") + 1
            height = int.from_bytes(data[27:30], "little") + 1
        elif chunk_type == b"VP8 " and len(data) >= 30 and data[23:26] == b"\x9d\x01\x2a":
            width = int.from_bytes(data[26:28], "little") & 0x3FFF
            height = int.from_bytes(data[28:30], "little") & 0x3FFF
        elif chunk_type == b"VP8L" and len(data) >= 25 and data[20] == 0x2F:
            packed = int.from_bytes(data[21:25], "little")
            width = (packed & 0x3FFF) + 1
            height = ((packed >> 14) & 0x3FFF) + 1
    if not mime_type or width <= 0 or height <= 0:
        raise image_error(415, "UNSUPPORTED_IMAGE_FORMAT", "仅支持有效的 PNG、JPEG、GIF 或 WebP 图片")
    if width > IMAGE_MAX_EDGE_PX or height > IMAGE_MAX_EDGE_PX or width * height > IMAGE_MAX_PIXELS:
        raise image_error(413, "IMAGE_TOO_LARGE", "图片像素尺寸超过限制")
    return mime_type, extension, width, height


SVG_FORBIDDEN_ELEMENTS = {
    "script",
    "style",
    "foreignobject",
    "iframe",
    "object",
    "embed",
    "audio",
    "video",
}
SVG_LENGTH_PATTERN = re.compile(
    r"^\s*([0-9]+(?:\.[0-9]+)?|\.[0-9]+)\s*(px|pt|pc|mm|cm|in|q)?\s*$",
    re.IGNORECASE,
)
SVG_URL_PATTERN = re.compile(r"url\(\s*(['\"]?)(.*?)\1\s*\)", re.IGNORECASE)


def svg_length_px(value: Any) -> float | None:
    if value is None:
        return None
    match = SVG_LENGTH_PATTERN.fullmatch(str(value))
    if not match:
        return None
    number = float(match.group(1))
    unit = (match.group(2) or "px").lower()
    scale = {
        "px": 1.0,
        "pt": 96.0 / 72.0,
        "pc": 16.0,
        "mm": 96.0 / 25.4,
        "cm": 96.0 / 2.54,
        "in": 96.0,
        "q": 96.0 / 101.6,
    }[unit]
    return number * scale


def sanitize_svg(data: bytes) -> tuple[bytes, int, int]:
    if (
        re.search(br"<!\s*(?:DOCTYPE|ENTITY)\b", data, re.IGNORECASE)
        or b"\x00" in data
    ):
        raise image_error(415, "UNSUPPORTED_IMAGE_FORMAT", "SVG 包含不允许的文档声明")
    try:
        text = data.decode("utf-8")
        root = ET.fromstring(text)
    except (UnicodeDecodeError, ET.ParseError) as error:
        raise image_error(415, "UNSUPPORTED_IMAGE_FORMAT", "SVG 不是有效的 UTF-8 XML") from error
    root_name = root.tag.rsplit("}", 1)[-1].lower()
    if root_name != "svg":
        raise image_error(415, "UNSUPPORTED_IMAGE_FORMAT", "SVG 根元素无效")
    for element in root.iter():
        element_name = str(element.tag).rsplit("}", 1)[-1].lower()
        if element_name in SVG_FORBIDDEN_ELEMENTS:
            raise image_error(415, "UNSUPPORTED_IMAGE_FORMAT", "SVG 包含不安全的活动内容")
        for attribute, raw_value in element.attrib.items():
            attribute_name = str(attribute).rsplit("}", 1)[-1].lower()
            value = str(raw_value).strip()
            lowered = value.lower()
            if attribute_name.startswith("on"):
                raise image_error(415, "UNSUPPORTED_IMAGE_FORMAT", "SVG 包含不安全的事件处理器")
            if (
                "javascript:" in lowered
                or "vbscript:" in lowered
                or "expression(" in lowered
                or "@import" in lowered
            ):
                raise image_error(415, "UNSUPPORTED_IMAGE_FORMAT", "SVG 包含不安全的脚本或样式")
            if attribute_name in {"href", "src"} and value and not value.startswith("#"):
                raise image_error(415, "UNSUPPORTED_IMAGE_FORMAT", "SVG 不允许引用外部资源")
            for url_match in SVG_URL_PATTERN.finditer(value):
                if not url_match.group(2).strip().startswith("#"):
                    raise image_error(415, "UNSUPPORTED_IMAGE_FORMAT", "SVG 不允许引用外部资源")
    view_box = str(root.attrib.get("viewBox", root.attrib.get("viewbox", ""))).strip()
    view_width = None
    view_height = None
    if view_box:
        try:
            values = [float(value) for value in re.split(r"[\s,]+", view_box) if value]
        except ValueError as error:
            raise image_error(415, "UNSUPPORTED_IMAGE_FORMAT", "SVG viewBox 无效") from error
        if len(values) != 4 or values[2] <= 0 or values[3] <= 0:
            raise image_error(415, "UNSUPPORTED_IMAGE_FORMAT", "SVG viewBox 无效")
        view_width, view_height = values[2], values[3]
    width = svg_length_px(root.attrib.get("width"))
    height = svg_length_px(root.attrib.get("height"))
    if width is None and height is not None and view_width and view_height:
        width = height * view_width / view_height
    if height is None and width is not None and view_width and view_height:
        height = width * view_height / view_width
    width = width if width is not None else (view_width or 300.0)
    height = height if height is not None else (view_height or 150.0)
    width_px = max(1, int(round(width)))
    height_px = max(1, int(round(height)))
    if (
        width_px > IMAGE_MAX_EDGE_PX
        or height_px > IMAGE_MAX_EDGE_PX
        or width_px * height_px > IMAGE_MAX_PIXELS
    ):
        raise image_error(413, "IMAGE_TOO_LARGE", "图片像素尺寸超过限制")
    try:
        ET.register_namespace("", "http://www.w3.org/2000/svg")
        ET.register_namespace("xlink", "http://www.w3.org/1999/xlink")
        sanitized = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    except (TypeError, ValueError) as error:
        raise image_error(415, "UNSUPPORTED_IMAGE_FORMAT", "SVG 无法安全规范化") from error
    if len(sanitized) > IMAGE_MAX_BYTES:
        raise image_error(413, "IMAGE_TOO_LARGE", "图片文件超过 8 MiB 限制")
    return sanitized, width_px, height_px


def decode_image_source(source: Any) -> tuple[bytes, str | None]:
    if not isinstance(source, dict):
        raise image_error(400, "INVALID_IMAGE_SOURCE", "source 必须是图片来源对象")
    source_type = str(source.get("type", ""))
    if source_type == "dataUrl":
        data_url = source.get("dataUrl")
        if not isinstance(data_url, str):
            raise image_error(400, "INVALID_IMAGE_SOURCE", "dataUrl 不能为空")
        match = re.fullmatch(
            r"data:(image/(?:png|jpeg|gif|webp)|image/svg\+xml);base64,([A-Za-z0-9+/]*={0,2})",
            data_url,
        )
        if not match:
            raise image_error(400, "INVALID_IMAGE_SOURCE", "Data URL 必须是受支持图片的严格 Base64 编码")
        try:
            data = base64.b64decode(match.group(2), validate=True)
        except (binascii.Error, ValueError) as error:
            raise image_error(400, "INVALID_IMAGE_SOURCE", "图片 Base64 数据无效") from error
        if not data or len(data) > IMAGE_MAX_BYTES:
            raise image_error(413, "IMAGE_TOO_LARGE", "图片文件超过 8 MiB 限制")
        return data, match.group(1)
    if source_type == "url":
        url = source.get("url")
        if not isinstance(url, str):
            raise image_error(400, "INVALID_IMAGE_SOURCE", "图片 URL 不能为空")
        validated_url = validate_image_url(url)
        request = urllib.request.Request(
            validated_url,
            headers={
                "Accept": "image/png,image/jpeg,image/gif,image/webp,image/svg+xml",
                "User-Agent": "OnlyOffice-ai-bridge/0.3",
            },
            method="GET",
        )
        opener = urllib.request.build_opener(SafeImageRedirectHandler())
        try:
            with opener.open(request, timeout=IMAGE_FETCH_TIMEOUT_SECONDS) as response:
                content_type = str(response.headers.get_content_type()).lower()
                content_length = response.headers.get("Content-Length")
                if content_length and int(content_length) > IMAGE_MAX_BYTES:
                    raise image_error(413, "IMAGE_TOO_LARGE", "图片文件超过 8 MiB 限制")
                data = response.read(IMAGE_MAX_BYTES + 1)
        except BridgeError:
            raise
        except (urllib.error.HTTPError, urllib.error.URLError, OSError, ValueError) as error:
            raise image_error(502, "IMAGE_FETCH_FAILED", "获取外部图片失败") from error
        if len(data) > IMAGE_MAX_BYTES:
            raise image_error(413, "IMAGE_TOO_LARGE", "图片文件超过 8 MiB 限制")
        if content_type not in IMAGE_MIME_EXTENSIONS:
            raise image_error(415, "UNSUPPORTED_IMAGE_FORMAT", "远程响应不是受支持的图片类型")
        return data, content_type
    raise image_error(400, "INVALID_IMAGE_SOURCE", "source.type 必须是 url 或 dataUrl")


def cleanup_image_assets(now: float | None = None) -> None:
    root = image_asset_root()
    if not os.path.isdir(root):
        return
    current = time.time() if now is None else now
    for name in os.listdir(root):
        if not IMAGE_ASSET_PATTERN.fullmatch(name):
            continue
        path = os.path.join(root, name)
        try:
            if current - os.path.getmtime(path) > IMAGE_RETENTION_SECONDS:
                os.remove(path)
        except FileNotFoundError:
            continue


def import_image(source: Any, claims: dict[str, Any], now: float | None = None) -> dict[str, Any]:
    data, declared_mime = decode_image_source(source)
    if declared_mime == "image/svg+xml":
        data, width, height = sanitize_svg(data)
        mime_type, extension = "image/svg+xml", "svg"
    else:
        mime_type, extension, width, height = sniff_image(data)
    if declared_mime != mime_type:
        raise image_error(415, "UNSUPPORTED_IMAGE_FORMAT", "图片声明类型与实际内容不一致")
    current = time.time() if now is None else now
    root = image_asset_root()
    os.makedirs(root, mode=0o700, exist_ok=True)
    cleanup_image_assets(current)
    digest = hashlib.sha256(data).hexdigest()
    asset_id = f"{digest}.{extension}"
    asset_path = os.path.join(root, asset_id)
    try:
        with open(asset_path, "xb") as stream:
            stream.write(data)
        os.chmod(asset_path, 0o600)
    except FileExistsError:
        os.utime(asset_path, (current, current))
    expires_at = int(current) + IMAGE_URL_TTL_SECONDS
    identity_digest = bridge_identity_digest(bridge_identity(claims))
    token = sign_jwt(
        {
            "scope": "ai-bridge-image",
            "assetId": asset_id,
            "identity": identity_digest,
            "exp": expires_at,
        },
        get_jwt_secret(),
    )
    return {
        "ok": True,
        "asset": {
            "assetId": asset_id,
            "path": f"/copilot-api/images/{asset_id}?token={token}",
            "mimeType": mime_type,
            "widthPx": width,
            "heightPx": height,
            "expiresAt": expires_at,
        },
    }


def verify_image_asset(asset_id: str, token: str, now: float | None = None) -> tuple[str, str]:
    if not IMAGE_ASSET_PATTERN.fullmatch(asset_id):
        raise image_error(404, "IMAGE_ASSET_EXPIRED", "图片资源不存在或已过期")
    try:
        encoded_header, encoded_payload, encoded_signature = token.split(".")
        unsigned = f"{encoded_header}.{encoded_payload}"
        expected = b64url(hmac.new(get_jwt_secret().encode(), unsigned.encode(), hashlib.sha256).digest())
        if not hmac.compare_digest(encoded_signature, expected):
            raise ValueError("signature")
        header = json.loads(base64.urlsafe_b64decode(encoded_header + "=" * (-len(encoded_header) % 4)))
        payload = json.loads(base64.urlsafe_b64decode(encoded_payload + "=" * (-len(encoded_payload) % 4)))
        if (
            header.get("alg") != "HS256"
            or not isinstance(payload, dict)
            or payload.get("scope") != "ai-bridge-image"
            or not hmac.compare_digest(str(payload.get("assetId", "")), asset_id)
        ):
            raise ValueError("claims")
        current = time.time() if now is None else now
        if float(payload.get("exp", 0)) < current:
            raise image_error(410, "IMAGE_ASSET_EXPIRED", "图片资源不存在或已过期")
    except BridgeError:
        raise
    except Exception as error:
        raise image_error(404, "IMAGE_ASSET_EXPIRED", "图片资源不存在或已过期") from error
    path = os.path.join(image_asset_root(), asset_id)
    if not os.path.isfile(path):
        raise image_error(404, "IMAGE_ASSET_EXPIRED", "图片资源不存在或已过期")
    extension = asset_id.rsplit(".", 1)[1]
    mime_type = next(
        mime for mime, configured_extension in IMAGE_MIME_EXTENSIONS.items()
        if configured_extension == extension
    )
    return path, mime_type


def bridge_binding_token(claims: dict[str, Any]) -> tuple[str, int]:
    file_name, file_type, editor_type, user_id = bridge_identity(claims)
    expires_at = int(time.time()) + BRIDGE_BINDING_TOKEN_TTL_SECONDS
    return (
        sign_jwt(
            {
                "scope": "ai-bridge-binding",
                "fileName": file_name,
                "fileType": file_type,
                "editorType": editor_type,
                "userId": user_id,
                "exp": expires_at,
            },
            get_jwt_secret(),
        ),
        expires_at,
    )


def verify_bridge_binding_token(token: str) -> dict[str, Any]:
    try:
        encoded_header, encoded_payload, encoded_signature = token.split(".")
        unsigned = f"{encoded_header}.{encoded_payload}"
        expected = b64url(hmac.new(get_jwt_secret().encode(), unsigned.encode(), hashlib.sha256).digest())
        if not hmac.compare_digest(encoded_signature, expected):
            raise ValueError("signature")
        header = json.loads(base64.urlsafe_b64decode(encoded_header + "=" * (-len(encoded_header) % 4)))
        payload = json.loads(base64.urlsafe_b64decode(encoded_payload + "=" * (-len(encoded_payload) % 4)))
        if header.get("alg") != "HS256" or not isinstance(payload, dict):
            raise ValueError("claims")
        if payload.get("scope") != "ai-bridge-binding":
            raise ValueError("scope")
        if float(payload.get("exp", 0)) < time.time():
            raise BridgeError(
                401,
                "BRIDGE_BINDING_TOKEN_EXPIRED",
                "ai-bridge 绑定凭证已过期，请重新绑定当前编辑器一次",
            )
        claims = {
            "fileName": str(payload.get("fileName", "")),
            "fileType": str(payload.get("fileType", "")),
            "editorType": str(payload.get("editorType", "")),
            "userId": str(payload.get("userId", "")),
            "authKind": "binding",
        }
        bridge_identity(claims)
        return claims
    except BridgeError:
        raise
    except Exception as error:
        raise BridgeError(
            401,
            "INVALID_BRIDGE_BINDING_TOKEN",
            "无效的 ai-bridge 绑定凭证，请重新绑定当前编辑器一次",
        ) from error


def bridge_token_scope(token: str) -> str:
    try:
        encoded_payload = token.split(".")[1]
        payload = json.loads(base64.urlsafe_b64decode(encoded_payload + "=" * (-len(encoded_payload) % 4)))
        return str(payload.get("scope", "")) if isinstance(payload, dict) else ""
    except Exception:
        return ""


def bridge_resume_token(session_id: str, claims: dict[str, Any]) -> str:
    return sign_jwt(
        {
            "scope": "ai-bridge-resume",
            "sessionId": session_id,
            "documentKey": claims["documentKey"],
            "fileName": claims.get("fileName", ""),
            "fileType": claims.get("fileType", ""),
            "editorType": claims["editorType"],
            "userId": claims.get("userId", ""),
            "exp": int(time.time()) + BRIDGE_RESUME_TOKEN_TTL_SECONDS,
        },
        get_jwt_secret(),
    )


def verify_bridge_resume_token(token: str, session_id: str) -> dict[str, Any]:
    try:
        encoded_header, encoded_payload, encoded_signature = token.split(".")
        unsigned = f"{encoded_header}.{encoded_payload}"
        expected = b64url(hmac.new(get_jwt_secret().encode(), unsigned.encode(), hashlib.sha256).digest())
        if not hmac.compare_digest(encoded_signature, expected):
            raise ValueError("signature")
        header = json.loads(base64.urlsafe_b64decode(encoded_header + "=" * (-len(encoded_header) % 4)))
        payload = json.loads(base64.urlsafe_b64decode(encoded_payload + "=" * (-len(encoded_payload) % 4)))
        if header.get("alg") != "HS256" or not isinstance(payload, dict):
            raise ValueError("claims")
        if payload.get("scope") != "ai-bridge-resume":
            raise ValueError("scope")
        if not hmac.compare_digest(str(payload.get("sessionId", "")), session_id):
            raise ValueError("session")
        if float(payload.get("exp", 0)) < time.time():
            raise BridgeError(
                401,
                "BRIDGE_RESUME_TOKEN_EXPIRED",
                "浏览器 Relay 续接凭证已过期，编辑器页面将自动刷新",
            )
        document_key = str(payload.get("documentKey", "")).strip()
        editor_type = str(payload.get("editorType", "")).strip()
        if not document_key or not editor_type:
            raise ValueError("document")
        return {
            "documentKey": document_key,
            "fileName": str(payload.get("fileName", "")),
            "fileType": str(payload.get("fileType", "")),
            "editorType": editor_type,
            "userId": str(payload.get("userId", "")),
            "authKind": "resume",
        }
    except BridgeError:
        raise
    except Exception as error:
        raise BridgeError(
            401,
            "INVALID_BRIDGE_RESUME_TOKEN",
            "浏览器 Relay 续接凭证无效，请重新加载编辑器页面",
        ) from error


def bridge_authorization(handler: BaseHTTPRequestHandler) -> dict[str, Any]:
    value = handler.headers.get("Authorization", "")
    if not value.startswith("Bearer "):
        raise BridgeError(401, "EDITOR_TOKEN_REQUIRED", "请先绑定当前编辑器")
    token = value.removeprefix("Bearer ").strip()
    if bridge_token_scope(token) == "ai-bridge-binding":
        return verify_bridge_binding_token(token)
    return verify_editor_jwt(token)


def bridge_public_session(session_id: str, session: dict[str, Any]) -> dict[str, Any]:
    state = session.get("state") if isinstance(session.get("state"), dict) else {}
    context = state.get("context") if isinstance(state.get("context"), dict) else {}
    return {
        "sessionId": session_id,
        "ready": bool(state.get("ready")),
        "editorType": state.get("editorType") or context.get("editorType"),
        "documentKey": context.get("documentKey"),
        "fileName": context.get("fileName"),
        "fileType": context.get("fileType"),
        "userId": context.get("userId"),
        "capabilities": state.get("capabilities"),
        "lastSeen": session.get("lastSeen"),
        "registeredAt": session.get("registeredAt"),
        "generation": session.get("generation"),
        "authoritative": bool(session.get("authoritative")),
    }


def bridge_cleanup_locked(now: float | None = None) -> None:
    current = now if now is not None else time.time()
    stale_sessions = [
        session_id
        for session_id, session in BRIDGE_SESSIONS.items()
        if current - float(session.get("lastSeen", 0)) > BRIDGE_SESSION_TTL_SECONDS
    ]
    for session_id in stale_sessions:
        session = BRIDGE_SESSIONS.pop(session_id)
        identity = session.get("identity")
        if identity and BRIDGE_AUTHORITATIVE.get(identity) == session_id:
            del BRIDGE_AUTHORITATIVE[identity]

    stale_superseded = [
        session_id
        for session_id, session in BRIDGE_SUPERSEDED.items()
        if float(session.get("expiresAt", 0)) < current
    ]
    for session_id in stale_superseded:
        del BRIDGE_SUPERSEDED[session_id]

    stale_credentials = [
        fingerprint
        for fingerprint, rejection in BRIDGE_REJECTED_CREDENTIALS.items()
        if float(rejection.get("expiresAt", 0)) < current
    ]
    for fingerprint in stale_credentials:
        del BRIDGE_REJECTED_CREDENTIALS[fingerprint]

    stale_commands = [
        command_id
        for command_id, command in BRIDGE_COMMANDS.items()
        if current - float(command.get("completedAt") or command.get("createdAt") or current)
        > (
            BRIDGE_RESULT_TTL_SECONDS
            if command.get("completedAt") is not None
            else BRIDGE_PENDING_TTL_SECONDS
        )
    ]
    for command_id in stale_commands:
        command = BRIDGE_COMMANDS.pop(command_id)
        request_key = (command.get("identity"), str(command.get("requestId", "")))
        if BRIDGE_REQUESTS.get(request_key) == command_id:
            del BRIDGE_REQUESTS[request_key]


def bridge_registration_fingerprint(payload: dict[str, Any]) -> str:
    resume_token = str(payload.get("resumeToken", "")).strip()
    editor_token = str(payload.get("editorToken", "")).strip()
    credential_kind = "resume" if resume_token else "editor"
    credential = resume_token or editor_token
    if not credential:
        credential = str(payload.get("sessionId", "")).strip()
    return hashlib.sha256(f"{credential_kind}\0{credential}".encode()).hexdigest()


def bridge_registration_error(payload: dict[str, Any], error: BridgeError) -> BridgeError:
    if error.code not in BRIDGE_REGISTRATION_CREDENTIAL_ERRORS:
        return error

    fingerprint = bridge_registration_fingerprint(payload)
    now = time.time()
    with BRIDGE_CONDITION:
        bridge_cleanup_locked(now)
        rejection = BRIDGE_REJECTED_CREDENTIALS.get(fingerprint)
        if rejection is None:
            BRIDGE_REJECTED_CREDENTIALS[fingerprint] = {
                "expiresAt": now + BRIDGE_REJECTED_CREDENTIAL_TTL_SECONDS,
                "terminalLogged": False,
            }
            return error

        rejection["expiresAt"] = now + BRIDGE_REJECTED_CREDENTIAL_TTL_SECONDS
        terminal = BridgeError(
            409,
            "SESSION_SUPERSEDED",
            "该浏览器 Relay 的认证刷新已进入终态，请重新打开编辑器页面",
            {"reason": "credential-refresh-limit"},
        )
        terminal.suppress_log = bool(rejection.get("terminalLogged"))
        rejection["terminalLogged"] = True
        return terminal


def bridge_registration_succeeded(payload: dict[str, Any]) -> None:
    fingerprint = bridge_registration_fingerprint(payload)
    with BRIDGE_CONDITION:
        BRIDGE_REJECTED_CREDENTIALS.pop(fingerprint, None)


def bridge_assert_state_matches(claims: dict[str, Any], state: Any) -> dict[str, Any]:
    if not isinstance(state, dict):
        raise BridgeError(400, "INVALID_STATE", "state 必须是对象")
    context = state.get("context") if isinstance(state.get("context"), dict) else {}
    if context.get("documentKey") != claims["documentKey"]:
        raise BridgeError(409, "DOCUMENT_MISMATCH", "编辑器页面与凭证绑定的 document.key 不一致")
    if (state.get("editorType") or context.get("editorType")) != claims["editorType"]:
        raise BridgeError(409, "EDITOR_MISMATCH", "编辑器页面与凭证的编辑器类型不一致")
    for key in ("fileName", "fileType", "userId"):
        if str(context.get(key, "")).strip() != str(claims.get(key, "")).strip():
            raise BridgeError(
                409,
                "DOCUMENT_IDENTITY_MISMATCH",
                "编辑器页面与凭证的稳定文档身份不一致",
            )
    return state


def bridge_session_matches_claims(session: dict[str, Any], claims: dict[str, Any]) -> bool:
    if claims.get("authKind") == "binding":
        return session.get("identity") == bridge_identity(claims)
    return session.get("documentKey") == claims.get("documentKey")


def bridge_supersede_locked(
    session_id: str,
    replacement_id: str,
    now: float,
) -> list[str]:
    session = BRIDGE_SESSIONS.pop(session_id, None)
    if session is None:
        return []
    session["authoritative"] = False
    identity = session.get("identity")
    if identity and BRIDGE_AUTHORITATIVE.get(identity) == session_id:
        del BRIDGE_AUTHORITATIVE[identity]
    BRIDGE_SUPERSEDED[session_id] = {
        **session,
        "replacedBy": replacement_id,
        "expiresAt": now + BRIDGE_RESUME_TOKEN_TTL_SECONDS,
    }

    transferred: list[str] = []
    for command_id in list(session.get("queue") or []):
        command = BRIDGE_COMMANDS.get(command_id)
        if (
            command is not None
            and command.get("response") is None
            and command.get("deliveredAt") is None
        ):
            command["sessionId"] = replacement_id
            transferred.append(command_id)
    session["queue"] = []
    return transferred


def bridge_register(payload: dict[str, Any]) -> dict[str, Any]:
    global BRIDGE_GENERATION

    session_id = str(payload.get("sessionId", "")).strip()
    if not BRIDGE_ID_PATTERN.fullmatch(session_id):
        raise BridgeError(400, "INVALID_SESSION_ID", "sessionId 格式无效")
    resume_token = str(payload.get("resumeToken", "")).strip()
    try:
        claims = (
            verify_bridge_resume_token(resume_token, session_id)
            if resume_token
            else verify_editor_jwt(str(payload.get("editorToken", "")))
        )
        state = bridge_assert_state_matches(claims, payload.get("state"))
        identity = bridge_identity(claims)
    except BridgeError as error:
        registration_error = bridge_registration_error(payload, error)
        if registration_error is error:
            raise
        raise registration_error from error
    bridge_registration_succeeded(payload)
    now = time.time()
    relay_key = secrets.token_urlsafe(32)
    with BRIDGE_CONDITION:
        bridge_cleanup_locked(now)
        superseded = BRIDGE_SUPERSEDED.get(session_id)
        if superseded is not None:
            raise BridgeError(
                409,
                "SESSION_SUPERSEDED",
                "当前浏览器 Relay 已被更新的编辑器页面接管",
                {"replacementSessionId": superseded.get("replacedBy")},
            )

        existing = BRIDGE_SESSIONS.get(session_id)
        if existing is not None:
            if existing.get("identity") != identity:
                raise BridgeError(409, "SESSION_ID_CONFLICT", "sessionId 已属于另一个编辑器页面")
            existing.update(
                {
                    "documentKey": claims["documentKey"],
                    "editorType": claims["editorType"],
                    "relayKey": relay_key,
                    "state": state,
                    "lastSeen": now,
                }
            )
            session = existing
        else:
            BRIDGE_GENERATION += 1
            transferred: list[str] = []
            previous_id = BRIDGE_AUTHORITATIVE.get(identity)
            if previous_id and previous_id != session_id:
                transferred = bridge_supersede_locked(previous_id, session_id, now)
            session = {
                "documentKey": claims["documentKey"],
                "editorType": claims["editorType"],
                "identity": identity,
                "relayKey": relay_key,
                "state": state,
                "lastSeen": now,
                "registeredAt": now,
                "generation": BRIDGE_GENERATION,
                "authoritative": True,
                "queue": transferred,
            }
            BRIDGE_SESSIONS[session_id] = session

        session["authoritative"] = True
        BRIDGE_AUTHORITATIVE[identity] = session_id
        print(
            "[bridge-session] "
            + compact_json(
                {
                    "event": "reconnected" if existing is not None else "registered",
                    "identity": bridge_identity_digest(identity),
                    "sessionId": session_id,
                    "generation": session.get("generation"),
                }
            ),
            flush=True,
        )
        startup_timings = bridge_startup_timings(state)
        if startup_timings:
            print(
                "[bridge-startup] "
                + compact_json(
                    {
                        "event": "reconnected" if existing is not None else "registered",
                        "identity": bridge_identity_digest(identity),
                        "sessionId": session_id,
                        **startup_timings,
                    }
                ),
                flush=True,
            )
        BRIDGE_CONDITION.notify_all()
        return {
            "ok": True,
            "relayKey": relay_key,
            "resumeToken": bridge_resume_token(session_id, claims),
            "session": bridge_public_session(session_id, session),
        }


def bridge_authenticate_page_locked(
    payload: dict[str, Any],
    allow_superseded: bool = False,
) -> tuple[str, dict[str, Any]]:
    session_id = str(payload.get("sessionId", "")).strip()
    relay_key = str(payload.get("relayKey", ""))
    session = BRIDGE_SESSIONS.get(session_id)
    if session is None:
        superseded = BRIDGE_SUPERSEDED.get(session_id)
        if (
            superseded is not None
            and hmac.compare_digest(relay_key, str(superseded.get("relayKey", "")))
        ):
            if allow_superseded:
                return session_id, superseded
            raise BridgeError(
                409,
                "SESSION_SUPERSEDED",
                "当前浏览器 Relay 已被更新的编辑器页面接管",
                {"replacementSessionId": superseded.get("replacedBy")},
            )
        raise BridgeError(401, "INVALID_RELAY_SESSION", "浏览器 Relay 会话无效，请重新加载编辑器页面")
    if not hmac.compare_digest(relay_key, str(session.get("relayKey", ""))):
        raise BridgeError(401, "INVALID_RELAY_SESSION", "浏览器 Relay 会话无效，请重新加载编辑器页面")
    state = payload.get("state")
    if state is not None:
        claims = {
            "documentKey": session["documentKey"],
            "fileName": session["identity"][0],
            "fileType": session["identity"][1],
            "editorType": session["identity"][2],
            "userId": session["identity"][3],
        }
        session["state"] = bridge_assert_state_matches(claims, state)
    session["lastSeen"] = time.time()
    return session_id, session


def bridge_sessions(claims: dict[str, Any]) -> dict[str, Any]:
    deadline = time.time() + BRIDGE_DISCOVERY_WAIT_SECONDS
    with BRIDGE_CONDITION:
        while True:
            bridge_cleanup_locked()
            sessions = [
                bridge_public_session(session_id, session)
                for session_id, session in BRIDGE_SESSIONS.items()
                if session.get("authoritative")
                and bridge_session_matches_claims(session, claims)
            ]
            if sessions:
                break
            remaining = deadline - time.time()
            if remaining <= 0:
                break
            BRIDGE_CONDITION.wait(timeout=remaining)
    sessions.sort(key=lambda item: int(item.get("generation") or 0), reverse=True)
    binding_claims = claims
    if claims.get("authKind") == "editor" and sessions:
        selected = next(
            (session for session in sessions if session.get("ready")),
            sessions[0],
        )
        binding_claims = {
            "fileName": selected.get("fileName"),
            "fileType": selected.get("fileType"),
            "editorType": selected.get("editorType"),
            "userId": selected.get("userId"),
            "authKind": "binding",
        }
    binding_token, expires_at = bridge_binding_token(binding_claims)
    return {
        "ok": True,
        "bindingToken": binding_token,
        "bindingExpiresAt": expires_at,
        "sessions": sessions,
    }


def bridge_select_session_locked(
    requested_session_id: Any,
    claims: dict[str, Any],
) -> tuple[str, dict[str, Any]]:
    deadline = time.time() + BRIDGE_DISCOVERY_WAIT_SECONDS
    requested = str(requested_session_id or "").strip()
    while True:
        bridge_cleanup_locked()
        candidates = [
            (candidate_id, session)
            for candidate_id, session in BRIDGE_SESSIONS.items()
            if session.get("authoritative")
            and bridge_session_matches_claims(session, claims)
            and bool((session.get("state") or {}).get("ready"))
        ]
        if requested:
            for candidate_id, session in candidates:
                if candidate_id == requested:
                    return candidate_id, session
            if candidates:
                raise BridgeError(
                    409,
                    "SESSION_NOT_AUTHORITATIVE",
                    "指定会话不是当前文档的权威编辑器页面",
                    {"authoritativeSessionId": candidates[0][0]},
                )
        elif candidates:
            candidates.sort(key=lambda item: int(item[1].get("generation") or 0), reverse=True)
            return candidates[0]

        remaining = deadline - time.time()
        if remaining <= 0:
            break
        BRIDGE_CONDITION.wait(timeout=remaining)

    if requested:
        raise BridgeError(404, "SESSION_NOT_FOUND", "指定的当前文档编辑器会话不存在或已离线")
    raise BridgeError(503, "NO_ACTIVE_EDITOR", "当前文档没有已就绪的 ai-bridge 编辑器页面")


def bridge_parse_json_parameter(
    payload: dict[str, Any],
    direct_key: str,
    json_key: str,
    expected_type: type,
    default: Any,
) -> Any:
    value = payload.get(direct_key)
    if value is None and payload.get(json_key) not in (None, ""):
        try:
            value = json.loads(str(payload[json_key]))
        except json.JSONDecodeError as error:
            raise BridgeError(400, "INVALID_ARGUMENTS", f"{json_key} 不是有效 JSON：{error.msg}") from error
    if value is None:
        value = default
    if not isinstance(value, expected_type):
        label = "对象" if expected_type is dict else "数组"
        raise BridgeError(400, "INVALID_ARGUMENTS", f"{direct_key} 必须是 JSON {label}")
    return value


def bridge_build_command(payload: dict[str, Any], session: dict[str, Any]) -> dict[str, Any]:
    method = str(payload.get("method", "")).strip()
    if method not in BRIDGE_ALLOWED_METHODS:
        raise BridgeError(400, "METHOD_NOT_ALLOWED", f"不允许 HTTP Relay 方法：{method or 'unknown'}")
    request_id = str(payload.get("requestId", "")).strip()
    if not BRIDGE_ID_PATTERN.fullmatch(request_id):
        raise BridgeError(400, "INVALID_REQUEST_ID", "requestId 格式无效")
    try:
        timeout_ms = int(payload.get("timeoutMs", 90000))
    except (TypeError, ValueError) as error:
        raise BridgeError(400, "INVALID_TIMEOUT", "timeoutMs 必须是整数") from error
    timeout_ms = min(300000, max(100, timeout_ms))

    params: dict[str, Any] = {}
    allowed_tools = set(((session.get("state") or {}).get("capabilities") or {}).get("tools") or [])
    if method == "executeTool":
        name = str(payload.get("name", "")).strip()
        if name not in allowed_tools:
            raise BridgeError(400, "TOOL_NOT_ALLOWED", f"当前编辑器不允许工具：{name or 'unknown'}")
        arguments = bridge_parse_json_parameter(payload, "arguments", "argumentsJson", dict, {})
        argument_limit = IMAGE_MAX_REQUEST_BYTES if name in IMAGE_SOURCE_TOOLS else 250000
        if len(compact_json(arguments)) > argument_limit:
            raise BridgeError(400, "ARGUMENTS_TOO_LARGE", f"arguments 超过 {argument_limit} 字符")
        params = {"name": name, "arguments": arguments}
    elif method == "executeBatch":
        tool_calls = bridge_parse_json_parameter(payload, "toolCalls", "toolCallsJson", list, [])
        if not tool_calls or len(tool_calls) > 20:
            raise BridgeError(400, "INVALID_TOOL_CALL", "toolCalls 数量必须为 1 到 20")
        for call in tool_calls:
            if not isinstance(call, dict) or call.get("name") not in allowed_tools:
                raise BridgeError(400, "TOOL_NOT_ALLOWED", "批量调用包含当前编辑器不允许的工具")
        arguments_limit = (
            IMAGE_MAX_REQUEST_BYTES
            if any(call.get("name") in IMAGE_SOURCE_TOOLS for call in tool_calls)
            else 250000
        )
        if len(compact_json(tool_calls)) > arguments_limit:
            raise BridgeError(400, "ARGUMENTS_TOO_LARGE", f"toolCalls 超过 {arguments_limit} 字符")
        params = {"toolCalls": tool_calls}

    return {
        "requestId": request_id,
        "method": method,
        "params": params,
        "timeoutMs": timeout_ms,
    }


def bridge_execute(payload: dict[str, Any], claims: dict[str, Any]) -> dict[str, Any]:
    with BRIDGE_CONDITION:
        session_id, session = bridge_select_session_locked(payload.get("sessionId"), claims)
        command_data = bridge_build_command(payload, session)
        identity = session["identity"]
        request_key = (identity, command_data["requestId"])
        command_id = BRIDGE_REQUESTS.get(request_key)
        cached = command_id is not None
        if command_id is None:
            command_id = f"command:{uuid.uuid4()}"
            created_at = time.time()
            created_monotonic = time.monotonic()
            command = {
                "commandId": command_id,
                "sessionId": session_id,
                "createdAt": created_at,
                "createdMonotonic": created_monotonic,
                "deliveredAt": None,
                "firstDeliveredMonotonic": None,
                "completedAt": None,
                "completedMonotonic": None,
                "response": None,
                "timingLogged": False,
                "identity": identity,
                **command_data,
            }
            BRIDGE_COMMANDS[command_id] = command
            BRIDGE_REQUESTS[request_key] = command_id
            session["queue"].append(command_id)
            BRIDGE_CONDITION.notify_all()
        else:
            command = BRIDGE_COMMANDS[command_id]
            if (
                command.get("method") != command_data["method"]
                or command.get("params") != command_data["params"]
            ):
                raise BridgeError(
                    409,
                    "REQUEST_ID_CONFLICT",
                    "同一 requestId 不能用于不同方法或参数",
                    {"requestId": command_data["requestId"], "sessionId": session_id},
                )
            if (
                command.get("response") is None
                and command.get("deliveredAt") is None
                and command.get("sessionId") != session_id
            ):
                command["sessionId"] = session_id
                if command_id not in session["queue"]:
                    session["queue"].append(command_id)
                    BRIDGE_CONDITION.notify_all()
            redelivery_after = float(command.get("timeoutMs", 90000)) / 1000 + 5
            delivered_at = command.get("deliveredAt")
            if (
                command.get("response") is None
                and delivered_at is not None
                and time.time() - float(delivered_at) > redelivery_after
                and command_id not in session["queue"]
            ):
                session["queue"].append(command_id)
                BRIDGE_CONDITION.notify_all()

        deadline = time.time() + float(command.get("timeoutMs", 90000)) / 1000 + 10
        while command.get("response") is None:
            remaining = deadline - time.time()
            if remaining <= 0:
                log_bridge_command(command, "timeout", time.monotonic())
                raise BridgeError(
                    504,
                    "BRIDGE_TIMEOUT",
                    "等待当前编辑器执行命令超时；重试时必须复用同一 requestId",
                    {"requestId": command_data["requestId"], "sessionId": session_id},
                )
            BRIDGE_CONDITION.wait(timeout=remaining)

        response = command["response"]
        if not response.get("ok"):
            error = response.get("error") if isinstance(response.get("error"), dict) else {}
            code = str(error.get("code") or "EXECUTION_FAILED")
            raise BridgeError(
                editor_error_http_status(code),
                code,
                str(error.get("message") or "当前编辑器执行失败"),
                error.get("details"),
            )
        return {
            "ok": True,
            "cached": cached,
            "requestId": command_data["requestId"],
            "session": bridge_public_session(session_id, session),
            "result": response.get("result"),
        }


def bridge_poll(payload: dict[str, Any]) -> dict[str, Any]:
    try:
        timeout_ms = int(payload.get("timeoutMs", 25000))
    except (TypeError, ValueError):
        timeout_ms = 25000
    timeout_seconds = min(BRIDGE_MAX_POLL_SECONDS, max(0.1, timeout_ms / 1000))
    deadline = time.time() + timeout_seconds
    with BRIDGE_CONDITION:
        session_id, session = bridge_authenticate_page_locked(payload)
        while True:
            if not session.get("authoritative"):
                superseded = BRIDGE_SUPERSEDED.get(session_id, session)
                raise BridgeError(
                    409,
                    "SESSION_SUPERSEDED",
                    "当前浏览器 Relay 已被更新的编辑器页面接管",
                    {"replacementSessionId": superseded.get("replacedBy")},
                )
            while session["queue"]:
                command_id = session["queue"].pop(0)
                command = BRIDGE_COMMANDS.get(command_id)
                if command is None or command.get("response") is not None:
                    continue
                command["deliveredAt"] = time.time()
                delivered_monotonic = time.monotonic()
                if command.get("firstDeliveredMonotonic") is None:
                    command["firstDeliveredMonotonic"] = delivered_monotonic
                return {
                    "ok": True,
                    "command": {
                        "commandId": command["commandId"],
                        "requestId": command["requestId"],
                        "method": command["method"],
                        "params": command["params"],
                        "timeoutMs": command["timeoutMs"],
                    },
                }
            remaining = deadline - time.time()
            if remaining <= 0:
                session["lastSeen"] = time.time()
                return {"ok": True, "command": None}
            BRIDGE_CONDITION.wait(timeout=remaining)


def bridge_result(payload: dict[str, Any]) -> dict[str, Any]:
    command_id = str(payload.get("commandId", "")).strip()
    with BRIDGE_CONDITION:
        session_id, _ = bridge_authenticate_page_locked(payload, allow_superseded=True)
        command = BRIDGE_COMMANDS.get(command_id)
        if command is None or command.get("sessionId") != session_id:
            raise BridgeError(404, "COMMAND_NOT_FOUND", "命令不存在或不属于该编辑器会话")
        if command.get("response") is None:
            completed_at = time.time()
            completed_monotonic = time.monotonic()
            command["response"] = {
                "ok": bool(payload.get("ok")),
                "result": payload.get("result"),
                "error": payload.get("error"),
            }
            command["completedAt"] = completed_at
            command["completedMonotonic"] = completed_monotonic
            log_bridge_command(
                command,
                "completed" if command["response"]["ok"] else "failed",
                completed_monotonic,
            )
            BRIDGE_CONDITION.notify_all()
        return {"ok": True, "accepted": True}


def bridge_unregister(payload: dict[str, Any]) -> dict[str, Any]:
    with BRIDGE_CONDITION:
        session_id, session = bridge_authenticate_page_locked(payload, allow_superseded=True)
        removed = BRIDGE_SESSIONS.pop(session_id, None) is not None
        if removed:
            identity = session.get("identity")
            if identity and BRIDGE_AUTHORITATIVE.get(identity) == session_id:
                del BRIDGE_AUTHORITATIVE[identity]
        else:
            removed = BRIDGE_SUPERSEDED.pop(session_id, None) is not None
        BRIDGE_CONDITION.notify_all()
    return {"ok": True, "removed": removed}


def json_response(
    handler: BaseHTTPRequestHandler,
    status: int,
    payload: dict[str, Any],
    headers: dict[str, str] | None = None,
) -> None:
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    try:
        handler.send_response(status)
        handler.send_header("Content-Type", "application/json; charset=utf-8")
        handler.send_header("Content-Length", str(len(body)))
        handler.send_header("Cache-Control", "no-store")
        handler.send_header("X-Content-Type-Options", "nosniff")
        for name, value in (headers or {}).items():
            handler.send_header(name, value)
        handler.end_headers()
        handler.wfile.write(body)
    except (BrokenPipeError, ConnectionResetError):
        return


def read_json(handler: BaseHTTPRequestHandler, max_bytes: int = 2_000_000) -> dict[str, Any]:
    length = int(handler.headers.get("Content-Length", "0"))
    if length <= 0 or length > max_bytes:
        raise ValueError("请求体为空或过大")
    payload = json.loads(handler.rfile.read(length).decode("utf-8"))
    if not isinstance(payload, dict):
        raise ValueError("请求体必须是 JSON 对象")
    return payload


def document_storage_directory() -> str:
    return os.environ.get(
        "DOCUMENT_STORAGE_DIR",
        os.path.join(EXAMPLE_FILES, EXAMPLE_STORAGE_ID),
    )


def document_template_path(file_type: str) -> str:
    if file_type not in DOCUMENT_TYPES:
        raise BridgeError(404, "DOCUMENT_TYPE_NOT_FOUND", "不支持的文档类型")
    template_root = os.environ.get("DOCUMENT_TEMPLATE_ROOT", DOCUMENT_TEMPLATE_ROOT)
    return os.path.join(template_root, f"new.{file_type}")


def normalize_document_id(value: Any) -> str:
    document_id = str(value or "").strip()
    if not DOCUMENT_UUID_PATTERN.fullmatch(document_id):
        raise BridgeError(404, "DOCUMENT_NOT_FOUND", "文档不存在")
    return document_id


def document_path(file_type: str, document_id: Any, must_exist: bool = True) -> str:
    document_id = normalize_document_id(document_id)
    if file_type not in DOCUMENT_TYPES:
        raise BridgeError(404, "DOCUMENT_NOT_FOUND", "文档不存在")
    path = os.path.join(document_storage_directory(), f"{document_id}.{file_type}")
    if must_exist and not os.path.isfile(path):
        raise BridgeError(404, "DOCUMENT_NOT_FOUND", "文档不存在")
    return path


def document_file_identity(file_name: Any) -> tuple[str, str]:
    match = DOCUMENT_FILE_PATTERN.fullmatch(str(file_name or "").strip())
    if not match:
        raise BridgeError(404, "DOCUMENT_NOT_FOUND", "文档不存在")
    file_type = match.group("fileType")
    document_id = match.group("documentId")
    document_path(file_type, document_id)
    return file_type, document_id


def document_public_origin() -> str:
    value = os.environ.get(
        "DOCUMENT_PUBLIC_ORIGIN",
        "http://localhost:8088",
    ).strip().rstrip("/")
    parsed = urllib.parse.urlsplit(value)
    if (
        parsed.scheme not in ("http", "https")
        or not parsed.netloc
        or parsed.path not in ("", "/")
        or parsed.query
        or parsed.fragment
    ):
        raise RuntimeError("DOCUMENT_PUBLIC_ORIGIN 必须是无路径的 HTTP(S) origin")
    return value


def document_editor_url(file_type: str, document_id: str) -> str:
    return f"{document_public_origin()}/{file_type}/{document_id}"


def create_document(file_type: str) -> dict[str, Any]:
    template = document_template_path(file_type)
    if not os.path.isfile(template):
        raise RuntimeError(f"找不到 {file_type.upper()} 空白模板")
    directory = document_storage_directory()
    os.makedirs(directory, mode=0o755, exist_ok=True)

    for _ in range(10):
        document_id = str(uuid.uuid4())
        destination = document_path(file_type, document_id, must_exist=False)
        try:
            descriptor = os.open(
                destination,
                os.O_WRONLY | os.O_CREAT | os.O_EXCL,
                0o644,
            )
        except FileExistsError:
            continue
        try:
            with open(template, "rb") as source, os.fdopen(descriptor, "wb") as output:
                descriptor = -1
                shutil.copyfileobj(source, output)
                output.flush()
                os.fsync(output.fileno())
        except Exception:
            if descriptor >= 0:
                os.close(descriptor)
            try:
                os.unlink(destination)
            except FileNotFoundError:
                pass
            raise
        return {
            "documentId": document_id,
            "fileName": f"{document_id}.{file_type}",
            "editorUrl": document_editor_url(file_type, document_id),
        }
    raise RuntimeError("无法分配唯一文档 ID")


def request_client_identity(handler: BaseHTTPRequestHandler) -> str:
    forwarded = handler.headers.get("X-Forwarded-For", "")
    candidate = forwarded.split(",", 1)[0].strip() if forwarded else ""
    if not candidate:
        candidate = str(handler.client_address[0])
    try:
        return str(ipaddress.ip_address(candidate))
    except ValueError:
        return str(handler.client_address[0])


def enforce_document_creation_rate(
    client_identity: str,
    now: float | None = None,
) -> None:
    current = time.monotonic() if now is None else float(now)
    with DOCUMENT_RATE_LIMIT_LOCK:
        tokens, updated = DOCUMENT_RATE_LIMITS.get(
            client_identity,
            (float(DOCUMENT_CREATE_BURST), current),
        )
        tokens = min(
            float(DOCUMENT_CREATE_BURST),
            tokens + max(0.0, current - updated) * DOCUMENT_CREATE_RATE_PER_SECOND,
        )
        if tokens < 1:
            retry_after = max(
                1,
                math.ceil((1 - tokens) / DOCUMENT_CREATE_RATE_PER_SECOND),
            )
            DOCUMENT_RATE_LIMITS[client_identity] = (tokens, current)
            raise BridgeError(
                429,
                "DOCUMENT_CREATE_RATE_LIMITED",
                "新建文档过于频繁，请稍后重试",
                {"retryAfter": retry_after},
            )
        DOCUMENT_RATE_LIMITS[client_identity] = (tokens - 1, current)


def list_uuid_documents() -> list[dict[str, Any]]:
    directory = document_storage_directory()
    if not os.path.isdir(directory):
        return []
    documents = []
    with os.scandir(directory) as entries:
        for entry in entries:
            if not entry.is_file(follow_symlinks=False):
                continue
            match = DOCUMENT_FILE_PATTERN.fullmatch(entry.name)
            if not match:
                continue
            metadata = entry.stat(follow_symlinks=False)
            file_type = match.group("fileType")
            document_id = match.group("documentId")
            documents.append(
                {
                    "documentId": document_id,
                    "fileName": entry.name,
                    "fileType": file_type,
                    "size": metadata.st_size,
                    "updatedAt": metadata.st_mtime,
                    "editorUrl": document_editor_url(file_type, document_id),
                }
            )
    documents.sort(
        key=lambda item: (float(item["updatedAt"]), str(item["fileName"])),
        reverse=True,
    )
    return documents


def require_document_admin(handler: BaseHTTPRequestHandler) -> None:
    expected_user = os.environ.get("DOCUMENT_ADMIN_USERNAME", "")
    expected_password = os.environ.get("DOCUMENT_ADMIN_PASSWORD", "")
    if not expected_user or not expected_password:
        raise BridgeError(
            503,
            "DOCUMENT_ADMIN_NOT_CONFIGURED",
            "管理员凭证尚未配置",
        )
    authorization = handler.headers.get("Authorization", "")
    if not authorization.startswith("Basic "):
        raise BridgeError(401, "DOCUMENT_ADMIN_AUTH_REQUIRED", "需要管理员认证")
    try:
        decoded = base64.b64decode(
            authorization.removeprefix("Basic ").strip(),
            validate=True,
        ).decode("utf-8")
        provided_user, provided_password = decoded.split(":", 1)
    except (binascii.Error, UnicodeDecodeError, ValueError) as error:
        raise BridgeError(
            401,
            "DOCUMENT_ADMIN_AUTH_REQUIRED",
            "无效的管理员认证",
        ) from error
    user_matches = hmac.compare_digest(provided_user, expected_user)
    password_matches = hmac.compare_digest(provided_password, expected_password)
    if not (user_matches and password_matches):
        raise BridgeError(401, "DOCUMENT_ADMIN_AUTH_REQUIRED", "无效的管理员认证")


def human_file_size(size: int) -> str:
    value = float(size)
    for unit in ("B", "KB", "MB", "GB"):
        if value < 1024 or unit == "GB":
            return f"{value:.0f} {unit}" if unit == "B" else f"{value:.2f} {unit}"
        value /= 1024
    return f"{value:.2f} GB"


def admin_documents_html(documents: list[dict[str, Any]]) -> str:
    rows = []
    for document in documents:
        editor_url = html.escape(str(document["editorUrl"]), quote=True)
        rows.append(
            "<tr>"
            f"<td><code>{html.escape(str(document['documentId']))}</code></td>"
            f"<td>{html.escape(str(document['fileType']).upper())}</td>"
            f"<td>{html.escape(human_file_size(int(document['size'])))}</td>"
            f"<td>{html.escape(time.strftime('%Y-%m-%d %H:%M:%S UTC', time.gmtime(float(document['updatedAt']))))}</td>"
            f'<td><a href="{editor_url}" rel="noreferrer"><code>{editor_url}</code></a></td>'
            "</tr>"
        )
    table_body = "".join(rows) or (
        '<tr><td class="empty" colspan="5">还没有 UUID 文档。</td></tr>'
    )
    return """<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>文档管理</title>
  <style>
    :root { color-scheme: light; font-family: system-ui, sans-serif; color: #1f2937; background: #f7f8fa; }
    body { margin: 0; padding: 32px; }
    main { max-width: 1180px; margin: 0 auto; background: #fff; border: 1px solid #e5e7eb; border-radius: 12px; overflow: hidden; }
    header { padding: 24px; border-bottom: 1px solid #e5e7eb; }
    h1 { margin: 0 0 6px; font-size: 22px; }
    p { margin: 0; color: #6b7280; }
    table { width: 100%; border-collapse: collapse; }
    th, td { padding: 14px 18px; border-bottom: 1px solid #eef0f2; text-align: left; }
    th { font-size: 12px; color: #6b7280; text-transform: uppercase; background: #fafafa; }
    code { font-size: 12px; }
    a { color: #2563eb; text-decoration: none; font-weight: 600; }
    a:hover { text-decoration: underline; }
    .empty { padding: 40px; text-align: center; color: #6b7280; }
  </style>
</head>
<body>
<main>
  <header>
    <h1>UUID 文档管理</h1>
    <p>仅展示通过公开新建接口创建的 DOCX、XLSX 和 PPTX。</p>
  </header>
  <table>
    <thead><tr><th>文档 ID</th><th>格式</th><th>大小</th><th>更新时间</th><th>编辑 URL</th></tr></thead>
    <tbody>""" + table_body + """</tbody>
  </table>
</main>
</body>
</html>"""


def html_response(
    handler: BaseHTTPRequestHandler,
    status: int,
    body: str,
    authenticate: bool = False,
) -> None:
    encoded = body.encode("utf-8")
    try:
        handler.send_response(status)
        handler.send_header("Content-Type", "text/html; charset=utf-8")
        handler.send_header("Content-Length", str(len(encoded)))
        handler.send_header("Cache-Control", "no-store")
        handler.send_header("X-Content-Type-Options", "nosniff")
        handler.send_header("X-Frame-Options", "DENY")
        handler.send_header("Referrer-Policy", "no-referrer")
        handler.send_header(
            "Content-Security-Policy",
            "default-src 'none'; style-src 'unsafe-inline'; "
            "base-uri 'none'; form-action 'none'; frame-ancestors 'none'",
        )
        if authenticate:
            handler.send_header(
                "WWW-Authenticate",
                'Basic realm="ONLYOFFICE documents", charset="UTF-8"',
            )
        handler.end_headers()
        handler.wfile.write(encoded)
    except (BrokenPipeError, ConnectionResetError):
        return


def document_gateway_response(
    handler: BaseHTTPRequestHandler,
    file_type: str,
    document_id: str,
) -> None:
    document_path(file_type, document_id)
    try:
        handler.send_response(200)
        handler.send_header(
            "X-Accel-Redirect",
            f"/__document_editor/{file_type}/{document_id}",
        )
        handler.send_header("Content-Length", "0")
        handler.send_header("Cache-Control", "no-store")
        handler.send_header("X-Content-Type-Options", "nosniff")
        handler.end_headers()
    except (BrokenPipeError, ConnectionResetError):
        return


def validate_calls(editor: str, calls: list[dict[str, Any]]) -> list[dict[str, Any]]:
    if len(calls) > 20:
        raise ValueError("单次最多执行 20 个工具")
    allowed = ALLOWED_BY_EDITOR[editor]
    validated = []
    for call in calls:
        name = call.get("name")
        if name not in allowed:
            raise ValueError(f"当前编辑器不允许工具：{name}")
        arguments = call.get("arguments", {})
        if isinstance(arguments, str):
            arguments = json.loads(arguments or "{}")
        if not isinstance(arguments, dict):
            raise ValueError(f"{name}.arguments 必须是对象")
        validated.append({"id": call.get("id"), "name": name, "arguments": arguments})
    return validated


def fallback_plan(editor: str, message: str) -> dict[str, Any]:
    text = message.strip()
    calls: list[dict[str, Any]] = []

    if editor == "word":
        replace = re.search(r"把\s*[‘'\"“](.+?)[’'\"”]\s*(?:替换|改)为\s*[‘'\"“](.+?)[’'\"”]", text)
        shrink = re.search(r"(?:字体|字号).*?(?:缩小|减小)\s*(\d+(?:\.\d+)?)\s*%", text)
        append = re.search(r"(?:在)?文末(?:添加|追加)(?:一段|段落)?[：:]?\s*(.+)", text)
        if replace:
            calls.append({"name": "word_replace_text", "arguments": {"search": replace.group(1), "replace": replace.group(2)}})
        if shrink:
            calls.append({"name": "word_scale_font", "arguments": {"scale": 1 - float(shrink.group(1)) / 100}})
        if append:
            calls.append({"name": "word_append_paragraph", "arguments": {"text": append.group(1).strip()}})
        if "全文" in text and "加粗" in text:
            calls.append({"name": "word_format_document", "arguments": {"bold": True}})

    elif editor == "slide":
        replace = re.search(r"把\s*[‘'\"“](.+?)[’'\"”]\s*(?:替换|改)为\s*[‘'\"“](.+?)[’'\"”]", text)
        shrink = re.search(r"第\s*(\d+)\s*页.*?(?:字体|字号).*?(?:缩小|减小)\s*(\d+(?:\.\d+)?)\s*%", text)
        duplicate = re.search(r"(?:复制|重复)第\s*(\d+)\s*页", text)
        delete = re.search(r"删除第\s*(\d+)\s*页", text)
        add = re.search(r"(?:新增|新建|添加)一页.*?(?:标题(?:为|是)|主题(?:为|是))[：:]?\s*(.+)", text)
        if replace:
            calls.append({"name": "slides_replace_text", "arguments": {"search": replace.group(1), "replace": replace.group(2)}})
        if shrink:
            calls.append({"name": "slides_scale_font", "arguments": {"slide": int(shrink.group(1)), "scale": 1 - float(shrink.group(2)) / 100}})
        if duplicate:
            calls.append({"name": "slides_duplicate_slide", "arguments": {"slide": int(duplicate.group(1))}})
        if delete:
            calls.append({"name": "slides_delete_slide", "arguments": {"slide": int(delete.group(1))}})
        if add:
            calls.append({"name": "slides_add_slide", "arguments": {"title": add.group(1).strip()}})

    elif editor == "cell":
        replace = re.search(r"把\s*[‘'\"“](.+?)[’'\"”]\s*(?:替换|改)为\s*[‘'\"“](.*?)[’'\"”]", text)
        cell_range = re.search(r"\b([A-Z]+\d+(?::[A-Z]+\d+)?)\b", text.upper())
        if replace:
            calls.append({"name": "sheets_replace_text", "arguments": {"search": replace.group(1), "replace": replace.group(2)}})
        if cell_range and ("加粗" in text or "填充" in text):
            args: dict[str, Any] = {"range": cell_range.group(1)}
            if "加粗" in text:
                args["bold"] = True
            color_match = re.search(r"#([0-9A-Fa-f]{6})", text)
            args["fillColor"] = f"#{color_match.group(1)}" if color_match else "#DDEBF7"
            calls.append({"name": "sheets_format_range", "arguments": args})
        chart = re.search(r"(?:用|基于)\s*([A-Z]+\d+:[A-Z]+\d+).*?(?:图表|柱状图|折线图|饼图)", text.upper())
        if chart:
            chart_type = "line" if "折线" in text else "pie" if "饼" in text else "bar"
            calls.append({"name": "sheets_add_chart", "arguments": {"range": chart.group(1), "type": chart_type}})
        set_value = re.search(r"(?:把|将)\s*([A-Z]+\d+(?::[A-Z]+\d+)?)\s*(?:设为|写入)[：:]?\s*(.+)", text, re.IGNORECASE)
        if set_value and not calls:
            value: Any = set_value.group(2).strip()
            try:
                value = json.loads(value)
            except json.JSONDecodeError:
                pass
            calls.append({"name": "sheets_set_values", "arguments": {"range": set_value.group(1), "values": value}})

    if calls:
        return {"message": f"已生成 {len(calls)} 个受控编辑操作，将在当前文档中执行。", "toolCalls": validate_calls(editor, calls), "provider": "local-fallback"}
    return {
        "message": "当前没有配置大模型，且本地解析器未能识别这条指令。请配置 COPILOT_API_KEY、COPILOT_API_BASE 和 COPILOT_MODEL，或换用明确的编辑描述。",
        "toolCalls": [],
        "provider": "local-fallback",
    }


def call_model(editor: str, request: dict[str, Any]) -> dict[str, Any]:
    api_key = os.environ.get("COPILOT_API_KEY", "").strip()
    model = os.environ.get("COPILOT_MODEL", "").strip()
    if not api_key or not model:
        return fallback_plan(editor, str(request.get("message", "")))

    api_base = os.environ.get("COPILOT_API_BASE", "https://api.openai.com/v1").rstrip("/")
    endpoint = api_base if api_base.endswith("/chat/completions") else f"{api_base}/chat/completions"
    context = compact_json(request.get("context", {}))[:60000]
    history = request.get("history", [])
    messages = [
        {
            "role": "system",
            "content": (
                "你是嵌入 ONLYOFFICE 的文档编辑 Copilot。根据用户目标和当前文档上下文选择工具。"
                "只能调用提供的工具，不能输出或要求执行任意 JavaScript。尽量用最少的工具完成修改。"
                "不要编造页码、工作表名、单元格范围或原文；信息不足时只用文字提出简短澄清，不调用工具。"
                "回复使用中文，简洁说明将做什么。"
            ),
        }
    ]
    for item in history[-8:]:
        role = item.get("role")
        content = item.get("content")
        if role in ("user", "assistant") and isinstance(content, str):
            messages.append({"role": role, "content": content[:8000]})
    messages.append({"role": "user", "content": f"文件：{request.get('fileName') or '未命名'}\n当前内容上下文：{context}\n\n用户要求：{request.get('message', '')}"})
    body = compact_json({
        "model": model,
        "messages": messages,
        "tools": TOOLS_BY_EDITOR[editor],
        "tool_choice": "auto",
        "temperature": 0.1,
    }).encode("utf-8")
    http_request = urllib.request.Request(
        endpoint,
        data=body,
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(http_request, timeout=90) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as error:
        detail = error.read().decode("utf-8", errors="replace")[:2000]
        raise RuntimeError(f"大模型接口返回 {error.code}: {detail}") from error
    message = payload["choices"][0]["message"]
    calls = []
    for entry in message.get("tool_calls") or []:
        function = entry.get("function") or {}
        calls.append({"id": entry.get("id"), "name": function.get("name"), "arguments": function.get("arguments") or {}})
    return {
        "message": message.get("content") or ("已生成编辑操作。" if calls else "没有需要执行的编辑操作。"),
        "toolCalls": validate_calls(editor, calls),
        "provider": "model",
        "model": model,
    }


def get_jwt_secret() -> str:
    explicit = os.environ.get("JWT_SECRET", "")
    if explicit:
        return explicit
    with open(LOCAL_CONFIG, "r", encoding="utf-8") as stream:
        config = json.load(stream)
    return str(config["services"]["CoAuthoring"]["secret"]["inbox"]["string"])


def b64url(value: bytes) -> str:
    return base64.urlsafe_b64encode(value).decode("ascii").rstrip("=")


def sign_jwt(payload: dict[str, Any], secret: str) -> str:
    header = {"alg": "HS256", "typ": "JWT"}
    unsigned = f"{b64url(compact_json(header).encode())}.{b64url(compact_json(payload).encode())}"
    signature = hmac.new(secret.encode(), unsigned.encode(), hashlib.sha256).digest()
    return f"{unsigned}.{b64url(signature)}"


def command_service(command: dict[str, Any]) -> dict[str, Any]:
    request_body = dict(command)
    request_body["token"] = sign_jwt(command, get_jwt_secret())
    http_request = urllib.request.Request(
        "http://127.0.0.1/coauthoring/CommandService.ashx",
        data=compact_json(request_body).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(http_request, timeout=20) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as error:
        detail = error.read().decode("utf-8", errors="replace")[:2000]
        raise RuntimeError(f"CommandService 返回 {error.code}: {detail}") from error


def storage_files(file_name: str | None) -> tuple[list[str], list[str]]:
    if not file_name:
        return [], []
    safe_name = os.path.basename(file_name)
    main_files = [
        path
        for path in glob.glob(os.path.join(EXAMPLE_FILES, "*", safe_name))
        if os.path.isfile(path) and "-history" not in path
    ]
    force_save_files = [
        path
        for path in glob.glob(os.path.join(EXAMPLE_FILES, "*", f"{safe_name}-history", safe_name))
        if os.path.isfile(path)
    ]
    return main_files, force_save_files


def newest_mtime(paths: list[str]) -> float | None:
    mtimes = [os.path.getmtime(path) for path in paths if os.path.isfile(path)]
    return max(mtimes) if mtimes else None


def canonical_file(file_name: str | None) -> str:
    if not file_name:
        raise ValueError("未取得文件名")
    main_files, _ = storage_files(file_name)
    if not main_files:
        raise ValueError(f"找不到原文件：{os.path.basename(file_name)}")
    return max(main_files, key=os.path.getmtime)


def version_root(main_file: str) -> str:
    return f"{main_file}.copilot-versions"


def version_files(main_file: str, kind: str) -> list[str]:
    if kind not in ("undo", "redo"):
        raise ValueError("无效版本类型")
    directory = os.path.join(version_root(main_file), kind)
    return sorted(
        path
        for path in glob.glob(os.path.join(directory, "*.snapshot"))
        if os.path.isfile(path)
    )


def snapshot_file(source: str, main_file: str, kind: str) -> str:
    directory = os.path.join(version_root(main_file), kind)
    os.makedirs(directory, exist_ok=True)
    snapshot_id = f"{time.time_ns():020d}-{uuid.uuid4().hex}"
    destination = os.path.join(directory, f"{snapshot_id}.snapshot")
    temporary = f"{destination}.saving"
    shutil.copy2(source, temporary)
    os.replace(temporary, destination)
    snapshots = version_files(main_file, kind)
    for stale in snapshots[:-MAX_VERSION_SNAPSHOTS]:
        os.unlink(stale)
    return destination


def clear_versions(main_file: str, kind: str) -> None:
    for path in version_files(main_file, kind):
        os.unlink(path)


def version_status(file_name: str | None) -> dict[str, Any]:
    main_file = canonical_file(file_name)
    undo = version_files(main_file, "undo")
    redo = version_files(main_file, "redo")
    return {
        "fileName": os.path.basename(main_file),
        "canUndo": bool(undo),
        "canRedo": bool(redo),
        "undoCount": len(undo),
        "redoCount": len(redo),
    }


def create_checkpoint(file_name: str | None) -> dict[str, Any]:
    main_file = canonical_file(file_name)
    snapshot = snapshot_file(main_file, main_file, "undo")
    clear_versions(main_file, "redo")
    return {
        **version_status(file_name),
        "checkpointId": os.path.basename(snapshot).removesuffix(".snapshot"),
    }


def disconnect_editor(key: str | None, user_id: str | None) -> dict[str, Any] | None:
    if not key or not user_id:
        return None
    if len(key) > 512 or len(user_id) > 512:
        raise ValueError("无效的编辑会话")
    result = command_service({"c": "drop", "key": key, "users": [user_id]})
    if result.get("error") not in (0, 1):
        raise RuntimeError(f"CommandService 拒绝断开旧编辑会话：{result}")
    if result.get("error") == 1:
        return result
    # The example callback can perform one final write after the user is dropped.
    # Let that write settle before replacing the canonical file with a snapshot.
    time.sleep(5.5)
    return result


def restore_version(file_name: str | None, direction: str, key: str | None, user_id: str | None) -> dict[str, Any]:
    if direction not in ("undo", "redo"):
        raise ValueError("无效恢复方向")
    main_file = canonical_file(file_name)
    source_kind = direction
    target_kind = "redo" if direction == "undo" else "undo"
    source_versions = version_files(main_file, source_kind)
    if not source_versions:
        label = "撤销" if direction == "undo" else "重做"
        raise ValueError(f"没有可{label}的 Copilot 修改")

    disconnect_editor(key, user_id)
    main_file = canonical_file(file_name)
    source_versions = version_files(main_file, source_kind)
    if not source_versions:
        label = "撤销" if direction == "undo" else "重做"
        raise ValueError(f"没有可{label}的 Copilot 修改")
    source = source_versions[-1]
    snapshot_file(main_file, main_file, target_kind)
    temporary = f"{main_file}.copilot-restoring"
    shutil.copy2(source, temporary)
    os.replace(temporary, main_file)
    os.utime(main_file, None)
    os.unlink(source)
    return {
        **version_status(file_name),
        "restored": True,
        "direction": direction,
        "reload": True,
    }


def promote_force_save(file_name: str) -> tuple[bool, str | None]:
    main_files, force_save_files = storage_files(file_name)
    if not force_save_files:
        return False, None
    source = max(force_save_files, key=os.path.getmtime)
    source_user_dir = os.path.dirname(os.path.dirname(source))
    destination = os.path.join(source_user_dir, os.path.basename(file_name))
    if main_files and destination not in main_files:
        destination = max(main_files, key=os.path.getmtime)
    temporary = f"{destination}.copilot-saving"
    shutil.copy2(source, temporary)
    os.replace(temporary, destination)
    return True, destination


def force_save(key: str, file_name: str | None, allow_no_changes: bool = False) -> dict[str, Any]:
    if not key or len(key) > 512:
        raise ValueError("无效 document.key")
    before_main_files, before_force_files = storage_files(file_name)
    before = newest_mtime(before_main_files)
    before_force = newest_mtime(before_force_files)
    prepromoted = False
    prepromoted_path = None
    if file_name and before_force is not None and (before is None or before_force > before + 0.0001):
        prepromoted, prepromoted_path = promote_force_save(os.path.basename(file_name))
    command: dict[str, Any] = {"c": "forcesave", "key": key}
    if file_name:
        command["userdata"] = compact_json({"fileName": os.path.basename(file_name), "source": "office-copilot"})
    retry_deadline = time.time() + 12
    while True:
        result = command_service(command)
        if result.get("error") != 4 or prepromoted or allow_no_changes or time.time() >= retry_deadline:
            break
        time.sleep(0.6)
    if result.get("error") == 4:
        main_files, _ = storage_files(file_name)
        return {
            "accepted": False,
            "noChanges": True,
            "persisted": bool(main_files),
            "promoted": prepromoted,
            "promotedPath": prepromoted_path,
            "command": result,
            "beforeMtime": before,
            "afterMtime": newest_mtime(main_files),
            "forceSaveMtime": before_force,
            "note": "No editor changes needed saving; the current canonical file remains downloadable.",
        }
    if result.get("error") != 0:
        raise RuntimeError(f"CommandService 拒绝 forcesave：{result}")

    persisted = False
    promoted = prepromoted
    promoted_path = prepromoted_path
    after = before
    after_force = before_force
    if file_name:
        deadline = time.time() + 15
        while time.time() < deadline:
            main_files, force_files = storage_files(file_name)
            after = newest_mtime(main_files)
            after_force = newest_mtime(force_files)
            main_changed = after is not None and (before is None or after > before + 0.0001)
            force_changed = after_force is not None and (before_force is None or after_force > before_force + 0.0001)
            force_is_newer = after_force is not None and (after is None or after_force > after + 0.0001)
            if (force_changed or force_is_newer) and not main_changed:
                promoted, promoted_path = promote_force_save(os.path.basename(file_name))
                main_files, _ = storage_files(file_name)
                after = newest_mtime(main_files)
                main_changed = promoted
            if main_changed:
                persisted = True
                break
            time.sleep(0.35)
    return {
        "accepted": True,
        "persisted": persisted,
        "promoted": promoted,
        "promotedPath": promoted_path,
        "command": result,
        "beforeMtime": before,
        "afterMtime": after,
        "forceSaveMtime": after_force,
    }


class Handler(BaseHTTPRequestHandler):
    server_version = "OnlyOfficeCopilot/1.3"

    def do_GET(self) -> None:
        parsed_path = urllib.parse.urlsplit(self.path)
        public_document = PUBLIC_DOCUMENT_PATH_PATTERN.fullmatch(parsed_path.path)
        if public_document:
            try:
                document_gateway_response(
                    self,
                    public_document.group("fileType"),
                    public_document.group("documentId"),
                )
            except BridgeError as error:
                json_response(self, error.status, error.payload())
            return
        if parsed_path.path in ("/admin", "/admin/"):
            try:
                require_document_admin(self)
                html_response(self, 200, admin_documents_html(list_uuid_documents()))
            except BridgeError as error:
                html_response(
                    self,
                    error.status,
                    (
                        "<!doctype html><html lang=\"zh-CN\"><meta charset=\"utf-8\">"
                        f"<title>文档管理</title><p>{html.escape(error.message)}</p></html>"
                    ),
                    authenticate=error.status == 401,
                )
            except Exception:  # noqa: BLE001 - 管理入口必须保持 fail-closed
                html_response(
                    self,
                    503,
                    (
                        "<!doctype html><html lang=\"zh-CN\"><meta charset=\"utf-8\">"
                        "<title>文档管理</title><p>管理员页面暂不可用</p></html>"
                    ),
                )
            return
        request_path = parsed_path.path.rstrip("/")
        if request_path == "/documents/editor":
            try:
                file_name = urllib.parse.parse_qs(parsed_path.query).get(
                    "fileName",
                    [""],
                )[0]
                file_type, document_id = document_file_identity(file_name)
                document_gateway_response(self, file_type, document_id)
            except BridgeError as error:
                json_response(self, error.status, error.payload())
            return
        if request_path == "/health":
            json_response(self, 200, {"ok": True, "modelConfigured": bool(os.environ.get("COPILOT_API_KEY") and os.environ.get("COPILOT_MODEL")), "editors": ["word", "slide", "cell"]})
            return
        if request_path.startswith("/images/"):
            try:
                asset_id = request_path.removeprefix("/images/")
                token = urllib.parse.parse_qs(parsed_path.query).get("token", [""])[0]
                asset_path, mime_type = verify_image_asset(asset_id, token)
                with open(asset_path, "rb") as stream:
                    body = stream.read()
                self.send_response(200)
                self.send_header("Content-Type", mime_type)
                self.send_header("Content-Length", str(len(body)))
                self.send_header("Cache-Control", "private, max-age=900")
                self.send_header("X-Content-Type-Options", "nosniff")
                self.end_headers()
                self.wfile.write(body)
            except BridgeError as error:
                self.log_bridge_error(error)
                json_response(self, error.status, error.payload())
            except (BrokenPipeError, ConnectionResetError):
                return
            return
        if request_path == "/bridge/sessions":
            try:
                json_response(self, 200, bridge_sessions(bridge_authorization(self)))
            except BridgeError as error:
                self.log_bridge_error(error)
                json_response(self, error.status, error.payload())
            return
        json_response(self, 404, {"error": "Not found"})

    def do_POST(self) -> None:
        payload: dict[str, Any] = {}
        request_path = ""
        try:
            parsed_path = urllib.parse.urlsplit(self.path)
            new_document = NEW_DOCUMENT_PATH_PATTERN.fullmatch(parsed_path.path)
            if new_document:
                try:
                    enforce_document_creation_rate(request_client_identity(self))
                    json_response(
                        self,
                        201,
                        create_document(new_document.group("fileType")),
                    )
                except BridgeError as error:
                    retry_after = (
                        str((error.details or {}).get("retryAfter"))
                        if error.status == 429 and isinstance(error.details, dict)
                        else None
                    )
                    json_response(
                        self,
                        error.status,
                        error.payload(),
                        {"Retry-After": retry_after} if retry_after else None,
                    )
                except Exception:  # noqa: BLE001 - API boundary
                    error = BridgeError(
                        502,
                        "DOCUMENT_CREATE_FAILED",
                        "暂时无法创建文档",
                    )
                    json_response(self, error.status, error.payload())
                return
            request_path = parsed_path.path.rstrip("/")
            max_bytes = (
                IMAGE_MAX_REQUEST_BYTES
                if request_path in ("/images/import", "/bridge/execute")
                else 2_000_000
            )
            payload = read_json(self, max_bytes)
            if request_path == "/images/import":
                json_response(
                    self,
                    200,
                    import_image(payload.get("source"), bridge_authorization(self)),
                )
                return
            if request_path == "/bridge/register":
                json_response(self, 200, bridge_register(payload))
                return
            if request_path == "/bridge/poll":
                json_response(self, 200, bridge_poll(payload))
                return
            if request_path == "/bridge/result":
                json_response(self, 200, bridge_result(payload))
                return
            if request_path == "/bridge/unregister":
                json_response(self, 200, bridge_unregister(payload))
                return
            if request_path == "/bridge/execute":
                json_response(self, 200, bridge_execute(payload, bridge_authorization(self)))
                return
            if request_path == "/editor-config/anonymous":
                json_response(self, 200, issue_anonymous_editor_config(payload))
                return
            if request_path == "/chat":
                editor = str(payload.get("editorType", ""))
                if editor not in TOOLS_BY_EDITOR:
                    raise ValueError(f"不支持 editorType：{editor}")
                if not str(payload.get("message", "")).strip():
                    raise ValueError("message 不能为空")
                json_response(self, 200, call_model(editor, payload))
                return
            if request_path == "/forcesave":
                json_response(
                    self,
                    200,
                    force_save(
                        str(payload.get("key", "")),
                        payload.get("fileName"),
                        bool(payload.get("allowNoChanges")),
                    ),
                )
                return
            if request_path == "/checkpoint":
                json_response(self, 200, create_checkpoint(payload.get("fileName")))
                return
            if request_path == "/history":
                json_response(self, 200, version_status(payload.get("fileName")))
                return
            if request_path in ("/undo", "/redo"):
                direction = request_path.strip("/")
                json_response(
                    self,
                    200,
                    restore_version(
                        payload.get("fileName"),
                        direction,
                        payload.get("key"),
                        payload.get("userId"),
                    ),
                )
                return
            json_response(self, 404, {"error": "Not found"})
        except BridgeError as error:
            self.log_bridge_error(error, payload)
            json_response(self, error.status, error.payload())
        except ValueError as error:
            json_response(self, 400, {"error": str(error)})
        except Exception as error:  # noqa: BLE001 - boundary must return JSON
            if request_path == "/images/import":
                safe_error = image_error(502, "IMAGE_FETCH_FAILED", "图片导入失败")
                self.log_bridge_error(safe_error, payload)
                json_response(self, safe_error.status, safe_error.payload())
            else:
                json_response(self, 502, {"error": str(error)})

    def log_message(self, format: str, *args: Any) -> None:
        if getattr(self, "_suppress_access_log", False):
            return
        rendered = format % args
        rendered = re.sub(
            r"(/images/[0-9a-f]{64}\.(?:png|jpg|gif|webp|svg))\?token=[^ ]+",
            r"\1?token=[redacted]",
            rendered,
        )
        print(f"[{self.log_date_time_string()}] {self.client_address[0]} {rendered}", flush=True)

    def log_bridge_error(self, error: BridgeError, payload: dict[str, Any] | None = None) -> None:
        self._suppress_access_log = bool(error.suppress_log)
        if error.suppress_log:
            return
        session_id = str((payload or {}).get("sessionId", "")).strip()
        request_id = str((payload or {}).get("requestId", "")).strip()
        event = {
            "path": self.path.split("?", 1)[0],
            "status": error.status,
            "code": error.code,
            "client": self.client_address[0],
        }
        if BRIDGE_ID_PATTERN.fullmatch(session_id):
            event["sessionId"] = session_id
        if BRIDGE_ID_PATTERN.fullmatch(request_id):
            event["requestId"] = request_id
        details = error.details if isinstance(error.details, dict) else {}
        phase = str(details.get("phase", "")).strip()
        if phase in BRIDGE_LOG_PHASES:
            event["phase"] = phase
        print(
            "[bridge-error] "
            + compact_json(event),
            flush=True,
        )


if __name__ == "__main__":
    server = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    print(f"ONLYOFFICE Copilot API listening on 127.0.0.1:{PORT}", flush=True)
    server.serve_forever()
