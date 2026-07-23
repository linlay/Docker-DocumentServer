#!/usr/bin/env python3
"""Planning and persistence API for the headless ONLYOFFICE ai-bridge plugin."""

from __future__ import annotations

import base64
import glob
import hashlib
import hmac
import json
import os
import re
import shutil
import time
import urllib.error
import urllib.request
import uuid
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from typing import Any


PORT = int(os.environ.get("COPILOT_PORT", "3001"))
LOCAL_CONFIG = "/etc/onlyoffice/documentserver/local.json"
EXAMPLE_FILES = "/var/lib/onlyoffice/documentserver-example/files"
MAX_VERSION_SNAPSHOTS = 20


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


WORD_TOOLS = [
    tool("word_replace_text", "在 Word 全文中查找替换文本", {"search": STRING, "replace": STRING, "matchCase": BOOLEAN}, ["search", "replace"]),
    tool("word_append_paragraph", "在 Word 文末追加段落", {"text": STRING, "fontSize": NUMBER, "fontFamily": STRING, "bold": BOOLEAN, "italic": BOOLEAN, "color": STRING, "align": {"type": "string", "enum": ["left", "center", "right", "both"]}, "spacingBefore": NUMBER, "spacingAfter": NUMBER}, ["text"]),
    tool("word_insert_paragraph", "在 Word 当前光标处插入段落", {"text": STRING, "inline": BOOLEAN, "fontSize": NUMBER, "fontFamily": STRING, "bold": BOOLEAN, "italic": BOOLEAN, "color": STRING, "align": STRING}, ["text"]),
    tool("word_format_document", "统一格式化 Word 全文", {"fontSize": NUMBER, "fontFamily": STRING, "bold": BOOLEAN, "italic": BOOLEAN, "underline": BOOLEAN, "strikeout": BOOLEAN, "color": STRING, "align": {"type": "string", "enum": ["left", "center", "right", "both"]}, "spacingBefore": NUMBER, "spacingAfter": NUMBER}),
    tool("word_format_selection", "格式化 Word 当前选中文本；仅当上下文包含 selectionText 或用户明确说选区时使用", {"fontSize": NUMBER, "fontFamily": STRING, "bold": BOOLEAN, "italic": BOOLEAN, "underline": BOOLEAN, "strikeout": BOOLEAN, "color": STRING}),
    tool("word_scale_font", "按比例缩放 Word 全文字号；0.8 表示缩小 20%", {"scale": NUMBER}, ["scale"]),
    tool("word_add_table", "在 Word 文末插入表格", {"rows": {"type": "integer"}, "cols": {"type": "integer"}, "data": {"type": "array", "items": {"type": "array", "items": {}}}, "widthPercent": NUMBER}, ["rows", "cols"]),
    tool("word_set_document_text", "用纯文本重写整个 Word 文档；仅在用户明确要求重写全文时使用", {"text": STRING}, ["text"]),
]


SLIDE_TOOLS = [
    tool("slides_replace_text", "在 PPT 全部或指定幻灯片中查找替换", {"search": STRING, "replace": STRING, "matchCase": BOOLEAN, "slide": {"type": "integer", "minimum": 1}}, ["search", "replace"]),
    tool("slides_scale_font", "按比例缩放 PPT 全部或指定页字号", {"slide": {"type": "integer", "minimum": 1}, "scale": NUMBER}, ["scale"]),
    tool("slides_format_text", "格式化 PPT 全部或指定页文本", {"slide": {"type": "integer", "minimum": 1}, "fontSize": NUMBER, "fontFamily": STRING, "bold": BOOLEAN, "italic": BOOLEAN, "underline": BOOLEAN, "color": STRING}),
    tool("slides_format_selection", "格式化 PPT 当前选中的文本框或形状", {"fontSize": NUMBER, "fontFamily": STRING, "bold": BOOLEAN, "italic": BOOLEAN, "underline": BOOLEAN, "color": STRING}),
    tool("slides_add_slide", "新建幻灯片", {"index": {"type": "integer", "minimum": 1}, "title": STRING, "titleFontSize": NUMBER, "backgroundColor": STRING}),
    tool("slides_duplicate_slide", "复制指定幻灯片", {"slide": {"type": "integer", "minimum": 1}}, ["slide"]),
    tool("slides_delete_slide", "删除指定幻灯片", {"slide": {"type": "integer", "minimum": 1}}, ["slide"]),
    tool("slides_add_textbox", "在指定 PPT 页添加文本框；位置与尺寸单位为毫米", {"slide": {"type": "integer", "minimum": 1}, "text": STRING, "xMm": NUMBER, "yMm": NUMBER, "widthMm": NUMBER, "heightMm": NUMBER, "fontSize": NUMBER, "fontFamily": STRING, "bold": BOOLEAN, "italic": BOOLEAN, "color": STRING, "fillColor": STRING, "align": STRING}, ["slide", "text"]),
]


SHEET_TOOLS = [
    tool("sheets_set_values", "向 XLSX 单元格区域写入单个值或二维数组", {"sheet": SHEET, "range": RANGE, "values": {}}, ["range", "values"]),
    tool("sheets_set_formula", "向 XLSX 单元格写入公式", {"sheet": SHEET, "range": RANGE, "formula": STRING}, ["range", "formula"]),
    tool("sheets_replace_text", "在 XLSX 使用区域或指定区域查找替换", {"sheet": SHEET, "range": RANGE, "search": STRING, "replace": STRING}, ["search", "replace"]),
    tool("sheets_format_range", "设置 XLSX 区域字体、填充、对齐、数字格式和行列尺寸", {"sheet": SHEET, "range": RANGE, "fontSize": NUMBER, "fontName": STRING, "bold": BOOLEAN, "italic": BOOLEAN, "underline": BOOLEAN, "fontColor": STRING, "fillColor": STRING, "horizontalAlign": STRING, "verticalAlign": STRING, "numberFormat": STRING, "wrap": BOOLEAN, "columnWidth": NUMBER, "rowHeight": NUMBER}, ["range"]),
    tool("sheets_add_sheet", "新建工作表", {"name": STRING}, ["name"]),
    tool("sheets_rename_sheet", "重命名工作表", {"sheet": SHEET, "newName": STRING}, ["newName"]),
    tool("sheets_delete_sheet", "删除工作表", {"sheet": SHEET}, ["sheet"]),
    tool("sheets_add_chart", "基于区域数据创建图表", {"sheet": SHEET, "range": RANGE, "type": STRING, "title": STRING, "titleFontSize": NUMBER, "inRows": BOOLEAN, "style": NUMBER, "widthMm": NUMBER, "heightMm": NUMBER, "fromColumn": NUMBER, "fromRow": NUMBER}, ["range"]),
]


TOOLS_BY_EDITOR = {"word": WORD_TOOLS, "slide": SLIDE_TOOLS, "cell": SHEET_TOOLS}
PREFIX_BY_EDITOR = {"word": "word_", "slide": "slides_", "cell": "sheets_"}
ALLOWED_BY_EDITOR = {
    editor: {entry["function"]["name"] for entry in entries}
    for editor, entries in TOOLS_BY_EDITOR.items()
}


def compact_json(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


def json_response(handler: BaseHTTPRequestHandler, status: int, payload: dict[str, Any]) -> None:
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    try:
        handler.send_response(status)
        handler.send_header("Content-Type", "application/json; charset=utf-8")
        handler.send_header("Content-Length", str(len(body)))
        handler.send_header("Cache-Control", "no-store")
        handler.send_header("X-Content-Type-Options", "nosniff")
        handler.end_headers()
        handler.wfile.write(body)
    except (BrokenPipeError, ConnectionResetError):
        return


def read_json(handler: BaseHTTPRequestHandler) -> dict[str, Any]:
    length = int(handler.headers.get("Content-Length", "0"))
    if length <= 0 or length > 2_000_000:
        raise ValueError("请求体为空或过大")
    payload = json.loads(handler.rfile.read(length).decode("utf-8"))
    if not isinstance(payload, dict):
        raise ValueError("请求体必须是 JSON 对象")
    return payload


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
    server_version = "OnlyOfficeCopilot/1.2"

    def do_GET(self) -> None:
        if self.path.rstrip("/") == "/health":
            json_response(self, 200, {"ok": True, "modelConfigured": bool(os.environ.get("COPILOT_API_KEY") and os.environ.get("COPILOT_MODEL")), "editors": ["word", "slide", "cell"]})
            return
        json_response(self, 404, {"error": "Not found"})

    def do_POST(self) -> None:
        try:
            payload = read_json(self)
            if self.path.rstrip("/") == "/chat":
                editor = str(payload.get("editorType", ""))
                if editor not in TOOLS_BY_EDITOR:
                    raise ValueError(f"不支持 editorType：{editor}")
                if not str(payload.get("message", "")).strip():
                    raise ValueError("message 不能为空")
                json_response(self, 200, call_model(editor, payload))
                return
            if self.path.rstrip("/") == "/forcesave":
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
            if self.path.rstrip("/") == "/checkpoint":
                json_response(self, 200, create_checkpoint(payload.get("fileName")))
                return
            if self.path.rstrip("/") == "/history":
                json_response(self, 200, version_status(payload.get("fileName")))
                return
            if self.path.rstrip("/") in ("/undo", "/redo"):
                direction = self.path.strip("/")
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
        except ValueError as error:
            json_response(self, 400, {"error": str(error)})
        except Exception as error:  # noqa: BLE001 - boundary must return JSON
            json_response(self, 502, {"error": str(error)})

    def log_message(self, format: str, *args: Any) -> None:
        print(f"[{self.log_date_time_string()}] {self.client_address[0]} {format % args}", flush=True)


if __name__ == "__main__":
    server = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    print(f"ONLYOFFICE Copilot API listening on 127.0.0.1:{PORT}", flush=True)
    server.serve_forever()
