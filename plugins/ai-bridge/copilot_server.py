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
import secrets
import shutil
import threading
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
    "slides_inspect_objects": "只读检查 PPT 图形、图表、图片、表格、组合等对象及填充、线条、位置和可选原始 JSON",
    "slides_replace_text": "在 PPT 全部或指定幻灯片中查找替换",
    "slides_scale_font": "按比例缩放 PPT 全部或指定页字号",
    "slides_format_text": "格式化 PPT 全部或指定页文本",
    "slides_format_selection": "格式化 PPT 当前选中的文本框或形状",
    "slides_add_slide": "新建幻灯片",
    "slides_duplicate_slide": "复制指定幻灯片",
    "slides_delete_slide": "删除指定幻灯片",
    "slides_add_textbox": "在指定 PPT 页添加可设置位置、尺寸、文本、渐变填充和线条的文本框；尺寸单位为毫米",
    "slides_set_background": "设置、清除或恢复 PPT 页背景；自定义背景支持纯色、线性渐变、径向渐变、图案和 raw 填充",
    "slides_add_shape": "在指定 PPT 页添加任意预设图形，并设置文本、几何、渐变填充、线条、旋转和内边距",
    "slides_update_shape": "按 objectId、objectIndex 或 name 更新 PPT 图形的文本、类型、位置、尺寸、旋转、填充和线条",
    "slides_delete_object": "按 objectId、objectIndex 或 name 删除 PPT 页中的任意对象",
    "slides_inspect_charts": "只读检查 PPT 图表类型、标题、系列、位置、样式和可选原始 JSON",
    "slides_add_chart": "用数值系列和分类在 PPT 页创建图表，并设置系列、坐标轴、图例、标签、渐变填充和位置",
    "slides_update_chart": "按 chartId、chartIndex 或 name 更新 PPT 图表系列、分类、坐标轴、图例、标签、颜色、位置和尺寸",
    "slides_delete_chart": "按 chartId、chartIndex 或 name 删除 PPT 图表",
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


def verify_editor_jwt(token: str) -> dict[str, Any]:
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
        editor_config = payload.get("editorConfig") if isinstance(payload.get("editorConfig"), dict) else {}
        user = editor_config.get("user") if isinstance(editor_config.get("user"), dict) else {}
        document_key = str(document.get("key", "")).strip()
        if not document_key:
            raise ValueError("document.key")
        return {
            "documentKey": document_key,
            "fileName": str(document.get("title", "")),
            "fileType": str(document.get("fileType", "")),
            "editorType": str(payload.get("documentType", "")),
            "userId": str(user.get("id", "")),
            "authKind": "editor",
        }
    except BridgeError:
        raise
    except Exception as error:
        raise BridgeError(401, "INVALID_EDITOR_TOKEN", "无效的 ONLYOFFICE 当前编辑器凭证") from error


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
    binding_token, expires_at = bridge_binding_token(claims)
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
        if len(compact_json(arguments)) > 250000:
            raise BridgeError(400, "ARGUMENTS_TOO_LARGE", "arguments 超过 250000 字符")
        params = {"name": name, "arguments": arguments}
    elif method == "executeBatch":
        tool_calls = bridge_parse_json_parameter(payload, "toolCalls", "toolCallsJson", list, [])
        if not tool_calls or len(tool_calls) > 20:
            raise BridgeError(400, "INVALID_TOOL_CALL", "toolCalls 数量必须为 1 到 20")
        for call in tool_calls:
            if not isinstance(call, dict) or call.get("name") not in allowed_tools:
                raise BridgeError(400, "TOOL_NOT_ALLOWED", "批量调用包含当前编辑器不允许的工具")
        if len(compact_json(tool_calls)) > 250000:
            raise BridgeError(400, "ARGUMENTS_TOO_LARGE", "toolCalls 超过 250000 字符")
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
            command = {
                "commandId": command_id,
                "sessionId": session_id,
                "createdAt": time.time(),
                "deliveredAt": None,
                "completedAt": None,
                "response": None,
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
            raise BridgeError(
                409,
                str(error.get("code") or "EXECUTION_FAILED"),
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
            command["response"] = {
                "ok": bool(payload.get("ok")),
                "result": payload.get("result"),
                "error": payload.get("error"),
            }
            command["completedAt"] = time.time()
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
    server_version = "OnlyOfficeCopilot/1.3"

    def do_GET(self) -> None:
        if self.path.rstrip("/") == "/health":
            json_response(self, 200, {"ok": True, "modelConfigured": bool(os.environ.get("COPILOT_API_KEY") and os.environ.get("COPILOT_MODEL")), "editors": ["word", "slide", "cell"]})
            return
        if self.path.rstrip("/") == "/bridge/sessions":
            try:
                json_response(self, 200, bridge_sessions(bridge_authorization(self)))
            except BridgeError as error:
                self.log_bridge_error(error)
                json_response(self, error.status, error.payload())
            return
        json_response(self, 404, {"error": "Not found"})

    def do_POST(self) -> None:
        payload: dict[str, Any] = {}
        try:
            payload = read_json(self)
            if self.path.rstrip("/") == "/bridge/register":
                json_response(self, 200, bridge_register(payload))
                return
            if self.path.rstrip("/") == "/bridge/poll":
                json_response(self, 200, bridge_poll(payload))
                return
            if self.path.rstrip("/") == "/bridge/result":
                json_response(self, 200, bridge_result(payload))
                return
            if self.path.rstrip("/") == "/bridge/unregister":
                json_response(self, 200, bridge_unregister(payload))
                return
            if self.path.rstrip("/") == "/bridge/execute":
                json_response(self, 200, bridge_execute(payload, bridge_authorization(self)))
                return
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
        except BridgeError as error:
            self.log_bridge_error(error, payload)
            json_response(self, error.status, error.payload())
        except ValueError as error:
            json_response(self, 400, {"error": str(error)})
        except Exception as error:  # noqa: BLE001 - boundary must return JSON
            json_response(self, 502, {"error": str(error)})

    def log_message(self, format: str, *args: Any) -> None:
        if getattr(self, "_suppress_access_log", False):
            return
        print(f"[{self.log_date_time_string()}] {self.client_address[0]} {format % args}", flush=True)

    def log_bridge_error(self, error: BridgeError, payload: dict[str, Any] | None = None) -> None:
        self._suppress_access_log = bool(error.suppress_log)
        if error.suppress_log:
            return
        session_id = str((payload or {}).get("sessionId", "")).strip()
        event = {
            "path": self.path.split("?", 1)[0],
            "status": error.status,
            "code": error.code,
            "client": self.client_address[0],
        }
        if BRIDGE_ID_PATTERN.fullmatch(session_id):
            event["sessionId"] = session_id
        print(
            "[bridge-error] "
            + compact_json(event),
            flush=True,
        )


if __name__ == "__main__":
    server = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    print(f"ONLYOFFICE Copilot API listening on 127.0.0.1:{PORT}", flush=True)
    server.serve_forever()
