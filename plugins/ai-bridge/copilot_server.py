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
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from typing import Any


PORT = int(os.environ.get("COPILOT_PORT", "3001"))
COPILOT_BIND_ADDRESS = os.environ.get("COPILOT_BIND_ADDRESS", "127.0.0.1")
LOCAL_CONFIG = "/etc/onlyoffice/documentserver/local.json"
DOCUMENT_STORAGE_ROOT = "/var/lib/onlyoffice/copilot/documents"
DOCUMENT_INTERNAL_ORIGIN = "http://127.0.0.1"
DOCUMENT_TEMPLATE_ROOT = (
    "/var/www/onlyoffice/documentserver/document-templates/new/zh-CN"
)
EDITOR_ASSET_REVISION = "0.2.1-39138552b98a8ed8be92ea60ddb7ac0caa4210dba02f9b691fd5ce2a2415d8e9"
EDITOR_TOKEN_TTL_SECONDS = 12 * 60 * 60
DOCUMENT_MAX_SAVE_BYTES = 200 * 1024 * 1024
AI_BRIDGE_GUID = "asc.{A17E5F31-64AA-4E37-9A42-8D430814C2F6}"
LOOPBACK_FRAME_ANCESTOR_PATTERNS = {
    "http://localhost:*",
    "http://127.0.0.1:*",
}
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
DOCUMENT_STORAGE_PATH_PATTERN = re.compile(
    r"^/documents/storage/(?P<action>download|callback)/"
    r"(?P<fileType>docx|xlsx|pptx)/"
    r"(?P<documentId>[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-"
    r"[89ab][0-9a-f]{3}-[0-9a-f]{12})$"
)
DOCUMENT_KEY_PATTERN = re.compile(
    r"^(?P<documentId>[0-9a-f-]{36})\.(?P<revision>[0-9a-f]{24})$"
)
DOCUMENT_CREATE_RATE_PER_SECOND = 10 / 60
DOCUMENT_CREATE_BURST = 20
DOCUMENT_RATE_LIMIT_LOCK = threading.Lock()
DOCUMENT_RATE_LIMITS: dict[str, tuple[float, float]] = {}
DOCUMENT_SAVE_LOCKS_LOCK = threading.Lock()
DOCUMENT_SAVE_LOCKS: dict[str, threading.Lock] = {}
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
BRIDGE_CONDITION = threading.Condition()
BRIDGE_SESSIONS: dict[str, dict[str, Any]] = {}
BRIDGE_AUTHORITATIVE: dict[tuple[str, str, str, str], str] = {}
BRIDGE_SUPERSEDED: dict[str, dict[str, Any]] = {}
BRIDGE_COMMANDS: dict[str, dict[str, Any]] = {}
BRIDGE_REQUESTS: dict[tuple[tuple[str, str, str, str], str], str] = {}
BRIDGE_REJECTED_CREDENTIALS: dict[str, dict[str, Any]] = {}
BRIDGE_GENERATION = 0


def configured_secret(name: str) -> str:
    direct = os.environ.get(name, "").strip()
    if direct:
        return direct
    path = os.environ.get(f"{name}_FILE", "").strip()
    if not path:
        return ""
    try:
        with open(path, "r", encoding="utf-8") as stream:
            return stream.read().strip()
    except OSError:
        return ""


AI_RELAY_INTERNAL_SECRET = configured_secret("AI_RELAY_INTERNAL_SECRET")

V1_RELAY_POST_PATHS = frozenset(
    {
        "/bridge/register",
        "/bridge/poll",
        "/bridge/result",
        "/bridge/unregister",
        "/bridge/internal/execute",
        "/bridge/internal/validate",
        "/bridge/internal/session",
        "/bridge/internal/drop",
        "/bridge/internal/images/import",
    }
)


def v1_relay_route_allowed(method: str, path: str) -> bool:
    normalized = path.rstrip("/") or "/"
    if method == "POST":
        return normalized in V1_RELAY_POST_PATHS
    if method == "GET":
        return (
            normalized == "/health"
            or normalized.startswith("/bridge/contract/")
            or normalized.startswith("/images/")
        )
    return False


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


def load_public_contract() -> dict[str, Any]:
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
        return json.load(stream)


PUBLIC_API_CONTRACT = load_public_contract()
PUBLIC_API_CONTRACT_VERSION = str(PUBLIC_API_CONTRACT.get("version") or "")
PUBLIC_API_CONTRACT_SHA256 = hashlib.sha256(
    json.dumps(
        PUBLIC_API_CONTRACT,
        ensure_ascii=False,
        separators=(",", ":"),
        sort_keys=True,
    ).encode("utf-8")
).hexdigest()
PUBLIC_API_LIMITS = PUBLIC_API_CONTRACT.get("limits") or {}
PUBLIC_TOOL_NAMING = PUBLIC_API_CONTRACT.get("toolNaming") or {}
PREFIX_BY_EDITOR = dict(PUBLIC_TOOL_NAMING.get("internalPrefixes") or {})
MAX_TOOL_CALLS = int(PUBLIC_API_LIMITS.get("maxToolCalls") or 0)
MAX_ARGUMENTS_JSON_CHARS = int(
    PUBLIC_API_LIMITS.get("maxArgumentsJsonChars") or 0
)
IMAGE_MAX_REQUEST_BYTES = int(
    PUBLIC_API_LIMITS.get("maxImageRequestBytes") or 0
)
IMAGE_MAX_BYTES = int(PUBLIC_API_LIMITS.get("maxImageBytes") or 0)
MAX_REQUEST_ID_CHARS = int(PUBLIC_API_LIMITS.get("maxRequestIdChars") or 0)
MAX_TIMEOUT_MS = int(PUBLIC_API_LIMITS.get("maxTimeoutMs") or 0)
if not all(
    (
        PUBLIC_API_CONTRACT_VERSION,
        MAX_TOOL_CALLS,
        MAX_ARGUMENTS_JSON_CHARS,
        IMAGE_MAX_REQUEST_BYTES,
        IMAGE_MAX_BYTES,
        MAX_REQUEST_ID_CHARS,
        MAX_TIMEOUT_MS,
    )
):
    raise RuntimeError("public-api.json 缺少必需的版本或限制")
if (
    PUBLIC_TOOL_NAMING.get("scope") != "editorType"
    or PUBLIC_TOOL_NAMING.get("publicNames") != "unprefixed"
    or set(PREFIX_BY_EDITOR) != {"word", "slide", "cell"}
    or not all(isinstance(value, str) and value for value in PREFIX_BY_EDITOR.values())
):
    raise RuntimeError("public-api.json 缺少有效的无前缀工具命名策略")
BRIDGE_ID_PATTERN = re.compile(
    rf"^[A-Za-z0-9._:-]{{1,{MAX_REQUEST_ID_CHARS}}}$"
)


def contract_identity() -> dict[str, str]:
    return {
        "contractVersion": PUBLIC_API_CONTRACT_VERSION,
        "contractSha256": PUBLIC_API_CONTRACT_SHA256,
    }


def public_tool_name(editor: str, internal_name: str) -> str:
    prefix = PREFIX_BY_EDITOR.get(editor, "")
    if not prefix or not internal_name.startswith(prefix):
        raise RuntimeError(f"{editor} 内部工具名缺少前缀：{internal_name}")
    public_name = internal_name[len(prefix):]
    if not public_name:
        raise RuntimeError(f"{editor} 内部工具名缺少公开名称：{internal_name}")
    return public_name


def public_tool_schemas(editor: str) -> dict[str, Any]:
    schemas = PUBLIC_API_CONTRACT.get("tools", {}).get(editor)
    if not isinstance(schemas, dict):
        raise BridgeError(404, "EDITOR_MISMATCH", f"不存在编辑器契约：{editor}")
    projected: dict[str, Any] = {}
    for internal_name, schema in schemas.items():
        public_name = public_tool_name(editor, internal_name)
        if public_name in projected:
            raise RuntimeError(f"{editor} 公开工具名重复：{public_name}")
        projected[public_name] = schema
    return projected


def internal_tool_name(editor: str, public_name: str) -> str:
    prefix = PREFIX_BY_EDITOR.get(editor, "")
    if not prefix or not public_name or public_name.startswith(prefix):
        raise BridgeError(
            400,
            "TOOL_NOT_ALLOWED",
            f"当前编辑器不允许公开工具：{public_name or 'unknown'}",
        )
    internal_name = prefix + public_name
    if internal_name not in ALLOWED_BY_EDITOR.get(editor, set()):
        raise BridgeError(
            400,
            "TOOL_NOT_ALLOWED",
            f"当前编辑器不允许公开工具：{public_name}",
        )
    return internal_name


def internalize_public_tool_calls(
    editor: str,
    tool_calls: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    internal_calls: list[dict[str, Any]] = []
    expected_shape = (
        '{"name":"set_size","arguments":{"preset":"wide"}}'
        if editor == "slide"
        else '{"name":"inspect","arguments":{}}'
    )

    def invalid_call(index: int, path: str, message: str) -> None:
        raise BridgeError(
            400,
            "INVALID_TOOL_CALL",
            message,
            {
                "toolCallIndex": index,
                "path": path,
                "expectedShape": expected_shape,
            },
        )

    for index, call in enumerate(tool_calls):
        if not isinstance(call, dict):
            invalid_call(index, f"toolCalls[{index}]", "工具调用必须是对象")

        has_name = "name" in call
        has_action = "action" in call
        raw_name = call.get("name")
        raw_action = call.get("action")
        if has_name and (not isinstance(raw_name, str) or not raw_name.strip()):
            invalid_call(
                index,
                f"toolCalls[{index}].name",
                "工具调用 name 必须是非空字符串",
            )
        if has_action and (not isinstance(raw_action, str) or not raw_action.strip()):
            invalid_call(
                index,
                f"toolCalls[{index}].action",
                "工具调用 action 必须是非空字符串",
            )
        if not has_name and not has_action:
            invalid_call(
                index,
                f"toolCalls[{index}].name",
                "工具调用缺少 name；action 仅作为兼容别名",
            )

        public_name = raw_name.strip() if has_name else raw_action.strip()
        if has_name and has_action and raw_name.strip() != raw_action.strip():
            invalid_call(
                index,
                f"toolCalls[{index}].action",
                "工具调用 name 与 action 不一致",
            )
        if has_action:
            print(
                "[bridge-tool-call-normalization] "
                + compact_json(
                    {
                        "editorType": editor,
                        "toolCallIndex": index,
                        "kind": "actionAlias",
                        "name": public_name,
                    }
                ),
                flush=True,
            )

        arguments = call.get("arguments", {})
        if not isinstance(arguments, dict):
            invalid_call(
                index,
                f"toolCalls[{index}].arguments",
                "工具调用 arguments 必须是对象",
            )

        normalized = {key: value for key, value in call.items() if key != "action"}
        normalized["name"] = internal_tool_name(editor, public_name)
        normalized["arguments"] = arguments
        internal_calls.append(normalized)
    return internal_calls


def project_public_capabilities(editor: str, capabilities: Any) -> Any:
    if not isinstance(capabilities, dict):
        return capabilities
    projected = dict(capabilities)
    tools = capabilities.get("tools")
    if isinstance(tools, list):
        projected["tools"] = [
            public_tool_name(editor, name)
            for name in tools
            if isinstance(name, str) and name.startswith(PREFIX_BY_EDITOR[editor])
        ]
    return projected


def project_public_tool_fields(editor: str, value: Any) -> Any:
    if editor not in PREFIX_BY_EDITOR:
        return value
    if isinstance(value, list):
        return [project_public_tool_fields(editor, item) for item in value]
    if not isinstance(value, dict):
        return value
    projected: dict[str, Any] = {}
    prefix = PREFIX_BY_EDITOR[editor]
    for key, item in value.items():
        if key == "capabilities":
            projected[key] = project_public_capabilities(editor, item)
        elif key in {"name", "tool"} and isinstance(item, str) and item.startswith(prefix):
            projected[key] = public_tool_name(editor, item)
        else:
            projected[key] = project_public_tool_fields(editor, item)
    return projected


def project_bridge_error(editor: str, error: "BridgeError") -> "BridgeError":
    error.details = project_public_tool_fields(editor, error.details)
    return error


def editor_public_contract(editor: str) -> dict[str, Any]:
    schemas = public_tool_schemas(editor)
    units_key = {"word": "word", "slide": "slide", "cell": "sheet"}[editor]
    normalization_key = {"word": "word", "slide": "slide", "cell": "sheet"}[editor]
    return {
        "name": PUBLIC_API_CONTRACT.get("name"),
        "version": PUBLIC_API_CONTRACT_VERSION,
        "protocolVersion": PUBLIC_API_CONTRACT.get("protocolVersion"),
        "contractSha256": PUBLIC_API_CONTRACT_SHA256,
        "editorType": editor,
        "toolPrefix": "",
        "limits": PUBLIC_API_LIMITS,
        "unitConventions": (
            (PUBLIC_API_CONTRACT.get("unitConventions") or {}).get(units_key) or {}
        ),
        "inputNormalization": (
            (PUBLIC_API_CONTRACT.get("inputNormalization") or {}).get(
                normalization_key
            )
            or {}
        ),
        "controls": PUBLIC_API_CONTRACT.get("controls") or [],
        "errors": PUBLIC_API_CONTRACT.get("errors") or [],
        "$defs": PUBLIC_API_CONTRACT.get("$defs") or {},
        "tools": schemas,
    }


def load_contract_tools(editor: str) -> list[dict[str, Any]]:
    contract = PUBLIC_API_CONTRACT
    schemas = contract.get("tools", {}).get(editor)
    if not isinstance(schemas, dict) or not schemas:
        raise RuntimeError(f"public-api.json 缺少 {editor} 工具契约")
    definitions = contract.get("$defs", {})
    missing_descriptions = sorted(
        name
        for name, schema in schemas.items()
        if not isinstance(schema, dict)
        or not isinstance(schema.get("description"), str)
        or not schema["description"].strip()
    )
    if missing_descriptions:
        raise RuntimeError(
            f"{editor} 工具契约缺少 description：{', '.join(missing_descriptions)}"
        )
    for name, schema in schemas.items():
        if schema.get("x-effects") not in {"read", "write"}:
            raise RuntimeError(
                f"public-api.json 的 {editor}.{name}.x-effects 必须是 read 或 write"
            )
        if schema.get("x-argumentLimitClass") not in {"standard", "image"}:
            raise RuntimeError(
                f"public-api.json 的 {editor}.{name}.x-argumentLimitClass 无效"
            )
    return [
        {
            "type": "function",
            "function": {
                "name": name,
                "description": schema["description"],
                "parameters": resolve_contract_schema(schema, definitions),
            },
        }
        for name, schema in schemas.items()
    ]


WORD_TOOLS = load_contract_tools("word")




SLIDE_TOOLS = load_contract_tools("slide")
SHEET_TOOLS = load_contract_tools("cell")


TOOLS_BY_EDITOR = {"word": WORD_TOOLS, "slide": SLIDE_TOOLS, "cell": SHEET_TOOLS}
ALLOWED_BY_EDITOR = {
    editor: {entry["function"]["name"] for entry in entries}
    for editor, entries in TOOLS_BY_EDITOR.items()
}
PUBLIC_ALLOWED_BY_EDITOR = {
    editor: {public_tool_name(editor, name) for name in names}
    for editor, names in ALLOWED_BY_EDITOR.items()
}
if any(
    len(PUBLIC_ALLOWED_BY_EDITOR[editor]) != len(ALLOWED_BY_EDITOR[editor])
    for editor in ALLOWED_BY_EDITOR
):
    raise RuntimeError("同一 editorType 内存在重复的公开工具名")
ARGUMENT_SCHEMAS_BY_EDITOR = {
    editor: {
        entry["function"]["name"]: entry["function"]["parameters"]
        for entry in entries
    }
    for editor, entries in TOOLS_BY_EDITOR.items()
}
TOOL_SCHEMAS = {
    name: schema
    for schemas in ARGUMENT_SCHEMAS_BY_EDITOR.values()
    for name, schema in schemas.items()
}
IMAGE_SOURCE_TOOLS = {
    name
    for name, schema in TOOL_SCHEMAS.items()
    if schema.get("x-argumentLimitClass") == "image"
}


def argument_limit_class_for_tool_calls(
    tool_calls: list[dict[str, Any]],
) -> str:
    return (
        "image"
        if any(
            TOOL_SCHEMAS.get(str(call.get("name") or ""), {}).get(
                "x-argumentLimitClass"
            )
            == "image"
            for call in tool_calls
            if isinstance(call, dict)
        )
        else "standard"
    )


def require_arguments_within_limit(
    value: Any,
    tool_calls: list[dict[str, Any]],
    label: str,
) -> None:
    encoded = compact_json(value).encode("utf-8")
    limit_class = argument_limit_class_for_tool_calls(tool_calls)
    if limit_class == "image":
        actual = len(encoded)
        limit = IMAGE_MAX_REQUEST_BYTES
        unit = "bytes"
    else:
        actual = len(encoded.decode("utf-8"))
        limit = MAX_ARGUMENTS_JSON_CHARS
        unit = "characters"
    if actual > limit:
        raise BridgeError(
            400,
            "ARGUMENTS_TOO_LARGE",
            f"{label} exceeds {limit} {unit}",
            {
                "argumentLimitClass": limit_class,
                "actual": actual,
                "limit": limit,
                "unit": unit,
            },
        )


NORMALIZATION_POLICY_BY_EDITOR = {
    "word": (
        (PUBLIC_API_CONTRACT.get("inputNormalization") or {}).get("word") or {}
    ),
    "cell": (
        (PUBLIC_API_CONTRACT.get("inputNormalization") or {}).get("sheet") or {}
    ),
    "slide": (
        (PUBLIC_API_CONTRACT.get("inputNormalization") or {}).get("slide") or {}
    ),
}
WORD_NORMALIZATION_POLICY = NORMALIZATION_POLICY_BY_EDITOR["word"]
WORD_PROPERTY_ALIASES = dict(WORD_NORMALIZATION_POLICY.get("propertyAliases") or {})
WORD_SUSPICIOUS_TWIPS_ALIASES = {
    "firstLineIndent",
    "leftIndent",
    "rightIndent",
}


def normalized_enum_token(value: str) -> str:
    return re.sub(r"[\s_-]+", "", value).casefold()


def schema_branches_for_value(schema: dict[str, Any], value: Any) -> list[dict[str, Any]]:
    branches: list[dict[str, Any]] = []
    for keyword in ("anyOf", "oneOf"):
        candidates = schema.get(keyword)
        if not isinstance(candidates, list):
            continue
        typed = [
            branch
            for branch in candidates
            if isinstance(branch, dict)
            and (
                branch.get("type") is None
                or (
                    isinstance(branch.get("type"), list)
                    and any(
                        isinstance(item, str) and json_schema_type_matches(value, item)
                        for item in branch["type"]
                    )
                )
                or (
                    isinstance(branch.get("type"), str)
                    and json_schema_type_matches(value, branch["type"])
                )
            )
        ]
        branches.extend(typed or [branch for branch in candidates if isinstance(branch, dict)])
    return branches


def normalize_editor_value(
    value: Any,
    schema: dict[str, Any],
    path: str,
    tool_call_index: int,
    normalizations: list[dict[str, Any]],
    property_aliases: dict[str, str],
    enum_aliases: dict[str, str],
    match_enum_tokens: bool,
) -> Any:
    for branch in schema_branches_for_value(schema, value):
        value = normalize_editor_value(
            value,
            branch,
            path,
            tool_call_index,
            normalizations,
            property_aliases,
            enum_aliases,
            match_enum_tokens,
        )

    if isinstance(value, dict):
        properties = schema.get("properties")
        if not isinstance(properties, dict):
            return value
        normalized = dict(value)
        for alias, canonical in property_aliases.items():
            if alias not in normalized or canonical not in properties:
                continue
            alias_path = f"{path}.{alias}"
            canonical_path = f"{path}.{canonical}"
            kind = "canonicalWins" if canonical in normalized else "propertyAlias"
            if canonical not in normalized:
                normalized[canonical] = normalized[alias]
            del normalized[alias]
            normalizations.append(
                {
                    "toolCallIndex": tool_call_index,
                    "path": alias_path,
                    "canonicalPath": canonical_path,
                    "kind": kind,
                }
            )
        for name, child_schema in properties.items():
            if name in normalized and isinstance(child_schema, dict):
                normalized[name] = normalize_editor_value(
                    normalized[name],
                    child_schema,
                    f"{path}.{name}",
                    tool_call_index,
                    normalizations,
                    property_aliases,
                    enum_aliases,
                    match_enum_tokens,
                )
        return normalized

    if isinstance(value, list):
        item_schema = schema.get("items")
        if isinstance(item_schema, dict):
            return [
                normalize_editor_value(
                    item,
                    item_schema,
                    f"{path}[{index}]",
                    tool_call_index,
                    normalizations,
                    property_aliases,
                    enum_aliases,
                    match_enum_tokens,
                )
                for index, item in enumerate(value)
            ]
        return value

    enum_values = schema.get("enum")
    if isinstance(value, str) and isinstance(enum_values, list):
        string_values = [item for item in enum_values if isinstance(item, str)]
        source_token = normalized_enum_token(value)
        explicit_target = enum_aliases.get(source_token)
        if explicit_target in string_values and value != explicit_target:
            normalizations.append(
                {
                    "toolCallIndex": tool_call_index,
                    "path": path,
                    "canonicalPath": path,
                    "kind": "enumAlias",
                }
            )
            return explicit_target
        matches = (
            [
                candidate
                for candidate in string_values
                if normalized_enum_token(candidate) == source_token
            ]
            if match_enum_tokens
            else []
        )
        if len(matches) == 1 and matches[0] != value:
            normalizations.append(
                {
                    "toolCallIndex": tool_call_index,
                    "path": path,
                    "canonicalPath": path,
                    "kind": "enumCanonicalization",
                }
            )
            return matches[0]
    return value


def suspicious_word_unit_errors(tool_calls: list[dict[str, Any]]) -> list[dict[str, Any]]:
    errors: list[dict[str, Any]] = []

    def visit(value: Any, path: str, index: int, name: str) -> None:
        if isinstance(value, dict):
            for field, child in value.items():
                child_path = f"{path}.{field}"
                if (
                    field in WORD_SUSPICIOUS_TWIPS_ALIASES
                    and WORD_PROPERTY_ALIASES.get(field) not in value
                    and isinstance(child, (int, float))
                    and not isinstance(child, bool)
                    and 300 <= abs(child) <= 2880
                    and abs(child % 20) < 0.000001
                ):
                    errors.append(
                        {
                            "toolCallIndex": index,
                            "tool": name,
                            "path": child_path,
                            "keyword": "semantic",
                            "message": (
                                f"{field} is point-valued and appears to contain twips"
                            ),
                        }
                    )
                visit(child, child_path, index, name)
        elif isinstance(value, list):
            for child_index, child in enumerate(value):
                visit(child, f"{path}[{child_index}]", index, name)

    for index, call in enumerate(tool_calls):
        name = str(call.get("name") or "")
        visit(call.get("arguments", {}), "arguments", index, name)
    return errors


def normalize_word_tool_calls(
    tool_calls: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    return normalize_editor_tool_calls("word", tool_calls)


def normalize_editor_tool_calls(
    editor: str,
    tool_calls: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    normalized_calls = json.loads(json.dumps(tool_calls, ensure_ascii=False))
    normalizations: list[dict[str, Any]] = []
    schemas = ARGUMENT_SCHEMAS_BY_EDITOR[editor]
    policy = NORMALIZATION_POLICY_BY_EDITOR.get(editor) or {}
    property_aliases = dict(policy.get("propertyAliases") or {})
    enum_aliases = {
        normalized_enum_token(str(alias)): str(canonical)
        for alias, canonical in (policy.get("enumAliases") or {}).items()
    }
    match_enum_tokens = (
        policy.get("enumTokenMatching")
        == "case-insensitive-separator-insensitive"
    )
    for index, call in enumerate(normalized_calls):
        name = str(call.get("name") or "")
        schema = schemas.get(name)
        if not isinstance(schema, dict):
            continue
        arguments = call.get("arguments", {})
        call["arguments"] = normalize_editor_value(
            arguments,
            schema,
            "arguments",
            index,
            normalizations,
            property_aliases,
            enum_aliases,
            match_enum_tokens,
        )
    return normalized_calls, normalizations


def json_schema_type_matches(value: Any, expected: str) -> bool:
    if expected == "object":
        return isinstance(value, dict)
    if expected == "array":
        return isinstance(value, list)
    if expected == "string":
        return isinstance(value, str)
    if expected == "boolean":
        return isinstance(value, bool)
    if expected == "integer":
        return isinstance(value, int) and not isinstance(value, bool)
    if expected == "number":
        return isinstance(value, (int, float)) and not isinstance(value, bool)
    if expected == "null":
        return value is None
    return True


def json_schema_type_label(expected: Any) -> str:
    if isinstance(expected, list):
        return " or ".join(str(item) for item in expected)
    return str(expected)


def append_argument_validation_error(
    errors: list[dict[str, Any]],
    path: str,
    keyword: str,
    message: str,
    **details: Any,
) -> None:
    error = {
        "path": path,
        "keyword": keyword,
        "message": message,
    }
    error.update(details)
    errors.append(error)


def enum_value_suggestions(value: Any, allowed_values: list[Any]) -> list[Any]:
    if not isinstance(value, str):
        return []
    token = re.sub(r"[^a-z0-9]", "", value.lower())
    if not token:
        return []
    ranked: list[tuple[int, int, int, str]] = []
    for index, candidate in enumerate(allowed_values):
        if not isinstance(candidate, str):
            continue
        candidate_token = re.sub(r"[^a-z0-9]", "", candidate.lower())
        if not candidate_token:
            continue
        if token.startswith(candidate_token) or candidate_token.startswith(token):
            ranked.append(
                (
                    0 if token.startswith(candidate_token) else 1,
                    abs(len(token) - len(candidate_token)),
                    index,
                    candidate,
                )
            )
    ranked.sort()
    return [candidate for _, _, _, candidate in ranked[:3]]


def validate_json_schema(
    value: Any,
    schema: Any,
    path: str,
    errors: list[dict[str, Any]],
) -> None:
    if not isinstance(schema, dict):
        return

    all_of = schema.get("allOf")
    if isinstance(all_of, list):
        for branch in all_of:
            validate_json_schema(value, branch, path, errors)

    conditional = schema.get("if")
    if isinstance(conditional, dict):
        conditional_errors: list[dict[str, Any]] = []
        validate_json_schema(value, conditional, path, conditional_errors)
        selected = schema.get("then") if not conditional_errors else schema.get("else")
        if isinstance(selected, dict):
            validate_json_schema(value, selected, path, errors)

    negated = schema.get("not")
    if isinstance(negated, dict):
        negated_errors: list[dict[str, Any]] = []
        validate_json_schema(value, negated, path, negated_errors)
        if not negated_errors:
            append_argument_validation_error(
                errors,
                path,
                "not",
                f"{path} matches a forbidden shape",
            )

    any_of = schema.get("anyOf")
    if isinstance(any_of, list) and any_of:
        any_of_matched = False
        any_of_failures: list[tuple[Any, list[dict[str, Any]]]] = []
        for branch in any_of:
            branch_errors: list[dict[str, Any]] = []
            validate_json_schema(value, branch, path, branch_errors)
            if not branch_errors:
                any_of_matched = True
                break
            any_of_failures.append((branch, branch_errors))
        if not any_of_matched:
            matching_type_failures = []
            for branch, branch_errors in any_of_failures:
                expected_type = branch.get("type") if isinstance(branch, dict) else None
                expected_types = (
                    expected_type if isinstance(expected_type, list) else [expected_type]
                )
                if any(
                    isinstance(item, str) and json_schema_type_matches(value, item)
                    for item in expected_types
                ):
                    matching_type_failures.append(branch_errors)
            if matching_type_failures:
                errors.extend(min(matching_type_failures, key=len))
            else:
                append_argument_validation_error(
                    errors,
                    path,
                    "anyOf",
                    f"{path} must match at least one allowed shape",
                )
            return

    one_of = schema.get("oneOf")
    if isinstance(one_of, list) and one_of:
        one_of_matches = 0
        for branch in one_of:
            branch_errors = []
            validate_json_schema(value, branch, path, branch_errors)
            if not branch_errors:
                one_of_matches += 1
        if one_of_matches != 1:
            append_argument_validation_error(
                errors,
                path,
                "oneOf",
                f"{path} must match exactly one allowed shape",
            )
            return

    expected_type = schema.get("type")
    if expected_type is not None:
        expected_types = expected_type if isinstance(expected_type, list) else [expected_type]
        if not any(
            isinstance(item, str) and json_schema_type_matches(value, item)
            for item in expected_types
        ):
            append_argument_validation_error(
                errors,
                path,
                "type",
                f"{path} must be {json_schema_type_label(expected_type)}",
            )
            return

    enum_values = schema.get("enum")
    if isinstance(enum_values, list) and not any(
        type(value) is type(candidate) and value == candidate
        for candidate in enum_values
    ):
        suggestions = enum_value_suggestions(value, enum_values)
        append_argument_validation_error(
            errors,
            path,
            "enum",
            (
                f"{path} must be one of the allowed values"
                + (
                    f"; suggested: {', '.join(str(item) for item in suggestions)}"
                    if suggestions
                    else ""
                )
            ),
            received=value,
            allowedValues=enum_values,
            suggestedValues=suggestions,
        )

    if "const" in schema and (
        type(value) is not type(schema["const"])
        or value != schema["const"]
    ):
        append_argument_validation_error(
            errors,
            path,
            "const",
            f"{path} must use the required constant value",
        )

    if isinstance(value, dict):
        minimum_properties = schema.get("minProperties")
        maximum_properties = schema.get("maxProperties")
        if isinstance(minimum_properties, int) and len(value) < minimum_properties:
            append_argument_validation_error(
                errors,
                path,
                "minProperties",
                f"{path} must contain at least {minimum_properties} properties",
            )
        if isinstance(maximum_properties, int) and len(value) > maximum_properties:
            append_argument_validation_error(
                errors,
                path,
                "maxProperties",
                f"{path} must contain at most {maximum_properties} properties",
            )
        required = schema.get("required")
        if isinstance(required, list):
            for name in required:
                if isinstance(name, str) and name not in value:
                    append_argument_validation_error(
                        errors,
                        f"{path}.{name}",
                        "required",
                        f"{name} is required",
                    )
        properties = schema.get("properties")
        if isinstance(properties, dict):
            if schema.get("additionalProperties") is False:
                for name in value:
                    if name not in properties:
                        append_argument_validation_error(
                            errors,
                            f"{path}.{name}",
                            "additionalProperties",
                            f"{name} is not allowed",
                        )
            for name, child_schema in properties.items():
                if name in value:
                    validate_json_schema(
                        value[name],
                        child_schema,
                        f"{path}.{name}",
                        errors,
                    )

    if isinstance(value, list):
        minimum_items = schema.get("minItems")
        maximum_items = schema.get("maxItems")
        if isinstance(minimum_items, int) and len(value) < minimum_items:
            append_argument_validation_error(
                errors,
                path,
                "minItems",
                f"{path} must contain at least {minimum_items} items",
            )
        if isinstance(maximum_items, int) and len(value) > maximum_items:
            append_argument_validation_error(
                errors,
                path,
                "maxItems",
                f"{path} must contain at most {maximum_items} items",
            )
        if schema.get("uniqueItems") is True:
            encoded_items = [
                json.dumps(
                    item,
                    ensure_ascii=False,
                    separators=(",", ":"),
                    sort_keys=True,
                )
                for item in value
            ]
            if len(encoded_items) != len(set(encoded_items)):
                append_argument_validation_error(
                    errors,
                    path,
                    "uniqueItems",
                    f"{path} must not contain duplicate items",
                )
        item_schema = schema.get("items")
        if isinstance(item_schema, dict):
            for index, item in enumerate(value):
                validate_json_schema(item, item_schema, f"{path}[{index}]", errors)

    if isinstance(value, str):
        minimum_length = schema.get("minLength")
        maximum_length = schema.get("maxLength")
        if isinstance(minimum_length, int) and len(value) < minimum_length:
            append_argument_validation_error(
                errors,
                path,
                "minLength",
                f"{path} must contain at least {minimum_length} characters",
            )
        if isinstance(maximum_length, int) and len(value) > maximum_length:
            append_argument_validation_error(
                errors,
                path,
                "maxLength",
                f"{path} must contain at most {maximum_length} characters",
            )
        pattern = schema.get("pattern")
        if isinstance(pattern, str) and re.search(pattern, value) is None:
            append_argument_validation_error(
                errors,
                path,
                "pattern",
                f"{path} does not match the required pattern",
            )

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        minimum = schema.get("minimum")
        maximum = schema.get("maximum")
        exclusive_minimum = schema.get("exclusiveMinimum")
        exclusive_maximum = schema.get("exclusiveMaximum")
        if isinstance(minimum, (int, float)) and value < minimum:
            append_argument_validation_error(
                errors,
                path,
                "minimum",
                f"{path} must be greater than or equal to {minimum}",
            )
        if isinstance(maximum, (int, float)) and value > maximum:
            append_argument_validation_error(
                errors,
                path,
                "maximum",
                f"{path} must be less than or equal to {maximum}",
            )
        if isinstance(exclusive_minimum, (int, float)) and value <= exclusive_minimum:
            append_argument_validation_error(
                errors,
                path,
                "exclusiveMinimum",
                f"{path} must be greater than {exclusive_minimum}",
            )
        if isinstance(exclusive_maximum, (int, float)) and value >= exclusive_maximum:
            append_argument_validation_error(
                errors,
                path,
                "exclusiveMaximum",
                f"{path} must be less than {exclusive_maximum}",
            )


def validate_word_semantics(
    name: str,
    arguments: Any,
    errors: list[dict[str, Any]],
) -> None:
    if not isinstance(arguments, dict):
        return

    def semantic_error(path: str, message: str) -> None:
        append_argument_validation_error(errors, f"arguments.{path}", "semantic", message)

    if name == "word_manage_section" and arguments.get("action", "configure") == "create":
        if not any(field in arguments for field in ("boundary", "endParagraphIndex", "paragraphIndex")):
            semantic_error("boundary", "boundary or endParagraphIndex is required when action is create")

    if name == "word_format_table_advanced":
        if "repeatHeader" in arguments and "row" not in arguments:
            semantic_error("row", "row is required when repeatHeader is provided")

    if name == "word_manage_fields" and arguments.get("action") == "add":
        instruction = arguments.get("instruction")
        if "instruction" in arguments and (
            not isinstance(instruction, str) or not instruction.strip()
        ):
            semantic_error("instruction", "instruction is required when action is add")

    if name == "word_insert_page_break":
        has_target = (
            arguments.get("all") is True
            or arguments.get("current") is True
            or "paragraphIndex" in arguments
            or "paragraphId" in arguments
            or "internalId" in arguments
            or (
                isinstance(arguments.get("paragraphIndexes"), list)
                and bool(arguments["paragraphIndexes"])
            )
            or (
                isinstance(arguments.get("search"), str)
                and bool(arguments["search"])
            )
        )
        if not has_target:
            semantic_error(
                "target",
                "paragraphId, paragraphIndex, paragraphIndexes, search, all=true, or current=true is required",
            )

    if name == "word_edit_table":
        action = arguments.get("action")
        if action == "mergeCells" and all(
            isinstance(arguments.get(field), int)
            and not isinstance(arguments.get(field), bool)
            for field in ("rowStart", "rowEnd", "columnStart", "columnEnd")
        ):
            if arguments["rowStart"] > arguments["rowEnd"]:
                semantic_error("rowEnd", "rowEnd must be greater than or equal to rowStart")
            if arguments["columnStart"] > arguments["columnEnd"]:
                semantic_error(
                    "columnEnd",
                    "columnEnd must be greater than or equal to columnStart",
                )

    if name == "word_set_header_footer" and arguments.get("action", "set") == "set":
        content_fields = {
            "text",
            "pageNumber",
            "pagesCount",
            "fields",
            "fontSize",
            "fontFamily",
            "bold",
            "italic",
            "color",
            "align",
        }
        if not any(field in arguments for field in content_fields):
            semantic_error(
                "action",
                "set requires text, a page field, fields, or formatting",
            )


def validate_slide_semantics(
    name: str,
    arguments: Any,
    errors: list[dict[str, Any]],
) -> None:
    if not isinstance(arguments, dict):
        return

    def semantic_error(path: str, message: str) -> None:
        append_argument_validation_error(errors, f"arguments.{path}", "semantic", message)

    if name in {"slides_add_textbox", "slides_add_shape"}:
        if "text" in arguments and "paragraphs" in arguments:
            semantic_error("paragraphs", "text and paragraphs are mutually exclusive")

    if name == "slides_add_slide" and "masterIndex" in arguments and "layoutIndex" not in arguments:
        semantic_error("layoutIndex", "layoutIndex is required when masterIndex is provided")

    if name == "slides_validate_layout":
        for index, pair in enumerate(arguments.get("allowedOverlapPairs") or []):
            if (
                isinstance(pair, dict)
                and pair.get("firstName") == pair.get("secondName")
            ):
                semantic_error(
                    f"allowedOverlapPairs[{index}]",
                    "overlap pair must contain two different object names",
                )


SHEET_A1_RANGE_PATTERN = re.compile(
    r"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)"
    r"(?::\$?([A-Za-z]{1,3})\$?([1-9][0-9]*))?$"
)


def sheet_column_number(column: str) -> int:
    number = 0
    for character in column.upper():
        number = number * 26 + ord(character) - ord("A") + 1
    return number


def sheet_a1_range_shape(value: Any) -> tuple[int, int] | None:
    if not isinstance(value, str) or value.casefold() == "selection":
        return None
    if "!" in value:
        sheet_prefix, address = value.rsplit("!", 1)
        if not sheet_prefix:
            return None
    else:
        address = value
    match = SHEET_A1_RANGE_PATTERN.fullmatch(address)
    if match is None:
        return None
    first_column, first_row, last_column, last_row = match.groups()
    if last_column is None:
        last_column = first_column
        last_row = first_row
    first_column_number = sheet_column_number(first_column)
    last_column_number = sheet_column_number(last_column)
    if (
        first_column_number > 16384
        or last_column_number > 16384
        or int(first_row) > 1048576
        or int(last_row) > 1048576
    ):
        return None
    return (
        abs(int(last_row) - int(first_row)) + 1,
        abs(last_column_number - first_column_number) + 1,
    )


def sheet_matrix_shape(value: Any) -> tuple[int, int] | None:
    if not isinstance(value, list) or not value:
        return None
    if not all(isinstance(row, list) and row for row in value):
        return None
    widths = {len(row) for row in value}
    if len(widths) != 1:
        return None
    return len(value), next(iter(widths))


def validate_sheet_semantics(
    name: str,
    arguments: Any,
    errors: list[dict[str, Any]],
) -> None:
    if not isinstance(arguments, dict):
        return

    def semantic_error(path: str, message: str) -> None:
        append_argument_validation_error(errors, f"arguments.{path}", "semantic", message)

    if name in {"sheets_set_values", "sheets_set_formula"}:
        field = "values" if name == "sheets_set_values" else "formula"
        value = arguments.get(field)
        target_range = arguments.get("range")
        target_shape = sheet_a1_range_shape(target_range)
        if (
            isinstance(target_range, str)
            and target_range.casefold() != "selection"
            and target_shape is None
        ):
            semantic_error(
                "range",
                "range must be a finite A1 cell range or selection",
            )
        if isinstance(value, list):
            matrix_shape = sheet_matrix_shape(value)
            if matrix_shape is None:
                semantic_error(field, f"{field} must be a non-empty rectangular 2D matrix")
            elif target_shape is not None and matrix_shape != target_shape:
                semantic_error(
                    field,
                    f"{field} matrix shape {matrix_shape[0]}x{matrix_shape[1]} "
                    f"must match target range shape {target_shape[0]}x{target_shape[1]}",
                )
        elif target_shape is not None and target_shape != (1, 1):
            semantic_error(
                field,
                f"scalar {field} is only allowed for a single-cell target",
            )

    if name == "sheets_manage_validation":
        action = arguments.get("action")
        if action in {"add", "modify"}:
            validation_type = arguments.get("type")
            if validation_type is None:
                semantic_error("type", "type is required when action is add or modify")
            input_only_types = {"inputOnly", "xlValidateInputOnly"}
            if validation_type not in input_only_types and "formula1" not in arguments:
                semantic_error(
                    "formula1",
                    "formula1 is required for this validation type",
                )
            comparison_types = {
                "wholeNumber",
                "decimal",
                "date",
                "time",
                "textLength",
                "xlValidateWholeNumber",
                "xlValidateDecimal",
                "xlValidateDate",
                "xlValidateTime",
                "xlValidateTextLength",
            }
            between_operators = {
                None,
                "between",
                "notBetween",
                "xlBetween",
                "xlNotBetween",
            }
            if (
                validation_type in comparison_types
                and arguments.get("operator") in between_operators
                and "formula2" not in arguments
            ):
                semantic_error(
                    "formula2",
                    "formula2 is required for between/notBetween validation",
                )

    if name == "sheets_manage_table":
        action = arguments.get("action")
        table_mode = arguments.get("tableMode", "auto")
        structured_fields = {
            "sourceType",
            "name",
            "newName",
            "style",
            "showTotals",
            "showHeaders",
            "rowStripes",
            "columnStripes",
            "firstColumn",
            "lastColumn",
            "showAutoFilter",
            "showAutoFilterDropDown",
            "summary",
            "alternativeText",
        }
        provided_structured_fields = sorted(structured_fields.intersection(arguments))
        allowed_create_fields = {
            "action",
            "sheet",
            "range",
            "tableMode",
            "sourceType",
            "name",
            "style",
            "showTotals",
            "showHeaders",
            "rowStripes",
            "columnStripes",
            "firstColumn",
            "lastColumn",
            "showAutoFilter",
            "showAutoFilterDropDown",
            "summary",
            "alternativeText",
        }
        if action == "create":
            invalid_create_fields = sorted(set(arguments).difference(allowed_create_fields))
            if invalid_create_fields:
                semantic_error(
                    invalid_create_fields[0],
                    "field is not allowed when action is create",
                )
        if action != "create" and "tableMode" in arguments:
            semantic_error("tableMode", "tableMode is only allowed when action is create")
        if (
            action == "create"
            and table_mode in {"basic", "rangeStyle"}
            and provided_structured_fields
        ):
            semantic_error(
                provided_structured_fields[0],
                "range-style tables do not support structured-table properties",
            )
        if action == "format":
            invalid_format_fields = sorted(
                set(arguments).difference({"action", "sheet", "range"})
            )
            if invalid_format_fields:
                semantic_error(
                    invalid_format_fields[0],
                    "action format only accepts sheet and range",
                )

SEMANTIC_VALIDATORS = {
    "word.fieldInstruction": validate_word_semantics,
    "word.pageBreakTarget": validate_word_semantics,
    "word.tableRangeOrder": validate_word_semantics,
    "word.headerFooterContent": validate_word_semantics,
    "slides.distinctOverlapPair": validate_slide_semantics,
    "sheets.a1MatrixShape": validate_sheet_semantics,
    "sheets.validationRule": validate_sheet_semantics,
    "sheets.tableOperation": validate_sheet_semantics,
}


def validate_semantic_validator_registry() -> None:
    unknown: list[str] = []
    for editor, schemas in ARGUMENT_SCHEMAS_BY_EDITOR.items():
        for name, schema in schemas.items():
            validator_ids = schema.get("x-semanticValidators", [])
            if not isinstance(validator_ids, list) or not all(
                isinstance(item, str) for item in validator_ids
            ):
                raise RuntimeError(
                    f"public-api.json 的 {editor}.{name}.x-semanticValidators 必须是字符串数组"
                )
            unknown.extend(
                f"{editor}.{name}:{validator_id}"
                for validator_id in validator_ids
                if validator_id not in SEMANTIC_VALIDATORS
            )
    if unknown:
        raise RuntimeError(
            "public-api.json 引用了未注册的语义校验器：" + ", ".join(sorted(unknown))
        )


validate_semantic_validator_registry()


def validate_editor_tool_calls(
    editor: str,
    tool_calls: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    validation_errors: list[dict[str, Any]] = []
    schemas = ARGUMENT_SCHEMAS_BY_EDITOR[editor]
    for index, call in enumerate(tool_calls):
        name = str(call.get("name") or "")
        arguments = call.get("arguments", {})
        call_errors: list[dict[str, Any]] = []
        schema = schemas.get(name)
        should_validate_schema = editor in {"word", "slide", "cell"}
        if schema is not None and should_validate_schema:
            validate_json_schema(arguments, schema, "arguments", call_errors)
        if schema is not None:
            for validator_id in schema.get("x-semanticValidators", []):
                SEMANTIC_VALIDATORS[validator_id](name, arguments, call_errors)
        for error in call_errors:
            validation_errors.append(
                {
                    "toolCallIndex": index,
                    "tool": name,
                    **error,
                }
            )
    return validation_errors


def validate_word_tool_calls(tool_calls: list[dict[str, Any]]) -> list[dict[str, Any]]:
    normalized_calls, _ = normalize_word_tool_calls(tool_calls)
    return validate_editor_tool_calls(
        "word",
        normalized_calls,
    ) + suspicious_word_unit_errors(tool_calls)


def require_valid_editor_tool_calls(
    editor: str,
    tool_calls: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    normalized_calls = tool_calls
    normalizations: list[dict[str, Any]] = []
    validation_errors: list[dict[str, Any]] = []
    if editor in NORMALIZATION_POLICY_BY_EDITOR:
        normalized_calls, normalizations = normalize_editor_tool_calls(
            editor,
            tool_calls,
        )
    validation_errors.extend(validate_editor_tool_calls(editor, normalized_calls))
    if editor == "word":
        validation_errors.extend(suspicious_word_unit_errors(tool_calls))
    if validation_errors:
        editor_label = {"word": "Word", "slide": "Slides", "cell": "Sheets"}.get(
            editor,
            editor,
        )
        raise BridgeError(
            422,
            "INVALID_TOOL_ARGUMENTS",
            f"{editor_label} 工具参数校验失败",
            {
                "validationErrors": validation_errors,
                "completedToolCalls": 0,
                "partialMutationPossible": False,
            },
        )
    return normalized_calls, normalizations


def require_valid_word_tool_calls(
    tool_calls: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    return require_valid_editor_tool_calls("word", tool_calls)


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
    "INVALID_TOOL_ARGUMENTS": 422,
    "INVALID_TARGET": 422,
    "INVALID_IMAGE_SOURCE": 422,
    "EXECUTION_FAILED": 500,
    "WORD_API_UNSUPPORTED": 501,
    "SLIDES_API_UNSUPPORTED": 501,
    "SHEETS_API_UNSUPPORTED": 501,
    "SHEETS_RUNTIME_INCOMPATIBLE": 500,
    "SHEETS_CHART_PARTIAL_MUTATION": 500,
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
        "documentId": str(payload.get("documentId", "")).strip(),
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
    stable_document_identity = str(claims.get("documentId", "")).strip()
    if not stable_document_identity:
        stable_document_identity = str(claims.get("fileName", "")).strip()
    identity = (
        stable_document_identity,
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
        os.path.join(document_storage_directory(), ".ai-bridge-images"),
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
                "User-Agent": "OnlyOffice-ai-bridge/0.2.1",
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
            "documentId": claims.get("documentId", ""),
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
            "documentId": str(payload.get("documentId", "")),
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


def require_internal_relay(handler: BaseHTTPRequestHandler) -> None:
    value = handler.headers.get("Authorization", "")
    supplied = value.removeprefix("Bearer ").strip() if value.startswith("Bearer ") else ""
    if (
        not AI_RELAY_INTERNAL_SECRET
        or not supplied
        or not hmac.compare_digest(supplied, AI_RELAY_INTERNAL_SECRET)
    ):
        raise BridgeError(
            401,
            "INTERNAL_RELAY_AUTH_REQUIRED",
            "内部 Relay 认证失败",
        )


def internal_bridge_claims(session_id: Any) -> dict[str, Any]:
    requested = str(session_id or "").strip()
    with BRIDGE_CONDITION:
        bridge_cleanup_locked()
        session = BRIDGE_SESSIONS.get(requested)
        if session is None or not session.get("authoritative"):
            raise BridgeError(401, "INVALID_RELAY_SESSION", "浏览器 Relay 会话已经失效")
        return {
            **bridge_session_claims(session),
            "authKind": "binding",
        }


def internal_bridge_session(session_id: Any) -> dict[str, Any]:
    requested = str(session_id or "").strip()
    with BRIDGE_CONDITION:
        bridge_cleanup_locked()
        session = BRIDGE_SESSIONS.get(requested)
        if session is None or not session.get("authoritative"):
            raise BridgeError(401, "INVALID_RELAY_SESSION", "浏览器 Relay 会话已经失效")
        return {
            "ok": True,
            **bridge_public_session(requested, session),
        }


def internal_bridge_drop(session_id: Any) -> dict[str, Any]:
    requested = str(session_id or "").strip()
    with BRIDGE_CONDITION:
        session = BRIDGE_SESSIONS.pop(requested, None)
        if session is None:
            return {"ok": True, "removed": False}
        identity = session.get("identity")
        if identity and BRIDGE_AUTHORITATIVE.get(identity) == requested:
            del BRIDGE_AUTHORITATIVE[identity]
        for command_id in list(session.get("queue") or []):
            command = BRIDGE_COMMANDS.get(command_id)
            if command is not None and command.get("response") is None:
                command["response"] = {
                    "ok": False,
                    "error": {
                        "code": "SESSION_REVOKED",
                        "message": "编辑器 Relay 会话已撤销",
                    },
                }
        BRIDGE_CONDITION.notify_all()
        return {"ok": True, "removed": True}


def bridge_public_session(session_id: str, session: dict[str, Any]) -> dict[str, Any]:
    state = session.get("state") if isinstance(session.get("state"), dict) else {}
    context = state.get("context") if isinstance(state.get("context"), dict) else {}
    editor_type = str(state.get("editorType") or context.get("editorType") or "")
    return {
        "sessionId": session_id,
        "documentId": context.get("documentId"),
        "ready": bool(state.get("ready")),
        "documentReady": bool(state.get("documentReady")),
        "capabilityProbeComplete": bool(state.get("capabilityProbeComplete")),
        "capabilityProbedAt": state.get("capabilityProbedAt"),
        "saveReady": bool(state.get("saveReady")),
        "saveStatus": state.get("saveStatus"),
        "saveStatusAt": state.get("saveStatusAt"),
        "editorType": editor_type,
        "documentKey": context.get("documentKey"),
        "fileName": context.get("fileName"),
        "fileType": context.get("fileType"),
        "userId": context.get("userId"),
        "capabilities": (
            project_public_capabilities(editor_type, state.get("capabilities"))
            if editor_type in PREFIX_BY_EDITOR
            else state.get("capabilities")
        ),
        "contractVersion": state.get("contractVersion"),
        "contractSha256": state.get("contractSha256"),
        "lastSeen": session.get("lastSeen"),
        "registeredAt": session.get("registeredAt"),
        "generation": session.get("generation"),
        "authoritative": bool(session.get("authoritative")),
    }


def bridge_session_claims(session: dict[str, Any]) -> dict[str, Any]:
    state = session.get("state") if isinstance(session.get("state"), dict) else {}
    context = state.get("context") if isinstance(state.get("context"), dict) else {}
    return {
        "documentId": str(context.get("documentId", "")).strip(),
        "documentKey": session["documentKey"],
        "fileName": str(context.get("fileName", "")).strip(),
        "fileType": str(context.get("fileType", "")).strip(),
        "editorType": str(state.get("editorType") or context.get("editorType") or "").strip(),
        "userId": str(context.get("userId", "")).strip(),
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
    claimed_document_id = str(claims.get("documentId", "")).strip()
    context_document_id = str(context.get("documentId", "")).strip()
    if claimed_document_id and context_document_id != claimed_document_id:
        raise BridgeError(
            409,
            "DOCUMENT_IDENTITY_MISMATCH",
            "编辑器页面与凭证的 documentId 不一致",
            {
                "identityField": "documentId",
                "claimedIdentityHash": hashlib.sha256(claimed_document_id.encode()).hexdigest()[:12],
                "contextIdentityHash": hashlib.sha256(context_document_id.encode()).hexdigest()[:12],
            },
        )
    if context.get("documentKey") != claims["documentKey"]:
        raise BridgeError(409, "DOCUMENT_MISMATCH", "编辑器页面与凭证绑定的 document.key 不一致")
    if (state.get("editorType") or context.get("editorType")) != claims["editorType"]:
        raise BridgeError(409, "EDITOR_MISMATCH", "编辑器页面与凭证的编辑器类型不一致")
    state_contract_sha256 = str(state.get("contractSha256") or "")
    if (
        state_contract_sha256
        and not hmac.compare_digest(
            state_contract_sha256,
            PUBLIC_API_CONTRACT_SHA256,
        )
    ):
        raise BridgeError(
            409,
            "CONTRACT_VERSION_MISMATCH",
            "编辑器页面与 HTTP Relay 使用了不同的 ai-bridge 契约",
            {
                "expectedContractVersion": PUBLIC_API_CONTRACT_VERSION,
                "expectedContractSha256": PUBLIC_API_CONTRACT_SHA256,
                "receivedContractVersion": state.get("contractVersion"),
                "receivedContractSha256": state_contract_sha256,
            },
        )
    stable_keys = ("fileType", "userId") if claimed_document_id else ("fileName", "fileType", "userId")
    for key in stable_keys:
        if str(context.get(key, "")).strip() != str(claims.get(key, "")).strip():
            raise BridgeError(
                409,
                "DOCUMENT_IDENTITY_MISMATCH",
                "编辑器页面与凭证的稳定文档身份不一致",
                {"identityField": key},
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
        claims = bridge_session_claims(session)
        session["state"] = bridge_assert_state_matches(claims, state)
    session["lastSeen"] = time.time()
    return session_id, session


def bridge_attach(payload: dict[str, Any]) -> dict[str, Any]:
    file_name = str(payload.get("fileName", "")).strip()
    editor_type = str(payload.get("editorType", "")).strip()
    if editor_type not in DOCUMENT_TYPES.values():
        raise BridgeError(
            422,
            "INVALID_ARGUMENTS",
            "editorType 必须是 word、cell 或 slide",
        )

    file_type, _ = document_file_identity(file_name)
    expected_editor_type = DOCUMENT_TYPES[file_type]
    if editor_type != expected_editor_type:
        raise BridgeError(
            409,
            "EDITOR_MISMATCH",
            "文件类型与目标编辑器类型不一致",
        )

    deadline = time.time() + BRIDGE_DISCOVERY_WAIT_SECONDS
    selected_id = ""
    selected: dict[str, Any] | None = None
    with BRIDGE_CONDITION:
        while True:
            bridge_cleanup_locked()
            candidates = [
                (session_id, session)
                for session_id, session in BRIDGE_SESSIONS.items()
                if session.get("authoritative")
                and bool((session.get("state") or {}).get("ready"))
                and session.get("identity")
                and session["identity"][0] == file_name
                and session["identity"][1] == file_type
                and session["identity"][2] == editor_type
            ]
            if candidates:
                candidates.sort(
                    key=lambda item: int(item[1].get("generation") or 0),
                    reverse=True,
                )
                selected_id, selected = candidates[0]
                break
            remaining = deadline - time.time()
            if remaining <= 0:
                break
            BRIDGE_CONDITION.wait(timeout=remaining)

        if selected is None:
            raise BridgeError(
                503,
                "NO_ACTIVE_EDITOR",
                "当前文档没有已就绪的 ai-bridge 编辑器页面",
            )
        public_session = bridge_public_session(selected_id, selected)
        identity = selected["identity"]

    binding_claims = {
        "fileName": identity[0],
        "fileType": identity[1],
        "editorType": identity[2],
        "userId": identity[3],
        "authKind": "binding",
    }
    binding_token, expires_at = bridge_binding_token(binding_claims)
    return {
        "ok": True,
        "bindingToken": binding_token,
        "bindingExpiresAt": expires_at,
        "session": public_session,
        **contract_identity(),
    }


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
        **contract_identity(),
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
                candidates.sort(key=lambda item: int(item[1].get("generation") or 0), reverse=True)
                return candidates[0]
        elif candidates:
            candidates.sort(key=lambda item: int(item[1].get("generation") or 0), reverse=True)
            return candidates[0]

        remaining = deadline - time.time()
        if remaining <= 0:
            break
        BRIDGE_CONDITION.wait(timeout=remaining)

    if requested and claims.get("authKind") != "binding":
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


def sheet_runtime_feature(
    capabilities: dict[str, Any],
    *path: str,
) -> bool | None:
    value: Any = capabilities.get("features")
    for name in path:
        if not isinstance(value, dict) or name not in value:
            return None
        value = value[name]
    return value if isinstance(value, bool) else None


def require_sheets_batch_dependencies(tool_calls: list[dict[str, Any]]) -> None:
    created_or_renamed: dict[str, int] = {}
    direct_sheet_fields = {
        "sheet",
        "destinationSheet",
        "beforeSheet",
        "sourceSheet",
    }
    formula_fields = {"formula", "formula1", "formula2", "categoryRange", "valuesRange", "xValuesRange"}

    def references_name(key: str, value: Any, name: str) -> bool:
        if isinstance(value, list):
            return any(references_name(key, item, name) for item in value)
        if isinstance(value, dict):
            return any(references_name(str(nested_key), nested, name) for nested_key, nested in value.items())
        if not isinstance(value, str):
            return False
        if key in direct_sheet_fields and value == name:
            return True
        if key not in formula_fields:
            return False
        escaped = name.replace("'", "''")
        return f"'{escaped}'!" in value or re.search(rf"(?<![A-Za-z0-9_]){re.escape(name)}!", value) is not None

    for index, call in enumerate(tool_calls):
        arguments = call.get("arguments") if isinstance(call.get("arguments"), dict) else {}
        for name, source_index in created_or_renamed.items():
            if any(references_name(str(key), value, name) for key, value in arguments.items()):
                raise BridgeError(
                    409,
                    "BATCH_DEPENDENCY_REQUIRES_SPLIT",
                    "同一批次不能引用刚新增或刚重命名的工作表；请先提交工作表结构，再发起下一批调用",
                    {
                        "sheet": name,
                        "sourceToolCallIndex": source_index,
                        "toolCallIndex": index,
                        "completedToolCalls": 0,
                        "partialMutationPossible": False,
                        "mutationState": "none",
                    },
                )
        name = str(call.get("name") or "")
        created_name: Any = None
        if name == "sheets_add_sheet":
            created_name = arguments.get("name")
        elif name == "sheets_rename_sheet":
            created_name = arguments.get("newName")
        elif name == "sheets_manage_sheet" and arguments.get("action") == "copy":
            created_name = arguments.get("newName")
        if isinstance(created_name, str) and created_name:
            created_or_renamed[created_name] = index


def require_sheets_runtime_capabilities(
    tool_calls: list[dict[str, Any]],
    capabilities: dict[str, Any],
) -> None:
    structured_fields = {
        "sourceType",
        "name",
        "style",
        "showTotals",
        "showHeaders",
        "rowStripes",
        "columnStripes",
        "firstColumn",
        "lastColumn",
        "showAutoFilter",
        "showAutoFilterDropDown",
        "summary",
        "alternativeText",
    }

    def reject(index: int, feature: str, message: str) -> None:
        raise BridgeError(
            501,
            "SHEETS_API_UNSUPPORTED",
            message,
            {
                "feature": feature,
                "retryable": False,
                "mutationState": "none",
                "toolCallIndex": index,
                "completedToolCalls": 0,
                "partialMutationPossible": False,
            },
        )

    for index, call in enumerate(tool_calls):
        name = str(call.get("name", ""))
        arguments = call.get("arguments")
        if not isinstance(arguments, dict):
            arguments = {}
        if name == "sheets_manage_table":
            action = str(arguments.get("action", ""))
            mode = str(arguments.get("tableMode", "auto"))
            if mode == "basic":
                mode = "rangeStyle"
            requires_native = mode == "structured" or bool(
                structured_fields.intersection(arguments)
            )
            if (
                action == "create"
                and requires_native
                and sheet_runtime_feature(
                    capabilities, "sheets", "nativeTables", "create"
                ) is False
            ):
                reject(
                    index,
                    "sheets.nativeTables.create",
                    "当前 ONLYOFFICE 运行时不支持创建原生结构化表格",
                )
            if (
                action == "create"
                and mode == "rangeStyle"
                and sheet_runtime_feature(
                    capabilities, "sheets", "rangeStyleTables", "create"
                ) is False
            ):
                reject(
                    index,
                    "sheets.rangeStyleTables.create",
                    "当前 ONLYOFFICE 运行时不支持普通区域表格样式",
                )
            if (
                action in {"update", "resize", "delete", "unlist"}
                and sheet_runtime_feature(
                    capabilities, "sheets", "nativeTables", "inspect"
                ) is False
            ):
                reject(
                    index,
                    "sheets.nativeTables.inspect",
                    "当前 ONLYOFFICE 运行时不支持访问原生结构化表格",
                )
        elif name == "sheets_manage_conditional_format":
            if sheet_runtime_feature(
                capabilities, "sheets", "conditionalFormatting", "create"
            ) is False:
                reject(
                    index,
                    "sheets.conditionalFormatting.create",
                    "当前 ONLYOFFICE 运行时不支持条件格式",
                )
        elif name == "sheets_manage_range":
            action = str(arguments.get("action", ""))
            fill_feature = {
                "fillDown": "fillDown",
                "fillUp": "fillUp",
                "fillLeft": "fillLeft",
                "fillRight": "fillRight",
            }.get(action)
            if fill_feature and sheet_runtime_feature(
                capabilities, "sheets", "rangeFill", fill_feature
            ) is False:
                reject(
                    index,
                    f"sheets.rangeFill.{fill_feature}",
                    f"当前 ONLYOFFICE 运行时不支持区域操作 {action}",
                )
        elif name == "sheets_set_array_formula":
            if sheet_runtime_feature(
                capabilities, "sheets", "arrayFormula", "set"
            ) is False:
                reject(index, "sheets.arrayFormula.set", "当前 ONLYOFFICE 运行时不支持数组公式")
        elif name == "sheets_manage_validation":
            if sheet_runtime_feature(
                capabilities, "sheets", "validation", "manage"
            ) is False:
                reject(index, "sheets.validation.manage", "当前 ONLYOFFICE 运行时不支持数据验证")
        elif name == "sheets_inspect_comments":
            if sheet_runtime_feature(
                capabilities, "sheets", "comments", "inspect"
            ) is False:
                reject(index, "sheets.comments.inspect", "当前 ONLYOFFICE 运行时不支持读取批注")
        elif name == "sheets_manage_comments":
            action = str(arguments.get("action", ""))
            feature = "create" if action == "add" else ("delete" if action == "delete" else "update")
            if sheet_runtime_feature(
                capabilities, "sheets", "comments", feature
            ) is False:
                reject(index, f"sheets.comments.{feature}", f"当前 ONLYOFFICE 运行时不支持批注操作 {action}")
        elif name == "sheets_inspect_freeze_panes":
            if sheet_runtime_feature(
                capabilities, "sheets", "freezePanes", "inspect"
            ) is False:
                reject(index, "sheets.freezePanes.inspect", "当前 ONLYOFFICE 运行时不支持读取冻结窗格")
        elif name == "sheets_manage_freeze_panes":
            if sheet_runtime_feature(
                capabilities, "sheets", "freezePanes", "manage"
            ) is False:
                reject(index, "sheets.freezePanes.manage", "当前 ONLYOFFICE 运行时不支持管理冻结窗格")
        elif name == "sheets_inspect_charts":
            if sheet_runtime_feature(
                capabilities, "sheets", "charts", "inspect"
            ) is False:
                reject(index, "sheets.charts.inspect", "当前 ONLYOFFICE 运行时不支持读取图表")
        elif name == "sheets_update_chart":
            if sheet_runtime_feature(
                capabilities, "sheets", "charts", "update"
            ) is False:
                reject(index, "sheets.charts.update", "当前 ONLYOFFICE 运行时不支持更新图表")
        elif name == "sheets_add_chart":
            if sheet_runtime_feature(
                capabilities, "sheets", "charts", "create"
            ) is False:
                reject(
                    index,
                    "sheets.charts.create",
                    "当前 ONLYOFFICE 运行时不支持创建图表",
                )
            has_advanced_series = any(
                key in arguments and arguments.get(key) not in (None, [], "")
                for key in ("categoryRange", "addSeries", "seriesUpdates", "removeSeries")
            )
            if (
                has_advanced_series
                and sheet_runtime_feature(
                    capabilities, "sheets", "charts", "addSeriesOnCreate"
                ) is False
            ):
                reject(
                    index,
                    "sheets.charts.addSeriesOnCreate",
                    "当前 ONLYOFFICE 运行时不支持在创建图表时追加或重定向系列",
                )
        elif name == "sheets_delete_chart":
            if sheet_runtime_feature(
                capabilities, "sheets", "charts", "delete"
            ) is False:
                reject(
                    index,
                    "sheets.charts.delete",
                    "当前 ONLYOFFICE 运行时不支持删除图表",
                )


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
    timeout_ms = min(MAX_TIMEOUT_MS, max(100, timeout_ms))

    params: dict[str, Any] = {}
    argument_normalizations: list[dict[str, Any]] = []
    editor_type = str((session.get("state") or {}).get("editorType") or "")
    if editor_type not in ARGUMENT_SCHEMAS_BY_EDITOR:
        raise BridgeError(400, "EDITOR_MISMATCH", "当前编辑器类型不支持公开工具调用")
    allowed_tools = set(((session.get("state") or {}).get("capabilities") or {}).get("tools") or [])
    if method == "executeTool":
        public_name = str(payload.get("name", "")).strip()
        name = internal_tool_name(editor_type, public_name)
        if name not in allowed_tools:
            raise BridgeError(400, "TOOL_NOT_ALLOWED", f"当前编辑器不允许公开工具：{public_name or 'unknown'}")
        arguments = bridge_parse_json_parameter(payload, "arguments", "argumentsJson", dict, {})
        require_arguments_within_limit(
            arguments,
            [{"name": name, "arguments": arguments}],
            "arguments",
        )
        try:
            normalized_calls, argument_normalizations = require_valid_editor_tool_calls(
                editor_type,
                [{"name": name, "arguments": arguments}],
            )
            arguments = normalized_calls[0]["arguments"]
        except BridgeError as error:
            raise project_bridge_error(editor_type, error)
        if editor_type == "cell":
            require_sheets_runtime_capabilities(
                [{"name": name, "arguments": arguments}],
                (session.get("state") or {}).get("capabilities") or {},
            )
        params = {"name": name, "arguments": arguments}
    elif method == "executeBatch":
        public_tool_calls = bridge_parse_json_parameter(payload, "toolCalls", "toolCallsJson", list, [])
        if not public_tool_calls or len(public_tool_calls) > MAX_TOOL_CALLS:
            raise BridgeError(
                400,
                "INVALID_TOOL_CALL",
                f"toolCalls 数量必须为 1 到 {MAX_TOOL_CALLS}",
            )
        tool_calls = internalize_public_tool_calls(editor_type, public_tool_calls)
        for call in tool_calls:
            if not isinstance(call, dict) or call.get("name") not in allowed_tools:
                raise BridgeError(400, "TOOL_NOT_ALLOWED", "批量调用包含当前编辑器不允许的工具")
        require_arguments_within_limit(tool_calls, tool_calls, "toolCalls")
        try:
            tool_calls, argument_normalizations = require_valid_editor_tool_calls(
                editor_type,
                tool_calls,
            )
        except BridgeError as error:
            raise project_bridge_error(editor_type, error)
        if editor_type == "cell":
            require_sheets_batch_dependencies(tool_calls)
            require_sheets_runtime_capabilities(
                tool_calls,
                (session.get("state") or {}).get("capabilities") or {},
            )
        params = {"toolCalls": tool_calls}

    return {
        "requestId": request_id,
        "method": method,
        "params": params,
        "timeoutMs": timeout_ms,
        "argumentNormalizations": argument_normalizations,
    }


def bridge_validate(payload: dict[str, Any], claims: dict[str, Any]) -> dict[str, Any]:
    if claims.get("authKind") != "binding":
        raise BridgeError(
            401,
            "INVALID_BRIDGE_BINDING_TOKEN",
            "批量预检必须使用 attach 返回的 ai-bridge 绑定凭证",
        )
    with BRIDGE_CONDITION:
        session_id, session = bridge_select_session_locked(payload.get("sessionId"), claims)
        editor_type = str((session.get("state") or {}).get("editorType") or "")
        if editor_type not in {"word", "slide", "cell"}:
            raise BridgeError(
                400,
                "EDITOR_MISMATCH",
                "批量预检只支持当前 Word、PPTX 或 XLSX 编辑器会话",
            )
        public_tool_calls = bridge_parse_json_parameter(
            payload,
            "toolCalls",
            "toolCallsJson",
            list,
            [],
        )
        if not public_tool_calls or len(public_tool_calls) > MAX_TOOL_CALLS:
            raise BridgeError(
                400,
                "INVALID_TOOL_CALL",
                f"toolCalls 数量必须为 1 到 {MAX_TOOL_CALLS}",
            )
        allowed_tools = set(
            ((session.get("state") or {}).get("capabilities") or {}).get("tools") or []
        )
        tool_calls = internalize_public_tool_calls(editor_type, public_tool_calls)
        for call in tool_calls:
            if not isinstance(call, dict) or call.get("name") not in allowed_tools:
                raise BridgeError(
                    400,
                    "TOOL_NOT_ALLOWED",
                    "批量调用包含当前编辑器不允许的工具",
                )
        require_arguments_within_limit(tool_calls, tool_calls, "toolCalls")
        try:
            normalized_calls, argument_normalizations = require_valid_editor_tool_calls(
                editor_type,
                tool_calls
            )
        except BridgeError as error:
            raise project_bridge_error(editor_type, error)
        if editor_type == "cell":
            require_sheets_batch_dependencies(normalized_calls)
            require_sheets_runtime_capabilities(
                normalized_calls,
                (session.get("state") or {}).get("capabilities") or {},
            )
        return {
            "ok": True,
            "valid": True,
            "editorType": editor_type,
            "toolCalls": len(normalized_calls),
            "argumentNormalizations": argument_normalizations,
            **contract_identity(),
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
                project_public_tool_fields(
                    str((session.get("state") or {}).get("editorType") or ""),
                    error.get("details"),
                ),
            )
        result_session_id = str(command.get("sessionId") or session_id)
        result_session = (
            BRIDGE_SESSIONS.get(result_session_id)
            or BRIDGE_SUPERSEDED.get(result_session_id)
            or session
        )
        completed_monotonic = command.get("completedMonotonic")
        bridge_timings = {
            "queueWait": bridge_elapsed_ms(
                command.get("createdMonotonic"),
                command.get("firstDeliveredMonotonic"),
            ),
            "editorRoundTrip": bridge_elapsed_ms(
                command.get("firstDeliveredMonotonic"),
                completed_monotonic,
            ),
            "total": bridge_elapsed_ms(
                command.get("createdMonotonic"),
                completed_monotonic,
            ),
        }
        return {
            "ok": True,
            "cached": cached,
            "requestId": command_data["requestId"],
            "session": bridge_public_session(result_session_id, result_session),
            "result": project_public_tool_fields(
                str((result_session.get("state") or {}).get("editorType") or ""),
                response.get("result"),
            ),
            "timingsMs": bridge_timings,
            "argumentNormalizations": command.get("argumentNormalizations", []),
            **contract_identity(),
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
        DOCUMENT_STORAGE_ROOT,
    )


def document_template_path(file_type: str) -> str:
    if file_type not in DOCUMENT_TYPES:
        raise BridgeError(404, "DOCUMENT_TYPE_NOT_FOUND", "不支持的文档类型")
    if file_type == "pptx":
        configured = os.environ.get("DOCUMENT_PPTX_TEMPLATE_PATH", "").strip()
        if configured:
            validate_pptx_template(configured)
            return configured
    template_root = os.environ.get("DOCUMENT_TEMPLATE_ROOT", DOCUMENT_TEMPLATE_ROOT)
    return os.path.join(template_root, f"new.{file_type}")


def validate_pptx_template(path: str) -> None:
    if not os.path.isfile(path):
        raise RuntimeError("找不到 DOCUMENT_PPTX_TEMPLATE_PATH 指定的 PPTX 模板")
    required_parts = {
        "[Content_Types].xml",
        "ppt/presentation.xml",
    }
    try:
        with zipfile.ZipFile(path) as archive:
            names = set(archive.namelist())
            missing = sorted(required_parts - names)
            slide_parts = sorted(
                name
                for name in names
                if re.fullmatch(r"ppt/slides/slide[1-9][0-9]*\.xml", name)
            )
            if missing or not slide_parts:
                details = missing or ["ppt/slides/slide*.xml"]
                raise RuntimeError(
                    "PPTX 模板缺少必要 OOXML 部件：" + ", ".join(details)
                )
            for part in ("[Content_Types].xml", "ppt/presentation.xml", slide_parts[0]):
                ET.fromstring(archive.read(part))
    except (OSError, zipfile.BadZipFile, KeyError, ET.ParseError) as error:
        raise RuntimeError("DOCUMENT_PPTX_TEMPLATE_PATH 不是有效的 PPTX 模板") from error


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


def document_frame_ancestors() -> tuple[str, ...]:
    raw = os.environ.get("DOCUMENT_FRAME_ANCESTORS", "'self'").strip()
    ancestors: list[str] = []
    for token in raw.split():
        if token == "'self'" or token in LOOPBACK_FRAME_ANCESTOR_PATTERNS:
            candidate = token
        else:
            parsed = urllib.parse.urlsplit(token)
            if (
                parsed.scheme not in ("http", "https")
                or not parsed.hostname
                or parsed.username
                or parsed.password
                or parsed.path not in ("", "/")
                or parsed.query
                or parsed.fragment
                or "*" in token
            ):
                raise RuntimeError(
                    "DOCUMENT_FRAME_ANCESTORS 只允许 'self'、受支持的回环来源"
                    "或精确 HTTP(S) origin"
                )
            candidate = token.rstrip("/")
        if candidate not in ancestors:
            ancestors.append(candidate)
    if not ancestors:
        raise RuntimeError("DOCUMENT_FRAME_ANCESTORS 不得为空")
    return tuple(ancestors)


def document_editor_url(file_type: str, document_id: str) -> str:
    return f"{document_public_origin()}/{file_type}/{document_id}"


def document_revision_key(file_type: str, document_id: str) -> str:
    path = document_path(file_type, document_id)
    metadata = os.stat(path, follow_symlinks=False)
    revision = hashlib.sha256(
        (
            f"{document_id}\0{file_type}\0"
            f"{metadata.st_size}\0{metadata.st_mtime_ns}"
        ).encode("utf-8")
    ).hexdigest()[:24]
    return f"{document_id}.{revision}"


def document_editor_config(file_type: str, document_id: str) -> dict[str, Any]:
    document_id = normalize_document_id(document_id)
    file_name = f"{document_id}.{file_type}"
    document_file_identity(file_name)
    document_key = document_revision_key(file_type, document_id)
    internal_base = (
        f"{DOCUMENT_INTERNAL_ORIGIN}/__document_storage"
        f"/download/{file_type}/{document_id}"
    )
    channel_id = f"channel-{uuid.uuid4()}"
    plugin_url = (
        f"{document_public_origin()}/sdkjs-plugins/"
        f"{{A17E5F31-64AA-4E37-9A42-8D430814C2F6}}"
        f"/config.json?v={EDITOR_ASSET_REVISION}"
    )
    config: dict[str, Any] = {
        "document": {
            "fileType": file_type,
            "key": document_key,
            "title": file_name,
            "url": f"{internal_base}?v={document_key.rsplit('.', 1)[-1]}",
            "info": {
                "owner": "Office",
                "uploaded": time.strftime(
                    "%Y-%m-%d %H:%M:%S UTC",
                    time.gmtime(os.path.getmtime(document_path(file_type, document_id))),
                ),
            },
            "permissions": {
                "chat": False,
                "comment": True,
                "copy": True,
                "download": True,
                "edit": True,
                "fillForms": True,
                "modifyContentControl": True,
                "modifyFilter": True,
                "print": True,
                "protect": True,
                "review": True,
            },
        },
        "documentType": DOCUMENT_TYPES[file_type],
        "editorConfig": {
            "callbackUrl": (
                f"{DOCUMENT_INTERNAL_ORIGIN}/__document_storage"
                f"/callback/{file_type}/{document_id}"
            ),
            "customization": {
                "compactToolbar": True,
                "feedback": False,
                "forcesave": False,
                "layout": {
                    "leftMenu": False,
                    "toolbar": {
                        "collaboration": False,
                        "plugins": False,
                    },
                },
            },
            "lang": "zh",
            "mode": "edit",
            "plugins": {
                "pluginsData": [plugin_url],
                "autostart": [AI_BRIDGE_GUID],
                "options": {
                    AI_BRIDGE_GUID: {
                        "hostOrigin": document_public_origin(),
                        "channelId": channel_id,
                    },
                },
            },
            "user": {
                "group": "",
                "id": "pending-local-guest",
                "image": "",
                "name": "访客",
                "roles": [],
            },
        },
        "height": "100%",
        "type": "desktop",
        "width": "100%",
    }
    now = int(time.time())
    token_payload = {
        **config,
        "iat": now,
        "exp": now + EDITOR_TOKEN_TTL_SECONDS,
    }
    config["token"] = sign_jwt(token_payload, get_jwt_secret())
    return config


def document_editor_html(file_type: str, document_id: str) -> str:
    config = document_editor_config(file_type, document_id)
    encoded_config = base64.urlsafe_b64encode(
        compact_json(config).encode("utf-8")
    ).decode("ascii")
    encoded_config = html.escape(encoded_config, quote=True)
    shard_key = urllib.parse.quote(str(config["document"]["key"]), safe="")
    revision = html.escape(EDITOR_ASSET_REVISION, quote=True)
    return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>{html.escape(str(config["document"]["title"]))}</title>
  <style>
    html, body, #iframeEditor {{ width: 100%; height: 100%; margin: 0; overflow: hidden; }}
    body {{ background: #f4f4f4; }}
  </style>
</head>
<body>
  <div id="iframeEditor" data-editor-config="{encoded_config}"></div>
  <script src="/web-apps/apps/api/documents/api.js?shardkey={shard_key}"></script>
  <script src="/sdkjs-plugins/{{A17E5F31-64AA-4E37-9A42-8D430814C2F6}}/local-guest.js?v={revision}"></script>
  <script src="/sdkjs-plugins/{{A17E5F31-64AA-4E37-9A42-8D430814C2F6}}/editor-shell.js?v={revision}"></script>
</body>
</html>"""


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


def document_editor_response(
    handler: BaseHTTPRequestHandler,
    file_type: str,
    document_id: str,
) -> None:
    body = document_editor_html(file_type, document_id).encode("utf-8")
    frame_ancestors = document_frame_ancestors()
    try:
        handler.send_response(200)
        handler.send_header("Content-Type", "text/html; charset=utf-8")
        handler.send_header("Content-Length", str(len(body)))
        handler.send_header("Cache-Control", "no-store")
        handler.send_header("Pragma", "no-cache")
        handler.send_header("X-Content-Type-Options", "nosniff")
        if frame_ancestors == ("'self'",):
            handler.send_header("X-Frame-Options", "SAMEORIGIN")
        handler.send_header("Referrer-Policy", "no-referrer")
        handler.send_header(
            "Content-Security-Policy",
            "default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; "
            "img-src 'self' data: blob:; connect-src 'self' ws: wss:; "
            "frame-src 'self' blob:; worker-src 'self' blob:; "
            "font-src 'self' data:; object-src 'none'; base-uri 'none'; "
            f"form-action 'none'; frame-ancestors {' '.join(frame_ancestors)}",
        )
        handler.end_headers()
        handler.wfile.write(body)
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


def verify_storage_jwt(token: str) -> dict[str, Any]:
    try:
        encoded_header, encoded_payload, encoded_signature = token.split(".")
        unsigned = f"{encoded_header}.{encoded_payload}"
        expected = b64url(
            hmac.new(
                get_jwt_secret().encode(),
                unsigned.encode(),
                hashlib.sha256,
            ).digest()
        )
        if not hmac.compare_digest(encoded_signature, expected):
            raise ValueError("signature")
        header = json.loads(
            base64.urlsafe_b64decode(
                encoded_header + "=" * (-len(encoded_header) % 4)
            )
        )
        payload = json.loads(
            base64.urlsafe_b64decode(
                encoded_payload + "=" * (-len(encoded_payload) % 4)
            )
        )
        if header.get("alg") != "HS256" or not isinstance(payload, dict):
            raise ValueError("claims")
        expiration = payload.get("exp")
        if expiration is not None and float(expiration) < time.time():
            raise ValueError("expired")
        return payload
    except Exception as error:
        raise BridgeError(
            401,
            "INVALID_DOCUMENT_STORAGE_TOKEN",
            "无效的文档存储凭证",
        ) from error


def request_bearer_token(handler: BaseHTTPRequestHandler) -> str:
    authorization = handler.headers.get("Authorization", "")
    if not authorization.startswith("Bearer "):
        raise BridgeError(
            401,
            "DOCUMENT_STORAGE_TOKEN_REQUIRED",
            "缺少文档存储凭证",
        )
    token = authorization.removeprefix("Bearer ").strip()
    if not token:
        raise BridgeError(
            401,
            "DOCUMENT_STORAGE_TOKEN_REQUIRED",
            "缺少文档存储凭证",
        )
    return token


def storage_download_response(
    handler: BaseHTTPRequestHandler,
    file_type: str,
    document_id: str,
) -> None:
    verify_storage_jwt(request_bearer_token(handler))
    document_path(file_type, document_id)
    mime_types = {
        "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    }
    file_name = f"{document_id}.{file_type}"
    try:
        handler.send_response(200)
        handler.send_header(
            "X-Accel-Redirect",
            f"/__document_files/{file_name}",
        )
        handler.send_header("Content-Type", mime_types[file_type])
        handler.send_header(
            "Content-Disposition",
            f'attachment; filename="{file_name}"',
        )
        handler.send_header("Cache-Control", "private, no-store")
        handler.send_header("X-Content-Type-Options", "nosniff")
        handler.end_headers()
    except (BrokenPipeError, ConnectionResetError):
        return


def callback_payload(
    handler: BaseHTTPRequestHandler,
    request_payload: dict[str, Any],
) -> dict[str, Any]:
    body_token = str(request_payload.get("token", "")).strip()
    token = body_token or request_bearer_token(handler)
    signed_payload = verify_storage_jwt(token)
    if "payload" in signed_payload and isinstance(signed_payload["payload"], dict):
        signed_payload = signed_payload["payload"]
    return signed_payload


def validate_callback_key(key: Any, document_id: str) -> str:
    candidate = str(key or "").strip()
    match = DOCUMENT_KEY_PATTERN.fullmatch(candidate)
    if not match or match.group("documentId") != document_id:
        raise BridgeError(
            409,
            "DOCUMENT_CALLBACK_KEY_MISMATCH",
            "保存回调与文档版本不匹配",
        )
    return candidate


def normalized_callback_download_url(value: Any) -> str:
    try:
        parsed = urllib.parse.urlsplit(str(value or "").strip())
    except ValueError as error:
        raise BridgeError(
            400,
            "INVALID_DOCUMENT_CALLBACK_URL",
            "保存回调缺少有效下载地址",
        ) from error
    if (
        parsed.scheme not in ("http", "https")
        or not parsed.hostname
        or parsed.username
        or parsed.password
        or parsed.fragment
        or not parsed.path.startswith("/cache/files/")
    ):
        raise BridgeError(
            400,
            "INVALID_DOCUMENT_CALLBACK_URL",
            "保存回调下载地址不受信任",
        )

    hostname = parsed.hostname.lower().rstrip(".")
    port = parsed.port
    if hostname in ("127.0.0.1", "localhost", "::1"):
        # Document Server builds cache URLs from the browser-facing origin, so a
        # loopback URL may carry Docker's host port (for example 8088 or 11981).
        # The cache endpoint itself lives on this container's port 80.
        return urllib.parse.urlunsplit(
            ("http", "127.0.0.1", parsed.path, parsed.query, "")
        )

    public = urllib.parse.urlsplit(document_public_origin())
    public_port = public.port or (443 if public.scheme == "https" else 80)
    callback_port = port or (443 if parsed.scheme == "https" else 80)
    if hostname != str(public.hostname or "").lower() or callback_port != public_port:
        raise BridgeError(
            400,
            "INVALID_DOCUMENT_CALLBACK_URL",
            "保存回调下载主机不受信任",
        )
    return urllib.parse.urlunsplit(
        ("http", "127.0.0.1", parsed.path, parsed.query, "")
    )


def document_save_lock(path: str) -> threading.Lock:
    with DOCUMENT_SAVE_LOCKS_LOCK:
        return DOCUMENT_SAVE_LOCKS.setdefault(path, threading.Lock())


def document_max_save_bytes() -> int:
    raw = os.environ.get(
        "DOCUMENT_MAX_SAVE_BYTES",
        str(DOCUMENT_MAX_SAVE_BYTES),
    )
    try:
        value = int(raw)
    except ValueError as error:
        raise RuntimeError("DOCUMENT_MAX_SAVE_BYTES 必须是整数") from error
    if value < 1 or value > 2 * 1024 * 1024 * 1024:
        raise RuntimeError("DOCUMENT_MAX_SAVE_BYTES 必须在 1 到 2147483648 之间")
    return value


def persist_callback_document(
    file_type: str,
    document_id: str,
    download_url: Any,
) -> dict[str, Any]:
    destination = document_path(file_type, document_id)
    safe_url = normalized_callback_download_url(download_url)
    maximum = document_max_save_bytes()
    temporary = (
        f"{destination}.saving-"
        f"{os.getpid()}-{threading.get_ident()}-{secrets.token_hex(6)}"
    )
    downloaded = 0
    try:
        with document_save_lock(destination):
            request = urllib.request.Request(
                safe_url,
                headers={"User-Agent": "OnlyOfficeCopilotStorage/0.2.1"},
                method="GET",
            )
            with urllib.request.urlopen(request, timeout=45) as response:
                declared = response.headers.get("Content-Length")
                if declared and int(declared) > maximum:
                    raise BridgeError(
                        413,
                        "DOCUMENT_SAVE_TOO_LARGE",
                        "保存后的文档超过大小限制",
                    )
                with open(temporary, "xb") as output:
                    while True:
                        chunk = response.read(1024 * 1024)
                        if not chunk:
                            break
                        downloaded += len(chunk)
                        if downloaded > maximum:
                            raise BridgeError(
                                413,
                                "DOCUMENT_SAVE_TOO_LARGE",
                                "保存后的文档超过大小限制",
                            )
                        output.write(chunk)
                    output.flush()
                    os.fsync(output.fileno())
            if downloaded < 4:
                raise BridgeError(
                    502,
                    "DOCUMENT_SAVE_EMPTY",
                    "文档服务返回了空文件",
                )
            with open(temporary, "rb") as saved:
                if saved.read(4) != b"PK\x03\x04":
                    raise BridgeError(
                        502,
                        "DOCUMENT_SAVE_INVALID_FORMAT",
                        "文档服务返回的文件格式无效",
                    )
            os.chmod(temporary, 0o644)
            os.replace(temporary, destination)
            os.utime(destination, None)
    except BridgeError:
        try:
            os.unlink(temporary)
        except FileNotFoundError:
            pass
        raise
    except Exception as error:
        try:
            os.unlink(temporary)
        except FileNotFoundError:
            pass
        raise BridgeError(
            502,
            "DOCUMENT_SAVE_DOWNLOAD_FAILED",
            "无法下载并保存编辑后的文档",
        ) from error
    return {
        "persisted": True,
        "bytes": downloaded,
        "fileName": os.path.basename(destination),
    }


def process_document_callback(
    handler: BaseHTTPRequestHandler,
    request_payload: dict[str, Any],
    file_type: str,
    document_id: str,
) -> dict[str, Any]:
    payload = callback_payload(handler, request_payload)
    validate_callback_key(payload.get("key"), document_id)
    status = payload.get("status")
    if not isinstance(status, int) or isinstance(status, bool) or status not in range(1, 8):
        raise BridgeError(
            400,
            "INVALID_DOCUMENT_CALLBACK_STATUS",
            "保存回调状态无效",
        )
    persisted = None
    if status in (2, 6):
        callback_file_type = str(payload.get("filetype", file_type)).lower()
        if callback_file_type != file_type:
            raise BridgeError(
                409,
                "DOCUMENT_CALLBACK_TYPE_MISMATCH",
                "保存回调文件类型不匹配",
            )
        persisted = persist_callback_document(
            file_type,
            document_id,
            payload.get("url"),
        )
    print(
        "[document-callback] "
        + compact_json(
            {
                "documentId": document_id,
                "fileType": file_type,
                "status": status,
                "persisted": bool(persisted),
                "bytes": (persisted or {}).get("bytes", 0),
            }
        ),
        flush=True,
    )
    return {"error": 0}


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


def canonical_file(file_name: str | None) -> str:
    if not file_name:
        raise ValueError("未取得文件名")
    safe_name = os.path.basename(file_name)
    if safe_name != file_name or not DOCUMENT_FILE_PATTERN.fullmatch(safe_name):
        raise ValueError("无效的 UUID 文档文件名")
    path = os.path.join(document_storage_directory(), safe_name)
    if not os.path.isfile(path):
        raise ValueError(f"找不到原文件：{safe_name}")
    return path


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
    # The storage callback can perform one final write after the user is dropped.
    # Let that atomic write settle before replacing the canonical file.
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


def force_save(key: str, file_name: str | None, allow_no_changes: bool = False) -> dict[str, Any]:
    if not key or len(key) > 512:
        raise ValueError("无效 document.key")
    main_file = canonical_file(file_name)
    before = os.path.getmtime(main_file)
    command: dict[str, Any] = {"c": "forcesave", "key": key}
    if file_name:
        command["userdata"] = compact_json({"fileName": os.path.basename(file_name), "source": "office-copilot"})
    retry_deadline = time.time() + 12
    while True:
        result = command_service(command)
        if result.get("error") != 4 or allow_no_changes or time.time() >= retry_deadline:
            break
        time.sleep(0.6)
    if result.get("error") == 4 and allow_no_changes:
        return {
            "accepted": False,
            "noChanges": True,
            "persisted": os.path.isfile(main_file),
            "status": "no_changes",
            "commandError": 4,
            "command": result,
            "beforeMtime": before,
            "afterMtime": os.path.getmtime(main_file),
            "note": "No editor changes needed saving; the current canonical file remains downloadable.",
        }
    if result.get("error") != 0:
        return {
            "accepted": False,
            "noChanges": False,
            "persisted": False,
            "status": "failed",
            "commandError": result.get("error"),
            "command": result,
            "beforeMtime": before,
            "afterMtime": os.path.getmtime(main_file),
        }

    persisted = False
    after = before
    deadline = time.time() + 15
    while time.time() < deadline:
        after = os.path.getmtime(main_file)
        if after > before + 0.0001:
            persisted = True
            break
        time.sleep(0.35)
    return {
        "accepted": True,
        "persisted": persisted,
        "status": "saved" if persisted else "failed",
        "commandError": 0,
        "command": result,
        "beforeMtime": before,
        "afterMtime": after,
    }


class Handler(BaseHTTPRequestHandler):
    server_version = "OnlyOfficeCopilot/2.0"

    def do_GET(self) -> None:
        parsed_path = urllib.parse.urlsplit(self.path)
        request_path = parsed_path.path.rstrip("/") or "/"
        if not v1_relay_route_allowed("GET", request_path):
            json_response(
                self,
                410,
                {
                    "code": "LEGACY_BUSINESS_DISABLED",
                    "error": "旧文档业务接口已迁移到 document-hub",
                },
            )
            return
        try:
            require_internal_relay(self)
        except BridgeError as error:
            self.log_bridge_error(error)
            json_response(self, error.status, error.payload())
            return
        storage_request = DOCUMENT_STORAGE_PATH_PATTERN.fullmatch(
            parsed_path.path
        )
        if storage_request:
            if storage_request.group("action") != "download":
                json_response(
                    self,
                    405,
                    {"error": "Method not allowed"},
                    {"Allow": "POST"},
                )
                return
            try:
                storage_download_response(
                    self,
                    storage_request.group("fileType"),
                    storage_request.group("documentId"),
                )
            except BridgeError as error:
                json_response(self, error.status, error.payload())
            return
        public_document = PUBLIC_DOCUMENT_PATH_PATTERN.fullmatch(parsed_path.path)
        if public_document:
            try:
                document_editor_response(
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
                document_editor_response(self, file_type, document_id)
            except BridgeError as error:
                json_response(self, error.status, error.payload())
            return
        if request_path == "/health":
            json_response(
                self,
                200,
                {
                    "ok": True,
                    "modelConfigured": bool(
                        os.environ.get("COPILOT_API_KEY")
                        and os.environ.get("COPILOT_MODEL")
                    ),
                    "editors": ["word", "slide", "cell"],
                    **contract_identity(),
                },
            )
            return
        contract_match = re.fullmatch(
            r"/bridge/contract/(?P<editor>word|slide|cell)",
            request_path,
        )
        if contract_match:
            json_response(
                self,
                200,
                editor_public_contract(contract_match.group("editor")),
            )
            return
        if request_path.startswith("/images/"):
            try:
                require_internal_relay(self)
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
        json_response(self, 404, {"error": "Not found"})

    def do_POST(self) -> None:
        payload: dict[str, Any] = {}
        request_path = ""
        try:
            parsed_path = urllib.parse.urlsplit(self.path)
            request_path = parsed_path.path.rstrip("/") or "/"
            if not v1_relay_route_allowed("POST", request_path):
                raise BridgeError(
                    410,
                    "LEGACY_BUSINESS_DISABLED",
                    "旧文档业务接口已迁移到 document-hub",
                )
            storage_request = DOCUMENT_STORAGE_PATH_PATTERN.fullmatch(
                parsed_path.path
            )
            if storage_request:
                request_path = parsed_path.path
                if storage_request.group("action") != "callback":
                    json_response(
                        self,
                        405,
                        {"error": "Method not allowed"},
                        {"Allow": "GET"},
                    )
                    return
                content_type = (
                    self.headers.get("Content-Type", "")
                    .split(";", 1)[0]
                    .strip()
                    .lower()
                )
                if content_type != "application/json":
                    raise BridgeError(
                        415,
                        "UNSUPPORTED_MEDIA_TYPE",
                        "Content-Type 必须是 application/json",
                    )
                payload = read_json(self)
                json_response(
                    self,
                    200,
                    process_document_callback(
                        self,
                        payload,
                        storage_request.group("fileType"),
                        storage_request.group("documentId"),
                    ),
                )
                return
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
                if request_path
                in (
                    "/bridge/internal/execute",
                    "/bridge/internal/validate",
                    "/bridge/internal/images/import",
                )
                else 2_000_000
            )
            content_type = (
                self.headers.get("Content-Type", "")
                .split(";", 1)[0]
                .strip()
                .lower()
            )
            payload = read_json(self, max_bytes)
            if request_path == "/bridge/register":
                require_internal_relay(self)
                json_response(self, 200, bridge_register(payload))
                return
            if request_path == "/bridge/poll":
                require_internal_relay(self)
                json_response(self, 200, bridge_poll(payload))
                return
            if request_path == "/bridge/result":
                require_internal_relay(self)
                json_response(self, 200, bridge_result(payload))
                return
            if request_path == "/bridge/unregister":
                require_internal_relay(self)
                json_response(self, 200, bridge_unregister(payload))
                return
            if request_path == "/bridge/internal/execute":
                require_internal_relay(self)
                claims = internal_bridge_claims(payload.get("relaySessionId"))
                payload["sessionId"] = payload.get("relaySessionId")
                json_response(self, 200, bridge_execute(payload, claims))
                return
            if request_path == "/bridge/internal/validate":
                require_internal_relay(self)
                claims = internal_bridge_claims(payload.get("relaySessionId"))
                payload["sessionId"] = payload.get("relaySessionId")
                json_response(self, 200, bridge_validate(payload, claims))
                return
            if request_path == "/bridge/internal/session":
                require_internal_relay(self)
                json_response(self, 200, internal_bridge_session(payload.get("relaySessionId")))
                return
            if request_path == "/bridge/internal/drop":
                require_internal_relay(self)
                json_response(self, 200, internal_bridge_drop(payload.get("relaySessionId")))
                return
            if request_path == "/bridge/internal/images/import":
                require_internal_relay(self)
                claims = internal_bridge_claims(payload.get("relaySessionId"))
                json_response(self, 200, import_image(payload.get("source"), claims))
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
            if request_path.startswith("/documents/storage/callback/"):
                json_response(self, error.status, {"error": 1})
            else:
                json_response(self, error.status, error.payload())
        except ValueError as error:
            if request_path.startswith("/documents/storage/callback/"):
                json_response(self, 400, {"error": 1})
            else:
                json_response(self, 400, {"error": str(error)})
        except Exception as error:  # noqa: BLE001 - boundary must return JSON
            if request_path.startswith("/documents/storage/callback/"):
                json_response(self, 502, {"error": 1})
            elif request_path == "/images/import":
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
        identity_field = str(details.get("identityField", "")).strip()
        if identity_field in {"documentId", "fileName", "fileType", "userId"}:
            event["identityField"] = identity_field
        for hash_field in ("claimedIdentityHash", "contextIdentityHash"):
            identity_hash = str(details.get(hash_field, "")).strip()
            if re.fullmatch(r"[0-9a-f]{12}", identity_hash):
                event[hash_field] = identity_hash
        print(
            "[bridge-error] "
            + compact_json(event),
            flush=True,
        )


if __name__ == "__main__":
    if COPILOT_BIND_ADDRESS not in {"127.0.0.1", "0.0.0.0"}:
        raise RuntimeError("COPILOT_BIND_ADDRESS must be 127.0.0.1 or 0.0.0.0")
    if COPILOT_BIND_ADDRESS == "0.0.0.0" and not AI_RELAY_INTERNAL_SECRET:
        raise RuntimeError("AI_RELAY_INTERNAL_SECRET(_FILE) is required for private-network Relay")
    server = ThreadingHTTPServer((COPILOT_BIND_ADDRESS, PORT), Handler)
    print(f"ONLYOFFICE Copilot API listening on {COPILOT_BIND_ADDRESS}:{PORT}", flush=True)
    server.serve_forever()
