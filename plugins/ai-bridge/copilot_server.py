#!/usr/bin/env python3
"""Private Relay, validation, and image service for ONLYOFFICE ai-bridge."""

from __future__ import annotations

import base64
import binascii
import hashlib
import hmac
import ipaddress
import json
import math
import os
import re
import secrets
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
COPILOT_BIND_ADDRESS = os.environ.get("COPILOT_BIND_ADDRESS", "127.0.0.1")
LOCAL_CONFIG = "/etc/onlyoffice/documentserver/local.json"
DOCUMENT_STORAGE_ROOT = "/var/lib/onlyoffice/copilot/documents"
BRIDGE_SESSION_TTL_SECONDS = 180
BRIDGE_RESUME_TOKEN_TTL_SECONDS = 12 * 60 * 60
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
# Keep the documented planning limit stable while allowing small counting
# errors from generated batches. This value is intentionally not projected
# into the public contract.
MAX_ACCEPTED_TOOL_CALLS = 30
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
if MAX_ACCEPTED_TOOL_CALLS < MAX_TOOL_CALLS:
    raise RuntimeError("内部批处理容错上限不得低于公开规划上限")
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


def load_contract_argument_schemas(editor: str) -> dict[str, dict[str, Any]]:
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
    return {
        name: resolve_contract_schema(schema, definitions)
        for name, schema in schemas.items()
    }


ARGUMENT_SCHEMAS_BY_EDITOR = {
    editor: load_contract_argument_schemas(editor)
    for editor in ("word", "slide", "cell")
}
ALLOWED_BY_EDITOR = {
    editor: set(schemas)
    for editor, schemas in ARGUMENT_SCHEMAS_BY_EDITOR.items()
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
TOOL_SCHEMAS = {
    name: schema
    for schemas in ARGUMENT_SCHEMAS_BY_EDITOR.values()
    for name, schema in schemas.items()
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
        if editor == "word" and name == "word_set_numbering":
            levels = call["arguments"].get("levels")
            if not isinstance(levels, list):
                continue
            for level_index, level in enumerate(levels):
                if not isinstance(level, dict):
                    continue
                level_path = f"arguments.levels[{level_index}]"
                if "level" not in level:
                    level["level"] = level_index
                    normalizations.append(
                        {
                            "toolCallIndex": index,
                            "path": f"{level_path}.level",
                            "canonicalPath": f"{level_path}.level",
                            "kind": "implicitDefault",
                        }
                    )
                if "numberFormat" in level:
                    canonical_wins = "format" in level
                    if not canonical_wins:
                        level["format"] = level["numberFormat"]
                    del level["numberFormat"]
                    normalizations.append(
                        {
                            "toolCallIndex": index,
                            "path": f"{level_path}.numberFormat",
                            "canonicalPath": f"{level_path}.format",
                            "kind": "canonicalWins" if canonical_wins else "propertyAlias",
                        }
                    )
                if "hangingIndentPt" in level:
                    canonical_wins = "firstLineIndentPt" in level
                    hanging_indent = level["hangingIndentPt"]
                    if canonical_wins:
                        del level["hangingIndentPt"]
                    elif (
                        isinstance(hanging_indent, (int, float))
                        and not isinstance(hanging_indent, bool)
                        and math.isfinite(hanging_indent)
                    ):
                        level["firstLineIndentPt"] = -abs(hanging_indent)
                        del level["hangingIndentPt"]
                    else:
                        continue
                    normalizations.append(
                        {
                            "toolCallIndex": index,
                            "path": f"{level_path}.hangingIndentPt",
                            "canonicalPath": f"{level_path}.firstLineIndentPt",
                            "kind": "canonicalWins" if canonical_wins else "propertyAlias",
                        }
                    )
                if type(level.get("restart")) is int and level["restart"] in (0, 1):
                    level["restart"] = bool(level["restart"])
                    normalizations.append(
                        {
                            "toolCallIndex": index,
                            "path": f"{level_path}.restart",
                            "canonicalPath": f"{level_path}.restart",
                            "kind": "typeAlias",
                        }
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

    if name == "word_set_numbering":
        assignments = arguments.get("assignments")
        has_legacy_target = (
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
        if (not isinstance(assignments, list) or not assignments) and not has_legacy_target:
            semantic_error(
                "target",
                "set_numbering requires assignments[] or one legacy top-level paragraph target; use one assignments[] call for a continuous multilevel list",
            )
        if isinstance(assignments, list):
            for index, assignment in enumerate(assignments):
                if not isinstance(assignment, dict):
                    continue
                if "paragraphIndexes" in assignment:
                    semantic_error(
                        f"assignments[{index}].paragraphIndexes",
                        "paragraphIndexes is not allowed inside an assignment; split it into one assignment per paragraph in the same call",
                    )
                target_count = sum(
                    assignment.get(field) not in (None, "")
                    for field in (
                        "paragraphId",
                        "internalId",
                        "paragraphIndex",
                        "search",
                    )
                )
                if target_count != 1:
                    semantic_error(
                        f"assignments[{index}]",
                        "each assignment must contain exactly one paragraph locator",
                    )
        levels = arguments.get("levels")
        if isinstance(levels, list):
            seen_levels: set[Any] = set()
            for index, level in enumerate(levels):
                if not isinstance(level, dict):
                    continue
                configured_level = level.get("level", index)
                if configured_level in seen_levels:
                    semantic_error(
                        f"levels[{index}].level",
                        "each numbering level may be configured only once",
                    )
                seen_levels.add(configured_level)


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

    if name == "sheets_filter":
        action = arguments.get("action")
        set_only_fields = {
            "range",
            "field",
            "criteria1",
            "operator",
            "criteria2",
            "visibleDropDown",
        }
        if action == "set":
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
            if target_shape is not None:
                rows, columns = target_shape
                if rows < 2:
                    semantic_error(
                        "range",
                        "filter range must contain a header row and at least one data row",
                    )
                field = arguments.get("field")
                if (
                    isinstance(field, int)
                    and not isinstance(field, bool)
                    and field > columns
                ):
                    semantic_error(
                        "field",
                        f"field must not exceed the filter range width ({columns})",
                    )

            criteria1_present = "criteria1" in arguments
            criteria2_present = "criteria2" in arguments
            criteria1 = arguments.get("criteria1")
            criteria2 = arguments.get("criteria2")
            operator = {
                "and": "xlAnd",
                "or": "xlOr",
                "filterValues": "xlFilterValues",
                "values": "xlFilterValues",
                "top10Items": "xlTop10Items",
                "bottom10Items": "xlBottom10Items",
                "top10Percent": "xlTop10Percent",
                "bottom10Percent": "xlBottom10Percent",
                "filterCellColor": "xlFilterCellColor",
                "filterFontColor": "xlFilterFontColor",
                "filterIcon": "xlFilterIcon",
                "dynamic": "xlFilterDynamic",
            }.get(arguments.get("operator"), arguments.get("operator"))

            def safe_filter_scalar(value: Any) -> bool:
                return (
                    isinstance(value, (str, bool))
                    or (
                        isinstance(value, (int, float))
                        and not isinstance(value, bool)
                        and math.isfinite(value)
                    )
                )

            def safe_filter_criterion(value: Any, allow_array: bool) -> bool:
                if safe_filter_scalar(value):
                    return True
                return bool(
                    allow_array
                    and isinstance(value, list)
                    and value
                    and all(safe_filter_scalar(item) for item in value)
                )

            if criteria1_present and not safe_filter_criterion(criteria1, True):
                semantic_error(
                    "criteria1",
                    "criteria1 must be a scalar or a non-empty array of scalar values",
                )
            if criteria2_present and not safe_filter_criterion(criteria2, False):
                semantic_error("criteria2", "criteria2 must be a scalar value")
            if operator is not None and not criteria1_present:
                semantic_error("criteria1", "criteria1 is required when operator is provided")
            if criteria2_present and operator not in {"xlAnd", "xlOr"}:
                semantic_error(
                    "criteria2",
                    "criteria2 is only allowed with xlAnd or xlOr",
                )
            if criteria2_present and not criteria1_present:
                semantic_error("criteria1", "criteria1 is required when criteria2 is provided")
            if operator == "xlFilterValues" and not (
                isinstance(criteria1, list) and criteria1
            ):
                semantic_error(
                    "criteria1",
                    "xlFilterValues requires a non-empty criteria1 array",
                )
            if isinstance(criteria1, list) and operator != "xlFilterValues":
                semantic_error(
                    "operator",
                    "array criteria1 is only allowed with xlFilterValues",
                )
            if operator in {
                "xlTop10Items",
                "xlBottom10Items",
                "xlTop10Percent",
                "xlBottom10Percent",
            }:
                try:
                    ranking = float(criteria1)
                except (TypeError, ValueError):
                    ranking = math.nan
                if not math.isfinite(ranking) or ranking <= 0:
                    semantic_error(
                        "criteria1",
                        "top/bottom filters require a positive numeric criteria1",
                    )
            if operator in {
                "xlFilterCellColor",
                "xlFilterFontColor",
                "xlFilterIcon",
            }:
                semantic_error(
                    "operator",
                    "color and icon filters require runtime objects and are not safe through the JSON bridge",
                )
        elif action in {"showAll", "reapply"}:
            invalid_fields = sorted(set_only_fields.intersection(arguments))
            if invalid_fields:
                semantic_error(
                    invalid_fields[0],
                    f"{invalid_fields[0]} is only allowed when action is set",
                )

SEMANTIC_VALIDATORS = {
    "word.fieldInstruction": validate_word_semantics,
    "word.pageBreakTarget": validate_word_semantics,
    "word.tableRangeOrder": validate_word_semantics,
    "word.headerFooterContent": validate_word_semantics,
    "word.numberingDefinition": validate_word_semantics,
    "slides.distinctOverlapPair": validate_slide_semantics,
    "sheets.a1MatrixShape": validate_sheet_semantics,
    "sheets.validationRule": validate_sheet_semantics,
    "sheets.tableOperation": validate_sheet_semantics,
    "sheets.filterOperation": validate_sheet_semantics,
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
    "INVALID_BRIDGE_RESUME_TOKEN": 401,
    "BRIDGE_RESUME_TOKEN_EXPIRED": 401,
    "INVALID_RELAY_SESSION": 401,
    "CONTROL_NOT_ALLOWED": 403,
    "IMAGE_FETCH_BLOCKED": 403,
    "COMMAND_NOT_FOUND": 404,
    "IMAGE_ASSET_EXPIRED": 410,
    "DOCUMENT_MISMATCH": 409,
    "DOCUMENT_IDENTITY_MISMATCH": 409,
    "EDITOR_MISMATCH": 409,
    "SESSION_ID_CONFLICT": 409,
    "SESSION_SUPERSEDED": 409,
    "REQUEST_ID_CONFLICT": 409,
    "REQUEST_IN_FLIGHT": 409,
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
    "PERSISTENCE_INVALID_RESPONSE": 502,
    "IMAGE_FETCH_FAILED": 502,
    "HTTP_RELAY_INVALID_RESPONSE": 502,
    "NOT_READY": 503,
    "DOCUMENT_NOT_CONFIGURED": 503,
    "PERSISTENCE_NOT_AVAILABLE": 503,
    "NO_ACTIVE_EDITOR": 503,
    "TIMEOUT": 504,
    "CONNECTION_TIMEOUT": 504,
    "BRIDGE_TIMEOUT": 504,
    "PERSISTENCE_TIMEOUT": 504,
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
        return bridge_session_claims(session)


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
    return session.get("identity") == bridge_identity(claims)


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
        if not public_tool_calls or len(public_tool_calls) > MAX_ACCEPTED_TOOL_CALLS:
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
        if not public_tool_calls or len(public_tool_calls) > MAX_ACCEPTED_TOOL_CALLS:
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
                    {
                        "requestId": command_data["requestId"],
                        "sessionId": session_id,
                        "retryable": False,
                        "reuseAllowed": False,
                        "requiredAction": "use_new_request_id",
                        "partialMutationPossible": False,
                    },
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
        request_path = parsed_path.path.rstrip("/")
        if request_path == "/health":
            json_response(
                self,
                200,
                {
                    "ok": True,
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
            require_internal_relay(self)
            payload = read_json(self, max_bytes)
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
            if request_path == "/bridge/internal/execute":
                claims = internal_bridge_claims(payload.get("relaySessionId"))
                payload["sessionId"] = payload.get("relaySessionId")
                json_response(self, 200, bridge_execute(payload, claims))
                return
            if request_path == "/bridge/internal/validate":
                claims = internal_bridge_claims(payload.get("relaySessionId"))
                payload["sessionId"] = payload.get("relaySessionId")
                json_response(self, 200, bridge_validate(payload, claims))
                return
            if request_path == "/bridge/internal/session":
                json_response(self, 200, internal_bridge_session(payload.get("relaySessionId")))
                return
            if request_path == "/bridge/internal/drop":
                json_response(self, 200, internal_bridge_drop(payload.get("relaySessionId")))
                return
            if request_path == "/bridge/internal/images/import":
                claims = internal_bridge_claims(payload.get("relaySessionId"))
                json_response(self, 200, import_image(payload.get("source"), claims))
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
