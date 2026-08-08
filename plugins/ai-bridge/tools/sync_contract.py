#!/usr/bin/env python3
"""Generate every ai-bridge contract projection from public-api.json.

The script intentionally uses only the Python standard library so both
repositories can run it in CI without a project-specific environment.
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
import re
import sys
from pathlib import Path
from typing import Any
from urllib.parse import urlsplit, urlunsplit


BRIDGE_ROOT = Path(__file__).resolve().parents[1]
CONTRACT_PATH = BRIDGE_ROOT / "public-api.json"
TYPES_PATH = BRIDGE_ROOT / "public-api.d.ts"
PLUGIN_PATH = BRIDGE_ROOT / "plugin.js"
HOST_PATH = BRIDGE_ROOT / "host-bridge.js"
CLIENT_PATH = BRIDGE_ROOT / "client-sdk.js"
CONFIG_PATH = BRIDGE_ROOT / "config.json"
INDEX_PATH = BRIDGE_ROOT / "index.html"
TYPE_START = "// <ai-bridge-generated:tool-arguments>"
TYPE_END = "// </ai-bridge-generated:tool-arguments>"
PLUGIN_START = "  // <ai-bridge-generated:contract-runtime>"
PLUGIN_END = "  // </ai-bridge-generated:contract-runtime>"
EDITOR_CONFIG = {
    "word": {
        "skill": "online-docx",
        "toml": "online-docx-bridge.toml",
        "site": "online-docx-bridge",
        "fileType": "docx",
        "prefix": "word_",
        "label": "Word",
        "sessionAction": "word_session",
        "unitKey": "word",
        "normalizationKey": "word",
    },
    "slide": {
        "skill": "online-pptx",
        "toml": "online-pptx-bridge.toml",
        "site": "online-pptx-bridge",
        "fileType": "pptx",
        "prefix": "slides_",
        "label": "Slides",
        "sessionAction": "slides_session",
        "unitKey": "slide",
        "normalizationKey": "slide",
    },
    "cell": {
        "skill": "online-xlsx",
        "toml": "online-xlsx-bridge.toml",
        "site": "online-xlsx-bridge",
        "fileType": "xlsx",
        "prefix": "sheets_",
        "label": "Sheets",
        "sessionAction": "sheets_session",
        "unitKey": "sheet",
        "normalizationKey": "sheet",
    },
}


def render_skill_version(source: str, version: str) -> str:
    """Update only metadata.version in a SKILL.md YAML frontmatter."""
    if not source.startswith("---\n"):
        raise ValueError("SKILL.md must start with YAML frontmatter")
    end = source.find("\n---", 4)
    if end < 0:
        raise ValueError("SKILL.md YAML frontmatter is not closed")

    frontmatter = source[4:end]
    lines = frontmatter.split("\n")
    metadata_index = next(
        (index for index, line in enumerate(lines) if line == "metadata:"),
        None,
    )
    rendered_version = f'  version: "{version}"'
    if metadata_index is None:
        lines.extend(["metadata:", rendered_version])
    else:
        block_end = len(lines)
        for index in range(metadata_index + 1, len(lines)):
            line = lines[index]
            if line and not line[0].isspace():
                block_end = index
                break
        version_index = next(
            (
                index
                for index in range(metadata_index + 1, block_end)
                if re.match(r"^  version\s*:", lines[index])
            ),
            None,
        )
        if version_index is None:
            lines.insert(metadata_index + 1, rendered_version)
        else:
            lines[version_index] = rendered_version
    return "---\n" + "\n".join(lines) + source[end:]


def canonical_json(value: Any) -> str:
    return json.dumps(
        value,
        ensure_ascii=False,
        separators=(",", ":"),
        sort_keys=True,
    )


def contract_sha256(contract: dict[str, Any]) -> str:
    return hashlib.sha256(canonical_json(contract).encode("utf-8")).hexdigest()


def json_literal(value: Any) -> str:
    if value is None:
        return "null"
    if value is True:
        return "true"
    if value is False:
        return "false"
    if isinstance(value, str):
        return json.dumps(value, ensure_ascii=False)
    return str(value)


def resolve_schema(schema: Any, definitions: dict[str, Any]) -> Any:
    if isinstance(schema, list):
        return [resolve_schema(item, definitions) for item in schema]
    if not isinstance(schema, dict):
        return schema
    reference = schema.get("$ref")
    if isinstance(reference, str) and reference.startswith("#/$defs/"):
        name = reference.removeprefix("#/$defs/")
        if name not in definitions:
            raise ValueError(f"unresolved contract reference: {reference}")
        merged = copy.deepcopy(definitions[name])
        merged.update({key: value for key, value in schema.items() if key != "$ref"})
        return resolve_schema(merged, definitions)
    return {
        key: resolve_schema(value, definitions)
        for key, value in schema.items()
    }


PLUGIN_SCHEMA_KEYWORDS = {
    "additionalProperties",
    "allOf",
    "anyOf",
    "const",
    "else",
    "enum",
    "exclusiveMaximum",
    "exclusiveMinimum",
    "if",
    "items",
    "maxItems",
    "maxLength",
    "maxProperties",
    "maximum",
    "minItems",
    "minLength",
    "minProperties",
    "minimum",
    "not",
    "oneOf",
    "pattern",
    "properties",
    "required",
    "then",
    "type",
    "uniqueItems",
}


def compact_plugin_schema(value: Any) -> Any:
    """Keep only the keywords enforced by the generated plugin validator."""
    if isinstance(value, list):
        return [compact_plugin_schema(item) for item in value]
    if not isinstance(value, dict):
        return value
    compacted = {}
    for key, child in value.items():
        if key not in PLUGIN_SCHEMA_KEYWORDS:
            continue
        if key == "properties" and isinstance(child, dict):
            compacted[key] = {
                name: compact_plugin_schema(property_schema)
                for name, property_schema in child.items()
            }
        else:
            compacted[key] = compact_plugin_schema(child)
    return compacted


def indent(value: str, amount: int = 2) -> str:
    prefix = " " * amount
    return "\n".join(prefix + line if line else line for line in value.splitlines())


def ts_property_name(name: str) -> str:
    return name if re.fullmatch(r"[A-Za-z_$][A-Za-z0-9_$]*", name) else json.dumps(name)


def ts_core_type(schema: dict[str, Any], definitions: dict[str, Any]) -> str:
    if "const" in schema:
        return json_literal(schema["const"])
    enum = schema.get("enum")
    if isinstance(enum, list) and enum:
        return " | ".join(json_literal(value) for value in enum)

    schema_type = schema.get("type")
    if isinstance(schema_type, list):
        return " | ".join(
            ts_core_type({**schema, "type": item}, definitions)
            for item in schema_type
        )
    if schema_type == "string":
        return "string"
    if schema_type in {"integer", "number"}:
        return "number"
    if schema_type == "boolean":
        return "boolean"
    if schema_type == "null":
        return "null"
    if schema_type == "array":
        item_type = ts_type(schema.get("items", {}), definitions)
        return f"Array<{item_type}>"
    if schema_type == "object" or isinstance(schema.get("properties"), dict):
        properties = schema.get("properties") or {}
        required = set(schema.get("required") or [])
        lines = []
        for name, child_schema in properties.items():
            optional = "" if name in required else "?"
            lines.append(
                f"{ts_property_name(name)}{optional}: "
                f"{ts_type(child_schema, definitions)};"
            )
        additional = schema.get("additionalProperties")
        if isinstance(additional, dict):
            lines.append(f"[key: string]: {ts_type(additional, definitions)};")
        elif additional is True and not properties:
            return "Record<string, unknown>"
        if not lines:
            return "Record<string, never>" if additional is False else "Record<string, unknown>"
        return "{\n" + indent("\n".join(lines)) + "\n}"
    return "unknown"


def required_projection(
    required: list[Any],
    base_schema: dict[str, Any],
    definitions: dict[str, Any],
) -> str:
    properties = base_schema.get("properties") or {}
    fields = []
    for name in required:
        if not isinstance(name, str):
            continue
        field_type = ts_type(properties.get(name, {}), definitions)
        fields.append(f"{ts_property_name(name)}: {field_type};")
    return "{ " + " ".join(fields) + " }"


def conditional_projection(
    conditional: dict[str, Any],
    base_schema: dict[str, Any],
    definitions: dict[str, Any],
) -> str:
    condition = conditional.get("if") or {}
    consequence = conditional.get("then") or {}
    then_required = list(consequence.get("required") or [])
    condition_required = [
        item for item in condition.get("required") or [] if isinstance(item, str)
    ]
    condition_properties = condition.get("properties") or {}
    positive_fields = []
    negative_fields = []
    base_properties = base_schema.get("properties") or {}

    for name in condition_required:
        property_condition = condition_properties.get(name) or {}
        property_type = ts_type(base_properties.get(name, {}), definitions)
        if "const" in property_condition:
            literal = json_literal(property_condition["const"])
            positive_fields.append(f"{ts_property_name(name)}: {literal};")
            negative_fields.append(
                f"{ts_property_name(name)}?: Exclude<{property_type}, {literal}>;"
            )
        else:
            positive_fields.append(f"{ts_property_name(name)}: {property_type};")
            negative_fields.append(f"{ts_property_name(name)}?: never;")
    for name in then_required:
        property_type = ts_type(base_properties.get(name, {}), definitions)
        positive_fields.append(f"{ts_property_name(name)}: {property_type};")
    if not positive_fields:
        return "unknown"
    positive = "{ " + " ".join(positive_fields) + " }"
    negative = "{ " + " ".join(negative_fields) + " }"
    return f"({positive} | {negative})"


def ts_type(schema: Any, definitions: dict[str, Any]) -> str:
    if not isinstance(schema, dict):
        return "unknown"
    resolved = resolve_schema(schema, definitions)
    core = ts_core_type(resolved, definitions)
    intersections = [core]

    for keyword in ("anyOf", "oneOf"):
        branches = resolved.get(keyword)
        if isinstance(branches, list) and branches:
            branch_types = []
            for branch in branches:
                if (
                    isinstance(branch, dict)
                    and set(branch).issubset({"required", "not"})
                ):
                    projected = required_projection(
                        list(branch.get("required") or []),
                        resolved,
                        definitions,
                    )
                    negated_required = (
                        (branch.get("not") or {}).get("required")
                        if isinstance(branch.get("not"), dict)
                        else None
                    )
                    if isinstance(negated_required, list):
                        never_fields = " ".join(
                            f"{ts_property_name(name)}?: never;"
                            for name in negated_required
                            if isinstance(name, str)
                        )
                        projected = f"{projected} & {{ {never_fields} }}"
                    branch_types.append(projected)
                else:
                    branch_types.append(ts_type(branch, definitions))
            intersections.append("(" + " | ".join(branch_types) + ")")

    for branch in resolved.get("allOf") or []:
        if not isinstance(branch, dict):
            continue
        if isinstance(branch.get("if"), dict):
            intersections.append(
                conditional_projection(branch, resolved, definitions)
            )
        elif isinstance(branch.get("anyOf"), list):
            projections = [
                required_projection(
                    list(item.get("required") or []),
                    resolved,
                    definitions,
                )
                for item in branch["anyOf"]
                if isinstance(item, dict)
            ]
            if projections:
                intersections.append("(" + " | ".join(projections) + ")")
        else:
            intersections.append(ts_type(branch, definitions))
    return " & ".join(item for item in intersections if item != "unknown") or "unknown"


def generated_type_name(tool_name: str) -> str:
    return "AiBridgeGenerated" + "".join(
        part[:1].upper() + part[1:] for part in tool_name.split("_")
    ) + "Args"


def render_types_region(contract: dict[str, Any]) -> str:
    definitions = contract.get("$defs") or {}
    aliases = []
    mappings = []
    for editor in ("word", "slide", "cell"):
        for name, schema in (contract.get("tools", {}).get(editor) or {}).items():
            type_name = generated_type_name(name)
            aliases.append(
                f"export type {type_name} = {ts_type(schema, definitions)};"
            )
            mappings.append(f"  {name}: {type_name};")
    return (
        f"{TYPE_START}\n"
        "// Generated by tools/sync_contract.py. Do not edit this region.\n"
        + "\n".join(aliases)
        + "\n\nexport interface AiBridgeToolArgumentsMap {\n"
        + "\n".join(mappings)
        + "\n}\n"
        + TYPE_END
    )


def replace_types_region(source: str, generated: str) -> str:
    if TYPE_START in source and TYPE_END in source:
        pattern = re.compile(
            re.escape(TYPE_START) + r".*?" + re.escape(TYPE_END),
            re.DOTALL,
        )
        return pattern.sub(generated, source)
    start = source.index("export interface AiBridgeToolArgumentsMap {")
    cursor = source.index("{", start) + 1
    depth = 1
    while depth and cursor < len(source):
        if source[cursor] == "{":
            depth += 1
        elif source[cursor] == "}":
            depth -= 1
        cursor += 1
    if depth:
        raise ValueError("could not find AiBridgeToolArgumentsMap end")
    if cursor < len(source) and source[cursor] == ";":
        cursor += 1
    return source[:start] + generated + source[cursor:]


def replace_types_contract_version(
    source: str,
    contract: dict[str, Any],
) -> str:
    return re.sub(
        (
            r'((?:readonly )?'
            r'(?:contractVersion\?|contractVersion|version): )'
            r'"\d+\.\d+\.\d+";'
        ),
        lambda match: match.group(1) + f'"{contract["version"]}";',
        source,
    )


def render_plugin_region(contract: dict[str, Any], sha256: str) -> str:
    limits = contract["limits"]
    definitions = contract.get("$defs") or {}
    argument_schemas = {
        editor: {
            name: compact_plugin_schema(resolve_schema(schema, definitions))
            for name, schema in contract["tools"][editor].items()
        }
        for editor in ("word", "slide", "cell")
    }
    read_only = [
        name
        for editor in ("word", "slide", "cell")
        for name, schema in contract["tools"][editor].items()
        if schema["x-effects"] == "read"
    ]
    image_tools = [
        name
        for editor in ("word", "slide", "cell")
        for name, schema in contract["tools"][editor].items()
        if schema["x-argumentLimitClass"] == "image"
    ]
    lines = [
        PLUGIN_START,
        "  // Generated by tools/sync_contract.py. Do not edit this region.",
        f'  const PLUGIN_VERSION = {json.dumps(contract["version"])};',
        f"  const PROTOCOL_VERSION = {int(contract['protocolVersion'])};",
        f'  const CONTRACT_SHA256 = "{sha256}";',
        f'  const PLUGIN_GUID = "{contract["pluginGuid"]}";',
        f"  const MAX_CALLS = {int(limits['maxToolCalls'])};",
        "  const MAX_CACHED_REQUESTS = 100;",
        (
            "  const REQUEST_ID_PATTERN = "
            f"/^[A-Za-z0-9._:-]{{1,{int(limits['maxRequestIdChars'])}}}$/;"
        ),
        (
            "  const CHANNEL_ID_PATTERN = "
            f"/^[A-Za-z0-9._:-]{{16,{int(limits['maxRequestIdChars'])}}}$/;"
        ),
        "  const READ_ONLY_TOOLS = new Set([",
    ]
    lines.extend(f'    "{name}",' for name in read_only)
    lines.extend(["  ]);", "  const IMAGE_SOURCE_TOOLS = new Set(["])
    lines.extend(f'    "{name}",' for name in image_tools)
    lines.extend(["  ]);", "", "  const ALLOWED_TOOLS = {"])
    for editor in ("word", "slide", "cell"):
        lines.append(f"    {editor}: new Set([")
        lines.extend(
            f'      "{name}",' for name in contract["tools"][editor]
        )
        lines.append("    ]),")
    lines.extend([
        "  };",
        (
            "  const ARGUMENT_SCHEMAS = "
            + json.dumps(
                argument_schemas,
                ensure_ascii=False,
                separators=(",", ":"),
                sort_keys=True,
            )
            + ";"
        ),
        PLUGIN_END,
    ])
    return "\n".join(lines)


def replace_plugin_region(source: str, generated: str) -> str:
    if PLUGIN_START in source and PLUGIN_END in source:
        pattern = re.compile(
            re.escape(PLUGIN_START) + r".*?" + re.escape(PLUGIN_END),
            re.DOTALL,
        )
        return pattern.sub(generated, source)
    start = source.index("  const PLUGIN_VERSION")
    end = source.index("\n\n  const state =", start)
    return source[:start] + generated + source[end:]


def replace_js_contract_header(
    source: str,
    contract: dict[str, Any],
    sha256: str,
    include_guid: bool,
) -> str:
    limits = contract["limits"]
    source = re.sub(
        r'  const VERSION = "[^"]+";',
        f'  const VERSION = "{contract["version"]}";',
        source,
        count=1,
    )
    source = re.sub(
        r"  const PROTOCOL_VERSION = \d+;",
        f"  const PROTOCOL_VERSION = {int(contract['protocolVersion'])};",
        source,
        count=1,
    )
    if "const CONTRACT_SHA256" in source:
        source = re.sub(
            r'  const CONTRACT_SHA256 = "[0-9a-f]+";',
            f'  const CONTRACT_SHA256 = "{sha256}";',
            source,
            count=1,
        )
    elif include_guid:
        protocol_line = f"  const PROTOCOL_VERSION = {int(contract['protocolVersion'])};"
        source = source.replace(
            protocol_line,
            protocol_line + f'\n  const CONTRACT_SHA256 = "{sha256}";',
            1,
        )
    source = re.sub(
        r"  const REQUEST_ID_PATTERN = /\^\[A-Za-z0-9\._:-\]\{1,\d+\}\$/;",
        (
            "  const REQUEST_ID_PATTERN = "
            f"/^[A-Za-z0-9._:-]{{1,{int(limits['maxRequestIdChars'])}}}$/;"
        ),
        source,
        count=1,
    )
    source = re.sub(
        r"  const CHANNEL_ID_PATTERN = /\^\[A-Za-z0-9\._:-\]\{16,\d+\}\$/;",
        (
            "  const CHANNEL_ID_PATTERN = "
            f"/^[A-Za-z0-9._:-]{{16,{int(limits['maxRequestIdChars'])}}}$/;"
        ),
        source,
        count=1,
    )
    if include_guid:
        required_images = []
        optional_images = []
        for editor in ("word", "slide", "cell"):
            for name, schema in contract["tools"][editor].items():
                if schema.get("x-argumentLimitClass") != "image":
                    continue
                target = (
                    required_images
                    if "source" in (schema.get("required") or [])
                    else optional_images
                )
                target.append(name)
        required_literal = json.dumps(
            required_images,
            ensure_ascii=False,
            separators=(",", ":"),
        )
        optional_literal = json.dumps(
            optional_images,
            ensure_ascii=False,
            separators=(",", ":"),
        )
        source = re.sub(
            r"  const IMAGE_TOOL_NAMES = new Set\([^\n]+\);",
            f"  const IMAGE_TOOL_NAMES = new Set({required_literal});",
            source,
            count=1,
        )
        source = re.sub(
            r"  const OPTIONAL_IMAGE_TOOL_NAMES = new Set\([^\n]+\);",
            (
                "  const OPTIONAL_IMAGE_TOOL_NAMES = "
                f"new Set({optional_literal});"
            ),
            source,
            count=1,
        )
    return source


def scoped_contract(
    contract: dict[str, Any],
    editor: str,
    sha256: str,
) -> dict[str, Any]:
    config = EDITOR_CONFIG[editor]
    return {
        "name": contract.get("name"),
        "version": contract["version"],
        "protocolVersion": contract["protocolVersion"],
        "contractSha256": sha256,
        "editorType": editor,
        "toolPrefix": config["prefix"],
        "limits": contract["limits"],
        "unitConventions": (
            contract.get("unitConventions", {}).get(config["unitKey"]) or {}
        ),
        "inputNormalization": (
            contract.get("inputNormalization", {}).get(
                config["normalizationKey"]
            )
            or {}
        ),
        "controls": contract.get("controls") or [],
        "errors": contract.get("errors") or [],
        "$defs": contract.get("$defs") or {},
        "tools": contract["tools"][editor],
    }


def collect_enums(
    schema: Any,
    definitions: dict[str, Any],
    path: str = "",
) -> list[tuple[str, list[Any]]]:
    if not isinstance(schema, dict):
        return []
    resolved = resolve_schema(schema, definitions)
    values = []
    if isinstance(resolved.get("enum"), list):
        values.append((path or "(root)", resolved["enum"]))
    for name, child in (resolved.get("properties") or {}).items():
        values.extend(
            collect_enums(
                child,
                definitions,
                f"{path}.{name}".strip("."),
            )
        )
    item_schema = resolved.get("items")
    if isinstance(item_schema, dict):
        values.extend(collect_enums(item_schema, definitions, path + "[]"))
    return values


def render_markdown(scoped: dict[str, Any]) -> str:
    lines = [
        f"# {EDITOR_CONFIG[scoped['editorType']]['label']} ai-bridge 生成契约",
        "",
        "> 由 ai-bridge `public-api.json` 生成，禁止手工编辑。",
        "",
        f"- 契约版本：`{scoped['version']}`",
        f"- 协议版本：`{scoped['protocolVersion']}`",
        f"- SHA-256：`{scoped['contractSha256']}`",
        f"- 工具数量：`{len(scoped['tools'])}`",
        "",
        "## 限制",
        "",
        "| 名称 | 值 |",
        "| --- | ---: |",
    ]
    for name, value in scoped["limits"].items():
        lines.append(f"| `{name}` | `{value}` |")
    lines.extend(
        [
            "",
            "## 单位约定",
            "",
            (
                scoped["unitConventions"]
                if isinstance(scoped["unitConventions"], str)
                else "```json\n"
                + json.dumps(
                    scoped["unitConventions"],
                    ensure_ascii=False,
                    indent=2,
                )
                + "\n```"
            ),
            "",
            "## 输入归一化",
            "",
            "```json",
            json.dumps(
                scoped["inputNormalization"],
                ensure_ascii=False,
                indent=2,
            ),
            "```",
            "",
            "## 工具",
            "",
        ]
    )
    definitions = scoped["$defs"]
    for name, schema in scoped["tools"].items():
        required = list(schema.get("required") or [])
        properties = list((schema.get("properties") or {}).keys())
        optional = [field for field in properties if field not in required]
        validators = schema.get("x-semanticValidators") or []
        enums = collect_enums(schema, definitions)
        lines.extend(
            [
                f"### `{name}`",
                "",
                str(schema.get("description") or ""),
                "",
                f"- 读写属性：`{schema.get('x-effects')}`",
                f"- 参数限制类：`{schema.get('x-argumentLimitClass')}`",
                "- 必填字段：" + (
                    ", ".join(f"`{item}`" for item in required) or "无"
                ),
                "- 可选字段：" + (
                    ", ".join(f"`{item}`" for item in optional) or "无"
                ),
                "- 语义校验器：" + (
                    ", ".join(f"`{item}`" for item in validators) or "无"
                ),
            ]
        )
        if enums:
            lines.append("- 枚举：")
            for enum_path, values in enums:
                lines.append(
                    "  - `"
                    + enum_path
                    + "`："
                    + " | ".join(f"`{value}`" for value in values)
                )
        else:
            lines.append("- 枚举：无")
        lines.extend(
            [
                "",
                "<details><summary>完整 JSON Schema</summary>",
                "",
                "```json",
                json.dumps(schema, ensure_ascii=False, indent=2),
                "```",
                "",
                "</details>",
                "",
            ]
        )
    lines.extend(
        [
            "## `$defs`",
            "",
            "```json",
            json.dumps(scoped["$defs"], ensure_ascii=False, indent=2),
            "```",
            "",
        ]
    )
    return "\n".join(lines)


def toml_string(value: str) -> str:
    return json.dumps(value, ensure_ascii=False)


def action_params(kind: str) -> list[str]:
    if kind == "tool":
        return [
            '  { name = "arguments_json", type = "json object string", required = true, description = "业务参数 JSON；字段与枚举以生成契约为准" },',
            '  { name = "request_id", type = "string", required = true, description = "稳定幂等 ID；写操作重试必须复用" },',
            '  { name = "timeout_ms", type = "integer", required = false, description = "等待页面响应的毫秒数" }',
        ]
    if kind in {"batch", "validate"}:
        return [
            '  { name = "tool_calls_json", type = "json array string", required = true, description = "工具调用数组；业务参数由 /bridge/validate 按生成契约校验" },',
            '  { name = "request_id", type = "string", required = true, description = "稳定幂等 ID" },',
            '  { name = "timeout_ms", type = "integer", required = false, description = "等待页面响应的毫秒数" }',
        ]
    return [
        '  { name = "request_id", type = "string", required = true, description = "稳定幂等 ID" },',
        '  { name = "timeout_ms", type = "integer", required = false, description = "等待页面响应的毫秒数" }',
    ]


def document_path_source(action: str) -> str:
    if action not in {"session", "validate", "execute"}:
        raise ValueError(f"unsupported document action: {action}")
    command = (
        "document_id=$DOCUMENT_HUB_DOCUMENT_ID; "
        "case \"$document_id\" in "
        "????????-????-4???-[89ab]???-????????????) ;; *) exit 64 ;; esac; "
        "case \"$document_id\" in *[!0-9a-f-]*) exit 64 ;; esac; "
        f"printf \"%s\" \"/api/v1/documents/$document_id/ai/{action}\""
    )
    return "{ from = \"shell\", cmd = '''" + command + "''', timeout_ms = 1000, trim = true }"


def normalize_httpx_base_url(value: str | None) -> str:
    if not isinstance(value, str) or not value.strip():
        raise ValueError(
            "--httpx-base-url is required when --zenmind-root is used"
        )
    candidate = value.strip()
    parsed = urlsplit(candidate)
    if parsed.scheme not in {"http", "https"} or not parsed.hostname:
        raise ValueError(
            "--httpx-base-url must be an absolute HTTP(S) origin"
        )
    if parsed.username is not None or parsed.password is not None:
        raise ValueError("--httpx-base-url must not contain credentials")
    if parsed.path not in {"", "/"} or parsed.query or parsed.fragment:
        raise ValueError(
            "--httpx-base-url must not contain a path, query, or fragment"
        )
    return urlunsplit((parsed.scheme, parsed.netloc, "", "", ""))


def render_toml(
    contract: dict[str, Any],
    editor: str,
    sha256: str,
    httpx_base_url: str,
) -> str:
    config = EDITOR_CONFIG[editor]
    file_type = config["fileType"]
    label = config["label"]
    session_action = config["sessionAction"]
    base_url = normalize_httpx_base_url(httpx_base_url)
    lines = [
        "version = 1",
        (
            "description = "
            + toml_string(
                f"OFFICE {label} bridge（document-hub 匿名访问）；"
                f"业务参数由 ai-bridge {contract['version']} 契约定义"
            )
        ),
        f"base_url = {toml_string(base_url)}",
        'timeout = "120s"',
        "retries = 0",
        'state_scope = "chat"',
        "",
        "[headers]",
        'Accept = "application/json"',
        f'User-Agent = "agent-platform-httpx/{config["site"]}"',
        "",
        "[actions.health]",
        'description = "检查 document-hub 服务健康状态"',
        'path = "/health/ready"',
        "expect_status = 200",
        'extract_type = "jq"',
        'extract_expr = ".body"',
        "",
    ]
    lines.extend(
        [
            f"[actions.new_{file_type}]",
            f'description = "以固定 Platform 主体新建 {file_type.upper()} 文档"',
            'method = "POST"',
            'path = "/api/v1/documents"',
            (
                'body = { title = { from = "param", key = "title" }, '
                f'fileType = {{ from = "literal", value = "{file_type}" }} }}'
            ),
            "expect_status = 201",
            "params = [",
            f'  {{ name = "title", type = "string", required = true, description = "新建 {file_type.upper()} 文档标题" }}',
            "]",
            'extract_type = "jq"',
            (
                "extract_expr = '''.body | . + "
                f'{{documentId:.id,fileName:(.id + ".{file_type}"),'
                f'editorUrl:("{base_url}/documents/" + .id)}}'
                "'''"
            ),
            "",
        ]
    )
    lines.extend(
        [
            f"[actions.{session_action}]",
            f'description = "确认固定 Platform 主体在目标文档上存在在线 {label} 编辑器"',
            f"path = {document_path_source('session')}",
            "expect_status = 200",
            'extract_type = "jq"',
            (
                "extract_expr = '''if (.body.online == true "
                f'and .body.editorType == "{editor}") then .body '
                f"else error(\"EDITOR_SESSION_UNAVAILABLE: {config['site']}\") end'''"
            ),
            "",
            "[actions.get_state]",
            f'description = "读取当前 {label} bridge state 和契约身份"',
            'method = "POST"',
            f"path = {document_path_source('execute')}",
            'body = { method = "getState", requestId = { from = "param", key = "request_id" }, timeoutMs = { from = "param", key = "timeout_ms", default = 30000 } }',
            "expect_status = 200",
            "params = [",
            *action_params("control"),
            "]",
            'extract_type = "jq"',
            'extract_expr = ".body | del(.session)"',
            "",
        ]
    )
    for name, schema in contract["tools"][editor].items():
        lines.extend(
            [
                f"[actions.{name}]",
                (
                    "description = "
                    + toml_string(
                        f"{schema['description']}；业务参数统一通过 arguments_json"
                    )
                ),
                'method = "POST"',
                f"path = {document_path_source('execute')}",
                (
                    'body = { method = "executeTool", '
                    f'name = "{name}", argumentsJson = {{ from = "param", key = "arguments_json" }}, '
                    'requestId = { from = "param", key = "request_id" }, '
                    'timeoutMs = { from = "param", key = "timeout_ms", default = 90000 } }'
                ),
                "expect_status = 200",
                "params = [",
                *action_params("tool"),
                "]",
                'extract_type = "jq"',
                'extract_expr = ".body | del(.session)"',
                "",
            ]
        )
    batch_name = f"execute_{config['prefix'].rstrip('_')}_batch"
    validate_name = f"validate_{config['prefix'].rstrip('_')}_batch"
    for action_name, document_action, method, kind in (
        (validate_name, "validate", None, "validate"),
        (batch_name, "execute", "executeBatch", "batch"),
    ):
        body_parts = []
        if method:
            body_parts.append(f'method = "{method}"')
        body_parts.extend(
            [
                'toolCallsJson = { from = "param", key = "tool_calls_json" }',
                'requestId = { from = "param", key = "request_id" }',
                'timeoutMs = { from = "param", key = "timeout_ms", default = 90000 }',
            ]
        )
        lines.extend(
            [
                f"[actions.{action_name}]",
                f'description = "{label} 批量{"预检" if kind == "validate" else "执行"}"',
                'method = "POST"',
                f"path = {document_path_source(document_action)}",
                "body = { " + ", ".join(body_parts) + " }",
                "expect_status = 200",
                "params = [",
                *action_params(kind),
                "]",
                'extract_type = "jq"',
                'extract_expr = ".body | del(.session)"',
                "",
            ]
        )
    for control in contract.get("controls") or []:
        lines.extend(
            [
                f"[actions.{control}]",
                f'description = "执行 {label} {control} 控制操作"',
                'method = "POST"',
                f"path = {document_path_source('execute')}",
                (
                    f'body = {{ method = "{control}", '
                    'requestId = { from = "param", key = "request_id" }, '
                    'timeoutMs = { from = "param", key = "timeout_ms", default = 90000 } }'
                ),
                "expect_status = 200",
                "params = [",
                *action_params("control"),
                "]",
                'extract_type = "jq"',
                'extract_expr = ".body | del(.session)"',
                "",
            ]
        )
    return "\n".join(lines).rstrip() + "\n"


def render_runtime_helper(scoped: dict[str, Any]) -> str:
    editor = scoped["editorType"]
    prefix = scoped["toolPrefix"]
    limits = scoped["limits"]
    return f'''"""Generated ai-bridge contract loader. Do not edit."""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any, Dict, Iterable, Optional


CONTRACT_PATH = Path(__file__).resolve().parents[1] / "references" / "contract.generated.json"
CONTRACT: Dict[str, Any] = json.loads(CONTRACT_PATH.read_text(encoding="utf-8"))
CONTRACT_VERSION = str(CONTRACT["version"])
CONTRACT_SHA256 = str(CONTRACT["contractSha256"])
EDITOR_TYPE = {editor!r}
TOOL_PREFIX = {prefix!r}
TOOL_SCHEMAS: Dict[str, Dict[str, Any]] = dict(CONTRACT["tools"])
TOOL_NAMES = frozenset(TOOL_SCHEMAS)
READ_ONLY_ACTIONS = frozenset(
    name for name, schema in TOOL_SCHEMAS.items()
    if schema.get("x-effects") == "read"
)
IMAGE_SOURCE_ACTIONS = frozenset(
    name for name, schema in TOOL_SCHEMAS.items()
    if schema.get("x-argumentLimitClass") == "image"
)
MAX_BATCH_CALLS = int(CONTRACT["limits"]["maxToolCalls"])
MAX_ARGUMENT_CHARS = int(CONTRACT["limits"]["maxArgumentsJsonChars"])
MAX_IMAGE_REQUEST_BYTES = int(CONTRACT["limits"]["maxImageRequestBytes"])
MAX_REQUEST_ID_CHARS = int(CONTRACT["limits"]["maxRequestIdChars"])
MAX_TIMEOUT_MS = int(CONTRACT["limits"]["maxTimeoutMs"])
REQUEST_ID_PATTERN = re.compile(
    r"^[A-Za-z0-9._:-]{{1," + str(MAX_REQUEST_ID_CHARS) + r"}}$"
)


def argument_limit_class(tool_names: Iterable[str]) -> str:
    return "image" if any(name in IMAGE_SOURCE_ACTIONS for name in tool_names) else "standard"


def validate_arguments_size(
    value: Any,
    label: str,
    tool_names: Iterable[str],
) -> None:
    encoded = json.dumps(value, ensure_ascii=False, separators=(",", ":"))
    limit_class = argument_limit_class(tool_names)
    actual = len(encoded.encode("utf-8")) if limit_class == "image" else len(encoded)
    limit = MAX_IMAGE_REQUEST_BYTES if limit_class == "image" else MAX_ARGUMENT_CHARS
    unit = "bytes" if limit_class == "image" else "characters"
    if actual > limit:
        raise ValueError(
            "{{0}} exceeds {{1}} serialized {{2}} ({{3}} limit class)".format(
                label, limit, unit, limit_class
            )
        )


def response_body(payload: Dict[str, Any]) -> Dict[str, Any]:
    body = payload.get("body")
    return body if isinstance(body, dict) else payload


def contract_mismatch(payload: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    body = response_body(payload)
    received_version = body.get("contractVersion")
    received_sha256 = body.get("contractSha256")
    if received_sha256 == CONTRACT_SHA256:
        return None
    status = payload.get("status")
    request_failed = (
        payload.get("ok") is False
        or body.get("ok") is False
        or isinstance(payload.get("error"), dict)
        or (isinstance(status, int) and not 200 <= status < 300)
    )
    if request_failed:
        return None
    if not isinstance(received_sha256, str) or not received_sha256.strip():
        return {{
            "ok": False,
            "phase": "contract",
            "error": {{
                "code": "CONTRACT_METADATA_MISSING",
                "message": "成功响应缺少 ai-bridge 契约元数据。",
                "details": {{
                    "expectedContractVersion": CONTRACT_VERSION,
                    "expectedContractSha256": CONTRACT_SHA256,
                    "receivedContractVersion": received_version,
                    "receivedContractSha256": received_sha256,
                }},
            }},
        }}
    return {{
        "ok": False,
        "phase": "contract",
        "error": {{
            "code": "CONTRACT_VERSION_MISMATCH",
            "message": "生成技能契约与当前 ai-bridge 服务不一致。",
            "details": {{
                "expectedContractVersion": CONTRACT_VERSION,
                "expectedContractSha256": CONTRACT_SHA256,
                "receivedContractVersion": received_version,
                "receivedContractSha256": received_sha256,
            }},
        }},
    }}
'''


def build_artifacts(
    contract: dict[str, Any],
    zenmind_root: Path | None,
    httpx_base_url: str | None = None,
) -> dict[Path, str]:
    sha256 = contract_sha256(contract)
    artifacts = {
        TYPES_PATH: replace_types_contract_version(
            replace_types_region(
                TYPES_PATH.read_text(encoding="utf-8"),
                render_types_region(contract),
            ),
            contract,
        ),
        PLUGIN_PATH: replace_plugin_region(
            PLUGIN_PATH.read_text(encoding="utf-8"),
            render_plugin_region(contract, sha256),
        ),
        HOST_PATH: replace_js_contract_header(
            HOST_PATH.read_text(encoding="utf-8"),
            contract,
            sha256,
            True,
        ),
        CLIENT_PATH: replace_js_contract_header(
            CLIENT_PATH.read_text(encoding="utf-8"),
            contract,
            sha256,
            False,
        ),
    }
    config = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    config["version"] = contract["version"]
    for variation in config.get("variations") or []:
        if isinstance(variation, dict) and isinstance(variation.get("url"), str):
            variation["url"] = re.sub(
                r"([?&]v=)[^-&]+(-rev\d+)?",
                lambda match: (
                    match.group(1)
                    + contract["version"]
                    + (match.group(2) or "-rev1")
                ),
                variation["url"],
            )
    artifacts[CONFIG_PATH] = (
        json.dumps(config, ensure_ascii=False, indent=2) + "\n"
    )
    artifacts[INDEX_PATH] = re.sub(
        r"([?&]v=)[^-\"&]+(-rev\d+)?",
        lambda match: (
            match.group(1)
            + contract["version"]
            + (match.group(2) or "-rev1")
        ),
        INDEX_PATH.read_text(encoding="utf-8"),
    )
    if zenmind_root is None:
        return artifacts
    if not zenmind_root.is_dir():
        raise ValueError(f"zenmind root does not exist: {zenmind_root}")
    normalized_httpx_base_url = normalize_httpx_base_url(httpx_base_url)
    for editor, config in EDITOR_CONFIG.items():
        scoped = scoped_contract(contract, editor, sha256)
        skill_root = (
            zenmind_root
            / "skills-center"
            / config["skill"]
        )
        references = skill_root / "references"
        skill_path = skill_root / "SKILL.md"
        artifacts[skill_path] = render_skill_version(
            skill_path.read_text(encoding="utf-8"),
            contract["version"],
        )
        artifacts[references / "contract.generated.json"] = (
            json.dumps(scoped, ensure_ascii=False, indent=2) + "\n"
        )
        artifacts[references / "contract.generated.md"] = render_markdown(scoped)
        artifacts[
            skill_root
            / "scripts"
            / f"_{config['fileType']}_contract_runtime.py"
        ] = render_runtime_helper(scoped)
        artifacts[
            skill_root / ".config" / "httpx" / config["toml"]
        ] = render_toml(
            contract,
            editor,
            sha256,
            normalized_httpx_base_url,
        )
    return artifacts


def validate_contract(contract: dict[str, Any]) -> None:
    if contract.get("version") != "0.1.0":
        raise ValueError("public-api.json version must be 0.1.0")
    if contract.get("protocolVersion") != 1:
        raise ValueError("protocolVersion must remain 1")
    names = []
    for editor in ("word", "slide", "cell"):
        schemas = contract.get("tools", {}).get(editor)
        if not isinstance(schemas, dict) or not schemas:
            raise ValueError(f"missing tools.{editor}")
        for name, schema in schemas.items():
            names.append(name)
            if not isinstance(schema.get("description"), str):
                raise ValueError(f"{editor}.{name} is missing description")
            if schema.get("x-effects") not in {"read", "write"}:
                raise ValueError(f"{editor}.{name} has invalid x-effects")
            if schema.get("x-argumentLimitClass") not in {"standard", "image"}:
                raise ValueError(
                    f"{editor}.{name} has invalid x-argumentLimitClass"
                )
            if not isinstance(schema.get("x-semanticValidators"), list):
                raise ValueError(
                    f"{editor}.{name} has invalid x-semanticValidators"
                )
            resolve_schema(schema, contract.get("$defs") or {})
    if len(names) != 155 or len(names) != len(set(names)):
        raise ValueError(
            f"expected 155 unique tools, found {len(names)} total/{len(set(names))} unique"
        )


def drifted_paths(artifacts: dict[Path, str]) -> list[Path]:
    return [
        path
        for path, expected in artifacts.items()
        if not path.exists() or path.read_text(encoding="utf-8") != expected
    ]


def main() -> int:
    parser = argparse.ArgumentParser()
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--write", action="store_true")
    mode.add_argument("--check", action="store_true")
    parser.add_argument(
        "--zenmind-root",
        type=Path,
        help="zenmind-env repository root; enables TOML and skill projections",
    )
    parser.add_argument(
        "--httpx-base-url",
        help=(
            "HTTP(S) origin used by generated HTTPX bridge sites; "
            "required with --zenmind-root"
        ),
    )
    args = parser.parse_args()
    contract = json.loads(CONTRACT_PATH.read_text(encoding="utf-8"))
    validate_contract(contract)
    try:
        artifacts = build_artifacts(
            contract,
            args.zenmind_root,
            args.httpx_base_url,
        )
    except ValueError as error:
        parser.error(str(error))
    drift = drifted_paths(artifacts)
    for path in drift:
        expected = artifacts[path]
        if args.write:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(expected, encoding="utf-8")
            print(f"updated {path}")
        else:
            print(f"contract projection drift: {path}", file=sys.stderr)
    if args.check and drift:
        return 1
    print(
        f"contract {contract['version']} "
        f"{contract_sha256(contract)}: {len(artifacts)} projections OK"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
