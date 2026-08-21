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
BOOTSTRAP_PATH = BRIDGE_ROOT / "bootstrap.js"
STATIC_RUNTIME_PATHS = (
    BOOTSTRAP_PATH,
    BRIDGE_ROOT / "bridges" / "word-bridge.js",
    BRIDGE_ROOT / "bridges" / "slides-bridge.js",
    BRIDGE_ROOT / "bridges" / "sheets-bridge.js",
)
ASSET_REVISION_PLACEHOLDER = "__AI_BRIDGE_ASSET_REVISION__"
REQUEST_ID_DESCRIPTION = (
    "不可变幂等 ID；仅在结果未知且 action、参数、timeout 完全不变时原样复用。"
    "修改任一字段必须使用新 ID；request_id_conflict/REQUEST_ID_CONFLICT 不可用旧 ID 重试"
)
# The public contract remains a planning recommendation. The generated plugin
# tolerates modest model counting errors without exposing this private ceiling.
MAX_ACCEPTED_TOOL_CALLS = 30
DOCUMENT_ID_PATTERN = (
    "^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-"
    "[89ab][0-9a-f]{3}-[0-9a-f]{12}$"
)
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
        "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "prefix": "word_",
        "label": "Word",
        "sessionAction": "session",
        "unitKey": "word",
        "normalizationKey": "word",
    },
    "slide": {
        "skill": "online-pptx",
        "toml": "online-pptx-bridge.toml",
        "site": "online-pptx-bridge",
        "fileType": "pptx",
        "mimeType": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "prefix": "slides_",
        "label": "Slides",
        "sessionAction": "session",
        "unitKey": "slide",
        "normalizationKey": "slide",
    },
    "cell": {
        "skill": "online-xlsx",
        "toml": "online-xlsx-bridge.toml",
        "site": "online-xlsx-bridge",
        "fileType": "xlsx",
        "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "prefix": "sheets_",
        "label": "Sheets",
        "sessionAction": "session",
        "unitKey": "sheet",
        "normalizationKey": "sheet",
    },
}

WORD_CONTRACT_GROUPS = {
    "content": (
        "inspect", "replace_text", "append_paragraph", "insert_paragraph",
        "format_document", "format_selection", "format_matches",
        "delete_matches", "add_hyperlink", "add_comment", "add_bookmark",
        "format_paragraphs", "set_paragraph_text", "delete_paragraphs",
        "set_list", "insert_page_break", "navigate", "scroll",
        "scale_font", "set_document_text",
    ),
    "tables": (
        "format_table_advanced", "add_nested_table", "add_table",
        "set_table_cell", "format_table", "edit_table",
    ),
    "structure": (
        "inspect_advanced", "set_document_properties", "manage_section",
        "manage_style", "set_tabs", "set_numbering", "set_page_layout",
        "set_header_footer",
    ),
    "objects": (
        "add_image", "manage_drawing", "add_shape", "add_chart",
        "add_math", "add_ole_object",
    ),
    "long-document": (
        "manage_fields", "manage_long_document", "manage_comments",
        "manage_revisions", "set_watermark",
    ),
    "special": (
        "set_protection", "manage_content_control", "manage_custom_xml",
        "inspect_macros", "set_macros",
    ),
}

SLIDE_CONTRACT_GROUPS = {
    "structure": (
        "inspect", "inspect_layouts", "inspect_backgrounds", "inspect_themes",
        "inspect_builtin_themes", "inspect_objects", "validate_layout",
        "add_slide", "duplicate_slide", "delete_slide", "move_slide",
        "set_visibility", "set_size", "apply_layout", "set_show_settings",
    ),
    "text": (
        "replace_text", "scale_font", "format_text", "format_selection",
        "set_text_content", "format_paragraphs", "add_textbox",
        "add_word_art", "add_math",
    ),
    "templates": (
        "apply_theme", "apply_builtin_theme", "set_theme", "create_layout",
        "add_template_shape", "manage_template_object",
        "set_template_background", "set_background",
    ),
    "objects": (
        "update_object", "set_hyperlink", "align_objects", "group_objects",
        "reorder_object", "add_connector", "add_freeform", "add_image",
        "add_image_shape", "add_ole_object", "add_shape", "update_shape",
        "delete_object",
    ),
    "tables": (
        "add_table", "set_table_cell", "edit_table", "format_table",
    ),
    "smartart": (
        "inspect_smartarts", "add_smartart", "update_smartart",
        "delete_smartart",
    ),
    "charts": (
        "inspect_charts", "add_chart", "update_chart", "delete_chart",
    ),
    "collaboration": (
        "set_notes", "add_comment", "inspect_comments", "manage_comment",
        "set_transition", "inspect_animations", "manage_animation",
        "inspect_macros", "set_macros", "control_slideshow",
    ),
}

CELL_CONTRACT_GROUPS = {
    "cells": (
        "inspect", "set_values", "set_formula", "inspect_range",
        "set_array_formula", "replace_text", "format_range",
        "set_rich_text", "recalculate",
    ),
    "structure": (
        "add_sheet", "rename_sheet", "delete_sheet", "manage_sheet",
        "manage_range", "inspect_names", "manage_names", "sort", "filter",
    ),
    "tables": (
        "inspect_tables", "manage_table", "manage_conditional_format",
        "manage_validation", "inspect_pivots", "manage_pivot",
    ),
    "objects": (
        "inspect_drawings", "manage_drawing", "manage_hyperlink",
        "inspect_comments", "manage_comments",
    ),
    "workbook": (
        "inspect_freeze_panes", "manage_freeze_panes", "inspect_properties",
        "manage_properties", "inspect_protected_ranges",
        "manage_protected_ranges", "inspect_page_layout",
        "manage_page_layout", "inspect_macros", "set_macros",
    ),
    "charts": (
        "add_chart", "inspect_charts", "update_chart", "delete_chart",
    ),
}

EDITOR_CONTRACT_GROUPS = {
    "word": WORD_CONTRACT_GROUPS,
    "slide": SLIDE_CONTRACT_GROUPS,
    "cell": CELL_CONTRACT_GROUPS,
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


def replace_asset_revision_queries(source: str, revision: str) -> str:
    """Replace every existing v query parameter with one atomic asset revision."""
    return re.sub(
        r"([?&]v=)[^&\"'\s<>]+",
        lambda match: match.group(1) + revision,
        source,
    )


def compute_asset_revision(version: str, sources: dict[str, str]) -> str:
    """Return a deterministic revision for an ordered set of browser assets."""
    digest = hashlib.sha256()
    for name in sorted(sources):
        digest.update(name.encode("utf-8"))
        digest.update(b"\0")
        digest.update(sources[name].encode("utf-8"))
        digest.update(b"\0")
    return f"{version}-{digest.hexdigest()}"


def runtime_asset_sources(
    contract: dict[str, Any],
    generated_artifacts: dict[Path, str],
    normalized_config: str,
    normalized_index: str,
) -> dict[str, str]:
    sources = {
        "public-api.json": canonical_json(contract) + "\n",
        "plugin.js": generated_artifacts[PLUGIN_PATH],
        "host-bridge.js": generated_artifacts[HOST_PATH],
        "client-sdk.js": generated_artifacts[CLIENT_PATH],
        "config.json": normalized_config,
        "index.html": normalized_index,
    }
    sources.update(
        {
            str(path.relative_to(BRIDGE_ROOT)): path.read_text(encoding="utf-8")
            for path in STATIC_RUNTIME_PATHS
        }
    )
    return sources


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
    consequence_properties = consequence.get("properties") or {}

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
        property_type = ts_type(
            consequence_properties.get(name, base_properties.get(name, {})),
            definitions,
        )
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


def replace_types_error_codes(source: str, contract: dict[str, Any]) -> str:
    codes = ["AI_BRIDGE_ERROR"]
    for code in contract.get("errors") or []:
        if isinstance(code, str) and code not in codes:
            codes.append(code)
    rendered = (
        "export type AiBridgeErrorCode =\n"
        + "\n".join(f"  | {json.dumps(code)}" for code in codes)
        + "\n  | string;"
    )
    pattern = re.compile(
        r"export type AiBridgeErrorCode =\n.*?\n  \| string;",
        re.DOTALL,
    )
    updated, count = pattern.subn(rendered, source, count=1)
    if count != 1:
        raise ValueError("public-api.d.ts must define AiBridgeErrorCode")
    return updated


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
        f"  const PUBLIC_MAX_CALLS = {int(limits['maxToolCalls'])};",
        f"  const MAX_CALLS = {MAX_ACCEPTED_TOOL_CALLS};",
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
    tools = {
        public_tool_name(editor, name): schema
        for name, schema in contract["tools"][editor].items()
    }
    return {
        "name": contract.get("name"),
        "version": contract["version"],
        "protocolVersion": contract["protocolVersion"],
        "contractSha256": sha256,
        "editorType": editor,
        "toolPrefix": "",
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
        "runtimeCapabilities": (
            (contract.get("runtimeCapabilities") or {}).get(editor) or {}
        ),
        "errors": contract.get("errors") or [],
        "$defs": contract.get("$defs") or {},
        "tools": tools,
    }


def referenced_definitions(
    value: Any,
    definitions: dict[str, Any],
) -> dict[str, Any]:
    names: set[str] = set()

    def visit(node: Any) -> None:
        if isinstance(node, list):
            for item in node:
                visit(item)
            return
        if not isinstance(node, dict):
            return
        reference = node.get("$ref")
        if isinstance(reference, str) and reference.startswith("#/$defs/"):
            name = reference.removeprefix("#/$defs/")
            if name not in definitions:
                raise ValueError(f"unresolved contract reference: {reference}")
            if name not in names:
                names.add(name)
                visit(definitions[name])
        for child in node.values():
            visit(child)

    visit(value)
    return {name: definitions[name] for name in definitions if name in names}


def scoped_contract_group(
    scoped: dict[str, Any],
    tool_names: tuple[str, ...],
) -> dict[str, Any]:
    missing = [name for name in tool_names if name not in scoped["tools"]]
    if missing:
        raise ValueError(f"missing grouped tools: {', '.join(missing)}")
    grouped = copy.deepcopy(scoped)
    grouped["tools"] = {name: scoped["tools"][name] for name in tool_names}
    grouped["$defs"] = referenced_definitions(
        grouped["tools"],
        scoped["$defs"],
    )
    return grouped


def public_tool_name(editor: str, internal_name: str) -> str:
    prefix = EDITOR_CONFIG[editor]["prefix"]
    if not internal_name.startswith(prefix) or internal_name == prefix:
        raise ValueError(
            f"{editor} tool {internal_name!r} must start with {prefix!r}"
        )
    return internal_name[len(prefix):]


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


def render_markdown(
    scoped: dict[str, Any],
    include_full_schema: bool = True,
) -> str:
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
        "## 运行时能力位",
        "",
        "```json",
        json.dumps(scoped.get("runtimeCapabilities") or {}, ensure_ascii=False, indent=2),
        "```",
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
        examples = schema.get("examples")
        if isinstance(examples, list) and examples:
            lines.extend(
                [
                    "",
                    "示例：",
                    "",
                    "```json",
                    json.dumps(examples[0], ensure_ascii=False, indent=2),
                    "```",
                ]
            )
        if include_full_schema:
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
        else:
            lines.append("")
    if include_full_schema:
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
    else:
        lines.extend(
            [
                "## 精确 Schema",
                "",
                "本文件只保留渐进式摘要；同名 `.generated.json` 含本组完整、自包含 JSON Schema。",
                "",
            ]
        )
    return "\n".join(lines)


def toml_string(value: str) -> str:
    return json.dumps(value, ensure_ascii=False)


def toml_value(value: Any) -> str:
    if isinstance(value, str):
        return toml_string(value)
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, list):
        return "[" + ", ".join(toml_value(item) for item in value) + "]"
    if isinstance(value, dict):
        if not value:
            return "{}"
        return (
            "{ "
            + ", ".join(
                f"{toml_string(str(key))} = {toml_value(item)}"
                for key, item in value.items()
            )
            + " }"
        )
    raise ValueError(f"unsupported TOML example value: {value!r}")


def schema_has_required_business_fields(schema: dict[str, Any]) -> bool:
    if schema.get("required"):
        return True
    for keyword in ("allOf", "anyOf", "oneOf"):
        branches = schema.get(keyword)
        if isinstance(branches, list) and any(
            isinstance(branch, dict) and branch.get("required")
            for branch in branches
        ):
            return True
    return False


def minimal_schema_example(schema: Any, definitions: dict[str, Any]) -> Any:
    if not isinstance(schema, dict):
        return None
    examples = schema.get("examples")
    if isinstance(examples, list) and examples:
        return copy.deepcopy(examples[0])
    if "$ref" in schema:
        ref = str(schema["$ref"])
        prefix = "#/$defs/"
        return minimal_schema_example(definitions.get(ref[len(prefix):], {}), definitions)
    if "const" in schema:
        return copy.deepcopy(schema["const"])
    enum = schema.get("enum")
    if isinstance(enum, list) and enum:
        return copy.deepcopy(enum[0])

    selected_branch = None
    for keyword in ("oneOf", "anyOf"):
        branches = schema.get(keyword)
        if isinstance(branches, list) and branches:
            selected_branch = branches[0]
            break

    schema_type = schema.get("type")
    if isinstance(schema_type, list):
        schema_type = next((item for item in schema_type if item != "null"), schema_type[0])
    if not schema_type and isinstance(selected_branch, dict):
        branch_example = minimal_schema_example(selected_branch, definitions)
        if not isinstance(branch_example, dict) or not schema.get("properties"):
            return branch_example
        schema_type = "object"

    if schema_type == "object" or "properties" in schema:
        properties = schema.get("properties") or {}
        required = list(schema.get("required") or [])
        if isinstance(selected_branch, dict):
            for key in selected_branch.get("required") or []:
                if key not in required:
                    required.append(key)
        return {
            key: minimal_schema_example(properties.get(key, {}), definitions)
            for key in required
        }
    if schema_type == "array":
        count = max(1, int(schema.get("minItems") or 0))
        return [minimal_schema_example(schema.get("items") or {}, definitions) for _ in range(count)]
    if schema_type == "integer":
        minimum = schema.get("minimum")
        return int(minimum) if isinstance(minimum, (int, float)) else 1
    if schema_type == "number":
        minimum = schema.get("minimum")
        if isinstance(minimum, (int, float)):
            return minimum + (1 if schema.get("exclusiveMinimum") else 0)
        return 1
    if schema_type == "boolean":
        return False
    if schema_type == "string":
        if schema.get("format") in {"uri", "url"}:
            return "https://example.test/resource"
        pattern = str(schema.get("pattern") or "")
        if "0-9A-Fa-f" in pattern and "6" in pattern:
            return "#000000"
        return "example"
    return {}


def action_params(
    kind: str,
    *,
    action_name: str = "request",
    schema: dict[str, Any] | None = None,
    definitions: dict[str, Any] | None = None,
) -> list[str]:
    request_example = f"{kind}-{action_name}-1"
    if kind == "tool":
        schema = schema or {}
        definitions = definitions or {}
        empty_read = (
            schema.get("x-effects") == "read"
            and not schema_has_required_business_fields(schema)
        )
        arguments_example = minimal_schema_example(schema, definitions)
        return [
            '  { name = "arguments", type = "object", required = '
            + ("false" if empty_read else "true")
            + ', description = "结构化业务参数；字段与枚举以生成契约为准，使用 --param-json-file 传入", example = '
            + toml_value(arguments_example)
            + " },",
            '  { name = "request_id", type = "string", required = true, description = '
            + toml_string(REQUEST_ID_DESCRIPTION)
            + ', example = '
            + toml_string(request_example)
            + " },",
            '  { name = "timeout_ms", type = "integer", required = false, description = "等待页面响应的毫秒数" }',
        ]
    if kind in {"batch", "validate"}:
        return [
            '  { name = "tool_calls", type = "array", required = true, description = "结构化工具调用数组；标准字段为 name，action 仅作兼容别名，使用 --param-json-file 传入", example = '
            + toml_value(
                [{"name": "set_size", "arguments": {"preset": "wide"}}]
            )
            + " },",
            '  { name = "request_id", type = "string", required = true, description = '
            + toml_string(REQUEST_ID_DESCRIPTION)
            + ', example = '
            + toml_string(request_example)
            + " },",
            '  { name = "timeout_ms", type = "integer", required = false, description = "等待页面响应的毫秒数" }',
        ]
    return [
        '  { name = "request_id", type = "string", required = true, description = '
        + toml_string(REQUEST_ID_DESCRIPTION)
        + ', example = '
        + toml_string(request_example)
        + " },",
        '  { name = "timeout_ms", type = "integer", required = false, description = "等待页面响应的毫秒数" }',
    ]


def document_path_source(action: str) -> str:
    if action not in {"session", "validate", "execute", "commit", "qa", "images/import"}:
        raise ValueError(f"unsupported document action: {action}")
    template = f"/api/v1/documents/{{{{value}}}}/ai/{action}"
    return (
        '{ from = "env", key = "DOCUMENT_HUB_DOCUMENT_ID", trim = true, '
        f"pattern = {toml_string(DOCUMENT_ID_PATTERN)}, "
        f"output_template = {toml_string(template)} }}"
    )


def document_download_path_source() -> str:
    return (
        '{ from = "env", key = "DOCUMENT_HUB_DOCUMENT_ID", trim = true, '
        f"pattern = {toml_string(DOCUMENT_ID_PATTERN)}, "
        'output_template = "/api/v1/documents/{{value}}/download" }'
    )


def local_image_data_url_source() -> str:
    return (
        '{ from = "file_data_url", '
        'path = { from = "env", key = "PPTX_IMAGE_PATH", trim = true }, '
        'max_bytes = 8388608, '
        'allowed_media_types = ["image/png", "image/jpeg", "image/gif", '
        '"image/webp", "image/svg+xml"] }'
    )


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
    mime_type = config["mimeType"]
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
            "[actions.create]",
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
            "[actions.upload]",
            f'description = "上传本地 {file_type.upper()}，返回真实文档 UUID 和编辑器地址"',
            'method = "POST"',
            'path = "/api/v1/documents/upload"',
            "retries = 0",
            "multipart = [",
            '  { name = "title", value = { from = "param", key = "title", default = "" } },',
            (
                '  { name = "file", file = { from = "param", key = "file_path" }, '
                f'content_type = "{mime_type}", max_bytes = 209715200 }}'
            ),
            "]",
            "expect_status = 201",
            "params = [",
            f'  {{ name = "file_path", type = "string", required = true, description = "本地 {file_type.upper()} 绝对路径" }},',
            f'  {{ name = "title", type = "string", required = false, description = "上传后的 {file_type.upper()} 文档标题；省略时使用文件名" }}',
            "]",
            'extract_type = "jq"',
            (
                "extract_expr = '''.body | "
                f'{{documentId:.id,fileName:(.id + ".{file_type}"),'
                f'editorUrl:("{base_url}/documents/" + .id)}}'
                "'''"
            ),
            "",
            "[actions.download]",
            f'description = "把当前持久化 {file_type.upper()} 版本安全下载到本地；默认不覆盖"',
            'method = "GET"',
            f"path = {document_download_path_source()}",
            "expect_status = 200",
            (
                'download = { path = { from = "param", key = "output_path" }, '
                'overwrite = { from = "param", key = "overwrite", default = false }, '
                "max_bytes = 209715200 }"
            ),
            "params = [",
            f'  {{ name = "output_path", type = "string", required = true, description = "本地 {file_type.upper()} 输出路径；父目录必须已存在" }},',
            '  { name = "overwrite", type = "boolean", required = false, description = "显式允许原子替换已有文件", example = false }',
            "]",
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
                "extract_expr = '''if (.body.contractVersion != "
                f'"{contract["version"]}" or .body.contractSha256 != "{sha256}") then '
                f"error(\"CONTRACT_MISMATCH: {config['site']}\") "
                "elif (.body.online == true "
                f'and .body.saveReady == true and .body.editorType == "{editor}") then .body '
                f"else error(\"EDITOR_SESSION_UNAVAILABLE: {config['site']}\") end'''"
            ),
            'save = { "session.lease" = ".body.sessionLease", "session.contract_sha256" = ".body.contractSha256" }',
            "",
        ]
    )
    if editor == "slide":
        lines.extend(
            [
                "[actions.import_local_image]",
                'description = "把 PPTX_IMAGE_PATH 指向的本地 PNG/JPEG/GIF/WebP/SVG 安全导入当前在线演示文稿，返回可供图片与背景工具复用的 relayAsset source"',
                'method = "POST"',
                f"path = {document_path_source('images/import')}",
                'headers = { "X-AI-Session-Lease" = { from = "state", scope = "chat", key = "session.lease" } }',
                'body = { source = { type = { from = "literal", value = "dataUrl" }, dataUrl = '
                + local_image_data_url_source()
                + " } }",
                "expect_status = 200",
                'extract_type = "jq"',
                'extract_expr = \'\'\'.body.asset | {source:{type:"relayAsset",path:.path,assetToken:.assetToken,assetId:.assetId,mimeType:.mimeType,widthPx:.widthPx,heightPx:.heightPx},expiresAt:.expiresAt}\'\'\'',
                "",
            ]
        )
    lines.extend(
        [
            "[actions.get_state]",
            f'description = "读取当前 {label} bridge state 和契约身份"',
            'method = "POST"',
            f"path = {document_path_source('execute')}",
            'headers = { "X-AI-Session-Lease" = { from = "state", scope = "chat", key = "session.lease" } }',
            'body = { method = "getState", requestId = { from = "param", key = "request_id" }, timeoutMs = { from = "param", key = "timeout_ms", default = 30000 } }',
            "expect_status = 200",
            "params = [",
            *action_params("control", action_name="get_state"),
            "]",
            'extract_type = "jq"',
            'extract_expr = ".body | del(.session)"',
            "",
        ]
    )
    for internal_name, schema in contract["tools"][editor].items():
        name = public_tool_name(editor, internal_name)
        empty_read_arguments = (
            schema.get("x-effects") == "read"
            and not schema_has_required_business_fields(schema)
        )
        arguments_source = '{ from = "param", key = "arguments"'
        if empty_read_arguments:
            arguments_source += ", default = {}"
        arguments_source += " }"
        arguments_note = (
            "无业务参数时可省略 arguments，默认使用空对象"
            if empty_read_arguments
            else "结构化业务参数统一通过 arguments"
        )
        lines.extend(
            [
                f"[actions.{name}]",
                (
                    "description = "
                    + toml_string(
                        f"{schema['description']}；{arguments_note}"
                    )
                ),
                'method = "POST"',
                f"path = {document_path_source('execute')}",
                'headers = { "X-AI-Session-Lease" = { from = "state", scope = "chat", key = "session.lease" } }',
                (
                    'body = { method = "executeTool", '
                    f'name = "{name}", arguments = {arguments_source}, '
                    'requestId = { from = "param", key = "request_id" }, '
                    'timeoutMs = { from = "param", key = "timeout_ms", default = 90000 } }'
                ),
                "expect_status = 200",
                "params = [",
                *action_params(
                    "tool",
                    action_name=name,
                    schema=schema,
                    definitions=contract.get("$defs") or {},
                ),
                "]",
                'extract_type = "jq"',
                'extract_expr = ".body | del(.session)"',
                "",
            ]
        )
    batch_name = "execute_batch"
    validate_name = "validate_batch"
    for action_name, document_action, method, kind in (
        (validate_name, "validate", None, "validate"),
        (batch_name, "execute", "executeBatch", "batch"),
    ):
        body_parts = []
        if method:
            body_parts.append(f'method = "{method}"')
        body_parts.extend(
            [
                'toolCalls = { from = "param", key = "tool_calls" }',
                'requestId = { from = "param", key = "request_id" }',
                'timeoutMs = { from = "param", key = "timeout_ms", default = 90000 }',
            ]
        )
        lines.extend(
            [
                f"[actions.{action_name}]",
                "description = "
                + toml_string(
                    f"{label} 批量{'预检' if kind == 'validate' else '执行'}；"
                    "inspect、validate、execute 必须使用相同 request_id 和逐字相同的 JSON；"
                    "预检失败后只要修改 JSON，就换用 -r1 等新 ID 并重新执行完整链路"
                ),
                'method = "POST"',
                f"path = {document_path_source(document_action)}",
                'headers = { "X-AI-Session-Lease" = { from = "state", scope = "chat", key = "session.lease" } }',
                "body = { " + ", ".join(body_parts) + " }",
                "expect_status = 200",
                "params = [",
                *action_params(kind, action_name=action_name),
                "]",
                'extract_type = "jq"',
                'extract_expr = ".body | del(.session)"',
                "",
            ]
        )
    lines.extend(
        [
            "[actions.commit]",
            f'description = "仅重试 {label} force-save，不重放工具调用；用于消费 mutation receipt"',
            'method = "POST"',
            f"path = {document_path_source('commit')}",
            'headers = { "X-AI-Session-Lease" = { from = "state", scope = "chat", key = "session.lease" } }',
            'body = { mutationReceipt = { from = "param", key = "mutation_receipt" } }',
            "expect_status = 200",
            "params = [",
            '  { name = "mutation_receipt", type = "string", required = true, description = "执行失败响应返回的 mutationReceipt；commit 不会重放业务修改" }',
            "]",
            'extract_type = "jq"',
            'extract_expr = ".body"',
            "",
        ]
    )
    if editor == "word":
        lines.extend(
            [
                "[actions.qa]",
                'description = "强制保存并对 DOCX 执行 OOXML、样式、编号、字段、批注和逐页渲染验收"',
                'method = "POST"',
                f"path = {document_path_source('qa')}",
                'headers = { "X-AI-Session-Lease" = { from = "state", scope = "chat", key = "session.lease" } }',
                'body = { requiredStyles = { from = "param", key = "required_styles", default = [] }, numberingSequences = { from = "param", key = "numbering_sequences", default = [] }, requireSeq = { from = "param", key = "require_seq", default = false }, requireRef = { from = "param", key = "require_ref", default = false }, requireCommentAnchors = { from = "param", key = "require_comment_anchors", default = false }, allowedBlankPages = { from = "param", key = "allowed_blank_pages", default = [] } }',
                "expect_status = 200",
                "params = [",
                '  { name = "required_styles", type = "array", required = false, description = "必须存在并达到最小使用次数的样式要求", example = [] },',
                '  { name = "numbering_sequences", type = "array", required = false, description = "连续编号要求；多级编号可提供与 paragraphIndexes 等长的 levels", example = [] },',
                '  { name = "require_seq", type = "boolean", required = false, description = "要求存在 SEQ 域", example = false },',
                '  { name = "require_ref", type = "boolean", required = false, description = "要求存在 REF 域", example = false },',
                '  { name = "require_comment_anchors", type = "boolean", required = false, description = "要求批注具有有效锚点", example = false },',
                '  { name = "allowed_blank_pages", type = "array", required = false, description = "允许为空白页的一基页码", example = [] }',
                "]",
                'extract_type = "jq"',
                'extract_expr = ".body"',
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
                'headers = { "X-AI-Session-Lease" = { from = "state", scope = "chat", key = "session.lease" } }',
                (
                    f'body = {{ method = "{control}", '
                    'requestId = { from = "param", key = "request_id" }, '
                    'timeoutMs = { from = "param", key = "timeout_ms", default = 90000 } }'
                ),
                "expect_status = 200",
                "params = [",
                *action_params("control", action_name=control),
                "]",
                'extract_type = "jq"',
                'extract_expr = ".body | del(.session)"',
                "",
            ]
        )
    rendered = "\n".join(lines).rstrip() + "\n"
    if 'from = "shell"' in rendered:
        raise ValueError(
            f"generated {config['site']} HTTPX site must remain shell-free"
        )
    return rendered


def build_artifacts(
    contract: dict[str, Any],
    zenmind_root: Path | None,
    httpx_base_url: str | None = None,
    editors: tuple[str, ...] | None = None,
) -> dict[Path, str]:
    sha256 = contract_sha256(contract)
    artifacts = {
        TYPES_PATH: replace_types_error_codes(
            replace_types_contract_version(
                replace_types_region(
                    TYPES_PATH.read_text(encoding="utf-8"),
                    render_types_region(contract),
                ),
                contract,
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
            normalized_url = replace_asset_revision_queries(
                variation["url"],
                ASSET_REVISION_PLACEHOLDER,
            )
            if normalized_url == variation["url"]:
                raise ValueError("every plugin variation URL must include a v query parameter")
            variation["url"] = normalized_url
    normalized_config = json.dumps(config, ensure_ascii=False, indent=2) + "\n"
    normalized_index = replace_asset_revision_queries(
        INDEX_PATH.read_text(encoding="utf-8"),
        ASSET_REVISION_PLACEHOLDER,
    )
    if normalized_index.count(f"v={ASSET_REVISION_PLACEHOLDER}") != 1:
        raise ValueError("index.html must version the ai-bridge bootstrap script")
    revision_sources = runtime_asset_sources(
        contract,
        artifacts,
        normalized_config,
        normalized_index,
    )
    asset_revision = compute_asset_revision(
        contract["version"],
        revision_sources,
    )
    artifacts[CONFIG_PATH] = normalized_config.replace(
        ASSET_REVISION_PLACEHOLDER,
        asset_revision,
    )
    artifacts[INDEX_PATH] = normalized_index.replace(
        ASSET_REVISION_PLACEHOLDER,
        asset_revision,
    )
    if zenmind_root is None:
        return artifacts
    if not zenmind_root.is_dir():
        raise ValueError(f"zenmind root does not exist: {zenmind_root}")
    selected_editors = editors or tuple(EDITOR_CONFIG)
    unknown_editors = sorted(set(selected_editors) - set(EDITOR_CONFIG))
    if unknown_editors:
        raise ValueError(f"unknown editors: {unknown_editors}")
    validate_skill_layout(zenmind_root, selected_editors)
    normalized_httpx_base_url = normalize_httpx_base_url(httpx_base_url)
    for editor in selected_editors:
        config = EDITOR_CONFIG[editor]
        scoped = scoped_contract(contract, editor, sha256)
        skill_root = (
            zenmind_root
            / "skills-center"
            / config["skill"]
        )
        references = skill_root / "references"
        skill_path = skill_root / "SKILL.md"
        # Skill releases are independent from the bridge contract. Keep the
        # authored SKILL.md byte-for-byte instead of projecting contract.version.
        artifacts[skill_path] = skill_path.read_text(encoding="utf-8")
        contract_groups = EDITOR_CONTRACT_GROUPS[editor]
        grouped_names = [
            name
            for names in contract_groups.values()
            for name in names
        ]
        if len(grouped_names) != len(set(grouped_names)):
            raise ValueError(f"{editor} contract groups contain duplicate tools")
        if set(grouped_names) != set(scoped["tools"]):
            missing = sorted(set(scoped["tools"]) - set(grouped_names))
            extra = sorted(set(grouped_names) - set(scoped["tools"]))
            raise ValueError(
                f"{editor} contract groups must cover every tool exactly once; "
                f"missing={missing}, extra={extra}"
            )
        for group_name, tool_names in contract_groups.items():
            grouped = scoped_contract_group(scoped, tool_names)
            stem = f"contract-{group_name}.generated"
            artifacts[references / f"{stem}.json"] = (
                json.dumps(grouped, ensure_ascii=False, indent=2) + "\n"
            )
            artifacts[references / f"{stem}.md"] = render_markdown(
                grouped,
                include_full_schema=False,
            )
        artifacts[
            skill_root / ".config" / "httpx" / config["toml"]
        ] = render_toml(
            contract,
            editor,
            sha256,
            normalized_httpx_base_url,
        )
    return artifacts


def validate_skill_layout(
    zenmind_root: Path,
    editors: tuple[str, ...] | None = None,
) -> None:
    selected_editors = editors or tuple(EDITOR_CONFIG)
    forbidden = [
        zenmind_root / "skills-center" / EDITOR_CONFIG[editor]["skill"] / "scripts"
        for editor in selected_editors
        if (
            zenmind_root
            / "skills-center"
            / EDITOR_CONFIG[editor]["skill"]
            / "scripts"
        ).exists()
    ]
    if forbidden:
        paths = ", ".join(str(path) for path in forbidden)
        raise ValueError(
            "generated online skills must not contain scripts directories: "
            + paths
        )

    missing_internal_boundary = []
    for editor in selected_editors:
        config = EDITOR_CONFIG[editor]
        skill_path = zenmind_root / "skills-center" / config["skill"] / "SKILL.md"
        body = skill_path.read_text(encoding="utf-8")
        if (
            "DocumentServerPublicOrigin" not in body
            or "不得用 `curl`" not in body
            or "重试一次" not in body
        ):
            missing_internal_boundary.append(skill_path)
    if missing_internal_boundary:
        paths = ", ".join(str(path) for path in missing_internal_boundary)
        raise ValueError(
            "generated online skills must keep the internal-service boundary and bounded session retry policy: "
            + paths
        )


def validate_contract(contract: dict[str, Any]) -> None:
    if contract.get("version") != "0.2.6":
        raise ValueError("public-api.json version must be 0.2.6")
    if contract.get("protocolVersion") != 1:
        raise ValueError("protocolVersion must remain 1")
    http_relay = (contract.get("transport") or {}).get("httpRelay") or {}
    expected_transport = {
        "externalBasePath": "/api/v1/editor-relay",
        "internalBasePath": "/copilot-api/bridge/internal",
        "imageBasePath": "/api/v1/editor-relay/images",
        "persistenceBasePath": "/api/v1/editor-relay/persistence",
        "requiresExplicitConfiguration": True,
    }
    if any(http_relay.get(key) != value for key, value in expected_transport.items()):
        raise ValueError("public-api.json must describe the document-hub v1 Relay boundary")
    if "httpRelaySchemas" in contract:
        raise ValueError("public-api.json must not expose legacy HTTP Relay schemas")
    errors = set(contract.get("errors") or [])
    required_runtime_errors = {
        "CONTRACT_MISMATCH",
        "DOCUMENT_LOAD_TIMEOUT",
        "EDITOR_CAPABILITY_PROBE_FAILED",
        "EDITOR_COMMAND_FAILED",
        "EDITOR_COMMAND_REJECTED",
        "EDITOR_APP_STARTUP_TIMEOUT",
        "HTTP_RELAY_INVALID_RESPONSE",
        "HTTP_RELAY_NETWORK_ERROR",
        "MESSAGE_NOT_SUPPORTED",
        "PERSISTENCE_NOT_AVAILABLE",
        "PERSISTENCE_INVALID_RESPONSE",
        "PERSISTENCE_TIMEOUT",
        "PLUGIN_ASSET_LOAD_FAILED",
        "PLUGIN_BRIDGE_MISSING",
        "PLUGIN_HANDSHAKE_CONFIG_INVALID",
        "PLUGIN_INITIALIZATION_FAILED",
        "PLUGIN_NOT_LOADED",
        "UNSUPPORTED_EDITOR_TYPE",
    }
    if not required_runtime_errors.issubset(errors):
        raise ValueError("public-api.json must expose current browser runtime errors")
    if {"MULTIPLE_ACTIVE_EDITORS", "SESSION_NOT_FOUND"} & errors:
        raise ValueError("public-api.json still exposes legacy Relay errors")
    naming = contract.get("toolNaming") or {}
    expected_prefixes = {
        editor: config["prefix"]
        for editor, config in EDITOR_CONFIG.items()
    }
    if (
        naming.get("scope") != "editorType"
        or naming.get("publicNames") != "unprefixed"
        or naming.get("internalPrefixes") != expected_prefixes
    ):
        raise ValueError("public-api.json has an invalid toolNaming policy")
    names = []
    for editor in ("word", "slide", "cell"):
        schemas = contract.get("tools", {}).get(editor)
        if not isinstance(schemas, dict) or not schemas:
            raise ValueError(f"missing tools.{editor}")
        public_names = set()
        for name, schema in schemas.items():
            names.append(name)
            public_name = public_tool_name(editor, name)
            if public_name in public_names:
                raise ValueError(
                    f"duplicate public tool name for {editor}: {public_name}"
                )
            public_names.add(public_name)
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
    if len(names) != 160 or len(names) != len(set(names)):
        raise ValueError(
            f"expected 160 unique tools, found {len(names)} total/{len(set(names))} unique"
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
    parser.add_argument(
        "--editor",
        action="append",
        choices=tuple(EDITOR_CONFIG),
        help="limit zenmind skill/TOML projection to one editor; repeatable",
    )
    args = parser.parse_args()
    contract = json.loads(CONTRACT_PATH.read_text(encoding="utf-8"))
    validate_contract(contract)
    try:
        artifacts = build_artifacts(
            contract,
            args.zenmind_root,
            args.httpx_base_url,
            tuple(args.editor) if args.editor else None,
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
