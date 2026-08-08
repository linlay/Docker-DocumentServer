#!/usr/bin/env python3
"""Contract-source, generated projection, normalization, and limit tests."""

from __future__ import annotations

import copy
import importlib.util
import json
import re
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from typing import Any


BRIDGE_ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(BRIDGE_ROOT))

import copilot_server  # noqa: E402

GENERATOR_SPEC = importlib.util.spec_from_file_location(
    "sync_contract",
    BRIDGE_ROOT / "tools" / "sync_contract.py",
)
assert GENERATOR_SPEC is not None and GENERATOR_SPEC.loader is not None
sync_contract = importlib.util.module_from_spec(GENERATOR_SPEC)
GENERATOR_SPEC.loader.exec_module(sync_contract)


class ContractGenerationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.contract = json.loads(
            (BRIDGE_ROOT / "public-api.json").read_text(encoding="utf-8")
        )

    def test_contract_has_155_complete_tool_metadata_entries(self) -> None:
        tools = {
            name: schema
            for editor_tools in self.contract["tools"].values()
            for name, schema in editor_tools.items()
        }
        self.assertEqual(len(tools), 155)
        self.assertEqual(self.contract["version"], "0.1.0")
        self.assertEqual(self.contract["protocolVersion"], 1)
        for name, schema in tools.items():
            with self.subTest(tool=name):
                self.assertTrue(schema["description"])
                self.assertIn(schema["x-effects"], {"read", "write"})
                self.assertIn(
                    schema["x-argumentLimitClass"],
                    {"standard", "image"},
                )
                self.assertIsInstance(schema["x-semanticValidators"], list)

    def test_slides_add_slide_exposes_generic_title_placement(self) -> None:
        schema = self.contract["tools"]["slide"]["slides_add_slide"]
        title_placement = schema["properties"]["titlePlacement"]
        self.assertEqual(
            title_placement["enum"],
            ["auto", "placeholder", "textbox"],
        )
        self.assertEqual(title_placement["default"], "auto")

    def test_slides_image_background_contract_is_safe_and_mutually_exclusive(self) -> None:
        slide_schema = self.contract["tools"]["slide"]["slides_set_background"]
        template_schema = self.contract["tools"]["slide"][
            "slides_set_template_background"
        ]
        for schema in (slide_schema, template_schema):
            with self.subTest(tool=schema["description"]):
                self.assertEqual(schema["x-argumentLimitClass"], "image")
                self.assertEqual(
                    schema["properties"]["source"]["$ref"],
                    "#/$defs/imageSource",
                )
                self.assertEqual(
                    schema["properties"]["fillMode"]["enum"],
                    ["stretch", "tile"],
                )
                self.assertEqual(
                    schema["properties"]["fillMode"]["default"],
                    "stretch",
                )
                self.assertIn("image", schema["properties"]["mode"]["enum"])

        valid_calls = [
            {
                "name": "slides_set_background",
                "arguments": {
                    "slide": 1,
                    "mode": "image",
                    "source": {
                        "type": "url",
                        "url": "https://images.example.test/background.png",
                    },
                },
            },
            {
                "name": "slides_set_template_background",
                "arguments": {
                    "scope": "master",
                    "masterIndex": 1,
                    "mode": "image",
                    "fillMode": "tile",
                    "source": {
                        "type": "dataUrl",
                        "dataUrl": "data:image/png;base64,AAAA",
                    },
                },
            },
            {
                "name": "slides_set_background",
                "arguments": {
                    "slide": 1,
                    "fill": {"type": "solid", "color": "#FFFFFF"},
                },
            },
            {
                "name": "slides_set_template_background",
                "arguments": {
                    "scope": "layout",
                    "masterIndex": 1,
                    "layoutIndex": 1,
                    "mode": "master",
                },
            },
        ]
        normalized, changes = copilot_server.require_valid_editor_tool_calls(
            "slide",
            valid_calls,
        )
        self.assertEqual(normalized, valid_calls)
        self.assertEqual(changes, [])

        invalid_arguments = [
            {"slide": 1, "mode": "image"},
            {
                "slide": 1,
                "mode": "image",
                "source": {
                    "type": "url",
                    "url": "https://images.example.test/background.png",
                },
                "fill": {"type": "solid", "color": "#FFFFFF"},
            },
            {
                "slide": 1,
                "mode": "custom",
                "fill": {"type": "solid", "color": "#FFFFFF"},
                "source": {
                    "type": "url",
                    "url": "https://images.example.test/background.png",
                },
            },
            {
                "slide": 1,
                "mode": "clear",
                "fill": {"type": "solid", "color": "#FFFFFF"},
            },
            {
                "slide": 1,
                "fill": {"type": "solid", "color": "#FFFFFF"},
                "fillMode": "tile",
            },
        ]
        for arguments in invalid_arguments:
            with self.subTest(arguments=arguments):
                with self.assertRaises(copilot_server.BridgeError) as caught:
                    copilot_server.require_valid_editor_tool_calls(
                        "slide",
                        [
                            {
                                "name": "slides_set_background",
                                "arguments": arguments,
                            }
                        ],
                    )
                self.assertEqual(caught.exception.code, "INVALID_TOOL_ARGUMENTS")
                self.assertEqual(
                    caught.exception.details["completedToolCalls"],
                    0,
                )
                self.assertFalse(
                    caught.exception.details["partialMutationPossible"]
                )

    def test_local_generated_projections_are_current_and_idempotent(self) -> None:
        command = [
            sys.executable,
            str(BRIDGE_ROOT / "tools" / "sync_contract.py"),
            "--check",
        ]
        first = subprocess.run(
            command,
            cwd=BRIDGE_ROOT,
            capture_output=True,
            text=True,
            check=False,
        )
        second = subprocess.run(
            command,
            cwd=BRIDGE_ROOT,
            capture_output=True,
            text=True,
            check=False,
        )
        self.assertEqual(first.returncode, 0, first.stderr)
        self.assertEqual(second.returncode, 0, second.stderr)
        self.assertEqual(first.stdout, second.stdout)

    def test_httpx_toml_uses_anonymous_document_hub_transport(
        self,
    ) -> None:
        base_url = "https://office.example.test/"
        sha256 = sync_contract.contract_sha256(self.contract)
        expected_session_actions = {
            "word": "word_session",
            "slide": "slides_session",
            "cell": "sheets_session",
        }
        for editor, session_action in expected_session_actions.items():
            with self.subTest(editor=editor):
                rendered = sync_contract.render_toml(
                    self.contract,
                    editor,
                    sha256,
                    base_url,
                )
                self.assertIn(
                    'base_url = "https://office.example.test"',
                    rendered,
                )
                self.assertIn(f"[actions.{session_action}]", rendered)
                self.assertNotIn("Authorization", rendered)
                self.assertNotIn("AP_ACCESS_TOKEN", rendered)
                self.assertNotIn("X-HTTPX-Direct-Secret", rendered)
                self.assertIn("document_id=$DOCUMENT_HUB_DOCUMENT_ID", rendered)
                self.assertIn('path = "/api/v1/documents"', rendered)
                self.assertIn("/ai/session", rendered)
                self.assertIn("/ai/validate", rendered)
                self.assertIn("/ai/execute", rendered)
                self.assertNotIn("/copilot-api/bridge/attach", rendered)
                self.assertNotIn("auth.bridge", rendered)
                self.assertNotIn("bindingToken", rendered)
                site = sync_contract.EDITOR_CONFIG[editor]["site"]
                self.assertIn(
                    f'User-Agent = "agent-platform-httpx/{site}"',
                    rendered,
                )
                self.assertIn(f"EDITOR_SESSION_UNAVAILABLE: {site}", rendered)

    def test_skill_version_projection_preserves_frontmatter_and_body(self) -> None:
        source = (
            "---\n"
            "name: online-docx\n"
            "description: Example\n"
            "metadata:\n"
            "  tags:\n"
            "    - office\n"
            "  version: \"9.9.9\"\n"
            "---\n\n"
            "# Body\n"
        )
        expected = source.replace('version: "9.9.9"', 'version: "0.1.0"')
        rendered = sync_contract.render_skill_version(source, "0.1.0")
        self.assertEqual(rendered, expected)
        self.assertEqual(
            sync_contract.render_skill_version(rendered, "0.1.0"),
            expected,
        )

        without_metadata = (
            "---\n"
            "name: online-xlsx\n"
            "description: Example\n"
            "---\n\n"
            "# Body\n"
        )
        self.assertEqual(
            sync_contract.render_skill_version(without_metadata, "0.1.0"),
            "---\n"
            "name: online-xlsx\n"
            "description: Example\n"
            "metadata:\n"
            "  version: \"0.1.0\"\n"
            "---\n\n"
            "# Body\n",
        )

    def test_cross_repo_projections_live_inside_each_skill(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            zenmind_root = Path(directory)
            for config in sync_contract.EDITOR_CONFIG.values():
                skill_root = (
                    zenmind_root / "skills-center" / config["skill"]
                )
                skill_root.mkdir(parents=True)
                (skill_root / "SKILL.md").write_text(
                    "---\n"
                    f'name: {config["skill"]}\n'
                    "description: Example\n"
                    "---\n\n"
                    "# Body\n",
                    encoding="utf-8",
                )

            artifacts = sync_contract.build_artifacts(
                self.contract,
                zenmind_root,
                "https://office.example.test",
            )
            self.assertFalse(
                any(
                    path.is_relative_to(zenmind_root / "agents")
                    for path in artifacts
                )
            )
            for config in sync_contract.EDITOR_CONFIG.values():
                skill_root = (
                    zenmind_root / "skills-center" / config["skill"]
                )
                skill_path = skill_root / "SKILL.md"
                toml_path = (
                    skill_root / ".config" / "httpx" / config["toml"]
                )
                self.assertIn(skill_path, artifacts)
                self.assertIn(toml_path, artifacts)
                self.assertIn(
                    '  version: "0.1.0"',
                    artifacts[skill_path],
                )
                self.assertIn(
                    f'agent-platform-httpx/{config["site"]}',
                    artifacts[toml_path],
                )

    def test_httpx_base_url_rejects_non_origin_values(self) -> None:
        invalid_values = (
            None,
            "",
            "office.example.test",
            "https://user@office.example.test",
            "https://office.example.test/path",
            "https://office.example.test?query=1",
            "https://office.example.test#fragment",
        )
        for value in invalid_values:
            with self.subTest(value=value), self.assertRaises(ValueError):
                sync_contract.normalize_httpx_base_url(value)

    def test_check_logic_detects_a_manual_generated_file_edit(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            projection = Path(directory) / "contract.generated.json"
            expected = '{"version":"0.1.0"}\n'
            projection.write_text(expected, encoding="utf-8")
            self.assertEqual(
                sync_contract.drifted_paths({projection: expected}),
                [],
            )
            projection.write_text(
                expected.replace("0.1.0", "manual-edit"),
                encoding="utf-8",
            )
            self.assertEqual(
                sync_contract.drifted_paths({projection: expected}),
                [projection],
            )

    def test_plugin_schema_projection_is_resolved_and_compact(self) -> None:
        schema = self.contract["tools"]["word"]["word_add_image"]
        resolved = sync_contract.resolve_schema(
            schema,
            self.contract["$defs"],
        )
        compacted = sync_contract.compact_plugin_schema(resolved)

        self.assertIn("source", compacted["properties"])
        self.assertIn("oneOf", compacted["properties"]["source"])
        self.assertNotIn("$ref", json.dumps(compacted, ensure_ascii=False))
        self.assertNotIn("description", compacted)
        self.assertNotIn("x-effects", compacted)

    def test_editor_contracts_are_self_contained_and_identified(self) -> None:
        def assert_refs_resolve(value: Any, definitions: dict[str, Any]) -> None:
            if isinstance(value, list):
                for item in value:
                    assert_refs_resolve(item, definitions)
                return
            if not isinstance(value, dict):
                return
            reference = value.get("$ref")
            if isinstance(reference, str):
                self.assertTrue(reference.startswith("#/$defs/"))
                self.assertIn(reference.removeprefix("#/$defs/"), definitions)
            for child in value.values():
                assert_refs_resolve(child, definitions)

        for editor in ("word", "slide", "cell"):
            scoped = copilot_server.editor_public_contract(editor)
            with self.subTest(editor=editor):
                self.assertEqual(
                    scoped["version"],
                    copilot_server.PUBLIC_API_CONTRACT_VERSION,
                )
                self.assertEqual(
                    scoped["contractSha256"],
                    copilot_server.PUBLIC_API_CONTRACT_SHA256,
                )
                self.assertIn("$defs", scoped)
                self.assertIn("tools", scoped)
                self.assertIn("limits", scoped)
                self.assertIn("unitConventions", scoped)
                self.assertIn("inputNormalization", scoped)
                self.assertIn("controls", scoped)
                self.assertIn("errors", scoped)
                assert_refs_resolve(scoped["tools"], scoped["$defs"])

    def test_unknown_semantic_validator_is_a_startup_error(self) -> None:
        schema = copilot_server.ARGUMENT_SCHEMAS_BY_EDITOR["word"][
            "word_inspect"
        ]
        original = copy.deepcopy(schema.get("x-semanticValidators"))
        schema["x-semanticValidators"] = ["unknown.validator"]
        try:
            with self.assertRaisesRegex(RuntimeError, "unknown.validator"):
                copilot_server.validate_semantic_validator_registry()
        finally:
            schema["x-semanticValidators"] = original

    def test_sheet_enum_normalization_uses_contract_policy(self) -> None:
        calls = [
            {
                "name": "sheets_filter",
                "arguments": {
                    "action": "set",
                    "range": "A1:B2",
                    "operator": "AND",
                },
            }
        ]
        normalized, changes = copilot_server.require_valid_editor_tool_calls(
            "cell",
            calls,
        )
        self.assertEqual(normalized[0]["arguments"]["operator"], "xlAnd")
        self.assertEqual(changes[0]["kind"], "enumAlias")

        canonicalized, changes = copilot_server.require_valid_editor_tool_calls(
            "cell",
            [
                {
                    "name": "sheets_manage_validation",
                    "arguments": {
                        "action": "add",
                        "range": "A1",
                        "type": "XL-VALIDATE-LIST",
                        "formula1": "A,B",
                    },
                }
            ],
        )
        self.assertEqual(
            canonicalized[0]["arguments"]["type"],
            "xlValidateList",
        )
        self.assertEqual(changes[0]["kind"], "enumCanonicalization")

    def test_chart_type_contract_accepts_supported_values_and_rejects_unknowns(self) -> None:
        supported = self.contract["$defs"]["chartType"]["enum"]
        self.assertIn("bar", supported)
        self.assertIn("column", supported)
        self.assertNotIn("columnClustered", supported)
        for chart_type in supported:
            with self.subTest(chart_type=chart_type):
                normalized, _ = copilot_server.require_valid_editor_tool_calls(
                    "cell",
                    [
                        {
                            "name": "sheets_add_chart",
                            "arguments": {
                                "range": "A1:B3",
                                "type": chart_type,
                            },
                        }
                    ],
                )
                self.assertEqual(
                    normalized[0]["arguments"]["type"],
                    chart_type,
                )

        with self.assertRaises(copilot_server.BridgeError) as caught:
            copilot_server.require_valid_editor_tool_calls(
                "cell",
                [
                    {
                        "name": "sheets_add_chart",
                        "arguments": {
                            "range": "A1:B3",
                            "type": "columnClustered",
                        },
                    }
                ],
            )
        self.assertEqual(caught.exception.status, 422)
        self.assertEqual(caught.exception.code, "INVALID_TOOL_ARGUMENTS")
        validation_error = caught.exception.details["validationErrors"][0]
        self.assertEqual(validation_error["path"], "arguments.type")
        self.assertEqual(validation_error["received"], "columnClustered")
        self.assertEqual(validation_error["allowedValues"], supported)
        self.assertEqual(validation_error["suggestedValues"], ["column"])
        self.assertEqual(caught.exception.details["completedToolCalls"], 0)
        self.assertFalse(caught.exception.details["partialMutationPossible"])

    def test_sheet_chart_runtime_guard_matches_contract(self) -> None:
        chart_type = self.contract["$defs"]["chartType"]
        aliases = chart_type["x-aliases"]
        expected_canonicals = [
            value for value in chart_type["enum"] if value not in aliases
        ]
        bridge_source = (BRIDGE_ROOT / "bridges" / "sheets-bridge.js").read_text(
            encoding="utf-8"
        )
        canonical_block = re.search(
            r"var CHART_TYPE_CANONICALS = \[(.*?)\n\s*\];",
            bridge_source,
            re.DOTALL,
        )
        alias_block = re.search(
            r"var CHART_TYPE_ALIASES = \{(.*?)\n\s*\};",
            bridge_source,
            re.DOTALL,
        )
        self.assertIsNotNone(canonical_block)
        self.assertIsNotNone(alias_block)
        runtime_canonicals = re.findall(r'"([^"]+)"', canonical_block.group(1))
        runtime_aliases = dict(
            re.findall(
                r"([A-Za-z][A-Za-z0-9]*):\s*\"([^\"]+)\"",
                alias_block.group(1),
            )
        )
        self.assertEqual(runtime_canonicals, expected_canonicals)
        self.assertEqual(runtime_aliases, aliases)

    def test_standard_and_image_argument_limits_share_one_rule(self) -> None:
        standard_calls = [
            {
                "name": "word_set_document_text",
                "arguments": {"text": "x" * 250_001},
            }
        ]
        with self.assertRaisesRegex(
            copilot_server.BridgeError,
            "250000",
        ):
            copilot_server.require_arguments_within_limit(
                standard_calls,
                standard_calls,
                "toolCalls",
            )

        image_calls = [
            {
                "name": "word_add_image",
                "arguments": {"source": "x" * 300_000},
            }
        ]
        copilot_server.require_arguments_within_limit(
            image_calls,
            image_calls,
            "toolCalls",
        )

        oversized_image_calls = [
            {
                "name": "word_add_image",
                "arguments": {"source": "x" * 12_000_001},
            }
        ]
        with self.assertRaisesRegex(
            copilot_server.BridgeError,
            "12000000",
        ):
            copilot_server.require_arguments_within_limit(
                oversized_image_calls,
                oversized_image_calls,
                "toolCalls",
            )

    def test_schema_conditionals_are_enforced(self) -> None:
        with self.assertRaises(copilot_server.BridgeError) as caught:
            copilot_server.require_valid_editor_tool_calls(
                "word",
                [
                    {
                        "name": "word_manage_section",
                        "arguments": {"action": "create"},
                    }
                ],
            )
        validation_errors = caught.exception.details["validationErrors"]
        self.assertTrue(
            any(
                error["path"] == "arguments.paragraphIndex"
                and error["keyword"] == "required"
                for error in validation_errors
            )
        )

    def test_relay_validation_matches_direct_plugin_regression_shape(self) -> None:
        with self.assertRaises(copilot_server.BridgeError) as caught:
            copilot_server.require_valid_editor_tool_calls(
                "word",
                [
                    {
                        "name": "word_append_paragraph",
                        "arguments": {
                            "text": 123,
                            "internalNote": "must-not-leak",
                        },
                    },
                    {
                        "name": "word_manage_section",
                        "arguments": {"action": "create"},
                    },
                    {
                        "name": "word_insert_page_break",
                        "arguments": {},
                    },
                ],
            )

        validation_errors = caught.exception.details["validationErrors"]
        self.assertEqual(
            [
                (
                    error["toolCallIndex"],
                    error["path"],
                    error["keyword"],
                )
                for error in validation_errors
            ],
            [
                (0, "arguments.internalNote", "additionalProperties"),
                (0, "arguments.text", "type"),
                (1, "arguments.paragraphIndex", "required"),
                (2, "arguments.target", "semantic"),
            ],
        )
        self.assertNotIn("must-not-leak", json.dumps(validation_errors))


if __name__ == "__main__":
    unittest.main()
