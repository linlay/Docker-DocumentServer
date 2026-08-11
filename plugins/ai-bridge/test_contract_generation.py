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

    def test_contract_has_160_complete_tool_metadata_entries(self) -> None:
        tools = {
            name: schema
            for editor_tools in self.contract["tools"].values()
            for name, schema in editor_tools.items()
        }
        self.assertEqual(len(tools), 160)
        self.assertEqual(self.contract["version"], "0.2.1")
        self.assertEqual(
            self.contract["toolNaming"],
            {
                "scope": "editorType",
                "publicNames": "unprefixed",
                "internalPrefixes": {
                    "word": "word_",
                    "slide": "slides_",
                    "cell": "sheets_",
                },
            },
        )
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

    def test_slides_contract_exposes_native_smartart_crud(self) -> None:
        slide_tools = self.contract["tools"]["slide"]
        self.assertEqual(
            {
                "slides_inspect_smartarts",
                "slides_add_smartart",
                "slides_update_smartart",
                "slides_delete_smartart",
            }.issubset(slide_tools),
            True,
        )
        smartart_names = self.contract["$defs"]["smartArtType"]["oneOf"][0]["enum"]
        self.assertEqual(len(smartart_names), 151)
        self.assertEqual(smartart_names[77], "BasicBlockList")
        self.assertEqual(smartart_names[150], "VerticalBracketList")
        add_schema = slide_tools["slides_add_smartart"]
        self.assertEqual(add_schema["required"], ["slide", "type"])
        self.assertEqual(add_schema["x-effects"], "write")
        node_schema = self.contract["$defs"]["smartArtNodeUpdate"]
        self.assertIn("nodeId", node_schema["properties"])
        self.assertIn("paragraphs", node_schema["properties"])

    def test_chart_contract_exposes_all_documented_bold_and_point_format_arguments(self) -> None:
        chart_axis = self.contract["$defs"]["chartAxis"]["properties"]
        chart_point = self.contract["$defs"]["chartPointUpdate"]["properties"]
        self.assertEqual(chart_axis["titleBold"], {"type": "boolean"})
        self.assertEqual(chart_point["allSeries"], {"type": "boolean"})
        for editor, tool_names in {
            "slide": ("slides_add_chart", "slides_update_chart"),
            "cell": ("sheets_add_chart", "sheets_update_chart"),
        }.items():
            for tool_name in tool_names:
                with self.subTest(tool=tool_name):
                    self.assertEqual(
                        self.contract["tools"][editor][tool_name]["properties"]["titleBold"],
                        {"type": "boolean"},
                    )

    def test_slides_doughnut_hole_size_contract_is_strict_and_optional(self) -> None:
        slide_tools = self.contract["tools"]["slide"]
        expected = {
            "type": "integer",
            "minimum": 10,
            "maximum": 90,
            "description": "环形图圆孔占直径的百分比；数值越小环越粗，仅适用于 doughnut",
        }
        self.assertEqual(
            slide_tools["slides_add_chart"]["properties"]["holeSizePercent"],
            expected,
        )
        self.assertEqual(
            slide_tools["slides_update_chart"]["properties"]["holeSizePercent"],
            expected,
        )
        self.assertIn("INVALID_TARGET", self.contract["errors"])
        self.assertIn("SLIDES_API_UNSUPPORTED", self.contract["errors"])
        self.assertIn(
            'holeSizePercent: number; type: "doughnut";',
            sync_contract.render_types_region(self.contract),
        )
        generated_types = sync_contract.replace_types_error_codes(
            'export type AiBridgeErrorCode =\n  | "OLD"\n  | string;',
            self.contract,
        )
        self.assertIn('| "INVALID_TARGET"', generated_types)
        self.assertIn('| "SLIDES_API_UNSUPPORTED"', generated_types)
        self.assertNotIn('| "OLD"', generated_types)

        for value in (10, 40, 90):
            calls, changes = copilot_server.require_valid_editor_tool_calls(
                "slide",
                [{
                    "name": "slides_add_chart",
                    "arguments": {
                        "slide": 1,
                        "type": "doughnut",
                        "series": [[1]],
                        "seriesNames": ["One"],
                        "categories": ["A"],
                        "holeSizePercent": value,
                    },
                }],
            )
            self.assertEqual(calls[0]["arguments"]["holeSizePercent"], value)
            self.assertEqual(changes, [])

        for value in (9, 91, 40.5, "40"):
            with self.subTest(invalid=value):
                with self.assertRaises(copilot_server.BridgeError) as caught:
                    copilot_server.require_valid_editor_tool_calls(
                        "slide",
                        [{
                            "name": "slides_add_chart",
                            "arguments": {
                                "slide": 1,
                                "type": "doughnut",
                                "series": [[1]],
                                "seriesNames": ["One"],
                                "categories": ["A"],
                                "holeSizePercent": value,
                            },
                        }],
                    )
                self.assertEqual(caught.exception.code, "INVALID_TOOL_ARGUMENTS")

        with self.assertRaises(copilot_server.BridgeError) as caught:
            copilot_server.require_valid_editor_tool_calls(
                "slide",
                [{
                    "name": "slides_add_chart",
                    "arguments": {
                        "slide": 1,
                        "type": "bar",
                        "series": [[1]],
                        "seriesNames": ["One"],
                        "categories": ["A"],
                        "holeSizePercent": 40,
                    },
                }],
            )
        self.assertEqual(caught.exception.code, "INVALID_TOOL_ARGUMENTS")

    def test_batch_envelope_accepts_action_alias_and_rejects_ambiguous_shapes(self) -> None:
        canonical = copilot_server.internalize_public_tool_calls(
            "slide",
            [{"name": "set_size", "arguments": {"preset": "wide"}}],
        )
        self.assertEqual(
            canonical,
            [{"name": "slides_set_size", "arguments": {"preset": "wide"}}],
        )
        alias = copilot_server.internalize_public_tool_calls(
            "slide",
            [{"action": "set_size", "arguments": {"preset": "wide"}}],
        )
        self.assertEqual(alias, canonical)
        both = copilot_server.internalize_public_tool_calls(
            "slide",
            [{"name": "inspect", "action": "inspect"}],
        )
        self.assertEqual(both, [{"name": "slides_inspect", "arguments": {}}])

        invalid_calls = (
            ({}, "toolCalls[0].name"),
            ({"name": ""}, "toolCalls[0].name"),
            ({"action": 42}, "toolCalls[0].action"),
            ({"name": "inspect", "action": "set_size"}, "toolCalls[0].action"),
            ({"name": "inspect", "arguments": []}, "toolCalls[0].arguments"),
        )
        for call, expected_path in invalid_calls:
            with self.subTest(call=call):
                with self.assertRaises(copilot_server.BridgeError) as caught:
                    copilot_server.internalize_public_tool_calls("slide", [call])
                self.assertEqual(caught.exception.status, 400)
                self.assertEqual(caught.exception.code, "INVALID_TOOL_CALL")
                self.assertEqual(caught.exception.details["toolCallIndex"], 0)
                self.assertEqual(caught.exception.details["path"], expected_path)
                self.assertIn('"name":"set_size"', caught.exception.details["expectedShape"])

        with self.assertRaises(copilot_server.BridgeError) as caught:
            copilot_server.internalize_public_tool_calls(
                "slide",
                [{"action": "definitely_unknown"}],
            )
        self.assertEqual(caught.exception.code, "TOOL_NOT_ALLOWED")
        self.assertIn("definitely_unknown", caught.exception.message)

    def test_sheets_runtime_capability_contract_is_granular(self) -> None:
        sheets = self.contract["runtimeCapabilities"]["cell"]["features"]["sheets"]
        self.assertEqual(
            sheets["rangeFill"],
            {
                "fillDown": "boolean",
                "fillUp": "boolean",
                "fillLeft": "boolean",
                "fillRight": "boolean",
            },
        )
        self.assertEqual(sheets["arrayFormula"], {"set": "boolean"})
        self.assertEqual(sheets["validation"], {"manage": "boolean"})
        self.assertEqual(
            sheets["comments"],
            {
                "create": "boolean",
                "inspect": "boolean",
                "update": "boolean",
                "delete": "boolean",
            },
        )
        self.assertEqual(
            sheets["freezePanes"],
            {"inspect": "boolean", "manage": "boolean"},
        )
        self.assertEqual(
            sheets["charts"],
            {
                "create": "boolean",
                "inspect": "boolean",
                "update": "boolean",
                "delete": "boolean",
                "addSeriesOnCreate": "boolean",
            },
        )

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
                    "mode": "image",
                    "source": {
                        "type": "relayAsset",
                        "path": "/api/v1/editor-relay/images/"
                        + "a" * 64
                        + ".png",
                        "assetToken": "leased-token",
                        "assetId": "a" * 64 + ".png",
                        "mimeType": "image/png",
                        "widthPx": 1600,
                        "heightPx": 900,
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
                "mode": "image",
                "source": {
                    "type": "relayAsset",
                    "path": "/api/v1/editor-relay/images/"
                    + "a" * 64
                    + ".png",
                    "mimeType": "image/png",
                    "widthPx": 1600,
                    "heightPx": 900,
                },
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

    def test_asset_revision_is_content_addressed_and_order_independent(self) -> None:
        first = sync_contract.compute_asset_revision(
            "0.2.1",
            {"b.js": "second\n", "a.js": "first\n"},
        )
        reordered = sync_contract.compute_asset_revision(
            "0.2.1",
            {"a.js": "first\n", "b.js": "second\n"},
        )
        changed = sync_contract.compute_asset_revision(
            "0.2.1",
            {"a.js": "first\n", "b.js": "changed\n"},
        )

        self.assertEqual(first, reordered)
        self.assertNotEqual(first, changed)
        self.assertRegex(first, r"^0\.2\.1-[0-9a-f]{64}$")

    def test_runtime_asset_revision_covers_every_browser_bridge(self) -> None:
        artifacts = sync_contract.build_artifacts(self.contract, None)
        config = json.loads(artifacts[sync_contract.CONFIG_PATH])
        revision = config["variations"][0]["url"].split("?v=", 1)[1]
        index_revisions = re.findall(
            r"[?&]v=([^&\"'\s<>]+)",
            artifacts[sync_contract.INDEX_PATH],
        )
        copilot_match = re.search(
            r'^EDITOR_ASSET_REVISION = "([^"]+)"$',
            artifacts[sync_contract.COPILOT_SERVER_PATH],
            re.MULTILINE,
        )

        self.assertEqual(index_revisions, [revision] * 4)
        self.assertIsNotNone(copilot_match)
        self.assertEqual(copilot_match.group(1), revision)
        self.assertRegex(revision, r"^0\.2\.1-[0-9a-f]{64}$")

        normalized_config = sync_contract.replace_asset_revision_queries(
            artifacts[sync_contract.CONFIG_PATH],
            sync_contract.ASSET_REVISION_PLACEHOLDER,
        )
        normalized_index = sync_contract.replace_asset_revision_queries(
            artifacts[sync_contract.INDEX_PATH],
            sync_contract.ASSET_REVISION_PLACEHOLDER,
        )
        sources = sync_contract.runtime_asset_sources(
            self.contract,
            artifacts,
            normalized_config,
            normalized_index,
        )
        self.assertEqual(
            set(sources),
            {
                "public-api.json",
                "plugin.js",
                "host-bridge.js",
                "client-sdk.js",
                "config.json",
                "index.html",
                "bridges/word-bridge.js",
                "bridges/slides-bridge.js",
                "bridges/sheets-bridge.js",
                "editor-shell.js",
                "local-guest.js",
            },
        )
        changed_sources = dict(sources)
        changed_sources["bridges/sheets-bridge.js"] += "\n// changed"
        self.assertNotEqual(
            sync_contract.compute_asset_revision(self.contract["version"], sources),
            sync_contract.compute_asset_revision(
                self.contract["version"],
                changed_sources,
            ),
        )

    def test_manual_revision_tokens_are_replaced_atomically(self) -> None:
        source = (
            '<script src="plugin.js?v=0.2.1-rev2"></script>\n'
            '<script src="bridges/sheets-bridge.js?v=stale"></script>\n'
        )
        rendered = sync_contract.replace_asset_revision_queries(
            source,
            "0.2.1-" + "a" * 64,
        )
        self.assertNotIn("rev2", rendered)
        self.assertNotIn("v=stale", rendered)
        self.assertEqual(rendered.count("v=0.2.1-" + "a" * 64), 2)

    def test_httpx_toml_uses_anonymous_document_hub_transport(
        self,
    ) -> None:
        base_url = "https://office.example.test/"
        sha256 = sync_contract.contract_sha256(self.contract)
        expected_session_actions = {editor: "session" for editor in ("word", "slide", "cell")}
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
                self.assertIn("[actions.create]", rendered)
                self.assertIn("[actions.validate_batch]", rendered)
                self.assertIn("[actions.execute_batch]", rendered)
                self.assertIn('key = "arguments_json"', rendered)
                self.assertIn('key = "tool_calls_json"', rendered)
                self.assertIn('argumentsJson = { from = "param"', rendered)
                self.assertIn('toolCallsJson = { from = "param"', rendered)
                inspect_action = rendered.split("[actions.inspect]", 1)[1].split(
                    "\n[actions.", 1
                )[0]
                self.assertIn(
                    'argumentsJson = { from = "param", key = "arguments_json", default = "{}" }',
                    inspect_action,
                )
                self.assertIn(
                    'name = "arguments_json", type = "json object string", required = false',
                    inspect_action,
                )
                self.assertIn('example = "{}"', inspect_action)
                self.assertIn('example = "tool-inspect-1"', inspect_action)
                batch_action = rendered.split("[actions.validate_batch]", 1)[1].split(
                    "\n[actions.", 1
                )[0]
                self.assertIn(
                    'example = "[{\\"name\\":\\"set_size\\",\\"arguments\\":{\\"preset\\":\\"wide\\"}}]"',
                    batch_action,
                )
                write_tool_name = {
                    "word": "set_document_text",
                    "slide": "add_chart",
                    "cell": "manage_range",
                }[editor]
                write_action = rendered.split(f"[actions.{write_tool_name}]", 1)[1].split(
                    "\n[actions.", 1
                )[0]
                self.assertNotIn('key = "arguments_json", default = "{}"', write_action)
                self.assertIn(
                    'name = "arguments_json", type = "json object string", required = true',
                    write_action,
                )
                self.assertNotIn(f"[actions.new_{sync_contract.EDITOR_CONFIG[editor]['fileType']}]", rendered)
                self.assertNotIn(f"[actions.{sync_contract.EDITOR_CONFIG[editor]['prefix']}", rendered)
                self.assertNotIn("Authorization", rendered)
                self.assertNotIn("AP_ACCESS_TOKEN", rendered)
                self.assertNotIn("X-HTTPX-Direct-Secret", rendered)
                self.assertIn("document_id=$DOCUMENT_HUB_DOCUMENT_ID", rendered)
                self.assertIn('path = "/api/v1/documents"', rendered)
                self.assertIn("/ai/session", rendered)
                self.assertIn("/ai/validate", rendered)
                self.assertIn("/ai/execute", rendered)
                self.assertIn("/ai/commit", rendered)
                self.assertIn("[actions.commit]", rendered)
                self.assertIn(".body.saveReady == true", rendered)
                self.assertIn('"session.lease" = ".body.sessionLease"', rendered)
                self.assertIn(
                    '"X-AI-Session-Lease" = { from = "state", scope = "chat", key = "session.lease" }',
                    rendered,
                )
                self.assertIn('key = "mutation_receipt"', rendered)
                if editor == "word":
                    self.assertIn("[actions.qa]", rendered)
                    self.assertIn("/ai/qa", rendered)
                    self.assertIn('key = "qa_json"', rendered)
                    qa_action = rendered.split("[actions.qa]", 1)[1].split("\n[actions.", 1)[0]
                    self.assertIn('"X-AI-Session-Lease"', qa_action)
                    self.assertIn('key = "session.lease"', qa_action)
                else:
                    self.assertNotIn("[actions.qa]", rendered)
                    self.assertNotIn("/ai/qa", rendered)
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
        expected = source.replace('version: "9.9.9"', 'version: "0.2.0"')
        rendered = sync_contract.render_skill_version(source, "0.2.0")
        self.assertEqual(rendered, expected)
        self.assertEqual(
            sync_contract.render_skill_version(rendered, "0.2.0"),
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
            sync_contract.render_skill_version(without_metadata, "0.2.0"),
            "---\n"
            "name: online-xlsx\n"
            "description: Example\n"
            "metadata:\n"
            "  version: \"0.2.0\"\n"
            "---\n\n"
            "# Body\n",
        )

    def test_cross_repo_projections_live_inside_each_skill(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            zenmind_root = Path(directory)
            for editor, config in sync_contract.EDITOR_CONFIG.items():
                skill_root = (
                    zenmind_root / "skills-center" / config["skill"]
                )
                skill_root.mkdir(parents=True)
                (skill_root / "SKILL.md").write_text(
                    "---\n"
                    f'name: {config["skill"]}\n'
                    "description: Example\n"
                    "---\n\n"
                    "# Body\n\n"
                    "`DocumentServerPublicOrigin` 是内部实现细节，不得用 `curl` 探测。\n"
                    "session 失败时最多重试一次。\n",
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
            for editor, config in sync_contract.EDITOR_CONFIG.items():
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
                    '  version: "0.2.1"',
                    artifacts[skill_path],
                )
                self.assertIn(
                    f'agent-platform-httpx/{config["site"]}',
                    artifacts[toml_path],
                )
                scripts_root = skill_root / "scripts"
                self.assertFalse(
                    any(scripts_root in path.parents for path in artifacts),
                    f"{config['skill']} must not generate scripts artifacts",
                )
                action_names = set(
                    re.findall(
                        r"^\[actions\.([^]]+)\]$",
                        artifacts[toml_path],
                        re.MULTILINE,
                    )
                )
                expected_actions = {
                    "health",
                    "create",
                    "session",
                    "get_state",
                    "validate_batch",
                    "execute_batch",
                    "commit",
                    *self.contract["controls"],
                    *(
                        sync_contract.public_tool_name(editor, name)
                        for name in self.contract["tools"][editor]
                    ),
                }
                if editor == "word":
                    expected_actions.add("qa")
                if editor == "slide":
                    expected_actions.add("import_local_image")
                    self.assertIn("PPTX_IMAGE_PATH", artifacts[toml_path])
                    self.assertIn("/ai/images/import", artifacts[toml_path])
                    self.assertIn('type:"relayAsset"', artifacts[toml_path])
                generated_tools: set[str] = set()
                contract_groups = sync_contract.EDITOR_CONTRACT_GROUPS[editor]
                for group_name, expected_tools in contract_groups.items():
                    json_path = (
                        skill_root
                        / "references"
                        / f"contract-{group_name}.generated.json"
                    )
                    markdown_path = (
                        skill_root
                        / "references"
                        / f"contract-{group_name}.generated.md"
                    )
                    self.assertIn(json_path, artifacts)
                    self.assertIn(markdown_path, artifacts)
                    grouped = json.loads(artifacts[json_path])
                    self.assertEqual(tuple(grouped["tools"]), expected_tools)
                    generated_tools.update(grouped["tools"])
                    self.assertNotIn("<details>", artifacts[markdown_path])
                    references = set(re.findall(
                        r'"#/\$defs/([^"/]+)"',
                        artifacts[json_path],
                    ))
                    self.assertLessEqual(references, set(grouped["$defs"]))
                self.assertEqual(generated_tools, set(
                    sync_contract.public_tool_name(editor, name)
                    for name in self.contract["tools"][editor]
                ))
                self.assertNotIn(
                    skill_root / "references" / "contract.generated.json",
                    artifacts,
                )
                self.assertNotIn(
                    skill_root / "references" / "contract.generated.md",
                    artifacts,
                )
                self.assertEqual(action_names, expected_actions)

    def test_generation_rejects_scripts_directories_for_every_online_skill(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            zenmind_root = Path(directory)
            for config in sync_contract.EDITOR_CONFIG.values():
                skill_root = zenmind_root / "skills-center" / config["skill"]
                skill_root.mkdir(parents=True)
                (skill_root / "SKILL.md").write_text(
                    "---\nname: example\n---\n\n# Body\n",
                    encoding="utf-8",
                )
                scripts_root = skill_root / "scripts"
                scripts_root.mkdir()
                with self.subTest(skill=config["skill"]), self.assertRaisesRegex(
                    ValueError,
                    "must not contain scripts directories",
                ):
                    sync_contract.build_artifacts(
                        self.contract,
                        zenmind_root,
                        "https://office.example.test",
                    )
                scripts_root.rmdir()

    def test_cross_repo_projection_can_be_limited_to_one_editor(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            zenmind_root = Path(directory)
            for editor, config in sync_contract.EDITOR_CONFIG.items():
                skill_root = zenmind_root / "skills-center" / config["skill"]
                skill_root.mkdir(parents=True)
                body = "# Body\n"
                if editor == "slide":
                    body += (
                        "`DocumentServerPublicOrigin` 是内部实现细节，不得用 `curl` 探测。\n"
                        "session 失败时最多重试一次。\n"
                    )
                (skill_root / "SKILL.md").write_text(
                    f"---\nname: {config['skill']}\n---\n\n{body}",
                    encoding="utf-8",
                )

            artifacts = sync_contract.build_artifacts(
                self.contract,
                zenmind_root,
                "https://office.example.test",
                ("slide",),
            )
            external_paths = [
                path
                for path in artifacts
                if path.is_relative_to(zenmind_root)
            ]
            self.assertTrue(external_paths)
            self.assertTrue(
                all("online-pptx" in str(path) for path in external_paths)
            )

    def test_generation_rejects_missing_internal_service_boundary(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            zenmind_root = Path(directory)
            for config in sync_contract.EDITOR_CONFIG.values():
                skill_root = zenmind_root / "skills-center" / config["skill"]
                skill_root.mkdir(parents=True)
                (skill_root / "SKILL.md").write_text(
                    "---\nname: example\n---\n\n# Body\n",
                    encoding="utf-8",
                )
            with self.assertRaisesRegex(
                ValueError,
                "must keep the internal-service boundary",
            ):
                sync_contract.build_artifacts(
                    self.contract,
                    zenmind_root,
                    "https://office.example.test",
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
            expected = '{"version":"0.2.0"}\n'
            projection.write_text(expected, encoding="utf-8")
            self.assertEqual(
                sync_contract.drifted_paths({projection: expected}),
                [],
            )
            projection.write_text(
                expected.replace("0.2.0", "manual-edit"),
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
                self.assertEqual(scoped["toolPrefix"], "")
                prefix = sync_contract.EDITOR_CONFIG[editor]["prefix"]
                self.assertTrue(scoped["tools"])
                self.assertFalse(any(name.startswith(prefix) for name in scoped["tools"]))
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

    def test_freeze_panes_actions_have_strict_mutually_exclusive_shapes(self) -> None:
        valid_arguments = (
            {"action": "unfreeze"},
            {"action": "freezeAt", "range": "A1:C4"},
            {"action": "freezeRows", "count": 2},
            {"action": "freezeColumns", "count": 3},
        )
        for arguments in valid_arguments:
            with self.subTest(valid=arguments):
                normalized, changes = copilot_server.require_valid_editor_tool_calls(
                    "cell",
                    [
                        {
                            "name": "sheets_manage_freeze_panes",
                            "arguments": arguments,
                        }
                    ],
                )
                self.assertEqual(normalized[0]["arguments"], arguments)
                self.assertEqual(changes, [])

        invalid_arguments = (
            {"action": "unfreeze", "range": "A1:B2"},
            {"action": "unfreeze", "count": 1},
            {"action": "freezeAt"},
            {"action": "freezeAt", "range": "A1:B2", "count": 1},
            {"action": "freezeRows"},
            {"action": "freezeRows", "count": 1, "range": "A1:B2"},
            {"action": "freezeRows", "count": 0},
            {"action": "freezeColumns"},
            {"action": "freezeColumns", "count": 1, "range": "A1:B2"},
            {"action": "freezeColumns", "count": -1},
        )
        for arguments in invalid_arguments:
            with self.subTest(invalid=arguments):
                with self.assertRaises(copilot_server.BridgeError) as caught:
                    copilot_server.require_valid_editor_tool_calls(
                        "cell",
                        [
                            {
                                "name": "sheets_manage_freeze_panes",
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
                error["path"] == "arguments"
                and error["keyword"] == "anyOf"
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
                (1, "arguments", "anyOf"),
                (2, "arguments.target", "semantic"),
            ],
        )
        self.assertNotIn("must-not-leak", json.dumps(validation_errors))


if __name__ == "__main__":
    unittest.main()
