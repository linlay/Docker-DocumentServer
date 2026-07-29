#!/usr/bin/env python3
"""Contract-source, generated projection, normalization, and limit tests."""

from __future__ import annotations

import copy
import importlib.util
import json
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
        self.assertEqual(self.contract["version"], "0.6.0")
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

    def test_check_logic_detects_a_manual_generated_file_edit(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            projection = Path(directory) / "contract.generated.json"
            expected = '{"version":"0.6.0"}\n'
            projection.write_text(expected, encoding="utf-8")
            self.assertEqual(
                sync_contract.drifted_paths({projection: expected}),
                [],
            )
            projection.write_text(
                expected.replace("0.6.0", "manual-edit"),
                encoding="utf-8",
            )
            self.assertEqual(
                sync_contract.drifted_paths({projection: expected}),
                [projection],
            )

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


if __name__ == "__main__":
    unittest.main()
