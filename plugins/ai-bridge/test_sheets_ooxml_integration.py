#!/usr/bin/env python3
"""Dedicated live Sheets bridge regression against an exported ONLYOFFICE XLSX.

Run with a fresh XLSX capability page open in a browser:

    AI_BRIDGE_LIVE_FILE=<uuid>.xlsx \
      python3 plugins/ai-bridge/test_sheets_ooxml_integration.py

The test uses the first-party localhost document app, copies the persisted file
from the local DocumentServer container, and verifies the resulting OOXML. It
intentionally checks behavior that JavaScript mocks cannot prove: cell/differential fills,
conditional-format operators, formula progression, validation XML, strict
values, range-style formatting, frozen panes, and worksheet view flags.
"""

from __future__ import annotations

import json
import os
import subprocess
import tempfile
import unittest
import urllib.error
import urllib.request
import uuid
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


BASE_URL = os.environ.get("AI_BRIDGE_LIVE_BASE_URL", "http://127.0.0.1:8088").rstrip("/")
FILE_NAME = os.environ.get("AI_BRIDGE_LIVE_FILE", "").strip()
CONTAINER = os.environ.get("AI_BRIDGE_LIVE_CONTAINER", "onlyoffice-documentserver").strip()
NS = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def post_json(
    path: str,
    payload: dict[str, object],
    bearer_token: str | None = None,
) -> dict[str, object]:
    headers = {"content-type": "application/json"}
    if bearer_token:
        headers["authorization"] = f"Bearer {bearer_token}"
    request = urllib.request.Request(
        f"{BASE_URL}{path}",
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers=headers,
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=120) as response:
            result = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as error:
        detail = error.read().decode("utf-8", "replace")
        raise AssertionError(f"{path} failed with HTTP {error.code}: {detail}") from error
    if not result.get("ok"):
        raise AssertionError(f"{path} failed: {result}")
    return result


def persisted_xlsx(file_name: str, destination: Path) -> None:
    search = subprocess.run(
        [
            "docker",
            "exec",
            CONTAINER,
            "find",
            "/var/lib/onlyoffice/copilot/documents",
            "-type",
            "f",
            "-name",
            file_name,
        ],
        check=True,
        capture_output=True,
        text=True,
    )
    matches = [line.strip() for line in search.stdout.splitlines() if line.strip()]
    if not matches:
        raise AssertionError(f"persisted XLSX not found in {CONTAINER}: {file_name}")
    subprocess.run(
        ["docker", "cp", f"{CONTAINER}:{matches[-1]}", str(destination)],
        check=True,
        capture_output=True,
        text=True,
    )


def rgb_values(root: ET.Element) -> set[str]:
    return {
        value.upper()
        for element in root.findall(".//*[@rgb]")
        if (value := element.get("rgb"))
    }


def normalized_a1(value: object) -> str:
    return str(value or "").rsplit("!", 1)[-1].replace("$", "")


class SheetsOoxmlIntegrationTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        if not FILE_NAME:
            raise RuntimeError(
                "AI_BRIDGE_LIVE_FILE is required; the dedicated live integration "
                "test must not silently skip"
            )

    def test_bridge_output_survives_export_with_correct_semantics(self) -> None:
        attached = post_json(
            "/copilot-api/bridge/attach",
            {"fileName": FILE_NAME, "editorType": "cell"},
        )
        binding_token = str(attached["bindingToken"])
        tool_calls: list[dict[str, object]] = [
            {
                "name": "sheets_set_values",
                "arguments": {
                    "range": "A1:C4",
                    "values": [
                        ["项目", "收入", "成本"],
                        ["一月", 100, 80],
                        ["二月", 90, 95],
                        ["三月", 120, 100],
                    ],
                },
            },
            {
                "name": "sheets_set_formula",
                "arguments": {
                    "range": "D2:D4",
                    "formula": [["=B2-C2"], ["=B3-C3"], ["=B4-C4"]],
                },
            },
            {
                "name": "sheets_format_range",
                "arguments": {
                    "range": "A1:D1",
                    "bold": True,
                    "fontColor": "#FFFFFF",
                    "fillColor": "#17365D",
                },
            },
            {
                "name": "sheets_manage_conditional_format",
                "arguments": {
                    "action": "add",
                    "range": "D2:D4",
                    "type": "cellValue",
                    "operator": "lessThan",
                    "formula1": 0,
                    "fillColor": "#FFF2CC",
                },
            },
            {
                "name": "sheets_manage_validation",
                "arguments": {
                    "action": "add",
                    "range": "E2:E4",
                    "type": "list",
                    "formula1": "一月,二月,三月",
                    "ignoreBlank": True,
                    "showError": True,
                    "errorTitle": "月份无效",
                    "errorMessage": "请从列表中选择",
                },
            },
            {
                "name": "sheets_manage_table",
                "arguments": {
                    "action": "create",
                    "tableMode": "rangeStyle",
                    "range": "A1:D4",
                },
            },
            {
                "name": "sheets_manage_freeze_panes",
                "arguments": {
                    "action": "freezeRows",
                    "count": 2,
                },
            },
            {
                "name": "sheets_manage_page_layout",
                "arguments": {
                    "orientation": "landscape",
                    "displayGridlines": False,
                    "printGridlines": False,
                },
            },
            {
                "name": "sheets_inspect_range",
                "arguments": {
                    "range": "A1:D4",
                    "includeFormat": True,
                    "includeConditionalFormats": True,
                },
            },
            {
                "name": "sheets_inspect_range",
                "arguments": {
                    "range": "E2:E4",
                    "includeValidation": True,
                },
            },
            {"name": "sheets_inspect_page_layout", "arguments": {}},
        ]
        validated = post_json(
            "/copilot-api/bridge/validate",
            {"toolCalls": tool_calls},
            binding_token,
        )
        self.assertTrue(validated.get("valid"), validated)
        executed = post_json(
            "/copilot-api/bridge/execute",
            {
                "bindingToken": binding_token,
                "method": "executeBatch",
                "requestId": f"sheets-ooxml-{uuid.uuid4()}",
                "toolCalls": tool_calls,
                "timeoutMs": 120000,
            },
            binding_token,
        )
        execution = executed.get("result") or {}
        self.assertTrue(execution.get("persisted"), executed)
        self.assertEqual(execution.get("changed"), 8)
        freeze_result = next(
            result
            for result in execution.get("results", [])
            if result.get("name") == "sheets_manage_freeze_panes"
        )
        self.assertTrue(freeze_result.get("verified"))
        self.assertEqual(freeze_result.get("frozenRows"), 2)
        self.assertEqual(freeze_result.get("frozenColumns"), 0)
        self.assertEqual(freeze_result.get("topLeftCell"), "A3")
        page_layout_result = next(
            result
            for result in execution.get("results", [])
            if result.get("name") == "sheets_manage_page_layout"
        )
        self.assertTrue(page_layout_result.get("verified"))
        self.assertFalse(page_layout_result.get("displayGridlines"))
        table_result = next(
            result
            for result in execution.get("results", [])
            if result.get("name") == "sheets_manage_table"
        )
        self.assertEqual(table_result.get("requestedKind"), "rangeStyle")
        self.assertEqual(table_result.get("actualKind"), "rangeStyle")
        self.assertEqual(table_result.get("tableKind"), "rangeStyle")
        self.assertFalse(table_result.get("degraded"))
        self.assertTrue(table_result.get("formatted"))
        self.assertIsNone(table_result.get("table"))
        validation_result = next(
            result
            for result in execution.get("results", [])
            if result.get("name") == "sheets_inspect_range"
            and normalized_a1((result.get("range") or {}).get("address")) == "E2:E4"
        )
        validation = validation_result["range"]["validation"]
        self.assertEqual(validation["type"], "xlValidateList")
        self.assertEqual(validation["formula1"], "一月,二月,三月")
        self.assertTrue(validation["ignoreBlank"])
        self.assertEqual(validation["errorTitle"], "月份无效")

        with tempfile.TemporaryDirectory(prefix="ai-bridge-sheets-ooxml-") as temp_dir:
            xlsx_path = Path(temp_dir) / FILE_NAME
            persisted_xlsx(FILE_NAME, xlsx_path)
            with zipfile.ZipFile(xlsx_path) as archive:
                styles = ET.fromstring(archive.read("xl/styles.xml"))
                sheet = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))

        self.assertIn("FF17365D", rgb_values(styles), "header fill did not survive XLSX export")
        self.assertIn("FFFFF2CC", rgb_values(styles), "conditional fill did not survive XLSX export")

        rules = sheet.findall(".//x:conditionalFormatting/x:cfRule", NS)
        self.assertEqual(len(rules), 1)
        self.assertEqual(rules[0].get("operator"), "lessThan")
        self.assertIsNotNone(rules[0].get("dxfId"))

        formulas = {
            cell.get("r"): cell.findtext("x:f", default="", namespaces=NS)
            for cell in sheet.findall(".//x:c", NS)
            if cell.find("x:f", NS) is not None
        }
        self.assertEqual(
            formulas,
            {"D2": "B2-C2", "D3": "B3-C3", "D4": "B4-C4"},
        )
        errors = [
            cell.findtext("x:v", default="", namespaces=NS)
            for cell in sheet.findall(".//x:c[@t='e']", NS)
        ]
        self.assertNotIn("#N/A", errors)

        validations = sheet.findall(".//x:dataValidations/x:dataValidation", NS)
        self.assertEqual(len(validations), 1)
        self.assertEqual(validations[0].get("type"), "list")
        self.assertEqual(validations[0].get("sqref"), "E2:E4")
        self.assertEqual(
            validations[0].findtext("x:formula1", default="", namespaces=NS),
            "一月,二月,三月",
        )

        sheet_view = sheet.find(".//x:sheetViews/x:sheetView", NS)
        self.assertIsNotNone(sheet_view)
        self.assertEqual(sheet_view.get("showGridLines"), "0")
        pane = sheet_view.find("x:pane", NS)
        self.assertIsNotNone(pane)
        self.assertEqual(pane.get("ySplit"), "2")
        self.assertEqual(pane.get("topLeftCell"), "A3")
        self.assertEqual(pane.get("state"), "frozen")
        page_setup = sheet.find("x:pageSetup", NS)
        self.assertIsNotNone(page_setup)
        self.assertEqual(page_setup.get("orientation"), "landscape")

    def test_freeze_and_screen_gridlines_are_verified_and_persisted(self) -> None:
        attached = post_json(
            "/copilot-api/bridge/attach",
            {"fileName": FILE_NAME, "editorType": "cell"},
        )
        binding_token = str(attached["bindingToken"])
        tool_calls: list[dict[str, object]] = [
            {
                "name": "sheets_manage_freeze_panes",
                "arguments": {"action": "freezeRows", "count": 2},
            },
            {
                "name": "sheets_inspect_freeze_panes",
                "arguments": {},
            },
            {
                "name": "sheets_manage_page_layout",
                "arguments": {
                    "displayGridlines": False,
                    "printGridlines": True,
                },
            },
            {
                "name": "sheets_inspect_page_layout",
                "arguments": {},
            },
        ]
        validated = post_json(
            "/copilot-api/bridge/validate",
            {"toolCalls": tool_calls},
            binding_token,
        )
        self.assertTrue(validated.get("valid"), validated)
        executed = post_json(
            "/copilot-api/bridge/execute",
            {
                "bindingToken": binding_token,
                "method": "executeBatch",
                "requestId": f"sheets-view-ooxml-{uuid.uuid4()}",
                "toolCalls": tool_calls,
                "timeoutMs": 120000,
            },
            binding_token,
        )
        execution = executed.get("result") or {}
        self.assertTrue(execution.get("persisted"), executed)
        self.assertEqual(execution.get("changed"), 2)
        freeze_results = [
            result
            for result in execution.get("results", [])
            if result.get("name") in {
                "sheets_manage_freeze_panes",
                "sheets_inspect_freeze_panes",
            }
        ]
        self.assertEqual(len(freeze_results), 2)
        for result in freeze_results:
            self.assertTrue(result.get("verified"), result)
            self.assertEqual(result.get("frozenRows"), 2)
            self.assertEqual(result.get("frozenColumns"), 0)
            self.assertEqual(result.get("topLeftCell"), "A3")

        page_results = [
            result
            for result in execution.get("results", [])
            if result.get("name") in {
                "sheets_manage_page_layout",
                "sheets_inspect_page_layout",
            }
        ]
        self.assertEqual(len(page_results), 2)
        for result in page_results:
            self.assertTrue(result.get("verified"), result)
            self.assertFalse(result.get("displayGridlines"), result)
            self.assertTrue(result.get("printGridlines"), result)

        with tempfile.TemporaryDirectory(prefix="ai-bridge-sheets-view-ooxml-") as temp_dir:
            xlsx_path = Path(temp_dir) / FILE_NAME
            persisted_xlsx(FILE_NAME, xlsx_path)
            with zipfile.ZipFile(xlsx_path) as archive:
                sheet = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))

        sheet_view = sheet.find(".//x:sheetViews/x:sheetView", NS)
        self.assertIsNotNone(sheet_view)
        self.assertEqual(sheet_view.get("showGridLines"), "0")
        pane = sheet_view.find("x:pane", NS)
        self.assertIsNotNone(pane)
        self.assertEqual(pane.get("ySplit"), "2")
        self.assertEqual(pane.get("topLeftCell"), "A3")
        self.assertEqual(pane.get("state"), "frozen")
        print_options = sheet.find("x:printOptions", NS)
        self.assertIsNotNone(print_options)
        self.assertEqual(print_options.get("gridLines"), "1")


if __name__ == "__main__":
    unittest.main()
