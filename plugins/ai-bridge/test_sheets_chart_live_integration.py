#!/usr/bin/env python3
"""Live chart regression that never exports or downloads the XLSX.

Run with a fresh temporary XLSX capability page open in a browser:

    AI_BRIDGE_LIVE_FILE=<uuid>.xlsx \
      python3 plugins/ai-bridge/test_sheets_chart_live_integration.py

The caller owns the disposable test document. This test writes neutral fixture
data, creates a canonical ``bar`` chart and a ``column`` alias chart, then uses
the live editor bridge to verify that both chart objects are visible online.
"""

from __future__ import annotations

import json
import os
import unittest
import urllib.error
import urllib.request
import uuid


BASE_URL = os.environ.get(
    "AI_BRIDGE_LIVE_BASE_URL",
    "http://127.0.0.1:8088",
).rstrip("/")
FILE_NAME = os.environ.get("AI_BRIDGE_LIVE_FILE", "").strip()


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
        raise AssertionError(
            f"{path} failed with HTTP {error.code}: {detail}"
        ) from error
    if not result.get("ok"):
        raise AssertionError(f"{path} failed: {result}")
    return result


class SheetsChartLiveIntegrationTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        if not FILE_NAME:
            raise RuntimeError(
                "AI_BRIDGE_LIVE_FILE is required; use a fresh disposable XLSX "
                "capability page"
            )

    def test_bar_and_column_alias_are_visible_in_the_live_editor(self) -> None:
        attached = post_json(
            "/copilot-api/bridge/attach",
            {"fileName": FILE_NAME, "editorType": "cell"},
        )
        binding_token = str(attached["bindingToken"])
        tool_calls: list[dict[str, object]] = [
            {
                "name": "sheets_set_values",
                "arguments": {
                    "range": "A1:B4",
                    "values": [
                        ["Category", "Value"],
                        ["A", 1],
                        ["B", 2],
                        ["C", 3],
                    ],
                },
            },
            {
                "name": "sheets_add_chart",
                "arguments": {
                    "range": "A1:B4",
                    "type": "bar",
                    "name": "AiBridgeBarCanonical",
                    "title": "Bar canonical",
                    "fromColumn": 3,
                    "fromRow": 0,
                },
            },
            {
                "name": "sheets_add_chart",
                "arguments": {
                    "range": "A1:B4",
                    "type": "column",
                    "name": "AiBridgeColumnAlias",
                    "title": "Column alias",
                    "fromColumn": 3,
                    "fromRow": 15,
                },
            },
            {
                "name": "sheets_inspect_charts",
                "arguments": {"maxCharts": 20, "includeRaw": False},
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
                "requestId": f"sheets-chart-live-{uuid.uuid4()}",
                "toolCalls": tool_calls,
                "timeoutMs": 120000,
            },
            binding_token,
        )
        execution = executed.get("result") or {}
        self.assertTrue(execution.get("persisted"), executed)
        add_results = [
            item
            for item in execution.get("results", [])
            if item.get("name") == "sheets_add_chart"
        ]
        self.assertEqual(len(add_results), 2)
        self.assertEqual([item.get("type") for item in add_results], ["bar", "bar"])

        inspection = next(
            item
            for item in execution.get("results", [])
            if item.get("name") == "sheets_inspect_charts"
        )
        charts = [
            chart
            for sheet in inspection.get("sheets", [])
            for chart in sheet.get("charts", [])
            if chart.get("name")
            in {"AiBridgeBarCanonical", "AiBridgeColumnAlias"}
        ]
        self.assertEqual(
            {chart.get("name") for chart in charts},
            {"AiBridgeBarCanonical", "AiBridgeColumnAlias"},
        )
        for chart in charts:
            self.assertEqual(chart.get("type"), "bar")
            self.assertGreater(chart.get("widthMm") or 0, 0)
            self.assertGreater(chart.get("heightMm") or 0, 0)
            self.assertTrue(chart.get("series"))


if __name__ == "__main__":
    unittest.main()
