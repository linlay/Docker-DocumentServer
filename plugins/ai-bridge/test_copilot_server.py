import base64
import io
import json
import os
import re
import socket
import struct
import tempfile
import threading
import time
import unittest
from unittest import mock

import copilot_server


class ContractAlignmentTests(unittest.TestCase):

    def test_document_app_has_no_example_proxy_or_html_rewriting(self):
        config_path = os.path.join(
            os.path.dirname(__file__),
            "nginx-document-app.conf",
        )
        with open(config_path, encoding="utf-8") as stream:
            config = stream.read()

        self.assertNotIn("proxy_pass http://example", config)
        self.assertNotIn("sub_filter", config)
        self.assertNotIn("/__document_editor/", config)
        self.assertNotIn("location", config)
        self.assertIn("document-hub", config)

    def test_gateway_exposes_no_legacy_document_routes(self):
        config_path = os.path.join(
            os.path.dirname(__file__),
            "nginx-document-app.conf",
        )
        with open(config_path, encoding="utf-8") as stream:
            config = stream.read()

        self.assertNotIn("docx|xlsx|pptx", config)
        self.assertNotIn("proxy_pass", config)

    def test_v1_private_relay_allows_only_internal_bridge_routes(self):
        for path in (
            "/new-docx",
            "/admin",
            "/documents/editor",
            "/documents/storage/callback/docx/example",
            "/chat",
            "/forcesave",
            "/editor-config/anonymous",
        ):
            with self.subTest(path=path):
                self.assertFalse(
                    copilot_server.v1_relay_route_allowed("POST", path)
                )
        for path in (
            "/bridge/register",
            "/bridge/internal/execute",
            "/bridge/internal/validate",
            "/bridge/internal/session",
            "/bridge/internal/drop",
            "/bridge/internal/images/import",
        ):
            with self.subTest(path=path):
                self.assertTrue(
                    copilot_server.v1_relay_route_allowed("POST", path)
                )
        self.assertFalse(
            copilot_server.v1_relay_route_allowed("GET", "/admin")
        )
        self.assertTrue(
            copilot_server.v1_relay_route_allowed("GET", "/health")
        )
        for method, path in (
            ("POST", "/bridge/direct/attach"),
            ("GET", "/bridge/direct/health"),
            ("POST", "/bridge/attach"),
            ("POST", "/bridge/execute"),
            ("POST", "/bridge/validate"),
            ("GET", "/bridge/sessions"),
        ):
            with self.subTest(method=method, path=path):
                self.assertFalse(
                    copilot_server.v1_relay_route_allowed(method, path)
                )

    def test_representative_legacy_routes_return_gone(self):
        class Request:
            def __init__(self, path):
                self.path = path
                self.headers = {}
                self.client_address = ("127.0.0.1", 0)

            def log_bridge_error(self, _error, _payload=None):
                return None

        responses = []

        def capture(_handler, status, payload, _headers=None):
            responses.append((status, payload))

        with mock.patch.object(copilot_server, "json_response", side_effect=capture):
            for method, path in (
                ("POST", "/bridge/attach"),
                ("POST", "/chat"),
                ("POST", "/checkpoint"),
                ("POST", "/new-docx"),
                ("GET", "/admin"),
                ("GET", "/docx/123e4567-e89b-42d3-a456-426614174000"),
            ):
                with self.subTest(method=method, path=path):
                    request = Request(path)
                    if method == "POST":
                        copilot_server.Handler.do_POST(request)
                    else:
                        copilot_server.Handler.do_GET(request)
                    status, payload = responses.pop()
                    self.assertEqual(status, 410)
                    code = payload.get("code") or (payload.get("error") or {}).get("code")
                    self.assertEqual(code, "LEGACY_BUSINESS_DISABLED")

    def test_compose_passes_only_private_relay_environment_to_copilot_process(self):
        base_dir = os.path.dirname(__file__)
        with open(
            os.path.join(base_dir, "..", "..", "docker-compose.copilot.yml"),
            encoding="utf-8",
        ) as stream:
            compose = stream.read()
        with open(
            os.path.join(base_dir, "supervisor-ds-copilot.conf"),
            encoding="utf-8",
        ) as stream:
            supervisor = stream.read()
        with open(
            os.path.join(base_dir, "run-copilot.sh"),
            encoding="utf-8",
        ) as stream:
            runner = stream.read()

        self.assertNotIn("DOCUMENT_ADMIN_USERNAME", compose)
        self.assertNotIn("DOCUMENT_ADMIN_PASSWORD", compose)
        self.assertNotIn("DOCUMENT_STORAGE_DIR", compose)
        self.assertNotIn("COPILOT_API_KEY", compose)
        self.assertIn("> /run/onlyoffice-copilot.env", compose)
        self.assertIn('EXAMPLE_ENABLED: "false"', compose)
        self.assertNotIn(
            '"./data/documents:/var/lib/onlyoffice/copilot/documents"',
            compose,
        )
        self.assertIn('AI_RELAY_INTERNAL_SECRET_FILE: "/run/secrets/ai_relay_internal_secret"', compose)
        self.assertIn('COPILOT_BIND_ADDRESS: "0.0.0.0"', compose)
        self.assertIn(
            '"AI_RELAY_INTERNAL_SECRET_FILE=/run/secrets/ai_relay_internal_secret"',
            compose,
        )
        self.assertIn("ONLYOFFICE JWT secret must contain at least 32 bytes", compose)
        self.assertNotIn("onlyoffice-copilot-local", compose)
        self.assertNotIn("nginx-copilot-loopback.conf", compose)
        self.assertIn(
            "command=/bin/bash /opt/onlyoffice-copilot/run-copilot.sh",
            supervisor,
        )
        self.assertIn('export "$assignment"', runner)

    def test_static_asset_cache_revision_is_consistent(self):
        base_dir = os.path.dirname(__file__)
        with open(os.path.join(base_dir, "index.html"), encoding="utf-8") as stream:
            plugin_index = stream.read()
        revisions = re.findall(
            r'[?&]v=([^&"\'\s<>]+)',
            plugin_index,
        )
        with open(os.path.join(base_dir, "config.json"), encoding="utf-8") as stream:
            plugin_config = json.load(stream)
        with open(os.path.join(base_dir, "public-api.json"), encoding="utf-8") as stream:
            public_contract = json.load(stream)
        config_revision = plugin_config["variations"][0]["url"].split("?v=", 1)[1]

        self.assertEqual(revisions, [config_revision])
        self.assertIn("bootstrap.js?v=" + config_revision, plugin_index)
        self.assertRegex(
            config_revision,
            rf"^{re.escape(public_contract['version'])}-[0-9a-f]{{64}}$",
        )
        for relative_path in ("README.md", "INTEGRATION.zh-CN.md"):
            with self.subTest(path=relative_path):
                with open(os.path.join(base_dir, relative_path), encoding="utf-8") as stream:
                    contents = stream.read()
                self.assertIn("?v=<asset-revision>", contents)
                self.assertNotRegex(contents, r"\?v=[^\s\"']*-rev\d+")

    def test_word_argument_schemas_match_the_public_contract(self):
        contract_path = os.path.join(os.path.dirname(__file__), "public-api.json")
        with open(contract_path, encoding="utf-8") as stream:
            contract = json.load(stream)

        contract_names = list(contract["tools"]["word"])
        model_names = list(copilot_server.ARGUMENT_SCHEMAS_BY_EDITOR["word"])

        self.assertEqual(model_names, contract_names)
        self.assertIn("word_add_image", model_names)
        for schema in copilot_server.ARGUMENT_SCHEMAS_BY_EDITOR["word"].values():
            self.assertTrue(schema["description"])
            self.assertNotIn("$ref", json.dumps(schema))

    def test_table_cell_contract_accepts_mixed_values_and_reports_exact_invalid_paths(self):
        valid_word_calls = [
            {
                "name": "word_add_table",
                "arguments": {
                    "rows": 2,
                    "cols": 3,
                    "data": [
                        [
                            "文本",
                            42,
                            {
                                "text": "格式",
                                "fontFamily": "微软雅黑",
                                "color": "#112233",
                                "backgroundColor": "#E8EEF5",
                                "align": "center",
                            },
                        ],
                        [True, None],
                    ],
                },
            },
            {
                "name": "word_add_nested_table",
                "arguments": {
                    "tableIndex": 1,
                    "row": 1,
                    "column": 1,
                    "rows": 1,
                    "cols": 2,
                    "data": [[{"text": "嵌套", "bold": True}, False]],
                },
            },
        ]
        valid_slide_calls = [{
            "name": "slides_add_table",
            "arguments": {
                "slide": 1,
                "rows": 1,
                "columns": 3,
                "data": [[
                    {"text": "格式", "color": "#112233", "align": "center"},
                    7,
                    None,
                ]],
            },
        }]

        self.assertEqual(
            copilot_server.validate_editor_tool_calls("word", valid_word_calls),
            [],
        )
        self.assertEqual(
            copilot_server.validate_editor_tool_calls("slide", valid_slide_calls),
            [],
        )
        self.assertEqual(
            copilot_server.validate_editor_tool_calls(
                "word",
                [{
                    "name": "word_manage_revisions",
                    "arguments": {"action": "setDisplay", "displayMode": "edit"},
                }],
            ),
            [],
        )
        revision_errors = copilot_server.validate_editor_tool_calls(
            "word",
            [{"name": "word_manage_revisions", "arguments": {"action": "setDisplay"}}],
        )
        self.assertEqual(
            [(error["path"], error["keyword"]) for error in revision_errors],
            [("arguments.displayMode", "required")],
        )

        compatibility_and_invalid_calls = [
            {
                "name": "word_add_table",
                "arguments": {
                    "rows": 1,
                    "cols": 1,
                    "data": [[{"text": "错误", "textColor": "#112233"}]],
                },
            },
            {
                "name": "word_add_nested_table",
                "arguments": {
                    "tableIndex": 1,
                    "row": 1,
                    "column": 1,
                    "rows": 1,
                    "cols": 1,
                    "data": [[{"bold": True}]],
                },
            },
        ]
        normalized_calls, normalizations = copilot_server.normalize_editor_tool_calls(
            "word",
            compatibility_and_invalid_calls,
        )
        word_errors = (
            copilot_server.validate_editor_tool_calls("word", normalized_calls)
            + copilot_server.suspicious_word_unit_errors(compatibility_and_invalid_calls)
        )
        slide_errors = copilot_server.validate_editor_tool_calls(
            "slide",
            [{
                "name": "slides_add_table",
                "arguments": {
                    "slide": 1,
                    "rows": 1,
                    "columns": 1,
                    "data": [[{"text": {"nested": "value"}}]],
                },
            }],
        )

        self.assertEqual(
            [(error["tool"], error["path"]) for error in word_errors],
            [
                ("word_add_nested_table", "arguments.data[0][0].text"),
            ],
        )
        self.assertEqual(
            normalized_calls[0]["arguments"]["data"][0][0]["color"],
            "#112233",
        )
        self.assertEqual(normalizations[0]["path"], "arguments.data[0][0].textColor")
        self.assertEqual(
            [(error["tool"], error["path"]) for error in slide_errors],
            [("slides_add_table", "arguments.data[0][0].text")],
        )

    def test_word_numbering_contract_normalizes_compatibility_fields_and_rejects_ambiguous_shapes(self):
        raw_calls = [{
            "name": "word_set_numbering",
            "arguments": {
                "kind": "multilevel",
                "levels": [
                    {
                        "numberFormat": "lowerRoman",
                        "hangingIndentPt": 18,
                        "leftIndentPt": 36,
                        "restart": 1,
                    },
                    {
                        "level": 1,
                        "format": "decimal",
                        "numberFormat": "upperLetter",
                        "firstLineIndentPt": -12,
                        "hangingIndentPt": 20,
                    },
                ],
                "assignments": [
                    {"paragraphIndex": 1, "level": 0},
                    {"paragraphIndex": 2, "level": 1},
                ],
            },
        }]
        normalized_calls, normalizations = copilot_server.normalize_editor_tool_calls(
            "word",
            raw_calls,
        )
        levels = normalized_calls[0]["arguments"]["levels"]
        self.assertEqual(
            levels,
            [
                {
                    "level": 0,
                    "format": "lowerRoman",
                    "firstLineIndentPt": -18,
                    "leftIndentPt": 36,
                    "restart": True,
                },
                {
                    "level": 1,
                    "format": "decimal",
                    "firstLineIndentPt": -12,
                },
            ],
        )
        self.assertEqual(
            {normalization["kind"] for normalization in normalizations},
            {"implicitDefault", "propertyAlias", "canonicalWins", "typeAlias"},
        )
        self.assertEqual(
            copilot_server.validate_editor_tool_calls("word", normalized_calls),
            [],
        )

        for arguments, expected_path in [
            ({}, "arguments.target"),
            (
                {
                    "levels": [{"level": 0}, {"level": 0}],
                    "assignments": [{"paragraphIndex": 1, "level": 0}],
                },
                "arguments.levels[1].level",
            ),
            (
                {"assignments": [{"paragraphIndexes": [1, 2], "level": 0}]},
                "arguments.assignments[0].paragraphIndexes",
            ),
            (
                {
                    "assignments": [{
                        "paragraphIndex": 1,
                        "search": "一级",
                        "level": 0,
                    }],
                },
                "arguments.assignments[0]",
            ),
        ]:
            normalized, _ = copilot_server.normalize_editor_tool_calls(
                "word",
                [{"name": "word_set_numbering", "arguments": arguments}],
            )
            errors = copilot_server.validate_editor_tool_calls("word", normalized)
            self.assertIn(expected_path, [error["path"] for error in errors])


class ImageImportTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.environment = mock.patch.dict(
            os.environ,
            {
                "COPILOT_IMAGE_DIR": self.temp_dir.name,
                "JWT_SECRET": "image-test-secret",
            },
        )
        self.environment.start()
        self.png = (
            b"\x89PNG\r\n\x1a\n"
            + b"\x00\x00\x00\rIHDR"
            + struct.pack(">II", 2, 1)
            + b"\x08\x06\x00\x00\x00"
        )
        self.claims = {
            "documentKey": "document-key-v1",
            "fileName": "demo.docx",
            "fileType": "docx",
            "editorType": "word",
            "userId": "uid-1",
        }

    def tearDown(self):
        self.environment.stop()
        self.temp_dir.cleanup()

    def data_source(self, mime_type="image/png", data=None):
        encoded = base64.b64encode(self.png if data is None else data).decode()
        return {
            "type": "dataUrl",
            "dataUrl": f"data:{mime_type};base64,{encoded}",
        }

    def test_data_url_import_is_content_addressed_and_signed(self):
        first = copilot_server.import_image(self.data_source(), self.claims, now=1000)
        second = copilot_server.import_image(self.data_source(), self.claims, now=1001)

        self.assertEqual(first["asset"]["assetId"], second["asset"]["assetId"])
        self.assertEqual(first["asset"]["mimeType"], "image/png")
        self.assertEqual(first["asset"]["widthPx"], 2)
        self.assertEqual(first["asset"]["heightPx"], 1)
        parsed = copilot_server.urllib.parse.urlsplit(first["asset"]["path"])
        token = copilot_server.urllib.parse.parse_qs(parsed.query)["token"][0]
        asset_id = parsed.path.rsplit("/", 1)[1]
        path, mime_type = copilot_server.verify_image_asset(asset_id, token, now=1001)
        self.assertEqual(mime_type, "image/png")
        with open(path, "rb") as stream:
            self.assertEqual(stream.read(), self.png)

    def test_signed_asset_expires(self):
        imported = copilot_server.import_image(self.data_source(), self.claims, now=1000)
        parsed = copilot_server.urllib.parse.urlsplit(imported["asset"]["path"])
        token = copilot_server.urllib.parse.parse_qs(parsed.query)["token"][0]
        asset_id = parsed.path.rsplit("/", 1)[1]

        with self.assertRaises(copilot_server.BridgeError) as raised:
            copilot_server.verify_image_asset(asset_id, token, now=2000)

        self.assertEqual(raised.exception.code, "IMAGE_ASSET_EXPIRED")

    def test_data_url_mime_must_match_magic_bytes(self):
        with self.assertRaises(copilot_server.BridgeError) as raised:
            copilot_server.import_image(
                self.data_source("image/jpeg"),
                self.claims,
            )

        self.assertEqual(raised.exception.code, "UNSUPPORTED_IMAGE_FORMAT")

    def test_svg_is_sanitized_imported_and_served(self):
        svg = (
            b'<svg xmlns="http://www.w3.org/2000/svg" width="25.4mm" height="12.7mm" '
            b'viewBox="0 0 96 48"><defs><linearGradient id="g">'
            b'<stop offset="0" stop-color="#fff"/></linearGradient></defs>'
            b'<rect width="96" height="48" fill="url(#g)"/></svg>'
        )
        imported = copilot_server.import_image(
            self.data_source("image/svg+xml", svg),
            self.claims,
            now=1000,
        )

        self.assertEqual(imported["asset"]["mimeType"], "image/svg+xml")
        self.assertEqual(imported["asset"]["widthPx"], 96)
        self.assertEqual(imported["asset"]["heightPx"], 48)
        self.assertTrue(imported["asset"]["assetId"].endswith(".svg"))
        parsed = copilot_server.urllib.parse.urlsplit(imported["asset"]["path"])
        token = copilot_server.urllib.parse.parse_qs(parsed.query)["token"][0]
        path, mime_type = copilot_server.verify_image_asset(
            imported["asset"]["assetId"],
            token,
            now=1001,
        )
        self.assertEqual(mime_type, "image/svg+xml")
        with open(path, "rb") as stream:
            sanitized = stream.read()
        self.assertIn(b"<rect", sanitized)
        self.assertNotIn(b"DOCTYPE", sanitized)

    def test_svg_rejects_scripts_events_and_external_resources(self):
        samples = [
            b'<svg xmlns="http://www.w3.org/2000/svg"><script>alert(1)</script></svg>',
            b'<svg xmlns="http://www.w3.org/2000/svg"><rect onclick="alert(1)"/></svg>',
            b'<svg xmlns="http://www.w3.org/2000/svg"><image href="https://evil.test/a.png"/></svg>',
            b'<!DOCTYPE svg [<!ENTITY xxe SYSTEM "file:///etc/passwd">]><svg>&xxe;</svg>',
        ]
        for svg in samples:
            with self.subTest(svg=svg[:50]):
                with self.assertRaises(copilot_server.BridgeError) as raised:
                    copilot_server.import_image(
                        self.data_source("image/svg+xml", svg),
                        self.claims,
                    )
                self.assertEqual(raised.exception.code, "UNSUPPORTED_IMAGE_FORMAT")

    def test_malformed_base64_is_rejected_without_echoing_input(self):
        sensitive = "not-valid-base64!"
        with self.assertRaises(copilot_server.BridgeError) as raised:
            copilot_server.import_image(
                {
                    "type": "dataUrl",
                    "dataUrl": f"data:image/png;base64,{sensitive}",
                },
                self.claims,
            )

        self.assertEqual(raised.exception.code, "INVALID_IMAGE_SOURCE")
        self.assertNotIn(sensitive, raised.exception.message)

    def test_raw_byte_limit_is_enforced_before_image_parsing(self):
        with mock.patch.object(copilot_server, "IMAGE_MAX_BYTES", 8):
            with self.assertRaises(copilot_server.BridgeError) as raised:
                copilot_server.import_image(self.data_source(), self.claims)

        self.assertEqual(raised.exception.code, "IMAGE_TOO_LARGE")

    def test_pixel_limits_are_enforced(self):
        oversized = (
            b"\x89PNG\r\n\x1a\n"
            + b"\x00\x00\x00\rIHDR"
            + struct.pack(">II", copilot_server.IMAGE_MAX_EDGE_PX + 1, 1)
            + b"\x08\x06\x00\x00\x00"
        )
        with self.assertRaises(copilot_server.BridgeError) as raised:
            copilot_server.import_image(self.data_source(data=oversized), self.claims)

        self.assertEqual(raised.exception.code, "IMAGE_TOO_LARGE")

    def test_private_remote_address_is_blocked(self):
        with mock.patch.object(
            copilot_server.socket,
            "getaddrinfo",
            return_value=[(socket.AF_INET, socket.SOCK_STREAM, 6, "", ("127.0.0.1", 443))],
        ):
            with self.assertRaises(copilot_server.BridgeError) as raised:
                copilot_server.validate_image_url("https://example.test/image.png")

        self.assertEqual(raised.exception.code, "IMAGE_FETCH_BLOCKED")

    def test_loopback_ipv6_address_is_blocked(self):
        with mock.patch.object(
            copilot_server.socket,
            "getaddrinfo",
            return_value=[(socket.AF_INET6, socket.SOCK_STREAM, 6, "", ("::1", 443, 0, 0))],
        ):
            with self.assertRaises(copilot_server.BridgeError) as raised:
                copilot_server.validate_image_url("https://example.test/image.png")

        self.assertEqual(raised.exception.code, "IMAGE_FETCH_BLOCKED")

    def test_hostname_allow_list_accepts_subdomains_and_blocks_others(self):
        with mock.patch.dict(
            os.environ,
            {"COPILOT_IMAGE_ALLOWED_HOSTS": "images.example.com"},
        ), mock.patch.object(
            copilot_server.socket,
            "getaddrinfo",
            return_value=[(socket.AF_INET, socket.SOCK_STREAM, 6, "", ("93.184.216.34", 443))],
        ):
            accepted = copilot_server.validate_image_url(
                "https://cdn.images.example.com/image.png"
            )
            with self.assertRaises(copilot_server.BridgeError) as raised:
                copilot_server.validate_image_url("https://example.net/image.png")

        self.assertEqual(accepted, "https://cdn.images.example.com/image.png")
        self.assertEqual(raised.exception.code, "IMAGE_FETCH_BLOCKED")

    def test_redirect_to_private_address_is_blocked(self):
        handler = copilot_server.SafeImageRedirectHandler()
        request = copilot_server.urllib.request.Request(
            "https://public.example/image.png"
        )
        with mock.patch.object(
            copilot_server.socket,
            "getaddrinfo",
            return_value=[(socket.AF_INET, socket.SOCK_STREAM, 6, "", ("10.0.0.8", 443))],
        ):
            with self.assertRaises(copilot_server.BridgeError) as raised:
                handler.redirect_request(
                    request,
                    None,
                    302,
                    "Found",
                    {},
                    "https://private.example/image.png",
                )

        self.assertEqual(raised.exception.code, "IMAGE_FETCH_BLOCKED")

    def test_public_https_image_is_downloaded_with_type_validation(self):
        response = mock.MagicMock()
        response.__enter__.return_value = response
        response.headers.get_content_type.return_value = "image/png"
        response.headers.get.return_value = str(len(self.png))
        response.read.return_value = self.png
        opener = mock.Mock()
        opener.open.return_value = response
        with mock.patch.object(
            copilot_server.socket,
            "getaddrinfo",
            return_value=[(socket.AF_INET, socket.SOCK_STREAM, 6, "", ("93.184.216.34", 443))],
        ), mock.patch.object(copilot_server.urllib.request, "build_opener", return_value=opener):
            data, content_type = copilot_server.decode_image_source({
                "type": "url",
                "url": "https://example.test/image.png",
            })

        self.assertEqual(data, self.png)
        self.assertEqual(content_type, "image/png")

    def test_remote_timeout_returns_safe_fetch_error(self):
        opener = mock.Mock()
        opener.open.side_effect = socket.timeout("remote response body")
        with mock.patch.object(
            copilot_server.socket,
            "getaddrinfo",
            return_value=[(socket.AF_INET, socket.SOCK_STREAM, 6, "", ("93.184.216.34", 443))],
        ), mock.patch.object(copilot_server.urllib.request, "build_opener", return_value=opener):
            with self.assertRaises(copilot_server.BridgeError) as raised:
                copilot_server.decode_image_source({
                    "type": "url",
                    "url": "https://example.test/image.png",
                })

        self.assertEqual(raised.exception.code, "IMAGE_FETCH_FAILED")
        self.assertNotIn("remote response body", raised.exception.message)


class HttpRelayTests(unittest.TestCase):
    def setUp(self):
        self.storage = tempfile.TemporaryDirectory()
        self.storage_env = mock.patch.dict(
            os.environ,
            {"DOCUMENT_STORAGE_DIR": self.storage.name},
        )
        self.storage_env.start()
        self.jwt_secret = "relay-test-secret"
        self.jwt_patcher = mock.patch.object(
            copilot_server,
            "get_jwt_secret",
            return_value=self.jwt_secret,
        )
        self.jwt_patcher.start()
        with copilot_server.BRIDGE_CONDITION:
            copilot_server.BRIDGE_SESSIONS.clear()
            copilot_server.BRIDGE_AUTHORITATIVE.clear()
            copilot_server.BRIDGE_SUPERSEDED.clear()
            copilot_server.BRIDGE_COMMANDS.clear()
            copilot_server.BRIDGE_REQUESTS.clear()
            copilot_server.BRIDGE_REJECTED_CREDENTIALS.clear()

    def tearDown(self):
        with copilot_server.BRIDGE_CONDITION:
            copilot_server.BRIDGE_SESSIONS.clear()
            copilot_server.BRIDGE_AUTHORITATIVE.clear()
            copilot_server.BRIDGE_SUPERSEDED.clear()
            copilot_server.BRIDGE_COMMANDS.clear()
            copilot_server.BRIDGE_REQUESTS.clear()
            copilot_server.BRIDGE_REJECTED_CREDENTIALS.clear()
        self.jwt_patcher.stop()
        self.storage_env.stop()
        self.storage.cleanup()

    def editor_token(
        self,
        document_key="document-key-v1",
        expires_in=300,
        *,
        document_id="",
        file_name="demo.docx",
        file_type="docx",
        editor_type="word",
        user_id="uid-1",
    ):
        return copilot_server.sign_jwt(
            {
                "documentId": document_id,
                "document": {
                    "key": document_key,
                    "title": file_name,
                    "fileType": file_type,
                },
                "documentType": editor_type,
                "editorConfig": {"user": {"id": user_id}},
                "exp": int(time.time()) + expires_in,
            },
            self.jwt_secret,
        )

    def editor_state(
        self,
        document_key="document-key-v1",
        *,
        document_id="",
        file_name="demo.docx",
        file_type="docx",
        editor_type="word",
        user_id="uid-1",
        ready=True,
    ):
        tool_name = {
            "word": "word_inspect",
            "cell": "sheets_inspect",
            "slide": "slides_inspect",
        }[editor_type]
        return {
            "ready": ready,
            "editorType": editor_type,
            "contractVersion": copilot_server.PUBLIC_API_CONTRACT_VERSION,
            "contractSha256": copilot_server.PUBLIC_API_CONTRACT_SHA256,
            "context": {
                "documentId": document_id,
                "documentKey": document_key,
                "fileName": file_name,
                "fileType": file_type,
                "editorType": editor_type,
                "userId": user_id,
            },
            "capabilities": {
                "tools": [tool_name],
                "controls": ["save", "history", "undo", "redo"],
            },
        }

    def register(
        self,
        session_id="http-session:test",
        document_key="document-key-v1",
        *,
        document_id="",
        file_name="demo.docx",
        file_type="docx",
        editor_type="word",
        user_id="uid-1",
        ready=True,
    ):
        return copilot_server.bridge_register(
            {
                "sessionId": session_id,
                "editorToken": self.editor_token(
                    document_key,
                    document_id=document_id,
                    file_name=file_name,
                    file_type=file_type,
                    editor_type=editor_type,
                    user_id=user_id,
                ),
                "state": self.editor_state(
                    document_key,
                    document_id=document_id,
                    file_name=file_name,
                    file_type=file_type,
                    editor_type=editor_type,
                    user_id=user_id,
                    ready=ready,
                ),
            }
        )

    def create_document_file(self, document_id, file_type):
        file_name = f"{document_id}.{file_type}"
        with open(os.path.join(self.storage.name, file_name), "wb") as stream:
            stream.write(b"test")
        return file_name

    def public_calls(self, editor, calls):
        return [
            {
                **call,
                "name": copilot_server.public_tool_name(editor, call["name"]),
            }
            for call in calls
        ]

    def editor_failure(self, code, message="editor failed", details=None, request_id="failure"):
        if "http-session:test" not in copilot_server.BRIDGE_SESSIONS:
            registration = self.register()
        else:
            registration = {
                "relayKey": copilot_server.BRIDGE_SESSIONS["http-session:test"]["relayKey"],
            }
        claims = copilot_server.verify_editor_jwt(self.editor_token())
        holder = {}

        def execute():
            try:
                copilot_server.bridge_execute(
                    {
                        "method": "executeTool",
                        "name": "inspect",
                        "arguments": {"maxChars": 1000},
                        "requestId": request_id,
                        "timeoutMs": 2000,
                    },
                    claims,
                )
            except copilot_server.BridgeError as error:
                holder["error"] = error

        worker = threading.Thread(target=execute)
        worker.start()
        command = copilot_server.bridge_poll(
            {
                "sessionId": "http-session:test",
                "relayKey": registration["relayKey"],
                "state": self.editor_state(),
                "timeoutMs": 1000,
            }
        )["command"]
        copilot_server.bridge_result(
            {
                "sessionId": "http-session:test",
                "relayKey": registration["relayKey"],
                "commandId": command["commandId"],
                "state": self.editor_state(),
                "ok": False,
                "error": {
                    "code": code,
                    "message": message,
                    "details": details,
                },
            }
        )
        worker.join(timeout=2)
        self.assertFalse(worker.is_alive())
        return holder["error"]

    def test_editor_error_status_mapping_is_centralized_and_unknown_is_500(self):
        cases = {
            "INVALID_COMMAND": 400,
            "EDITOR_TOKEN_EXPIRED": 401,
            "CONTROL_NOT_ALLOWED": 403,
            "IMAGE_ASSET_EXPIRED": 410,
            "REQUEST_ID_CONFLICT": 409,
            "ARGUMENTS_TOO_LARGE": 413,
            "UNSUPPORTED_IMAGE_FORMAT": 415,
            "INVALID_ARGUMENTS": 422,
            "INVALID_TOOL_ARGUMENTS": 422,
            "EDITOR_CAPABILITY_PROBE_FAILED": 500,
            "EDITOR_COMMAND_REJECTED": 500,
            "EDITOR_COMMAND_FAILED": 500,
            "EXECUTION_FAILED": 500,
            "WORD_API_UNSUPPORTED": 501,
            "SHEETS_API_UNSUPPORTED": 501,
            "SHEETS_RUNTIME_INCOMPATIBLE": 500,
            "SHEETS_CHART_PARTIAL_MUTATION": 500,
            "PERSISTENCE_INVALID_RESPONSE": 502,
            "PERSISTENCE_FAILED": 502,
            "PERSISTENCE_NOT_AVAILABLE": 503,
            "NOT_READY": 503,
            "BRIDGE_TIMEOUT": 504,
            "PERSISTENCE_TIMEOUT": 504,
            "FUTURE_EDITOR_ERROR": 500,
        }
        for code, expected_status in cases.items():
            with self.subTest(code=code):
                self.assertEqual(
                    copilot_server.editor_error_http_status(code),
                    expected_status,
                )

    def test_editor_failure_uses_status_map_and_preserves_safe_details(self):
        details = {
            "phase": "word-command",
            "tool": "format_table_advanced",
            "toolCallIndex": 0,
            "completedToolCalls": 0,
            "partialMutationPossible": False,
        }
        error = self.editor_failure(
            "INVALID_ARGUMENTS",
            "设置重复表头时必须提供 row",
            details,
            request_id="invalid-table-header",
        )
        self.assertEqual(error.status, 422)
        self.assertEqual(error.code, "INVALID_ARGUMENTS")
        self.assertEqual(error.message, "设置重复表头时必须提供 row")
        self.assertEqual(error.details, details)

    def test_editor_token_binds_registration_to_document_key(self):
        with self.assertRaises(copilot_server.BridgeError) as raised:
            copilot_server.bridge_register(
                {
                    "sessionId": "http-session:test",
                    "editorToken": self.editor_token("document-key-v2"),
                    "state": self.editor_state("document-key-v1"),
                }
            )

        self.assertEqual(raised.exception.code, "DOCUMENT_MISMATCH")

    def test_document_id_is_stable_when_onlyoffice_normalizes_title(self):
        document_id = "99999999-9999-4999-8999-999999999999"
        normalized_state = self.editor_state(
            document_id=document_id,
            file_name="Platform 匿名模式验收.docx",
        )
        registration = copilot_server.bridge_register(
            {
                "sessionId": "http-session:title-normalized",
                "editorToken": self.editor_token(
                    document_id=document_id,
                    file_name="Platform 匿名模式验收",
                ),
                "state": normalized_state,
            }
        )

        self.assertTrue(registration["ok"])
        self.assertEqual(
            registration["session"]["documentId"],
            document_id,
        )
        polled = copilot_server.bridge_poll(
            {
                "sessionId": "http-session:title-normalized",
                "relayKey": registration["relayKey"],
                "timeoutMs": 100,
                "state": normalized_state,
            }
        )
        self.assertIsNone(polled["command"])

    def test_expired_editor_token_is_rejected(self):
        with self.assertRaises(copilot_server.BridgeError) as raised:
            copilot_server.verify_editor_jwt(self.editor_token(expires_in=-1))

        self.assertEqual(raised.exception.code, "EDITOR_TOKEN_EXPIRED")
        self.assertNotIn("bind_current_word", raised.exception.message)

    def test_repeated_expired_registration_enters_terminal_state(self):
        payload = {
            "sessionId": "http-session:expired",
            "editorToken": self.editor_token(expires_in=-1),
            "state": self.editor_state(),
        }

        with self.assertRaises(copilot_server.BridgeError) as first:
            copilot_server.bridge_register(payload)
        with self.assertRaises(copilot_server.BridgeError) as second:
            copilot_server.bridge_register(payload)
        with self.assertRaises(copilot_server.BridgeError) as third:
            copilot_server.bridge_register(payload)

        self.assertEqual(first.exception.code, "EDITOR_TOKEN_EXPIRED")
        self.assertEqual(second.exception.code, "SESSION_SUPERSEDED")
        self.assertFalse(second.exception.suppress_log)
        self.assertEqual(third.exception.code, "SESSION_SUPERSEDED")
        self.assertTrue(third.exception.suppress_log)

    def test_successful_registration_clears_credential_failure_latch(self):
        token = self.editor_token()
        payload = {
            "sessionId": "http-session:mismatch",
            "editorToken": token,
            "state": self.editor_state("document-key-v2"),
        }
        with self.assertRaises(copilot_server.BridgeError):
            copilot_server.bridge_register(payload)

        payload["state"] = self.editor_state()
        registration = copilot_server.bridge_register(payload)

        self.assertTrue(registration["ok"])
        self.assertNotIn(
            copilot_server.bridge_registration_fingerprint(payload),
            copilot_server.BRIDGE_REJECTED_CREDENTIALS,
        )


    def test_http_relay_only_allows_large_arguments_for_image_tools(self):
        large_data_url = "data:image/png;base64," + ("A" * 300000)
        session = {
            "state": {
                "editorType": "word",
                "capabilities": {
                    "tools": ["word_add_image", "word_replace_text"],
                },
            },
        }
        image_command = copilot_server.bridge_build_command(
            {
                "method": "executeTool",
                "requestId": "large-image",
                "name": "add_image",
                "arguments": {
                    "source": {"type": "dataUrl", "dataUrl": large_data_url},
                },
            },
            session,
        )
        self.assertEqual(image_command["params"]["name"], "word_add_image")

        with self.assertRaises(copilot_server.BridgeError) as raised:
            copilot_server.bridge_build_command(
                {
                    "method": "executeTool",
                    "requestId": "large-text",
                    "name": "replace_text",
                    "arguments": {"search": "a", "replace": "A" * 300000},
                },
                session,
            )

        self.assertEqual(raised.exception.code, "ARGUMENTS_TOO_LARGE")

    def test_public_inspect_resolves_to_each_editor_internal_tool(self):
        for editor, internal_name in {
            "word": "word_inspect",
            "slide": "slides_inspect",
            "cell": "sheets_inspect",
        }.items():
            with self.subTest(editor=editor):
                command = copilot_server.bridge_build_command(
                    {
                        "method": "executeTool",
                        "requestId": f"short-inspect-{editor}",
                        "name": "inspect",
                        "arguments": {},
                    },
                    {
                        "state": {
                            "editorType": editor,
                            "capabilities": {"tools": [internal_name]},
                        },
                    },
                )
                self.assertEqual(command["params"]["name"], internal_name)

    def test_http_relay_rejects_prefixed_public_tool_names_before_delivery(self):
        for editor, internal_name in {
            "word": "word_inspect",
            "slide": "slides_inspect",
            "cell": "sheets_inspect",
        }.items():
            with self.subTest(editor=editor):
                with self.assertRaises(copilot_server.BridgeError) as raised:
                    copilot_server.bridge_build_command(
                        {
                            "method": "executeTool",
                            "requestId": f"legacy-prefixed-name-{editor}",
                            "name": internal_name,
                            "arguments": {},
                        },
                        {
                            "state": {
                                "editorType": editor,
                                "capabilities": {"tools": [internal_name]},
                            },
                        },
                    )

                self.assertEqual(raised.exception.status, 400)
                self.assertEqual(raised.exception.code, "TOOL_NOT_ALLOWED")

    def test_batch_endpoints_accept_action_alias_and_report_envelope_paths(self):
        session = {
            "state": {
                "editorType": "slide",
                "capabilities": {"tools": ["slides_set_size"]},
            },
        }
        command = copilot_server.bridge_build_command(
            {
                "method": "executeBatch",
                "requestId": "batch-action-alias-1",
                "toolCalls": [
                    {"action": "set_size", "arguments": {"preset": "wide"}},
                ],
            },
            session,
        )
        self.assertEqual(
            command["params"]["toolCalls"],
            [{"name": "slides_set_size", "arguments": {"preset": "wide"}}],
        )

        self.register(
            file_name="demo.pptx",
            file_type="pptx",
            editor_type="slide",
        )
        with copilot_server.BRIDGE_CONDITION:
            copilot_server.BRIDGE_SESSIONS["http-session:test"]["state"]["capabilities"]["tools"] = [
                "slides_set_size",
            ]
        claims = {
            "fileName": "demo.pptx",
            "fileType": "pptx",
            "editorType": "slide",
            "userId": "uid-1",
        }
        validated = copilot_server.bridge_validate(
            {
                "toolCalls": [
                    {"action": "set_size", "arguments": {"preset": "wide"}},
                ],
            },
            claims,
        )
        self.assertTrue(validated["valid"])
        self.assertEqual(validated["toolCalls"], 1)

        with self.assertRaises(copilot_server.BridgeError) as raised:
            copilot_server.bridge_validate(
                {
                    "toolCalls": [
                        {"name": "set_size", "action": "inspect"},
                    ],
                },
                claims,
            )
        self.assertEqual(raised.exception.code, "INVALID_TOOL_CALL")
        self.assertEqual(raised.exception.details["toolCallIndex"], 0)
        self.assertEqual(raised.exception.details["path"], "toolCalls[0].action")

    def test_batch_endpoints_tolerate_up_to_30_calls_for_every_editor(self):
        editors = {
            "word": ("docx", "word_inspect"),
            "slide": ("pptx", "slides_inspect"),
            "cell": ("xlsx", "sheets_inspect"),
        }
        # 22 and 21 are the page-2/page-3 regression batch sizes.
        accepted_counts = (20, 21, 22, 30)

        for editor, (file_type, internal_name) in editors.items():
            session = {
                "state": {
                    "editorType": editor,
                    "capabilities": {"tools": [internal_name]},
                },
            }
            session_id = f"http-session:batch-limit-{editor}"
            document_key = f"batch-limit-{editor}-key"
            file_name = f"batch-limit.{file_type}"
            self.register(
                session_id=session_id,
                document_key=document_key,
                file_name=file_name,
                file_type=file_type,
                editor_type=editor,
            )
            claims = {
                "fileName": file_name,
                "fileType": file_type,
                "editorType": editor,
                "userId": "uid-1",
            }

            for count in accepted_counts:
                public_calls = [
                    {"name": "inspect", "arguments": {}}
                    for _index in range(count)
                ]
                with self.subTest(editor=editor, endpoint="execute", count=count):
                    command = copilot_server.bridge_build_command(
                        {
                            "method": "executeBatch",
                            "requestId": f"batch-limit-{editor}-{count}",
                            "toolCalls": public_calls,
                        },
                        session,
                    )
                    self.assertEqual(len(command["params"]["toolCalls"]), count)
                    self.assertTrue(all(
                        call["name"] == internal_name
                        for call in command["params"]["toolCalls"]
                    ))

                with self.subTest(editor=editor, endpoint="validate", count=count):
                    validated = copilot_server.bridge_validate(
                        {"toolCalls": public_calls},
                        claims,
                    )
                    self.assertTrue(validated["valid"])
                    self.assertEqual(validated["toolCalls"], count)

            for count in (0, 31):
                public_calls = [
                    {"name": "inspect", "arguments": {}}
                    for _index in range(count)
                ]
                with self.subTest(editor=editor, endpoint="execute", count=count):
                    with self.assertRaises(copilot_server.BridgeError) as raised:
                        copilot_server.bridge_build_command(
                            {
                                "method": "executeBatch",
                                "requestId": f"batch-limit-{editor}-rejected-{count}",
                                "toolCalls": public_calls,
                            },
                            session,
                        )
                    self.assertEqual(raised.exception.code, "INVALID_TOOL_CALL")
                    self.assertIn("1 到 20", raised.exception.message)

                with self.subTest(editor=editor, endpoint="validate", count=count):
                    with self.assertRaises(copilot_server.BridgeError) as raised:
                        copilot_server.bridge_validate(
                            {"toolCalls": public_calls},
                            claims,
                        )
                    self.assertEqual(raised.exception.code, "INVALID_TOOL_CALL")
                    self.assertIn("1 到 20", raised.exception.message)

    def test_public_tool_from_another_editor_is_rejected(self):
        with self.assertRaises(copilot_server.BridgeError) as raised:
            copilot_server.bridge_build_command(
                {
                    "method": "executeTool",
                    "requestId": "wrong-editor-short-name",
                    "name": "manage_freeze_panes",
                    "arguments": {"action": "unfreeze"},
                },
                {
                    "state": {
                        "editorType": "word",
                        "capabilities": {"tools": ["word_inspect"]},
                    },
                },
            )
        self.assertEqual(raised.exception.code, "TOOL_NOT_ALLOWED")

    def test_public_session_and_results_only_expose_short_tool_names(self):
        session = {
            "state": {
                "ready": True,
                "editorType": "word",
                "capabilities": {
                    "tools": ["word_inspect", "word_append_paragraph"],
                    "controls": ["save"],
                },
            },
        }

        public_session = copilot_server.bridge_public_session("session-1", session)
        self.assertEqual(
            public_session["capabilities"]["tools"],
            ["inspect", "append_paragraph"],
        )
        projected = copilot_server.project_public_tool_fields(
            "word",
            {
                "results": [{"name": "word_inspect", "tool": "word_inspect"}],
                "details": {
                    "tool": "word_append_paragraph",
                    "validationErrors": [{"tool": "word_inspect"}],
                },
            },
        )
        self.assertEqual(projected["results"][0], {"name": "inspect", "tool": "inspect"})
        self.assertEqual(projected["details"]["tool"], "append_paragraph")
        self.assertEqual(projected["details"]["validationErrors"][0]["tool"], "inspect")

    def test_validate_accepts_image_arguments_between_standard_and_image_limits(self):
        self.register()
        with copilot_server.BRIDGE_CONDITION:
            copilot_server.BRIDGE_SESSIONS["http-session:test"]["state"]["capabilities"]["tools"] = [
                "word_add_image",
            ]
        claims = {
            "fileName": "demo.docx",
            "fileType": "docx",
            "editorType": "word",
            "userId": "uid-1",
        }

        result = copilot_server.bridge_validate(
            {
                "toolCalls": [
                    {
                        "name": "add_image",
                        "arguments": {
                            "source": {
                                "type": "dataUrl",
                                "dataUrl": "data:image/png;base64," + ("A" * 300000),
                            },
                        },
                    },
                ],
            },
            claims,
        )
        self.assertTrue(result["valid"])

        with self.assertRaises(copilot_server.BridgeError) as raised:
            copilot_server.bridge_validate(
                {
                    "toolCalls": [
                        {
                            "name": "add_image",
                            "arguments": {
                                "source": {
                                    "type": "dataUrl",
                                    "dataUrl": "data:image/png;base64,"
                                    + ("A" * 12_000_001),
                                },
                            },
                        },
                    ],
                },
                claims,
            )
        self.assertEqual(raised.exception.code, "ARGUMENTS_TOO_LARGE")

    def test_word_arguments_are_normalized_before_validation_and_fingerprinting(self):
        allowed = [
            "word_set_document_properties",
            "word_set_page_layout",
            "word_manage_section",
            "word_manage_style",
            "word_add_table",
        ]
        session = {
            "state": {
                "editorType": "word",
                "capabilities": {"tools": allowed},
            },
        }
        compatibility_calls = [
            {
                "name": "word_set_document_properties",
                "arguments": {"author": "兼容作者"},
            },
            {
                "name": "word_set_page_layout",
                "arguments": {"differentFirstPage": True, "pageSize": "letter"},
            },
            {
                "name": "word_manage_section",
                "arguments": {"type": "Next Page"},
            },
            {
                "name": "word_manage_style",
                "arguments": {
                    "name": "正文",
                    "align": "justified",
                    "lineRule": "multiple",
                },
            },
            {
                "name": "word_add_table",
                "arguments": {
                    "rows": 1,
                    "columns": 1,
                    "align": "centre",
                    "data": [[{"text": "值", "textColor": "#112233"}]],
                },
            },
        ]
        canonical_calls = [
            {
                "name": "word_set_document_properties",
                "arguments": {"creator": "兼容作者"},
            },
            {
                "name": "word_set_page_layout",
                "arguments": {"titlePage": True, "pageSize": "Letter"},
            },
            {
                "name": "word_manage_section",
                "arguments": {"type": "nextPage"},
            },
            {
                "name": "word_manage_style",
                "arguments": {
                    "name": "正文",
                    "align": "both",
                    "lineRule": "auto",
                },
            },
            {
                "name": "word_add_table",
                "arguments": {
                    "rows": 1,
                    "cols": 1,
                    "align": "center",
                    "data": [[{"text": "值", "color": "#112233"}]],
                },
            },
        ]

        compatibility = copilot_server.bridge_build_command(
            {
                "method": "executeBatch",
                "requestId": "normalize-compatible",
                "toolCalls": self.public_calls("word", compatibility_calls),
            },
            session,
        )
        canonical = copilot_server.bridge_build_command(
            {
                "method": "executeBatch",
                "requestId": "normalize-canonical",
                "toolCalls": self.public_calls("word", canonical_calls),
            },
            session,
        )

        self.assertEqual(compatibility["params"], canonical["params"])
        self.assertGreaterEqual(len(compatibility["argumentNormalizations"]), 9)
        self.assertEqual(canonical["argumentNormalizations"], [])
        encoded = json.dumps(
            compatibility["argumentNormalizations"],
            ensure_ascii=False,
        )
        self.assertNotIn("兼容作者", encoded)
        self.assertNotIn("#112233", encoded)

    def test_sheets_schema_accepts_formula_matrices_and_only_documented_aliases(self):
        allowed = [
            "sheets_set_formula",
            "sheets_format_range",
            "sheets_manage_table",
            "sheets_manage_conditional_format",
            "sheets_manage_validation",
            "sheets_manage_page_layout",
        ]
        session = {
            "state": {
                "editorType": "cell",
                "capabilities": {"tools": allowed},
            },
        }
        tool_calls = [
            {
                "name": "sheets_set_formula",
                "arguments": {
                    "range": "D2:D4",
                    "formula": [["=B2-C2"], ["=B3-C3"], ["=B4-C4"]],
                },
            },
            {
                "name": "sheets_manage_conditional_format",
                "arguments": {
                    "action": "add",
                    "range": "D2:D4",
                    "operator": "lessThan",
                },
            },
            {
                "name": "sheets_manage_validation",
                "arguments": {
                    "action": "add",
                    "range": "A2:A4",
                    "type": "list",
                    "alertStyle": "warning",
                    "formula1": "A,B,C",
                },
            },
            {
                "name": "sheets_manage_page_layout",
                "arguments": {
                    "orientation": "landscape",
                    "displayGridlines": False,
                },
            },
        ]
        command = copilot_server.bridge_build_command(
            {
                "method": "executeBatch",
                "requestId": "sheets-public-aliases",
                "toolCalls": self.public_calls("cell", tool_calls),
            },
            session,
        )
        expected_calls = json.loads(json.dumps(tool_calls))
        expected_calls[1]["arguments"]["operator"] = "xlLess"
        expected_calls[2]["arguments"]["type"] = "xlValidateList"
        expected_calls[3]["arguments"]["orientation"] = "xlLandscape"
        self.assertEqual(command["params"]["toolCalls"], expected_calls)
        self.assertEqual(len(command["argumentNormalizations"]), 3)

        invalid_calls = [
            {
                "name": "sheets_manage_conditional_format",
                "arguments": {
                    "action": "add",
                    "range": "A1",
                    "operator": "smallerMaybe",
                },
            },
            {
                "name": "sheets_set_formula",
                "arguments": {"range": "A1", "formula": 42},
            },
            {
                "name": "sheets_set_formula",
                "arguments": {
                    "range": "A1:B2",
                    "formula": [["=1"], ["=2", "=3"]],
                },
            },
            {
                "name": "sheets_format_range",
                "arguments": {"range": "A1:B2"},
            },
            {
                "name": "sheets_manage_table",
                "arguments": {
                    "action": "create",
                    "range": "A1:B2",
                    "hasHeaders": True,
                },
            },
            {
                "name": "sheets_manage_page_layout",
                "arguments": {"marginsPt": {}},
            },
        ]
        for call in invalid_calls:
            with self.subTest(name=call["name"]):
                with self.assertRaises(copilot_server.BridgeError) as raised:
                    copilot_server.bridge_build_command(
                        {
                            "method": "executeTool",
                            "requestId": "sheets-invalid-public-enum",
                            "name": copilot_server.public_tool_name("cell", call["name"]),
                            "arguments": call["arguments"],
                        },
                        session,
                    )
                self.assertEqual(raised.exception.code, "INVALID_TOOL_ARGUMENTS")

    def test_word_canonical_property_wins_without_relaxing_unsafe_inputs(self):
        normalized, normalizations = copilot_server.normalize_editor_tool_calls(
            "word",
            [{
                "name": "word_set_document_properties",
                "arguments": {"author": 42, "creator": "标准作者"},
            }],
        )
        self.assertEqual(normalized[0]["arguments"], {"creator": "标准作者"})
        self.assertEqual(normalizations[0]["kind"], "canonicalWins")

        invalid_cases = [
            (
                {
                    "name": "word_set_document_properties",
                    "arguments": {"unknownProperty": "x"},
                },
                "arguments.unknownProperty",
            ),
            (
                {
                    "name": "word_set_page_layout",
                    "arguments": {"differentFirstPage": "true"},
                },
                "arguments.titlePage",
            ),
            (
                {
                    "name": "word_append_paragraph",
                    "arguments": {"text": "x", "leftIndent": 720},
                },
                "arguments.leftIndent",
            ),
        ]
        for call, expected_path in invalid_cases:
            with self.subTest(expected_path=expected_path):
                with self.assertRaises(copilot_server.BridgeError) as raised:
                    copilot_server.require_valid_editor_tool_calls("word", [call])
                self.assertEqual(raised.exception.code, "INVALID_TOOL_ARGUMENTS")
                self.assertIn(
                    expected_path,
                    [
                        error["path"]
                        for error in raised.exception.details["validationErrors"]
                    ],
                )

    def test_word_validate_batch_is_read_only_and_returns_safe_normalizations(self):
        self.register()
        with copilot_server.BRIDGE_CONDITION:
            session = copilot_server.BRIDGE_SESSIONS["http-session:test"]
            session["state"]["capabilities"]["tools"] = [
                "word_set_document_properties",
                "word_manage_style",
            ]
        claims = {
            "fileName": "demo.docx",
            "fileType": "docx",
            "editorType": "word",
            "userId": "uid-1",
        }
        commands_before = dict(copilot_server.BRIDGE_COMMANDS)
        requests_before = dict(copilot_server.BRIDGE_REQUESTS)

        result = copilot_server.bridge_validate(
            {
                "toolCalls": [
                    {
                        "name": "set_document_properties",
                        "arguments": {"author": "预检作者"},
                    },
                    {
                        "name": "manage_style",
                        "arguments": {"name": "正文", "align": "JUSTIFY"},
                    },
                ],
            },
            claims,
        )

        self.assertTrue(result["valid"])
        self.assertEqual(result["toolCalls"], 2)
        self.assertEqual(len(result["argumentNormalizations"]), 2)
        self.assertEqual(copilot_server.BRIDGE_COMMANDS, commands_before)
        self.assertEqual(copilot_server.BRIDGE_REQUESTS, requests_before)
        self.assertNotIn(
            "预检作者",
            json.dumps(result["argumentNormalizations"], ensure_ascii=False),
        )

    def test_slides_validate_batch_is_read_only_and_uses_full_schema(self):
        self.register(
            file_name="demo.pptx",
            file_type="pptx",
            editor_type="slide",
        )
        with copilot_server.BRIDGE_CONDITION:
            copilot_server.BRIDGE_SESSIONS["http-session:test"]["state"]["capabilities"]["tools"] = [
                "slides_add_slide",
                "slides_add_textbox",
                "slides_validate_layout",
            ]
        claims = {
            "fileName": "demo.pptx",
            "fileType": "pptx",
            "editorType": "slide",
            "userId": "uid-1",
        }
        commands_before = dict(copilot_server.BRIDGE_COMMANDS)
        requests_before = dict(copilot_server.BRIDGE_REQUESTS)

        result = copilot_server.bridge_validate(
            {
                "toolCalls": [
                    {
                        "name": "add_slide",
                        "arguments": {"masterIndex": 1, "layoutIndex": 2},
                    },
                    {
                        "name": "add_textbox",
                        "arguments": {
                            "slide": 1,
                            "paragraphs": [
                                {
                                    "text": "提升交付效率",
                                    "listType": "bullet",
                                    "fontSize": 18,
                                }
                            ],
                        },
                    },
                ],
            },
            claims,
        )

        self.assertTrue(result["valid"])
        self.assertEqual(result["editorType"], "slide")
        self.assertEqual(result["toolCalls"], 2)
        self.assertEqual(result["argumentNormalizations"], [])
        self.assertEqual(copilot_server.BRIDGE_COMMANDS, commands_before)
        self.assertEqual(copilot_server.BRIDGE_REQUESTS, requests_before)

        with self.assertRaises(copilot_server.BridgeError) as raised:
            copilot_server.bridge_validate(
                {
                    "toolCalls": [
                        {
                            "name": "add_textbox",
                            "arguments": {
                                "slide": 1,
                                "text": "普通文本",
                                "paragraphs": [{"text": "重复内容入口"}],
                            },
                        },
                        {
                            "name": "validate_layout",
                            "arguments": {"unknownRule": True},
                        },
                    ],
                },
                claims,
            )
        self.assertEqual(raised.exception.code, "INVALID_TOOL_ARGUMENTS")
        self.assertEqual(raised.exception.details["completedToolCalls"], 0)
        self.assertFalse(raised.exception.details["partialMutationPossible"])
        paths = {
            error["path"]
            for error in raised.exception.details["validationErrors"]
        }
        self.assertIn("arguments", paths)
        self.assertIn("arguments.unknownRule", paths)
        self.assertEqual(copilot_server.BRIDGE_COMMANDS, commands_before)
        self.assertEqual(copilot_server.BRIDGE_REQUESTS, requests_before)

        with self.assertRaises(copilot_server.BridgeError) as policy_error:
            copilot_server.bridge_validate(
                {
                    "toolCalls": [
                        {
                            "name": "validate_layout",
                            "arguments": {
                                "minFontSize": 9,
                                "fontPolicy": {
                                    "id": "policy/conflict",
                                    "source": "user",
                                    "defaultMinPt": 12,
                                    "rules": [
                                        {"name": "title", "role": "contentTitle", "minPt": 28},
                                        {"name": "title", "role": "body", "minPt": 16},
                                    ],
                                },
                            },
                        }
                    ],
                },
                claims,
            )
        policy_paths = {
            error["path"]
            for error in policy_error.exception.details["validationErrors"]
        }
        self.assertIn("arguments.fontPolicy", policy_paths)
        self.assertIn("arguments.fontPolicy.rules[1].name", policy_paths)

    def test_sheets_validate_batch_rejects_semantic_errors_without_editor_commands(self):
        self.register(
            file_name="demo.xlsx",
            file_type="xlsx",
            editor_type="cell",
        )
        tools = [
            "sheets_set_values",
            "sheets_set_formula",
            "sheets_manage_validation",
            "sheets_manage_table",
            "sheets_manage_freeze_panes",
        ]
        with copilot_server.BRIDGE_CONDITION:
            copilot_server.BRIDGE_SESSIONS["http-session:test"]["state"]["capabilities"]["tools"] = tools
        claims = {
            "fileName": "demo.xlsx",
            "fileType": "xlsx",
            "editorType": "cell",
            "userId": "uid-1",
        }
        commands_before = dict(copilot_server.BRIDGE_COMMANDS)
        requests_before = dict(copilot_server.BRIDGE_REQUESTS)

        valid = copilot_server.bridge_validate(
            {
                "toolCalls": [
                    {
                        "name": "set_values",
                        "arguments": {
                            "sheet": "Sheet1",
                            "range": "A1:B2",
                            "values": [["A", "B"], [1, 2]],
                        },
                    },
                    {
                        "name": "set_formula",
                        "arguments": {
                            "sheet": "Sheet1",
                            "range": "selection",
                            "formula": "=SUM(A1:B1)",
                        },
                    },
                    {
                        "name": "manage_freeze_panes",
                        "arguments": {
                            "sheet": "Sheet1",
                            "action": "freezeRows",
                            "count": 2,
                        },
                    },
                ],
            },
            claims,
        )
        self.assertTrue(valid["valid"])
        self.assertEqual(valid["editorType"], "cell")
        self.assertEqual(valid["toolCalls"], 3)

        invalid_calls = [
            {
                "name": "sheets_set_values",
                "arguments": {"range": "A1:B2", "values": 1},
            },
            {
                "name": "sheets_set_values",
                "arguments": {"range": "not-a-range", "values": 1},
            },
            {
                "name": "sheets_set_values",
                "arguments": {"range": "XFE1", "values": 1},
            },
            {
                "name": "sheets_set_formula",
                "arguments": {"range": "A1:B2", "formula": [["=1"], ["=2"]]},
            },
            {
                "name": "sheets_manage_validation",
                "arguments": {
                    "action": "add",
                    "range": "A1:A2",
                    "type": "wholeNumber",
                    "formula1": 1,
                },
            },
            {
                "name": "sheets_manage_table",
                "arguments": {
                    "action": "create",
                    "tableMode": "basic",
                    "range": "A1:B2",
                    "name": "StructuredName",
                },
            },
            {
                "name": "sheets_manage_freeze_panes",
                "arguments": {"action": "freezeRows", "count": 0},
            },
            {
                "name": "sheets_manage_freeze_panes",
                "arguments": {"action": "freezeAt"},
            },
        ]
        for call in invalid_calls:
            with self.subTest(tool=call["name"]):
                with self.assertRaises(copilot_server.BridgeError) as raised:
                    copilot_server.bridge_validate(
                        {"toolCalls": self.public_calls("cell", [call])},
                        claims,
                    )
                self.assertEqual(raised.exception.status, 422)
                self.assertEqual(raised.exception.code, "INVALID_TOOL_ARGUMENTS")
                self.assertEqual(
                    raised.exception.details["completedToolCalls"],
                    0,
                )
                self.assertFalse(
                    raised.exception.details["partialMutationPossible"],
                )

        self.assertEqual(copilot_server.BRIDGE_COMMANDS, commands_before)
        self.assertEqual(copilot_server.BRIDGE_REQUESTS, requests_before)

    def test_sheets_validate_batch_rejects_known_runtime_capability_gaps(self):
        self.register(
            file_name="demo.xlsx",
            file_type="xlsx",
            editor_type="cell",
        )
        capabilities = {
            "tools": [
                "sheets_manage_table",
                "sheets_manage_range",
                "sheets_set_array_formula",
                "sheets_manage_validation",
                "sheets_manage_comments",
                "sheets_manage_freeze_panes",
                "sheets_inspect_charts",
                "sheets_update_chart",
                "sheets_add_chart",
                "sheets_delete_chart",
            ],
            "features": {
                "sheets": {
                    "nativeTables": {"create": False, "inspect": False},
                    "rangeStyleTables": {"create": True},
                    "conditionalFormatting": {"create": True},
                    "rangeFill": {
                        "fillDown": False,
                        "fillUp": False,
                        "fillLeft": False,
                        "fillRight": False,
                    },
                    "arrayFormula": {"set": False},
                    "validation": {"manage": False},
                    "comments": {
                        "create": False,
                        "inspect": True,
                        "update": False,
                        "delete": False,
                    },
                    "freezePanes": {"inspect": True, "manage": False},
                    "charts": {
                        "create": True,
                        "inspect": False,
                        "update": False,
                        "addSeriesOnCreate": False,
                        "delete": False,
                    },
                },
            },
        }
        with copilot_server.BRIDGE_CONDITION:
            copilot_server.BRIDGE_SESSIONS["http-session:test"]["state"]["capabilities"] = capabilities
        claims = {
            "fileName": "demo.xlsx",
            "fileType": "xlsx",
            "editorType": "cell",
            "userId": "uid-1",
        }

        valid = copilot_server.bridge_validate(
            {
                "toolCalls": [{
                    "name": "manage_table",
                    "arguments": {
                        "action": "create",
                        "tableMode": "rangeStyle",
                        "range": "A1:B3",
                    },
                }],
            },
            claims,
        )
        self.assertTrue(valid["valid"])

        invalid_calls = [
            (
                {
                    "name": "sheets_manage_range",
                    "arguments": {"action": "fillDown", "range": "A1:A3"},
                },
                "sheets.rangeFill.fillDown",
            ),
            (
                {
                    "name": "sheets_set_array_formula",
                    "arguments": {"range": "A1:A3", "formula": "=ROW(A1:A3)"},
                },
                "sheets.arrayFormula.set",
            ),
            (
                {
                    "name": "sheets_manage_validation",
                    "arguments": {"action": "delete", "range": "A1:A3"},
                },
                "sheets.validation.manage",
            ),
            (
                {
                    "name": "sheets_manage_comments",
                    "arguments": {"action": "add", "range": "A1", "text": "note"},
                },
                "sheets.comments.create",
            ),
            (
                {
                    "name": "sheets_manage_freeze_panes",
                    "arguments": {"action": "freezeRows", "count": 1},
                },
                "sheets.freezePanes.manage",
            ),
            (
                {
                    "name": "sheets_inspect_charts",
                    "arguments": {},
                },
                "sheets.charts.inspect",
            ),
            (
                {
                    "name": "sheets_update_chart",
                    "arguments": {"chartIndex": 0, "title": "Updated"},
                },
                "sheets.charts.update",
            ),
            (
                {
                    "name": "sheets_manage_table",
                    "arguments": {
                        "action": "create",
                        "tableMode": "structured",
                        "range": "A1:B3",
                    },
                },
                "sheets.nativeTables.create",
            ),
            (
                {
                    "name": "sheets_add_chart",
                    "arguments": {
                        "range": "A1:B3",
                        "addSeries": [{"name": "S2", "valuesRange": "C1:C3"}],
                    },
                },
                "sheets.charts.addSeriesOnCreate",
            ),
            (
                {
                    "name": "sheets_delete_chart",
                    "arguments": {"chartIndex": 0},
                },
                "sheets.charts.delete",
            ),
        ]
        for call, feature in invalid_calls:
            with self.subTest(feature=feature):
                with self.assertRaises(copilot_server.BridgeError) as raised:
                    copilot_server.bridge_validate(
                        {"toolCalls": self.public_calls("cell", [call])},
                        claims,
                    )
                self.assertEqual(raised.exception.status, 501)
                self.assertEqual(raised.exception.code, "SHEETS_API_UNSUPPORTED")
                self.assertEqual(raised.exception.details["feature"], feature)
                self.assertEqual(raised.exception.details["completedToolCalls"], 0)
                self.assertFalse(raised.exception.details["partialMutationPossible"])

    def test_sheets_batch_rejects_new_sheet_dependency_before_enqueue(self):
        self.register(file_name="demo.xlsx", file_type="xlsx", editor_type="cell")
        capabilities = copilot_server.BRIDGE_SESSIONS["http-session:test"]["state"]["capabilities"]
        capabilities["tools"] = ["sheets_add_sheet", "sheets_set_values"]
        claims = {
            "fileName": "demo.xlsx",
            "fileType": "xlsx",
            "editorType": "cell",
            "userId": "uid-1",
        }
        commands_before = dict(copilot_server.BRIDGE_COMMANDS)
        with self.assertRaises(copilot_server.BridgeError) as raised:
            copilot_server.bridge_validate(
                {
                    "toolCalls": [
                        {"name": "add_sheet", "arguments": {"name": "NewData"}},
                        {
                            "name": "set_values",
                            "arguments": {"sheet": "NewData", "range": "A1", "values": 1},
                        },
                    ],
                },
                claims,
            )
        self.assertEqual(raised.exception.status, 409)
        self.assertEqual(raised.exception.code, "BATCH_DEPENDENCY_REQUIRES_SPLIT")
        self.assertEqual(raised.exception.details["completedToolCalls"], 0)
        self.assertFalse(raised.exception.details["partialMutationPossible"])
        self.assertEqual(copilot_server.BRIDGE_COMMANDS, commands_before)

    def test_word_batch_collects_all_argument_errors_before_enqueue(self):
        self.register()
        tool_calls = [
            {
                "name": "word_manage_section",
                "arguments": {"action": "create", "internalNote": "敏感正文"},
            },
            {
                "name": "word_format_table_advanced",
                "arguments": {"tableIndex": 1, "repeatHeader": True},
            },
            {
                "name": "word_manage_fields",
                "arguments": {"action": "add"},
            },
            {
                "name": "word_insert_page_break",
                "arguments": {"position": "end"},
            },
            {
                "name": "word_edit_table",
                "arguments": {
                    "tableIndex": 1,
                    "action": "mergeCells",
                    "startRow": 1,
                    "endRow": 2,
                },
            },
            {
                "name": "word_set_header_footer",
                "arguments": {"kind": "header", "action": "append"},
            },
            {
                "name": "word_add_image",
                "arguments": {
                    "source": {"type": "url", "url": "http://blocked.example/image.png"},
                },
            },
            {
                "name": "word_set_table_cell",
                "arguments": {
                    "tableIndex": 1,
                    "row": 1,
                    "column": 1,
                    "color": "not-a-color",
                },
            },
        ]
        with copilot_server.BRIDGE_CONDITION:
            session = copilot_server.BRIDGE_SESSIONS["http-session:test"]
            session["state"]["capabilities"]["tools"] = [
                call["name"] for call in tool_calls
            ]
        claims = copilot_server.verify_editor_jwt(self.editor_token())
        commands_before = dict(copilot_server.BRIDGE_COMMANDS)
        requests_before = dict(copilot_server.BRIDGE_REQUESTS)

        with self.assertRaises(copilot_server.BridgeError) as raised:
            copilot_server.bridge_execute(
                {
                    "method": "executeBatch",
                    "requestId": "invalid-word-batch",
                    "toolCalls": self.public_calls("word", tool_calls),
                },
                claims,
            )

        error = raised.exception
        self.assertEqual(error.status, 422)
        self.assertEqual(error.code, "INVALID_TOOL_ARGUMENTS")
        self.assertEqual(error.details["completedToolCalls"], 0)
        self.assertFalse(error.details["partialMutationPossible"])
        self.assertGreaterEqual(len(error.details["validationErrors"]), 10)
        self.assertEqual(copilot_server.BRIDGE_COMMANDS, commands_before)
        self.assertEqual(copilot_server.BRIDGE_REQUESTS, requests_before)
        encoded = json.dumps(error.details, ensure_ascii=False)
        self.assertNotIn("敏感正文", encoded)
        self.assertIn('"keyword": "anyOf"', encoded)
        self.assertIn("arguments.row", encoded)
        self.assertIn("arguments.instruction", encoded)
        self.assertIn("arguments.target", encoded)
        self.assertIn("arguments.rowStart", encoded)
        self.assertIn("arguments.source", encoded)
        self.assertIn("arguments.color", encoded)

    def test_resume_token_restores_a_cleaned_session_without_editor_token(self):
        registration = self.register()
        with copilot_server.BRIDGE_CONDITION:
            copilot_server.BRIDGE_SESSIONS.clear()

        restored = copilot_server.bridge_register(
            {
                "sessionId": "http-session:test",
                "resumeToken": registration["resumeToken"],
                "state": self.editor_state(),
            }
        )

        self.assertNotEqual(restored["relayKey"], registration["relayKey"])
        self.assertTrue(restored["resumeToken"])
        self.assertEqual(restored["session"]["documentKey"], "document-key-v1")

    def test_expired_resume_token_requires_a_page_reload(self):
        with mock.patch.object(copilot_server, "BRIDGE_RESUME_TOKEN_TTL_SECONDS", -1):
            registration = self.register()

        with self.assertRaises(copilot_server.BridgeError) as raised:
            copilot_server.bridge_register(
                {
                    "sessionId": "http-session:test",
                    "resumeToken": registration["resumeToken"],
                    "state": self.editor_state(),
                }
            )

        self.assertEqual(raised.exception.code, "BRIDGE_RESUME_TOKEN_EXPIRED")

    def test_resume_token_cannot_be_replayed_for_another_session(self):
        registration = self.register()

        with self.assertRaises(copilot_server.BridgeError) as raised:
            copilot_server.bridge_register(
                {
                    "sessionId": "http-session:other",
                    "resumeToken": registration["resumeToken"],
                    "state": self.editor_state(),
                }
            )

        self.assertEqual(raised.exception.code, "INVALID_BRIDGE_RESUME_TOKEN")


    def test_internal_session_reports_ready_authoritative_contract_state(self):
        self.register(
            document_id="11111111-1111-4111-8111-111111111111",
            ready=True,
        )

        response = copilot_server.internal_bridge_session("http-session:test")

        self.assertTrue(response["ok"])
        self.assertTrue(response["ready"])
        self.assertTrue(response["authoritative"])
        self.assertEqual(response["editorType"], "word")
        self.assertEqual(response["contractVersion"], copilot_server.PUBLIC_API_CONTRACT_VERSION)
        self.assertEqual(response["contractSha256"], copilot_server.PUBLIC_API_CONTRACT_SHA256)
        self.assertIsInstance(response["lastSeen"], float)

    def test_internal_session_exposes_registering_state_and_rejects_superseded_session(self):
        self.register(ready=False)
        registering = copilot_server.internal_bridge_session("http-session:test")
        self.assertFalse(registering["ready"])

        self.register(session_id="http-session:newer", ready=True)
        with self.assertRaises(copilot_server.BridgeError) as raised:
            copilot_server.internal_bridge_session("http-session:test")
        self.assertEqual(raised.exception.code, "INVALID_RELAY_SESSION")
        self.assertTrue(
            copilot_server.internal_bridge_session("http-session:newer")["ready"]
        )


    def test_new_registration_supersedes_old_document_version(self):
        first = self.register("http-session:first", "document-key-v1")
        claims = copilot_server.verify_editor_jwt(self.editor_token("document-key-v1"))

        second = self.register("http-session:second", "document-key-v2")

        with self.assertRaises(copilot_server.BridgeError) as raised:
            copilot_server.bridge_poll(
                {
                    "sessionId": "http-session:first",
                    "relayKey": first["relayKey"],
                    "state": self.editor_state("document-key-v1"),
                    "timeoutMs": 1,
                }
            )
        self.assertEqual(raised.exception.code, "SESSION_SUPERSEDED")

        with copilot_server.BRIDGE_CONDITION:
            selected_id, selected = copilot_server.bridge_select_session_locked(
                "http-session:first",
                claims,
            )
        self.assertEqual(selected_id, "http-session:second")
        self.assertEqual(selected["documentKey"], "document-key-v2")
        self.assertGreater(
            selected["generation"],
            first["session"]["generation"],
        )
        self.assertTrue(second["session"]["authoritative"])

    def test_stale_session_hint_routes_execute_to_latest_page(self):
        self.register("http-session:first", "document-key-v1")
        claims = copilot_server.verify_editor_jwt(self.editor_token("document-key-v1"))
        second = self.register("http-session:second", "document-key-v2")
        holder = {}
        request = {
            "sessionId": "http-session:first",
            "method": "executeTool",
            "name": "inspect",
            "arguments": {"maxChars": 500},
            "requestId": "stale-session-hint-1",
            "timeoutMs": 2000,
        }

        worker = threading.Thread(
            target=lambda: holder.update(
                result=copilot_server.bridge_execute(request, claims)
            )
        )
        worker.start()
        command = copilot_server.bridge_poll(
            {
                "sessionId": "http-session:second",
                "relayKey": second["relayKey"],
                "state": self.editor_state("document-key-v2"),
                "timeoutMs": 1000,
            }
        )["command"]
        copilot_server.bridge_result(
            {
                "sessionId": "http-session:second",
                "relayKey": second["relayKey"],
                "commandId": command["commandId"],
                "state": self.editor_state("document-key-v2"),
                "ok": True,
                "result": {"text": "latest page"},
            }
        )
        worker.join(timeout=2)

        self.assertFalse(worker.is_alive())
        self.assertEqual(holder["result"]["session"]["sessionId"], "http-session:second")
        self.assertEqual(holder["result"]["result"]["text"], "latest page")

    def test_latest_page_close_does_not_restore_superseded_page(self):
        first = self.register("http-session:first", "document-key-v1")
        claims = copilot_server.verify_editor_jwt(self.editor_token("document-key-v1"))
        second = self.register("http-session:second", "document-key-v2")

        copilot_server.bridge_unregister(
            {
                "sessionId": "http-session:second",
                "relayKey": second["relayKey"],
                "state": self.editor_state("document-key-v2"),
            }
        )

        with (
            mock.patch.object(copilot_server, "BRIDGE_DISCOVERY_WAIT_SECONDS", 0),
            copilot_server.BRIDGE_CONDITION,
            self.assertRaises(copilot_server.BridgeError) as raised,
        ):
            copilot_server.bridge_select_session_locked("http-session:first", claims)
        self.assertEqual(raised.exception.code, "NO_ACTIVE_EDITOR")

        with self.assertRaises(copilot_server.BridgeError) as old_register:
            self.register("http-session:first", "document-key-v1")
        self.assertEqual(old_register.exception.code, "SESSION_SUPERSEDED")

        refreshed = self.register("http-session:refreshed", "document-key-v3")
        with copilot_server.BRIDGE_CONDITION:
            selected_id, _ = copilot_server.bridge_select_session_locked(
                "http-session:first",
                claims,
            )
        self.assertEqual(selected_id, "http-session:refreshed")
        self.assertTrue(refreshed["session"]["authoritative"])
        self.assertIn("http-session:first", copilot_server.BRIDGE_SUPERSEDED)
        self.assertNotIn("http-session:first", copilot_server.BRIDGE_SESSIONS)
        self.assertEqual(first["session"]["sessionId"], "http-session:first")

    def test_stable_identity_cannot_select_another_document(self):
        self.register()
        foreign_claims = {
            "fileName": "another.docx",
            "fileType": "docx",
            "editorType": "word",
            "userId": "uid-1",
        }

        with (
            mock.patch.object(copilot_server, "BRIDGE_DISCOVERY_WAIT_SECONDS", 0),
            copilot_server.BRIDGE_CONDITION,
            self.assertRaises(copilot_server.BridgeError) as raised,
        ):
            copilot_server.bridge_select_session_locked("", foreign_claims)

        self.assertEqual(raised.exception.code, "NO_ACTIVE_EDITOR")

    def test_undelivered_command_moves_once_to_new_authoritative_page(self):
        first = self.register("http-session:first", "document-key-v1")
        claims = copilot_server.verify_editor_jwt(self.editor_token("document-key-v1"))
        holder = {}
        request = {
            "method": "executeTool",
            "name": "inspect",
            "arguments": {"maxChars": 500},
            "requestId": "turn-handoff-undelivered-1",
            "timeoutMs": 2000,
        }

        worker = threading.Thread(
            target=lambda: holder.update(result=copilot_server.bridge_execute(request, claims))
        )
        worker.start()
        deadline = time.time() + 1
        while not copilot_server.BRIDGE_COMMANDS and time.time() < deadline:
            time.sleep(0.001)

        second = self.register("http-session:second", "document-key-v2")
        command = copilot_server.bridge_poll(
            {
                "sessionId": "http-session:second",
                "relayKey": second["relayKey"],
                "state": self.editor_state("document-key-v2"),
                "timeoutMs": 1000,
            }
        )["command"]
        self.assertEqual(command["requestId"], "turn-handoff-undelivered-1")
        copilot_server.bridge_result(
            {
                "sessionId": "http-session:second",
                "relayKey": second["relayKey"],
                "commandId": command["commandId"],
                "state": self.editor_state("document-key-v2"),
                "ok": True,
                "result": {"text": "moved once"},
            }
        )
        worker.join(timeout=2)

        self.assertFalse(worker.is_alive())
        self.assertEqual(holder["result"]["result"]["text"], "moved once")
        self.assertEqual(holder["result"]["session"]["sessionId"], "http-session:second")
        self.assertEqual(len(copilot_server.BRIDGE_COMMANDS), 1)
        self.assertEqual(
            next(iter(copilot_server.BRIDGE_COMMANDS.values()))["sessionId"],
            "http-session:second",
        )
        with self.assertRaises(copilot_server.BridgeError) as old_poll:
            copilot_server.bridge_poll(
                {
                    "sessionId": "http-session:first",
                    "relayKey": first["relayKey"],
                    "state": self.editor_state("document-key-v1"),
                    "timeoutMs": 1,
                }
            )
        self.assertEqual(old_poll.exception.code, "SESSION_SUPERSEDED")

    def test_delivered_command_finishes_after_new_page_takes_authority(self):
        first = self.register("http-session:first", "document-key-v1")
        claims = copilot_server.verify_editor_jwt(self.editor_token("document-key-v1"))
        request = {
            "method": "executeTool",
            "name": "inspect",
            "arguments": {"maxChars": 500},
            "requestId": "turn-handoff-tool-1",
            "timeoutMs": 2000,
        }
        holder = {}

        worker = threading.Thread(
            target=lambda: holder.update(result=copilot_server.bridge_execute(request, claims))
        )
        worker.start()
        command = copilot_server.bridge_poll(
            {
                "sessionId": "http-session:first",
                "relayKey": first["relayKey"],
                "state": self.editor_state("document-key-v1"),
                "timeoutMs": 1000,
            }
        )["command"]

        self.register("http-session:second", "document-key-v2")
        accepted = copilot_server.bridge_result(
            {
                "sessionId": "http-session:first",
                "relayKey": first["relayKey"],
                "commandId": command["commandId"],
                "state": self.editor_state("document-key-v1"),
                "ok": True,
                "result": {"text": "completed on first page"},
            }
        )
        worker.join(timeout=2)

        self.assertTrue(accepted["accepted"])
        self.assertFalse(worker.is_alive())
        self.assertEqual(holder["result"]["result"]["text"], "completed on first page")

    def test_command_round_trip_and_request_id_cache(self):
        registration = self.register()
        claims = copilot_server.verify_editor_jwt(self.editor_token())
        request = {
            "method": "executeTool",
            "name": "inspect",
            "arguments": {"maxChars": 1000},
            "requestId": "turn-1-tool-1",
            "timeoutMs": 2000,
        }
        holder = {}

        def execute():
            holder["result"] = copilot_server.bridge_execute(request, claims)

        worker = threading.Thread(target=execute)
        worker.start()
        command = copilot_server.bridge_poll(
            {
                "sessionId": "http-session:test",
                "relayKey": registration["relayKey"],
                "state": self.editor_state(),
                "timeoutMs": 1000,
            }
        )["command"]
        self.assertEqual(command["method"], "executeTool")
        self.assertEqual(command["params"]["name"], "word_inspect")

        copilot_server.bridge_result(
            {
                "sessionId": "http-session:test",
                "relayKey": registration["relayKey"],
                "commandId": command["commandId"],
                "state": self.editor_state(),
                "ok": True,
                "result": {"text": "current document"},
            }
        )
        worker.join(timeout=2)

        self.assertFalse(worker.is_alive())
        self.assertEqual(holder["result"]["result"]["text"], "current document")
        self.assertEqual(
            set(holder["result"]["timingsMs"]),
            {"queueWait", "editorRoundTrip", "total"},
        )
        self.assertGreaterEqual(holder["result"]["timingsMs"]["total"], 0)
        cached = copilot_server.bridge_execute(request, claims)
        self.assertTrue(cached["cached"])
        self.assertEqual(cached["result"]["text"], "current document")
        self.assertEqual(cached["timingsMs"], holder["result"]["timingsMs"])

        conflicting = dict(request)
        conflicting["arguments"] = {"maxChars": 2000}
        with self.assertRaises(copilot_server.BridgeError) as raised:
            copilot_server.bridge_execute(conflicting, claims)
        self.assertEqual(raised.exception.code, "REQUEST_ID_CONFLICT")
        self.assertEqual(raised.exception.status, 409)
        self.assertEqual(
            raised.exception.details,
            {
                "requestId": request["requestId"],
                "sessionId": "http-session:test",
                "retryable": False,
                "reuseAllowed": False,
                "requiredAction": "use_new_request_id",
                "partialMutationPossible": False,
            },
        )

        self.register("http-session:next", "document-key-v2")
        cached_after_handoff = copilot_server.bridge_execute(request, claims)
        self.assertTrue(cached_after_handoff["cached"])
        self.assertEqual(cached_after_handoff["result"]["text"], "current document")

    def test_registration_logs_only_sanitized_startup_timings(self):
        state = self.editor_state()
        state["_diagnostics"] = {
            "startup": {
                "hostScriptMs": 12.4,
                "windowLoadMs": "25",
                "editorFrameSeenMs": -1,
                "editorFrameLoadMs": float("inf"),
                "appReadyMs": True,
                "documentReadyMs": 49_001,
                "bridgeReadyMs": 49_800,
                "secret": "敏感正文",
            }
        }
        output = io.StringIO()
        with mock.patch("sys.stdout", output):
            registration = copilot_server.bridge_register(
                {
                    "sessionId": "http-session:startup",
                    "editorToken": self.editor_token(),
                    "state": state,
                }
            )

        startup_line = next(
            line for line in output.getvalue().splitlines()
            if line.startswith("[bridge-startup] ")
        )
        event = json.loads(startup_line.split(" ", 1)[1])
        self.assertEqual(
            event,
            {
                "event": "registered",
                "identity": copilot_server.bridge_identity_digest(
                    ("demo.docx", "docx", "word", "uid-1")
                ),
                "sessionId": "http-session:startup",
                "hostScriptMs": 12,
                "documentReadyMs": 49001,
                "bridgeReadyMs": 49800,
            },
        )
        self.assertNotIn("_diagnostics", registration["session"])
        self.assertNotIn("敏感正文", output.getvalue())


class BridgeLoggingTests(unittest.TestCase):
    def test_bridge_command_log_contains_sanitized_timings_once(self):
        identity = ("demo.docx", "docx", "word", "uid-1")
        command = {
            "requestId": "turn-4-batch",
            "sessionId": "http-session:test",
            "identity": identity,
            "method": "executeBatch",
            "params": {
                "toolCalls": [
                    {
                        "name": "word_append_paragraph",
                        "arguments": {"text": "敏感正文"},
                    }
                    for _ in range(10)
                ]
            },
            "createdMonotonic": 10.0,
            "firstDeliveredMonotonic": 10.125,
            "timingLogged": False,
        }

        output = io.StringIO()
        with mock.patch("sys.stdout", output):
            copilot_server.log_bridge_command(command, "completed", 11.5)
            copilot_server.log_bridge_command(command, "completed", 12.0)

        lines = output.getvalue().splitlines()
        self.assertEqual(len(lines), 1)
        prefix, encoded = lines[0].split(" ", 1)
        event = json.loads(encoded)
        self.assertEqual(prefix, "[bridge-command]")
        self.assertEqual(
            event,
            {
                "event": "completed",
                "requestId": "turn-4-batch",
                "sessionId": "http-session:test",
                "identity": copilot_server.bridge_identity_digest(identity),
                "method": "executeBatch",
                "toolCount": 10,
                "queueWaitMs": 125,
                "editorRoundTripMs": 1375,
                "totalMs": 1500,
            },
        )
        self.assertNotIn("敏感正文", output.getvalue())

    def test_bridge_error_log_includes_only_validated_correlation_fields(self):
        handler = object.__new__(copilot_server.Handler)
        handler.path = "/bridge/execute?debug=true"
        handler.client_address = ("127.0.0.1", 12345)
        error = copilot_server.BridgeError(
            409,
            "EXECUTION_FAILED",
            "包含敏感正文的错误消息",
            {"phase": "word-command", "secret": "敏感正文"},
        )

        output = io.StringIO()
        with mock.patch("sys.stdout", output):
            handler.log_bridge_error(
                error,
                {
                    "requestId": "turn-3-set-text",
                    "sessionId": "http-session:test",
                    "arguments": {"text": "敏感正文"},
                },
            )

        prefix, encoded = output.getvalue().strip().split(" ", 1)
        event = json.loads(encoded)
        self.assertEqual(prefix, "[bridge-error]")
        self.assertEqual(
            event,
            {
                "path": "/bridge/execute",
                "status": 409,
                "code": "EXECUTION_FAILED",
                "client": "127.0.0.1",
                "sessionId": "http-session:test",
                "requestId": "turn-3-set-text",
                "phase": "word-command",
            },
        )
        self.assertNotIn("敏感正文", output.getvalue())

    def test_bridge_error_log_omits_untrusted_ids_and_phases(self):
        handler = object.__new__(copilot_server.Handler)
        handler.path = "/bridge/execute"
        handler.client_address = ("127.0.0.1", 12345)
        error = copilot_server.BridgeError(
            409,
            "EXECUTION_FAILED",
            "failed",
            {"phase": "untrusted-phase"},
        )

        output = io.StringIO()
        with mock.patch("sys.stdout", output):
            handler.log_bridge_error(
                error,
                {
                    "requestId": "invalid request id",
                    "sessionId": "invalid session id",
                },
            )

        _, encoded = output.getvalue().strip().split(" ", 1)
        event = json.loads(encoded)
        self.assertNotIn("requestId", event)
        self.assertNotIn("sessionId", event)
        self.assertNotIn("phase", event)


if __name__ == "__main__":
    unittest.main()
