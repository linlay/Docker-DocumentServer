import base64
import http.client
import io
import json
import os
import socket
import struct
import tempfile
import threading
import time
import unittest
from unittest import mock

import copilot_server


class ContractAlignmentTests(unittest.TestCase):
    def test_capability_gateway_defaults_to_stable_guest_without_forcing_a_user(self):
        config_path = os.path.join(os.path.dirname(__file__), "nginx-ds-example.conf")
        with open(config_path, encoding="utf-8") as stream:
            config = stream.read()

        self.assertIn(
            """sub_filter '<option value="uid-0">Anonymous</option>' '<option value="uid-0" selected>访客</option>';""",
            config,
        )
        self.assertIn("OnlyOfficeLocalGuest.prepareExample(config)", config)
        self.assertNotRegex(config, r"editorConfig\.user\s*=")
        self.assertNotIn("uid-1", config)
        self.assertNotIn("uid-2", config)
        self.assertNotIn("uid-3", config)

    def test_capability_gateway_enables_native_compact_toolbar_before_editor_events(self):
        config_path = os.path.join(os.path.dirname(__file__), "nginx-ds-example.conf")
        with open(config_path, encoding="utf-8") as stream:
            config = stream.read()

        compact_toolbar = "config.editorConfig.customization.compactToolbar = true;"
        self.assertEqual(config.count(compact_toolbar), 1)
        self.assertLess(config.index(compact_toolbar), config.rindex("config.events = {"))

    def test_capability_gateway_uses_simplified_chinese_editor_config(self):
        config_path = os.path.join(os.path.dirname(__file__), "nginx-ds-example.conf")
        with open(config_path, encoding="utf-8") as stream:
            config = stream.read()

        self.assertIn(
            """sub_filter '<option value="zh">zh</option>' '<option value="zh" selected>zh</option>';""",
            config,
        )
        self.assertIn("absolute_redirect off;", config)
        self.assertIn("&lang=zh;", config)
        self.assertNotIn('config.editorConfig.lang = "zh"', config)

    def test_legacy_example_is_closed_and_storage_is_loopback_only(self):
        config_path = os.path.join(os.path.dirname(__file__), "nginx-ds-example.conf")
        with open(config_path, encoding="utf-8") as stream:
            config = stream.read()

        self.assertIn("location = /example {", config)
        self.assertIn("location ^~ /example/ {", config)
        self.assertEqual(config.count("return 404;"), 2)
        self.assertIn("location ^~ /__document_storage/ {", config)
        self.assertIn("allow 127.0.0.1;", config)
        self.assertIn("deny all;", config)
        self.assertIn("internal;", config)
        self.assertIn("proxy_set_header X-Forwarded-Host 127.0.0.1;", config)
        self.assertIn("proxy_set_header X-Forwarded-Proto http;", config)
        with open(
            os.path.join(os.path.dirname(__file__), "copilot_server.py"),
            encoding="utf-8",
        ) as stream:
            self.assertIn("X-Accel-Redirect", stream.read())

    def test_gateway_exposes_three_create_and_capability_routes(self):
        config_path = os.path.join(os.path.dirname(__file__), "nginx-ds-example.conf")
        with open(config_path, encoding="utf-8") as stream:
            config = stream.read()

        self.assertIn("docx|xlsx|pptx", config)
        self.assertIn("^/new-(docx|xlsx|pptx)$", config)
        self.assertIn("/__document_editor/", config)
        self.assertIn("proxy_pass http://127.0.0.1:3001;", config)
        self.assertIn("proxy_set_header X-Forwarded-For $remote_addr;", config)

    def test_compose_passes_document_admin_environment_to_copilot_process(self):
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

        self.assertIn('"DOCUMENT_ADMIN_USERNAME=$${DOCUMENT_ADMIN_USERNAME}"', compose)
        self.assertIn('"DOCUMENT_ADMIN_PASSWORD=$${DOCUMENT_ADMIN_PASSWORD}"', compose)
        self.assertIn("> /run/onlyoffice-copilot.env", compose)
        self.assertIn(
            '"./data/documents:/var/lib/onlyoffice/documentserver-example/files"',
            compose,
        )
        self.assertIn(
            "/var/lib/onlyoffice/documentserver-example/files",
            compose,
        )
        self.assertIn(
            "command=/bin/bash /opt/onlyoffice-copilot/run-copilot.sh",
            supervisor,
        )
        self.assertIn('export "$assignment"', runner)

    def test_static_asset_cache_revision_is_consistent(self):
        base_dir = os.path.dirname(__file__)
        revision = "0.6.0-rev4"
        paths = [
            "config.json",
            "index.html",
            "nginx-ds-example.conf",
            "README.md",
            "INTEGRATION.zh-CN.md",
        ]
        for relative_path in paths:
            with self.subTest(path=relative_path):
                with open(os.path.join(base_dir, relative_path), encoding="utf-8") as stream:
                    contents = stream.read()
                self.assertIn(revision, contents)
                self.assertNotIn("0.4.0-rev30", contents)

    def test_word_model_tools_match_the_public_contract(self):
        contract_path = os.path.join(os.path.dirname(__file__), "public-api.json")
        with open(contract_path, encoding="utf-8") as stream:
            contract = json.load(stream)

        contract_names = list(contract["tools"]["word"])
        model_names = [
            entry["function"]["name"]
            for entry in copilot_server.WORD_TOOLS
        ]

        self.assertEqual(model_names, contract_names)
        self.assertIn("word_add_image", model_names)
        for entry in copilot_server.WORD_TOOLS:
            function = entry["function"]
            self.assertTrue(function["description"])
            self.assertNotIn("$ref", json.dumps(function["parameters"]))

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
        normalized_calls, normalizations = copilot_server.normalize_word_tool_calls(
            compatibility_and_invalid_calls
        )
        word_errors = copilot_server.validate_word_tool_calls(
            compatibility_and_invalid_calls
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


class DocumentGatewayTests(unittest.TestCase):
    def setUp(self):
        self.storage = tempfile.TemporaryDirectory()
        self.templates = tempfile.TemporaryDirectory()
        for extension in copilot_server.DOCUMENT_TYPES:
            with open(
                os.path.join(self.templates.name, f"new.{extension}"),
                "wb",
            ) as stream:
                stream.write(f"blank-{extension}".encode())
        self.environment = mock.patch.dict(
            os.environ,
            {
                "DOCUMENT_STORAGE_DIR": self.storage.name,
                "DOCUMENT_TEMPLATE_ROOT": self.templates.name,
                "DOCUMENT_PUBLIC_ORIGIN": "https://office.test",
                "DOCUMENT_ADMIN_USERNAME": "document-admin",
                "DOCUMENT_ADMIN_PASSWORD": "correct horse battery staple",
            },
        )
        self.environment.start()
        with copilot_server.DOCUMENT_RATE_LIMIT_LOCK:
            copilot_server.DOCUMENT_RATE_LIMITS.clear()

    def tearDown(self):
        self.environment.stop()
        self.templates.cleanup()
        self.storage.cleanup()
        with copilot_server.DOCUMENT_RATE_LIMIT_LOCK:
            copilot_server.DOCUMENT_RATE_LIMITS.clear()

    def test_creates_all_supported_formats_with_capability_urls(self):
        for extension in copilot_server.DOCUMENT_TYPES:
            with self.subTest(extension=extension):
                created = copilot_server.create_document(extension)
                self.assertRegex(
                    created["documentId"],
                    copilot_server.DOCUMENT_UUID_PATTERN,
                )
                self.assertEqual(
                    created["fileName"],
                    f"{created['documentId']}.{extension}",
                )
                self.assertEqual(
                    created["editorUrl"],
                    f"https://office.test/{extension}/{created['documentId']}",
                )
                with open(
                    os.path.join(self.storage.name, created["fileName"]),
                    "rb",
                ) as stream:
                    self.assertEqual(stream.read(), f"blank-{extension}".encode())

    def test_uuid_collision_retries_without_overwriting_existing_file(self):
        first_id = "123e4567-e89b-42d3-a456-426614174000"
        second_id = "223e4567-e89b-42d3-a456-426614174000"
        existing = os.path.join(self.storage.name, f"{first_id}.docx")
        with open(existing, "wb") as stream:
            stream.write(b"existing")

        with mock.patch.object(
            copilot_server.uuid,
            "uuid4",
            side_effect=[copilot_server.uuid.UUID(first_id), copilot_server.uuid.UUID(second_id)],
        ):
            created = copilot_server.create_document("docx")

        self.assertEqual(created["documentId"], second_id)
        with open(existing, "rb") as stream:
            self.assertEqual(stream.read(), b"existing")

    def test_invalid_and_unknown_capability_ids_are_indistinguishable(self):
        for document_id in (
            "not-a-uuid",
            "123e4567-e89b-42d3-a456-426614174000",
        ):
            with self.subTest(document_id=document_id):
                with self.assertRaises(copilot_server.BridgeError) as raised:
                    copilot_server.document_path("docx", document_id)
                self.assertEqual(raised.exception.status, 404)
                self.assertEqual(raised.exception.code, "DOCUMENT_NOT_FOUND")

    def test_httpx_binding_helper_accepts_only_existing_uuid_file_names(self):
        created = copilot_server.create_document("docx")

        self.assertEqual(
            copilot_server.document_file_identity(created["fileName"]),
            ("docx", created["documentId"]),
        )
        for file_name in (
            "legacy-name.docx",
            f"{created['documentId']}.xlsx",
            "../" + created["fileName"],
        ):
            with self.subTest(file_name=file_name):
                with self.assertRaises(copilot_server.BridgeError) as raised:
                    copilot_server.document_file_identity(file_name)
                self.assertEqual(raised.exception.status, 404)
                self.assertEqual(raised.exception.code, "DOCUMENT_NOT_FOUND")

    def test_creation_rate_limit_allows_burst_then_refills(self):
        for _ in range(copilot_server.DOCUMENT_CREATE_BURST):
            copilot_server.enforce_document_creation_rate("192.0.2.10", now=100)

        with self.assertRaises(copilot_server.BridgeError) as raised:
            copilot_server.enforce_document_creation_rate("192.0.2.10", now=100)
        self.assertEqual(raised.exception.status, 429)
        self.assertEqual(raised.exception.details["retryAfter"], 6)

        copilot_server.enforce_document_creation_rate("192.0.2.10", now=106)

    def test_admin_list_only_contains_uuid_documents(self):
        first = copilot_server.create_document("docx")
        second = copilot_server.create_document("xlsx")
        with open(os.path.join(self.storage.name, "legacy-name.docx"), "wb") as stream:
            stream.write(b"legacy")
        os.makedirs(
            os.path.join(self.storage.name, f"{first['fileName']}-history"),
        )
        with open(os.path.join(self.storage.name, ".private"), "wb") as stream:
            stream.write(b"hidden")

        listed = copilot_server.list_uuid_documents()

        self.assertEqual(
            {item["fileName"] for item in listed},
            {first["fileName"], second["fileName"]},
        )
        page = copilot_server.admin_documents_html(listed)
        self.assertIn(
            f"<code>{first['editorUrl']}</code>",
            page,
        )
        self.assertNotIn("legacy-name.docx", page)

    def test_basic_auth_fails_closed_and_uses_constant_credentials(self):
        handler = mock.Mock()
        handler.headers = {}
        with self.assertRaises(copilot_server.BridgeError) as missing:
            copilot_server.require_document_admin(handler)
        self.assertEqual(missing.exception.status, 401)

        handler.headers = {
            "Authorization": "Basic "
            + base64.b64encode(b"document-admin:wrong").decode()
        }
        with self.assertRaises(copilot_server.BridgeError) as wrong:
            copilot_server.require_document_admin(handler)
        self.assertEqual(wrong.exception.status, 401)

        handler.headers = {
            "Authorization": "Basic "
            + base64.b64encode(
                b"document-admin:correct horse battery staple"
            ).decode()
        }
        copilot_server.require_document_admin(handler)

        with mock.patch.dict(
            os.environ,
            {"DOCUMENT_ADMIN_USERNAME": "", "DOCUMENT_ADMIN_PASSWORD": ""},
        ):
            with self.assertRaises(copilot_server.BridgeError) as unconfigured:
                copilot_server.require_document_admin(handler)
        self.assertEqual(unconfigured.exception.status, 503)

    def test_gateway_uses_internal_acceleration_without_redirecting_browser(self):
        created = copilot_server.create_document("pptx")
        handler = mock.Mock()

        copilot_server.document_gateway_response(
            handler,
            "pptx",
            created["documentId"],
        )

        handler.send_response.assert_called_once_with(200)
        handler.send_header.assert_any_call(
            "X-Accel-Redirect",
            f"/__document_editor/pptx/{created['documentId']}",
        )


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


class LocalGuestTokenTests(unittest.TestCase):
    def setUp(self):
        self.jwt_secret = "local-guest-test-secret"
        self.jwt_patcher = mock.patch.object(
            copilot_server,
            "get_jwt_secret",
            return_value=self.jwt_secret,
        )
        self.jwt_patcher.start()
        self.expires_at = int(time.time()) + 300

    def tearDown(self):
        self.jwt_patcher.stop()

    def editor_token(self, expires_at=None):
        return copilot_server.sign_jwt(
            {
                "document": {
                    "key": "document-key-v1",
                    "title": "demo.docx",
                    "fileType": "docx",
                    "permissions": {
                        "edit": True,
                        "review": True,
                        "download": False,
                    },
                },
                "documentType": "word",
                "editorConfig": {
                    "callbackUrl": "https://app.test/callback",
                    "customization": {"compactToolbar": True},
                    "user": {
                        "id": "uid-1",
                        "name": "John Smith",
                        "roles": ["reviewer"],
                    },
                },
                "iat": int(time.time()),
                "exp": self.expires_at if expires_at is None else expires_at,
            },
            self.jwt_secret,
        )

    def test_valid_editor_token_is_reissued_for_stable_guest(self):
        anonymous_id = "123e4567-e89b-42d3-a456-426614174000"
        response = copilot_server.issue_anonymous_editor_config(
            {
                "editorToken": self.editor_token(),
                "anonymousId": anonymous_id,
                "name": "Attacker-selected name",
                "permissions": {"download": True},
            }
        )
        payload = copilot_server.verify_editor_jwt_payload(response["token"])

        self.assertTrue(response["ok"])
        self.assertEqual(response["expiresAt"], self.expires_at)
        self.assertEqual(
            payload["editorConfig"]["user"],
            {
                "group": "",
                "id": f"local-guest:{anonymous_id}",
                "image": "",
                "name": "访客",
                "roles": [],
            },
        )
        self.assertEqual(payload["editorConfig"]["customization"]["anonymous"], {
            "request": False,
            "label": "访客",
        })
        self.assertEqual(payload["document"]["key"], "document-key-v1")
        self.assertEqual(
            payload["document"]["permissions"],
            {"edit": True, "review": True, "download": False},
        )
        self.assertEqual(payload["editorConfig"]["callbackUrl"], "https://app.test/callback")
        self.assertEqual(payload["exp"], self.expires_at)

    def test_invalid_guest_uuid_is_rejected(self):
        with self.assertRaises(copilot_server.BridgeError) as raised:
            copilot_server.issue_anonymous_editor_config(
                {
                    "editorToken": self.editor_token(),
                    "anonymousId": "not-a-uuid",
                }
            )

        self.assertEqual(raised.exception.status, 400)
        self.assertEqual(raised.exception.code, "INVALID_ANONYMOUS_ID")

    def test_expired_editor_token_is_not_extended(self):
        with self.assertRaises(copilot_server.BridgeError) as raised:
            copilot_server.issue_anonymous_editor_config(
                {
                    "editorToken": self.editor_token(int(time.time()) - 1),
                    "anonymousId": "123e4567-e89b-42d3-a456-426614174000",
                }
            )

        self.assertEqual(raised.exception.code, "EDITOR_TOKEN_EXPIRED")

    def test_tampered_editor_token_is_rejected(self):
        token = self.editor_token()
        tampered = token[:-1] + ("A" if token[-1] != "A" else "B")

        with self.assertRaises(copilot_server.BridgeError) as raised:
            copilot_server.issue_anonymous_editor_config(
                {
                    "editorToken": tampered,
                    "anonymousId": "123e4567-e89b-42d3-a456-426614174000",
                }
            )

        self.assertEqual(raised.exception.code, "INVALID_EDITOR_TOKEN")


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
        file_name="demo.docx",
        file_type="docx",
        editor_type="word",
        user_id="uid-1",
    ):
        return copilot_server.sign_jwt(
            {
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
            "context": {
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
                    file_name=file_name,
                    file_type=file_type,
                    editor_type=editor_type,
                    user_id=user_id,
                ),
                "state": self.editor_state(
                    document_key,
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
                        "name": "word_inspect",
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
            "SESSION_NOT_FOUND": 404,
            "IMAGE_ASSET_EXPIRED": 410,
            "REQUEST_ID_CONFLICT": 409,
            "ARGUMENTS_TOO_LARGE": 413,
            "UNSUPPORTED_IMAGE_FORMAT": 415,
            "INVALID_ARGUMENTS": 422,
            "INVALID_TOOL_ARGUMENTS": 422,
            "EXECUTION_FAILED": 500,
            "WORD_API_UNSUPPORTED": 501,
            "SHEETS_API_UNSUPPORTED": 501,
            "PERSISTENCE_FAILED": 502,
            "NOT_READY": 503,
            "BRIDGE_TIMEOUT": 504,
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
            "tool": "word_format_table_advanced",
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

    def test_missing_editor_token_message_is_editor_neutral(self):
        handler = mock.Mock()
        handler.headers = {}

        with self.assertRaises(copilot_server.BridgeError) as raised:
            copilot_server.bridge_authorization(handler)

        self.assertEqual(raised.exception.code, "EDITOR_TOKEN_REQUIRED")
        self.assertNotIn("bind_current_word", raised.exception.message)

    def test_http_relay_only_allows_large_arguments_for_image_tools(self):
        large_data_url = "data:image/png;base64," + ("A" * 300000)
        session = {
            "state": {
                "capabilities": {
                    "tools": ["word_add_image", "word_replace_text"],
                },
            },
        }
        image_command = copilot_server.bridge_build_command(
            {
                "method": "executeTool",
                "requestId": "large-image",
                "name": "word_add_image",
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
                    "name": "word_replace_text",
                    "arguments": {"search": "a", "replace": "A" * 300000},
                },
                session,
            )

        self.assertEqual(raised.exception.code, "ARGUMENTS_TOO_LARGE")

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
            "authKind": "binding",
        }

        result = copilot_server.bridge_validate(
            {
                "toolCalls": [
                    {
                        "name": "word_add_image",
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
                            "name": "word_add_image",
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
                "toolCalls": compatibility_calls,
            },
            session,
        )
        canonical = copilot_server.bridge_build_command(
            {
                "method": "executeBatch",
                "requestId": "normalize-canonical",
                "toolCalls": canonical_calls,
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
                "toolCalls": tool_calls,
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
                            "name": call["name"],
                            "arguments": call["arguments"],
                        },
                        session,
                    )
                self.assertEqual(raised.exception.code, "INVALID_TOOL_ARGUMENTS")

    def test_word_canonical_property_wins_without_relaxing_unsafe_inputs(self):
        normalized, normalizations = copilot_server.normalize_word_tool_calls(
            [{
                "name": "word_set_document_properties",
                "arguments": {"author": 42, "creator": "标准作者"},
            }]
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
                    copilot_server.require_valid_word_tool_calls([call])
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
            "authKind": "binding",
        }
        commands_before = dict(copilot_server.BRIDGE_COMMANDS)
        requests_before = dict(copilot_server.BRIDGE_REQUESTS)

        result = copilot_server.bridge_validate(
            {
                "toolCalls": [
                    {
                        "name": "word_set_document_properties",
                        "arguments": {"author": "预检作者"},
                    },
                    {
                        "name": "word_manage_style",
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
            "authKind": "binding",
        }
        commands_before = dict(copilot_server.BRIDGE_COMMANDS)
        requests_before = dict(copilot_server.BRIDGE_REQUESTS)

        result = copilot_server.bridge_validate(
            {
                "toolCalls": [
                    {
                        "name": "slides_add_slide",
                        "arguments": {"masterIndex": 1, "layoutIndex": 2},
                    },
                    {
                        "name": "slides_add_textbox",
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
                            "name": "slides_add_textbox",
                            "arguments": {
                                "slide": 1,
                                "text": "普通文本",
                                "paragraphs": [{"text": "重复内容入口"}],
                            },
                        },
                        {
                            "name": "slides_validate_layout",
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
            "authKind": "binding",
        }
        commands_before = dict(copilot_server.BRIDGE_COMMANDS)
        requests_before = dict(copilot_server.BRIDGE_REQUESTS)

        valid = copilot_server.bridge_validate(
            {
                "toolCalls": [
                    {
                        "name": "sheets_set_values",
                        "arguments": {
                            "sheet": "Sheet1",
                            "range": "A1:B2",
                            "values": [["A", "B"], [1, 2]],
                        },
                    },
                    {
                        "name": "sheets_set_formula",
                        "arguments": {
                            "sheet": "Sheet1",
                            "range": "selection",
                            "formula": "=SUM(A1:B1)",
                        },
                    },
                    {
                        "name": "sheets_manage_freeze_panes",
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
                    copilot_server.bridge_validate({"toolCalls": [call]}, claims)
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
                    "toolCalls": tool_calls,
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
        self.assertIn("arguments.paragraphIndex", encoded)
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

    def test_sessions_exchanges_editor_jwt_for_scoped_binding_token(self):
        self.register()
        editor_claims = copilot_server.verify_editor_jwt(self.editor_token())

        response = copilot_server.bridge_sessions(editor_claims)
        binding_claims = copilot_server.verify_bridge_binding_token(response["bindingToken"])

        self.assertEqual(binding_claims["authKind"], "binding")
        self.assertEqual(binding_claims["fileName"], "demo.docx")
        self.assertGreater(response["bindingExpiresAt"], int(time.time()))
        self.assertEqual(len(response["sessions"]), 1)
        self.assertTrue(response["sessions"][0]["authoritative"])

    def test_attach_issues_binding_for_ready_docx_xlsx_and_pptx_relays(self):
        cases = [
            (
                "11111111-1111-4111-8111-111111111111",
                "docx",
                "word",
            ),
            (
                "22222222-2222-4222-8222-222222222222",
                "xlsx",
                "cell",
            ),
            (
                "33333333-3333-4333-8333-333333333333",
                "pptx",
                "slide",
            ),
        ]
        for index, (document_id, file_type, editor_type) in enumerate(cases):
            with self.subTest(file_type=file_type):
                file_name = self.create_document_file(document_id, file_type)
                session_id = f"http-session:attach-{file_type}"
                user_id = f"uid-{index + 1}"
                self.register(
                    session_id,
                    f"document-key-{file_type}",
                    file_name=file_name,
                    file_type=file_type,
                    editor_type=editor_type,
                    user_id=user_id,
                )

                response = copilot_server.bridge_attach(
                    {
                        "fileName": file_name,
                        "editorType": editor_type,
                    }
                )
                claims = copilot_server.verify_bridge_binding_token(
                    response["bindingToken"]
                )

                self.assertTrue(response["ok"])
                self.assertGreater(response["bindingExpiresAt"], int(time.time()))
                self.assertEqual(response["session"]["sessionId"], session_id)
                self.assertTrue(response["session"]["ready"])
                self.assertEqual(response["session"]["editorType"], editor_type)
                self.assertEqual(claims["fileName"], file_name)
                self.assertEqual(claims["fileType"], file_type)
                self.assertEqual(claims["editorType"], editor_type)
                self.assertEqual(claims["userId"], user_id)

    def test_attach_selects_latest_ready_relay_for_same_file(self):
        file_name = self.create_document_file(
            "44444444-4444-4444-8444-444444444444",
            "docx",
        )
        self.register(
            "http-session:attach-first",
            "document-key-attach",
            file_name=file_name,
            user_id="uid-first",
        )
        self.register(
            "http-session:attach-second",
            "document-key-attach",
            file_name=file_name,
            user_id="uid-second",
        )

        response = copilot_server.bridge_attach(
            {"fileName": file_name, "editorType": "word"}
        )
        claims = copilot_server.verify_bridge_binding_token(response["bindingToken"])

        self.assertEqual(
            response["session"]["sessionId"],
            "http-session:attach-second",
        )
        self.assertEqual(claims["userId"], "uid-second")

    def test_attach_rejects_missing_editor_type_mismatch_and_missing_document(self):
        file_name = self.create_document_file(
            "55555555-5555-4555-8555-555555555555",
            "docx",
        )
        cases = [
            (
                {"fileName": file_name},
                422,
                "INVALID_ARGUMENTS",
            ),
            (
                {"fileName": file_name, "editorType": "cell"},
                409,
                "EDITOR_MISMATCH",
            ),
            (
                {
                    "fileName": "66666666-6666-4666-8666-666666666666.docx",
                    "editorType": "word",
                },
                404,
                "DOCUMENT_NOT_FOUND",
            ),
        ]
        for payload, status, code in cases:
            with self.subTest(code=code):
                with (
                    mock.patch.object(
                        copilot_server,
                        "bridge_binding_token",
                    ) as token_issuer,
                    self.assertRaises(copilot_server.BridgeError) as raised,
                ):
                    copilot_server.bridge_attach(payload)
                self.assertEqual(raised.exception.status, status)
                self.assertEqual(raised.exception.code, code)
                token_issuer.assert_not_called()

    def test_attach_requires_a_ready_active_relay(self):
        file_name = self.create_document_file(
            "77777777-7777-4777-8777-777777777777",
            "docx",
        )
        self.register(
            "http-session:attach-not-ready",
            "document-key-not-ready",
            file_name=file_name,
            ready=False,
        )

        with (
            mock.patch.object(copilot_server, "BRIDGE_DISCOVERY_WAIT_SECONDS", 0),
            self.assertRaises(copilot_server.BridgeError) as raised,
        ):
            copilot_server.bridge_attach(
                {"fileName": file_name, "editorType": "word"}
            )

        self.assertEqual(raised.exception.status, 503)
        self.assertEqual(raised.exception.code, "NO_ACTIVE_EDITOR")

    def test_attach_binding_routes_to_refreshed_relay(self):
        file_name = self.create_document_file(
            "88888888-8888-4888-8888-888888888888",
            "docx",
        )
        self.register(
            "http-session:attach-before-refresh",
            "document-key-before-refresh",
            file_name=file_name,
        )
        response = copilot_server.bridge_attach(
            {"fileName": file_name, "editorType": "word"}
        )
        binding_claims = copilot_server.verify_bridge_binding_token(
            response["bindingToken"]
        )

        self.register(
            "http-session:attach-after-refresh",
            "document-key-after-refresh",
            file_name=file_name,
        )
        refreshed = copilot_server.bridge_sessions(binding_claims)

        self.assertEqual(
            refreshed["sessions"][0]["sessionId"],
            "http-session:attach-after-refresh",
        )

    def test_attach_http_endpoint_is_post_only_and_no_store(self):
        file_name = self.create_document_file(
            "99999999-9999-4999-8999-999999999999",
            "docx",
        )
        self.register(
            "http-session:attach-http",
            "document-key-http",
            file_name=file_name,
        )
        server = copilot_server.ThreadingHTTPServer(
            ("127.0.0.1", 0),
            copilot_server.Handler,
        )
        worker = threading.Thread(target=server.serve_forever, daemon=True)
        worker.start()
        self.addCleanup(server.server_close)
        self.addCleanup(server.shutdown)

        connection = http.client.HTTPConnection(
            "127.0.0.1",
            server.server_address[1],
            timeout=2,
        )
        self.addCleanup(connection.close)
        body = json.dumps(
            {"fileName": file_name, "editorType": "word"}
        ).encode("utf-8")
        connection.request(
            "POST",
            "/bridge/attach",
            body=body,
            headers={"Content-Type": "application/json"},
        )
        response = connection.getresponse()
        payload = json.loads(response.read().decode("utf-8"))

        self.assertEqual(response.status, 200)
        self.assertEqual(response.getheader("Cache-Control"), "no-store")
        self.assertTrue(payload["bindingToken"])
        self.assertEqual(payload["session"]["sessionId"], "http-session:attach-http")

        connection.close()
        connection = http.client.HTTPConnection(
            "127.0.0.1",
            server.server_address[1],
            timeout=2,
        )
        connection.request("GET", "/bridge/attach")
        get_response = connection.getresponse()
        get_response.read()
        self.assertEqual(get_response.status, 404)

        connection.request(
            "POST",
            "/bridge/attach",
            body=json.dumps({"fileName": file_name, "editorType": "word"}),
            headers={"Content-Type": "text/plain"},
        )
        content_type_response = connection.getresponse()
        content_type_payload = json.loads(
            content_type_response.read().decode("utf-8")
        )
        self.assertEqual(content_type_response.status, 415)
        self.assertEqual(
            content_type_payload["error"]["code"],
            "UNSUPPORTED_MEDIA_TYPE",
        )

    def test_editor_jwt_binds_to_the_ready_local_guest_session(self):
        anonymous_id = "123e4567-e89b-42d3-a456-426614174000"
        guest_user_id = f"local-guest:{anonymous_id}"
        initial_token = self.editor_token()
        guest_token = copilot_server.issue_anonymous_editor_config(
            {
                "editorToken": initial_token,
                "anonymousId": anonymous_id,
            }
        )["token"]
        guest_state = self.editor_state()
        guest_state["context"]["userId"] = guest_user_id
        copilot_server.bridge_register(
            {
                "sessionId": "http-session:guest",
                "editorToken": guest_token,
                "state": guest_state,
            }
        )

        initial_claims = copilot_server.verify_editor_jwt(initial_token)
        response = copilot_server.bridge_sessions(initial_claims)
        binding_claims = copilot_server.verify_bridge_binding_token(
            response["bindingToken"]
        )

        self.assertEqual(binding_claims["userId"], guest_user_id)
        with copilot_server.BRIDGE_CONDITION:
            session_id, _ = copilot_server.bridge_select_session_locked(
                "",
                binding_claims,
            )
        self.assertEqual(session_id, "http-session:guest")

    def test_expired_binding_token_is_rejected(self):
        with mock.patch.object(copilot_server, "BRIDGE_BINDING_TOKEN_TTL_SECONDS", -1):
            token, _ = copilot_server.bridge_binding_token(
                copilot_server.verify_editor_jwt(self.editor_token())
            )

        with self.assertRaises(copilot_server.BridgeError) as raised:
            copilot_server.verify_bridge_binding_token(token)

        self.assertEqual(raised.exception.code, "BRIDGE_BINDING_TOKEN_EXPIRED")

    def test_new_registration_supersedes_old_document_version(self):
        first = self.register("http-session:first", "document-key-v1")
        editor_claims = copilot_server.verify_editor_jwt(self.editor_token("document-key-v1"))
        binding = copilot_server.bridge_sessions(editor_claims)["bindingToken"]
        binding_claims = copilot_server.verify_bridge_binding_token(binding)

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

        response = copilot_server.bridge_sessions(binding_claims)
        self.assertEqual([item["sessionId"] for item in response["sessions"]], ["http-session:second"])
        self.assertEqual(response["sessions"][0]["documentKey"], "document-key-v2")
        self.assertGreater(
            response["sessions"][0]["generation"],
            first["session"]["generation"],
        )
        self.assertTrue(second["session"]["authoritative"])

    def test_stale_binding_session_hint_routes_execute_to_latest_page(self):
        self.register("http-session:first", "document-key-v1")
        editor_claims = copilot_server.verify_editor_jwt(self.editor_token("document-key-v1"))
        binding = copilot_server.bridge_sessions(editor_claims)["bindingToken"]
        binding_claims = copilot_server.verify_bridge_binding_token(binding)
        second = self.register("http-session:second", "document-key-v2")
        holder = {}
        request = {
            "sessionId": "http-session:first",
            "method": "executeTool",
            "name": "word_inspect",
            "arguments": {"maxChars": 500},
            "requestId": "stale-session-hint-1",
            "timeoutMs": 2000,
        }

        worker = threading.Thread(
            target=lambda: holder.update(
                result=copilot_server.bridge_execute(request, binding_claims)
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
        editor_claims = copilot_server.verify_editor_jwt(self.editor_token("document-key-v1"))
        binding = copilot_server.bridge_sessions(editor_claims)["bindingToken"]
        binding_claims = copilot_server.verify_bridge_binding_token(binding)
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
            copilot_server.bridge_select_session_locked("http-session:first", binding_claims)
        self.assertEqual(raised.exception.code, "NO_ACTIVE_EDITOR")

        with self.assertRaises(copilot_server.BridgeError) as old_register:
            self.register("http-session:first", "document-key-v1")
        self.assertEqual(old_register.exception.code, "SESSION_SUPERSEDED")

        refreshed = self.register("http-session:refreshed", "document-key-v3")
        with copilot_server.BRIDGE_CONDITION:
            selected_id, _ = copilot_server.bridge_select_session_locked(
                "http-session:first",
                binding_claims,
            )
        self.assertEqual(selected_id, "http-session:refreshed")
        self.assertTrue(refreshed["session"]["authoritative"])
        self.assertIn("http-session:first", copilot_server.BRIDGE_SUPERSEDED)
        self.assertNotIn("http-session:first", copilot_server.BRIDGE_SESSIONS)
        self.assertEqual(first["session"]["sessionId"], "http-session:first")

    def test_binding_token_cannot_select_another_stable_document(self):
        self.register()
        foreign_claims = {
            "fileName": "another.docx",
            "fileType": "docx",
            "editorType": "word",
            "userId": "uid-1",
            "authKind": "binding",
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
            "name": "word_inspect",
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
            "name": "word_inspect",
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
            "name": "word_inspect",
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
        cached = copilot_server.bridge_execute(request, claims)
        self.assertTrue(cached["cached"])
        self.assertEqual(cached["result"]["text"], "current document")

        conflicting = dict(request)
        conflicting["arguments"] = {"maxChars": 2000}
        with self.assertRaises(copilot_server.BridgeError) as raised:
            copilot_server.bridge_execute(conflicting, claims)
        self.assertEqual(raised.exception.code, "REQUEST_ID_CONFLICT")
        self.assertEqual(raised.exception.status, 409)

        binding = copilot_server.bridge_sessions(claims)["bindingToken"]
        binding_claims = copilot_server.verify_bridge_binding_token(binding)
        self.register("http-session:next", "document-key-v2")
        cached_after_handoff = copilot_server.bridge_execute(request, binding_claims)
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


class VersionHistoryTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.original_root = copilot_server.EXAMPLE_FILES
        copilot_server.EXAMPLE_FILES = self.temporary.name
        user_directory = os.path.join(self.temporary.name, "test-user")
        os.makedirs(user_directory)
        self.main_file = os.path.join(user_directory, "demo.docx")
        with open(self.main_file, "wb") as stream:
            stream.write(b"version-one")

    def tearDown(self):
        copilot_server.EXAMPLE_FILES = self.original_root
        self.temporary.cleanup()

    def test_checkpoint_undo_and_redo_restore_file_bytes(self):
        checkpoint = copilot_server.create_checkpoint("demo.docx")
        self.assertTrue(checkpoint["canUndo"])

        with open(self.main_file, "wb") as stream:
            stream.write(b"version-two")

        undone = copilot_server.restore_version("demo.docx", "undo", None, None)
        with open(self.main_file, "rb") as stream:
            self.assertEqual(stream.read(), b"version-one")
        self.assertTrue(undone["canRedo"])

        redone = copilot_server.restore_version("demo.docx", "redo", None, None)
        with open(self.main_file, "rb") as stream:
            self.assertEqual(stream.read(), b"version-two")
        self.assertTrue(redone["canUndo"])

    def test_forcesave_error_four_is_a_downloadable_noop(self):
        with mock.patch.object(copilot_server, "command_service", return_value={"error": 4, "key": "key"}):
            result = copilot_server.force_save("key", "demo.docx", allow_no_changes=True)

        self.assertTrue(result["noChanges"])
        self.assertTrue(result["persisted"])
        self.assertFalse(result["accepted"])

    def test_missing_editor_session_does_not_delay_restore(self):
        with (
            mock.patch.object(copilot_server, "command_service", return_value={"error": 1, "key": "old-key"}),
            mock.patch.object(copilot_server.time, "sleep") as sleep,
        ):
            result = copilot_server.disconnect_editor("old-key", "uid-1")

        self.assertEqual(result["error"], 1)
        sleep.assert_not_called()


if __name__ == "__main__":
    unittest.main()
