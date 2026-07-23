import json
import os
import tempfile
import threading
import time
import unittest
from unittest import mock

import copilot_server


class ContractAlignmentTests(unittest.TestCase):
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
        self.assertEqual(len(model_names), 26)
        for entry in copilot_server.WORD_TOOLS:
            function = entry["function"]
            self.assertTrue(function["description"])
            self.assertNotIn("$ref", json.dumps(function["parameters"]))


class HttpRelayTests(unittest.TestCase):
    def setUp(self):
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

    def editor_token(self, document_key="document-key-v1", expires_in=300):
        return copilot_server.sign_jwt(
            {
                "document": {
                    "key": document_key,
                    "title": "demo.docx",
                    "fileType": "docx",
                },
                "documentType": "word",
                "editorConfig": {"user": {"id": "uid-1"}},
                "exp": int(time.time()) + expires_in,
            },
            self.jwt_secret,
        )

    def editor_state(self, document_key="document-key-v1"):
        return {
            "ready": True,
            "editorType": "word",
            "context": {
                "documentKey": document_key,
                "fileName": "demo.docx",
                "fileType": "docx",
                "editorType": "word",
                "userId": "uid-1",
            },
            "capabilities": {
                "tools": ["word_inspect", "word_replace_text"],
                "controls": ["save", "history", "undo", "redo"],
            },
        }

    def register(self, session_id="http-session:test", document_key="document-key-v1"):
        return copilot_server.bridge_register(
            {
                "sessionId": session_id,
                "editorToken": self.editor_token(document_key),
                "state": self.editor_state(document_key),
            }
        )

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

        binding = copilot_server.bridge_sessions(claims)["bindingToken"]
        binding_claims = copilot_server.verify_bridge_binding_token(binding)
        self.register("http-session:next", "document-key-v2")
        cached_after_handoff = copilot_server.bridge_execute(request, binding_claims)
        self.assertTrue(cached_after_handoff["cached"])
        self.assertEqual(cached_after_handoff["result"]["text"], "current document")


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
