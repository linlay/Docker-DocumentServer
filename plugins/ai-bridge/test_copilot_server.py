import os
import tempfile
import unittest
from unittest import mock

import copilot_server


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
