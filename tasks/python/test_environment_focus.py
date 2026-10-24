"""environment.txt decides whether Tasks opens on work or personal."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from task_server import read_environment_focus


class ReadEnvironmentFocusTest(unittest.TestCase):
    def test_utf8_bom_personal_is_personal(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "environment.txt"
            path.write_bytes(b"\xef\xbb\xbfpersonal")
            self.assertEqual(read_environment_focus(path), "personal")

    def test_plain_work(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "environment.txt"
            path.write_bytes(b"work\n")
            self.assertEqual(read_environment_focus(path), "work")

    def test_missing_file_defaults_to_work(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            self.assertEqual(
                read_environment_focus(Path(tmp) / "environment.txt"), "work"
            )

    def test_unknown_value_defaults_to_work(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "environment.txt"
            path.write_text("office", encoding="utf-8")
            self.assertEqual(read_environment_focus(path), "work")
