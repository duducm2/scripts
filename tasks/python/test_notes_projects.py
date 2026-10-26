"""Each environment keeps one Notes project."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from task_store import TaskStore


class NotesProjectsTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.store = TaskStore(Path(self.tmp.name))

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def _notes(self) -> dict[str, dict]:
        out = {}
        for row in self.store.load("projects"):
            if (row.get("title") or "").strip().lower() != "notes":
                continue
            if (row.get("active") or "1") == "0":
                continue
            out[row.get("filter") or ""] = row
        return out

    def test_new_store_has_notes_in_each_environment(self) -> None:
        notes = self._notes()
        self.assertEqual(set(notes), {"work", "personal", "habits"})
        for row in notes.values():
            sections = [
                s
                for s in self.store.load("sections")
                if s.get("project_id") == row["id"]
            ]
            self.assertTrue(any(s.get("title") == "General" for s in sections))

    def test_ensure_does_not_duplicate(self) -> None:
        before = {row["id"] for row in self._notes().values()}
        self.store.ensure_notes_projects()
        after = {row["id"] for row in self._notes().values()}
        self.assertEqual(before, after)


if __name__ == "__main__":
    unittest.main()
