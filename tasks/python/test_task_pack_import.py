"""TASK_PACK info parents when the same task title exists in two projects."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest import mock

from task_pack_import import commit_pack, pack_echoes_catalog, preview_labels
from task_store import TaskStore


class InfoParentDisambiguationTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.store = TaskStore(Path(self.tmp.name))
        self.rise = self.store.upsert_project({"title": "RISE", "filter": "work"})[
            "project"
        ]
        self.xtower = self.store.upsert_project({"title": "X-Tower", "filter": "work"})[
            "project"
        ]
        self.rise_task = self.store.upsert_task(
            {
                "project_id": self.rise["id"],
                "title": "Reuniões",
                "filter": "work",
                "kind": "punctual",
            }
        )["task"]
        self.xtower_task = self.store.upsert_task(
            {
                "project_id": self.xtower["id"],
                "title": "Reuniões",
                "filter": "work",
                "kind": "punctual",
            }
        )["task"]

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def _pack(self, info_rows: list[dict], task_rows: list[dict] | None = None) -> dict:
        return {
            "ok": True,
            "path": str(Path(self.tmp.name) / "missing-pack.txt"),
            "projects": [],
            "tasks": task_rows or [],
            "info": info_rows,
        }

    def _note(self, **extra: str) -> dict:
        row = {
            "attach_to": "task",
            "parent_title": "Reuniões",
            "filter": "work",
            "title": "PSO/BDO",
            "body": "PSO/BDO",
            "section_path": "General",
        }
        row.update(extra)
        return row

    def test_project_title_attaches_only_to_x_tower(self) -> None:
        result = commit_pack(
            self.store,
            self._pack([self._note(project_title="X-Tower")]),
        )
        self.assertTrue(result["ok"])
        self.assertEqual(result["errors"], [])
        infos = self.store.load("info_points")
        self.assertEqual(len(infos), 1)
        self.assertEqual(infos[0]["parent_type"], "task")
        self.assertEqual(infos[0]["parent_id"], self.xtower_task["id"])
        self.assertNotEqual(infos[0]["parent_id"], self.rise_task["id"])

    def test_missing_project_title_is_ambiguous_and_creates_no_info(self) -> None:
        with mock.patch("task_pack_import.write_fix_file", return_value=""):
            result = commit_pack(self.store, self._pack([self._note()]))
        self.assertFalse(result["ok"])
        self.assertIn("INFO ambiguous parent task: Reuniões", result["errors"])
        self.assertEqual(self.store.load("info_points"), [])

    def test_unique_title_still_matches_without_project_title(self) -> None:
        only = self.store.upsert_task(
            {
                "project_id": self.rise["id"],
                "title": "Check Job openings",
                "filter": "work",
                "kind": "punctual",
            }
        )["task"]
        result = commit_pack(
            self.store,
            self._pack(
                [
                    {
                        "attach_to": "task",
                        "parent_title": "Check Job openings",
                        "filter": "work",
                        "title": "Alumni",
                        "body": "Unicamp page",
                    }
                ]
            ),
        )
        self.assertTrue(result["ok"])
        infos = self.store.load("info_points")
        self.assertEqual(len(infos), 1)
        self.assertEqual(infos[0]["parent_id"], only["id"])

    def test_task_created_earlier_in_the_pack_is_a_valid_parent(self) -> None:
        result = commit_pack(
            self.store,
            self._pack(
                [
                    {
                        "attach_to": "task",
                        "parent_title": "Workshop with Amanda",
                        "project_title": "X-Tower",
                        "filter": "work",
                        "title": "Week of the 27th",
                        "body": "In person",
                    }
                ],
                [
                    {
                        "project_title": "X-Tower",
                        "filter": "work",
                        "title": "Workshop with Amanda",
                        "emoji": "general",
                        "kind": "punctual",
                        "section_path": "General",
                    }
                ],
            ),
        )
        self.assertTrue(result["ok"])
        infos = self.store.load("info_points")
        tasks = [
            t
            for t in self.store.load("tasks")
            if t.get("title") == "Workshop with Amanda"
        ]
        self.assertEqual(len(tasks), 1)
        self.assertEqual(tasks[0]["project_id"], self.xtower["id"])
        self.assertEqual(infos[0]["parent_id"], tasks[0]["id"])

    def test_preview_marks_catalog_clone_and_info_project(self) -> None:
        self.assertTrue(
            pack_echoes_catalog(
                self.store,
                [{"title": "Reuniões", "filter": "work"}],
            )
        )
        self.assertFalse(
            pack_echoes_catalog(
                self.store,
                [{"title": "Workshop with Amanda", "filter": "work"}],
            )
        )
        labels = preview_labels(
            {
                "counts": {"projects": 0, "tasks": 1, "info": 1},
                "catalog_clone": True,
                "projects": [],
                "tasks": [
                    {
                        "title": "Reuniões",
                        "filter": "work",
                        "kind": "punctual",
                        "emoji": "general",
                        "project_title": "X-Tower",
                    }
                ],
                "info": [self._note(project_title="X-Tower")],
            }
        )
        text = "\n".join(labels)
        self.assertIn("catalog clone", text)
        self.assertIn("[INFO] → Reuniões (task) · PSO/BDO @ X-Tower", text)


if __name__ == "__main__":
    unittest.main()
