"""Section reorder returns swapped sibling (cross-project sort_order collisions)."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from task_store import TaskStore


class MoveSectionTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.store = TaskStore(Path(self.tmp.name))

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def _set_sort(self, section_id: str, sort_order: str) -> None:
        rows = self.store.load("sections")
        for r in rows:
            if r.get("id") == section_id:
                r["sort_order"] = sort_order
        self.store.save("sections", rows)

    def test_move_section_returns_swapped_despite_shared_sort_orders(self) -> None:
        """Two projects reuse 10/20; move must swap the real neighbor, not guess globally."""
        a = self.store.upsert_project({"title": "Project A", "filter": "work"})[
            "project"
        ]
        b = self.store.upsert_project({"title": "Project B", "filter": "work"})[
            "project"
        ]
        daily = self.store.upsert_section({"project_id": a["id"], "title": "Daily"})[
            "section"
        ]
        weekly = self.store.upsert_section({"project_id": a["id"], "title": "Weekly"})[
            "section"
        ]
        decoy = self.store.upsert_section({"project_id": b["id"], "title": "Decoy"})[
            "section"
        ]

        # Force the collision the SPA used to mishandle: same numeric orders
        # across projects, with decoy appearing first in the CSV list.
        self._set_sort(daily["id"], "10")
        self._set_sort(weekly["id"], "20")
        self._set_sort(decoy["id"], "20")

        result = self.store.move_section(daily["id"], 1)
        self.assertTrue(result["ok"])
        self.assertTrue(result["moved"])
        self.assertEqual(result["section"]["id"], daily["id"])
        self.assertEqual(result["section"]["sort_order"], "20")
        self.assertIsNotNone(result["swapped"])
        self.assertEqual(result["swapped"]["id"], weekly["id"])
        self.assertEqual(result["swapped"]["sort_order"], "10")
        self.assertNotEqual(result["swapped"]["id"], decoy["id"])

        by_id = {r["id"]: r for r in self.store.load("sections")}
        self.assertEqual(by_id[daily["id"]]["sort_order"], "20")
        self.assertEqual(by_id[weekly["id"]]["sort_order"], "10")
        self.assertEqual(by_id[decoy["id"]]["sort_order"], "20")

    def test_move_section_edge_has_no_swapped(self) -> None:
        proj = self.store.upsert_project({"title": "Solo", "filter": "personal"})[
            "project"
        ]
        only = self.store.upsert_section({"project_id": proj["id"], "title": "Only"})[
            "section"
        ]
        result = self.store.move_section(only["id"], -1)
        self.assertTrue(result["ok"])
        self.assertFalse(result["moved"])
        self.assertIsNone(result["swapped"])


class DeleteSectionTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.store = TaskStore(Path(self.tmp.name))

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def test_delete_section_removes_its_tasks(self) -> None:
        proj = self.store.upsert_project(
            {"title": "Doctoral degree", "filter": "work"}
        )["project"]
        outreach = self.store.upsert_section(
            {"project_id": proj["id"], "title": "Outreach"}
        )["section"]
        writing = self.store.upsert_section(
            {"project_id": proj["id"], "title": "Writing"}
        )["section"]
        gone = self.store.upsert_task(
            {
                "project_id": proj["id"],
                "section_id": outreach["id"],
                "title": "Email universities",
                "filter": "work",
                "kind": "punctual",
            }
        )["task"]
        kept = self.store.upsert_task(
            {
                "project_id": proj["id"],
                "section_id": writing["id"],
                "title": "Chapter one",
                "filter": "work",
                "kind": "punctual",
            }
        )["task"]
        note = self.store.upsert_info(
            {
                "parent_type": "task",
                "parent_id": gone["id"],
                "title": "Contacts",
                "body": "List of labs",
            }
        )["info"]
        self.store.add_attachment(
            "info", note["id"], "image", r"attachments\labs.png", "labs.png"
        )
        self.store.upsert_info(
            {
                "parent_type": "project",
                "parent_id": proj["id"],
                "title": "Project note",
                "body": "Stays",
            }
        )
        general = self.store.ensure_general_section(proj["id"])

        result = self.store.delete_section(outreach["id"])
        self.assertTrue(result["ok"])
        self.assertIsNone(self.store.find("sections", outreach["id"]))
        titles = {t["title"] for t in self.store.load("tasks")}
        self.assertNotIn("Email universities", titles)
        self.assertIn("Chapter one", titles)
        self.assertEqual(
            self.store.find("tasks", kept["id"])["section_id"], writing["id"]
        )
        self.assertIsNone(self.store.find("tasks", gone["id"]))
        infos = self.store.load("info_points")
        self.assertFalse(any(i["id"] == note["id"] for i in infos))
        self.assertTrue(any(i["title"] == "Project note" for i in infos))
        self.assertFalse(
            any(
                a.get("parent_id") == note["id"] for a in self.store.load("attachments")
            )
        )
        general_tasks = [
            t for t in self.store.load("tasks") if t.get("section_id") == general["id"]
        ]
        self.assertEqual(general_tasks, [])

        blocked = self.store.delete_section(general["id"])
        self.assertFalse(blocked["ok"])
        self.assertEqual(blocked["error"], "cannot delete General")


if __name__ == "__main__":
    unittest.main()
