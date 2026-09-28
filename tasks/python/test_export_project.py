"""Project JSON export for an AI companion."""

from __future__ import annotations

import tempfile
import unittest
from datetime import date
from pathlib import Path

from task_store import STATUS_EMOJIS, TaskStore


class ExportProjectTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.store = TaskStore(Path(self.tmp.name))

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def test_missing_project(self) -> None:
        result = self.store.export_project("PROJ_MISSING")
        self.assertFalse(result["ok"])
        self.assertEqual(result["error"], "project not found")

    def test_snapshot_includes_sections_tasks_info_and_images(self) -> None:
        proj = self.store.upsert_project(
            {"title": "Doctoral degree", "filter": "work"}
        )["project"]
        other = self.store.upsert_project({"title": "Other project", "filter": "work"})[
            "project"
        ]
        sec = self.store.upsert_section(
            {"project_id": proj["id"], "title": "Outreach"}
        )["section"]
        open_task = self.store.upsert_task(
            {
                "project_id": proj["id"],
                "section_id": sec["id"],
                "title": "Email universities",
                "filter": "work",
                "emoji": "waiting",
                "kind": "punctual",
            }
        )["task"]
        done = self.store.upsert_task(
            {
                "project_id": proj["id"],
                "section_id": sec["id"],
                "title": "Draft outline",
                "filter": "work",
                "emoji": "done",
                "kind": "punctual",
            }
        )["task"]
        self.store.upsert_task(
            {
                "project_id": other["id"],
                "title": "Should stay out",
                "filter": "work",
                "kind": "punctual",
            }
        )
        note = self.store.upsert_info(
            {
                "parent_type": "task",
                "parent_id": open_task["id"],
                "title": "Contacts",
                "body": "List of labs",
                "section_path": "People",
            }
        )["info"]
        self.store.add_attachment(
            "info",
            note["id"],
            "image",
            r"attachments\labs.png",
            "labs.png",
        )
        self.store.upsert_info(
            {
                "parent_type": "project",
                "parent_id": proj["id"],
                "title": "Birth date & age",
                "birth_date": "2025-01-15",
            }
        )
        orphan = self.store.upsert_task(
            {
                "project_id": proj["id"],
                "title": "Loose task",
                "filter": "work",
                "kind": "punctual",
            }
        )["task"]
        rows = self.store.load("tasks")
        for t in rows:
            if t["id"] == orphan["id"]:
                t["section_id"] = ""
        self.store.save("tasks", rows)

        result = self.store.export_project(proj["id"], as_of=date(2026, 9, 28))
        self.assertTrue(result["ok"])
        self.assertEqual(result["filename"], f"Doctoral-degree__{proj['id']}.json")
        doc = result["document"]
        self.assertEqual(doc["schema"], "tasks.project")
        self.assertEqual(doc["project"]["title"], "Doctoral degree")
        self.assertEqual(doc["project"]["filter"], "work")
        titles = [s["title"] for s in doc["project"]["sections"]]
        self.assertIn("Outreach", titles)
        self.assertIn("General", titles)
        self.assertIn("Other", titles)
        outreach = next(
            s for s in doc["project"]["sections"] if s["title"] == "Outreach"
        )
        by_title = {t["title"]: t for t in outreach["tasks"]}
        self.assertEqual(by_title["Email universities"]["status"], "waiting")
        self.assertEqual(
            by_title["Email universities"]["emoji"], STATUS_EMOJIS["waiting"]
        )
        self.assertTrue(by_title["Email universities"]["open"])
        self.assertEqual(
            by_title["Email universities"]["info"][0]["category"], "People"
        )
        self.assertEqual(
            by_title["Email universities"]["info"][0]["body"], "List of labs"
        )
        self.assertEqual(
            by_title["Email universities"]["info"][0]["attachments"][0]["file"],
            "attachments/labs.png",
        )
        self.assertEqual(by_title["Draft outline"]["status"], "done")
        self.assertFalse(by_title["Draft outline"]["open"])
        other_sec = next(s for s in doc["project"]["sections"] if s["title"] == "Other")
        self.assertEqual(other_sec["tasks"][0]["title"], "Loose task")
        birth = next(i for i in doc["project"]["info"] if i.get("birth_date"))
        self.assertEqual(birth["age"], "20 months")
        dumped = str(doc)
        self.assertNotIn("Should stay out", dumped)
        self.assertEqual(doc["counts"]["tasks"], 3)
        self.assertEqual(doc["counts"]["open_tasks"], 2)
        self.assertEqual(doc["counts"]["completed_tasks"], 1)
        self.assertEqual(doc["counts"]["info_points"], 2)
        self.assertEqual(doc["counts"]["attachments"], 1)


if __name__ == "__main__":
    unittest.main()
