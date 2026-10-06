"""Project JSON export for an AI companion."""

from __future__ import annotations

import json
import tempfile
import unittest
from datetime import date
from pathlib import Path

from task_store import (
    STATUS_EMOJIS,
    TaskStore,
    project_export_filename,
    task_export_filename,
)


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

    def test_awkward_title_keeps_completed_task_and_multiline_note(self) -> None:
        proj = self.store.upsert_project({"title": "«—»", "filter": "personal"})[
            "project"
        ]
        done = self.store.upsert_task(
            {
                "project_id": proj["id"],
                "title": "Finished draft",
                "filter": "personal",
                "emoji": "done",
                "kind": "punctual",
            }
        )["task"]
        note = 'He said "hello"\nand left'
        self.store.upsert_info(
            {
                "parent_type": "task",
                "parent_id": done["id"],
                "title": "Quote",
                "body": note,
            }
        )
        result = self.store.export_project(proj["id"], as_of=date(2026, 9, 28))
        self.assertTrue(result["ok"])
        filename = result["filename"]
        self.assertEqual(filename, project_export_filename(proj))
        self.assertEqual(filename, f"project__{proj['id']}.json")
        self.assertRegex(filename, r"^project__[A-Za-z0-9_-]+\.json$")
        self.assertTrue(filename.isascii())
        tasks = [
            t for s in result["document"]["project"]["sections"] for t in s["tasks"]
        ]
        finished = next(t for t in tasks if t["title"] == "Finished draft")
        self.assertEqual(finished["status"], "done")
        self.assertFalse(finished["open"])
        self.assertEqual(finished["info"][0]["body"], note)
        self.assertNotRegex(filename, r"[^\x00-\x7F]")

    def test_json_files_follow_add_and_remove(self) -> None:
        alpha = self.store.upsert_project({"title": "Alpha", "filter": "work"})[
            "project"
        ]
        beta = self.store.upsert_project({"title": "Beta", "filter": "personal"})[
            "project"
        ]
        folder = self.store.project_json_dir()
        alpha_path = folder / project_export_filename(alpha)
        beta_path = folder / project_export_filename(beta)
        self.assertTrue(alpha_path.is_file())
        self.assertTrue(beta_path.is_file())
        index = json.loads((folder / "projects.json").read_text(encoding="utf-8"))
        self.assertEqual(
            [p["id"] for p in index["projects"]], [alpha["id"], beta["id"]]
        )

        self.store.upsert_task(
            {
                "project_id": beta["id"],
                "title": "Hello from beta",
                "filter": "personal",
                "kind": "punctual",
            }
        )
        self.assertIn("Hello from beta", beta_path.read_text(encoding="utf-8"))

        renamed = self.store.upsert_project(
            {"id": alpha["id"], "title": "Alpha Renamed", "filter": "work"}
        )["project"]
        self.assertFalse(alpha_path.is_file())
        renamed_path = folder / project_export_filename(renamed)
        self.assertTrue(renamed_path.is_file())

        self.store.delete_project(alpha["id"])
        self.assertFalse(renamed_path.is_file())
        self.assertTrue(beta_path.is_file())
        index = json.loads((folder / "projects.json").read_text(encoding="utf-8"))
        self.assertEqual([p["id"] for p in index["projects"]], [beta["id"]])
        self.assertNotIn(alpha["id"], beta_path.read_text(encoding="utf-8"))

    def test_waiting_status_exports_on_project_and_section(self) -> None:
        proj = self.store.upsert_project({"title": "Paused work", "filter": "work"})[
            "project"
        ]
        sec = self.store.upsert_section({"project_id": proj["id"], "title": "Later"})[
            "section"
        ]
        rejected = self.store.set_project_status(proj["id"], "nope")
        self.assertEqual(rejected["error"], "invalid status")
        marked = self.store.set_project_status(proj["id"], "waiting")
        self.assertTrue(marked["ok"])
        self.assertEqual(marked["project"]["status"], "waiting")
        sec_marked = self.store.set_section_status(sec["id"], "waiting")
        self.assertTrue(sec_marked["ok"])
        cleared = self.store.set_section_status(sec["id"], "")
        self.assertEqual(cleared["section"]["status"], "")
        self.store.set_section_status(sec["id"], "waiting")

        doc = self.store.export_project(proj["id"])["document"]["project"]
        self.assertEqual(doc["status"], "waiting")
        later = next(s for s in doc["sections"] if s["title"] == "Later")
        self.assertEqual(later["status"], "waiting")
        general = next(s for s in doc["sections"] if s["title"] == "General")
        self.assertNotIn("status", general)


class ExportTaskTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.store = TaskStore(Path(self.tmp.name))

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def test_missing_task(self) -> None:
        result = self.store.export_task("TASK_MISSING")
        self.assertFalse(result["ok"])
        self.assertEqual(result["error"], "task not found")

    def test_snapshot_is_one_task_with_its_info_points(self) -> None:
        proj = self.store.upsert_project(
            {"title": "Doctoral degree", "filter": "work"}
        )["project"]
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
        sibling = self.store.upsert_task(
            {
                "project_id": proj["id"],
                "section_id": sec["id"],
                "title": "Draft outline",
                "filter": "work",
                "emoji": "done",
                "kind": "punctual",
            }
        )["task"]
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
        self.store.add_attachment(
            "task",
            open_task["id"],
            "image",
            r"attachments\campus.png",
            "campus.png",
        )
        self.store.upsert_info(
            {
                "parent_type": "task",
                "parent_id": sibling["id"],
                "title": "Sibling note",
                "body": "Should stay out",
            }
        )
        self.store.upsert_info(
            {
                "parent_type": "project",
                "parent_id": proj["id"],
                "title": "Project note",
                "body": "Also stay out",
            }
        )

        result = self.store.export_task(open_task["id"], as_of=date(2026, 9, 28))
        self.assertTrue(result["ok"])
        self.assertEqual(result["filename"], task_export_filename(open_task))
        self.assertEqual(
            result["filename"], f"Email-universities__{open_task['id']}.json"
        )
        doc = result["document"]
        self.assertEqual(doc["schema"], "tasks.task")
        self.assertEqual(doc["project"], {"id": proj["id"], "title": "Doctoral degree"})
        self.assertEqual(doc["section"], {"id": sec["id"], "title": "Outreach"})
        self.assertEqual(doc["task"]["title"], "Email universities")
        self.assertEqual(doc["task"]["status"], "waiting")
        self.assertEqual(doc["task"]["info"][0]["title"], "Contacts")
        self.assertEqual(doc["task"]["info"][0]["body"], "List of labs")
        self.assertEqual(doc["task"]["info"][0]["category"], "People")
        self.assertEqual(
            doc["task"]["info"][0]["attachments"][0]["file"], "attachments/labs.png"
        )
        self.assertEqual(
            doc["task"]["attachments"][0]["file"], "attachments/campus.png"
        )
        self.assertEqual(doc["counts"]["info_points"], 1)
        self.assertEqual(doc["counts"]["attachments"], 2)
        dumped = str(doc)
        self.assertNotIn("Draft outline", dumped)
        self.assertNotIn("Sibling note", dumped)
        self.assertNotIn("Should stay out", dumped)
        self.assertNotIn("Project note", dumped)
        self.assertNotIn("Also stay out", dumped)
        self.assertNotIn("sections", doc)


if __name__ == "__main__":
    unittest.main()
