"""Convert to Task clipboard text pins the selected project. Same prompt file, no new pack."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from task_server import (
    build_convert_to_task_clipboard,
    convert_to_task_prompt_path,
    format_target_project,
    resolve_target_section,
)
from task_store import TaskStore, resolve_status_emoji

SCRIPTS_ROOT = Path(__file__).resolve().parents[2]


class ConvertToTaskPromptTest(unittest.TestCase):
    def test_uses_existing_prompt_file(self) -> None:
        path = convert_to_task_prompt_path(SCRIPTS_ROOT)
        self.assertTrue(path.is_file())
        self.assertEqual(path.name, "convert-to-task.txt")

    def test_clipboard_pins_project_and_keeps_pack_prompt(self) -> None:
        text = build_convert_to_task_clipboard(
            SCRIPTS_ROOT,
            title="X-Tower",
            filt="work",
            section="Outreach",
        )
        self.assertTrue(text.startswith("DATA OUTPUT CONTRACT"))
        self.assertIn("DELIVERY MANDATE", text)
        self.assertIn("TASK_PACK.txt", text)
        self.assertIn("===FILE: TASK_TASKS.csv===", text)
        self.assertEqual(text.count("TARGET PROJECT"), 2)
        self.assertIn("- project_title: X-Tower", text)
        self.assertIn("- filter: work", text)
        self.assertIn("- section_path: Outreach", text)
        self.assertIn("the emoji cell is general", text)
        self.assertIn("One activity full of info points", text)
        self.assertIn("Notes on a task they already name", text)
        self.assertNotIn("TASK_PROJECTS row for this project.\n- project_title: ", text)

    def test_general_section_is_omitted(self) -> None:
        block = format_target_project("RISE", "personal", "General")
        self.assertIn("- project_title: RISE", block)
        self.assertIn("- filter: personal", block)
        self.assertNotIn("- section_path:", block)

    def test_unknown_section_is_dropped(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            store = TaskStore(Path(tmp))
            proj = store.upsert_project({"title": "Dog", "filter": "habits"})["project"]
            store.upsert_section({"project_id": proj["id"], "title": "Feeding"})
            self.assertEqual(
                resolve_target_section(store, proj["id"], "Feeding"), "Feeding"
            )
            self.assertEqual(resolve_target_section(store, proj["id"], "General"), "")
            self.assertEqual(resolve_target_section(store, proj["id"], "Other"), "")

    def test_none_emoji_imports_as_general(self) -> None:
        self.assertEqual(resolve_status_emoji("none"), "🔲")
        self.assertEqual(resolve_status_emoji(""), "🔲")

    def test_title_newlines_collapse(self) -> None:
        block = format_target_project("Line\nTwo", "nope", "")
        self.assertIn("- project_title: Line Two", block)
        self.assertIn("- filter: work", block)


if __name__ == "__main__":
    unittest.main()
