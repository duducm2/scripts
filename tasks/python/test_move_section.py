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


if __name__ == "__main__":
    unittest.main()
