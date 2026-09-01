"""CSV store for the Memory Palace web app."""

from __future__ import annotations

import os
import re
import shutil
import sys
import threading
from copy import deepcopy
from pathlib import Path
from typing import Any

from data_aggregator import load_all
from plan_csv import load_plan_tables, next_id, save_plan_tables
from schemas import (
    ATOMS_HEADERS,
    BEASTS_HEADERS,
    PALACE_IMAGES_HEADERS,
    PALACES_HEADERS,
    PLAN_ITEMS_HEADERS,
    PLAN_RESOURCES_HEADERS,
    PLANS_HEADERS,
    STUDIES_HEADERS,
    ensure_data_dir,
    validate_beast_atoms,
)
import csv as _csv
from study_plan_parser import slug_filename
from study_practice_md import write_study
from study_plans_md import sync_one_plan

PALACE_DEFER_MD_SYNC = os.environ.get("PALACE_DEFER_MD_SYNC", "").strip().lower() in (
    "1",
    "true",
    "yes",
)

CACHE_KINDS = (
    "studies",
    "palaces",
    "palace_images",
    "beasts",
    "atoms",
    "plans",
    "plan_items",
    "plan_resources",
)
PLAN_KINDS = ("plans", "plan_items", "plan_resources")

ENTITY_PREFIX = {
    "studies": "STUDY_",
    "palaces": "PALACE_",
    "palace_images": "PALIMG_",
    "beasts": "BEAST_",
    "atoms": "ATOM_",
    "plans": "PLAN_",
    "plan_items": "PITEM_",
    "plan_resources": "PRES_",
}

ENTITY_HEADERS = {
    "studies": STUDIES_HEADERS,
    "palaces": PALACES_HEADERS,
    "palace_images": PALACE_IMAGES_HEADERS,
    "beasts": BEASTS_HEADERS,
    "atoms": ATOMS_HEADERS,
    "plans": PLANS_HEADERS,
    "plan_items": PLAN_ITEMS_HEADERS,
    "plan_resources": PLAN_RESOURCES_HEADERS,
}


def _next_sort(rows: list[dict[str, str]], key: str = "sort_order") -> str:
    mx = 0
    for r in rows:
        try:
            mx = max(mx, int(r.get(key) or 0))
        except ValueError:
            pass
    return str(mx + 1)


def _next_palace_number(palaces: list[dict[str, str]], study_id: str) -> str:
    mx = 0
    for p in palaces:
        if p.get("study_id") != study_id:
            continue
        try:
            mx = max(mx, int(p.get("palace_number") or 0))
        except ValueError:
            pass
    return str(mx + 1)


class PalaceStore:
    def __init__(
        self,
        data_dir: Path,
        output_dir: Path | None = None,
        studies_root: Path | None = None,
    ):
        self.data_dir = data_dir
        self.output_dir = output_dir or (data_dir.parent / "output")
        self.studies_root = studies_root
        ensure_data_dir(data_dir)
        self._cache_tree: dict[str, list[dict[str, str]]] | None = None
        self._cache_stamp: float = -1.0
        self._pending_practice: set[str] = set()
        self._pending_plans: set[str] = set()
        self._sync_lock = threading.Lock()
        self._sync_timer: threading.Timer | None = None

    def _csv_stamp(self) -> float:
        mx = 0.0
        for kind in CACHE_KINDS:
            path = self.data_dir / f"{kind}.csv"
            if path.is_file():
                mx = max(mx, path.stat().st_mtime)
        return mx

    def _invalidate_cache(self) -> None:
        self._cache_tree = None
        self._cache_stamp = -1.0

    def _load_tree(self) -> dict[str, list[dict[str, str]]]:
        stamp = self._csv_stamp()
        if self._cache_tree is not None and stamp == self._cache_stamp:
            return deepcopy(self._cache_tree)
        data = load_all(self.data_dir)
        plans = load_plan_tables(self.data_dir)
        data["plans"] = plans["plans"]
        data["plan_items"] = plans["plan_items"]
        data["plan_resources"] = plans["plan_resources"]
        self._cache_tree = deepcopy(data)
        self._cache_stamp = stamp
        return data

    def _write_kind(
        self, kind: str, headers: list[str], rows: list[dict[str, str]]
    ) -> None:
        path = self.data_dir / f"{kind}.csv"
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", encoding="utf-8", newline="") as f:
            w = _csv.DictWriter(
                f, fieldnames=headers, extrasaction="ignore", lineterminator="\n"
            )
            w.writeheader()
            for row in rows:
                w.writerow({h: row.get(h, "") for h in headers})

    def _save_tree(
        self,
        data: dict[str, list[dict[str, str]]],
        kinds: list[str] | None = None,
    ) -> None:
        write_all = kinds is None
        kinds_set = set(kinds or [])
        if write_all or "studies" in kinds_set:
            self._write_kind("studies", STUDIES_HEADERS, data.get("studies", []))
        if write_all or "palaces" in kinds_set:
            self._write_kind("palaces", PALACES_HEADERS, data.get("palaces", []))
        if write_all or "palace_images" in kinds_set:
            self._write_kind(
                "palace_images", PALACE_IMAGES_HEADERS, data.get("palace_images", [])
            )
        if write_all or "beasts" in kinds_set:
            self._write_kind("beasts", BEASTS_HEADERS, data.get("beasts", []))
        if write_all or "atoms" in kinds_set:
            self._write_kind("atoms", ATOMS_HEADERS, data.get("atoms", []))
        if write_all or kinds_set.intersection(PLAN_KINDS):
            save_plan_tables(
                self.data_dir,
                data.get("plans", []),
                data.get("plan_items", []),
                data.get("plan_resources", []),
            )
        self._cache_tree = deepcopy(data)
        self._cache_stamp = self._csv_stamp()

    def _arm_sync_timer(self) -> None:
        if self._sync_timer is not None:
            self._sync_timer.cancel()
        self._sync_timer = threading.Timer(0.5, self._flush_deferred_sync)
        self._sync_timer.daemon = True
        self._sync_timer.start()

    def _flush_deferred_sync(self) -> None:
        with self._sync_lock:
            practice_ids = list(self._pending_practice)
            plan_ids = list(self._pending_plans)
            self._pending_practice.clear()
            self._pending_plans.clear()
            self._sync_timer = None
        for study_id in practice_ids:
            self._sync_practice_now(study_id)
        for study_id in plan_ids:
            self._sync_plans_now(study_id)

    def _schedule_sync_practice(self, study_id: str) -> None:
        if not study_id:
            return
        if not PALACE_DEFER_MD_SYNC:
            self._sync_practice_now(study_id)
            return
        with self._sync_lock:
            self._pending_practice.add(study_id)
            self._arm_sync_timer()

    def _schedule_sync_plans(self, study_id: str) -> None:
        if not study_id:
            return
        if not PALACE_DEFER_MD_SYNC:
            self._sync_plans_now(study_id)
            return
        with self._sync_lock:
            self._pending_plans.add(study_id)
            self._arm_sync_timer()

    def _sync_practice_now(self, study_id: str) -> None:
        if not study_id:
            return
        try:
            write_study(study_id, self.data_dir, self.output_dir, self.studies_root)
        except Exception as e:
            print(
                f"palace_store: practice sync failed for {study_id}: {e}",
                file=sys.stderr,
            )

    def _sync_plans_now(
        self,
        study_id: str,
        studies: list[dict[str, str]] | None = None,
    ) -> None:
        """Export plan MD via study_plans_md (CSV → output/plans/{slug}.md)."""
        if not study_id:
            return
        try:
            if studies is None:
                studies = self._load_tree().get("studies", [])
            study = next(
                (s for s in studies if s.get("id") == study_id),
                None,
            )
            slug = ((study or {}).get("notes_rel_path") or "").strip()
            if not slug:
                return
            sync_one_plan(
                None,
                slug,
                self.output_dir,
                data_dir=self.data_dir,
                study_id=study_id,
            )
        except Exception as e:
            print(
                f"palace_store: plan sync failed for {study_id}: {e}",
                file=sys.stderr,
            )

    def _sync_practice(self, study_id: str) -> None:
        self._schedule_sync_practice(study_id)

    def _sync_plans(self, study_id: str) -> None:
        self._schedule_sync_plans(study_id)

    def state(self) -> dict[str, Any]:
        data = self._load_tree()
        return {
            "ok": True,
            "studies": data["studies"],
            "palaces": data["palaces"],
            "palace_images": data.get("palace_images", []),
            "beasts": data["beasts"],
            "atoms": data["atoms"],
            "plans": data["plans"],
            "plan_items": data["plan_items"],
            "plan_resources": data["plan_resources"],
            "meta": {
                "practice_github": "https://github.com/duducm2/scripts/tree/main/mnemonics/output/practice",
                "plans_github": "https://github.com/duducm2/scripts/tree/main/mnemonics/output/plans",
                "server_port": 8767,
            },
        }

    def upsert(self, entity: str, payload: dict[str, Any]) -> dict[str, Any]:
        if entity not in ENTITY_HEADERS:
            return {"ok": False, "error": f"unknown entity {entity}"}
        data = self._load_tree()
        rows = list(data.get(entity, []))
        headers = ENTITY_HEADERS[entity]
        rid = str(payload.get("id") or "").strip()
        existing = next((r for r in rows if r.get("id") == rid), None) if rid else None

        if entity == "studies":
            return self._upsert_study(data, rows, payload, existing)
        if entity == "palaces":
            return self._upsert_palace(data, rows, payload, existing)
        if entity == "beasts":
            return self._upsert_beast(data, rows, payload, existing)
        if entity == "atoms":
            return self._upsert_atom(data, rows, payload, existing)
        if entity == "plans":
            return self._upsert_plan(data, rows, payload, existing)
        if entity == "plan_items":
            return self._upsert_plan_item(data, rows, payload, existing)

        # generic
        row = {
            h: str(payload.get(h, existing.get(h, "") if existing else ""))
            for h in headers
        }
        if not rid:
            ids = {r.get("id") or "" for r in rows}
            rid = next_id(ENTITY_PREFIX[entity], ids)
            row["id"] = rid
            if "sort_order" in headers and not row.get("sort_order"):
                row["sort_order"] = _next_sort(rows)
            rows.append(row)
        else:
            rows = [row if r.get("id") == rid else r for r in rows]
            if not existing:
                rows.append(row)
        data[entity] = rows
        self._save_tree(data, [entity])
        return {"ok": True, "id": rid, "row": row}

    def _upsert_study(
        self,
        data: dict[str, list[dict[str, str]]],
        rows: list[dict[str, str]],
        payload: dict[str, Any],
        existing: dict[str, str] | None,
    ) -> dict[str, Any]:
        title = str(payload.get("title") or (existing or {}).get("title") or "").strip()
        if not title:
            return {"ok": False, "error": "title required"}
        rid = str(payload.get("id") or "").strip()
        if not rid:
            slug = re.sub(r"[^A-Za-z0-9]+", "", title.upper())[:12] or "STUDY"
            rid = f"STUDY_{slug}" if not slug.startswith("STUDY") else slug
            if not rid.startswith("STUDY_"):
                rid = f"STUDY_{rid}"
            base = rid
            n = 2
            ids = {r.get("id") for r in rows}
            while rid in ids:
                rid = f"{base}_{n}"
                n += 1
        notes = str(
            payload.get("notes_rel_path")
            or (existing or {}).get("notes_rel_path")
            or slug_filename(title)
        ).strip()
        row = {
            "id": rid,
            "title": title,
            "notes_rel_path": notes,
            "sort_order": str(
                payload.get("sort_order")
                or (existing or {}).get("sort_order")
                or _next_sort(rows)
            ),
            "active": str(payload.get("active", (existing or {}).get("active", "1"))),
        }
        if existing:
            data["studies"] = [row if r.get("id") == rid else r for r in rows]
        else:
            rows.append(row)
            data["studies"] = rows
        self._save_tree(data, ["studies"])
        return {"ok": True, "id": rid, "row": row}

    def _upsert_palace(
        self,
        data: dict[str, list[dict[str, str]]],
        rows: list[dict[str, str]],
        payload: dict[str, Any],
        existing: dict[str, str] | None,
    ) -> dict[str, Any]:
        study_id = str(
            payload.get("study_id") or (existing or {}).get("study_id") or ""
        ).strip()
        if not study_id:
            return {"ok": False, "error": "study_id required"}
        title = str(payload.get("title") or (existing or {}).get("title") or "").strip()
        if not title:
            return {"ok": False, "error": "title required"}
        rid = str(payload.get("id") or "").strip()
        if not rid:
            ids = {r.get("id") or "" for r in rows}
            rid = next_id("PALACE_", ids)
            num = _next_palace_number(rows, study_id)
        else:
            num = str(
                payload.get("palace_number")
                or (existing or {}).get("palace_number")
                or _next_palace_number(rows, study_id)
            )
        row = {
            "id": rid,
            "study_id": study_id,
            "palace_number": (
                num
                if not existing
                else str(
                    payload.get("palace_number")
                    or (existing or {}).get("palace_number")
                    or num
                )
            ),
            "title": title,
            "character_name": str(
                payload.get("character_name")
                or (existing or {}).get("character_name")
                or "(unassigned)"
            ),
            "image_rel_path": str(
                payload.get(
                    "image_rel_path",
                    (existing or {}).get("image_rel_path", ""),
                )
            ),
            "depth_slots_used": str(
                payload.get(
                    "depth_slots_used",
                    (existing or {}).get("depth_slots_used", "0"),
                )
            ),
            "image_prompt": str(
                payload.get("image_prompt", (existing or {}).get("image_prompt", ""))
            ),
            "palace_notes": str(
                payload.get("palace_notes", (existing or {}).get("palace_notes", ""))
            ),
        }
        if existing:
            data["palaces"] = [row if r.get("id") == rid else r for r in rows]
        else:
            if not payload.get("palace_number"):
                row["palace_number"] = _next_palace_number(rows, study_id)
            rows.append(row)
            data["palaces"] = rows
        self._save_tree(data, ["palaces"])
        self._sync_practice(study_id)
        return {"ok": True, "id": rid, "row": row}

    def _upsert_beast(
        self,
        data: dict[str, list[dict[str, str]]],
        rows: list[dict[str, str]],
        payload: dict[str, Any],
        existing: dict[str, str] | None,
    ) -> dict[str, Any]:
        palace_id = str(
            payload.get("palace_id") or (existing or {}).get("palace_id") or ""
        ).strip()
        if not palace_id:
            return {"ok": False, "error": "palace_id required"}
        siblings = [
            r
            for r in rows
            if r.get("palace_id") == palace_id
            and (not existing or r.get("id") != existing.get("id"))
        ]
        if not existing and len(siblings) >= 5:
            return {"ok": False, "error": "Palace may have at most 5 beasts"}
        peg = str(
            payload.get("peg_code") or (existing or {}).get("peg_code") or ""
        ).strip()
        if peg.isdigit():
            return {"ok": False, "error": "peg_code must not be numeric"}
        rid = str(payload.get("id") or "").strip()
        if not rid:
            ids = {r.get("id") or "" for r in rows}
            rid = next_id("BEAST_", ids)
        row = {
            "id": rid,
            "palace_id": palace_id,
            "peg_code": peg,
            "beast_name": str(
                payload.get("beast_name") or (existing or {}).get("beast_name") or ""
            ),
            "beast_source": str(
                payload.get("beast_source")
                or (existing or {}).get("beast_source")
                or ""
            ),
            "sensory_channel": str(
                payload.get("sensory_channel")
                or (existing or {}).get("sensory_channel")
                or ""
            ),
            "is_smashed": str(
                payload.get("is_smashed", (existing or {}).get("is_smashed", "0"))
            ),
            "sort_order": str(
                payload.get("sort_order")
                or (existing or {}).get("sort_order")
                or _next_sort(siblings)
            ),
        }
        if existing:
            data["beasts"] = [row if r.get("id") == rid else r for r in rows]
        else:
            rows.append(row)
            data["beasts"] = rows
        palace = next((p for p in data["palaces"] if p.get("id") == palace_id), None)
        self._save_tree(data, ["beasts"])
        if palace:
            self._sync_practice(palace.get("study_id") or "")
        return {"ok": True, "id": rid, "row": row}

    def _upsert_atom(
        self,
        data: dict[str, list[dict[str, str]]],
        rows: list[dict[str, str]],
        payload: dict[str, Any],
        existing: dict[str, str] | None,
    ) -> dict[str, Any]:
        beast_id = str(
            payload.get("beast_id") or (existing or {}).get("beast_id") or ""
        ).strip()
        if not beast_id:
            return {"ok": False, "error": "beast_id required"}
        rid = str(payload.get("id") or "").strip()
        if not rid:
            ids = {r.get("id") or "" for r in rows}
            rid = next_id("ATOM_", ids)
        kind = (
            str(payload.get("kind") or (existing or {}).get("kind") or "single")
            .strip()
            .lower()
        )
        row = {
            "id": rid,
            "beast_id": beast_id,
            "kind": kind,
            "zone": str(payload.get("zone", (existing or {}).get("zone", ""))),
            "zone_label": str(
                payload.get("zone_label", (existing or {}).get("zone_label", ""))
            ),
            "concept": str(payload.get("concept", (existing or {}).get("concept", ""))),
            "quote": str(payload.get("quote", (existing or {}).get("quote", ""))),
            "story": str(payload.get("story", (existing or {}).get("story", ""))),
            "sensory": str(payload.get("sensory", (existing or {}).get("sensory", ""))),
            "ipa": str(payload.get("ipa", (existing or {}).get("ipa", ""))),
            "sort_order": str(
                payload.get("sort_order")
                or (existing or {}).get("sort_order")
                or _next_sort([r for r in rows if r.get("beast_id") == beast_id])
            ),
        }
        trial = [
            row if r.get("id") == rid else r
            for r in rows
            if r.get("beast_id") == beast_id or r.get("id") == rid
        ]
        if not existing:
            trial = [r for r in rows if r.get("beast_id") == beast_id] + [row]
        else:
            trial = [
                row if r.get("id") == rid else r
                for r in rows
                if r.get("beast_id") == beast_id
            ]
        err = validate_beast_atoms(trial)
        if err:
            return {"ok": False, "error": err}
        if existing:
            data["atoms"] = [row if r.get("id") == rid else r for r in rows]
        else:
            rows.append(row)
            data["atoms"] = rows
        # smash flag
        zoned = any(
            (a.get("kind") or "").lower() in ("zoned", "subtopic")
            for a in data["atoms"]
            if a.get("beast_id") == beast_id
        )
        new_beasts = []
        for b in data["beasts"]:
            if b.get("id") == beast_id:
                bb = dict(b)
                bb["is_smashed"] = "1" if zoned else "0"
                new_beasts.append(bb)
            else:
                new_beasts.append(b)
        data["beasts"] = new_beasts
        beast = next((b for b in data["beasts"] if b.get("id") == beast_id), None)
        palace = None
        if beast:
            palace = next(
                (p for p in data["palaces"] if p.get("id") == beast.get("palace_id")),
                None,
            )
        self._save_tree(data, ["beasts", "atoms"])
        if palace:
            self._sync_practice(palace.get("study_id") or "")
        return {"ok": True, "id": rid, "row": row}

    def _upsert_plan(
        self,
        data: dict[str, list[dict[str, str]]],
        rows: list[dict[str, str]],
        payload: dict[str, Any],
        existing: dict[str, str] | None,
    ) -> dict[str, Any]:
        study_id = str(
            payload.get("study_id") or (existing or {}).get("study_id") or ""
        ).strip()
        if not study_id:
            return {"ok": False, "error": "study_id required"}
        if not existing:
            active = [
                r
                for r in rows
                if r.get("study_id") == study_id and r.get("active", "1") != "0"
            ]
            if active:
                return {"ok": False, "error": "Study already has an active plan"}
        rid = str(payload.get("id") or "").strip()
        if not rid:
            ids = {r.get("id") or "" for r in rows}
            rid = next_id("PLAN_", ids)
        row = {
            "id": rid,
            "study_id": study_id,
            "title": str(
                payload.get("title") or (existing or {}).get("title") or "Study Plan"
            ),
            "sort_order": str(
                payload.get("sort_order") or (existing or {}).get("sort_order") or "1"
            ),
            "active": str(payload.get("active", (existing or {}).get("active", "1"))),
        }
        if existing:
            data["plans"] = [row if r.get("id") == rid else r for r in rows]
        else:
            rows.append(row)
            data["plans"] = rows
        self._save_tree(data, ["plans"])
        self._sync_plans(study_id)
        return {"ok": True, "id": rid, "row": row}

    def _upsert_plan_item(
        self,
        data: dict[str, list[dict[str, str]]],
        rows: list[dict[str, str]],
        payload: dict[str, Any],
        existing: dict[str, str] | None,
    ) -> dict[str, Any]:
        plan_id = str(
            payload.get("plan_id") or (existing or {}).get("plan_id") or ""
        ).strip()
        if not plan_id:
            return {"ok": False, "error": "plan_id required"}
        rid = str(payload.get("id") or "").strip()
        if not rid:
            ids = {r.get("id") or "" for r in rows}
            rid = next_id("PITEM_", ids)
        row = {
            "id": rid,
            "plan_id": plan_id,
            "section_path": str(
                payload.get("section_path")
                or (existing or {}).get("section_path")
                or "Backlog"
            ),
            "text": str(payload.get("text") or (existing or {}).get("text") or ""),
            "checked": str(
                payload.get("checked", (existing or {}).get("checked", "0"))
            ),
            "sort_order": str(
                payload.get("sort_order")
                or (existing or {}).get("sort_order")
                or _next_sort([r for r in rows if r.get("plan_id") == plan_id])
            ),
        }
        if existing:
            data["plan_items"] = [row if r.get("id") == rid else r for r in rows]
        else:
            rows.append(row)
            data["plan_items"] = rows
        plan = next((p for p in data["plans"] if p.get("id") == plan_id), None)
        self._save_tree(data, ["plan_items"])
        if plan:
            self._sync_plans(plan.get("study_id") or "")
        return {"ok": True, "id": rid, "row": row}

    def delete(self, entity: str, entity_id: str) -> dict[str, Any]:
        entity_id = (entity_id or "").strip()
        if not entity_id or entity not in ENTITY_HEADERS:
            return {"ok": False, "error": "invalid delete"}
        data = self._load_tree()
        study_id = ""

        if entity == "studies":
            study_id = entity_id
            palace_ids = {
                p["id"]
                for p in data["palaces"]
                if p.get("study_id") == entity_id and p.get("id")
            }
            beast_ids = {
                b["id"]
                for b in data["beasts"]
                if b.get("palace_id") in palace_ids and b.get("id")
            }
            plan_ids = {
                p["id"]
                for p in data["plans"]
                if p.get("study_id") == entity_id and p.get("id")
            }
            data["studies"] = [r for r in data["studies"] if r.get("id") != entity_id]
            data["palaces"] = [
                r for r in data["palaces"] if r.get("id") not in palace_ids
            ]
            data["palace_images"] = [
                r
                for r in data.get("palace_images", [])
                if r.get("palace_id") not in palace_ids
            ]
            data["beasts"] = [r for r in data["beasts"] if r.get("id") not in beast_ids]
            data["atoms"] = [
                r for r in data["atoms"] if r.get("beast_id") not in beast_ids
            ]
            data["plans"] = [r for r in data["plans"] if r.get("id") not in plan_ids]
            data["plan_items"] = [
                r for r in data["plan_items"] if r.get("plan_id") not in plan_ids
            ]
            data["plan_resources"] = [
                r for r in data["plan_resources"] if r.get("plan_id") not in plan_ids
            ]
        elif entity == "palaces":
            palace = next(
                (p for p in data["palaces"] if p.get("id") == entity_id), None
            )
            study_id = (palace or {}).get("study_id") or ""
            beast_ids = {
                b["id"]
                for b in data["beasts"]
                if b.get("palace_id") == entity_id and b.get("id")
            }
            data["palaces"] = [r for r in data["palaces"] if r.get("id") != entity_id]
            data["palace_images"] = [
                r
                for r in data.get("palace_images", [])
                if r.get("palace_id") != entity_id
            ]
            data["beasts"] = [r for r in data["beasts"] if r.get("id") not in beast_ids]
            data["atoms"] = [
                r for r in data["atoms"] if r.get("beast_id") not in beast_ids
            ]
        elif entity == "beasts":
            beast = next((b for b in data["beasts"] if b.get("id") == entity_id), None)
            palace = None
            if beast:
                palace = next(
                    (
                        p
                        for p in data["palaces"]
                        if p.get("id") == beast.get("palace_id")
                    ),
                    None,
                )
            study_id = (palace or {}).get("study_id") or ""
            data["beasts"] = [r for r in data["beasts"] if r.get("id") != entity_id]
            data["atoms"] = [r for r in data["atoms"] if r.get("beast_id") != entity_id]
        elif entity == "atoms":
            atom = next((a for a in data["atoms"] if a.get("id") == entity_id), None)
            beast = None
            if atom:
                beast = next(
                    (b for b in data["beasts"] if b.get("id") == atom.get("beast_id")),
                    None,
                )
            palace = None
            if beast:
                palace = next(
                    (
                        p
                        for p in data["palaces"]
                        if p.get("id") == beast.get("palace_id")
                    ),
                    None,
                )
            study_id = (palace or {}).get("study_id") or ""
            data["atoms"] = [r for r in data["atoms"] if r.get("id") != entity_id]
        elif entity == "plans":
            plan = next((p for p in data["plans"] if p.get("id") == entity_id), None)
            study_id = (plan or {}).get("study_id") or ""
            data["plans"] = [r for r in data["plans"] if r.get("id") != entity_id]
            data["plan_items"] = [
                r for r in data["plan_items"] if r.get("plan_id") != entity_id
            ]
            data["plan_resources"] = [
                r for r in data["plan_resources"] if r.get("plan_id") != entity_id
            ]
        elif entity == "plan_items":
            item = next(
                (r for r in data["plan_items"] if r.get("id") == entity_id), None
            )
            plan = None
            if item:
                plan = next(
                    (p for p in data["plans"] if p.get("id") == item.get("plan_id")),
                    None,
                )
            study_id = (plan or {}).get("study_id") or ""
            data["plan_items"] = [
                r for r in data["plan_items"] if r.get("id") != entity_id
            ]
        else:
            data[entity] = [r for r in data.get(entity, []) if r.get("id") != entity_id]

        self._save_tree(data)
        if entity in ("studies", "palaces", "beasts", "atoms"):
            self._sync_practice(study_id)
        if entity in ("studies", "plans", "plan_items"):
            self._sync_plans(study_id)
        return {"ok": True, "id": entity_id}

    def quick_image(
        self, palace_id: str, desktop: Path | None = None
    ) -> dict[str, Any]:
        palace_id = (palace_id or "").strip()
        data = self._load_tree()
        palace = next((p for p in data["palaces"] if p.get("id") == palace_id), None)
        if not palace:
            return {"ok": False, "error": "unknown palace_id"}
        desk = desktop or Path.home() / "OneDrive" / "Desktop"
        if not desk.is_dir():
            desk = Path.home() / "Desktop"
        candidates: list[Path] = []
        for pat in ("*.png", "*.jpg", "*.jpeg", "*.webp"):
            candidates.extend(desk.glob(pat))
        if not candidates:
            return {"ok": False, "error": "no image on Desktop"}
        newest = max(candidates, key=lambda p: p.stat().st_mtime)
        study = next(
            (s for s in data["studies"] if s.get("id") == palace.get("study_id")),
            None,
        )
        slug = slug_filename(
            (study or {}).get("notes_rel_path") or palace.get("study_id") or "study"
        )
        num = palace.get("palace_number") or "1"
        ext = newest.suffix.lower() or ".png"
        dest_dir = self.output_dir / "practice" / "images" / slug
        dest_dir.mkdir(parents=True, exist_ok=True)
        dest = dest_dir / f"{num}{ext}"
        shutil.copy2(newest, dest)
        rel = f"practice/images/{slug}/{num}{ext}"
        new_palaces = []
        for p in data["palaces"]:
            if p.get("id") == palace_id:
                pp = dict(p)
                pp["image_rel_path"] = rel
                new_palaces.append(pp)
            else:
                new_palaces.append(p)
        data["palaces"] = new_palaces
        self._save_tree(data, ["palaces"])
        self._sync_practice(palace.get("study_id") or "")
        return {
            "ok": True,
            "palace_id": palace_id,
            "image_rel_path": rel,
            "source": str(newest),
        }

    def regen_all(self) -> dict[str, Any]:
        import subprocess
        import sys

        py_dir = Path(__file__).resolve().parent
        results = []
        for script, args in (
            (
                "study_practice_md.py",
                [
                    "--data-dir",
                    str(self.data_dir),
                    "--output-dir",
                    str(self.output_dir),
                    "--sync-all",
                ],
            ),
            (
                "study_plans_md.py",
                [
                    "--data-dir",
                    str(self.data_dir),
                    "--output-dir",
                    str(self.output_dir),
                    "--studies-root",
                    str(self.studies_root or self.data_dir.parent / "studies"),
                    "--sync-all",
                ],
            ),
        ):
            cmd = [sys.executable, str(py_dir / script), *args]
            try:
                r = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
                results.append(
                    {
                        "script": script,
                        "ok": r.returncode == 0,
                        "stdout": (r.stdout or "")[-500:],
                        "stderr": (r.stderr or "")[-500:],
                    }
                )
            except Exception as e:
                results.append({"script": script, "ok": False, "error": str(e)})
        return {"ok": all(x.get("ok") for x in results), "results": results}
