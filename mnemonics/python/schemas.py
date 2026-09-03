"""CSV schema constants and validators for Memory Palace."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

STUDIES_HEADERS = ["id", "title", "notes_rel_path", "sort_order", "active"]
PALACES_HEADERS = [
    "id",
    "study_id",
    "palace_number",
    "title",
    "character_name",
    "image_rel_path",
    "depth_slots_used",
    "image_prompt",
    "palace_notes",
]
PALACE_IMAGES_HEADERS = [
    "id",
    "palace_id",
    "image_rel_path",
    "caption",
    "sort_order",
]
BEASTS_HEADERS = [
    "id",
    "palace_id",
    "peg_code",
    "beast_name",
    "beast_source",
    "sensory_channel",
    "is_smashed",
    "sort_order",
]
ATOMS_HEADERS = [
    "id",
    "beast_id",
    "kind",
    "zone",
    "zone_label",
    "concept",
    "keywords",
    "quote",
    "story",
    "sensory",
    "ipa",
    "sort_order",
]

# Max Keyword | Recognizable Word pairs per Knowledge Atom (CSV `keywords` column).
ATOM_KEYWORDS_MAX_PAIRS = 10


def normalize_atom_keywords(raw: str | None) -> str:
    """Normalize `keywords` to `Keyword | Word || …` with at most ATOM_KEYWORDS_MAX_PAIRS pairs."""
    text = (raw or "").strip()
    if not text:
        return ""
    chunks = [c.strip() for c in re.split(r"\s*\|\|\s*", text) if c.strip()]
    pairs: list[str] = []
    for chunk in chunks:
        if " | " in chunk:
            left, right = chunk.split(" | ", 1)
        elif "|" in chunk:
            left, right = chunk.split("|", 1)
        else:
            continue
        left, right = left.strip(), right.strip()
        if not left or not right:
            continue
        pairs.append(f"{left} | {right}")
        if len(pairs) >= ATOM_KEYWORDS_MAX_PAIRS:
            break
    return " || ".join(pairs)


PLANS_HEADERS = ["id", "study_id", "title", "sort_order", "active"]
PLAN_ITEMS_HEADERS = [
    "id",
    "plan_id",
    "section_path",
    "text",
    "checked",
    "sort_order",
]
PLAN_RESOURCES_HEADERS = ["id", "plan_id", "section_path", "line", "sort_order"]

HEADERS = {
    "studies": STUDIES_HEADERS,
    "palaces": PALACES_HEADERS,
    "palace_images": PALACE_IMAGES_HEADERS,
    "beasts": BEASTS_HEADERS,
    "atoms": ATOMS_HEADERS,
    "plans": PLANS_HEADERS,
    "plan_items": PLAN_ITEMS_HEADERS,
    "plan_resources": PLAN_RESOURCES_HEADERS,
}


def validate_beast_atoms(atoms_for_beast: list[dict[str, Any]]) -> str | None:
    singles = 0
    zoned = 0
    for a in atoms_for_beast:
        kind = (a.get("kind") or "single").strip().lower()
        if kind in ("zoned", "subtopic"):
            zoned += 1
        else:
            singles += 1
    if singles and zoned:
        return "Beast cannot mix a single Knowledge Atom with zoned Knowledge Atoms."
    if singles > 1:
        return "Beast may carry only one comprehensive Knowledge Atom."
    if zoned > 4:
        return "Beast may carry at most four zoned Knowledge Atoms (Z1–Z4)."
    return None


def validate_dataset(
    studies: list[dict[str, str]],
    palaces: list[dict[str, str]],
    beasts: list[dict[str, str]],
    atoms: list[dict[str, str]],
    palace_images: list[dict[str, str]] | None = None,
) -> list[str]:
    issues: list[str] = []
    study_ids = {s["id"] for s in studies if s.get("id")}
    palace_ids = {s["id"] for s in palaces if s.get("id")}
    beast_ids = {b["id"] for b in beasts if b.get("id")}

    for st in palaces:
        if st.get("study_id") not in study_ids:
            issues.append(
                f"Palace {st.get('id')} has unknown study_id {st.get('study_id')}"
            )
        if not (st.get("character_name") or "").strip():
            issues.append(f"Palace {st.get('id')} missing character_name")
        elif (st.get("character_name") or "").startswith("(unassigned"):
            issues.append(
                f"Palace {st.get('id')} has placeholder character (fill via Palaces module)"
            )

    chars_by_study: dict[str, set[str]] = {}
    for st in palaces:
        sid = st.get("study_id") or ""
        ch = (st.get("character_name") or "").strip()
        if not ch:
            continue
        chars_by_study.setdefault(sid, set())
        if ch in chars_by_study[sid]:
            issues.append(f"Duplicate character '{ch}' in study {sid}")
        chars_by_study[sid].add(ch)

    beasts_on_palace: dict[str, int] = {}
    for b in beasts:
        if b.get("palace_id") not in palace_ids:
            issues.append(
                f"Beast {b.get('id')} has unknown palace_id {b.get('palace_id')}"
            )
        palace = b.get("palace_id") or ""
        beasts_on_palace[palace] = beasts_on_palace.get(palace, 0) + 1
        peg = (b.get("peg_code") or "").strip()
        if peg.isdigit():
            issues.append(f"Beast {b.get('id')} has numeric peg_code")

    for palace, count in beasts_on_palace.items():
        if count > 5:
            issues.append(f"Palace {palace} has {count} beasts (max 5)")

    by_beast: dict[str, list[dict[str, str]]] = {}
    for a in atoms:
        bid = a.get("beast_id") or ""
        if bid not in beast_ids:
            issues.append(f"Atom {a.get('id')} has unknown beast_id {bid}")
        by_beast.setdefault(bid, []).append(a)

    for bid, group in by_beast.items():
        err = validate_beast_atoms(group)
        if err:
            issues.append(f"Beast {bid}: {err}")

    for img in palace_images or []:
        if img.get("palace_id") not in palace_ids:
            issues.append(
                f"Palace image {img.get('id')} has unknown palace_id {img.get('palace_id')}"
            )

    return issues


def ensure_data_dir(data_dir: Path) -> None:
    data_dir.mkdir(parents=True, exist_ok=True)
    (data_dir / "imported").mkdir(parents=True, exist_ok=True)
