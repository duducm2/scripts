"""CSV schema constants and validators for Memory Palace."""

from __future__ import annotations

from pathlib import Path
from typing import Any

STUDIES_HEADERS = ["id", "slug", "title", "notes_rel_path", "sort_order", "active"]
STREETS_HEADERS = [
    "id",
    "study_id",
    "street_number",
    "title",
    "character_name",
    "image_rel_path",
    "depth_slots_used",
]
BEASTS_HEADERS = [
    "id",
    "street_id",
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
    "context",
    "quote",
    "narrative",
    "ipa",
    "sensory_channel",
    "sort_order",
]

HEADERS = {
    "studies": STUDIES_HEADERS,
    "streets": STREETS_HEADERS,
    "beasts": BEASTS_HEADERS,
    "atoms": ATOMS_HEADERS,
}


def validate_beast_atoms(atoms_for_beast: list[dict[str, Any]]) -> str | None:
    singles = 0
    subs = 0
    for a in atoms_for_beast:
        kind = (a.get("kind") or "single").strip().lower()
        if kind == "subtopic":
            subs += 1
        else:
            singles += 1
    if singles and subs:
        return "Beast cannot mix a single atom with smashed subtopics."
    if singles > 1:
        return "Beast may carry only one comprehensive Knowledge Atom."
    if subs > 4:
        return "Beast may carry at most four smashed subtopics (Z1–Z4)."
    return None


def validate_dataset(
    studies: list[dict[str, str]],
    streets: list[dict[str, str]],
    beasts: list[dict[str, str]],
    atoms: list[dict[str, str]],
) -> list[str]:
    issues: list[str] = []
    study_ids = {s["id"] for s in studies if s.get("id")}
    street_ids = {s["id"] for s in streets if s.get("id")}
    beast_ids = {b["id"] for b in beasts if b.get("id")}

    for st in streets:
        if st.get("study_id") not in study_ids:
            issues.append(
                f"Street {st.get('id')} has unknown study_id {st.get('study_id')}"
            )
        if not (st.get("character_name") or "").strip():
            issues.append(f"Street {st.get('id')} missing character_name")
        elif (st.get("character_name") or "").startswith("(unassigned"):
            issues.append(
                f"Street {st.get('id')} has placeholder character (fill via Streets module)"
            )

    chars_by_study: dict[str, set[str]] = {}
    for st in streets:
        sid = st.get("study_id") or ""
        ch = (st.get("character_name") or "").strip()
        if not ch:
            continue
        chars_by_study.setdefault(sid, set())
        if ch in chars_by_study[sid]:
            issues.append(f"Duplicate character '{ch}' in study {sid}")
        chars_by_study[sid].add(ch)

    beasts_on_street: dict[str, int] = {}
    for b in beasts:
        if b.get("street_id") not in street_ids:
            issues.append(
                f"Beast {b.get('id')} has unknown street_id {b.get('street_id')}"
            )
        street = b.get("street_id") or ""
        beasts_on_street[street] = beasts_on_street.get(street, 0) + 1
        peg = (b.get("peg_code") or "").strip()
        if peg.isdigit():
            issues.append(f"Beast {b.get('id')} has numeric peg_code")

    for street, count in beasts_on_street.items():
        if count > 5:
            issues.append(f"Street {street} has {count} beasts (max 5)")

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

    return issues


def ensure_data_dir(data_dir: Path) -> None:
    data_dir.mkdir(parents=True, exist_ok=True)
    (data_dir / "imported").mkdir(parents=True, exist_ok=True)
