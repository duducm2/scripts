"""Load and aggregate Memory Palace CSV data."""

from __future__ import annotations

import csv
import re
from pathlib import Path
from typing import Any

from schemas import HEADERS, validate_dataset

# Legacy Story Architect / smash lines often append channel to concept, e.g.
# "... · sensory: gustatory 👅" while leaving the sensory column empty.
_SENSORY_TAIL_RE = re.compile(
    r"\s*[·•]\s*sensory:\s*([A-Za-z]+)(?:\s+\S+)?\s*$",
    re.IGNORECASE,
)


def normalize_atom_concept_sensory(row: dict[str, str]) -> dict[str, str]:
    """Move trailing `· sensory: channel [emoji]` from concept into sensory."""
    concept = (row.get("concept") or row.get("context") or "").strip()
    sensory = (row.get("sensory") or row.get("sensory_channel") or "").strip()
    m = _SENSORY_TAIL_RE.search(concept)
    if not m:
        return row
    channel = m.group(1).strip().lower()
    cleaned = concept[: m.start()].rstrip()
    out = dict(row)
    out["concept"] = cleaned
    if not sensory:
        out["sensory"] = channel
    return out


def _read_csv(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        return []
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        rows: list[dict[str, str]] = []
        for row in reader:
            rows.append({k: (v if v is not None else "") for k, v in row.items()})
        return rows


def _write_csv(path: Path, headers: list[str], rows: list[dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=headers, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow({h: row.get(h, "") for h in headers})


def load_all(data_dir: Path) -> dict[str, list[dict[str, str]]]:
    data = {
        kind: _read_csv(data_dir / f"{kind}.csv")
        for kind in ("studies", "palaces", "beasts", "atoms")
    }
    data["atoms"] = [normalize_atom_concept_sensory(a) for a in data["atoms"]]
    return data


def save_all(data_dir: Path, data: dict[str, list[dict[str, str]]]) -> None:
    for kind, headers in HEADERS.items():
        _write_csv(data_dir / f"{kind}.csv", headers, data.get(kind, []))


def resolve_image(
    notes_root: Path | None,
    image_rel: str,
    output_dir: Path | None = None,
) -> Path | None:
    if not image_rel:
        return None
    p = Path(image_rel)
    if p.is_absolute() and p.exists():
        return p
    norm = image_rel.replace("\\", "/")
    if output_dir is not None and norm.startswith("practice/images/"):
        cand = output_dir / norm.replace("/", "\\")
        if cand.exists():
            return cand
        cand2 = output_dir / norm
        if cand2.exists():
            return cand2
    if notes_root is None:
        return None
    cand = notes_root / norm
    if cand.exists():
        return cand
    cand2 = notes_root / image_rel.replace("/", "\\")
    if cand2.exists():
        return cand2
    return None


def snapshot(
    data_dir: Path,
    notes_root: Path | None = None,
    study_id: str | None = None,
    output_dir: Path | None = None,
) -> dict[str, Any]:
    data = load_all(data_dir)
    studies = [s for s in data["studies"] if s.get("active", "1") != "0"]
    if study_id:
        studies = [s for s in studies if s.get("id") == study_id] or studies

    palaces = data["palaces"]
    beasts = data["beasts"]
    atoms = data["atoms"]

    issues = validate_dataset(data["studies"], palaces, beasts, atoms)

    study_cards: list[dict[str, Any]] = []
    for study in studies:
        sid = study["id"]
        study_palaces = [st for st in palaces if st.get("study_id") == sid]
        palace_cards: list[dict[str, Any]] = []
        for st in sorted(
            study_palaces, key=lambda x: int(x.get("palace_number") or 0), reverse=True
        ):
            st_id = st["id"]
            st_beasts = [b for b in beasts if b.get("palace_id") == st_id]
            palace_atoms: list[dict[str, Any]] = []
            for b in sorted(st_beasts, key=lambda x: int(x.get("sort_order") or 0)):
                for a in atoms:
                    if a.get("beast_id") != b["id"]:
                        continue
                    palace_atoms.append(
                        {
                            "id": a.get("id", ""),
                            "kind": a.get("kind", ""),
                            "zone": a.get("zone", ""),
                            "zone_label": a.get("zone_label", ""),
                            "concept": a.get("concept", a.get("context", "")),
                            "quote": a.get("quote", ""),
                            "story": a.get("story", a.get("narrative", "")),
                            "sensory": a.get("sensory", a.get("sensory_channel", "")),
                            "beast": f"[{b.get('peg_code', '')}] {b.get('beast_name', '')}".strip(),
                            "sort_order": a.get("sort_order", ""),
                        }
                    )
            atom_count = len(palace_atoms)
            img_path = resolve_image(
                notes_root, st.get("image_rel_path", ""), output_dir
            )
            palace_cards.append(
                {
                    "id": st_id,
                    "number": st.get("palace_number", ""),
                    "title": st.get("title", ""),
                    "character": st.get("character_name", ""),
                    "image_rel": st.get("image_rel_path", ""),
                    "image_abs": str(img_path) if img_path else "",
                    "image_exists": bool(img_path),
                    "image_prompt": st.get("image_prompt", ""),
                    "beast_count": len(st_beasts),
                    "atom_count": atom_count,
                    "depth_slots": st.get("depth_slots_used", ""),
                    "atoms": palace_atoms,
                }
            )
        study_cards.append(
            {
                "id": sid,
                "title": study.get("title", ""),
                "notes_rel_path": study.get("notes_rel_path", ""),
                "palace_count": len(palace_cards),
                "palaces": palace_cards,
            }
        )

    return {
        "studies": study_cards,
        "totals": {
            "studies": len(data["studies"]),
            "palaces": len(palaces),
            "beasts": len(beasts),
            "atoms": len(atoms),
        },
        "issues": issues,
        "all_studies": [
            {
                "id": s["id"],
                "title": s.get("title", ""),
                "notes_rel_path": s.get("notes_rel_path", ""),
            }
            for s in data["studies"]
            if s.get("active", "1") != "0"
        ],
        "selected_study_id": study_id or (studies[0]["id"] if studies else ""),
    }
