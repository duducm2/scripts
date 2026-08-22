"""Load and aggregate Memory Palace CSV data."""
from __future__ import annotations

import csv
from pathlib import Path
from typing import Any

from schemas import HEADERS, validate_dataset


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
    return {
        kind: _read_csv(data_dir / f"{kind}.csv")
        for kind in ("studies", "streets", "beasts", "atoms")
    }


def save_all(data_dir: Path, data: dict[str, list[dict[str, str]]]) -> None:
    for kind, headers in HEADERS.items():
        _write_csv(data_dir / f"{kind}.csv", headers, data.get(kind, []))


def resolve_image(notes_root: Path | None, image_rel: str) -> Path | None:
    if not image_rel:
        return None
    p = Path(image_rel)
    if p.is_absolute() and p.exists():
        return p
    if notes_root is None:
        return None
    cand = notes_root / Path(image_rel.replace("/", "\\") if "\\" in image_rel else image_rel)
    # Path handles both separators on Windows
    cand = notes_root / image_rel.replace("\\", "/")
    if cand.exists():
        return cand
    # try with backslashes
    cand2 = notes_root / image_rel.replace("/", "\\")
    if cand2.exists():
        return cand2
    return None


def snapshot(
    data_dir: Path,
    notes_root: Path | None = None,
    study_id: str | None = None,
) -> dict[str, Any]:
    data = load_all(data_dir)
    studies = [s for s in data["studies"] if s.get("active", "1") != "0"]
    if study_id:
        studies = [s for s in studies if s.get("id") == study_id] or studies

    streets = data["streets"]
    beasts = data["beasts"]
    atoms = data["atoms"]

    issues = validate_dataset(data["studies"], streets, beasts, atoms)

    study_cards: list[dict[str, Any]] = []
    for study in studies:
        sid = study["id"]
        study_streets = [st for st in streets if st.get("study_id") == sid]
        street_cards: list[dict[str, Any]] = []
        for st in sorted(study_streets, key=lambda x: int(x.get("street_number") or 0)):
            st_id = st["id"]
            st_beasts = [b for b in beasts if b.get("street_id") == st_id]
            atom_count = sum(
                1 for a in atoms if any(b["id"] == a.get("beast_id") for b in st_beasts)
            )
            img_path = resolve_image(notes_root, st.get("image_rel_path", ""))
            street_cards.append(
                {
                    "id": st_id,
                    "number": st.get("street_number", ""),
                    "title": st.get("title", ""),
                    "character": st.get("character_name", ""),
                    "image_rel": st.get("image_rel_path", ""),
                    "image_abs": str(img_path) if img_path else "",
                    "image_exists": bool(img_path),
                    "beast_count": len(st_beasts),
                    "atom_count": atom_count,
                    "depth_slots": st.get("depth_slots_used", ""),
                }
            )
        study_cards.append(
            {
                "id": sid,
                "slug": study.get("slug", ""),
                "title": study.get("title", ""),
                "notes_rel_path": study.get("notes_rel_path", ""),
                "street_count": len(street_cards),
                "streets": street_cards,
            }
        )

    return {
        "studies": study_cards,
        "totals": {
            "studies": len(data["studies"]),
            "streets": len(streets),
            "beasts": len(beasts),
            "atoms": len(atoms),
        },
        "issues": issues,
        "all_studies": [
            {"id": s["id"], "title": s.get("title", ""), "slug": s.get("slug", "")}
            for s in data["studies"]
            if s.get("active", "1") != "0"
        ],
        "selected_study_id": study_id or (studies[0]["id"] if studies else ""),
    }
