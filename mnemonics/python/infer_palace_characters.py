"""Infer Memory Palace characters from legacy mnemonics-*.md blocks.

Updates palaces.csv character_name when a canon characters.json name appears
in the matching ## Street N body. Does not invent names outside the canon.
"""

from __future__ import annotations

import argparse
import csv
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from migrate_md_to_csv import (  # noqa: E402
    discover_mnemonic_files,
    infer_character,
    load_character_names,
    split_streets,
)


def _read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        return [{k: (v or "") for k, v in row.items()} for row in csv.DictReader(f)]


def _write_csv(path: Path, headers: list[str], rows: list[dict[str, str]]) -> None:
    with path.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=headers, extrasaction="ignore")
        w.writeheader()
        for row in rows:
            w.writerow({h: row.get(h, "") for h in headers})


def build_street_blocks(
    notes_root: Path,
) -> dict[tuple[str, int], str]:
    """Map (folder, street_number) -> street body text."""
    out: dict[tuple[str, int], str] = {}
    for md_path in discover_mnemonic_files(notes_root):
        rel = md_path.relative_to(notes_root).parts
        folder = rel[0] if rel else md_path.stem
        text = md_path.read_text(encoding="utf-8", errors="replace")
        for num, _title, block in split_streets(text):
            out[(folder, num)] = block
    return out


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="Infer palace characters from MD")
    p.add_argument("--data-dir", type=Path, required=True)
    p.add_argument("--notes-root", type=Path, required=True)
    p.add_argument(
        "--technique-dir",
        type=Path,
        default=None,
        help="Default: mnemonics/technique, else <notes-root>/technique",
    )
    p.add_argument("--dry-run", action="store_true")
    args = p.parse_args(argv)

    data_dir = args.data_dir.resolve()
    notes_root = args.notes_root.resolve()
    script_technique = Path(__file__).resolve().parent.parent / "technique"
    if args.technique_dir:
        technique_dir = args.technique_dir.resolve()
    elif script_technique.is_dir() and (script_technique / "characters.json").exists():
        technique_dir = script_technique
    else:
        technique_dir = (notes_root / "technique").resolve()

    palaces_path = data_dir / "palaces.csv"
    studies_path = data_dir / "studies.csv"
    beasts_path = data_dir / "beasts.csv"
    atoms_path = data_dir / "atoms.csv"
    if not palaces_path.exists() or not studies_path.exists():
        print("Missing palaces.csv or studies.csv", file=sys.stderr)
        return 1

    names = load_character_names(technique_dir)
    if not names:
        print(f"No characters loaded from {technique_dir}", file=sys.stderr)
        return 1

    studies = {s["id"]: s for s in _read_csv(studies_path)}
    palaces = _read_csv(palaces_path)
    beasts = _read_csv(beasts_path) if beasts_path.exists() else []
    atoms = _read_csv(atoms_path) if atoms_path.exists() else []
    blocks = build_street_blocks(notes_root)

    # Concatenate beast + atom story text per palace for secondary inference
    beasts_by_palace: dict[str, list[dict[str, str]]] = {}
    for b in beasts:
        beasts_by_palace.setdefault(b.get("palace_id", ""), []).append(b)
    atoms_by_beast: dict[str, list[dict[str, str]]] = {}
    for a in atoms:
        atoms_by_beast.setdefault(a.get("beast_id", ""), []).append(a)

    def story_blob(palace_id: str) -> str:
        parts: list[str] = []
        for b in beasts_by_palace.get(palace_id, []):
            parts.append(b.get("beast_name", ""))
            for a in atoms_by_beast.get(b.get("id", ""), []):
                parts.append(a.get("concept", ""))
                parts.append(a.get("quote", ""))
                parts.append(a.get("story", ""))
                parts.append(a.get("sensory", ""))
        return "\n".join(parts)

    headers = [
        "id",
        "study_id",
        "palace_number",
        "title",
        "character_name",
        "image_rel_path",
        "depth_slots_used",
        "image_prompt",
    ]

    used_by_study: dict[str, set[str]] = {}
    assigned = 0
    still_open = 0
    report: list[str] = []

    for st in palaces:
        sid = st.get("study_id", "")
        ch = (st.get("character_name") or "").strip()
        if ch and not ch.startswith("(unassigned"):
            used_by_study.setdefault(sid, set()).add(ch)

    for st in palaces:
        sid = st.get("study_id", "")
        study = studies.get(sid) or {}
        folder = (study.get("notes_rel_path") or "").strip()
        try:
            num = int(st.get("palace_number") or 0)
        except ValueError:
            num = 0
        used = used_by_study.setdefault(sid, set())
        block = blocks.get((folder, num), "")
        current = (st.get("character_name") or "").strip()
        needs = (not current) or current.startswith("(unassigned")

        if needs:
            char_name, found = "", False
            src = "md"
            if block:
                char_name, found = infer_character(block, names, used)
            if not found:
                char_name, found = infer_character(
                    story_blob(st.get("id", "")), names, used
                )
                src = "csv-stories"
            if found:
                st["character_name"] = char_name
                used.add(char_name)
                assigned += 1
                report.append(f"OK {st.get('id')} -> {char_name} ({src})")
                continue

        if needs:
            st["character_name"] = f"(unassigned palace {num})"
            still_open += 1
            report.append(f"OPEN {st.get('id')} (no match)")
        elif current.startswith("(unassigned street"):
            st["character_name"] = current.replace(
                "(unassigned street", "(unassigned palace", 1
            )
            still_open += 1
            report.append(f"RENAMED placeholder {st.get('id')}")

    report_path = data_dir / "character_inference_report.md"
    lines = [
        "# Palace character inference",
        "",
        f"- Canon characters: {len(names)}",
        f"- Assigned: {assigned}",
        f"- Still unassigned: {still_open}",
        f"- Dry run: {args.dry_run}",
        "",
        "## Log",
        "",
    ]
    lines.extend(f"- {r}" for r in report)
    report_text = "\n".join(lines) + "\n"

    if not args.dry_run:
        _write_csv(palaces_path, headers, palaces)
        report_path.write_text(report_text, encoding="utf-8")

    print(f"characters_canon={len(names)}")
    print(f"assigned={assigned} still_unassigned={still_open} dry_run={args.dry_run}")
    print(f"report={report_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
