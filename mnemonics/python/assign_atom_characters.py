"""Assign Characters and Justifications to Knowledge Atoms based on their primary keyword.

Can inspect unassigned atoms, enforce a strict ±6 memory palaces proximity buffer to prevent
character repetition, generate a prompt for batch assignment, or apply assignments from a
CSV/JSON file to mnemonics/data/atoms.csv, then sync practice markdown.
"""

from __future__ import annotations

import argparse
import csv
import sys
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent))

from schemas import ATOMS_HEADERS, extract_primary_keyword
from study_practice_md import sync_all


def load_pao_catalog(data_dir: Path) -> list[dict[str, str]]:
    pao_file = data_dir / "pao_characters.csv"
    if not pao_file.exists():
        return []
    with pao_file.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f, delimiter=";")
        return list(reader)


def load_palace_and_beast_maps(
    data_dir: Path,
) -> tuple[dict[str, dict[str, Any]], dict[str, dict[str, Any]]]:
    """Map beast_id -> beast row, and palace_id -> palace row."""
    beasts_map: dict[str, dict[str, Any]] = {}
    beasts_file = data_dir / "beasts.csv"
    if beasts_file.exists():
        with beasts_file.open("r", encoding="utf-8-sig", newline="") as f:
            for r in csv.DictReader(f):
                bid = r.get("\ufeffid") or r.get("id") or ""
                if bid:
                    beasts_map[bid] = r

    palaces_map: dict[str, dict[str, Any]] = {}
    palaces_file = data_dir / "palaces.csv"
    if palaces_file.exists():
        with palaces_file.open("r", encoding="utf-8-sig", newline="") as f:
            for r in csv.DictReader(f):
                pid = r.get("\ufeffid") or r.get("id") or ""
                if pid:
                    palaces_map[pid] = r

    return beasts_map, palaces_map


def build_study_palace_character_index(
    atoms: list[dict[str, str]],
    beasts_map: dict[str, dict[str, Any]],
    palaces_map: dict[str, dict[str, Any]],
) -> dict[str, dict[int, set[str]]]:
    """Return map: study_id -> { palace_number: set(characters_used) }."""
    index: dict[str, dict[int, set[str]]] = {}
    for a in atoms:
        char = (a.get("character") or "").strip()
        if not char:
            continue
        bid = a.get("beast_id", "")
        beast = beasts_map.get(bid) or {}
        pid = beast.get("palace_id", "")
        palace = palaces_map.get(pid) or {}
        sid = palace.get("study_id", "UNKNOWN")
        try:
            num = int(palace.get("palace_number", "0") or 0)
        except ValueError:
            num = 0

        index.setdefault(sid, {}).setdefault(num, set()).add(char.lower())
    return index


def get_forbidden_nearby_characters(
    study_id: str,
    palace_number: int,
    char_index: dict[str, dict[int, set[str]]],
    min_distance: int = 6,
) -> set[str]:
    """Return set of lower-cased characters used in palaces within ±min_distance."""
    study_map = char_index.get(study_id, {})
    forbidden: set[str] = set()
    for num, chars in study_map.items():
        if abs(num - palace_number) <= min_distance:
            forbidden.update(chars)
    return forbidden


def build_batch_prompt(
    unassigned: list[dict[str, str]],
    pao_catalog: list[dict[str, str]],
    beasts_map: dict[str, dict[str, Any]],
    palaces_map: dict[str, dict[str, Any]],
    char_index: dict[str, dict[int, set[str]]],
) -> str:
    lines = [
        "You are the Character Matcher for the Memory Palace system.",
        "Your task is to assign a specific Character and a Justification to each Knowledge Atom.",
        "",
        "RULES (HARD MANDATES):",
        "1. Base each character assignment strictly on the atom's PRIMARY KEYWORD.",
        "2. The Justification must be 1-2 sentences explaining why this character relates to that specific keyword.",
        "3. Do NOT invent or simulate any interactions between the character and the beast.",
        "4. NO-REPETITION PROXIMITY RULE (±6 MEMORY PALACES BUFFER):",
        "   A character assigned to an atom MUST NOT appear in any memory palace within AT LEAST 6 palaces of the current one.",
        "   Do not reuse the same character in the same palace or in the 6 preceding/succeeding palaces.",
        "5. ALWAYS ASSIGN / BEST JUDGMENT MANDATE:",
        "   You MUST always assign an iconic character. If a direct link to the primary keyword is not immediately obvious,",
        "   make your best creative/metaphorical judgment to find an iconic figure and explain the link. NEVER leave character empty.",
        "6. Output a pure CSV table with columns: id,character,justification",
        "",
        "Available PAO characters (can also use well-known iconic pop-culture/cinema/history figures):",
    ]
    sample_chars = [c.get("name", "") for c in pao_catalog[:50] if c.get("name")]
    if sample_chars:
        lines.append(", ".join(sample_chars) + "...")
    lines.extend(
        [
            "",
            "Atoms to assign:",
            "id | Primary Keyword | Palace (Num) | Forbidden Nearby Characters (±6 palaces) | Concept",
        ]
    )
    for a in unassigned:
        kw = a.get("keywords", "")
        primary = extract_primary_keyword(kw)
        concept = a.get("concept", "")
        aid = a.get("id", "")
        bid = a.get("beast_id", "")
        beast = beasts_map.get(bid, {})
        pid = beast.get("palace_id", "")
        palace = palaces_map.get(pid, {})
        sid = palace.get("study_id", "STUDY")
        try:
            num = int(palace.get("palace_number", "0") or 0)
        except ValueError:
            num = 0
        forbidden = get_forbidden_nearby_characters(
            sid, num, char_index, min_distance=6
        )
        forb_str = ", ".join(sorted(forbidden)) if forbidden else "none"
        lines.append(
            f"{aid} | {primary} | {pid} (#{num}) | Forbidden: [{forb_str}] | {concept}"
        )
    lines.extend(
        [
            "",
            "Output format (pure CSV only):",
            "```csv",
            "id,character,justification",
            "<id>,<character>,<justification>",
            "```",
        ]
    )
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Assign characters and justifications to atoms."
    )
    parser.add_argument(
        "--data-dir",
        type=Path,
        default=Path(__file__).resolve().parent.parent / "data",
        help="Path to mnemonics/data directory",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path(__file__).resolve().parent.parent / "output",
        help="Path to mnemonics/output directory",
    )
    parser.add_argument(
        "--report",
        action="store_true",
        help="Report unassigned atoms and primary keywords",
    )
    parser.add_argument(
        "--prompt",
        action="store_true",
        help="Output a formatted prompt to assign characters via LLM",
    )
    parser.add_argument(
        "--apply",
        type=Path,
        help="Path to CSV file with id,character,justification to merge into atoms.csv",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Simulate changes without writing to disk",
    )
    args = parser.parse_args(argv)

    data_dir = args.data_dir.resolve()
    output_dir = args.output_dir.resolve()
    atoms_path = data_dir / "atoms.csv"

    if not atoms_path.exists():
        print(f"Error: {atoms_path} does not exist.", file=sys.stderr)
        return 1

    with atoms_path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        all_atoms = list(reader)

    beasts_map, palaces_map = load_palace_and_beast_maps(data_dir)
    char_index = build_study_palace_character_index(all_atoms, beasts_map, palaces_map)

    unassigned = [
        a
        for a in all_atoms
        if not (a.get("character") or "").strip()
        or not (a.get("justification") or "").strip()
    ]

    pao_catalog = load_pao_catalog(data_dir)

    if args.report or (not args.prompt and not args.apply):
        print(f"Total atoms: {len(all_atoms)}")
        print(f"Unassigned atoms: {len(unassigned)}")
        print(f"PAO characters in catalog: {len(pao_catalog)}")

        # Check existing collisions within ±6 palaces
        collisions = 0
        checked_chars: dict[tuple[str, str], int] = {}
        for a in all_atoms:
            ch = (a.get("character") or "").strip().lower()
            if not ch:
                continue
            bid = a.get("beast_id", "")
            beast = beasts_map.get(bid, {})
            pid = beast.get("palace_id", "")
            palace = palaces_map.get(pid, {})
            sid = palace.get("study_id", "")
            try:
                num = int(palace.get("palace_number", "0") or 0)
            except ValueError:
                num = 0

            key = (sid, ch)
            if key in checked_chars:
                prev_num = checked_chars[key]
                if abs(num - prev_num) <= 6 and num != prev_num:
                    print(
                        f"  [COLLISION] Character '{ch}' appears in {sid} palace {prev_num} and {num} (distance {abs(num - prev_num)} <= 6)!"
                    )
                    collisions += 1
            checked_chars[key] = num

        if collisions == 0:
            print("Proximity check: 0 collisions detected within ±6 palaces window.")

        if unassigned:
            print("\nSample unassigned atoms:")
            for a in unassigned[:10]:
                prim = extract_primary_keyword(a.get("keywords"))
                print(
                    f"  [{a.get('id')}] Primary Keyword: '{prim}' -> Concept: {a.get('concept', '')[:60]}..."
                )
            if len(unassigned) > 10:
                print(f"  ... and {len(unassigned) - 10} more.")

    if args.prompt:
        prompt_text = build_batch_prompt(
            unassigned, pao_catalog, beasts_map, palaces_map, char_index
        )
        print(prompt_text)

    if args.apply:
        apply_path = args.apply.resolve()
        if not apply_path.exists():
            print(f"Error: {apply_path} does not exist.", file=sys.stderr)
            return 1

        with apply_path.open("r", encoding="utf-8-sig", newline="") as f:
            apply_reader = csv.DictReader(f)
            apply_map = {
                r.get("id", "").strip(): r for r in apply_reader if r.get("id")
            }

        updated_count = 0
        collision_warnings = 0
        for a in all_atoms:
            aid = a.get("id", "").strip()
            if aid in apply_map:
                row_update = apply_map[aid]
                new_char = (row_update.get("character") or "").strip()
                new_just = (row_update.get("justification") or "").strip()
                if new_char:
                    # Check proximity collision
                    bid = a.get("beast_id", "")
                    beast = beasts_map.get(bid, {})
                    pid = beast.get("palace_id", "")
                    palace = palaces_map.get(pid, {})
                    sid = palace.get("study_id", "")
                    try:
                        num = int(palace.get("palace_number", "0") or 0)
                    except ValueError:
                        num = 0
                    forbidden = get_forbidden_nearby_characters(
                        sid, num, char_index, min_distance=6
                    )
                    if new_char.lower() in forbidden:
                        print(
                            f"  [WARNING] '{new_char}' for {aid} collides with a character used within ±6 palaces in {sid}!"
                        )
                        collision_warnings += 1
                    a["character"] = new_char
                    # Update index dynamically
                    char_index.setdefault(sid, {}).setdefault(num, set()).add(
                        new_char.lower()
                    )
                if new_just:
                    a["justification"] = new_just
                updated_count += 1

        print(
            f"Matched and updated {updated_count} atoms ({collision_warnings} proximity warnings)."
        )
        if not args.dry_run:
            with atoms_path.open("w", encoding="utf-8", newline="") as f:
                writer = csv.DictWriter(f, fieldnames=ATOMS_HEADERS)
                writer.writeheader()
                writer.writerows(all_atoms)
            print("Wrote updated atoms.csv.")
            print("Syncing practice markdown...")
            sync_all(data_dir, output_dir, None)
            print("Practice markdown synced successfully.")
        else:
            print("Dry run: no files written.")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
