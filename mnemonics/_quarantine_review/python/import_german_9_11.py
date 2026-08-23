"""Append German palaces 9-11 from archived mnemonics-german.md into existing CSV."""

from __future__ import annotations

import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from data_aggregator import load_all, save_all  # noqa: E402
from migrate_md_to_csv import (  # noqa: E402
    infer_character,
    load_character_names,
    parse_beast_blocks,
    slugify,
    split_streets,
)

STUDY_ID = "STUDY_GERMAN"
FOLDER = "german"
WANT_NUMS = {9, 10, 11}


def next_atom_id(atoms: list[dict[str, str]]) -> int:
    best = 0
    for a in atoms:
        aid = a.get("id", "")
        if aid.startswith("ATOM_"):
            try:
                best = max(best, int(aid[5:]))
            except ValueError:
                pass
    return best + 1


def main() -> int:
    root = Path(__file__).resolve().parent.parent
    data_dir = root / "data"
    md_path = root / "studies" / "german" / "mnemonics-german.md"
    technique_dir = root / "technique"
    if not md_path.exists():
        print(f"Missing {md_path}", file=sys.stderr)
        return 1

    data = load_all(data_dir)
    existing_nums = {
        int(p["palace_number"])
        for p in data["palaces"]
        if p.get("study_id") == STUDY_ID and str(p.get("palace_number", "")).isdigit()
    }
    used_chars = {
        p.get("character_name", "")
        for p in data["palaces"]
        if p.get("study_id") == STUDY_ID and p.get("character_name")
    }
    char_names = load_character_names(technique_dir)

    text = md_path.read_text(encoding="utf-8", errors="replace")
    # Streets 9–11 use plain "Street N:" without ## — normalize for split_streets
    lines_out: list[str] = []
    for line in text.splitlines():
        if re.match(r"^Street\s+\d+\s*:", line):
            lines_out.append("## " + line)
        else:
            lines_out.append(line)
    text = "\n".join(lines_out)

    streets = split_streets(text)
    print(f"Parsed streets: {[n for n, _, _ in streets]}")
    atom_seq = next_atom_id(data["atoms"])
    added_p = added_b = added_a = 0

    for num, title, block in streets:
        if num not in WANT_NUMS:
            continue
        if num in existing_nums:
            print(f"Skip palace {num} (already in CSV)")
            continue

        palace_id = f"PALACE_{slugify(FOLDER)}_{num:02d}"
        if any(p["id"] == palace_id for p in data["palaces"]):
            print(f"Skip {palace_id} (id exists)")
            continue

        char_name, found = infer_character(block, char_names, used_chars)
        if found:
            used_chars.add(char_name)
        else:
            char_name = f"(unassigned palace {num})"
            used_chars.add(char_name)

        beasts_parsed = parse_beast_blocks(block)
        img_rel = f"practice/images/german/{num}.png"
        data["palaces"].append(
            {
                "id": palace_id,
                "study_id": STUDY_ID,
                "palace_number": str(num),
                "title": title,
                "character_name": char_name,
                "image_rel_path": img_rel,
                "depth_slots_used": str(min(5, len(beasts_parsed))),
                "image_prompt": "",
            }
        )
        added_p += 1
        existing_nums.add(num)

        for bi, bp in enumerate(beasts_parsed, start=1):
            beast_id = (
                f"BEAST_{slugify(FOLDER)}_{num:02d}_{slugify(bp['peg_code'], 6)}"
            )
            data["beasts"].append(
                {
                    "id": beast_id,
                    "palace_id": palace_id,
                    "peg_code": bp["peg_code"],
                    "beast_name": bp["beast_name"],
                    "beast_source": "",
                    "sensory_channel": bp.get("sensory_channel", ""),
                    "is_smashed": bp.get("is_smashed", "0"),
                    "sort_order": str(bi),
                }
            )
            added_b += 1
            for ai, atom in enumerate(bp.get("atoms", []), start=1):
                data["atoms"].append(
                    {
                        "id": f"ATOM_{atom_seq:04d}",
                        "beast_id": beast_id,
                        "kind": atom.get("kind", "single"),
                        "zone": atom.get("zone", ""),
                        "zone_label": atom.get("zone_label", ""),
                        "concept": atom.get("concept", atom.get("context", "")),
                        "quote": atom.get("quote", ""),
                        "story": atom.get("story", atom.get("narrative", "")),
                        "ipa": atom.get("ipa", ""),
                        "sensory": atom.get(
                            "sensory", atom.get("sensory_channel", "")
                        ),
                        "sort_order": str(ai),
                    }
                )
                atom_seq += 1
                added_a += 1

        print(f"Added palace {num}: {title} ({len(beasts_parsed)} beasts)")

    save_all(data_dir, data)
    print(f"Saved: +{added_p} palaces, +{added_b} beasts, +{added_a} atoms")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
