"""Migrate legacy mnemonics-*.md files into Memory Palace CSV structure."""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

# Allow running as script from mnemonics/python
sys.path.insert(0, str(Path(__file__).resolve().parent))

from data_aggregator import save_all  # noqa: E402
from schemas import ensure_data_dir, validate_dataset  # noqa: E402

ORANGE = "🟧"
BLUE = "🟦"
BULB = "💡"

RE_TOPIC = re.compile(r"^#\s+(?:Topic:\s*)?(.+?)\s*$", re.M)
RE_STREET = re.compile(
    r"^##\s+Street\s+(\d+)\s*:\s*(.+?)\s*$",
    re.M | re.I,
)
RE_IMAGE = re.compile(r"!\[.*?\]\(([^)]+)\)")
RE_BEAST = re.compile(
    r"(?:(?:🟧|\?\?)\s*)?\*\*\\?\[([^\]]+)\]\s*\\?\[([^\]]+)\]\*\*(?:\s*·\s*sensory:\s*(\w+)\s*\S*)?",
    re.I,
)
# Legacy bare peg lines: [A] [Arachne]  or  [A] [Arachne] · sensory: ...
RE_BEAST_BARE = re.compile(
    r"^\[([A-Za-z]{1,2})\]\s*\[([^\]]+)\](?:\s*·\s*sensory:\s*(\w+)\s*\S*)?\s*$",
    re.M,
)
# Skills-style: 🟧 [A] Arachne   (name not bracketed)
RE_BEAST_LOOSE = re.compile(
    r"(?:🟧|\?\?)\s*\[([A-Za-z]{1,2})\]\s+([^\n·*]+?)(?:\s*·\s*sensory:\s*(\w+)\s*\S*)?\s*$",
    re.M,
)
RE_SUB = re.compile(
    r"🟦\s*\*\*Z([1-4])\s+([^|]+)\|\s*([^:]+):\*\*\s*(.*?)(?:\s*·\s*sensory:\s*(\w+)\s*\S*)?\s*$",
    re.I,
)
# Piano-style smash without zones: 🟦 **Label:** text
RE_SUB_LOOSE = re.compile(
    r"🟦\s*\*\*([^:*]+):\*\*\s*(.*)$",
    re.M,
)
# Even looser: 🟦 Label: text
RE_SUB_PLAIN = re.compile(
    r"🟦\s+\*\*?([^:*\n]+):?\*\*?\s*(.*)$",
    re.M,
)
RE_CONTEXT = re.compile(r"(?:💡\s*)?\*\*Context:\*\*\s*(.+)", re.I)
RE_CONTEXT_LOOSE = re.compile(r"(?:💡\s*)?Context:\s*(.+)", re.I)
RE_QUOTE = re.compile(r"\*\*Quote:\*\*\s*[\"“]?(.+?)[\"”]?\s*$", re.I)
RE_NARRATIVE = re.compile(r"\*\*Narrative(?:\s*\([^)]*\))?:\*\*\s*(.+)", re.I)
RE_CONCLUSION = re.compile(r"\*\*Conclusion(?:\s*\([^)]*\))?:\*\*\s*(.+)", re.I)
RE_IPA = re.compile(r"^IPA:\s*(.+)$", re.I | re.M)


def slugify(text: str, maxlen: int = 12) -> str:
    out = []
    for ch in text.upper():
        if ("A" <= ch <= "Z") or ("0" <= ch <= "9"):
            out.append(ch)
    s = "".join(out)[:maxlen]
    return s or "X"


def load_character_names(technique_dir: Path) -> list[str]:
    path = technique_dir / "characters.json"
    if not path.exists():
        return []
    try:
        text = path.read_text(encoding="utf-8")
        data, _ = json.JSONDecoder().raw_decode(text.lstrip())
    except Exception:
        return []
    names: list[str] = []
    for section in data.get("sections", []):
        for c in section.get("characters", []):
            n = (c.get("name") or "").strip()
            if n:
                names.append(n)
    # longest first for matching
    names.sort(key=len, reverse=True)
    return names


def infer_character(text: str, names: list[str], used: set[str]) -> tuple[str, bool]:
    for name in names:
        if name in used:
            continue
        if name and name in text:
            return name, True
    return "", False


def split_streets(body: str) -> list[tuple[int, str, str]]:
    """Return list of (number, title, block)."""
    matches = list(RE_STREET.finditer(body))
    if not matches:
        return []
    out: list[tuple[int, str, str]] = []
    for i, m in enumerate(matches):
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(body)
        out.append((int(m.group(1)), m.group(2).strip(), body[start:end]))
    return out


def parse_field_block(block: str) -> dict[str, str]:
    context = quote = narrative = ipa = ""
    lines = block.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if not line:
            i += 1
            continue
        m = RE_CONTEXT.search(line) or RE_CONTEXT_LOOSE.search(line)
        if m:
            context = m.group(1).strip()
            i += 1
            continue
        m = RE_QUOTE.search(line)
        if m:
            quote = m.group(1).strip().strip('"“”')
            i += 1
            continue
        m = RE_NARRATIVE.search(line)
        if m:
            narrative = m.group(1).strip()
            i += 1
            # gather continuation until blank or next marker
            while i < len(lines):
                nxt = lines[i].strip()
                if (
                    not nxt
                    or nxt.startswith("**")
                    or nxt.startswith(ORANGE)
                    or nxt.startswith(BLUE)
                    or nxt.startswith("??")
                ):
                    break
                if nxt.upper().startswith("IPA:"):
                    break
                narrative += " " + nxt
                i += 1
            continue
        m = RE_CONCLUSION.search(line)
        if m:
            # legacy: treat as quote if quote empty else narrative
            val = m.group(1).strip()
            if not quote:
                quote = val
            elif not narrative:
                narrative = val
            i += 1
            continue
        if line.upper().startswith("IPA:"):
            ipa = line.split(":", 1)[1].strip()
            i += 1
            continue
        i += 1
    return {
        "context": context,
        "quote": quote,
        "narrative": narrative,
        "ipa": ipa,
    }


def parse_beast_blocks(street_body: str) -> list[dict[str, Any]]:
    # Normalize mojibake orange
    text = street_body.replace("?? **[", f"{ORANGE} **[")
    text = text.replace("??**[", f"{ORANGE} **[")

    # Find beast starts (modern markdown bold headers)
    starts: list[tuple[int, str, str, str]] = []
    for m in RE_BEAST.finditer(text):
        starts.append(
            (
                m.start(),
                m.group(1).strip(),
                m.group(2).strip(),
                (m.group(3) or "").strip().lower(),
            )
        )
    if not starts:
        for m in RE_BEAST_BARE.finditer(text):
            starts.append(
                (
                    m.start(),
                    m.group(1).strip(),
                    m.group(2).strip(),
                    (m.group(3) or "").strip().lower(),
                )
            )
    if not starts:
        for m in RE_BEAST_LOOSE.finditer(text):
            starts.append(
                (
                    m.start(),
                    m.group(1).strip(),
                    m.group(2).strip(),
                    (m.group(3) or "").strip().lower(),
                )
            )
    if not starts:
        return []

    starts.sort(key=lambda x: x[0])
    beasts: list[dict[str, Any]] = []
    for i, (start, peg, name, sensory) in enumerate(starts):
        end = starts[i + 1][0] if i + 1 < len(starts) else len(text)
        chunk = text[start:end]
        # Skip past first line
        nl = chunk.find("\n")
        rest = chunk[nl + 1 :] if nl >= 0 else ""

        # Sub-atoms (canonical Z1–Z4 first)
        sub_matches = list(RE_SUB.finditer(rest))
        atoms: list[dict[str, Any]] = []
        if sub_matches:
            for j, sm in enumerate(sub_matches):
                s_end = (
                    sub_matches[j + 1].start()
                    if j + 1 < len(sub_matches)
                    else len(rest)
                )
                sub_chunk = rest[sm.start() : s_end]
                fields = parse_field_block(sub_chunk)
                if not fields["context"]:
                    after = rest[sm.end() : s_end]
                    cm = RE_CONTEXT.search(after) or RE_CONTEXT_LOOSE.search(after)
                    if cm:
                        fields["context"] = cm.group(1).strip()
                atoms.append(
                    {
                        "kind": "zoned",
                        "zone": f"Z{sm.group(1)}",
                        "zone_label": sm.group(3).strip(),
                        "sensory_channel": (sm.group(5) or "").strip().lower(),
                        **fields,
                    }
                )
        else:
            loose = list(RE_SUB_LOOSE.finditer(rest))
            if not loose:
                loose = [
                    m
                    for m in RE_SUB_PLAIN.finditer(rest)
                    if m.group(0).strip().startswith("🟦")
                ]
            if loose:
                # Cap at 4 zoned Knowledge Atoms (technique max)
                for j, sm in enumerate(loose[:4]):
                    label = sm.group(1).strip()
                    hook = (sm.group(2) or "").strip()
                    fields = parse_field_block(rest[sm.start() :])
                    if not fields["context"]:
                        fields["context"] = hook or label
                    if not fields["narrative"] and hook:
                        fields["narrative"] = hook
                    atoms.append(
                        {
                            "kind": "zoned",
                            "zone": f"Z{j + 1}",
                            "zone_label": label,
                            "sensory_channel": "",
                            **fields,
                        }
                    )

        if atoms and any(a.get("kind") in ("zoned", "subtopic") for a in atoms):
            beasts.append(
                {
                    "peg_code": peg,
                    "beast_name": name,
                    "sensory_channel": sensory,
                    "is_smashed": "1",
                    "atoms": atoms,
                }
            )
        else:
            fields = parse_field_block(rest)
            # Legacy: definition line, narrative prose, then quoted line
            if (
                not fields["context"]
                and not fields["quote"]
                and not fields["narrative"]
            ):
                prose_lines = []
                quoted = ""
                for line in rest.splitlines():
                    t = line.strip()
                    if not t or t.startswith("![") or t.startswith("#"):
                        continue
                    if (t.startswith('"') or t.startswith("“")) and (
                        t.endswith('"') or t.endswith("”")
                    ):
                        quoted = t.strip('"“”')
                        continue
                    if t.startswith("[") and "] [" in t:
                        continue
                    prose_lines.append(t)
                if prose_lines:
                    fields["context"] = prose_lines[0][:240]
                    if len(prose_lines) > 1:
                        fields["narrative"] = " ".join(prose_lines[1:])
                    else:
                        fields["narrative"] = prose_lines[0]
                if quoted:
                    fields["quote"] = quoted
            # Skills Conclusion → quote
            if not fields["quote"]:
                cm = RE_CONCLUSION.search(rest)
                if cm:
                    fields["quote"] = cm.group(1).strip().strip('"“”')
            atoms = [
                {
                    "kind": "single",
                    "zone": "",
                    "zone_label": "",
                    "sensory_channel": sensory,
                    **fields,
                }
            ]
            beasts.append(
                {
                    "peg_code": peg,
                    "beast_name": name,
                    "sensory_channel": sensory,
                    "is_smashed": "0",
                    "atoms": atoms,
                }
            )
    return beasts


def discover_mnemonic_files(notes_root: Path) -> list[Path]:
    files: list[Path] = []
    for path in notes_root.rglob("mnemonics-*.md"):
        if "technique" in path.parts:
            continue
        if "portals" in path.parts:
            continue
        files.append(path)
    return sorted(files)


def migrate(
    notes_root: Path,
    out_dir: Path,
    technique_dir: Path | None = None,
    dry_run: bool = False,
) -> dict[str, Any]:
    ensure_data_dir(out_dir)
    if technique_dir is None:
        technique_dir = notes_root / "technique"
    char_names = load_character_names(technique_dir)

    studies: list[dict[str, str]] = []
    palaces: list[dict[str, str]] = []
    beasts: list[dict[str, str]] = []
    atoms: list[dict[str, str]] = []
    report: list[str] = []
    warnings: list[str] = []

    atom_seq = 0
    files = discover_mnemonic_files(notes_root)
    if not files:
        report.append("No mnemonics-*.md files found under notes root.")

    study_order = 0
    for md_path in files:
        rel_parts = md_path.relative_to(notes_root).parts
        folder = rel_parts[0] if rel_parts else md_path.stem

        study_order += 1
        study_id = f"STUDY_{slugify(folder)}"
        # Reuse study if same folder already migrated (multiple files unlikely)
        existing_study = next(
            (s for s in studies if s.get("notes_rel_path") == folder), None
        )
        if existing_study:
            study_id = existing_study["id"]
            report.append(f"Append to study {study_id} from {md_path}")
        else:
            base_id = study_id
            n = 2
            existing_ids = {s["id"] for s in studies}
            while study_id in existing_ids:
                study_id = f"{base_id}{n}"
                n += 1
            title = folder.replace("-", " ").title()
            studies.append(
                {
                    "id": study_id,
                    "title": title,
                    "notes_rel_path": folder,
                    "sort_order": str(study_order),
                    "active": "1",
                }
            )
            report.append(f"Study {study_id} from {md_path}")

        text = md_path.read_text(encoding="utf-8", errors="replace")
        # Legacy MD still uses "## Street N" headings; emit palaces.csv.
        street_blocks = split_streets(text)
        if not street_blocks:
            warnings.append(f"No Memory Palaces (## Street N) in {md_path}")
            continue

        used_chars: set[str] = set()
        for st in palaces:
            if st.get("study_id") == study_id and st.get("character_name"):
                used_chars.add(st["character_name"])

        for num, st_title, block in street_blocks:
            img_m = RE_IMAGE.search(block)
            img_rel = ""
            if img_m:
                raw = img_m.group(1).replace("\\", "/")
                if raw.startswith("images/"):
                    img_rel = f"{folder}/{raw}"
                elif "/" in raw:
                    img_rel = raw if raw.startswith(folder) else f"{folder}/{raw}"
                else:
                    img_rel = f"{folder}/images/{raw}"
            else:
                # default convention
                for ext in (".png", ".jpg", ".jpeg"):
                    cand = notes_root / folder / "images" / f"{num}{ext}"
                    if cand.exists():
                        img_rel = f"{folder}/images/{num}{ext}"
                        break
                if not img_rel:
                    img_rel = f"{folder}/images/{num}.png"
                    warnings.append(f"Missing image for {folder} palace {num}")

            char_name, found = infer_character(block, char_names, used_chars)
            if found:
                used_chars.add(char_name)
            else:
                char_name = f"(unassigned palace {num})"
                warnings.append(
                    f"Low-confidence character for {folder} palace {num} → placeholder"
                )
                used_chars.add(char_name)

            palace_id = f"PALACE_{slugify(folder)}_{num:02d}"
            beasts_parsed = parse_beast_blocks(block)
            palaces.append(
                {
                    "id": palace_id,
                    "study_id": study_id,
                    "palace_number": str(num),
                    "title": st_title,
                    "character_name": char_name,
                    "image_rel_path": img_rel,
                    "depth_slots_used": str(min(5, len(beasts_parsed))),
                    "image_prompt": "",
                }
            )

            for bi, bp in enumerate(beasts_parsed, start=1):
                beast_id = (
                    f"BEAST_{slugify(folder)}_{num:02d}_{slugify(bp['peg_code'], 6)}"
                )
                beasts.append(
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
                for ai, atom in enumerate(bp.get("atoms", []), start=1):
                    atom_seq += 1
                    atoms.append(
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

            if not beasts_parsed:
                warnings.append(f"No beasts parsed on {folder} palace {num}")

    data = {
        "studies": studies,
        "palaces": palaces,
        "beasts": beasts,
        "atoms": atoms,
    }
    issues = validate_dataset(studies, palaces, beasts, atoms)
    report_path = out_dir / "migration_report.md"
    lines = [
        "# Memory Palace migration report",
        "",
        f"- Studies: {len(studies)}",
        f"- Palaces: {len(palaces)}",
        f"- Beasts: {len(beasts)}",
        f"- Atoms: {len(atoms)}",
        f"- Dry run: {dry_run}",
        "",
        "## Log",
        "",
    ]
    lines.extend(f"- {r}" for r in report)
    if warnings:
        lines += ["", "## Warnings", ""]
        lines.extend(f"- {w}" for w in warnings)
    if issues:
        lines += ["", "## Validation issues", ""]
        lines.extend(f"- {i}" for i in issues)

    report_text = "\n".join(lines) + "\n"
    if not dry_run:
        save_all(out_dir, data)
        report_path.write_text(report_text, encoding="utf-8")
    else:
        report_path.write_text(report_text, encoding="utf-8")

    return {
        "counts": {
            "studies": len(studies),
            "palaces": len(palaces),
            "beasts": len(beasts),
            "atoms": len(atoms),
        },
        "warnings": warnings,
        "issues": issues,
        "report_path": str(report_path),
        "dry_run": dry_run,
    }


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(
        description="Migrate mnemonic markdown to Memory Palace CSV"
    )
    p.add_argument(
        "--notes-root",
        type=Path,
        required=True,
        help="Path to notes/studies",
    )
    p.add_argument(
        "--out",
        type=Path,
        required=True,
        help="Output data directory (mnemonics/data)",
    )
    p.add_argument(
        "--technique-dir",
        type=Path,
        default=None,
        help="Path to studies/technique (characters.json)",
    )
    p.add_argument("--dry-run", action="store_true")
    args = p.parse_args(argv)

    result = migrate(
        notes_root=args.notes_root.resolve(),
        out_dir=args.out.resolve(),
        technique_dir=args.technique_dir.resolve() if args.technique_dir else None,
        dry_run=args.dry_run,
    )
    print(json.dumps(result["counts"], indent=2))
    print(f"Report: {result['report_path']}")
    print(f"Warnings: {len(result['warnings'])}  Issues: {len(result['issues'])}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
