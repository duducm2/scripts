"""Pure Markdown render helpers for practice export (dashboard-aligned hierarchy).

Mirrors chart_generator.js: groupAtomsByBeast → beast clusters → atom field cards.
No I/O — used by study_practice_md.build_study_markdown().
"""

from __future__ import annotations

from typing import Any

from beast_thumb_base import beast_thumb_md_image, short_label
from keyword_images import practice_markdown_row
from schemas import (
    embed_keyword_mnemonics,
    format_concept_thought_groups,
)


def md_escape(text: str) -> str:
    if not text:
        return ""
    return text.replace("\r\n", "\n").strip()


def dash(value: str | None) -> str:
    t = (value or "").strip()
    return t if t else "—"


def _peg_fold(value: str | None) -> str:
    return (value or "").strip().casefold()


def _atom_sort_order(atom: dict[str, Any]) -> int:
    try:
        return int(atom.get("sort_order") or 0)
    except (TypeError, ValueError):
        return 0


def _desc_text(text: str) -> str:
    """Invert code points so an ascending sort reads Z to A."""
    return "".join(chr(0x10FFFF - ord(ch)) for ch in text)


def palace_recency_key(palace: dict[str, Any]) -> int:
    """Newest palace first (the one normally recalled)."""
    try:
        return -int(palace.get("number") or 0)
    except (TypeError, ValueError):
        return 0


def group_atoms_by_beast(atoms: list[dict[str, Any]] | None) -> list[dict[str, Any]]:
    """Group atoms by beast, then order groups from the latest peg backward.

    Atoms inside one beast stay in sort_order.
    """
    groups: list[dict[str, Any]] = []
    seen: dict[str, dict[str, Any]] = {}
    for a in atoms or []:
        key = (a.get("beast") or "").strip() or "—"
        if key not in seen:
            g = {"beast": key, "atoms": []}
            seen[key] = g
            groups.append(g)
        seen[key]["atoms"].append(a)
    for g in groups:
        g["atoms"].sort(key=_atom_sort_order)

    def _group_key(group: dict[str, Any]) -> tuple[int, str, int]:
        first = (group.get("atoms") or [{}])[0]
        peg = _peg_fold(first.get("peg_code"))
        return (0 if peg else 1, _desc_text(peg), _atom_sort_order(first))

    groups.sort(key=_group_key)
    return groups


def format_concept(value: str | None, keywords: str | None = None) -> str:
    t = md_escape(format_concept_thought_groups(value or ""))
    if not t:
        return "—"
    t = embed_keyword_mnemonics(t, keywords)
    return f"💡 {t}"


def format_quote(value: str | None) -> str:
    t = md_escape(value or "")
    if not t:
        return "—"
    return f"\u201c{t}\u201d"


def format_field_block(label: str, body: str) -> list[str]:
    """Stacked label + body (mobile-friendly; matches dashboard .lbl / .field)."""
    return [f"**{label}**", body, ""]


def render_atom_block_md(atom: dict[str, Any]) -> list[str]:
    lines: list[str] = []
    zone = (atom.get("zone") or "").strip()
    zone_label = (atom.get("zone_label") or "").strip()
    if zone or zone_label:
        if zone and zone_label:
            tag = f"{zone} · {zone_label}"
        else:
            tag = zone or zone_label
        lines.append(f"🟦 **{tag}**")
        lines.append("")

    concept = format_concept(atom.get("concept"), atom.get("keywords"))
    pictures = practice_markdown_row(atom.get("keywords"))
    if pictures:
        concept = f"{concept}\n\n{pictures}"
    lines.extend(format_field_block("Concept", concept))
    lines.extend(format_field_block("Quote", format_quote(atom.get("quote"))))
    return lines


def render_beast_cluster_md(beast: str, atoms: list[dict[str, Any]]) -> list[str]:
    """Flat beast block (not collapsible) — only Memory Palaces use <details>."""
    beast_label = dash(beast)
    lines: list[str] = [
        f"### {beast_label}",
        "",
    ]
    # Thumb from first atom's peg/name when available
    peg = ""
    name = ""
    if atoms:
        peg = (atoms[0].get("peg_code") or "").strip()
        name = short_label(atoms[0].get("beast_name") or "")
    if not name and beast_label and beast_label != "—":
        # Fallback: strip "[P] " prefix from cluster key
        raw = beast_label
        if raw.startswith("[") and "]" in raw:
            name = raw[raw.index("]") + 1 :].strip()
            peg = peg or raw[1 : raw.index("]")].strip()
        else:
            name = raw
    img = beast_thumb_md_image(
        name or beast_label,
        code=peg or None,
        source=(atoms[0].get("beast_source") if atoms else None) or None,
        from_dir="practice",
        width=64,
    )
    if img:
        lines.append(img)
        lines.append("")
    for i, atom in enumerate(atoms):
        if i > 0:
            lines.append("---")
            lines.append("")
        lines.extend(render_atom_block_md(atom))
    return lines


def _beast_noun(n: int) -> str:
    return "beast" if n == 1 else "beasts"


def _atom_noun(n: int) -> str:
    return "Knowledge Atom" if n == 1 else "Knowledge Atoms"


def render_palace_section_md(
    palace: dict[str, Any],
    image_rel: str,
    *,
    open_default: bool = False,
    gallery_md_paths: list[tuple[str, str]] | None = None,
) -> list[str]:
    num = palace.get("number", "")
    ptitle = md_escape(palace.get("title") or "")
    character = dash(md_escape(palace.get("character") or ""))
    beast_count = int(palace.get("beast_count") or 0)
    atom_count = int(palace.get("atom_count") or 0)

    open_attr = " open" if open_default else ""
    atom_short = "atom" if atom_count == 1 else "atoms"
    summary = (
        f"<strong>Memory Palace {num}: {ptitle}</strong>"
        f" · Character: {character}"
        f" · {beast_count} {_beast_noun(beast_count)}"
        f" · {atom_count} {atom_short}"
    )

    lines: list[str] = [
        f"<details{open_attr}>",
        f"<summary>{summary}</summary>",
        "",
    ]

    if image_rel:
        lines.append(f"![Memory Palace {num}]({image_rel})")
    else:
        lines.append("_No image_")
    lines.append("")
    lines.append(
        f"<p><em>{beast_count} {_beast_noun(beast_count)} · {atom_count} {_atom_noun(atom_count)}</em></p>"
    )
    lines.append("")

    atoms = palace.get("atoms") or []
    if not atoms:
        lines.append("_No Knowledge Atoms on this Memory Palace._")
        lines.append("")
    else:
        lines.append("#### Knowledge Atoms")
        lines.append("")
        for group in group_atoms_by_beast(atoms):
            lines.extend(render_beast_cluster_md(group["beast"], group["atoms"]))

    notes = md_escape(palace.get("palace_notes") or "")
    lines.append("#### Notes")
    lines.append("")
    if notes:
        lines.append(notes)
    else:
        lines.append("_No notes._")
    lines.append("")

    lines.append("#### Gallery")
    lines.append("")
    gallery = gallery_md_paths or []
    if not gallery:
        lines.append("_No gallery images._")
        lines.append("")
    else:
        for caption, rel in gallery:
            cap = md_escape(caption) if caption else f"Gallery image"
            lines.append(f"![{cap}]({rel})")
            lines.append("")

    lines.append("</details>")
    lines.append("")
    return lines
