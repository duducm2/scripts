"""Pure Markdown render helpers for practice export (dashboard-aligned hierarchy).

Mirrors chart_generator.js: groupAtomsByBeast → beast clusters → atom field cards.
No I/O — used by study_practice_md.build_study_markdown().
"""

from __future__ import annotations

from typing import Any

from schemas import format_concept_thought_groups, normalize_atom_keywords

SENSORY_EMOJI = {
    "visual": "👁️",
    "auditory": "👂",
    "tactile": "✋",
    "olfactory": "👃",
    "gustatory": "👅",
    "thermal": "🌡️",
}


def md_escape(text: str) -> str:
    if not text:
        return ""
    return text.replace("\r\n", "\n").strip()


def dash(value: str | None) -> str:
    t = (value or "").strip()
    return t if t else "—"


def group_atoms_by_beast(atoms: list[dict[str, Any]] | None) -> list[dict[str, Any]]:
    """Stable beast order (first-seen), same key as dashboard JS groupAtomsByBeast."""
    groups: list[dict[str, Any]] = []
    seen: dict[str, dict[str, Any]] = {}
    for a in atoms or []:
        key = (a.get("beast") or "").strip() or "—"
        if key not in seen:
            g = {"beast": key, "atoms": []}
            seen[key] = g
            groups.append(g)
        seen[key]["atoms"].append(a)
    return groups


def format_concept(value: str | None) -> str:
    t = md_escape(format_concept_thought_groups(value or ""))
    if not t:
        return "—"
    return f"💡 {t}"


def format_quote(value: str | None) -> str:
    t = md_escape(value or "")
    if not t:
        return "—"
    return f"\u201c{t}\u201d"


def format_sensory(value: str | None) -> str:
    """Dashboard chip order: emoji then channel word (e.g. 👁️ visual)."""
    t = (value or "").strip()
    if not t:
        return "—"
    emoji = SENSORY_EMOJI.get(t.lower(), "")
    body = md_escape(t)
    return f"{emoji} {body}".strip() if emoji else body


def format_keywords_lines(value: str | None) -> list[str]:
    """One [concept word] -> [tangible keyword] pair per line; empty if none."""
    normalized = normalize_atom_keywords(value)
    if not normalized:
        return []
    out: list[str] = []
    for pair in normalized.split(" || "):
        pair = pair.strip()
        if not pair:
            continue
        if " | " in pair:
            left, right = pair.split(" | ", 1)
        elif "|" in pair:
            left, right = pair.split("|", 1)
        else:
            continue
        left, right = left.strip(), right.strip()
        if left and right:
            out.append(f"[{right}] \u2192 [{left}]")
    return out


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

    lines.extend(format_field_block("Concept", format_concept(atom.get("concept"))))
    kw_lines = format_keywords_lines(atom.get("keywords"))
    if kw_lines:
        lines.append("🔑 **Keywords**")
        lines.append("")
        lines.extend(f"- {row}" for row in kw_lines)
        lines.append("")
    else:
        lines.append("**Keywords**")
        lines.append("_No keywords yet_")
        lines.append("")
    lines.extend(format_field_block("Quote", format_quote(atom.get("quote"))))
    lines.extend(format_field_block("Sensory", format_sensory(atom.get("sensory"))))
    lines.extend(format_field_block("Story", dash(atom.get("story"))))
    return lines


def render_beast_cluster_md(beast: str, atoms: list[dict[str, Any]]) -> list[str]:
    """Flat beast block (not collapsible) — only Memory Palaces use <details>."""
    beast_label = dash(beast)
    lines: list[str] = [
        f"### 🟧 {beast_label}",
        "",
    ]
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
