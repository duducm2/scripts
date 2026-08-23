"""Pure Markdown render helpers for practice export (dashboard-aligned hierarchy).

Mirrors chart_generator.js: groupAtomsByBeast → beast clusters → atom field cards.
No I/O — used by study_practice_md.build_study_markdown().
"""

from __future__ import annotations

from typing import Any

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
    t = md_escape(value or "")
    if not t:
        return "—"
    return f"💡 {t}"


def format_quote(value: str | None) -> str:
    t = md_escape(value or "")
    if not t:
        return "—"
    return f"\u201c{t}\u201d"


def format_sensory(value: str | None) -> str:
    t = (value or "").strip()
    if not t:
        return "—"
    emoji = SENSORY_EMOJI.get(t.lower(), "")
    body = md_escape(t)
    return f"{emoji} {body}".strip() if emoji else body


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
        lines.append(f"**{tag}**")
        lines.append("")

    lines.extend(format_field_block("Concept", format_concept(atom.get("concept"))))
    lines.extend(format_field_block("Quote", format_quote(atom.get("quote"))))
    lines.extend(format_field_block("Story", dash(atom.get("story"))))
    lines.extend(format_field_block("Sensory", format_sensory(atom.get("sensory"))))
    return lines


def render_beast_cluster_md(beast: str, atoms: list[dict[str, Any]]) -> list[str]:
    beast_label = dash(beast)
    lines: list[str] = [
        "<details>",
        f"<summary><strong>Beast</strong> · {beast_label}</summary>",
        "",
    ]
    for i, atom in enumerate(atoms):
        if i > 0:
            lines.append("---")
            lines.append("")
        lines.extend(render_atom_block_md(atom))
    lines.append("</details>")
    lines.append("")
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
) -> list[str]:
    num = palace.get("number", "")
    ptitle = md_escape(palace.get("title") or "")
    character = dash(md_escape(palace.get("character") or ""))
    beast_count = int(palace.get("beast_count") or 0)
    atom_count = int(palace.get("atom_count") or 0)
    prompt = md_escape(palace.get("image_prompt") or "")

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

    if prompt:
        lines.append("**Image prompt**")
        lines.append("")
        lines.append("```")
        lines.append(prompt)
        lines.append("```")
    else:
        lines.append("_No image prompt saved._")
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

    lines.append("</details>")
    lines.append("")
    return lines
