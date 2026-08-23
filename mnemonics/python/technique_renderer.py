"""Render mirrored technique docs into dashboard Method HTML + canon JSON."""

from __future__ import annotations

import html
import json
import re
from pathlib import Path
from typing import Any

try:
    import markdown
    from markdown.extensions.fenced_code import FencedCodeExtension
    from markdown.extensions.tables import TableExtension
    from markdown.extensions.toc import TocExtension
except ImportError:  # pragma: no cover
    markdown = None  # type: ignore


def _md_to_html(text: str) -> tuple[str, str]:
    if markdown is None:
        body = (
            '<pre class="method-fallback">'
            + html.escape(text)
            + "</pre>"
            + '<p class="method-warn">Install <code>markdown</code> '
            "(pip install markdown&gt;=3.5) for formatted Method view.</p>"
        )
        return body, ""

    # Convert mermaid fences to <pre class="mermaid"> so mermaid.js can render
    def _mermaid_fence(m: re.Match[str]) -> str:
        body = m.group(1).strip("\n")
        return f'<pre class="mermaid">\n{html.escape(body)}\n</pre>'

    text = re.sub(
        r"```mermaid\s*\n(.*?)```",
        _mermaid_fence,
        text,
        flags=re.S | re.I,
    )
    md = markdown.Markdown(
        extensions=[
            TableExtension(),
            FencedCodeExtension(),
            TocExtension(permalink=False, toc_depth="2-3"),
        ]
    )
    body = md.convert(text)
    toc = getattr(md, "toc", "") or ""
    return body, toc


def _load_json(path: Path) -> Any:
    if not path.exists():
        return None
    try:
        text = path.read_text(encoding="utf-8")
        data, _ = json.JSONDecoder().raw_decode(text.lstrip())
        return data
    except (json.JSONDecodeError, OSError, ValueError):
        return None


def _prompt_preview(path: Path, lines: int = 20) -> tuple[str, str, bool]:
    """Return (name, preview_html_escaped, truncated)."""
    text = path.read_text(encoding="utf-8", errors="replace")
    all_lines = text.splitlines()
    truncated = len(all_lines) > lines
    preview = "\n".join(all_lines[:lines])
    if truncated:
        preview += "\n…"
    return path.name, html.escape(preview), truncated


def _research_details(research_dir: Path) -> str:
    if not research_dir.is_dir():
        return ""
    blocks: list[str] = []
    for path in sorted(research_dir.glob("*.md")):
        raw = path.read_text(encoding="utf-8", errors="replace")
        body, _toc = _md_to_html(raw)
        title = path.stem.replace("-", " ").title()
        blocks.append(
            f'<details class="method-research">'
            f"<summary>{html.escape(title)}</summary>"
            f'<div class="method-prose">{body}</div>'
            f"</details>"
        )
    if not blocks:
        return ""
    return (
        '<section class="method-section" id="method-research">'
        "<h2>Research notes</h2>" + "".join(blocks) + "</section>"
    )


def _prompts_section(prompts_dir: Path) -> str:
    if not prompts_dir.is_dir():
        return ""
    cards: list[str] = []
    for path in sorted(prompts_dir.glob("*.txt")):
        name, preview, truncated = _prompt_preview(path)
        full = html.escape(path.read_text(encoding="utf-8", errors="replace"))
        toggle = ""
        if truncated:
            toggle = (
                f'<button type="button" class="btn-prompt-full" '
                f'data-prompt="{html.escape(name)}">Show all</button>'
            )
        cards.append(
            f'<details class="method-prompt" data-prompt-name="{html.escape(name)}">'
            f"<summary>{html.escape(name)}</summary>"
            f'<pre class="prompt-preview" data-preview="1">{preview}</pre>'
            f'<pre class="prompt-full" hidden>{full}</pre>'
            f"{toggle}"
            f"</details>"
        )
    if not cards:
        return ""
    return (
        '<section class="method-section" id="method-prompts">'
        "<h2>Workflow prompts</h2>"
        '<p class="method-note">Mirrored from technique/prompts. Expand a file to preview.</p>'
        + "".join(cards)
        + "</section>"
    )


def build_canon_payload(technique_dir: Path) -> dict[str, Any]:
    characters = _load_json(technique_dir / "characters.json") or {}
    bestiary = _load_json(technique_dir / "bestiary.json") or {}

    char_sections: list[dict[str, Any]] = []
    for sec in characters.get("sections") or []:
        title = re.sub(r"\*+", "", str(sec.get("title") or "")).strip()
        names = [
            {"number": c.get("number", ""), "name": c.get("name", "")}
            for c in (sec.get("characters") or [])
            if c.get("name")
        ]
        if names:
            char_sections.append({"title": title, "characters": names})

    beasts: list[dict[str, Any]] = []
    for item in bestiary.get("items") or []:
        beasts.append(
            {
                "code": item.get("code", ""),
                "name": item.get("name", ""),
                "source": item.get("source", ""),
                "order": item.get("order", ""),
            }
        )

    return {
        "characters": char_sections,
        "bestiary": beasts,
    }


def build_method_panel(technique_dir: Path) -> tuple[str, dict[str, Any]]:
    """Return (method_panel_html, canon_dict)."""
    technique_dir = technique_dir.resolve()
    if not technique_dir.is_dir():
        empty = (
            '<div class="method-empty">'
            "<p>Technique mirror not found. Run "
            "<code>sync_technique.py</code> or open Dashboard after notes sync.</p>"
            "</div>"
        )
        return empty, {"characters": [], "bestiary": []}

    readme_path = technique_dir / "README.md"
    readme_raw = (
        readme_path.read_text(encoding="utf-8", errors="replace")
        if readme_path.exists()
        else "_README.md missing from technique mirror._"
    )
    readme_html, toc_html = _md_to_html(readme_raw)

    research_html = _research_details(technique_dir / "research")
    prompts_html = _prompts_section(technique_dir / "prompts")
    canon = build_canon_payload(technique_dir)

    toc_block = ""
    if toc_html.strip():
        toc_block = (
            f'<nav class="method-toc" aria-label="Method contents">'
            f"<h2>On this page</h2>{toc_html}</nav>"
        )

    panel = f"""
<div class="method-layout">
  {toc_block}
  <div class="method-main">
    <p class="method-banner">Software uses <strong>Memory Palace</strong>; method doc may say <strong>Street</strong>.</p>
    <article class="method-prose" id="method-readme">
      {readme_html}
    </article>
    {research_html}
    {prompts_html}
    <section class="method-section" id="method-characters">
      <h2>Characters (canon)</h2>
      <label class="method-search-label" for="charSearch">Search characters</label>
      <input type="search" id="charSearch" placeholder="Name…" autocomplete="off"/>
      <div id="charResults" class="canon-grid"></div>
    </section>
    <section class="method-section" id="method-bestiary">
      <h2>Bestiary (canon)</h2>
      <label class="method-search-label" for="beastSearch">Search beasts</label>
      <input type="search" id="beastSearch" placeholder="Code or name…" autocomplete="off"/>
      <div id="beastResults" class="canon-table-wrap"></div>
    </section>
  </div>
</div>
"""
    return panel, canon


def default_technique_dir(script_parent: Path | None = None) -> Path:
    """mnemonics/technique relative to python/ or scripts root."""
    here = script_parent or Path(__file__).resolve().parent
    # mnemonics/python -> mnemonics/technique
    return (here.parent / "technique").resolve()
