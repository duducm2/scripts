#!/usr/bin/env python3
"""Build a local HTML tree map of click_sequences.ini (Finance-style: static HTML + CDN)."""

from __future__ import annotations

import html
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
INI_PATH = ROOT / "assets" / "data" / "click_sequences.ini"
VOCAB_PATH = ROOT / "docs" / "click-sequence-vocabulary.md"
OUT_DIR = Path(__file__).resolve().parent / "output"
OUT_HTML = OUT_DIR / "map.html"
MERMAID_CDN = "https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.min.js"


def read_text(path: Path) -> str:
    raw = path.read_bytes()
    if raw.startswith(b"\xff\xfe") or raw.startswith(b"\xfe\xff"):
        return raw.decode("utf-16")
    if raw.startswith(b"\xef\xbb\xbf"):
        return raw.decode("utf-8-sig")
    return raw.decode("utf-8")


def parse_ini(path: Path) -> dict[str, dict[str, str]]:
    sections: dict[str, dict[str, str]] = {}
    current = ""
    if not path.exists():
        return sections
    for line in read_text(path).splitlines():
        line = line.strip()
        if not line or line.startswith(";") or line.startswith("#"):
            continue
        if line.startswith("[") and line.endswith("]"):
            current = line[1:-1]
            sections.setdefault(current, {})
            continue
        if current and "=" in line:
            key, val = line.split("=", 1)
            sections[current][key.strip()] = val.strip()
    return sections


def norm(val: str | None) -> str:
    if val is None or val == "ERROR":
        return ""
    return val


def mermaid_label(text: str) -> str:
    cleaned = re.sub(r"[\[\]{}|\\]", " ", text)
    cleaned = cleaned.replace('"', "'")
    cleaned = " ".join(cleaned.split())
    if len(cleaned) > 72:
        cleaned = cleaned[:69] + "..."
    return cleaned or "(empty)"


def nid(*parts: object) -> str:
    raw = "_".join(str(p) for p in parts)
    out = re.sub(r"[^A-Za-z0-9_]", "_", raw)
    if out and out[0].isdigit():
        out = "n_" + out
    return out or "node"


def parse_slots(raw: str) -> list[tuple[str, str]]:
    slots: list[tuple[str, str]] = []
    text = norm(raw)
    if not text:
        return [
            ("hardcoded", "scrollFeedBottom"),
            ("seqGroup", "clicks"),
            ("hardcoded", "desktopWait"),
            ("hardcoded", "desktopCut"),
        ]
    for part in text.split("|"):
        part = part.strip()
        if ":" not in part:
            continue
        kind, rest = part.split(":", 1)
        kind = kind.strip().lower()
        rest = rest.strip()
        if not rest:
            continue
        if kind in ("hardcoded", "script"):
            slots.append(("hardcoded", rest))
        elif kind in ("seqgroup", "group"):
            slots.append(("seqGroup", rest))
    return slots


def parse_search_order(raw: str) -> str:
    text = norm(raw)
    order = "bottomUp"
    for part in text.split("|"):
        if ":" not in part:
            continue
        k, v = part.split(":", 1)
        if k.strip().lower() == "searchorder":
            val = v.strip().lower()
            if val in ("topdown", "top-down", "top"):
                order = "topDown"
            elif val in ("firstmatch", "first", "tree"):
                order = "firstMatch"
            else:
                order = "bottomUp"
    return order


def group_indexes(
    sections: dict[str, dict[str, str]], macro_id: str, group_id: str, seq_count: int
) -> list[int]:
    for i in range(1, 21):
        sec = sections.get(f"SeqGroup_{macro_id}_{i}", {})
        gid = norm(sec.get("Id")).lower()
        if gid == "" and not norm(sec.get("Sequences")):
            break
        if gid == "" or gid == group_id.lower():
            idxs: list[int] = []
            for tok in norm(sec.get("Sequences")).split("|"):
                tok = tok.strip()
                if tok.isdigit():
                    n = int(tok)
                    if 1 <= n <= seq_count:
                        idxs.append(n)
            if gid == group_id.lower() or (gid == "" and group_id == "clicks"):
                return idxs or list(range(1, seq_count + 1))
    if group_id == "clicks":
        return list(range(1, seq_count + 1))
    return []


def sequence_count(sections: dict[str, dict[str, str]], macro_id: str) -> int:
    n = 0
    for i in range(1, 81):
        sec = sections.get(f"Seq_{macro_id}_{i}")
        if not sec:
            break
        if not norm(sec.get("Name")) and not any(
            k.startswith("Click_")
            for k in sections
            if k.startswith(f"Click_{macro_id}_{i}_")
        ):
            if i == 1:
                continue
            break
        n = i
    if n:
        return n
    for i in range(1, 81):
        if f"Seq_{macro_id}_{i}" in sections:
            n = i
        else:
            break
    return n


def click_count(
    sections: dict[str, dict[str, str]], macro_id: str, seq_idx: int
) -> int:
    n = 0
    for i in range(1, 41):
        if f"Click_{macro_id}_{seq_idx}_{i}" in sections:
            n = i
        else:
            break
    return n


def split_aliases(raw: str) -> list[str]:
    return [a.strip() for a in norm(raw).split("|") if a.strip()]


def md_to_html(md: str) -> str:
    lines = md.replace("\r\n", "\n").split("\n")
    out: list[str] = []
    in_table = False
    in_code = False
    for line in lines:
        if line.startswith("```"):
            if in_code:
                out.append("</pre>")
                in_code = False
            else:
                out.append("<pre>")
                in_code = True
            continue
        if in_code:
            out.append(html.escape(line))
            continue
        if line.startswith("|") and "---" not in line:
            cells = [c.strip() for c in line.strip("|").split("|")]
            tag = "th" if not in_table else "td"
            if not in_table:
                out.append("<table>")
                in_table = True
            row = "".join(f"<{tag}>{html.escape(c)}</{tag}>" for c in cells)
            out.append(f"<tr>{row}</tr>")
            continue
        if in_table:
            if line.startswith("|"):
                continue
            out.append("</table>")
            in_table = False
        if not line.strip():
            continue
        if line.startswith("# "):
            out.append(f"<h1>{html.escape(line[2:].strip())}</h1>")
        elif line.startswith("## "):
            out.append(f"<h2>{html.escape(line[3:].strip())}</h2>")
        elif line.startswith("- "):
            out.append(f"<li>{inline_md(line[2:].strip())}</li>")
        else:
            out.append(f"<p>{inline_md(line)}</p>")
    if in_table:
        out.append("</table>")
    if in_code:
        out.append("</pre>")
    return "\n".join(out)


def inline_md(text: str) -> str:
    escaped = html.escape(text)
    escaped = re.sub(r"`([^`]+)`", r"<code>\1</code>", escaped)
    escaped = re.sub(r"\*\*([^*]+)\*\*", r"<strong>\1</strong>", escaped)
    return escaped


def build_macros(sections: dict[str, dict[str, str]]) -> list[dict]:
    index = sections.get("Index", {})
    ids = [p.strip() for p in norm(index.get("Macros")).split("|") if p.strip()]
    if not ids:
        ids = ["ai-quick-download"]
    macros: list[dict] = []
    for mid in ids:
        msec = sections.get(f"Macro_{mid}", {})
        seq_n = sequence_count(sections, mid)
        slots_out: list[dict] = []
        for kind, rest in parse_slots(msec.get("Slots", "")):
            if kind == "hardcoded":
                slots_out.append({"kind": "hardcoded", "id": rest, "siblings": []})
                continue
            siblings: list[dict] = []
            for seq_idx in group_indexes(sections, mid, rest, seq_n):
                ssec = sections.get(f"Seq_{mid}_{seq_idx}", {})
                clicks: list[dict] = []
                for ci in range(1, click_count(sections, mid, seq_idx) + 1):
                    csec = sections.get(f"Click_{mid}_{seq_idx}_{ci}", {})
                    aliases = split_aliases(csec.get("Selectors", ""))
                    clicks.append({"index": ci, "aliases": aliases})
                siblings.append(
                    {
                        "index": seq_idx,
                        "name": norm(ssec.get("Name")) or f"sequence {seq_idx}",
                        "context": norm(ssec.get("Context")) or "*",
                        "clicks": clicks,
                    }
                )
            slots_out.append({"kind": "seqGroup", "id": rest, "siblings": siblings})
        macros.append(
            {
                "id": mid,
                "name": norm(msec.get("Name")) or mid,
                "trigger": norm(msec.get("Trigger")),
                "searchOrder": parse_search_order(msec.get("Rules", "")),
                "slots": slots_out,
            }
        )
    return macros


def build_outline(macros: list[dict]) -> str:
    parts: list[str] = []
    for macro in macros:
        title = f"{macro['name']}  {macro['trigger']}  ({macro['searchOrder']})"
        parts.append(
            f'<details class="lvl-shortcut" open><summary>{html.escape(title)}</summary>'
        )
        for slot in macro["slots"]:
            if slot["kind"] == "hardcoded":
                parts.append(
                    f'<div class="lvl-hardcoded">Hardcoded Script: '
                    f'<code>{html.escape(slot["id"])}</code></div>'
                )
                continue
            n = len(slot["siblings"])
            slot_title = (
                f'Sequence Group: {slot["id"]}  ({n} sibling{"s" if n != 1 else ""})'
            )
            parts.append(
                f'<details class="lvl-slot" open><summary>{html.escape(slot_title)}</summary>'
            )
            for sib in slot["siblings"]:
                sib_title = f'Sibling {sib["index"]}: {sib["name"]}  [{sib["context"]}]'
                parts.append(
                    f'<details class="lvl-sibling"><summary>{html.escape(sib_title)}</summary>'
                )
                for click in sib["clicks"]:
                    n_alias = len(click["aliases"])
                    click_title = f'Click {click["index"]}  ·  {n_alias} alias{"es" if n_alias != 1 else ""}'
                    parts.append(
                        f'<details class="lvl-click"><summary>{html.escape(click_title)}</summary>'
                    )
                    parts.append('<ul class="aliases">')
                    if click["aliases"]:
                        for alias in click["aliases"]:
                            parts.append(f"<li><code>{html.escape(alias)}</code></li>")
                    else:
                        parts.append("<li>(no aliases)</li>")
                    parts.append("</ul></details>")
                parts.append("</details>")
            parts.append("</details>")
        parts.append("</details>")
    return "\n".join(parts)


def build_mermaid(macros: list[dict]) -> str:
    lines = ["flowchart TB"]
    for macro in macros:
        mnode = nid("M", macro["id"])
        label = mermaid_label(
            f"{macro['name']} {macro['trigger']} ({macro['searchOrder']})"
        )
        lines.append(f'  {mnode}["{label}"]')
        for si, slot in enumerate(macro["slots"], start=1):
            snode = nid("S", macro["id"], si)
            if slot["kind"] == "hardcoded":
                lines.append(
                    f'  {snode}["{mermaid_label("Hardcoded: " + slot["id"])}"]'
                )
                lines.append(f"  {mnode} --> {snode}")
                continue
            lines.append(
                f'  {snode}["{mermaid_label("Sequence Group: " + slot["id"])}"]'
            )
            lines.append(f"  {mnode} --> {snode}")
            for sib in slot["siblings"]:
                qnode = nid("Q", macro["id"], sib["index"])
                sib_label = mermaid_label(
                    f"Sibling {sib['index']}: {sib['name']} [{sib['context']}]"
                )
                lines.append(f'  {qnode}["{sib_label}"]')
                lines.append(f"  {snode} --> {qnode}")
                for click in sib["clicks"]:
                    n_alias = len(click["aliases"])
                    alias_word = "alias" if n_alias == 1 else "aliases"
                    cnode = nid("C", macro["id"], sib["index"], click["index"])
                    click_label = mermaid_label(
                        f"Click {click['index']} · {n_alias} {alias_word}"
                    )
                    lines.append(f'  {cnode}["{click_label}"]')
                    lines.append(f"  {qnode} --> {cnode}")
    return "\n".join(lines)


def main() -> int:
    sections = parse_ini(INI_PATH)
    macros = build_macros(sections)
    mermaid = build_mermaid(macros)
    outline = build_outline(macros)
    vocab = ""
    if VOCAB_PATH.exists():
        vocab = md_to_html(read_text(VOCAB_PATH))
    page = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>Click Sequence map</title>
  <script src="{MERMAID_CDN}"></script>
  <style>
    :root {{
      --bg: #1e1e1e; --header: #252526; --panel: #2c2c2c; --text: #ddd;
      --muted: #9aa0a6; --border: #3c3c3c; --accent: #4fc3f7; --btn: #3c3c3c;
    }}
    body {{ margin: 0; font-family: Segoe UI, sans-serif; background: var(--bg); color: var(--text); }}
    header {{ padding: 10px 16px; background: var(--header); border-bottom: 1px solid var(--border);
      display: flex; align-items: flex-start; justify-content: space-between; gap: 16px; flex-wrap: wrap; }}
    header h1 {{ margin: 0; font-size: 16px; }}
    header p {{ margin: 4px 0 0; color: var(--muted); font-size: 12px; }}
    .header-actions {{ display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }}
    button, .toggle {{ background: var(--btn); color: var(--text); border: 1px solid var(--border);
      padding: 6px 10px; border-radius: 4px; cursor: pointer; font-size: 12px; }}
    button:hover, .toggle:hover {{ border-color: var(--accent); }}
    .toggle.active {{ background: #1b4d66; border-color: var(--accent); color: #fff; }}
    .layout {{ display: grid; grid-template-columns: minmax(0, 1.6fr) minmax(280px, 0.8fr);
      min-height: calc(100vh - 62px); }}
    .main, .vocab {{ padding: 16px; overflow: auto; }}
    .vocab {{ border-left: 1px solid var(--border); background: var(--panel); }}
    .vocab h1 {{ font-size: 18px; }}
    .vocab h2 {{ font-size: 14px; color: var(--accent); margin-top: 18px; }}
    .vocab p, .vocab li {{ font-size: 13px; line-height: 1.45; }}
    .vocab table {{ border-collapse: collapse; width: 100%; font-size: 12px; margin: 8px 0; }}
    .vocab th, .vocab td {{ border: 1px solid var(--border); padding: 4px 6px; text-align: left; }}
    .vocab code, .main code {{ background: #1e1e1e; padding: 1px 4px; }}
    #view-tree {{ display: none; }}
    #view-tree.visible {{ display: block; }}
    #view-outline.hidden {{ display: none; }}
    #outline details {{ margin: 4px 0 4px 12px; }}
    #outline > details {{ margin-left: 0; }}
    #outline summary {{ cursor: pointer; font-size: 13px; line-height: 1.5; }}
    #outline .lvl-hardcoded {{ margin: 4px 0 4px 24px; font-size: 13px; color: var(--muted); }}
    #outline ul.aliases {{ margin: 4px 0 8px 18px; padding: 0; list-style: disc; font-size: 13px; }}
    .outline-tools {{ margin-bottom: 10px; display: flex; gap: 8px; }}
    .tree-hint {{ color: var(--muted); font-size: 12px; margin: 0 0 8px; }}
    .tree-pan {{ background: var(--panel); border-radius: 6px; overflow: hidden; height: calc(100vh - 140px);
      cursor: grab; position: relative; }}
    .tree-pan.dragging {{ cursor: grabbing; }}
    .tree-inner {{ transform-origin: 0 0; padding: 16px; width: max-content; }}
    .mermaid {{ background: transparent; }}
  </style>
</head>
<body>
  <header>
    <div>
      <h1>Click Sequence map</h1>
      <p>Shortcut → Slots → Sibling Sequences → Clicks → Aliases. Generated locally from click_sequences.ini.</p>
    </div>
    <div class="header-actions">
      <button type="button" class="toggle active" id="btn-outline">Outline</button>
      <button type="button" class="toggle" id="btn-tree">Tree</button>
    </div>
  </header>
  <div class="layout">
    <div class="main">
      <div id="view-outline">
        <div class="outline-tools">
          <button type="button" id="btn-expand">Expand all</button>
          <button type="button" id="btn-collapse">Collapse all</button>
        </div>
        <div id="outline">
{outline}
        </div>
      </div>
      <div id="view-tree">
        <p class="tree-hint">Wheel to zoom, drag to pan. Aliases are listed in Outline.</p>
        <div class="tree-pan" id="tree-pan">
          <div class="tree-inner" id="tree-inner">
            <pre class="mermaid" id="mermaid-src">
{mermaid}
            </pre>
          </div>
        </div>
      </div>
    </div>
    <aside class="vocab">
{vocab}
    </aside>
  </div>
  <script>
    const outlineView = document.getElementById("view-outline");
    const treeView = document.getElementById("view-tree");
    const btnOutline = document.getElementById("btn-outline");
    const btnTree = document.getElementById("btn-tree");
    let mermaidReady = false;

    mermaid.initialize({{ startOnLoad: false, theme: "dark", flowchart: {{ useMaxWidth: false, htmlLabels: true }} }});

    function showOutline() {{
      outlineView.classList.remove("hidden");
      treeView.classList.remove("visible");
      btnOutline.classList.add("active");
      btnTree.classList.remove("active");
    }}
    async function showTree() {{
      outlineView.classList.add("hidden");
      treeView.classList.add("visible");
      btnTree.classList.add("active");
      btnOutline.classList.remove("active");
      if (!mermaidReady) {{
        await mermaid.run({{ querySelector: "#mermaid-src" }});
        mermaidReady = true;
      }}
    }}
    btnOutline.addEventListener("click", showOutline);
    btnTree.addEventListener("click", showTree);

    document.getElementById("btn-expand").addEventListener("click", () => {{
      document.querySelectorAll("#outline details").forEach((d) => {{ d.open = true; }});
    }});
    document.getElementById("btn-collapse").addEventListener("click", () => {{
      document.querySelectorAll("#outline details").forEach((d) => {{ d.open = false; }});
    }});

    const pan = document.getElementById("tree-pan");
    const inner = document.getElementById("tree-inner");
    let scale = 1, x = 16, y = 16, dragging = false, lastX = 0, lastY = 0;
    function applyTransform() {{
      inner.style.transform = `translate(${{x}}px, ${{y}}px) scale(${{scale}})`;
    }}
    applyTransform();
    pan.addEventListener("pointerdown", (e) => {{
      dragging = true;
      lastX = e.clientX;
      lastY = e.clientY;
      pan.classList.add("dragging");
      pan.setPointerCapture(e.pointerId);
    }});
    pan.addEventListener("pointermove", (e) => {{
      if (!dragging) return;
      x += e.clientX - lastX;
      y += e.clientY - lastY;
      lastX = e.clientX;
      lastY = e.clientY;
      applyTransform();
    }});
    function endDrag(e) {{
      dragging = false;
      pan.classList.remove("dragging");
      try {{ pan.releasePointerCapture(e.pointerId); }} catch (_) {{}}
    }}
    pan.addEventListener("pointerup", endDrag);
    pan.addEventListener("pointercancel", endDrag);
    pan.addEventListener("wheel", (e) => {{
      e.preventDefault();
      const next = Math.min(2.5, Math.max(0.25, scale * (e.deltaY < 0 ? 1.1 : 0.9)));
      scale = next;
      applyTransform();
    }}, {{ passive: false }});
  </script>
</body>
</html>
"""
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    OUT_HTML.write_text(page, encoding="utf-8")
    print(OUT_HTML)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
