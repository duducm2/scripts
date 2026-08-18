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


def group_indexes(sections: dict[str, dict[str, str]], macro_id: str, group_id: str, seq_count: int) -> list[int]:
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
            k.startswith("Click_") for k in sections if k.startswith(f"Click_{macro_id}_{i}_")
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


def click_count(sections: dict[str, dict[str, str]], macro_id: str, seq_idx: int) -> int:
    n = 0
    for i in range(1, 41):
        if f"Click_{macro_id}_{seq_idx}_{i}" in sections:
            n = i
        else:
            break
    return n


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


def build_mermaid(sections: dict[str, dict[str, str]]) -> str:
    lines = ["flowchart TD"]
    index = sections.get("Index", {})
    ids = [p.strip() for p in norm(index.get("Macros")).split("|") if p.strip()]
    if not ids:
        ids = ["ai-quick-download"]
    for mid in ids:
        msec = sections.get(f"Macro_{mid}", {})
        name = norm(msec.get("Name")) or mid
        trigger = norm(msec.get("Trigger"))
        order = parse_search_order(msec.get("Rules", ""))
        mnode = nid("M", mid)
        lines.append(f'  {mnode}["{mermaid_label(f"{name} {trigger} ({order})") }"]')
        slots = parse_slots(msec.get("Slots", ""))
        seq_n = sequence_count(sections, mid)
        for si, (kind, rest) in enumerate(slots, start=1):
            snode = nid("S", mid, si)
            if kind == "hardcoded":
                lines.append(f'  {snode}["{mermaid_label("Hardcoded: " + rest)}"]')
                lines.append(f"  {mnode} --> {snode}")
                continue
            lines.append(f'  {snode}["{mermaid_label("Sequence Group: " + rest)}"]')
            lines.append(f"  {mnode} --> {snode}")
            for seq_idx in group_indexes(sections, mid, rest, seq_n):
                ssec = sections.get(f"Seq_{mid}_{seq_idx}", {})
                sname = norm(ssec.get("Name")) or f"sequence {seq_idx}"
                ctx = norm(ssec.get("Context")) or "*"
                qnode = nid("Q", mid, seq_idx)
                lines.append(f'  {qnode}["{mermaid_label(f"Sibling {seq_idx}: {sname} [{ctx}]")}"]')
                lines.append(f"  {snode} --> {qnode}")
                for ci in range(1, click_count(sections, mid, seq_idx) + 1):
                    csec = sections.get(f"Click_{mid}_{seq_idx}_{ci}", {})
                    preview = norm(csec.get("Selectors"))
                    cnode = nid("C", mid, seq_idx, ci)
                    lines.append(f'  {cnode}["{mermaid_label(f"Click {ci}: {preview}")}"]')
                    lines.append(f"  {qnode} --> {cnode}")
                    aliases = [a.strip() for a in preview.split("|") if a.strip()]
                    for ai, alias in enumerate(aliases, start=1):
                        anode = nid("A", mid, seq_idx, ci, ai)
                        lines.append(f'  {anode}["{mermaid_label("Alias: " + alias)}"]')
                        lines.append(f"  {cnode} --> {anode}")
    return "\n".join(lines)


def main() -> int:
    sections = parse_ini(INI_PATH)
    mermaid = build_mermaid(sections)
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
      --muted: #9aa0a6; --border: #3c3c3c; --accent: #4fc3f7;
    }}
    body {{ margin: 0; font-family: Segoe UI, sans-serif; background: var(--bg); color: var(--text); }}
    header {{ padding: 10px 16px; background: var(--header); border-bottom: 1px solid var(--border); }}
    header h1 {{ margin: 0; font-size: 16px; }}
    header p {{ margin: 4px 0 0; color: var(--muted); font-size: 12px; }}
    .layout {{ display: grid; grid-template-columns: minmax(0, 1.6fr) minmax(280px, 0.8fr); min-height: calc(100vh - 58px); }}
    .tree, .vocab {{ padding: 16px; overflow: auto; }}
    .vocab {{ border-left: 1px solid var(--border); background: var(--panel); }}
    .vocab h1 {{ font-size: 18px; }}
    .vocab h2 {{ font-size: 14px; color: var(--accent); margin-top: 18px; }}
    .vocab p, .vocab li {{ font-size: 13px; line-height: 1.45; }}
    .vocab table {{ border-collapse: collapse; width: 100%; font-size: 12px; margin: 8px 0; }}
    .vocab th, .vocab td {{ border: 1px solid var(--border); padding: 4px 6px; text-align: left; }}
    .vocab code {{ background: #1e1e1e; padding: 1px 4px; }}
    .mermaid {{ background: var(--panel); padding: 12px; border-radius: 6px; }}
  </style>
</head>
<body>
  <header>
    <h1>Click Sequence map</h1>
    <p>Shortcut → Slots → Sibling Sequences → Clicks → Aliases. Generated locally from click_sequences.ini.</p>
  </header>
  <div class="layout">
    <div class="tree">
      <pre class="mermaid">
{mermaid}
      </pre>
    </div>
    <aside class="vocab">
{vocab}
    </aside>
  </div>
  <script>mermaid.initialize({{ startOnLoad: true, theme: "dark" }});</script>
</body>
</html>
"""
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    OUT_HTML.write_text(page, encoding="utf-8")
    print(OUT_HTML)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
