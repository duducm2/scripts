"""Generate Memory Palace HTML dashboard (study → street images)."""
from __future__ import annotations

import argparse
import html
import sys
from pathlib import Path
from urllib.parse import quote

sys.path.insert(0, str(Path(__file__).resolve().parent))

from data_aggregator import snapshot  # noqa: E402


def file_uri(path: str) -> str:
    p = Path(path).resolve()
    # Chrome file URI on Windows
    return "file:///" + quote(p.as_posix(), safe="/:")


def build_html(snap: dict) -> str:
    studies = snap.get("all_studies") or []
    cards = snap.get("studies") or []
    selected = snap.get("selected_study_id") or ""
    totals = snap.get("totals") or {}
    issues = snap.get("issues") or []

    options = []
    for s in studies:
        sel = " selected" if s["id"] == selected else ""
        options.append(
            f'<option value="{html.escape(s["id"])}"{sel}>{html.escape(s["title"])}</option>'
        )

    # JS study data embedded
    study_blocks = []
    for study in cards:
        street_html = []
        for st in study.get("streets", []):
            if st.get("image_exists") and st.get("image_abs"):
                img = (
                    f'<img src="{html.escape(file_uri(st["image_abs"]))}" '
                    f'alt="Street {html.escape(str(st["number"]))}" loading="lazy"/>'
                )
            else:
                img = '<div class="missing">No image</div>'
            street_html.append(
                f"""
                <article class="street">
                  {img}
                  <div class="meta">
                    <h3>Street {html.escape(str(st["number"]))}: {html.escape(st["title"])}</h3>
                    <p class="char">{html.escape(st.get("character") or "—")}</p>
                    <p class="counts">{st.get("beast_count", 0)} beasts · {st.get("atom_count", 0)} atoms</p>
                  </div>
                </article>
                """
            )
        study_blocks.append(
            f"""
            <section class="study" data-study-id="{html.escape(study["id"])}"
                     style="display:{'block' if study['id'] == selected or (not selected and cards and study is cards[0]) else 'none'}">
              <header class="study-head">
                <h2>{html.escape(study["title"])}</h2>
                <p>{study.get("street_count", 0)} streets · slug {html.escape(study.get("slug", ""))}</p>
              </header>
              <div class="grid">
                {''.join(street_html) if street_html else '<p class="empty">No streets yet.</p>'}
              </div>
            </section>
            """
        )

    issues_html = ""
    if issues:
        items = "".join(f"<li>{html.escape(i)}</li>" for i in issues[:20])
        issues_html = f'<aside class="issues"><h3>Validation</h3><ul>{items}</ul></aside>'

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1"/>
  <title>Memory Palace</title>
  <style>
    :root {{
      --bg: #12141a;
      --panel: #1c1f28;
      --text: #e8e6e0;
      --muted: #9a958c;
      --gold: #d4a017;
      --line: #2a2e3a;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: "Segoe UI", "Iowan Old Style", Georgia, serif;
      background: radial-gradient(1200px 600px at 10% -10%, #2a2418 0%, var(--bg) 55%);
      color: var(--text);
      min-height: 100vh;
    }}
    header.top {{
      padding: 1.25rem 1.5rem 0.75rem;
      border-bottom: 1px solid var(--line);
      display: flex;
      flex-wrap: wrap;
      gap: 1rem;
      align-items: end;
      justify-content: space-between;
    }}
    header.top h1 {{
      margin: 0;
      font-size: 1.6rem;
      font-weight: 650;
      letter-spacing: 0.02em;
    }}
    header.top .sub {{ color: var(--muted); margin: 0.25rem 0 0; font-size: 0.95rem; }}
    label {{ color: var(--muted); font-size: 0.85rem; display: block; margin-bottom: 0.35rem; }}
    select {{
      background: var(--panel);
      color: var(--text);
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 0.45rem 0.75rem;
      min-width: 220px;
      font-size: 1rem;
    }}
    main {{ padding: 1.25rem 1.5rem 2.5rem; }}
    .study-head h2 {{ margin: 0 0 0.25rem; color: var(--gold); }}
    .study-head p {{ margin: 0 0 1rem; color: var(--muted); }}
    .grid {{
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(280px, 1fr));
      gap: 1rem;
    }}
    .street {{
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 10px;
      overflow: hidden;
      display: flex;
      flex-direction: column;
    }}
    .street img {{
      width: 100%;
      aspect-ratio: 4/3;
      object-fit: cover;
      background: #0a0b0e;
      display: block;
    }}
    .missing {{
      width: 100%;
      aspect-ratio: 4/3;
      display: grid;
      place-items: center;
      background: #0a0b0e;
      color: var(--muted);
      border-bottom: 1px solid var(--line);
    }}
    .meta {{ padding: 0.75rem 0.9rem 1rem; }}
    .meta h3 {{ margin: 0 0 0.35rem; font-size: 1.05rem; }}
    .char {{ margin: 0; color: var(--gold); font-size: 0.92rem; }}
    .counts {{ margin: 0.35rem 0 0; color: var(--muted); font-size: 0.85rem; }}
    .empty {{ color: var(--muted); }}
    .issues {{
      margin-top: 1.5rem;
      padding: 1rem;
      background: #2a1c1c;
      border: 1px solid #5a3030;
      border-radius: 8px;
    }}
    .issues h3 {{ margin: 0 0 0.5rem; color: #f0a0a0; }}
    .issues ul {{ margin: 0; padding-left: 1.2rem; color: #e0c0c0; }}
    footer {{
      padding: 0 1.5rem 1.5rem;
      color: var(--muted);
      font-size: 0.85rem;
    }}
  </style>
</head>
<body>
  <header class="top">
    <div>
      <h1>Memory Palace</h1>
      <p class="sub">{totals.get("studies", 0)} studies · {totals.get("streets", 0)} streets ·
         {totals.get("beasts", 0)} beasts · {totals.get("atoms", 0)} atoms</p>
    </div>
    <div>
      <label for="study">Study</label>
      <select id="study">
        {''.join(options) if options else '<option value="">No studies</option>'}
      </select>
    </div>
  </header>
  <main>
    {''.join(study_blocks)}
    {issues_html}
  </main>
  <footer>Street images resolve from notes/studies.</footer>
  <script>
    const sel = document.getElementById('study');
    const sections = [...document.querySelectorAll('.study')];
    function show(id) {{
      sections.forEach(s => {{
        s.style.display = (s.dataset.studyId === id) ? 'block' : 'none';
      }});
    }}
    if (sel) {{
      sel.addEventListener('change', () => show(sel.value));
      if (sel.value) show(sel.value);
    }}
  </script>
</body>
</html>
"""


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="Build Memory Palace dashboard HTML")
    p.add_argument("--data-dir", type=Path, required=True)
    p.add_argument("--output-dir", type=Path, required=True)
    p.add_argument("--notes-root", type=Path, default=None)
    p.add_argument("--study-id", type=str, default=None)
    args = p.parse_args(argv)

    # Always load every study for the picker; study_id only sets the default tab.
    snap = snapshot(
        data_dir=args.data_dir.resolve(),
        notes_root=args.notes_root.resolve() if args.notes_root else None,
        study_id=None,
    )
    if args.study_id:
        snap["selected_study_id"] = args.study_id
    elif snap.get("studies"):
        snap["selected_study_id"] = snap["studies"][0]["id"]

    html_out = build_html(snap)
    out_dir = args.output_dir.resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    out = out_dir / "dashboard.html"
    out.write_text(html_out, encoding="utf-8")
    print(f"Wrote {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
