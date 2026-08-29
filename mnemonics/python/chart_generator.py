"""Generate Memory Palace HTML dashboard (study → palace images + practice)."""

from __future__ import annotations

import argparse
import html
import json
import sys
from pathlib import Path
from urllib.parse import quote

sys.path.insert(0, str(Path(__file__).resolve().parent))

from data_aggregator import snapshot  # noqa: E402
from technique_renderer import (  # noqa: E402
    build_method_panel,
    default_technique_dir,
)
from sync_technique import resolve_technique_source, sync_technique  # noqa: E402
from study_plan_renderer import (  # noqa: E402
    PLANS_GITHUB_URL,
    build_plans_panel_shell,
    build_plans_payload,
)
from study_plan_parser import default_studies_root  # noqa: E402
from study_plans_md import sync_all as sync_plans_all  # noqa: E402


def file_uri(path: str) -> str:
    p = Path(path).resolve()
    return "file:///" + quote(p.as_posix(), safe="/:")


def build_html(
    snap: dict,
    method_html: str = "",
    method_canon: dict | None = None,
    plans_payload: dict | None = None,
) -> str:
    studies = snap.get("all_studies") or []
    cards = snap.get("studies") or []
    selected = snap.get("selected_study_id") or ""
    totals = snap.get("totals") or {}

    options = []
    for s in studies:
        sel = " selected" if s["id"] == selected else ""
        options.append(
            f'<option value="{html.escape(s["id"])}"{sel}>{html.escape(s["title"])}</option>'
        )

    palace_data: dict[str, dict] = {}
    study_latest: dict[str, str] = {}
    study_blocks = []
    for study in cards:
        palace_html = []
        latest_id = ""
        for st in study.get("palaces", []):
            sid = st["id"]
            if not latest_id:
                latest_id = sid  # palaces already descending by number
            img_abs = st.get("image_abs") or ""
            prompt = st.get("image_prompt") or ""
            gallery = []
            for gi in st.get("gallery_images") or []:
                gi_abs = gi.get("image_abs") or ""
                gallery.append(
                    {
                        "id": gi.get("id", ""),
                        "caption": gi.get("caption", ""),
                        "sort_order": gi.get("sort_order", ""),
                        "image_rel_path": gi.get("image_rel_path", ""),
                        "image_uri": (
                            file_uri(gi_abs)
                            if gi.get("image_exists") and gi_abs
                            else ""
                        ),
                    }
                )
            palace_data[sid] = {
                "id": sid,
                "study_id": study["id"],
                "number": st.get("number", ""),
                "title": st.get("title", ""),
                "character": st.get("character", ""),
                "atom_count": st.get("atom_count", 0),
                "image_prompt": prompt,
                "image_uri": (
                    file_uri(img_abs) if st.get("image_exists") and img_abs else ""
                ),
                "palace_notes": st.get("palace_notes") or "",
                "gallery_images": gallery,
                "atoms": st.get("atoms") or [],
            }
            if st.get("image_exists") and img_abs:
                img = (
                    f'<img src="{html.escape(file_uri(img_abs))}" '
                    f'alt="Memory Palace {html.escape(str(st["number"]))}" loading="lazy"/>'
                )
            else:
                img = '<div class="missing">No image</div>'
            if prompt.strip():
                prompt_block = (
                    f'<div class="prompt-block">'
                    f'<div class="prompt-head"><span>Image prompt</span>'
                    f'<button type="button" class="btn-copy" data-copy-prompt="{html.escape(sid)}">Copy</button>'
                    f"</div>"
                    f'<pre class="prompt-text">{html.escape(prompt)}</pre>'
                    f"</div>"
                )
            else:
                prompt_block = '<p class="prompt-empty">No image prompt saved.</p>'
            palace_html.append(
                f"""
                <article class="palace" data-palace-id="{html.escape(sid)}" tabindex="0" role="button">
                  {img}
                  <div class="meta">
                    <h3>Memory Palace {html.escape(str(st["number"]))}: {html.escape(st["title"])}</h3>
                    <p class="char">{html.escape(st.get("character") or "—")}</p>
                    <p class="counts">{st.get("beast_count", 0)} beasts · {st.get("atom_count", 0)} Knowledge Atoms</p>
                    {prompt_block}
                  </div>
                </article>
                """
            )
        if latest_id:
            study_latest[study["id"]] = latest_id
        study_blocks.append(
            f"""
            <section class="study" data-study-id="{html.escape(study["id"])}"
                     style="display:{'block' if study['id'] == selected or (not selected and cards and study is cards[0]) else 'none'}">
              <header class="study-head">
                <h2>{html.escape(study["title"])}</h2>
                <p>{study.get("palace_count", 0)} Memory Palaces · newest first</p>
              </header>
              <div class="grid">
                {''.join(palace_html) if palace_html else '<p class="empty">No Memory Palaces yet.</p>'}
              </div>
            </section>
            """
        )

    palace_json = json.dumps(palace_data, ensure_ascii=False)
    latest_json = json.dumps(study_latest, ensure_ascii=False)
    studies_json = json.dumps(
        [{"id": s["id"], "title": s.get("title") or s["id"]} for s in studies],
        ensure_ascii=False,
    )
    method_canon_json = json.dumps(
        method_canon or {"characters": [], "bestiary": []}, ensure_ascii=False
    )
    plans_data_json = json.dumps(
        plans_payload or {"plans": {}, "github_url": PLANS_GITHUB_URL},
        ensure_ascii=False,
    )
    method_body = (
        method_html
        or '<div class="method-empty"><p>No technique docs mirrored.</p></div>'
    )
    plans_body = build_plans_panel_shell()

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
    .controls {{
      display: flex;
      flex-wrap: wrap;
      gap: 0.75rem 1rem;
      align-items: end;
    }}
    label {{ color: var(--muted); font-size: 0.85rem; display: block; margin-bottom: 0.35rem; }}
    select, input[type="search"] {{
      background: var(--panel);
      color: var(--text);
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 0.45rem 0.75rem;
      min-width: 200px;
      font-size: 1rem;
    }}
    input[type="search"] {{ min-width: 260px; }}
    #btnLatest {{
      background: var(--gold);
      color: #1a1408;
      border: none;
      border-radius: 6px;
      padding: 0.5rem 0.9rem;
      font-size: 0.95rem;
      font-weight: 650;
      cursor: pointer;
      height: 2.35rem;
    }}
    #btnLatest:hover {{ filter: brightness(1.08); }}
    #btnLatest:focus {{
      outline: 2px solid #fff4d0;
      outline-offset: 2px;
      filter: brightness(1.1);
    }}
    #searchResults {{
      display: none;
      margin: 0 1.5rem 0.5rem;
      padding: 0.75rem 1rem;
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      max-height: 240px;
      overflow: auto;
    }}
    #searchResults.open {{ display: block; }}
    #searchResults .hit {{
      display: block;
      width: 100%;
      text-align: left;
      background: transparent;
      border: none;
      border-bottom: 1px solid var(--line);
      color: var(--text);
      padding: 0.55rem 0.25rem;
      cursor: pointer;
      font: inherit;
    }}
    #searchResults .hit:last-child {{ border-bottom: none; }}
    #searchResults .hit:hover {{ color: var(--gold); }}
    #searchResults .hit .meta {{ color: var(--muted); font-size: 0.85rem; }}
    #searchResults .empty {{ color: var(--muted); margin: 0; }}
    #studyModal {{
      display: none;
      position: fixed;
      inset: 0;
      z-index: 2000;
      background: rgba(8, 9, 12, 0.82);
      align-items: center;
      justify-content: center;
      padding: 1.5rem;
    }}
    #studyModal.open {{ display: flex; }}
    #studyModal .modal-panel {{
      width: min(420px, 100%);
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 12px;
      padding: 1.15rem 1.25rem 1.25rem;
      box-shadow: 0 16px 48px rgba(0, 0, 0, 0.45);
    }}
    #studyModal h2 {{
      margin: 0 0 0.35rem;
      font-size: 1.2rem;
      font-weight: 650;
    }}
    #studyModal .hint {{
      margin: 0 0 1rem;
      color: var(--muted);
      font-size: 0.9rem;
    }}
    #studyModal .pick-list {{
      list-style: none;
      margin: 0;
      padding: 0;
      display: flex;
      flex-direction: column;
      gap: 0.35rem;
      max-height: min(60vh, 420px);
      overflow: auto;
    }}
    #studyModal .pick-list button {{
      display: flex;
      align-items: center;
      gap: 0.75rem;
      width: 100%;
      text-align: left;
      background: #242830;
      border: 1px solid var(--line);
      border-radius: 8px;
      color: var(--text);
      padding: 0.65rem 0.8rem;
      font: inherit;
      cursor: pointer;
    }}
    #studyModal .pick-list button:hover,
    #studyModal .pick-list button:focus {{
      outline: none;
      border-color: var(--gold);
      color: var(--gold);
    }}
    #studyModal .key {{
      flex: 0 0 1.6rem;
      font-weight: 700;
      color: var(--gold);
      font-variant-numeric: tabular-nums;
    }}
    #studyModal .pick-label {{ flex: 1; }}
    #studyModal .modal-foot {{
      margin: 0.9rem 0 0;
      color: var(--muted);
      font-size: 0.82rem;
    }}
    main {{ padding: 1.25rem 1.5rem 2.5rem; }}
    .study-head h2 {{ margin: 0 0 0.25rem; color: var(--gold); }}
    .study-head p {{ margin: 0 0 1rem; color: var(--muted); }}
    .grid {{
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(280px, 1fr));
      gap: 1rem;
    }}
    .palace {{
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 10px;
      overflow: hidden;
      display: flex;
      flex-direction: column;
      cursor: pointer;
      transition: border-color 0.15s ease, transform 0.15s ease;
    }}
    .palace:hover, .palace:focus {{
      border-color: var(--gold);
      outline: none;
      transform: translateY(-2px);
    }}
    .palace img {{
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
    .prompt-block {{
      margin-top: 0.65rem;
      padding-top: 0.55rem;
      border-top: 1px solid var(--line);
    }}
    .prompt-head {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 0.5rem;
      color: var(--gold);
      font-size: 0.78rem;
      letter-spacing: 0.04em;
      text-transform: uppercase;
      margin-bottom: 0.35rem;
    }}
    .btn-copy {{
      background: transparent;
      color: var(--gold);
      border: 1px solid var(--line);
      border-radius: 4px;
      padding: 0.15rem 0.45rem;
      font-size: 0.75rem;
      cursor: pointer;
      text-transform: none;
      letter-spacing: 0;
    }}
    .btn-copy:hover {{ border-color: var(--gold); }}
    .prompt-text {{
      margin: 0;
      white-space: pre-wrap;
      word-break: break-word;
      font-family: ui-monospace, Consolas, monospace;
      font-size: 0.78rem;
      line-height: 1.35;
      color: var(--muted);
      max-height: 4.8em;
      overflow: hidden;
    }}
    .prompt-empty {{
      margin: 0.55rem 0 0;
      color: var(--muted);
      font-size: 0.82rem;
      font-style: italic;
    }}
    .empty {{ color: var(--muted); }}
    footer {{
      padding: 0 1.5rem 1.5rem;
      color: var(--muted);
      font-size: 0.85rem;
    }}
    #overlay {{
      display: none;
      position: fixed;
      inset: 0;
      z-index: 1000;
      background: rgba(8, 9, 12, 0.96);
      overflow: auto;
    }}
    #overlay.open {{ display: block; }}
    .overlay-bar {{
      position: sticky;
      top: 0;
      z-index: 2;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 1rem;
      padding: 0.85rem 1.25rem;
      background: #14161c;
      border-bottom: 1px solid var(--line);
    }}
    .overlay-bar h2 {{
      margin: 0;
      font-size: 1.15rem;
      font-weight: 650;
    }}
    .overlay-bar .count {{
      color: var(--gold);
      font-size: 0.95rem;
      margin: 0.2rem 0 0;
    }}
    #btnClose {{
      background: var(--gold);
      color: #1a1408;
      border: none;
      border-radius: 6px;
      padding: 0.55rem 1.1rem;
      font-size: 1rem;
      font-weight: 650;
      cursor: pointer;
    }}
    #btnClose:hover {{ filter: brightness(1.08); }}
    .overlay-nav {{
      display: flex;
      align-items: center;
      gap: 0.5rem;
    }}
    .overlay-nav button {{
      background: #242830;
      color: var(--text);
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 0.5rem 0.85rem;
      font-size: 0.95rem;
      font-weight: 600;
      cursor: pointer;
    }}
    .overlay-nav button:hover:not(:disabled) {{
      border-color: var(--gold);
      color: var(--gold);
    }}
    .overlay-nav button:disabled {{
      opacity: 0.35;
      cursor: default;
    }}
    .overlay-nav .nav-pos {{
      color: var(--muted);
      font-size: 0.9rem;
      min-width: 4.5rem;
      text-align: center;
    }}
    .overlay-nav #btnClose {{
      background: var(--gold);
      color: #1a1408;
      border: none;
      margin-left: 0.35rem;
    }}
    .overlay-nav #btnClose:hover:not(:disabled) {{
      filter: brightness(1.08);
      color: #1a1408;
      border: none;
    }}
    .overlay-body {{
      display: grid;
      grid-template-columns: 1fr;
      gap: 1.5rem;
      padding: 0.75rem 0.5rem 2rem;
      width: 100%;
      max-width: none;
      margin: 0;
      box-sizing: border-box;
    }}
    .overlay-image {{
      background: #0a0b0e;
      border: 1px solid var(--line);
      border-radius: 0;
      overflow: hidden;
      min-height: 200px;
      width: 100%;
    }}
    .overlay-image img {{
      width: 100%;
      height: auto;
      max-height: 70vh;
      object-fit: contain;
      display: block;
    }}
    .overlay-image .missing-lg {{
      color: var(--muted);
      padding: 3rem;
      text-align: center;
    }}
    .overlay-prompt {{
      margin-top: 0;
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 10px;
      padding: 0.85rem 1rem;
    }}
    .overlay-prompt .prompt-text {{
      max-height: none;
      color: var(--text);
      font-size: 0.88rem;
    }}
    .overlay-notes, .overlay-gallery {{
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 10px;
      padding: 0.85rem 1rem;
    }}
    .overlay-notes h3, .overlay-gallery h3 {{
      margin: 0 0 0.65rem;
      color: var(--gold);
      font-size: 1.05rem;
    }}
    .overlay-notes textarea {{
      width: 100%;
      min-height: 12rem;
      box-sizing: border-box;
      background: #0a0b0e;
      color: var(--text);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 0.75rem 0.85rem;
      font-family: ui-monospace, Consolas, monospace;
      font-size: 0.92rem;
      line-height: 1.5;
      resize: vertical;
    }}
    .overlay-notes .notes-toolbar {{
      display: flex;
      align-items: center;
      gap: 0.65rem;
      margin-top: 0.65rem;
      flex-wrap: wrap;
    }}
    .overlay-notes .notes-status {{
      color: var(--muted);
      font-size: 0.82rem;
    }}
    .overlay-notes .notes-status.ok {{ color: #7dcea0; }}
    .overlay-notes .notes-status.err {{ color: #e07070; }}
    .overlay-gallery .gallery-toolbar {{
      display: flex;
      align-items: center;
      gap: 0.55rem;
      margin-bottom: 0.75rem;
      flex-wrap: wrap;
    }}
    .overlay-gallery .gallery-status {{
      color: var(--muted);
      font-size: 0.82rem;
    }}
    .overlay-gallery .gallery-status.ok {{ color: #7dcea0; }}
    .overlay-gallery .gallery-status.err {{ color: #e07070; }}
    .gallery-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(220px, 1fr));
      gap: 1rem;
    }}
    .gallery-card {{
      background: #0a0b0e;
      border: 1px solid var(--line);
      border-radius: 8px;
      overflow: hidden;
      display: flex;
      flex-direction: column;
    }}
    .gallery-card img {{
      width: 100%;
      height: auto;
      max-height: 280px;
      object-fit: contain;
      display: block;
      background: #050608;
    }}
    .gallery-card .gallery-missing {{
      color: var(--muted);
      padding: 2.5rem 1rem;
      text-align: center;
      min-height: 120px;
    }}
    .gallery-card-body {{
      padding: 0.65rem 0.75rem 0.75rem;
      display: flex;
      flex-direction: column;
      gap: 0.45rem;
    }}
    .gallery-card-body input[type="text"] {{
      width: 100%;
      box-sizing: border-box;
      background: var(--panel);
      color: var(--text);
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 0.35rem 0.5rem;
      font-size: 0.88rem;
    }}
    .gallery-card-actions {{
      display: flex;
      gap: 0.4rem;
      flex-wrap: wrap;
    }}
    .btn-gallery, .btn-notes-save {{
      background: transparent;
      color: var(--gold);
      border: 1px solid var(--gold);
      border-radius: 6px;
      padding: 0.35rem 0.65rem;
      font-size: 0.85rem;
      cursor: pointer;
    }}
    .btn-gallery:hover, .btn-notes-save:hover {{
      background: var(--gold);
      color: #1a1408;
    }}
    .btn-gallery.danger {{
      color: #e07070;
      border-color: #e07070;
    }}
    .btn-gallery.danger:hover {{
      background: #e07070;
      color: #1a1408;
    }}
    .gallery-empty {{
      color: var(--muted);
      margin: 0;
      padding: 1rem 0;
    }}
    #galleryFileInput {{
      display: none;
    }}
    .practice {{
      display: flex;
      flex-direction: column;
      gap: 0.85rem;
    }}
    .practice h3 {{
      margin: 0;
      color: var(--gold);
      font-size: 1.05rem;
    }}
    #ovAtoms {{
      display: grid;
      grid-template-columns: repeat(6, minmax(0, 1fr));
      gap: 1rem;
      width: 100%;
      align-items: stretch;
    }}
    #ovAtoms > .empty {{
      grid-column: 1 / -1;
      margin: 0;
    }}
    @media (max-width: 1500px) {{
      #ovAtoms {{ grid-template-columns: repeat(5, minmax(0, 1fr)); }}
    }}
    @media (max-width: 1100px) {{
      #ovAtoms {{ grid-template-columns: repeat(3, minmax(0, 1fr)); }}
    }}
    @media (max-width: 720px) {{
      #ovAtoms {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
    }}
    @media (max-width: 480px) {{
      #ovAtoms {{ grid-template-columns: 1fr; }}
    }}
    .beast-cluster {{
      display: flex;
      flex-direction: column;
      gap: 0.65rem;
      min-width: 0;
      background: #181c24;
      border: 1px solid var(--line);
      border-left: 3px solid var(--gold);
      border-radius: 10px;
      padding: 0.75rem 0.8rem 0.85rem;
      box-sizing: border-box;
    }}
    .beast-cluster-head {{
      margin: 0;
      padding-bottom: 0.45rem;
      border-bottom: 1px solid var(--line);
    }}
    .beast-cluster-head .beast-head-row {{
      display: flex;
      align-items: baseline;
      justify-content: space-between;
      gap: 0.5rem 0.75rem;
      margin-top: 0.1rem;
    }}
    .beast-cluster-head .beast-name {{
      margin: 0;
      flex: 1;
      min-width: 0;
      font-size: 1.05rem;
      font-weight: 650;
      line-height: 1.3;
      color: var(--text);
    }}
    .beast-cluster-head .beast-name .emoji {{
      margin-right: 0.28rem;
    }}
    .beast-cluster-head .sensory-chip {{
      flex-shrink: 0;
      margin: 0;
      max-width: min(14rem, 48%);
      color: var(--text);
      font-size: 0.82rem;
      font-weight: 600;
      line-height: 1.3;
      text-align: right;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }}
    .beast-cluster-head .sensory-chip .lbl {{
      display: inline;
      margin: 0 0.35rem 0 0;
      color: var(--gold);
      font-size: 0.68rem;
      letter-spacing: 0.03em;
      text-transform: uppercase;
      vertical-align: baseline;
    }}
    .beast-cluster-atoms {{
      display: grid;
      grid-template-columns: repeat(var(--beast-cols, 1), minmax(0, 1fr));
      gap: 0.75rem;
      flex: 1;
      min-width: 0;
    }}
    .atom-card {{
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 0.75rem 0.85rem;
      scroll-margin-top: 5rem;
      min-width: 0;
      box-sizing: border-box;
      height: 100%;
    }}
    .atom-card.highlight {{
      border-color: var(--gold);
      box-shadow: 0 0 0 1px var(--gold);
    }}
    .atom-card .atom-card-head {{
      display: flex;
      align-items: flex-start;
      justify-content: space-between;
      gap: 0.5rem 0.75rem;
      margin: 0 0 0.4rem;
      padding-bottom: 0.3rem;
      border-bottom: 1px dashed var(--line);
    }}
    .atom-card .atom-card-head:empty {{
      display: none;
    }}
    .atom-card .field {{
      margin: 0.45rem 0 0;
      font-size: 0.95rem;
      line-height: 1.45;
    }}
    .atom-card .field:first-child {{ margin-top: 0; }}
    .atom-card .lbl {{
      display: block;
      color: var(--gold);
      font-size: 0.78rem;
      letter-spacing: 0.04em;
      text-transform: uppercase;
      margin-bottom: 0.15rem;
    }}
    .atom-card .field-concept {{
      margin-top: 0.55rem;
      padding: 0.55rem 0.65rem;
      background: #242830;
      border: 1px solid var(--line);
      border-left: 3px solid var(--gold);
      border-radius: 8px;
    }}
    .atom-card > .field-concept:first-child {{
      margin-top: 0;
    }}
    .atom-card .emoji {{
      margin-right: 0.28rem;
    }}
    .atom-card .zone-tag {{
      color: var(--gold);
      font-size: 0.82rem;
      font-weight: 650;
      margin: 0;
      flex: 1;
      min-width: 0;
      padding: 0;
      border: none;
    }}
    .atom-card .zone-tag .emoji {{
      margin-right: 0.28rem;
    }}
    .atom-card .sensory-chip {{
      flex-shrink: 0;
      margin-left: auto;
      max-width: 45%;
      color: var(--text);
      font-size: 0.82rem;
      font-weight: 600;
      line-height: 1.3;
      text-align: right;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }}
    .atom-card .sensory-chip .lbl {{
      display: inline;
      margin: 0 0.35rem 0 0;
      font-size: 0.68rem;
      letter-spacing: 0.03em;
      vertical-align: baseline;
    }}
    #btnMethod {{
      background: transparent;
      color: var(--gold);
      border: 1px solid var(--gold);
      border-radius: 6px;
      padding: 0.5rem 0.9rem;
      font-size: 0.95rem;
      font-weight: 650;
      cursor: pointer;
      height: 2.35rem;
    }}
    #btnMethod:hover, #btnMethod.active {{
      background: var(--gold);
      color: #1a1408;
    }}
    #btnMethod:focus {{
      outline: 2px solid #fff4d0;
      outline-offset: 2px;
    }}
    #methodPanel {{
      padding: 1rem 1.5rem 2.5rem;
    }}
    .method-layout {{
      display: grid;
      grid-template-columns: minmax(160px, 220px) minmax(0, 1fr);
      gap: 1.5rem;
      align-items: start;
    }}
    @media (max-width: 900px) {{
      .method-layout {{ grid-template-columns: 1fr; }}
      .method-toc {{ position: static !important; max-height: none !important; }}
    }}
    .method-toc {{
      position: sticky;
      top: 0.75rem;
      max-height: calc(100vh - 1.5rem);
      overflow: auto;
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 10px;
      padding: 0.85rem 1rem;
      font-size: 0.88rem;
    }}
    .method-toc h2 {{
      margin: 0 0 0.5rem;
      font-size: 0.95rem;
      color: var(--gold);
    }}
    .method-toc ul {{ margin: 0; padding-left: 1.1rem; }}
    .method-toc a {{ color: var(--muted); text-decoration: none; }}
    .method-toc a:hover {{ color: var(--gold); }}
    .method-main {{ min-width: 0; }}
    .method-banner {{
      margin: 0 0 1rem;
      padding: 0.65rem 0.85rem;
      background: #242018;
      border: 1px solid var(--line);
      border-left: 3px solid var(--gold);
      border-radius: 8px;
      color: var(--muted);
      font-size: 0.95rem;
    }}
    .method-prose {{
      max-width: 52rem;
      line-height: 1.65;
      font-size: 1.02rem;
    }}
    .method-prose h1 {{ font-size: 1.75rem; margin: 0 0 0.75rem; color: var(--text); }}
    .method-prose h2 {{
      font-size: 1.35rem;
      margin: 1.75rem 0 0.65rem;
      padding-bottom: 0.35rem;
      border-bottom: 1px solid var(--line);
      color: var(--gold);
    }}
    .method-prose h3 {{ font-size: 1.12rem; margin: 1.25rem 0 0.45rem; color: var(--text); }}
    .method-prose p, .method-prose li {{ color: var(--text); }}
    .method-prose table {{
      width: 100%;
      border-collapse: collapse;
      margin: 0.75rem 0 1.25rem;
      font-size: 0.95rem;
    }}
    .method-prose th, .method-prose td {{
      border: 1px solid var(--line);
      padding: 0.45rem 0.65rem;
      text-align: left;
      vertical-align: top;
    }}
    .method-prose th {{ background: var(--panel); color: var(--gold); }}
    .method-prose code {{
      background: #242830;
      padding: 0.1rem 0.35rem;
      border-radius: 4px;
      font-size: 0.9em;
    }}
    .method-prose pre {{
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 0.85rem 1rem;
      overflow: auto;
      font-size: 0.88rem;
      line-height: 1.45;
    }}
    .method-prose pre.mermaid {{
      background: #161820;
      text-align: center;
    }}
    .method-prose blockquote {{
      margin: 0.75rem 0;
      padding: 0.35rem 0.85rem;
      border-left: 3px solid var(--gold);
      color: var(--muted);
    }}
    .method-section {{
      margin-top: 2rem;
      max-width: 52rem;
    }}
    .method-section > h2 {{
      color: var(--gold);
      border-bottom: 1px solid var(--line);
      padding-bottom: 0.35rem;
    }}
    .method-note, .method-search-label {{ color: var(--muted); font-size: 0.9rem; }}
    .method-section input[type="search"] {{
      display: block;
      width: min(360px, 100%);
      margin: 0.35rem 0 0.85rem;
    }}
    .method-research, .method-prompt {{
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 0.55rem 0.85rem;
      margin: 0.55rem 0;
    }}
    .method-research summary, .method-prompt summary {{
      cursor: pointer;
      color: var(--gold);
      font-weight: 650;
    }}
    .method-prompt pre {{
      white-space: pre-wrap;
      word-break: break-word;
      max-height: 320px;
      overflow: auto;
      background: #161820;
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 0.65rem;
      font-size: 0.82rem;
    }}
    .btn-prompt-full {{
      margin-top: 0.4rem;
      background: transparent;
      border: 1px solid var(--line);
      color: var(--muted);
      border-radius: 6px;
      padding: 0.3rem 0.65rem;
      cursor: pointer;
    }}
    .btn-prompt-full:hover {{ color: var(--gold); border-color: var(--gold); }}
    .canon-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(160px, 1fr));
      gap: 0.45rem;
    }}
    .canon-card {{
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 0.45rem 0.6rem;
      font-size: 0.9rem;
    }}
    .canon-card .sec {{ color: var(--muted); font-size: 0.75rem; display: block; margin-bottom: 0.15rem; }}
    .canon-table-wrap {{ overflow: auto; max-height: 420px; border: 1px solid var(--line); border-radius: 8px; }}
    .canon-table {{ width: 100%; border-collapse: collapse; font-size: 0.92rem; }}
    .canon-table th, .canon-table td {{
      border-bottom: 1px solid var(--line);
      padding: 0.4rem 0.65rem;
      text-align: left;
    }}
    .canon-table th {{ position: sticky; top: 0; background: var(--panel); color: var(--gold); }}
    .canon-table tr:hover td {{ background: #242830; }}
    .method-empty {{ color: var(--muted); padding: 2rem; }}
    .method-warn {{ color: #e0a060; }}
    #btnPlans {{
      background: transparent;
      color: var(--gold);
      border: 1px solid var(--gold);
      border-radius: 6px;
      padding: 0.5rem 0.9rem;
      font-size: 0.95rem;
      font-weight: 650;
      cursor: pointer;
      height: 2.35rem;
    }}
    #btnPlans:hover, #btnPlans.active {{
      background: var(--gold);
      color: #1a1408;
    }}
    #btnPlans:focus {{
      outline: 2px solid #fff4d0;
      outline-offset: 2px;
    }}
    #practicePanel.hidden, #methodPanel.hidden, #plansPanel.hidden {{ display: none !important; }}
    #plansPanel {{
      padding: 1rem 1.5rem 2.5rem;
    }}
    .plans-layout {{
      display: grid;
      grid-template-columns: minmax(160px, 240px) minmax(0, 1fr);
      gap: 1.25rem;
      align-items: start;
      transition: grid-template-columns 0.2s ease;
    }}
    .plans-layout.toc-collapsed {{
      grid-template-columns: 2.75rem minmax(0, 1fr);
      gap: 0.85rem;
    }}
    @media (max-width: 900px) {{
      .plans-layout {{ grid-template-columns: 1fr; }}
      .plans-layout.toc-collapsed {{ grid-template-columns: 1fr; }}
      .plans-toc {{
        position: static !important;
        max-height: none !important;
      }}
      .plans-layout.toc-collapsed .plans-toc {{
        max-height: none !important;
      }}
      .plans-layout.toc-collapsed .plans-toc-head h2,
      .plans-layout.toc-collapsed #plansTocList {{
        display: none !important;
      }}
    }}
    .plans-toc {{
      position: sticky;
      top: 0.75rem;
      max-height: min(70vh, calc(100vh - 8rem));
      overflow: hidden;
      display: flex;
      flex-direction: column;
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 10px;
      padding: 0.65rem 0.75rem 0.75rem;
      font-size: 0.88rem;
    }}
    .plans-toc-head {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 0.35rem;
      margin-bottom: 0.45rem;
      flex-shrink: 0;
    }}
    .plans-toc-head h2 {{
      margin: 0;
      font-size: 0.95rem;
      color: var(--gold);
      white-space: nowrap;
    }}
    .btn-plans-toc-toggle {{
      flex-shrink: 0;
      width: 1.7rem;
      height: 1.7rem;
      padding: 0;
      background: transparent;
      border: 1px solid var(--line);
      border-radius: 6px;
      color: var(--muted);
      cursor: pointer;
      font-size: 0.7rem;
      line-height: 1;
    }}
    .btn-plans-toc-toggle:hover {{
      color: var(--gold);
      border-color: var(--gold);
    }}
    .plans-layout.toc-collapsed .plans-toc {{
      padding: 0.45rem 0.3rem;
      align-items: center;
      max-height: none;
      overflow: visible;
    }}
    .plans-layout.toc-collapsed .plans-toc-head {{
      flex-direction: column;
      margin-bottom: 0;
    }}
    .plans-layout.toc-collapsed .plans-toc-head h2 {{
      writing-mode: vertical-rl;
      transform: rotate(180deg);
      font-size: 0.78rem;
      letter-spacing: 0.06em;
      margin: 0.35rem 0;
    }}
    .plans-layout.toc-collapsed #plansTocList {{
      display: none;
    }}
    .plans-layout.toc-collapsed .btn-plans-toc-toggle {{
      order: -1;
    }}
    .plans-main {{ min-width: 0; }}
    .plans-header {{
      display: flex;
      flex-wrap: wrap;
      gap: 0.75rem 1.5rem;
      justify-content: space-between;
      align-items: start;
      margin-bottom: 0.85rem;
    }}
    .plans-header h2 {{
      margin: 0;
      font-size: 1.45rem;
      color: var(--text);
    }}
    .plans-sub {{ margin: 0.25rem 0 0; color: var(--muted); font-size: 0.92rem; }}
    .plans-header-actions {{
      display: flex;
      flex-wrap: wrap;
      gap: 0.5rem;
      align-items: center;
    }}
    .plans-github {{
      color: var(--gold);
      text-decoration: none;
      font-size: 0.9rem;
      font-weight: 650;
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 0.35rem 0.65rem;
    }}
    .plans-github:hover {{ border-color: var(--gold); }}
    .btn-plans-reset {{
      background: transparent;
      border: 1px solid var(--line);
      color: var(--muted);
      border-radius: 6px;
      padding: 0.35rem 0.65rem;
      cursor: pointer;
      font-size: 0.85rem;
    }}
    .btn-plans-reset:hover {{ color: var(--gold); border-color: var(--gold); }}
    .btn-plans-save {{
      background: var(--gold);
      color: #1a1408;
      border: none;
      border-radius: 6px;
      padding: 0.35rem 0.75rem;
      cursor: pointer;
      font-size: 0.85rem;
      font-weight: 650;
    }}
    .btn-plans-save:hover {{ filter: brightness(1.08); }}
    .btn-plans-save:disabled {{ opacity: 0.55; cursor: wait; }}
    .plans-save-status {{
      margin: 0 0 0.75rem;
      min-height: 1.2rem;
      font-size: 0.88rem;
      color: var(--muted);
    }}
    .plans-save-status.ok {{ color: #7dcea0; }}
    .plans-save-status.err {{ color: #e08080; }}
    .plans-progress-wrap {{
      display: flex;
      align-items: center;
      gap: 0.85rem;
      margin-bottom: 1.25rem;
      width: 100%;
    }}
    .plans-progress-bar {{
      flex: 1 1 auto;
      min-width: 0;
      width: 100%;
      height: 10px;
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 999px;
      overflow: hidden;
    }}
    .plans-progress-bar::after {{
      content: "";
      display: block;
      height: 100%;
      width: var(--pct, 0%);
      background: linear-gradient(90deg, #8a6914, var(--gold));
      border-radius: 999px;
      transition: width 0.2s ease;
    }}
    .plans-progress-label {{
      color: var(--muted);
      font-size: 0.88rem;
      white-space: nowrap;
      flex-shrink: 0;
    }}
    .plans-content {{ max-width: none; width: 100%; }}
    .plans-toc ul {{
      margin: 0;
      padding-left: 0;
      list-style: none;
      overflow: auto;
      flex: 1;
      min-height: 0;
      scrollbar-width: thin;
      scrollbar-color: var(--line) transparent;
    }}
    .plans-toc ul ul {{
      padding-left: 0.85rem;
      margin-top: 0.15rem;
      border-left: 1px solid var(--line);
      margin-left: 0.45rem;
    }}
    .plans-toc li {{
      margin: 0.1rem 0;
    }}
    .plans-toc .toc-row {{
      display: flex;
      align-items: flex-start;
      gap: 0.15rem;
      min-width: 0;
    }}
    .plans-toc .toc-twist {{
      flex-shrink: 0;
      width: 1.15rem;
      height: 1.15rem;
      margin-top: 0.05rem;
      padding: 0;
      background: transparent;
      border: none;
      color: var(--muted);
      cursor: pointer;
      font-size: 0.65rem;
      line-height: 1.15rem;
      text-align: center;
      border-radius: 3px;
    }}
    .plans-toc .toc-twist:hover {{ color: var(--gold); background: #242830; }}
    .plans-toc .toc-twist.leaf {{
      visibility: hidden;
      pointer-events: none;
    }}
    .plans-toc .toc-link {{
      flex: 1;
      min-width: 0;
      color: var(--muted);
      text-decoration: none;
      line-height: 1.35;
      padding: 0.1rem 0.2rem;
      border-radius: 4px;
    }}
    .plans-toc .toc-link:hover {{ color: var(--gold); }}
    .plans-toc li.toc-collapsed > ul {{ display: none; }}
    .plans-toc .toc-level-2 > .toc-row .toc-link {{ color: var(--text); font-weight: 650; }}
    .plans-toc .toc-level-3 > .toc-row .toc-link {{ font-size: 0.86rem; }}
    .plans-toc .toc-level-4 > .toc-row .toc-link,
    .plans-toc .toc-level-5 > .toc-row .toc-link {{ font-size: 0.82rem; }}
    .plans-backlog {{
      background: #242018;
      border: 1px solid var(--line);
      border-left: 3px solid var(--gold);
      border-radius: 8px;
      padding: 0.75rem 1rem;
      margin-bottom: 1.5rem;
    }}
    .plans-backlog h3 {{
      margin: 0 0 0.5rem;
      font-size: 1rem;
      color: var(--gold);
    }}
    .plans-backlog .plan-todos {{
      margin: 0.35rem 0 0.75rem;
    }}
    .plans-backlog-empty {{
      margin: 0.35rem 0 0.75rem;
      color: var(--muted);
      font-size: 0.9rem;
    }}
    .plans-backlog-add {{
      display: flex;
      gap: 0.5rem;
      align-items: stretch;
      margin-top: 0.25rem;
    }}
    .plans-backlog-add input {{
      flex: 1;
      min-width: 0;
      background: #1a1710;
      border: 1px solid var(--line);
      border-radius: 6px;
      color: var(--text);
      padding: 0.45rem 0.65rem;
      font: inherit;
    }}
    .plans-backlog-add input:focus {{
      outline: none;
      border-color: var(--gold);
    }}
    .plans-backlog-add button {{
      flex-shrink: 0;
      background: var(--gold);
      color: #1a1408;
      border: none;
      border-radius: 6px;
      padding: 0.45rem 0.85rem;
      font: inherit;
      font-weight: 650;
      cursor: pointer;
    }}
    .plans-backlog-add button:disabled {{
      opacity: 0.55;
      cursor: default;
    }}
    .plan-section {{
      margin-bottom: 1rem;
      scroll-margin-top: 0.75rem;
    }}
    .plan-section.level-2 {{
      background: #161922;
      border-left: 3px solid var(--gold);
      border-radius: 0 10px 10px 0;
      padding: 0.85rem 1rem 1rem;
      margin-bottom: 1.5rem;
    }}
    .plan-section.level-3 {{
      background: #1a1e28;
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 0.65rem 0.85rem 0.75rem;
      margin-top: 0.65rem;
    }}
    .plan-section.plan-topic-card {{
      background: #1e2430;
      border: 1px solid #3a4560;
      border-radius: 8px;
      padding: 0.75rem 1rem;
      margin-top: 0.5rem;
      margin-bottom: 0.65rem;
    }}
    .plan-section.level-2.plan-topic-card {{
      background: #181c26;
      border: 1px solid #3a4560;
      border-left: 3px solid var(--gold);
    }}
    .plan-section.level-3.plan-topic-card {{
      background: #1c2130;
    }}
    .plan-topic-body {{
      min-height: 0;
    }}
    .plan-section h3, .plan-section h4, .plan-section h5, .plan-section h6 {{
      color: var(--text);
      margin: 0 0 0.55rem;
      line-height: 1.35;
    }}
    .plan-section.level-2 > h3 {{
      font-size: 1.35rem;
      color: var(--gold);
      border-bottom: 1px solid var(--line);
      padding-bottom: 0.35rem;
      margin-top: 0;
    }}
    .plan-section.level-3 > h4 {{ font-size: 1.12rem; margin-top: 0; }}
    .plan-section.level-4 > h5 {{ font-size: 1.02rem; color: var(--text); }}
    .plan-section.level-5 > h6 {{ font-size: 0.95rem; color: var(--muted); }}
    .plan-topic-card > h5, .plan-topic-card > h6 {{
      color: var(--text);
      font-size: 1rem;
      margin-bottom: 0.45rem;
    }}
    .plan-todos {{
      list-style: none;
      margin: 0 0 0.65rem;
      padding: 0;
    }}
    .plan-todo {{
      display: flex;
      align-items: flex-start;
      gap: 0.55rem;
      padding: 0.35rem 0.25rem;
      border-radius: 6px;
      line-height: 1.45;
    }}
    .plan-todo:hover {{ background: #1e222c; }}
    .plan-todo input[type="checkbox"] {{
      margin-top: 0.2rem;
      width: 1rem;
      height: 1rem;
      accent-color: var(--gold);
      flex-shrink: 0;
      cursor: pointer;
    }}
    .plan-todo label {{
      cursor: pointer;
      flex: 1;
      min-width: 0;
    }}
    .plan-todo.done label {{
      color: var(--muted);
      text-decoration: line-through;
      text-decoration-color: #5a5548;
    }}
    .plan-todo-remove {{
      flex-shrink: 0;
      margin-left: auto;
      background: transparent;
      border: 1px solid transparent;
      color: var(--muted);
      border-radius: 4px;
      padding: 0.1rem 0.4rem;
      font: inherit;
      font-size: 0.95rem;
      line-height: 1.2;
      cursor: pointer;
    }}
    .plan-todo-remove:hover {{
      color: #e8b4b4;
      border-color: #5a3030;
      background: #2a1a1a;
    }}
    .plan-sources {{
      margin: 0.65rem 0 0;
      padding-top: 0.45rem;
      border-top: 1px solid var(--line);
      display: flex;
      flex-wrap: wrap;
      gap: 0.3rem 0.4rem;
      align-items: center;
    }}
    .plan-sources-footer {{
      margin-top: auto;
    }}
    .plan-sources-label {{
      width: 100%;
      font-size: 0.72rem;
      font-weight: 500;
      color: var(--muted);
      letter-spacing: 0.03em;
      text-transform: uppercase;
      margin-bottom: 0.05rem;
    }}
    .plan-source-chip {{
      display: inline-flex;
      align-items: center;
      gap: 0.25rem;
      max-width: 100%;
      background: rgba(28, 31, 40, 0.6);
      border: 1px solid rgba(42, 46, 58, 0.9);
      border-radius: 999px;
      padding: 0.12rem 0.35rem 0.12rem 0.28rem;
      font-size: 0.78rem;
      line-height: 1.25;
    }}
    .plan-source-chip:hover, .plan-source-chip:focus-within {{
      border-color: #4a6a9a;
      background: rgba(26, 36, 56, 0.85);
    }}
    .plan-source-link {{
      display: inline-flex;
      align-items: center;
      gap: 0.28rem;
      color: #a8c4e8;
      text-decoration: none;
      min-width: 0;
    }}
    .plan-source-link:hover {{ color: #dce8f8; }}
    .plan-source-title {{
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
      max-width: 14rem;
    }}
    .plan-source-icon {{
      display: inline-flex;
      align-items: center;
      justify-content: center;
      width: 0.95rem;
      height: 0.95rem;
      flex-shrink: 0;
    }}
    .plan-source-icon svg {{
      width: 0.9rem;
      height: 0.9rem;
      display: block;
    }}
    .plan-source-actions {{
      display: inline-flex;
      gap: 0.05rem;
      margin-left: 0.05rem;
      opacity: 0;
      transition: opacity 0.12s ease;
    }}
    .plan-source-chip:hover .plan-source-actions,
    .plan-source-chip:focus-within .plan-source-actions {{
      opacity: 1;
    }}
    .plan-source-actions button {{
      background: transparent;
      border: 1px solid transparent;
      color: var(--muted);
      border-radius: 3px;
      padding: 0 0.2rem;
      font: inherit;
      font-size: 0.72rem;
      line-height: 1.1;
      cursor: pointer;
    }}
    .plan-source-actions button:hover {{
      color: #e8b4b4;
      border-color: #5a3030;
      background: #2a1a1a;
    }}
    .plan-source-empty {{
      color: var(--muted);
      font-size: 0.75rem;
      font-style: normal;
      opacity: 0.85;
    }}
    .plan-source-add {{
      display: inline;
      border: none;
      padding: 0;
      background: transparent;
      color: var(--muted);
      font: inherit;
      font-size: 0.75rem;
      cursor: pointer;
      text-decoration: underline;
      text-underline-offset: 2px;
    }}
    .plan-source-add:hover {{
      color: #a8c4e8;
      background: transparent;
    }}
    .plan-source-form {{
      width: 100%;
      display: flex;
      flex-wrap: nowrap;
      gap: 0.3rem;
      margin-top: 0.2rem;
      align-items: center;
    }}
    .plan-source-form.hidden {{ display: none; }}
    .plan-source-form input {{
      flex: 1 1 6rem;
      min-width: 5rem;
      background: #12161f;
      border: 1px solid var(--line);
      border-radius: 4px;
      color: var(--text);
      padding: 0.22rem 0.4rem;
      font: inherit;
      font-size: 0.78rem;
    }}
    .plan-source-form input:focus {{
      outline: none;
      border-color: #4a6a9a;
    }}
    .plan-source-form button {{
      flex-shrink: 0;
      background: transparent;
      border: 1px solid var(--line);
      color: var(--muted);
      border-radius: 4px;
      padding: 0.22rem 0.45rem;
      font: inherit;
      font-size: 0.75rem;
      cursor: pointer;
    }}
    .plan-source-form button:hover {{
      color: var(--text);
      border-color: #4a6a9a;
    }}
    .plan-source-form button.plan-source-save {{
      color: #a8c4e8;
    }}
    .plan-source-form button.plan-source-cancel {{
      background: transparent;
      border-color: transparent;
      color: var(--muted);
    }}
    .plans-empty {{ color: var(--muted); padding: 2rem 0; }}
    .plans-empty.hidden {{ display: none; }}
  </style>
</head>
<body>
  <header class="top">
    <div>
      <h1>Memory Palace</h1>
      <p class="sub">{totals.get("studies", 0)} studies · {totals.get("palaces", 0)} Memory Palaces ·
         {totals.get("beasts", 0)} beasts · {totals.get("atoms", 0)} Knowledge Atoms</p>
    </div>
    <div class="controls">
      <div>
        <label for="study">Study</label>
        <select id="study">
          {''.join(options) if options else '<option value="">No studies</option>'}
        </select>
      </div>
      <button type="button" id="btnLatest">Latest palace</button>
      <button type="button" id="btnPlans" title="Study plans (P)">Plans</button>
      <button type="button" id="btnMethod" title="Method docs (M)">Method</button>
      <div>
        <label for="atomSearch">Search Knowledge Atoms</label>
        <input type="search" id="atomSearch" placeholder="Beast, concept, quote, story…" autocomplete="off"/>
      </div>
    </div>
  </header>
  <div id="searchResults" aria-live="polite"></div>
  <main id="practicePanel">
    {''.join(study_blocks)}
  </main>
  <div id="methodPanel" class="hidden" aria-hidden="true">
    {method_body}
  </div>
  {plans_body}
  <footer>Click a Memory Palace for fullscreen practice. Overlay: ← Older / Newer → · <strong>C</strong> or Copy prompt · Esc close. Notes auto-save at bottom; gallery add/edit/delete. <strong>P</strong> plans · <strong>M</strong> method · Latest palace (or L) opens the highest palace number.</footer>

  <div id="studyModal" aria-hidden="true" role="dialog" aria-modal="true" aria-labelledby="studyModalTitle">
    <div class="modal-panel">
      <h2 id="studyModalTitle">Select a study</h2>
      <p class="hint" id="studyModalHint"></p>
      <ul class="pick-list" id="studyModalList"></ul>
      <p class="modal-foot" id="studyModalFoot"></p>
    </div>
  </div>

  <div id="overlay" aria-hidden="true">
    <div class="overlay-bar">
      <div>
        <h2 id="ovTitle">Memory Palace</h2>
        <p class="count" id="ovCount"></p>
      </div>
      <div class="overlay-nav">
        <button type="button" id="btnPrevPalace" title="Older palace (←)">← Older</button>
        <span class="nav-pos" id="ovNavPos"></span>
        <button type="button" id="btnNextPalace" title="Newer palace (→)">Newer →</button>
        <button type="button" id="btnCopyPrompt" title="Copy image prompt (C)" disabled>Copy prompt</button>
        <button type="button" id="btnClose">Close</button>
      </div>
    </div>
    <div class="overlay-body">
      <div class="overlay-image" id="ovImage"></div>
      <div class="practice">
        <h3>Practice — Knowledge Atoms</h3>
        <div id="ovAtoms"></div>
      </div>
      <div class="overlay-prompt" id="ovPrompt"></div>
      <div class="overlay-notes" id="ovNotes">
        <h3>Notes</h3>
        <textarea id="ovNotesInput" placeholder="Palace notes for recall, links, reminders…"></textarea>
        <div class="notes-toolbar">
          <button type="button" class="btn-notes-save" id="btnNotesSave">Save notes</button>
          <span class="notes-status" id="ovNotesStatus"></span>
        </div>
      </div>
      <div class="overlay-gallery" id="ovGallery">
        <h3>Gallery</h3>
        <div class="gallery-toolbar">
          <button type="button" class="btn-gallery" id="btnGalleryAdd">Add image</button>
          <span class="gallery-status" id="ovGalleryStatus"></span>
        </div>
        <input type="file" id="galleryFileInput" accept="image/*"/>
        <div class="gallery-grid" id="ovGalleryGrid"></div>
      </div>
    </div>
  </div>

  <script>
    const PALACE_DATA = {palace_json};
    const STUDY_LATEST = {latest_json};
    const STUDIES = {studies_json};
    const METHOD_CANON = {method_canon_json};
    const PLANS_DATA = {plans_data_json};
    const sel = document.getElementById('study');
    const sections = [...document.querySelectorAll('.study')];
    const overlay = document.getElementById('overlay');
    const ovTitle = document.getElementById('ovTitle');
    const ovCount = document.getElementById('ovCount');
    const ovImage = document.getElementById('ovImage');
    const ovPrompt = document.getElementById('ovPrompt');
    const ovNotesInput = document.getElementById('ovNotesInput');
    const ovNotesStatus = document.getElementById('ovNotesStatus');
    const btnNotesSave = document.getElementById('btnNotesSave');
    const ovGalleryGrid = document.getElementById('ovGalleryGrid');
    const ovGalleryStatus = document.getElementById('ovGalleryStatus');
    const btnGalleryAdd = document.getElementById('btnGalleryAdd');
    const galleryFileInput = document.getElementById('galleryFileInput');
    const ovAtoms = document.getElementById('ovAtoms');
    const btnClose = document.getElementById('btnClose');
    const btnPrevPalace = document.getElementById('btnPrevPalace');
    const btnNextPalace = document.getElementById('btnNextPalace');
    const btnCopyPrompt = document.getElementById('btnCopyPrompt');
    const ovNavPos = document.getElementById('ovNavPos');
    const btnLatest = document.getElementById('btnLatest');
    const btnMethod = document.getElementById('btnMethod');
    const btnPlans = document.getElementById('btnPlans');
    const practicePanel = document.getElementById('practicePanel');
    const methodPanel = document.getElementById('methodPanel');
    const plansPanel = document.getElementById('plansPanel');
    const atomSearch = document.getElementById('atomSearch');
    const searchResults = document.getElementById('searchResults');
    const studyModal = document.getElementById('studyModal');
    const studyModalTitle = document.getElementById('studyModalTitle');
    const studyModalHint = document.getElementById('studyModalHint');
    const studyModalList = document.getElementById('studyModalList');
    const studyModalFoot = document.getElementById('studyModalFoot');
    let searchTimer = null;
    let currentOverlayPalaceId = '';
    let notesSaveTimer = null;
    let notesDirty = false;
    let notesSaving = false;
    let galleryBusy = false;
    let dashboardView = 'practice';
    let mermaidReady = false;
    const PLANS_STORAGE_KEY = 'mnemonics_plans_v1';
    const PLANS_TOC_COLLAPSED_KEY = 'mnemonics_plans_toc_collapsed';
    const PLANS_TOC_NODES_KEY = 'mnemonics_plans_toc_nodes';

    // Newest-first order per study (same as grid)
    const STUDY_PALACE_ORDER = {{}};
    Object.values(PALACE_DATA).forEach(st => {{
      const sid = st.study_id || '';
      if (!STUDY_PALACE_ORDER[sid]) STUDY_PALACE_ORDER[sid] = [];
      STUDY_PALACE_ORDER[sid].push(st);
    }});
    Object.keys(STUDY_PALACE_ORDER).forEach(sid => {{
      STUDY_PALACE_ORDER[sid].sort((a, b) => Number(b.number) - Number(a.number));
      STUDY_PALACE_ORDER[sid] = STUDY_PALACE_ORDER[sid].map(s => s.id);
    }});

    function palaceOrderFor(st) {{
      if (!st) return [];
      return STUDY_PALACE_ORDER[st.study_id] || [];
    }}

    function updatePalaceNav(id) {{
      const st = PALACE_DATA[id];
      const order = palaceOrderFor(st);
      const idx = order.indexOf(id);
      const hasPrev = idx >= 0 && idx < order.length - 1; // older = higher index
      const hasNext = idx > 0; // newer = lower index
      if (btnPrevPalace) btnPrevPalace.disabled = !hasPrev;
      if (btnNextPalace) btnNextPalace.disabled = !hasNext;
      if (ovNavPos) {{
        ovNavPos.textContent = (idx >= 0 && order.length)
          ? ((idx + 1) + ' / ' + order.length)
          : '';
      }}
    }}

    function stepPalace(delta) {{
      // delta +1 = older (toward end of newest-first list), -1 = newer
      if (notesDirty) savePalaceNotes(true);
      const st = PALACE_DATA[currentOverlayPalaceId];
      const order = palaceOrderFor(st);
      const idx = order.indexOf(currentOverlayPalaceId);
      if (idx < 0) return;
      const nextIdx = idx + delta;
      if (nextIdx < 0 || nextIdx >= order.length) return;
      openPalace(order[nextIdx]);
    }}

    function show(id) {{
      sections.forEach(s => {{
        s.style.display = (s.dataset.studyId === id) ? 'block' : 'none';
      }});
    }}

    function focusLatestBtn() {{
      if (btnLatest) requestAnimationFrame(() => btnLatest.focus());
    }}

    function selectStudy(id, {{ focusLatest = true }} = {{}}) {{
      if (!sel || !id) return;
      sel.value = id;
      show(id);
      if (focusLatest) focusLatestBtn();
    }}

    if (sel) {{
      sel.addEventListener('change', () => {{
        show(sel.value);
        focusLatestBtn();
        if (dashboardView === 'plans') renderPlansForStudy(sel.value);
      }});
      if (sel.value) show(sel.value);
    }}

    function studyModalOpen() {{
      return studyModal && studyModal.classList.contains('open');
    }}

    function closeStudyModal({{ focusLatest = true }} = {{}}) {{
      if (!studyModal) return;
      studyModal.classList.remove('open');
      studyModal.setAttribute('aria-hidden', 'true');
      if (focusLatest) focusLatestBtn();
    }}

    function openStudyModal() {{
      if (!studyModal || !STUDIES.length) {{
        focusLatestBtn();
        return;
      }}
      studyModal.classList.add('open');
      studyModal.setAttribute('aria-hidden', 'false');
      renderStudyModal();
    }}

    function studyShortcut(idx) {{
      // Letters a–z; numbers 1–9 also map to the first nine studies
      if (idx < 0 || idx >= STUDIES.length || idx > 25) return null;
      return String.fromCharCode(97 + idx);
    }}

    function renderStudyModal() {{
      studyModalTitle.textContent = 'Select a study';
      studyModalHint.textContent = 'Press a letter to choose a study (1–9 also work for the first nine).';
      studyModalFoot.textContent = 'Esc keeps the current study';
      studyModalList.innerHTML = '';
      STUDIES.forEach((st, idx) => {{
        const letter = studyShortcut(idx);
        if (!letter) return;
        const li = document.createElement('li');
        const btn = document.createElement('button');
        btn.type = 'button';
        const numHint = (idx < 9)
          ? ' <span style="color:var(--muted)">(' + (idx + 1) + ')</span>'
          : '';
        btn.innerHTML = '<span class="key">' + esc(letter.toUpperCase()) + '</span>'
          + '<span class="pick-label">' + esc(st.title) + numHint + '</span>';
        btn.addEventListener('click', () => pickStudyItem(st.id));
        li.appendChild(btn);
        studyModalList.appendChild(li);
      }});
      const firstBtn = studyModalList.querySelector('button');
      if (firstBtn) requestAnimationFrame(() => firstBtn.focus());
    }}

    function pickStudyItem(id) {{
      selectStudy(id, {{ focusLatest: false }});
      closeStudyModal({{ focusLatest: false }});
      if (dashboardView !== 'practice') setDashboardView('practice');
      openLatestPalace();
    }}

    function handleStudyModalKey(e) {{
      if (!studyModalOpen()) return false;
      if (e.key === 'Escape') {{
        e.preventDefault();
        closeStudyModal({{ focusLatest: true }});
        return true;
      }}
      if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {{
        const ch = e.key.toLowerCase();
        let idx = -1;
        if (ch >= 'a' && ch <= 'z') {{
          idx = ch.charCodeAt(0) - 97;
        }} else if (ch >= '1' && ch <= '9') {{
          idx = parseInt(ch, 10) - 1;
        }}
        if (idx >= 0 && idx < STUDIES.length) {{
          e.preventDefault();
          pickStudyItem(STUDIES[idx].id);
          return true;
        }}
      }}
      return true;
    }}

    function esc(s) {{
      return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
    }}

    function dash(v) {{
      const t = (v ?? '').toString().trim();
      return t ? esc(t) : '—';
    }}

    function sensoryEmoji(v) {{
      const k = (v ?? '').toString().trim().toLowerCase();
      const map = {{
        visual: '👁️',
        auditory: '👂',
        tactile: '✋',
        olfactory: '👃',
        gustatory: '👅',
        thermal: '🌡️'
      }};
      return map[k] || '';
    }}

    function formatSensory(v) {{
      const t = (v ?? '').toString().trim();
      if (!t) return '—';
      const emoji = sensoryEmoji(t);
      return emoji
        ? '<span class="emoji" aria-hidden="true">' + emoji + '</span>' + esc(t)
        : esc(t);
    }}

    function formatConcept(v) {{
      const t = (v ?? '').toString().trim();
      if (!t) return '—';
      return '<span class="emoji" aria-hidden="true">💡</span>' + esc(t);
    }}

    function formatQuote(v) {{
      const t = (v ?? '').toString().trim();
      if (!t) return '—';
      return '\u201c' + esc(t) + '\u201d';
    }}

    async function copyPrompt(text, btn) {{
      const t = (text || '').toString();
      if (!t.trim()) return;
      try {{
        await navigator.clipboard.writeText(t);
        if (btn) {{
          btn.textContent = 'Copied';
          setTimeout(() => {{ btn.textContent = 'Copy prompt'; }}, 1200);
        }}
      }} catch (e) {{
        if (btn) {{
          btn.textContent = 'Copy failed';
          setTimeout(() => {{ btn.textContent = 'Copy prompt'; }}, 1600);
        }}
      }}
    }}

    function dashboardSaveBase() {{
      return (PLANS_DATA.save_url || 'http://127.0.0.1:8765').replace(/\\/$/, '');
    }}

    function setNotesStatus(text, kind) {{
      if (!ovNotesStatus) return;
      ovNotesStatus.textContent = text || '';
      ovNotesStatus.className = 'notes-status' + (kind ? (' ' + kind) : '');
    }}

    function setGalleryStatus(text, kind) {{
      if (!ovGalleryStatus) return;
      ovGalleryStatus.textContent = text || '';
      ovGalleryStatus.className = 'gallery-status' + (kind ? (' ' + kind) : '');
    }}

    function palaceActionError(data, res, verb) {{
      let msg = (data && data.error) ? data.error : '';
      if (!msg && data && data.ok === false) {{
        msg = 'Save server outdated — press [D] again';
      }}
      if (!msg) msg = verb + ' failed (' + res.status + ')';
      return msg + ' — reopen dashboard from Memory Palace [D].';
    }}

    async function savePalaceNotes(force) {{
      if (!currentOverlayPalaceId || !ovNotesInput) return;
      if (notesSaving) return;
      const st = PALACE_DATA[currentOverlayPalaceId];
      if (!st) return;
      const notes = ovNotesInput.value;
      if (!force && notes === (st.palace_notes || '')) {{
        notesDirty = false;
        return;
      }}
      notesSaving = true;
      setNotesStatus('Saving…', '');
      if (btnNotesSave) btnNotesSave.disabled = true;
      try {{
        const res = await fetch(dashboardSaveBase() + '/palace/notes', {{
          method: 'POST',
          headers: {{ 'Content-Type': 'application/json' }},
          body: JSON.stringify({{ palace_id: currentOverlayPalaceId, notes }})
        }});
        const data = await res.json().catch(() => ({{}}));
        if (!res.ok || !data.ok) {{
          setNotesStatus(palaceActionError(data, res, 'Notes save'), 'err');
          return;
        }}
        st.palace_notes = notes;
        notesDirty = false;
        setNotesStatus('Saved', 'ok');
      }} catch (e) {{
        setNotesStatus('Save server unavailable — reopen dashboard from Memory Palace [D].', 'err');
      }} finally {{
        notesSaving = false;
        if (btnNotesSave) btnNotesSave.disabled = false;
      }}
    }}

    function scheduleNotesSave() {{
      notesDirty = true;
      if (notesSaveTimer) clearTimeout(notesSaveTimer);
      notesSaveTimer = setTimeout(() => savePalaceNotes(false), 800);
    }}

    function renderGalleryGrid(st) {{
      if (!ovGalleryGrid) return;
      const items = (st && st.gallery_images) ? st.gallery_images.slice() : [];
      items.sort((a, b) => Number(a.sort_order || 0) - Number(b.sort_order || 0));
      if (!items.length) {{
        ovGalleryGrid.innerHTML = '<p class="gallery-empty">No gallery images yet. Use Add image for supplementary recall photos.</p>';
        return;
      }}
      ovGalleryGrid.innerHTML = items.map(img => {{
        const preview = img.image_uri
          ? '<img src="' + esc(img.image_uri) + '" alt="' + esc(img.caption || 'Gallery image') + '" loading="lazy"/>'
          : '<div class="gallery-missing">Image missing</div>';
        return '<article class="gallery-card" data-gallery-id="' + esc(img.id) + '">'
          + preview
          + '<div class="gallery-card-body">'
          + '<input type="text" class="gallery-caption" value="' + esc(img.caption || '') + '" placeholder="Caption (optional)" data-gallery-id="' + esc(img.id) + '"/>'
          + '<div class="gallery-card-actions">'
          + '<button type="button" class="btn-gallery" data-gallery-save="' + esc(img.id) + '">Save caption</button>'
          + '<button type="button" class="btn-gallery" data-gallery-up="' + esc(img.id) + '">↑</button>'
          + '<button type="button" class="btn-gallery" data-gallery-down="' + esc(img.id) + '">↓</button>'
          + '<button type="button" class="btn-gallery danger" data-gallery-delete="' + esc(img.id) + '">Delete</button>'
          + '</div></div></article>';
      }}).join('');
      ovGalleryGrid.querySelectorAll('[data-gallery-save]').forEach(btn => {{
        btn.addEventListener('click', () => updateGalleryCaption(btn.getAttribute('data-gallery-save')));
      }});
      ovGalleryGrid.querySelectorAll('[data-gallery-up]').forEach(btn => {{
        btn.addEventListener('click', () => moveGalleryImage(btn.getAttribute('data-gallery-up'), -1));
      }});
      ovGalleryGrid.querySelectorAll('[data-gallery-down]').forEach(btn => {{
        btn.addEventListener('click', () => moveGalleryImage(btn.getAttribute('data-gallery-down'), 1));
      }});
      ovGalleryGrid.querySelectorAll('[data-gallery-delete]').forEach(btn => {{
        btn.addEventListener('click', () => deleteGalleryImage(btn.getAttribute('data-gallery-delete')));
      }});
    }}

    async function postGalleryAction(payload, verb) {{
      if (galleryBusy) return null;
      galleryBusy = true;
      setGalleryStatus(verb + '…', '');
      if (btnGalleryAdd) btnGalleryAdd.disabled = true;
      try {{
        const res = await fetch(dashboardSaveBase() + '/palace/images', {{
          method: 'POST',
          headers: {{ 'Content-Type': 'application/json' }},
          body: JSON.stringify(payload)
        }});
        const data = await res.json().catch(() => ({{}}));
        if (!res.ok || !data.ok) {{
          setGalleryStatus(palaceActionError(data, res, verb), 'err');
          return null;
        }}
        return data;
      }} catch (e) {{
        setGalleryStatus('Save server unavailable — reopen dashboard from Memory Palace [D].', 'err');
        return null;
      }} finally {{
        galleryBusy = false;
        if (btnGalleryAdd) btnGalleryAdd.disabled = false;
      }}
    }}

    function syncGalleryFromResponse(st, data) {{
      if (!st || !data) return;
      if (data.action === 'delete') {{
        st.gallery_images = (st.gallery_images || []).filter(img => img.id !== data.image_id);
      }} else if (data.image) {{
        const imgs = st.gallery_images || [];
        const idx = imgs.findIndex(img => img.id === data.image.id);
        const entry = {{
          id: data.image.id,
          caption: data.image.caption || '',
          sort_order: data.image.sort_order || '',
          image_rel_path: data.image.image_rel_path || '',
          image_uri: ''
        }};
        if (idx >= 0) imgs[idx] = entry;
        else imgs.push(entry);
        st.gallery_images = imgs;
      }}
      renderGalleryGrid(st);
      setGalleryStatus('Saved', 'ok');
    }}

    async function updateGalleryCaption(imageId) {{
      const st = PALACE_DATA[currentOverlayPalaceId];
      if (!st) return;
      const input = ovGalleryGrid.querySelector('.gallery-caption[data-gallery-id="' + CSS.escape(imageId) + '"]');
      const caption = input ? input.value : '';
      const data = await postGalleryAction({{
        action: 'update',
        palace_id: currentOverlayPalaceId,
        image_id: imageId,
        caption
      }}, 'Save caption');
      syncGalleryFromResponse(st, data);
    }}

    async function moveGalleryImage(imageId, delta) {{
      const st = PALACE_DATA[currentOverlayPalaceId];
      if (!st) return;
      const items = (st.gallery_images || []).slice().sort((a, b) => Number(a.sort_order || 0) - Number(b.sort_order || 0));
      const idx = items.findIndex(img => img.id === imageId);
      if (idx < 0) return;
      const swapIdx = idx + delta;
      if (swapIdx < 0 || swapIdx >= items.length) return;
      const tmp = items[idx];
      items[idx] = items[swapIdx];
      items[swapIdx] = tmp;
      items.forEach((img, i) => {{ img.sort_order = String(i + 1); }});
      st.gallery_images = items;
      renderGalleryGrid(st);
      for (const img of items) {{
        const data = await postGalleryAction({{
          action: 'update',
          palace_id: currentOverlayPalaceId,
          image_id: img.id,
          sort_order: img.sort_order
        }}, 'Reorder');
        if (!data) return;
      }}
      setGalleryStatus('Order saved', 'ok');
    }}

    async function deleteGalleryImage(imageId) {{
      if (!window.confirm('Delete this gallery image?')) return;
      const st = PALACE_DATA[currentOverlayPalaceId];
      if (!st) return;
      const data = await postGalleryAction({{
        action: 'delete',
        palace_id: currentOverlayPalaceId,
        image_id: imageId
      }}, 'Delete');
      syncGalleryFromResponse(st, data);
    }}

    function readFileAsDataUrl(file) {{
      return new Promise((resolve, reject) => {{
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = () => reject(reader.error);
        reader.readAsDataURL(file);
      }});
    }}

    async function addGalleryImageFromFile(file) {{
      if (!file || !currentOverlayPalaceId) return;
      const st = PALACE_DATA[currentOverlayPalaceId];
      if (!st) return;
      let dataUrl = '';
      try {{
        dataUrl = await readFileAsDataUrl(file);
      }} catch (e) {{
        setGalleryStatus('Could not read image file.', 'err');
        return;
      }}
      const data = await postGalleryAction({{
        action: 'add',
        palace_id: currentOverlayPalaceId,
        data_b64: dataUrl,
        mime: file.type || ''
      }}, 'Add image');
      if (!data || !data.image) return;
      const entry = {{
        id: data.image.id,
        caption: data.image.caption || '',
        sort_order: data.image.sort_order || '',
        image_rel_path: data.image.image_rel_path || '',
        image_uri: dataUrl
      }};
      st.gallery_images = (st.gallery_images || []).concat([entry]);
      renderGalleryGrid(st);
      setGalleryStatus('Image added — reopen dashboard [D] if preview is missing.', 'ok');
    }}

    function renderPromptBlock(prompt, id) {{
      const t = (prompt || '').toString().trim();
      if (!t) return '<p class="prompt-empty">No image prompt saved.</p>';
      return '<div class="prompt-head"><span>Image prompt</span>'
        + '<button type="button" class="btn-copy" data-copy-id="' + esc(id) + '">Copy prompt</button></div>'
        + '<pre class="prompt-text">' + esc(prompt) + '</pre>';
    }}

    function groupAtomsByBeast(atoms) {{
      const groups = [];
      const seen = new Map();
      (atoms || []).forEach(a => {{
        const key = (a.beast || '').toString().trim() || '—';
        if (!seen.has(key)) {{
          const g = {{ beast: key, atoms: [] }};
          seen.set(key, g);
          groups.push(g);
        }}
        seen.get(key).atoms.push(a);
      }});
      return groups;
    }}

    function sensoryChipHtml(sensory) {{
      return '<p class="sensory-chip"><span class="lbl">Sensory</span>'
        + formatSensory(sensory) + '</p>';
    }}

    function renderAtomCard(a, focusAtomId, {{ sensoryOnCard = true }} = {{}}) {{
      const zone = (a.zone || a.zone_label)
        ? '<p class="zone-tag"><span class="emoji" aria-hidden="true">🟦</span> '
          + dash(a.zone) + (a.zone_label ? ' · ' + dash(a.zone_label) : '') + '</p>'
        : '';
      const sensory = sensoryOnCard ? sensoryChipHtml(a.sensory) : '';
      const head = (zone || sensory)
        ? '<div class="atom-card-head">' + zone + sensory + '</div>'
        : '';
      const hl = (focusAtomId && a.id === focusAtomId) ? ' highlight' : '';
      return '<article class="atom-card' + hl + '" data-atom-id="' + esc(a.id) + '">'
        + head
        + '<p class="field field-concept"><span class="lbl">Concept</span>'
        + formatConcept(a.concept) + '</p>'
        + '<p class="field"><span class="lbl">Quote</span>' + formatQuote(a.quote) + '</p>'
        + '<p class="field"><span class="lbl">Story</span>' + dash(a.story) + '</p>'
        + '</article>';
    }}

    function renderPalaceAtoms(atoms, focusAtomId) {{
      if (!atoms || !atoms.length) {{
        return '<p class="empty">No Knowledge Atoms on this Memory Palace.</p>';
      }}
      const maxCols = 6;
      return groupAtomsByBeast(atoms).map(g => {{
        const n = g.atoms.length;
        const span = Math.min(Math.max(n, 1), maxCols);
        const innerCols = Math.min(n, maxCols);
        const sensoryOnBeast = n === 1;
        const headSensory = sensoryOnBeast ? sensoryChipHtml(g.atoms[0].sensory) : '';
        const cards = g.atoms.map(a => renderAtomCard(a, focusAtomId, {{
          sensoryOnCard: !sensoryOnBeast
        }})).join('');
        return '<section class="beast-cluster" style="grid-column: span ' + span
          + '; --beast-cols: ' + innerCols + '">'
          + '<header class="beast-cluster-head">'
          + '<span class="lbl">Beast</span>'
          + '<div class="beast-head-row">'
          + '<p class="beast-name"><span class="emoji" aria-hidden="true">🟧</span> '
          + dash(g.beast) + '</p>'
          + headSensory
          + '</div>'
          + '</header>'
          + '<div class="beast-cluster-atoms">' + cards + '</div>'
          + '</section>';
      }}).join('');
    }}

    function openPalace(id, focusAtomId) {{
      const st = PALACE_DATA[id];
      if (!st) return;
      currentOverlayPalaceId = id;
      ovTitle.textContent = 'Memory Palace ' + st.number + ': ' + st.title;
      const n = st.atom_count || (st.atoms || []).length;
      ovCount.textContent = n + ' Knowledge Atom' + (n === 1 ? '' : 's')
        + (st.character ? ' · Character: ' + st.character : '');
      if (st.image_uri) {{
        ovImage.innerHTML = '<img src="' + esc(st.image_uri) + '" alt="Memory Palace ' + esc(st.number) + '"/>';
      }} else {{
        ovImage.innerHTML = '<div class="missing-lg">No image</div>';
      }}
      ovPrompt.innerHTML = renderPromptBlock(st.image_prompt, st.id);
      ovPrompt.querySelectorAll('.btn-copy').forEach(btn => {{
        btn.addEventListener('click', (e) => {{
          e.stopPropagation();
          copyPrompt(st.image_prompt, btn);
        }});
      }});
      if (btnCopyPrompt) {{
        const hasPrompt = !!(st.image_prompt || '').toString().trim();
        btnCopyPrompt.disabled = !hasPrompt;
        btnCopyPrompt.onclick = (e) => {{
          e.stopPropagation();
          copyPrompt(st.image_prompt, btnCopyPrompt);
        }};
      }}
      const atoms = st.atoms || [];
      ovAtoms.innerHTML = renderPalaceAtoms(atoms, focusAtomId);
      if (ovNotesInput) {{
        ovNotesInput.value = st.palace_notes || '';
        notesDirty = false;
        setNotesStatus('', '');
      }}
      renderGalleryGrid(st);
      setGalleryStatus('', '');
      updatePalaceNav(id);
      overlay.classList.add('open');
      overlay.setAttribute('aria-hidden', 'false');
      if (btnClose) btnClose.focus();
      if (focusAtomId) {{
        requestAnimationFrame(() => {{
          const el = ovAtoms.querySelector('[data-atom-id="' + CSS.escape(focusAtomId) + '"]');
          if (el) el.scrollIntoView({{ behavior: 'smooth', block: 'center' }});
        }});
      }}
    }}

    function closeOverlay() {{
      if (notesDirty) savePalaceNotes(true);
      overlay.classList.remove('open');
      overlay.setAttribute('aria-hidden', 'true');
      currentOverlayPalaceId = '';
      if (btnCopyPrompt) {{
        btnCopyPrompt.disabled = true;
        btnCopyPrompt.onclick = null;
      }}
    }}

    document.querySelectorAll('.palace').forEach(el => {{
      el.addEventListener('click', (e) => {{
        if (e.target.closest('.btn-copy')) return;
        openPalace(el.dataset.palaceId);
      }});
      el.addEventListener('keydown', (e) => {{
        if (e.key === 'Enter' || e.key === ' ') {{
          e.preventDefault();
          openPalace(el.dataset.palaceId);
        }}
      }});
    }});
    document.querySelectorAll('.btn-copy[data-copy-prompt]').forEach(btn => {{
      btn.addEventListener('click', (e) => {{
        e.stopPropagation();
        const id = btn.getAttribute('data-copy-prompt');
        const st = PALACE_DATA[id];
        if (st) copyPrompt(st.image_prompt, btn);
      }});
    }});
    btnClose.addEventListener('click', closeOverlay);
    if (ovNotesInput) ovNotesInput.addEventListener('input', scheduleNotesSave);
    if (btnNotesSave) btnNotesSave.addEventListener('click', () => savePalaceNotes(true));
    if (btnGalleryAdd && galleryFileInput) {{
      btnGalleryAdd.addEventListener('click', () => galleryFileInput.click());
      galleryFileInput.addEventListener('change', () => {{
        const file = galleryFileInput.files && galleryFileInput.files[0];
        galleryFileInput.value = '';
        if (file) addGalleryImageFromFile(file);
      }});
    }}
    if (btnPrevPalace) btnPrevPalace.addEventListener('click', () => stepPalace(1));
    if (btnNextPalace) btnNextPalace.addEventListener('click', () => stepPalace(-1));
    function openLatestPalace() {{
      const studyId = sel ? sel.value : '';
      const palaceId = STUDY_LATEST[studyId];
      if (palaceId) openPalace(palaceId);
    }}
    btnLatest.addEventListener('click', openLatestPalace);

    function setDashboardView(view) {{
      const next = (view === 'method' || view === 'plans') ? view : 'practice';
      if (dashboardView === next && view !== 'toggle-method' && view !== 'toggle-plans') {{
        if (next === 'practice') return;
      }}
      if (view === 'toggle-method') {{
        dashboardView = dashboardView === 'method' ? 'practice' : 'method';
      }} else if (view === 'toggle-plans') {{
        dashboardView = dashboardView === 'plans' ? 'practice' : 'plans';
      }} else {{
        dashboardView = next;
      }}
      const isPractice = dashboardView === 'practice';
      const isMethod = dashboardView === 'method';
      const isPlans = dashboardView === 'plans';
      if (practicePanel) practicePanel.classList.toggle('hidden', !isPractice);
      if (methodPanel) {{
        methodPanel.classList.toggle('hidden', !isMethod);
        methodPanel.setAttribute('aria-hidden', isMethod ? 'false' : 'true');
      }}
      if (plansPanel) {{
        plansPanel.classList.toggle('hidden', !isPlans);
        plansPanel.setAttribute('aria-hidden', isPlans ? 'false' : 'true');
      }}
      if (btnMethod) btnMethod.classList.toggle('active', isMethod);
      if (btnPlans) btnPlans.classList.toggle('active', isPlans);
      if (searchResults && !isPractice) {{
        searchResults.classList.remove('open');
        searchResults.innerHTML = '';
      }}
      if (isMethod) {{
        ensureMermaid();
        renderCharacters('');
        renderBestiary('');
      }}
      if (isPlans) renderPlansForStudy(sel ? sel.value : '');
    }}

    function loadPlanStorage() {{
      try {{
        const raw = localStorage.getItem(PLANS_STORAGE_KEY);
        return raw ? JSON.parse(raw) : {{}};
      }} catch (e) {{
        return {{}};
      }}
    }}

    function savePlanStorage(data) {{
      try {{
        localStorage.setItem(PLANS_STORAGE_KEY, JSON.stringify(data));
      }} catch (e) {{}}
    }}

    function planChecked(plan, todo) {{
      const store = loadPlanStorage();
      if (Object.prototype.hasOwnProperty.call(store, todo.id)) {{
        return !!store[todo.id];
      }}
      return !!todo.checked;
    }}

    function setPlanChecked(todoId, checked) {{
      const store = loadPlanStorage();
      store[todoId] = !!checked;
      savePlanStorage(store);
    }}

    function resetPlanProgress(slug) {{
      const store = loadPlanStorage();
      const prefix = slug + ':';
      Object.keys(store).forEach(k => {{
        if (k.startsWith(prefix)) delete store[k];
      }});
      savePlanStorage(store);
    }}

    function setPlansTocCollapsed(collapsed) {{
      const layout = document.getElementById('plansLayout');
      const btn = document.getElementById('btnPlansTocToggle');
      if (layout) layout.classList.toggle('toc-collapsed', !!collapsed);
      if (btn) {{
        btn.setAttribute('aria-expanded', collapsed ? 'false' : 'true');
        btn.title = collapsed ? 'Expand sections' : 'Collapse sections';
        btn.textContent = collapsed ? '▶' : '◀';
      }}
      try {{
        localStorage.setItem(PLANS_TOC_COLLAPSED_KEY, collapsed ? '1' : '0');
      }} catch (e) {{}}
    }}

    function initPlansTocToggle() {{
      const btn = document.getElementById('btnPlansTocToggle');
      if (!btn) return;
      let collapsed = false;
      try {{
        collapsed = localStorage.getItem(PLANS_TOC_COLLAPSED_KEY) === '1';
      }} catch (e) {{}}
      setPlansTocCollapsed(collapsed);
      btn.addEventListener('click', () => {{
        const layout = document.getElementById('plansLayout');
        const next = !(layout && layout.classList.contains('toc-collapsed'));
        setPlansTocCollapsed(next);
      }});
    }}
    initPlansTocToggle();

    function collectPlansSavePayload() {{
      const plans = {{}};
      Object.entries(PLANS_DATA.plans || {{}}).forEach(([studyId, plan]) => {{
        const todos = (plan.todos || []).map(t => ({{
          id: t.id,
          checked: planChecked(plan, t),
        }}));
        if (todos.length) {{
          plans[studyId] = {{ study_id: studyId, slug: plan.slug, todos }};
        }}
      }});
      return {{ plans }};
    }}

    function clearSavedPlanProgress(studyIds) {{
      const store = loadPlanStorage();
      studyIds.forEach(studyId => {{
        const plan = (PLANS_DATA.plans || {{}})[studyId];
        if (!plan || !plan.slug) return;
        const prefix = plan.slug + ':';
        Object.keys(store).forEach(k => {{
          if (k.startsWith(prefix)) delete store[k];
        }});
      }});
      savePlanStorage(store);
    }}

    function applyTodoStatesToPlan(plan, byId) {{
      function patchTodo(t) {{
        if (Object.prototype.hasOwnProperty.call(byId, t.id)) {{
          t.checked = byId[t.id];
        }}
      }}
      (plan.todos || []).forEach(patchTodo);
      (plan.backlog || []).forEach(patchTodo);
      function walk(sections) {{
        (sections || []).forEach(sec => {{
          (sec.todos || []).forEach(patchTodo);
          walk(sec.children);
        }});
      }}
      walk(plan.sections);
      plan.checked_count = (plan.todos || []).filter(t => t.checked).length;
      plan.todo_count = (plan.todos || []).length;
    }}

    function applySavedPlanStates(payload) {{
      Object.entries(payload.plans || {{}}).forEach(([studyId, entry]) => {{
        const plan = (PLANS_DATA.plans || {{}})[studyId];
        if (!plan) return;
        const byId = {{}};
        (entry.todos || []).forEach(t => {{ byId[t.id] = !!t.checked; }});
        applyTodoStatesToPlan(plan, byId);
      }});
    }}

    function setPlansSaveStatus(text, kind) {{
      const el = document.getElementById('plansSaveStatus');
      if (!el) return;
      el.textContent = text || '';
      el.classList.remove('ok', 'err');
      if (kind) el.classList.add(kind);
    }}

    async function savePlansToDisk() {{
      const btn = document.getElementById('btnPlansSave');
      const saveBase = (PLANS_DATA.save_url || 'http://127.0.0.1:8765').replace(/\\/$/, '');
      const payload = collectPlansSavePayload();
      const studyIds = Object.keys(payload.plans || {{}});
      if (!studyIds.length) {{
        setPlansSaveStatus('No plan todos to save.', 'err');
        return;
      }}
      if (btn) btn.disabled = true;
      setPlansSaveStatus('Saving…', '');
      try {{
        const res = await fetch(saveBase + '/save', {{
          method: 'POST',
          headers: {{ 'Content-Type': 'application/json' }},
          body: JSON.stringify(payload),
        }});
        const data = await res.json().catch(() => ({{}}));
        if (!res.ok || !data.ok) {{
          const msg = (data && data.error) ? data.error : ('Save failed (' + res.status + ')');
          setPlansSaveStatus(msg + ' — reopen dashboard from Memory Palace [D].', 'err');
          return;
        }}
        const changed = data.total_changed != null ? data.total_changed : 0;
        applySavedPlanStates(payload);
        clearSavedPlanProgress(studyIds);
        const studyId = sel ? sel.value : '';
        if (dashboardView === 'plans') renderPlansForStudy(studyId);
        setPlansSaveStatus(
          'Saved ' + changed + ' checkbox line(s) to plan files. Push [P] when ready.',
          'ok'
        );
      }} catch (e) {{
        setPlansSaveStatus(
          'Save server unavailable — reopen dashboard from Memory Palace [D].',
          'err'
        );
      }} finally {{
        if (btn) btn.disabled = false;
      }}
    }}

    function mdInline(text) {{
      let s = esc(text);
      s = s.replace(/\\[([^\\]]+)\\]\\(([^)]+)\\)/g, '<a href="$2" target="_blank" rel="noopener">$1</a>');
      s = s.replace(/(https?:\\/\\/[^\\s<]+)/g, '<a href="$1" target="_blank" rel="noopener">$1</a>');
      return s;
    }}

    function resourceIconSvg(kind) {{
      if (kind === 'youtube') {{
        return '<span class="plan-source-icon" aria-hidden="true">'
          + '<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">'
          + '<rect x="2" y="5" width="20" height="14" rx="3" fill="#FF0000"/>'
          + '<path d="M10 9.5v5l5-2.5-5-2.5z" fill="#fff"/>'
          + '</svg></span>';
      }}
      const emoji = {{
        book: '📖', docs: '📄', article: '📝', search: '🔍', link: '🔗'
      }}[kind] || '🔗';
      return '<span class="plan-source-icon" aria-hidden="true">' + emoji + '</span>';
    }}

    function normalizeResources(resources) {{
      if (!resources || !resources.length) return [];
      return resources.map(r => {{
        if (typeof r === 'string') {{
          const m = r.match(/\\[([^\\]]*)\\]\\(([^)]+)\\)/);
          const title = m ? m[1] : r;
          const url = m ? m[2] : '';
          return {{ id: '', title: title, url: url, kind: 'link', icon: 'link', line: r }};
        }}
        return r;
      }});
    }}

    function renderPlanSources(section, studyId) {{
      const resources = normalizeResources(section.resources || []);
      const sectionPath = section.section_path || '';
      const hasTodos = (section.todos || []).length > 0;
      const hasChildren = (section.children || []).length > 0;
      const hasResources = resources.length > 0;
      const showFooter = sectionPath && (hasTodos || hasResources || hasChildren);
      if (!showFooter) return '';

      let html = '<div class="plan-sources plan-sources-footer" data-section-path="' + esc(sectionPath) + '">';
      html += '<div class="plan-sources-label">Sources</div>';
      if (!hasResources) {{
        html += '<span class="plan-source-empty">No sources</span>';
      }} else {{
        resources.forEach(r => {{
          const kind = r.kind || r.icon || 'link';
          const title = r.title || r.url || 'Link';
          const url = r.url || '';
          const rid = r.id || '';
          html += '<span class="plan-source-chip">';
          if (url) {{
            html += '<a class="plan-source-link" href="' + esc(url) + '" target="_blank" rel="noopener">'
              + resourceIconSvg(kind)
              + '<span class="plan-source-title">' + esc(title) + '</span></a>';
          }} else {{
            html += '<span class="plan-source-link">' + resourceIconSvg(kind)
              + '<span class="plan-source-title">' + esc(title) + '</span></span>';
          }}
          html += '<span class="plan-source-actions">'
            + '<button type="button" class="plan-source-edit" data-resource-id="' + esc(rid)
            + '" data-resource-title="' + esc(title) + '" data-resource-url="' + esc(url)
            + '" title="Edit source" aria-label="Edit source">✎</button>'
            + '<button type="button" class="plan-source-remove" data-resource-id="' + esc(rid)
            + '" title="Remove source" aria-label="Remove source">×</button>'
            + '</span></span>';
        }});
      }}
      html += ' <button type="button" class="plan-source-add" data-section-path="' + esc(sectionPath)
        + '">+ add source</button>';
      html += '<div class="plan-source-form hidden" data-form-section="' + esc(sectionPath) + '">'
        + '<input type="text" class="plan-source-title-input" placeholder="Title" autocomplete="off"/>'
        + '<input type="url" class="plan-source-url-input" placeholder="URL" autocomplete="off"/>'
        + '<button type="button" class="plan-source-save">Save</button>'
        + '<button type="button" class="plan-source-cancel">Cancel</button>'
        + '</div>';
      html += '</div>';
      return html;
    }}

    function renderPlanResources(lines) {{
      return renderPlanSources({{ resources: lines, section_path: '', todos: [], children: [] }}, '');
    }}

    function stripPlanTopicEmoji(text) {{
      return (text || '').toString().replace(/^\\s*(?:🟧|🟦)\\s*/u, '').trim();
    }}

    function renderPlanTodos(todos, plan) {{
      if (!todos || !todos.length) return '';
      const rows = todos.map(t => {{
        const checked = planChecked(plan, t);
        const doneCls = checked ? ' done' : '';
        return '<li class="plan-todo' + doneCls + '">'
          + '<input type="checkbox" id="' + esc(t.id) + '" data-todo-id="' + esc(t.id) + '"'
          + (checked ? ' checked' : '') + '/>'
          + '<label for="' + esc(t.id) + '">' + mdInline(stripPlanTopicEmoji(t.text)) + '</label></li>';
      }}).join('');
      return '<ul class="plan-todos">' + rows + '</ul>';
    }}

    function renderBacklogTodos(todos, plan) {{
      if (!todos || !todos.length) return '';
      const rows = todos.map(t => {{
        const checked = planChecked(plan, t);
        const doneCls = checked ? ' done' : '';
        return '<li class="plan-todo' + doneCls + '">'
          + '<input type="checkbox" id="' + esc(t.id) + '" data-todo-id="' + esc(t.id) + '"'
          + (checked ? ' checked' : '') + '/>'
          + '<label for="' + esc(t.id) + '">' + mdInline(stripPlanTopicEmoji(t.text)) + '</label>'
          + '<button type="button" class="plan-todo-remove" data-backlog-remove="' + esc(t.id)
          + '" title="Remove from backlog" aria-label="Remove from backlog">×</button></li>';
      }}).join('');
      return '<ul class="plan-todos">' + rows + '</ul>';
    }}

    function headingTag(level) {{
      const n = Math.min(Math.max(level, 2), 6);
      return 'h' + n;
    }}

    function renderPlanSection(section, plan, studyId) {{
      const lvl = section.level || 2;
      const tag = headingTag(lvl);
      const hasTodos = (section.todos || []).length > 0;
      const topicCls = hasTodos ? ' plan-topic-card' : '';
      let html = '<section class="plan-section level-' + lvl + topicCls + '" id="' + esc(section.anchor || '') + '">'
        + '<' + tag + '>' + esc(section.title || '') + '</' + tag + '>'
        + '<div class="plan-topic-body">'
        + renderPlanTodos(section.todos || [], plan);
      (section.children || []).forEach(child => {{
        html += renderPlanSection(child, plan, studyId);
      }});
      html += '</div>'
        + renderPlanSources(section, studyId);
      html += '</section>';
      return html;
    }}

    function loadTocNodeState() {{
      try {{
        const raw = localStorage.getItem(PLANS_TOC_NODES_KEY);
        return raw ? JSON.parse(raw) : {{}};
      }} catch (e) {{
        return {{}};
      }}
    }}

    function saveTocNodeState(state) {{
      try {{
        localStorage.setItem(PLANS_TOC_NODES_KEY, JSON.stringify(state));
      }} catch (e) {{}}
    }}

    function isTocNodeCollapsed(anchor, level, hasChildren) {{
      if (!hasChildren) return false;
      const state = loadTocNodeState();
      if (Object.prototype.hasOwnProperty.call(state, anchor)) {{
        return !!state[anchor];
      }}
      // Default: keep Phases open; collapse Month+ topic nests for a cleaner TOC
      return level >= 3;
    }}

    function setTocNodeCollapsed(anchor, collapsed) {{
      const state = loadTocNodeState();
      state[anchor] = !!collapsed;
      saveTocNodeState(state);
    }}

    function renderTocNode(section) {{
      const children = section.children || [];
      const hasChildren = children.length > 0;
      const level = section.level || 2;
      const anchor = section.anchor || '';
      const collapsed = isTocNodeCollapsed(anchor, level, hasChildren);
      const twistCls = hasChildren ? 'toc-twist' : 'toc-twist leaf';
      const twistChar = hasChildren ? (collapsed ? '▶' : '▼') : '•';
      const liCls = 'toc-level-' + level + (collapsed ? ' toc-collapsed' : '');
      let html = '<li class="' + liCls + '" data-toc-anchor="' + esc(anchor) + '">'
        + '<div class="toc-row">'
        + '<button type="button" class="' + twistCls + '" data-toc-twist="' + esc(anchor) + '"'
        + ' aria-expanded="' + (collapsed ? 'false' : 'true') + '"'
        + (hasChildren ? '' : ' tabindex="-1"') + '>' + twistChar + '</button>'
        + '<a class="toc-link" href="#' + esc(anchor) + '">' + esc(section.title || '') + '</a>'
        + '</div>';
      if (hasChildren) {{
        html += '<ul>' + children.map(renderTocNode).join('') + '</ul>';
      }}
      html += '</li>';
      return html;
    }}

    function renderPlansToc(plan) {{
      const list = document.getElementById('plansTocList');
      if (!list) return;
      if (!plan) {{
        list.innerHTML = '';
        return;
      }}
      let html = '<li class="toc-level-2">'
        + '<div class="toc-row">'
        + '<button type="button" class="toc-twist leaf" tabindex="-1">•</button>'
        + '<a class="toc-link" href="#plan-backlog">Backlog</a>'
        + '</div></li>';
      (plan.sections || []).forEach(sec => {{
        html += renderTocNode(sec);
      }});
      list.innerHTML = html || '<li class="empty">No sections</li>';
      list.querySelectorAll('[data-toc-twist]').forEach(btn => {{
        if (btn.classList.contains('leaf')) return;
        btn.addEventListener('click', (e) => {{
          e.preventDefault();
          e.stopPropagation();
          const anchor = btn.getAttribute('data-toc-twist');
          const li = btn.closest('li');
          if (!li) return;
          const next = !li.classList.contains('toc-collapsed');
          li.classList.toggle('toc-collapsed', next);
          btn.textContent = next ? '▶' : '▼';
          btn.setAttribute('aria-expanded', next ? 'false' : 'true');
          setTocNodeCollapsed(anchor, next);
        }});
      }});
    }}

    function updatePlansProgress(plan) {{
      const bar = document.getElementById('plansProgressBar');
      const label = document.getElementById('plansProgressLabel');
      if (!plan) {{
        if (bar) bar.style.setProperty('--pct', '0%');
        if (label) label.textContent = '0 / 0 complete';
        if (bar) bar.setAttribute('aria-valuenow', '0');
        return;
      }}
      const todos = plan.todos || [];
      const total = todos.length;
      let done = 0;
      todos.forEach(t => {{ if (planChecked(plan, t)) done += 1; }});
      const pct = total ? Math.round((done / total) * 100) : 0;
      if (bar) {{
        bar.style.setProperty('--pct', pct + '%');
        bar.setAttribute('aria-valuenow', String(pct));
      }}
      if (label) label.textContent = done + ' / ' + total + ' complete';
    }}

    function renderPlansForStudy(studyId) {{
      const plan = (PLANS_DATA.plans || {{}})[studyId];
      const titleEl = document.getElementById('plansTitle');
      const subEl = document.getElementById('plansSub');
      const content = document.getElementById('plansContent');
      const empty = document.getElementById('plansEmpty');
      const github = document.getElementById('plansGithub');
      if (github && PLANS_DATA.github_url) github.href = PLANS_DATA.github_url;
      if (!plan) {{
        if (titleEl) titleEl.textContent = 'Study Plan';
        if (subEl) subEl.textContent = 'No plan file for this study';
        if (content) content.innerHTML = '';
        if (empty) empty.classList.remove('hidden');
        renderPlansToc(null);
        updatePlansProgress(null);
        return;
      }}
      if (empty) empty.classList.add('hidden');
      if (titleEl) titleEl.textContent = plan.title || plan.study_title || 'Study Plan';
      if (subEl) subEl.textContent = (plan.todo_count || 0) + ' tasks · source ' + (plan.source_rel || '');
      const backlog = plan.backlog || [];
      let html = '<div class="plans-backlog" id="plan-backlog"><h3>Backlog</h3>';
      if (backlog.length) {{
        html += renderBacklogTodos(backlog, plan);
      }} else {{
        html += '<p class="plans-backlog-empty">No backlog items yet.</p>';
      }}
      html += '<div class="plans-backlog-add">'
        + '<input type="text" id="plansBacklogInput" placeholder="Add backlog item…" autocomplete="off"/>'
        + '<button type="button" id="btnPlansBacklogAdd">Add</button>'
        + '</div></div>';
      (plan.sections || []).forEach(sec => {{
        html += renderPlanSection(sec, plan, studyId);
      }});
      if (content) {{
        content.innerHTML = html;
        content.querySelectorAll('input[type="checkbox"][data-todo-id]').forEach(box => {{
          box.addEventListener('change', () => {{
            setPlanChecked(box.getAttribute('data-todo-id'), box.checked);
            const row = box.closest('.plan-todo');
            if (row) row.classList.toggle('done', box.checked);
            updatePlansProgress(plan);
          }});
        }});
        content.querySelectorAll('[data-backlog-remove]').forEach(btn => {{
          btn.addEventListener('click', () => {{
            const todoId = btn.getAttribute('data-backlog-remove');
            removeBacklogItem(studyId, todoId);
          }});
        }});
        content.querySelectorAll('.plan-source-add').forEach(btn => {{
          btn.addEventListener('click', () => {{
            const wrap = btn.closest('.plan-sources');
            const form = wrap ? wrap.querySelector('.plan-source-form') : null;
            if (form) {{
              form.classList.remove('hidden');
              form.removeAttribute('data-edit-id');
              form.querySelectorAll('input').forEach(inp => {{ inp.value = ''; }});
              const ti = form.querySelector('.plan-source-title-input');
              if (ti) ti.focus();
            }}
          }});
        }});
        content.querySelectorAll('.plan-source-form .plan-source-cancel').forEach(btn => {{
          btn.addEventListener('click', () => {{
            const form = btn.closest('.plan-source-form');
            if (form) {{
              form.classList.add('hidden');
              form.removeAttribute('data-edit-id');
              form.querySelectorAll('input').forEach(inp => {{ inp.value = ''; }});
            }}
          }});
        }});
        content.querySelectorAll('.plan-source-form .plan-source-save').forEach(btn => {{
          btn.addEventListener('click', () => {{
            const form = btn.closest('.plan-source-form');
            if (!form) return;
            const sp = form.getAttribute('data-form-section') || '';
            const title = (form.querySelector('.plan-source-title-input') || {{}}).value || '';
            const url = (form.querySelector('.plan-source-url-input') || {{}}).value || '';
            const editId = form.getAttribute('data-edit-id') || '';
            if (editId) {{
              editResource(studyId, editId, title.trim(), url.trim());
            }} else {{
              addResource(studyId, sp, title.trim(), url.trim());
            }}
          }});
        }});
        content.querySelectorAll('.plan-source-edit').forEach(btn => {{
          btn.addEventListener('click', () => {{
            const wrap = btn.closest('.plan-sources');
            const form = wrap ? wrap.querySelector('.plan-source-form') : null;
            if (!form) return;
            form.classList.remove('hidden');
            form.setAttribute('data-edit-id', btn.getAttribute('data-resource-id') || '');
            const ti = form.querySelector('.plan-source-title-input');
            const ui = form.querySelector('.plan-source-url-input');
            if (ti) ti.value = btn.getAttribute('data-resource-title') || '';
            if (ui) ui.value = btn.getAttribute('data-resource-url') || '';
            if (ti) ti.focus();
          }});
        }});
        content.querySelectorAll('.plan-source-remove').forEach(btn => {{
          btn.addEventListener('click', () => {{
            removeResource(studyId, btn.getAttribute('data-resource-id') || '');
          }});
        }});
        const addBtn = document.getElementById('btnPlansBacklogAdd');
        const addInput = document.getElementById('plansBacklogInput');
        const submitBacklog = () => addBacklogItem(studyId, plan);
        if (addBtn) addBtn.addEventListener('click', submitBacklog);
        if (addInput) {{
          addInput.addEventListener('keydown', (e) => {{
            if (e.key === 'Enter') {{
              e.preventDefault();
              submitBacklog();
            }}
          }});
        }}
      }}
      renderPlansToc(plan);
      updatePlansProgress(plan);
    }}

    function backlogActionError(data, res, verb) {{
      let msg = (data && data.error) ? data.error : '';
      if (!msg && data && data.ok === false && Array.isArray(data.results) && !data.results.length) {{
        msg = 'Save server outdated — press [D] again';
      }}
      if (!msg) msg = verb + ' failed (' + res.status + ')';
      return msg + ' — reopen dashboard from Memory Palace [D].';
    }}

    async function addBacklogItem(studyId, plan) {{
      const input = document.getElementById('plansBacklogInput');
      const btn = document.getElementById('btnPlansBacklogAdd');
      const text = input ? input.value.trim() : '';
      if (!text) {{
        setPlansSaveStatus('Enter a backlog item first.', 'err');
        if (input) input.focus();
        return;
      }}
      const saveBase = (PLANS_DATA.save_url || 'http://127.0.0.1:8765').replace(/\\/$/, '');
      if (btn) btn.disabled = true;
      setPlansSaveStatus('Adding backlog item…', '');
      try {{
        const res = await fetch(saveBase + '/save', {{
          method: 'POST',
          headers: {{ 'Content-Type': 'application/json' }},
          body: JSON.stringify({{ action: 'add_backlog', study_id: studyId, text }}),
        }});
        const data = await res.json().catch(() => ({{}}));
        if (!res.ok || !data.ok || !data.plan) {{
          setPlansSaveStatus(backlogActionError(data, res, 'Add'), 'err');
          return;
        }}
        PLANS_DATA.plans[studyId] = data.plan;
        if (input) input.value = '';
        renderPlansForStudy(studyId);
        setPlansSaveStatus('Backlog item added to plan file.', 'ok');
      }} catch (e) {{
        setPlansSaveStatus(
          'Save server unavailable — reopen dashboard from Memory Palace [D].',
          'err'
        );
      }} finally {{
        if (btn) btn.disabled = false;
      }}
    }}

    async function removeBacklogItem(studyId, todoId) {{
      if (!todoId) return;
      const saveBase = (PLANS_DATA.save_url || 'http://127.0.0.1:8765').replace(/\\/$/, '');
      setPlansSaveStatus('Removing backlog item…', '');
      try {{
        const res = await fetch(saveBase + '/save', {{
          method: 'POST',
          headers: {{ 'Content-Type': 'application/json' }},
          body: JSON.stringify({{
            action: 'remove_backlog',
            study_id: studyId,
            todo_id: todoId,
          }}),
        }});
        const data = await res.json().catch(() => ({{}}));
        if (!res.ok || !data.ok || !data.plan) {{
          setPlansSaveStatus(backlogActionError(data, res, 'Remove'), 'err');
          return;
        }}
        const store = loadPlanStorage();
        if (Object.prototype.hasOwnProperty.call(store, todoId)) {{
          delete store[todoId];
          savePlanStorage(store);
        }}
        PLANS_DATA.plans[studyId] = data.plan;
        renderPlansForStudy(studyId);
        setPlansSaveStatus('Backlog item removed from plan file.', 'ok');
      }} catch (e) {{
        setPlansSaveStatus(
          'Save server unavailable — reopen dashboard from Memory Palace [D].',
          'err'
        );
      }}
    }}

    async function addResource(studyId, sectionPath, title, url) {{
      if (!title || !url) {{
        setPlansSaveStatus('Enter both title and URL for the source.', 'err');
        return;
      }}
      const saveBase = (PLANS_DATA.save_url || 'http://127.0.0.1:8765').replace(/\\/$/, '');
      setPlansSaveStatus('Adding source…', '');
      try {{
        const res = await fetch(saveBase + '/save', {{
          method: 'POST',
          headers: {{ 'Content-Type': 'application/json' }},
          body: JSON.stringify({{
            action: 'add_resource',
            study_id: studyId,
            section_path: sectionPath,
            title: title,
            url: url,
          }}),
        }});
        const data = await res.json().catch(() => ({{}}));
        if (!res.ok || !data.ok || !data.plan) {{
          setPlansSaveStatus(backlogActionError(data, res, 'Add source'), 'err');
          return;
        }}
        PLANS_DATA.plans[studyId] = data.plan;
        renderPlansForStudy(studyId);
        setPlansSaveStatus('Source added to plan.', 'ok');
      }} catch (e) {{
        setPlansSaveStatus(
          'Save server unavailable — reopen dashboard from Memory Palace [D].',
          'err'
        );
      }}
    }}

    async function editResource(studyId, resourceId, title, url) {{
      if (!resourceId || !title || !url) {{
        setPlansSaveStatus('Enter both title and URL for the source.', 'err');
        return;
      }}
      const saveBase = (PLANS_DATA.save_url || 'http://127.0.0.1:8765').replace(/\\/$/, '');
      setPlansSaveStatus('Updating source…', '');
      try {{
        const res = await fetch(saveBase + '/save', {{
          method: 'POST',
          headers: {{ 'Content-Type': 'application/json' }},
          body: JSON.stringify({{
            action: 'edit_resource',
            study_id: studyId,
            resource_id: resourceId,
            title: title,
            url: url,
          }}),
        }});
        const data = await res.json().catch(() => ({{}}));
        if (!res.ok || !data.ok || !data.plan) {{
          setPlansSaveStatus(backlogActionError(data, res, 'Edit source'), 'err');
          return;
        }}
        PLANS_DATA.plans[studyId] = data.plan;
        renderPlansForStudy(studyId);
        setPlansSaveStatus('Source updated.', 'ok');
      }} catch (e) {{
        setPlansSaveStatus(
          'Save server unavailable — reopen dashboard from Memory Palace [D].',
          'err'
        );
      }}
    }}

    async function removeResource(studyId, resourceId) {{
      if (!resourceId) return;
      const saveBase = (PLANS_DATA.save_url || 'http://127.0.0.1:8765').replace(/\\/$/, '');
      setPlansSaveStatus('Removing source…', '');
      try {{
        const res = await fetch(saveBase + '/save', {{
          method: 'POST',
          headers: {{ 'Content-Type': 'application/json' }},
          body: JSON.stringify({{
            action: 'remove_resource',
            study_id: studyId,
            resource_id: resourceId,
          }}),
        }});
        const data = await res.json().catch(() => ({{}}));
        if (!res.ok || !data.ok || !data.plan) {{
          setPlansSaveStatus(backlogActionError(data, res, 'Remove source'), 'err');
          return;
        }}
        PLANS_DATA.plans[studyId] = data.plan;
        renderPlansForStudy(studyId);
        setPlansSaveStatus('Source removed from plan.', 'ok');
      }} catch (e) {{
        setPlansSaveStatus(
          'Save server unavailable — reopen dashboard from Memory Palace [D].',
          'err'
        );
      }}
    }}

    function ensureMermaid() {{
      if (mermaidReady) {{
        if (window.mermaid) window.mermaid.run();
        return;
      }}
      mermaidReady = true;
      const s = document.createElement('script');
      s.src = 'https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.min.js';
      s.onload = () => {{
        if (window.mermaid) {{
          window.mermaid.initialize({{ startOnLoad: false, theme: 'dark' }});
          window.mermaid.run();
        }}
      }};
      document.head.appendChild(s);
    }}

    function renderCharacters(q) {{
      const el = document.getElementById('charResults');
      if (!el) return;
      const query = (q || '').trim().toLowerCase();
      const cards = [];
      (METHOD_CANON.characters || []).forEach(sec => {{
        (sec.characters || []).forEach(c => {{
          const name = (c.name || '').toString();
          if (query && !name.toLowerCase().includes(query)) return;
          cards.push(
            '<div class="canon-card"><span class="sec">' + esc(sec.title || '') + '</span>'
            + esc(name) + '</div>'
          );
        }});
      }});
      if (!cards.length) {{
        el.innerHTML = '<p class="empty">No characters match.</p>';
        return;
      }}
      el.innerHTML = cards.slice(0, 200).join('');
    }}

    function renderBestiary(q) {{
      const el = document.getElementById('beastResults');
      if (!el) return;
      const query = (q || '').trim().toLowerCase();
      const rows = [];
      (METHOD_CANON.bestiary || []).forEach(b => {{
        const code = (b.code || '').toString();
        const name = (b.name || '').toString();
        const source = (b.source || '').toString();
        const blob = (code + ' ' + name + ' ' + source).toLowerCase();
        if (query && !blob.includes(query)) return;
        rows.push(
          '<tr><td>' + esc(code) + '</td><td>' + esc(name) + '</td><td>' + esc(source) + '</td></tr>'
        );
      }});
      if (!rows.length) {{
        el.innerHTML = '<p class="empty">No beasts match.</p>';
        return;
      }}
      const limit = query ? 200 : 80;
      el.innerHTML = '<table class="canon-table"><thead><tr><th>Code</th><th>Name</th><th>Source</th></tr></thead><tbody>'
        + rows.slice(0, limit).join('')
        + '</tbody></table>'
        + (rows.length > limit
          ? '<p class="method-note">Showing ' + limit + ' of ' + rows.length + ' — refine search to narrow.</p>'
          : '');
    }}

    if (btnMethod) {{
      btnMethod.addEventListener('click', () => setDashboardView('toggle-method'));
    }}
    if (btnPlans) {{
      btnPlans.addEventListener('click', () => setDashboardView('toggle-plans'));
    }}
    const btnPlansReset = document.getElementById('btnPlansReset');
    const btnPlansSave = document.getElementById('btnPlansSave');
    if (btnPlansSave) {{
      btnPlansSave.addEventListener('click', () => savePlansToDisk());
    }}
    if (btnPlansReset) {{
      btnPlansReset.addEventListener('click', () => {{
        const studyId = sel ? sel.value : '';
        const plan = (PLANS_DATA.plans || {{}})[studyId];
        if (!plan || !plan.slug) return;
        resetPlanProgress(plan.slug);
        renderPlansForStudy(studyId);
      }});
    }}
    const charSearch = document.getElementById('charSearch');
    const beastSearch = document.getElementById('beastSearch');
    if (charSearch) {{
      charSearch.addEventListener('input', () => {{
        clearTimeout(searchTimer);
        searchTimer = setTimeout(() => renderCharacters(charSearch.value), 80);
      }});
    }}
    if (beastSearch) {{
      beastSearch.addEventListener('input', () => {{
        clearTimeout(searchTimer);
        searchTimer = setTimeout(() => renderBestiary(beastSearch.value), 80);
      }});
    }}
    document.querySelectorAll('.btn-prompt-full').forEach(btn => {{
      btn.addEventListener('click', () => {{
        const det = btn.closest('.method-prompt');
        if (!det) return;
        const preview = det.querySelector('.prompt-preview');
        const full = det.querySelector('.prompt-full');
        if (!preview || !full) return;
        const showing = !full.hidden;
        full.hidden = showing;
        preview.hidden = !showing;
        btn.textContent = showing ? 'Show all' : 'Show less';
      }});
    }});

    document.addEventListener('keydown', (e) => {{
      if (handleStudyModalKey(e)) return;
      if (e.key === 'Escape' && overlay.classList.contains('open')) {{
        e.preventDefault();
        closeOverlay();
        return;
      }}
      if (overlay.classList.contains('open')) {{
        if (e.key === 'ArrowLeft') {{
          e.preventDefault();
          stepPalace(1);
          return;
        }}
        if (e.key === 'ArrowRight') {{
          e.preventDefault();
          stepPalace(-1);
          return;
        }}
        if ((e.key === 'c' || e.key === 'C') && currentOverlayPalaceId
            && !e.ctrlKey && !e.metaKey && !e.altKey) {{
          const st = PALACE_DATA[currentOverlayPalaceId];
          const prompt = st ? (st.image_prompt || '').toString().trim() : '';
          if (!prompt) return;
          e.preventDefault();
          copyPrompt(st.image_prompt, btnCopyPrompt);
          return;
        }}
      }}
      const tag = (e.target && e.target.tagName) ? e.target.tagName.toLowerCase() : '';
      if (tag === 'input' || tag === 'textarea' || tag === 'select' || (e.target && e.target.isContentEditable))
        return;
      if ((e.key === 'p' || e.key === 'P') && !overlay.classList.contains('open')
          && !e.ctrlKey && !e.metaKey && !e.altKey) {{
        e.preventDefault();
        setDashboardView('toggle-plans');
        return;
      }}
      if ((e.key === 'm' || e.key === 'M') && !overlay.classList.contains('open')
          && !e.ctrlKey && !e.metaKey && !e.altKey) {{
        e.preventDefault();
        setDashboardView('toggle-method');
        return;
      }}
      if (e.key === 'l' || e.key === 'L') {{
        if (dashboardView !== 'practice') setDashboardView('practice');
        e.preventDefault();
        openLatestPalace();
      }}
    }});

    function runSearch(q) {{
      const query = (q || '').trim().toLowerCase();
      if (!query) {{
        searchResults.classList.remove('open');
        searchResults.innerHTML = '';
        return;
      }}
      const studyId = sel ? sel.value : '';
      const hits = [];
      Object.values(PALACE_DATA).forEach(st => {{
        if (studyId && st.study_id !== studyId) return;
        (st.atoms || []).forEach(a => {{
          const blob = [a.beast, a.concept, a.quote, a.story, a.sensory, a.zone, a.zone_label]
            .map(x => (x || '').toString().toLowerCase()).join(' ');
          if (blob.includes(query)) {{
            hits.push({{ palaceId: st.id, palaceNumber: st.number, atom: a }});
          }}
        }});
      }});
      searchResults.classList.add('open');
      if (!hits.length) {{
        searchResults.innerHTML = '<p class="empty">No Knowledge Atoms found.</p>';
        return;
      }}
      searchResults.innerHTML = hits.slice(0, 40).map(h => {{
        const concept = (h.atom.concept || '').toString();
        const snip = concept.length > 90 ? concept.slice(0, 87) + '…' : concept;
        return '<button type="button" class="hit" data-palace-id="' + esc(h.palaceId)
          + '" data-atom-id="' + esc(h.atom.id) + '">'
          + '<span class="meta">Memory Palace ' + esc(h.palaceNumber) + ' · ' + dash(h.atom.beast) + '</span><br/>'
          + dash(snip || h.atom.quote || h.atom.story)
          + '</button>';
      }}).join('');
      searchResults.querySelectorAll('.hit').forEach(btn => {{
        btn.addEventListener('click', () => {{
          openPalace(btn.dataset.palaceId, btn.dataset.atomId);
        }});
      }});
    }}

    atomSearch.addEventListener('input', () => {{
      clearTimeout(searchTimer);
      searchTimer = setTimeout(() => runSearch(atomSearch.value), 120);
    }});

    openStudyModal();
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
    p.add_argument("--technique-source", type=Path, default=None)
    p.add_argument("--technique-dir", type=Path, default=None)
    p.add_argument("--skip-technique-sync", action="store_true")
    p.add_argument("--studies-root", type=Path, default=None)
    p.add_argument("--skip-plans-sync", action="store_true")
    args = p.parse_args(argv)

    out_dir = args.output_dir.resolve()
    studies_root = (
        args.studies_root.resolve() if args.studies_root else default_studies_root()
    )
    technique_dir = (
        args.technique_dir.resolve() if args.technique_dir else default_technique_dir()
    )

    if not args.skip_technique_sync:
        src = resolve_technique_source(
            args.technique_source,
            args.notes_root.resolve() if args.notes_root else None,
        )
        if src and src.is_dir():
            try:
                copied, skipped = sync_technique(src, technique_dir)
                print(f"Technique sync: {copied} copied, {skipped} skipped")
            except OSError as e:
                print(f"Technique sync warning: {e}", file=sys.stderr)
        else:
            print(
                "Technique sync skipped (source missing); using existing mirror if any",
                file=sys.stderr,
            )

    method_html, method_canon = build_method_panel(technique_dir)

    if not args.skip_plans_sync:
        try:
            entries = sync_plans_all(
                args.data_dir.resolve(),
                studies_root,
                out_dir,
            )
            print(f"Plans sync: {len(entries)} plan(s)")
        except OSError as e:
            print(f"Plans sync warning: {e}", file=sys.stderr)

    plans_payload = build_plans_payload(studies_root, args.data_dir.resolve())

    snap = snapshot(
        data_dir=args.data_dir.resolve(),
        notes_root=args.notes_root.resolve() if args.notes_root else None,
        study_id=None,
        output_dir=out_dir,
    )
    if args.study_id:
        snap["selected_study_id"] = args.study_id
    elif snap.get("studies"):
        snap["selected_study_id"] = snap["studies"][0]["id"]

    html_out = build_html(
        snap,
        method_html=method_html,
        method_canon=method_canon,
        plans_payload=plans_payload,
    )
    out_dir.mkdir(parents=True, exist_ok=True)
    out = out_dir / "dashboard.html"
    out.write_text(html_out, encoding="utf-8")
    print(f"Wrote {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
