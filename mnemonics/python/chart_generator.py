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


def file_uri(path: str) -> str:
    p = Path(path).resolve()
    return "file:///" + quote(p.as_posix(), safe="/:")


def build_html(snap: dict) -> str:
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
      display: flex;
      flex-direction: row;
      flex-wrap: wrap;
      gap: 1.75rem;
      align-items: stretch;
    }}
    #ovAtoms > .empty {{
      flex: 1 1 100%;
      margin: 0;
    }}
    .atom-card {{
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 10px;
      padding: 0.95rem 1.05rem;
      scroll-margin-top: 5rem;
      flex: 1 1 200px;
      min-width: min(200px, 100%);
      max-width: 100%;
      box-sizing: border-box;
    }}
    .atom-card.highlight {{
      border-color: var(--gold);
      box-shadow: 0 0 0 1px var(--gold);
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
    .atom-card .beast-name {{
      font-size: 1.12rem;
      font-weight: 650;
      line-height: 1.35;
    }}
    .atom-card .field-concept {{
      margin-top: 0.65rem;
      padding: 0.65rem 0.75rem;
      background: #242830;
      border: 1px solid var(--line);
      border-left: 3px solid var(--gold);
      border-radius: 8px;
    }}
    .atom-card .emoji {{
      margin-right: 0.28rem;
    }}
    .atom-card .zone-tag {{
      color: var(--muted);
      font-size: 0.85rem;
      margin: 0 0 0.35rem;
    }}
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
      <div>
        <label for="atomSearch">Search Knowledge Atoms</label>
        <input type="search" id="atomSearch" placeholder="Beast, concept, quote, story…" autocomplete="off"/>
      </div>
    </div>
  </header>
  <div id="searchResults" aria-live="polite"></div>
  <main>
    {''.join(study_blocks)}
  </main>
  <footer>Click a Memory Palace for fullscreen practice. Latest palace (or L) opens the highest palace number.</footer>

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
      <button type="button" id="btnClose">Close</button>
    </div>
    <div class="overlay-body">
      <div class="overlay-image" id="ovImage"></div>
      <div class="practice">
        <h3>Practice — Knowledge Atoms</h3>
        <div id="ovAtoms"></div>
      </div>
      <div class="overlay-prompt" id="ovPrompt"></div>
    </div>
  </div>

  <script>
    const PALACE_DATA = {palace_json};
    const STUDY_LATEST = {latest_json};
    const STUDIES = {studies_json};
    const STUDY_PAGE_SIZE = 8;
    const sel = document.getElementById('study');
    const sections = [...document.querySelectorAll('.study')];
    const overlay = document.getElementById('overlay');
    const ovTitle = document.getElementById('ovTitle');
    const ovCount = document.getElementById('ovCount');
    const ovImage = document.getElementById('ovImage');
    const ovPrompt = document.getElementById('ovPrompt');
    const ovAtoms = document.getElementById('ovAtoms');
    const btnClose = document.getElementById('btnClose');
    const btnLatest = document.getElementById('btnLatest');
    const atomSearch = document.getElementById('atomSearch');
    const searchResults = document.getElementById('searchResults');
    const studyModal = document.getElementById('studyModal');
    const studyModalTitle = document.getElementById('studyModalTitle');
    const studyModalHint = document.getElementById('studyModalHint');
    const studyModalList = document.getElementById('studyModalList');
    const studyModalFoot = document.getElementById('studyModalFoot');
    let searchTimer = null;
    let studyPickerTier = 'group'; // 'group' | 'study'
    let studyPickerGroup = 0;

    function show(id) {{
      sections.forEach(s => {{
        s.style.display = (s.dataset.studyId === id) ? 'block' : 'none';
      }});
    }}

    function selectStudy(id, {{ focusSelect = true }} = {{}}) {{
      if (!sel || !id) return;
      sel.value = id;
      show(id);
      sel.dispatchEvent(new Event('change', {{ bubbles: true }}));
      if (focusSelect) {{
        requestAnimationFrame(() => sel.focus());
      }}
    }}

    if (sel) {{
      sel.addEventListener('change', () => show(sel.value));
      if (sel.value) show(sel.value);
    }}

    function studyGroups() {{
      const groups = [];
      for (let i = 0; i < STUDIES.length; i += STUDY_PAGE_SIZE) {{
        groups.push(STUDIES.slice(i, i + STUDY_PAGE_SIZE));
      }}
      return groups;
    }}

    function studyModalOpen() {{
      return studyModal && studyModal.classList.contains('open');
    }}

    function closeStudyModal({{ focusSelect = true }} = {{}}) {{
      if (!studyModal) return;
      studyModal.classList.remove('open');
      studyModal.setAttribute('aria-hidden', 'true');
      if (focusSelect && sel) requestAnimationFrame(() => sel.focus());
    }}

    function openStudyModal() {{
      if (!studyModal || !STUDIES.length) {{
        if (sel) requestAnimationFrame(() => sel.focus());
        return;
      }}
      studyPickerTier = 'group';
      studyPickerGroup = 0;
      studyModal.classList.add('open');
      studyModal.setAttribute('aria-hidden', 'false');
      renderStudyModal();
    }}

    function renderStudyModal() {{
      const groups = studyGroups();
      studyModalList.innerHTML = '';
      if (studyPickerTier === 'group') {{
        studyModalTitle.textContent = 'Select a study group';
        studyModalHint.textContent = 'Press a number to open a group.';
        studyModalFoot.textContent = 'Esc keeps the current study';
        groups.forEach((g, idx) => {{
          const n = String(idx + 1);
          const first = g[0] ? g[0].title : '';
          const last = g[g.length - 1] ? g[g.length - 1].title : '';
          const label = (first === last)
            ? first
            : (first + ' — ' + last);
          const li = document.createElement('li');
          const btn = document.createElement('button');
          btn.type = 'button';
          btn.innerHTML = '<span class="key">' + esc(n) + '</span>'
            + '<span class="pick-label">' + esc(label) + ' <span style="color:var(--muted)">('
            + g.length + ')</span></span>';
          btn.addEventListener('click', () => pickStudyGroup(idx));
          li.appendChild(btn);
          studyModalList.appendChild(li);
        }});
      }} else {{
        const g = groups[studyPickerGroup] || [];
        studyModalTitle.textContent = 'Select a study';
        studyModalHint.textContent = 'Press a letter to choose a study.';
        studyModalFoot.textContent = 'Esc returns to groups';
        g.forEach((st, idx) => {{
          const letter = String.fromCharCode(97 + idx); // a, b, c...
          const li = document.createElement('li');
          const btn = document.createElement('button');
          btn.type = 'button';
          btn.innerHTML = '<span class="key">' + esc(letter.toUpperCase()) + '</span>'
            + '<span class="pick-label">' + esc(st.title) + '</span>';
          btn.addEventListener('click', () => pickStudyItem(st.id));
          li.appendChild(btn);
          studyModalList.appendChild(li);
        }});
      }}
      const firstBtn = studyModalList.querySelector('button');
      if (firstBtn) requestAnimationFrame(() => firstBtn.focus());
    }}

    function pickStudyGroup(idx) {{
      const groups = studyGroups();
      if (idx < 0 || idx >= groups.length) return;
      studyPickerGroup = idx;
      studyPickerTier = 'study';
      renderStudyModal();
    }}

    function pickStudyItem(id) {{
      selectStudy(id, {{ focusSelect: true }});
      closeStudyModal({{ focusSelect: true }});
    }}

    function handleStudyModalKey(e) {{
      if (!studyModalOpen()) return false;
      if (e.key === 'Escape') {{
        e.preventDefault();
        if (studyPickerTier === 'study') {{
          studyPickerTier = 'group';
          renderStudyModal();
        }} else {{
          closeStudyModal({{ focusSelect: true }});
        }}
        return true;
      }}
      if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {{
        const ch = e.key.toLowerCase();
        if (studyPickerTier === 'group') {{
          if (ch >= '1' && ch <= '9') {{
            e.preventDefault();
            pickStudyGroup(parseInt(ch, 10) - 1);
            return true;
          }}
        }} else if (ch >= 'a' && ch <= 'z') {{
          const idx = ch.charCodeAt(0) - 97;
          const g = studyGroups()[studyPickerGroup] || [];
          if (idx >= 0 && idx < g.length) {{
            e.preventDefault();
            pickStudyItem(g[idx].id);
            return true;
          }}
        }}
      }}
      return true; // swallow other keys while modal is open
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
          const prev = btn.textContent;
          btn.textContent = 'Copied';
          setTimeout(() => {{ btn.textContent = prev; }}, 1200);
        }}
      }} catch (e) {{
        if (btn) btn.textContent = 'Copy failed';
      }}
    }}

    function renderPromptBlock(prompt, id) {{
      const t = (prompt || '').toString().trim();
      if (!t) return '<p class="prompt-empty">No image prompt saved.</p>';
      return '<div class="prompt-head"><span>Image prompt</span>'
        + '<button type="button" class="btn-copy" data-copy-id="' + esc(id) + '">Copy prompt</button></div>'
        + '<pre class="prompt-text">' + esc(prompt) + '</pre>';
    }}

    function openPalace(id, focusAtomId) {{
      const st = PALACE_DATA[id];
      if (!st) return;
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
      const atoms = st.atoms || [];
      if (!atoms.length) {{
        ovAtoms.innerHTML = '<p class="empty">No Knowledge Atoms on this Memory Palace.</p>';
      }} else {{
        ovAtoms.innerHTML = atoms.map(a => {{
          const zone = (a.zone || a.zone_label)
            ? '<p class="zone-tag">' + dash(a.zone) + (a.zone_label ? ' · ' + dash(a.zone_label) : '') + '</p>'
            : '';
          const hl = (focusAtomId && a.id === focusAtomId) ? ' highlight' : '';
          return '<article class="atom-card' + hl + '" data-atom-id="' + esc(a.id) + '">'
            + zone
            + '<p class="field field-beast"><span class="lbl">Beast</span>'
            + '<span class="beast-name">' + dash(a.beast) + '</span></p>'
            + '<p class="field field-concept"><span class="lbl">Concept</span>'
            + formatConcept(a.concept) + '</p>'
            + '<p class="field"><span class="lbl">Quote</span>' + formatQuote(a.quote) + '</p>'
            + '<p class="field"><span class="lbl">Story</span>' + dash(a.story) + '</p>'
            + '<p class="field"><span class="lbl">Sensory</span>' + formatSensory(a.sensory) + '</p>'
            + '</article>';
        }}).join('');
      }}
      overlay.classList.add('open');
      overlay.setAttribute('aria-hidden', 'false');
      btnClose.focus();
      if (focusAtomId) {{
        requestAnimationFrame(() => {{
          const el = ovAtoms.querySelector('[data-atom-id="' + CSS.escape(focusAtomId) + '"]');
          if (el) el.scrollIntoView({{ behavior: 'smooth', block: 'center' }});
        }});
      }}
    }}

    function closeOverlay() {{
      overlay.classList.remove('open');
      overlay.setAttribute('aria-hidden', 'true');
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
    function openLatestPalace() {{
      const studyId = sel ? sel.value : '';
      const palaceId = STUDY_LATEST[studyId];
      if (palaceId) openPalace(palaceId);
    }}
    btnLatest.addEventListener('click', openLatestPalace);
    document.addEventListener('keydown', (e) => {{
      if (handleStudyModalKey(e)) return;
      if (e.key === 'Escape' && overlay.classList.contains('open')) {{
        e.preventDefault();
        closeOverlay();
        return;
      }}
      const tag = (e.target && e.target.tagName) ? e.target.tagName.toLowerCase() : '';
      if (tag === 'input' || tag === 'textarea' || tag === 'select' || (e.target && e.target.isContentEditable))
        return;
      if (e.key === 'l' || e.key === 'L') {{
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
    args = p.parse_args(argv)

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
