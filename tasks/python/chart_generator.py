"""Build Tasks dashboard HTML from CSV data."""

from __future__ import annotations

import argparse
import csv
import html
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

RECURRENCE_ORDER = [
    "daily",
    "weekly",
    "monthly",
    "quarterly",
    "biannual",
    "yearly",
    "every_2y",
    "every_3y",
    "every_5y",
    "every_10y",
]

RECURRENCE_LABELS = {
    "daily": "Daily",
    "weekly": "Weekly",
    "monthly": "Monthly",
    "quarterly": "Quarterly",
    "biannual": "Bi-yearly",
    "yearly": "Yearly",
    "every_2y": "Every 2 years",
    "every_3y": "Every 3 years",
    "every_5y": "Every 5 years",
    "every_10y": "Every 10 years",
    "": "Other",
}


def read_csv(path: Path) -> list[dict]:
    if not path.exists():
        return []
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def esc(s: str) -> str:
    return html.escape(s or "")


def sort_key_order_title(row: dict) -> tuple:
    try:
        so = int(row.get("sort_order") or 0)
    except ValueError:
        so = 0
    return (so, (row.get("title") or "").lower())


def build_html(data_dir: Path) -> str:
    projects = read_csv(data_dir / "projects.csv")
    tasks = read_csv(data_dir / "tasks.csv")
    infos = read_csv(data_dir / "info_points.csv")
    attachments = read_csv(data_dir / "attachments.csv")

    today = datetime.now().strftime("%Y-%m-%d")

    active_projects = [p for p in projects if p.get("active", "1") != "0"]
    active_tasks = [t for t in tasks if t.get("active", "1") != "0"]

    open_punctual = [
        t
        for t in active_tasks
        if t.get("kind") != "habitual" and t.get("emoji") != "✅"
    ]
    habitual = [t for t in active_tasks if t.get("kind") == "habitual"]

    def is_due(t: dict) -> bool:
        due = (t.get("next_due") or t.get("due_date") or "").strip()
        return (not due) or due <= today

    habits_due_n = sum(1 for t in habitual if is_due(t))
    open_personal = [t for t in open_punctual if t.get("filter") == "personal"]
    open_work = [t for t in open_punctual if t.get("filter") == "work"]

    by_emoji = Counter(t.get("emoji") or "🔲" for t in open_punctual)

    tasks_by_project: dict[str, list[dict]] = defaultdict(list)
    for t in open_punctual:
        tasks_by_project[t.get("project_id") or ""].append(t)
    for pid in tasks_by_project:
        tasks_by_project[pid].sort(key=sort_key_order_title)

    images_by_parent: dict[tuple[str, str], list[dict]] = defaultdict(list)
    for a in attachments:
        if (a.get("kind") or "").strip().lower() != "image":
            continue
        ptype = (a.get("parent_type") or "").strip()
        pid = (a.get("parent_id") or "").strip()
        if ptype and pid:
            images_by_parent[(ptype, pid)].append(a)

    def attachment_web_src(a: dict) -> str:
        ref = (a.get("ref") or "").strip().replace("\\", "/")
        if not ref.startswith("attachments/"):
            return ""
        disk = data_dir / Path(ref)
        if not disk.is_file():
            return ""
        return "../data/" + ref

    def thumbs_html(parent_type: str, parent_id: str) -> str:
        imgs = images_by_parent.get((parent_type, parent_id), [])
        if not imgs:
            return ""
        chips: list[str] = []
        for a in imgs:
            src = attachment_web_src(a)
            if not src:
                continue
            desc = (a.get("description") or a.get("ref") or "image").strip()
            chips.append(
                f'<button type="button" class="thumb" data-full="{esc(src)}" title="{esc(desc)}">'
                f'<img src="{esc(src)}" alt="{esc(desc)}" loading="lazy">'
                f"</button>"
            )
        if not chips:
            return ""
        return f'<div class="thumbs">{"".join(chips)}</div>'

    def projects_for_filter(filt: str) -> list[dict]:
        rows = [p for p in active_projects if p.get("filter") == filt]
        rows.sort(key=sort_key_order_title)
        return rows

    def render_task_li(t: dict) -> str:
        due = (t.get("due_date") or t.get("next_due") or "").strip()
        due_html = f'<span class="due">{esc(due)}</span>' if due else ""
        thumbs = thumbs_html("task", t.get("id") or "")
        em = t.get("emoji") or "🔲"
        return (
            f'<li class="task" data-emoji="{esc(em)}">'
            f'<span class="em">{esc(em)}</span>'
            f'<span class="title">{esc(t.get("title"))}</span>'
            f"{due_html}"
            f"{thumbs}"
            f"</li>"
        )

    def render_column(heading: str, filt: str) -> str:
        projs = projects_for_filter(filt)
        blocks: list[str] = []
        for p in projs:
            pid = p.get("id") or ""
            pts = tasks_by_project.get(pid, [])
            proj_thumbs = thumbs_html("project", pid)
            if not pts and not proj_thumbs:
                continue
            items = "".join(render_task_li(t) for t in pts)
            task_block = f'<ul class="task-list">{items}</ul>' if items else ""
            blocks.append(
                f'<details class="project" open>'
                f"<summary>"
                f'<span class="proj-title">{esc(p.get("title"))}</span>'
                f'<span class="proj-count">{len(pts)}</span>'
                f"</summary>"
                f"{proj_thumbs}"
                f"{task_block}"
                f"</details>"
            )
        # Orphan open tasks (project missing / inactive)
        known = {p.get("id") for p in projs}
        orphans = [
            t
            for t in open_punctual
            if t.get("filter") == filt and t.get("project_id") not in known
        ]
        if orphans:
            orphans.sort(key=sort_key_order_title)
            items = "".join(render_task_li(t) for t in orphans)
            blocks.append(
                f'<details class="project" open>'
                f"<summary>"
                f'<span class="proj-title">Other</span>'
                f'<span class="proj-count">{len(orphans)}</span>'
                f"</summary>"
                f'<ul class="task-list">{items}</ul>'
                f"</details>"
            )
        body = "".join(blocks) if blocks else '<p class="empty">No open tasks</p>'
        return (
            f'<section class="column panel">'
            f"<h2>{esc(heading)}</h2>"
            f"{body}"
            f"</section>"
        )

    def habit_cards_html() -> str:
        by_rec: dict[str, list[dict]] = defaultdict(list)
        for t in habitual:
            by_rec[(t.get("recurrence") or "").strip().lower()].append(t)

        ordered_keys: list[str] = []
        for key in RECURRENCE_ORDER:
            if key in by_rec:
                ordered_keys.append(key)
        for key in sorted(by_rec.keys()):
            if key not in ordered_keys:
                ordered_keys.append(key)

        if not ordered_keys:
            return '<p class="empty">No habits</p>'

        cards: list[str] = []
        for key in ordered_keys:
            rows = by_rec[key]
            rows.sort(
                key=lambda t: (
                    (t.get("next_due") or t.get("due_date") or "9999-99-99"),
                    (t.get("title") or "").lower(),
                )
            )
            label = RECURRENCE_LABELS.get(key, key.replace("_", " ").title() or "Other")
            lis = []
            for t in rows:
                due = (t.get("next_due") or t.get("due_date") or "").strip()
                due_cls = " due-now" if is_due(t) else ""
                due_html = f'<span class="due">{esc(due)}</span>' if due else ""
                em = t.get("emoji") or "🔲"
                lis.append(
                    f'<li class="habit-row{due_cls}" data-emoji="{esc(em)}">'
                    f'<span class="em">{esc(em)}</span>'
                    f'<span class="title">{esc(t.get("title"))}</span>'
                    f"{due_html}"
                    f"</li>"
                )
            cards.append(
                f'<article class="habit-card panel">'
                f"<h2>{esc(label)}</h2>"
                f'<ul class="habit-list">{"".join(lis)}</ul>'
                f"</article>"
            )
        return '<div class="habit-cards">' + "".join(cards) + "</div>"

    # Dashboard filter keys (letter → emoji). Toggle same key or Esc to clear.
    EMOJI_FILTER_KEYS = [
        ("t", "⚡", "thunderbolt"),
        ("w", "⏳", "waiting"),
        ("g", "🔲", "general"),
        ("u", "❓", "doubt"),
    ]
    key_by_emoji = {em: k for k, em, _ in EMOJI_FILTER_KEYS}

    emoji_legend_parts: list[str] = []
    for e, n in by_emoji.most_common():
        key = key_by_emoji.get(e, "")
        key_html = f'<kbd>{esc(key.upper())}</kbd> ' if key else ""
        emoji_legend_parts.append(
            f'<button type="button" class="emoji-chip" data-emoji="{esc(e)}" '
            f'data-key="{esc(key)}" title="Filter {esc(e)}'
            + (f' ({esc(key.upper())})' if key else "")
            + f'">'
            f"{key_html}"
            f'<span class="em">{esc(e)}</span>'
            f'<span class="n">×{n}</span></button>'
        )
    emoji_legend = (
        "".join(emoji_legend_parts)
        or '<span class="empty-inline">No open tasks</span>'
    )
    filter_hint = " · ".join(
        f"<kbd>{k.upper()}</kbd> {em}" for k, em, _ in EMOJI_FILTER_KEYS
    )

    personal_col = render_column("Personal", "personal")
    work_col = render_column("Work", "work")
    habits_block = habit_cards_html()

    generated = datetime.now().strftime("%Y-%m-%d %H:%M")
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8"/>
<title>Tasks Dashboard</title>
<style>
:root {{
  --bg:#121212; --panel:#1e1e1e; --text:#f2f2f2; --muted:#a0a0a0;
  --gold:#f1c40f; --line:#333; --due:#e67e22;
}}
* {{ box-sizing:border-box; }}
body {{ margin:0; font-family:Segoe UI,system-ui,sans-serif; background:var(--bg); color:var(--text); }}
header {{ padding:20px 28px 4px; }}
h1 {{ margin:0; font-size:26px; }}
.sub {{ color:var(--muted); margin-top:6px; font-size:13px; }}
.strip {{ display:flex; flex-wrap:wrap; align-items:center; gap:10px 16px; padding:12px 28px 8px; }}
.kpi {{ background:var(--panel); border:1px solid var(--line); border-radius:8px; padding:8px 12px; min-width:88px; }}
.kpi .lbl {{ color:var(--muted); font-size:10px; text-transform:uppercase; letter-spacing:.04em; }}
.kpi .val {{ font-size:18px; margin-top:2px; color:var(--gold); }}
.emoji-legend {{ display:flex; flex-wrap:wrap; gap:6px; align-items:center; }}
.emoji-legend .legend-lbl {{ color:var(--muted); font-size:11px; margin-right:4px; }}
.emoji-chip {{
  display:inline-flex; align-items:center; gap:4px;
  background:var(--panel); border:1px solid var(--line);
  border-radius:999px; padding:2px 8px; font-size:12px;
  color:inherit; cursor:pointer; font:inherit;
}}
.emoji-chip:hover, .emoji-chip.active {{ border-color:var(--gold); color:var(--gold); }}
.emoji-chip .n {{ color:var(--muted); }}
.task.emoji-hidden, .habit-row.emoji-hidden, .project.emoji-hidden,
.habit-card.emoji-hidden {{ display:none !important; }}
#emojiFilterLbl {{
  color:var(--gold); font-size:12px; min-height:1.2em;
}}
.top-row {{
  display:grid; grid-template-columns:1fr 1fr; gap:16px;
  padding:8px 28px 12px;
}}
.habits-wrap {{ padding:8px 28px 32px; }}
.habits-wrap > .section-title {{
  margin:8px 0 14px; font-size:18px; color:var(--gold);
}}
.panel {{
  background:var(--panel); border:1px solid var(--line);
  border-radius:12px; padding:14px 16px;
}}
.column h2 {{ margin:0 0 12px; font-size:18px; color:var(--gold); }}
.project {{
  border:1px solid var(--line); border-radius:8px;
  margin-bottom:8px; background:#252528;
}}
.project summary {{
  cursor:pointer; list-style:none; display:flex; align-items:center;
  justify-content:space-between; gap:10px; padding:10px 12px;
  user-select:none;
}}
.project summary::-webkit-details-marker {{ display:none; }}
.project summary::before {{
  content:"▸"; color:var(--gold); margin-right:8px; font-size:12px;
}}
.project[open] summary::before {{ content:"▾"; }}
.proj-title {{ font-weight:600; flex:1; }}
.proj-count {{
  color:var(--muted); font-size:12px; background:#1a1a1c;
  border-radius:999px; padding:2px 8px;
}}
.task-list, .habit-list {{ list-style:none; margin:0; padding:0 12px 10px; }}
.task, .habit-row {{
  display:flex; flex-wrap:wrap; align-items:baseline; gap:8px;
  padding:6px 0; border-top:1px solid var(--line); font-size:13px;
}}
.task .em, .habit-row .em {{ flex:0 0 auto; }}
.task .title, .habit-row .title {{ flex:1; min-width:120px; }}
.due {{ color:var(--muted); font-size:11px; white-space:nowrap; }}
.thumbs {{
  flex:1 1 100%; display:flex; flex-wrap:wrap; gap:6px;
  padding:4px 0 2px 28px;
}}
.project > .thumbs {{ padding:0 12px 10px; }}
.thumb {{
  display:block; padding:0; margin:0; border:1px solid var(--line);
  border-radius:6px; background:#1a1a1c; cursor:zoom-in; overflow:hidden;
  width:72px; height:54px; flex:0 0 auto;
}}
.thumb img {{
  display:block; width:100%; height:100%; object-fit:cover;
}}
.thumb:hover {{ border-color:var(--gold); }}
#lightbox {{
  position:fixed; inset:0; z-index:1000;
  background:rgba(0,0,0,.82);
  display:flex; align-items:center; justify-content:center;
  padding:24px; cursor:zoom-out;
}}
#lightbox[hidden] {{ display:none; }}
#lightbox img {{
  max-width:min(96vw, 1200px); max-height:92vh;
  border-radius:8px; box-shadow:0 8px 40px rgba(0,0,0,.5);
  cursor:default;
}}
#lightbox .lb-close {{
  position:absolute; top:16px; right:20px;
  background:transparent; border:0; color:#fff;
  font-size:28px; line-height:1; cursor:pointer; padding:4px 10px;
}}
.habit-row.due-now .due {{ color:var(--due); }}
.habit-cards {{
  display:grid; grid-template-columns:repeat(auto-fill, minmax(260px, 1fr));
  gap:14px;
}}
.habit-card h2 {{
  margin:0 0 12px; font-size:22px; font-weight:700; color:var(--gold);
  letter-spacing:.02em;
}}
.empty {{ color:var(--muted); font-size:13px; margin:0; }}
.empty-inline {{ color:var(--muted); font-size:12px; }}
kbd {{
  font:inherit; font-size:11px; color:var(--gold);
  border:1px solid var(--line); border-radius:4px;
  padding:0 5px; background:#252528;
}}
@media (max-width:900px) {{
  .top-row {{ grid-template-columns:1fr; }}
}}
</style>
</head>
<body>
<header>
  <h1>Tasks</h1>
  <div class="sub">Generated {esc(generated)} · {len(active_projects)} projects · {len(active_tasks)} tasks · {len(infos)} info · {len(attachments)} attachments · <kbd>E</kbd> expand/collapse · {filter_hint} · Esc clear filter</div>
</header>
<div class="strip">
  <div class="kpi"><div class="lbl">Personal</div><div class="val">{len(open_personal)}</div></div>
  <div class="kpi"><div class="lbl">Work</div><div class="val">{len(open_work)}</div></div>
  <div class="kpi"><div class="lbl">Habits due</div><div class="val">{habits_due_n}</div></div>
  <div class="emoji-legend">
    <span class="legend-lbl">Emoji</span>
    {emoji_legend}
  </div>
  <div id="emojiFilterLbl"></div>
</div>
<div class="top-row">
  {personal_col}
  {work_col}
</div>
<div class="habits-wrap">
  <h2 class="section-title">Habits</h2>
  {habits_block}
</div>
<div id="lightbox" hidden>
  <button type="button" class="lb-close" aria-label="Close">×</button>
  <img src="" alt="">
</div>
<script>
(function () {{
  const EMOJI_KEYS = {{ t: "⚡", w: "⏳", g: "🔲", u: "❓" }};
  let activeEmoji = "";
  const filterLbl = document.getElementById("emojiFilterLbl");

  function toggleAllProjects() {{
    const nodes = document.querySelectorAll("details.project");
    if (!nodes.length) return;
    const openAll = Array.from(nodes).some((n) => !n.open);
    nodes.forEach((n) => {{ n.open = openAll; }});
  }}

  function applyEmojiFilter(emoji) {{
    activeEmoji = emoji || "";
    document.querySelectorAll(".task, .habit-row").forEach((el) => {{
      const match = !activeEmoji || el.getAttribute("data-emoji") === activeEmoji;
      el.classList.toggle("emoji-hidden", !match);
    }});
    document.querySelectorAll("details.project").forEach((proj) => {{
      if (!activeEmoji) {{
        proj.classList.remove("emoji-hidden");
        return;
      }}
      const visible = proj.querySelectorAll(".task:not(.emoji-hidden)").length;
      proj.classList.toggle("emoji-hidden", visible === 0);
      if (visible > 0) proj.open = true;
    }});
    document.querySelectorAll(".habit-card").forEach((card) => {{
      if (!activeEmoji) {{
        card.classList.remove("emoji-hidden");
        return;
      }}
      const visible = card.querySelectorAll(".habit-row:not(.emoji-hidden)").length;
      card.classList.toggle("emoji-hidden", visible === 0);
    }});
    document.querySelectorAll(".emoji-chip[data-emoji]").forEach((chip) => {{
      chip.classList.toggle("active", !!activeEmoji && chip.getAttribute("data-emoji") === activeEmoji);
    }});
    if (filterLbl) {{
      filterLbl.textContent = activeEmoji
        ? ("Showing " + activeEmoji + " only — Esc or same key to clear")
        : "";
    }}
  }}

  function toggleEmojiFilter(emoji) {{
    if (!emoji) {{
      applyEmojiFilter("");
      return;
    }}
    applyEmojiFilter(activeEmoji === emoji ? "" : emoji);
  }}

  const lb = document.getElementById("lightbox");
  const lbImg = lb ? lb.querySelector("img") : null;
  function openLightbox(src) {{
    if (!lb || !lbImg || !src) return;
    lbImg.src = src;
    lb.hidden = false;
  }}
  function closeLightbox() {{
    if (!lb || !lbImg) return;
    lb.hidden = true;
    lbImg.src = "";
  }}

  document.addEventListener("click", (e) => {{
    const chip = e.target.closest(".emoji-chip[data-emoji]");
    if (chip) {{
      e.preventDefault();
      toggleEmojiFilter(chip.getAttribute("data-emoji") || "");
      return;
    }}
    const btn = e.target.closest(".thumb");
    if (btn) {{
      e.preventDefault();
      e.stopPropagation();
      openLightbox(btn.getAttribute("data-full") || "");
      return;
    }}
    if (e.target.closest(".lb-close") || e.target === lb) {{
      closeLightbox();
    }}
  }});

  document.addEventListener("keydown", (e) => {{
    if (e.ctrlKey || e.metaKey || e.altKey) return;
    const tag = (e.target && e.target.tagName) ? e.target.tagName.toLowerCase() : "";
    if (tag === "input" || tag === "textarea" || tag === "select" || (e.target && e.target.isContentEditable))
      return;

    if (e.key === "Escape") {{
      e.preventDefault();
      if (lb && !lb.hidden) {{
        closeLightbox();
        return;
      }}
      if (activeEmoji) {{
        applyEmojiFilter("");
        return;
      }}
      return;
    }}

    const k = (e.key || "").toLowerCase();
    if (k === "e") {{
      e.preventDefault();
      toggleAllProjects();
      return;
    }}
    if (EMOJI_KEYS[k]) {{
      e.preventDefault();
      toggleEmojiFilter(EMOJI_KEYS[k]);
    }}
  }});
}})();
</script>
</body>
</html>
"""


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--data-dir", required=True)
    ap.add_argument("--output-dir", required=True)
    args = ap.parse_args()
    data_dir = Path(args.data_dir)
    out_dir = Path(args.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    html_out = build_html(data_dir)
    (out_dir / "dashboard.html").write_text(html_out, encoding="utf-8")
    print("Wrote", out_dir / "dashboard.html")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
