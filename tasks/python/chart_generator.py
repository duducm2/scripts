"""Build Tasks dashboard HTML from CSV data."""

from __future__ import annotations

import argparse
import csv
import html
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path


def read_csv(path: Path) -> list[dict]:
    if not path.exists():
        return []
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def esc(s: str) -> str:
    return html.escape(s or "")


def build_html(data_dir: Path) -> str:
    projects = read_csv(data_dir / "projects.csv")
    tasks = read_csv(data_dir / "tasks.csv")
    infos = read_csv(data_dir / "info_points.csv")
    attachments = read_csv(data_dir / "attachments.csv")

    active_tasks = [t for t in tasks if t.get("active", "1") != "0"]
    open_tasks = [t for t in active_tasks if t.get("emoji") != "✅"]
    done_tasks = [t for t in active_tasks if t.get("emoji") == "✅"]
    habitual = [t for t in open_tasks if t.get("kind") == "habitual"]

    by_filter = Counter(t.get("filter") or "—" for t in open_tasks)
    by_emoji = Counter(t.get("emoji") or "🔲" for t in open_tasks)

    proj_title = {p["id"]: p.get("title") or p["id"] for p in projects}

    def sort_due(t: dict) -> str:
        return t.get("next_due") or t.get("due_date") or "9999-99-99"

    habitual_sorted = sorted(habitual, key=sort_due)[:40]
    recent_done = sorted(
        done_tasks, key=lambda t: t.get("completed_at") or "", reverse=True
    )[:30]

    info_by_parent: dict[tuple[str, str], int] = defaultdict(int)
    for i in infos:
        info_by_parent[(i.get("parent_type", ""), i.get("parent_id", ""))] += 1

    filter_cards = "".join(
        f'<div class="kpi"><div class="lbl">{esc(k)}</div><div class="val">{v}</div></div>'
        for k, v in sorted(by_filter.items())
    )
    emoji_rows = "".join(
        f"<tr><td>{esc(e)}</td><td>{n}</td></tr>" for e, n in by_emoji.most_common()
    )
    habit_rows = "".join(
        "<tr>"
        f"<td>{esc(t.get('emoji'))}</td>"
        f"<td>{esc(t.get('title'))}</td>"
        f"<td>{esc(t.get('recurrence'))}</td>"
        f"<td>{esc(t.get('next_due') or t.get('due_date'))}</td>"
        f"<td>{esc(t.get('filter'))}</td>"
        "</tr>"
        for t in habitual_sorted
    )
    done_rows = "".join(
        "<tr>"
        f"<td>{esc(t.get('title'))}</td>"
        f"<td>{esc(proj_title.get(t.get('project_id', ''), ''))}</td>"
        f"<td>{esc(t.get('completed_at'))}</td>"
        f"<td>{esc(t.get('filter'))}</td>"
        "</tr>"
        for t in recent_done
    )
    project_rows = "".join(
        "<tr>"
        f"<td>{esc(p.get('title'))}</td>"
        f"<td>{esc(p.get('filter'))}</td>"
        f"<td>{sum(1 for t in open_tasks if t.get('project_id') == p.get('id'))}</td>"
        f"<td>{info_by_parent.get(('project', p.get('id', '')), 0)}</td>"
        "</tr>"
        for p in projects
        if p.get("active", "1") != "0"
    )

    generated = datetime.now().strftime("%Y-%m-%d %H:%M")
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8"/>
<title>Tasks Dashboard</title>
<style>
:root {{ --bg:#121212; --panel:#1e1e1e; --text:#f2f2f2; --muted:#a0a0a0; --gold:#f1c40f; --line:#333; }}
* {{ box-sizing:border-box; }}
body {{ margin:0; font-family:Segoe UI,system-ui,sans-serif; background:var(--bg); color:var(--text); }}
header {{ padding:24px 28px 8px; }}
h1 {{ margin:0; font-size:28px; }}
.sub {{ color:var(--muted); margin-top:6px; }}
.kpis {{ display:flex; flex-wrap:wrap; gap:12px; padding:16px 28px; }}
.kpi {{ background:var(--panel); border:1px solid var(--line); border-radius:10px; padding:14px 18px; min-width:120px; }}
.kpi .lbl {{ color:var(--muted); font-size:12px; text-transform:uppercase; letter-spacing:.04em; }}
.kpi .val {{ font-size:24px; margin-top:4px; color:var(--gold); }}
.grid {{ display:grid; grid-template-columns:1fr 1fr; gap:16px; padding:8px 28px 32px; }}
.panel {{ background:var(--panel); border:1px solid var(--line); border-radius:12px; padding:16px 18px; }}
.panel h2 {{ margin:0 0 12px; font-size:16px; color:var(--gold); }}
table {{ width:100%; border-collapse:collapse; font-size:13px; }}
th, td {{ text-align:left; padding:6px 8px; border-bottom:1px solid var(--line); vertical-align:top; }}
th {{ color:var(--muted); font-weight:600; }}
.wide {{ grid-column:1 / -1; }}
@media (max-width:900px) {{ .grid {{ grid-template-columns:1fr; }} }}
</style>
</head>
<body>
<header>
  <h1>Tasks</h1>
  <div class="sub">Generated {esc(generated)} · {len(projects)} projects · {len(active_tasks)} tasks · {len(infos)} info · {len(attachments)} attachments</div>
</header>
<div class="kpis">
  <div class="kpi"><div class="lbl">Open</div><div class="val">{len(open_tasks)}</div></div>
  <div class="kpi"><div class="lbl">Done</div><div class="val">{len(done_tasks)}</div></div>
  <div class="kpi"><div class="lbl">Habits open</div><div class="val">{len(habitual)}</div></div>
  {filter_cards}
</div>
<div class="grid">
  <div class="panel">
    <h2>Open by emoji</h2>
    <table><thead><tr><th>Emoji</th><th>Count</th></tr></thead><tbody>{emoji_rows or '<tr><td colspan="2">None</td></tr>'}</tbody></table>
  </div>
  <div class="panel">
    <h2>Projects</h2>
    <table><thead><tr><th>Title</th><th>Filter</th><th>Open</th><th>Info</th></tr></thead><tbody>{project_rows or '<tr><td colspan="4">None</td></tr>'}</tbody></table>
  </div>
  <div class="panel wide">
    <h2>Habits upcoming</h2>
    <table><thead><tr><th></th><th>Title</th><th>Recurrence</th><th>Next</th><th>Filter</th></tr></thead><tbody>{habit_rows or '<tr><td colspan="5">None</td></tr>'}</tbody></table>
  </div>
  <div class="panel wide">
    <h2>Recently completed</h2>
    <table><thead><tr><th>Title</th><th>Project</th><th>Completed</th><th>Filter</th></tr></thead><tbody>{done_rows or '<tr><td colspan="4">None</td></tr>'}</tbody></table>
  </div>
</div>
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
