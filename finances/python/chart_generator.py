"""Build local Plotly cockpit dashboard (all charts on one page)."""

from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from data_aggregator import (  # noqa: E402
    OUTPUT,
    format_brl,
    parse_decimal,
    snapshot,
    widget_on,
    cat_index,
    cat_label,
)
from seed_from_ini import seed  # noqa: E402


def pie_spec(rows):
    return {
        "labels": [r[0] for r in rows],
        "values": [round(r[1], 2) for r in rows],
        "colors": [r[2] for r in rows],
        "custom": [format_brl(r[1]) for r in rows],
    }


def build_html(data: dict) -> str:
    s = data["settings"]
    notes = data["notifications"] if widget_on(s, "ShowNotifications") else []
    note_html = (
        "".join(f'<div class="note">{n}</div>' for n in notes)
        or '<div class="note ok">No alerts</div>'
    )

    cards_html = ""
    if widget_on(s, "ShowBalance"):
        cards_html = f"""
        <div class="kpis">
          <div class="kpi"><div class="lbl">Balance</div><div class="val">{format_brl(data['balance'])}</div></div>
          <div class="kpi"><div class="lbl">Incomes</div><div class="val pos">{format_brl(data['totals']['income'])}</div></div>
          <div class="kpi"><div class="lbl">Expenses</div><div class="val neg">{format_brl(data['totals']['expense'])}</div></div>
          <div class="kpi"><div class="lbl">Card avail.</div><div class="val">{format_brl(data['card_available'])} <span class="dim">/ {format_brl(data['card_limit'])}</span></div></div>
        </div>"""

    perf_html = ""
    if widget_on(s, "ShowPerformance"):
        prev_b = data["prev_totals"]["balance"]
        cur_b = data["totals"]["balance"]
        vs = ((cur_b - prev_b) / abs(prev_b) * 100) if prev_b else 0
        top = (
            " · ".join(
                f"{name} {format_brl(amt)}" for name, amt, _ in data["top_expenses"][:4]
            )
            or "No expenses"
        )
        perf_html = f"""
        <div class="panel panel-slim">
          <div class="perf-line">Saved last month {format_brl(prev_b)} · This month {format_brl(cur_b)} ({vs:+.0f}%) · Kept {data['saved_pct']:.0f}% · Top: {top}</div>
        </div>"""

    pies_html = ""
    if widget_on(s, "ShowPies"):
        pies_html = """
          <div class="panel chart-cell"><h2>Expenses by category</h2><div id="pieExp" class="chart"></div></div>
          <div class="panel chart-cell"><h2>Incomes by category</h2><div id="pieInc" class="chart"></div></div>"""

    reports_html = """
          <div class="panel chart-cell"><h2>Spent per main category</h2><div id="barCat" class="chart"></div></div>
          <div class="panel chart-cell"><h2>Monthly balance</h2><div id="barBal" class="chart"></div></div>
          <div class="panel chart-cell"><h2>Cash flow</h2><div id="lineCf" class="chart"></div></div>
          <div class="panel chart-cell"><h2>Annual cash flow</h2><div id="lineYear" class="chart"></div></div>"""

    goals_html = ""
    if widget_on(s, "ShowGoals"):
        items = []
        for g in data["goals"]:
            name = g.get("name", "")
            status = (g.get("status") or "in_progress").strip().lower()
            cur = parse_decimal(g.get("current_amount"))
            tgt = parse_decimal(g.get("target_amount"))
            pct = (cur / tgt * 100) if tgt > 0 else (100.0 if cur > 0 else 0.0)
            width = min(pct, 100.0)
            rem = tgt - cur
            if status == "completed" or (tgt > 0 and cur >= tgt):
                fill = "#f1c40f"
                status = "completed"
                cap = "Reached"
            elif status == "expired":
                fill = "#7f8c8d"
                cap = f"Rem {format_brl(rem)}" if rem > 0 else "Reached"
            else:
                fill = "#3498db"
                cap = f"Rem {format_brl(rem)}" if rem > 0 else "Reached"
            items.append(
                f'<div class="bar-row">'
                f'<div class="bar-head"><span>{name}</span>'
                f"<span>{format_brl(cur)} / {format_brl(tgt)} · {pct:.0f}%</span></div>"
                f'<div class="bar-track"><div class="bar-fill" style="width:{width:.1f}%;background:{fill}"></div></div>'
                f'<div class="bar-meta">{status} · {cap}</div>'
                f"</div>"
            )
        goals_html = f"""
        <div class="panel"><h2>Goals</h2>
          {''.join(items) or '<p class="empty">No goals</p>'}</div>"""

    bud_html = ""
    if widget_on(s, "ShowBudgets"):
        by_id = cat_index(data["categories"])
        items = []
        for b in data["budgets"]:
            crow = by_id.get(b.get("category_id", ""))
            name = cat_label(crow, b.get("category_id", ""))
            p = parse_decimal(b.get("planned_amount"))
            sp = parse_decimal(b.get("spent_amount"))
            rem = p - sp
            pct = (sp / p * 100) if p > 0 else (100.0 if sp > 0 else 0.0)
            width = min(pct, 100.0)
            over = sp > p and p > 0
            fill = "#e74c3c" if over else "#2ecc71"
            cap = (
                f"Exceeded {format_brl(-rem)}" if over else f"Remain {format_brl(rem)}"
            )
            items.append(
                f'<div class="bar-row">'
                f'<div class="bar-head"><span>{name}</span>'
                f"<span>{format_brl(sp)} / {format_brl(p)} · {pct:.0f}%</span></div>"
                f'<div class="bar-track"><div class="bar-fill" style="width:{width:.1f}%;background:{fill}"></div></div>'
                f'<div class="bar-meta">{cap}</div>'
                f"</div>"
            )
        bud_html = f"""
        <div class="panel"><h2>Budgets</h2>
          {''.join(items) or '<p class="empty">No budgets this month</p>'}</div>"""

    payload = {
        "expensePie": pie_spec(data["expense_pie"]),
        "incomePie": pie_spec(data["income_pie"]),
        "series": data["series"],
        "annual": data["annual"],
        "spentMain": [{"name": n, "value": v} for n, v, _ in data["expense_pie"]],
    }
    payload_json = json.dumps(payload, ensure_ascii=False)

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8"/>
  <title>Finance cockpit</title>
  <script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
  <style>
    :root, [data-theme="dark"] {{
      --bg: #121212;
      --header: #1a1a1a;
      --border: #2a2a2a;
      --panel: #1e1e1e;
      --text: #eee;
      --muted: #888;
      --muted2: #777;
      --heading: #bbb;
      --perf: #ccc;
      --track: #333;
      --empty: #666;
      --note-bg: #3d2b00;
      --note-fg: #f1c40f;
      --note-ok-bg: #143d27;
      --note-ok-fg: #2ecc71;
      --toggle-bg: #2c2c2c;
      --toggle-fg: #eee;
      --plot-paper: #1e1e1e;
      --plot-font: #ccc;
    }}
    [data-theme="light"] {{
      --bg: #f4f5f7;
      --header: #ffffff;
      --border: #dde1e6;
      --panel: #ffffff;
      --text: #1a1a1a;
      --muted: #6b7280;
      --muted2: #6b7280;
      --heading: #374151;
      --perf: #4b5563;
      --track: #e5e7eb;
      --empty: #9ca3af;
      --note-bg: #fff7e0;
      --note-fg: #92400e;
      --note-ok-bg: #ecfdf5;
      --note-ok-fg: #047857;
      --toggle-bg: #eef0f3;
      --toggle-fg: #1a1a1a;
      --plot-paper: #ffffff;
      --plot-font: #374151;
    }}
    body {{ font-family: Segoe UI, sans-serif; background:var(--bg); color:var(--text); margin:0; font-size:13px; }}
    header {{ padding:8px 14px; background:var(--header); border-bottom:1px solid var(--border);
      display:flex; align-items:center; justify-content:space-between; gap:12px; }}
    header h1 {{ margin:0; font-size:15px; font-weight:600; }}
    #themeToggle {{ background:var(--toggle-bg); color:var(--toggle-fg); border:1px solid var(--border);
      border-radius:6px; padding:5px 10px; font-size:12px; cursor:pointer; }}
    #themeToggle:hover {{ filter:brightness(1.08); }}
    main {{ padding:12px 16px 20px; }}
    .kpis {{ display:grid; grid-template-columns:repeat(4,1fr); gap:10px; margin-bottom:10px; }}
    .kpi {{ background:var(--panel); padding:8px 10px; border-radius:6px; border:1px solid var(--border); }}
    .kpi .lbl {{ color:var(--muted); font-size:11px; text-transform:uppercase; letter-spacing:.03em; }}
    .kpi .val {{ font-size:16px; margin-top:2px; font-weight:600; }}
    .dim {{ color:var(--muted2); font-size:12px; font-weight:400; }}
    .pos {{ color:#2ecc71; }} .neg {{ color:#e74c3c; }}
    .charts {{ display:grid; grid-template-columns:1fr 1fr; gap:10px; margin-bottom:10px; }}
    .split {{ display:grid; grid-template-columns:1fr 1fr; gap:10px; }}
    .panel {{ background:var(--panel); padding:10px 12px; border-radius:6px; margin-bottom:0; border:1px solid var(--border); }}
    .panel-slim {{ margin-bottom:10px; }}
    .panel h2 {{ margin:0 0 4px; font-size:12px; color:var(--heading); font-weight:600; }}
    .chart {{ height:320px; }}
    .chart-cell {{ min-width:0; }}
    .note {{ background:var(--note-bg); color:var(--note-fg); padding:6px 10px; border-radius:4px; margin-bottom:10px; font-size:12px; }}
    .note.ok {{ background:var(--note-ok-bg); color:var(--note-ok-fg); }}
    .perf-line {{ color:var(--perf); font-size:12px; line-height:1.4; }}
    .bar-row {{ margin:6px 0 8px; }}
    .bar-head {{ display:flex; justify-content:space-between; gap:8px; font-size:12px; margin-bottom:3px; }}
    .bar-track {{ height:7px; background:var(--track); border-radius:4px; overflow:hidden; }}
    .bar-fill {{ height:100%; border-radius:4px; }}
    .bar-meta {{ color:var(--muted2); font-size:11px; margin-top:2px; }}
    .empty {{ color:var(--empty); margin:0; }}
    @media (max-width:900px) {{
      .kpis, .charts, .split {{ grid-template-columns:1fr; }}
    }}
  </style>
</head>
<body>
<header>
  <h1>Finance cockpit · {data['year_month']}</h1>
  <button type="button" id="themeToggle" aria-label="Toggle theme">Light</button>
</header>
<main>
  {note_html}
  {cards_html}
  {perf_html}
  <div class="charts">
    {pies_html}
    {reports_html}
  </div>
  <div class="split">
    {goals_html}
    {bud_html}
  </div>
</main>
<script>
const DATA = {payload_json};
const THEME_KEY = 'finance-cockpit-theme';
function themeColors() {{
  const cs = getComputedStyle(document.documentElement);
  return {{
    paper: cs.getPropertyValue('--plot-paper').trim() || '#1e1e1e',
    font: cs.getPropertyValue('--plot-font').trim() || '#ccc'
  }};
}}
function baseLayout() {{
  const t = themeColors();
  return {{paper_bgcolor:t.paper, plot_bgcolor:t.paper, font:{{color:t.font, size:11}},
    margin:{{t:28,b:48,l:42,r:16}}, height:320, legend:{{orientation:'h', y:1.12, font:{{size:10}}}}}};
}}
function pie(id, spec) {{
  const el = document.getElementById(id);
  if (!el) return;
  if (!spec.values.length) {{ el.innerHTML = '<p class="empty">No data</p>'; return; }}
  const L = baseLayout();
  Plotly.newPlot(id, [{{
    type:'pie', labels:spec.labels, values:spec.values, marker:{{colors:spec.colors}},
    customdata: spec.custom, textfont:{{size:10}},
    domain:{{x:[0, 0.62], y:[0, 1]}},
    hovertemplate: '%{{label}}<br>%{{percent}}<br>%{{customdata}}<extra></extra>'
  }}], Object.assign({{}}, L, {{
    showlegend:true,
    margin:{{t:28, b:20, l:8, r:8}},
    legend:{{orientation:'v', x:1.02, y:0.5, xanchor:'left', font:{{size:10}}}}
  }}), {{responsive:true, displayModeBar:false}});
}}
function drawAll() {{
  const L = baseLayout();
  pie('pieExp', DATA.expensePie);
  pie('pieInc', DATA.incomePie);
  const names = DATA.spentMain.map(x => x.name);
  const vals = DATA.spentMain.map(x => x.value);
  Plotly.newPlot('barCat', [{{type:'bar', x:names, y:vals, marker:{{color:'#e67e22'}},
    hovertemplate:'%{{x}}<br>R$ %{{y:.2f}}<extra></extra>'}}],
    Object.assign({{}}, L, {{
      showlegend:false,
      height:320,
      margin:{{t:28, b:110, l:48, r:16}},
      xaxis:{{tickangle:-45, automargin:true, tickfont:{{size:10}}}}
    }}), {{responsive:true, displayModeBar:false}});
  const months = DATA.series.map(x => x.month);
  const bals = DATA.series.map(x => x.balance);
  const colors = bals.map(v => v >= 0 ? '#27ae60' : '#c0392b');
  Plotly.newPlot('barBal', [{{type:'bar', x:months, y:bals, marker:{{color:colors}},
    hovertemplate:'%{{x}}<br>R$ %{{y:.2f}}<extra></extra>'}}],
    Object.assign({{}}, L, {{showlegend:false}}), {{responsive:true, displayModeBar:false}});
  Plotly.newPlot('lineCf', [
    {{type:'scatter', mode:'lines+markers', name:'In', x:months, y:DATA.series.map(x=>x.income), line:{{color:'#2ecc71'}}}},
    {{type:'scatter', mode:'lines+markers', name:'Out', x:months, y:DATA.series.map(x=>x.expense), line:{{color:'#e74c3c'}}}}
  ], L, {{responsive:true, displayModeBar:false}});
  Plotly.newPlot('lineYear', [
    {{type:'scatter', mode:'lines+markers', name:'In', x:DATA.annual.map(x=>x.label), y:DATA.annual.map(x=>x.income), line:{{color:'#2ecc71'}}}},
    {{type:'scatter', mode:'lines+markers', name:'Out', x:DATA.annual.map(x=>x.label), y:DATA.annual.map(x=>x.expense), line:{{color:'#e74c3c'}}}},
    {{type:'scatter', mode:'lines+markers', name:'Bal', x:DATA.annual.map(x=>x.label), y:DATA.annual.map(x=>x.balance), line:{{color:'#f1c40f'}}}}
  ], L, {{responsive:true, displayModeBar:false}});
}}
function applyTheme(theme) {{
  document.documentElement.setAttribute('data-theme', theme);
  localStorage.setItem(THEME_KEY, theme);
  const btn = document.getElementById('themeToggle');
  if (btn) btn.textContent = theme === 'dark' ? 'Light' : 'Dark';
  drawAll();
}}
(function initTheme() {{
  let theme = localStorage.getItem(THEME_KEY);
  if (theme !== 'light' && theme !== 'dark') theme = 'dark';
  document.documentElement.setAttribute('data-theme', theme);
  const btn = document.getElementById('themeToggle');
  if (btn) {{
    btn.textContent = theme === 'dark' ? 'Light' : 'Dark';
    btn.addEventListener('click', () => {{
      const next = document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
      applyTheme(next);
    }});
  }}
  drawAll();
}})();
</script>
</body></html>"""


def main():
    seed()
    OUTPUT.mkdir(parents=True, exist_ok=True)
    html = build_html(snapshot())
    out = OUTPUT / "dashboard.html"
    out.write_text(html, encoding="utf-8")
    print(str(out))


if __name__ == "__main__":
    main()
