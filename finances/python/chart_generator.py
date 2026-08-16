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
            name = by_id.get(b.get("category_id", ""), {}).get(
                "name", b.get("category_id")
            )
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
    body {{ font-family: Segoe UI, sans-serif; background:#121212; color:#eee; margin:0; font-size:13px; }}
    header {{ padding:8px 14px; background:#1a1a1a; border-bottom:1px solid #2a2a2a; }}
    header h1 {{ margin:0; font-size:15px; font-weight:600; }}
    main {{ padding:10px 12px 16px; }}
    .kpis {{ display:grid; grid-template-columns:repeat(4,1fr); gap:8px; margin-bottom:8px; }}
    .kpi {{ background:#1e1e1e; padding:8px 10px; border-radius:6px; }}
    .kpi .lbl {{ color:#888; font-size:11px; text-transform:uppercase; letter-spacing:.03em; }}
    .kpi .val {{ font-size:16px; margin-top:2px; font-weight:600; }}
    .dim {{ color:#777; font-size:12px; font-weight:400; }}
    .pos {{ color:#2ecc71; }} .neg {{ color:#e74c3c; }}
    .charts {{ display:grid; grid-template-columns:1fr 1fr; gap:8px; margin-bottom:8px; }}
    .split {{ display:grid; grid-template-columns:1fr 1fr; gap:8px; }}
    .panel {{ background:#1e1e1e; padding:8px 10px; border-radius:6px; margin-bottom:0; }}
    .panel-slim {{ margin-bottom:8px; }}
    .panel h2 {{ margin:0 0 4px; font-size:12px; color:#bbb; font-weight:600; }}
    .chart {{ height:260px; }}
    .chart-cell {{ min-width:0; }}
    .note {{ background:#3d2b00; color:#f1c40f; padding:4px 8px; border-radius:4px; margin-bottom:6px; font-size:12px; }}
    .note.ok {{ background:#143d27; color:#2ecc71; }}
    .perf-line {{ color:#ccc; font-size:12px; line-height:1.4; }}
    .bar-row {{ margin:6px 0 8px; }}
    .bar-head {{ display:flex; justify-content:space-between; gap:8px; font-size:12px; margin-bottom:3px; }}
    .bar-track {{ height:7px; background:#333; border-radius:4px; overflow:hidden; }}
    .bar-fill {{ height:100%; border-radius:4px; }}
    .bar-meta {{ color:#777; font-size:11px; margin-top:2px; }}
    .empty {{ color:#666; margin:0; }}
    @media (max-width:900px) {{
      .kpis, .charts, .split {{ grid-template-columns:1fr; }}
    }}
  </style>
</head>
<body>
<header>
  <h1>Finance cockpit · {data['year_month']}</h1>
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
const L = {{paper_bgcolor:'#1e1e1e', plot_bgcolor:'#1e1e1e', font:{{color:'#ccc', size:11}},
  margin:{{t:28,b:36,l:42,r:16}}, height:260, legend:{{orientation:'h', y:1.12, font:{{size:10}}}}}};
function pie(id, spec) {{
  const el = document.getElementById(id);
  if (!el) return;
  if (!spec.values.length) {{ el.innerHTML = '<p class="empty">No data</p>'; return; }}
  Plotly.newPlot(id, [{{
    type:'pie', labels:spec.labels, values:spec.values, marker:{{colors:spec.colors}},
    customdata: spec.custom, textfont:{{size:10}},
    hovertemplate: '%{{label}}<br>%{{percent}}<br>%{{customdata}}<extra></extra>'
  }}], Object.assign({{}}, L, {{showlegend:true}}), {{responsive:true, displayModeBar:false}});
}}
function drawAll() {{
  pie('pieExp', DATA.expensePie);
  pie('pieInc', DATA.incomePie);
  const names = DATA.spentMain.map(x => x.name);
  const vals = DATA.spentMain.map(x => x.value);
  Plotly.newPlot('barCat', [{{type:'bar', x:names, y:vals, marker:{{color:'#e67e22'}},
    hovertemplate:'%{{x}}<br>R$ %{{y:.2f}}<extra></extra>'}}],
    Object.assign({{}}, L, {{showlegend:false}}), {{responsive:true, displayModeBar:false}});
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
drawAll();
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
