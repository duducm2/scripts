"""Build local Plotly dashboard + reports HTML."""

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


def build_html(data: dict, open_reports: bool) -> str:
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
          <div class="kpi"><div class="lbl">Current balance</div><div class="val">{format_brl(data['balance'])}</div></div>
          <div class="kpi"><div class="lbl">Incomes</div><div class="val pos">{format_brl(data['totals']['income'])}</div></div>
          <div class="kpi"><div class="lbl">Expenses</div><div class="val neg">{format_brl(data['totals']['expense'])}</div></div>
          <div class="kpi"><div class="lbl">Card available</div><div class="val">{format_brl(data['card_available'])} / {format_brl(data['card_limit'])}</div></div>
        </div>"""

    perf_html = ""
    if widget_on(s, "ShowPerformance"):
        prev_b = data["prev_totals"]["balance"]
        cur_b = data["totals"]["balance"]
        vs = ((cur_b - prev_b) / abs(prev_b) * 100) if prev_b else 0
        top = (
            "".join(
                f"<li>{name} — {format_brl(amt)}</li>"
                for name, amt, _ in data["top_expenses"]
            )
            or "<li>No expenses this month</li>"
        )
        perf_html = f"""
        <div class="panel">
          <h2>Performance</h2>
          <p>Saved last month: {format_brl(prev_b)} · This month: {format_brl(cur_b)} ({vs:+.0f}% vs prior)</p>
          <p>Share of income kept: {data['saved_pct']:.1f}%</p>
          <h3>Top spending categories</h3>
          <ul>{top}</ul>
        </div>"""

    pies_html = ""
    if widget_on(s, "ShowPies"):
        pies_html = """
        <div class="charts">
          <div class="panel"><h2>Expenses by category</h2><div id="pieExp"></div></div>
          <div class="panel"><h2>Incomes by category</h2><div id="pieInc"></div></div>
        </div>"""

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
                cap = f"Remaining {format_brl(rem)}" if rem > 0 else "Reached"
            else:
                # in_progress / paused
                fill = "#3498db"
                cap = f"Remaining {format_brl(rem)}" if rem > 0 else "Reached"
            items.append(
                f'<div class="bar-row">'
                f'<div class="bar-head"><span>{name}</span>'
                f"<span>{format_brl(cur)} / {format_brl(tgt)} · {pct:.0f}%</span></div>"
                f'<div class="bar-track"><div class="bar-fill" style="width:{width:.1f}%;background:{fill}"></div></div>'
                f'<div class="bar-meta">{status} · {cap} · Current {format_brl(cur)} · Target {format_brl(tgt)}</div>'
                f"</div>"
            )
        goals_html = f"""
        <div class="panel"><h2>Goals</h2>
          {''.join(items) or '<p>No goals</p>'}</div>"""

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
                f'<div class="bar-meta">{cap} · Planned {format_brl(p)} · Spent {format_brl(sp)} · Remaining {format_brl(rem)}</div>'
                f"</div>"
            )
        bud_html = f"""
        <div class="panel"><h2>Budgets</h2>
          {''.join(items) or '<p>No budgets this month</p>'}</div>"""

    payload = {
        "expensePie": pie_spec(data["expense_pie"]),
        "incomePie": pie_spec(data["income_pie"]),
        "series": data["series"],
        "annual": data["annual"],
        "spentMain": [{"name": n, "value": v} for n, v, _ in data["expense_pie"]],
        "openReports": open_reports,
    }
    payload_json = json.dumps(payload, ensure_ascii=False)

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8"/>
  <title>Finance dashboard</title>
  <script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
  <style>
    body {{ font-family: Segoe UI, sans-serif; background:#121212; color:#eee; margin:0; }}
    header {{ padding:16px 24px; background:#1e1e1e; display:flex; gap:16px; align-items:center; }}
    header h1 {{ margin:0; font-size:20px; }}
    .tabs button {{ background:#2c2c2c; color:#eee; border:0; padding:8px 14px; margin-right:8px; cursor:pointer; border-radius:6px; }}
    .tabs button.active {{ background:#f1c40f; color:#111; }}
    main {{ padding:20px 24px; }}
    .kpis {{ display:grid; grid-template-columns:repeat(4,1fr); gap:12px; margin-bottom:16px; }}
    .kpi {{ background:#1e1e1e; padding:14px; border-radius:10px; }}
    .kpi .lbl {{ color:#aaa; font-size:12px; }}
    .kpi .val {{ font-size:22px; margin-top:6px; }}
    .pos {{ color:#2ecc71; }} .neg {{ color:#e74c3c; }}
    .charts {{ display:grid; grid-template-columns:1fr 1fr; gap:12px; }}
    .panel {{ background:#1e1e1e; padding:14px; border-radius:10px; margin-bottom:12px; }}
    table {{ width:100%; border-collapse:collapse; }}
    th,td {{ text-align:left; padding:6px 8px; border-bottom:1px solid #333; }}
    .note {{ background:#3d2b00; color:#f1c40f; padding:8px 12px; border-radius:8px; margin-bottom:8px; }}
    .note.ok {{ background:#143d27; color:#2ecc71; }}
    .hidden {{ display:none; }}
    .bar-row {{ margin:12px 0 16px; }}
    .bar-head {{ display:flex; justify-content:space-between; gap:12px; font-size:14px; margin-bottom:6px; }}
    .bar-track {{ height:10px; background:#333; border-radius:6px; overflow:hidden; }}
    .bar-fill {{ height:100%; border-radius:6px; }}
    .bar-meta {{ color:#888; font-size:12px; margin-top:4px; }}
  </style>
</head>
<body>
<header>
  <h1>Finance · {data['year_month']}</h1>
  <div class="tabs">
    <button id="tabDash" class="active">Dashboard</button>
    <button id="tabRep">Reports</button>
  </div>
</header>
<main>
  <section id="dash">
    {note_html}
    {cards_html}
    {perf_html}
    {pies_html}
    {goals_html}
    {bud_html}
  </section>
  <section id="rep" class="hidden">
    <div class="panel"><h2>Spent per main category</h2><div id="barCat"></div></div>
    <div class="panel"><h2>Monthly balance</h2><div id="barBal"></div></div>
    <div class="panel"><h2>Cash flow</h2><div id="lineCf"></div></div>
    <div class="panel"><h2>Annual cash flow</h2><div id="lineYear"></div></div>
  </section>
</main>
<script>
const DATA = {payload_json};
function pie(id, spec, title) {{
  if (!spec.values.length) {{
    document.getElementById(id).innerHTML = '<p>No data</p>';
    return;
  }}
  Plotly.newPlot(id, [{{
    type:'pie', labels:spec.labels, values:spec.values, marker:{{colors:spec.colors}},
    customdata: spec.custom,
    hovertemplate: '%{{label}}<br>%{{percent}}<br>%{{customdata}}<extra></extra>'
  }}], {{paper_bgcolor:'#1e1e1e', font:{{color:'#eee'}}, title:title, showlegend:true}}, {{responsive:true}});
}}
function drawDash() {{
  if (document.getElementById('pieExp')) pie('pieExp', DATA.expensePie, '');
  if (document.getElementById('pieInc')) pie('pieInc', DATA.incomePie, '');
}}
function drawRep() {{
  const names = DATA.spentMain.map(x => x.name);
  const vals = DATA.spentMain.map(x => x.value);
  Plotly.newPlot('barCat', [{{type:'bar', x:names, y:vals, marker:{{color:'#e67e22'}},
    hovertemplate:'%{{x}}<br>R$ %{{y:.2f}}<extra></extra>'}}],
    {{paper_bgcolor:'#1e1e1e', plot_bgcolor:'#1e1e1e', font:{{color:'#eee'}}}});
  const months = DATA.series.map(x => x.month);
  const bals = DATA.series.map(x => x.balance);
  const colors = bals.map(v => v >= 0 ? '#27ae60' : '#c0392b');
  Plotly.newPlot('barBal', [{{type:'bar', x:months, y:bals, marker:{{color:colors}},
    hovertemplate:'%{{x}}<br>R$ %{{y:.2f}}<extra></extra>'}}],
    {{paper_bgcolor:'#1e1e1e', plot_bgcolor:'#1e1e1e', font:{{color:'#eee'}}}});
  Plotly.newPlot('lineCf', [
    {{type:'scatter', mode:'lines+markers', name:'Incomes', x:months, y:DATA.series.map(x=>x.income), line:{{color:'#2ecc71'}}}},
    {{type:'scatter', mode:'lines+markers', name:'Expenses', x:months, y:DATA.series.map(x=>x.expense), line:{{color:'#e74c3c'}}}}
  ], {{paper_bgcolor:'#1e1e1e', plot_bgcolor:'#1e1e1e', font:{{color:'#eee'}}}});
  Plotly.newPlot('lineYear', [
    {{type:'scatter', mode:'lines+markers', name:'Incomes', x:DATA.annual.map(x=>x.label), y:DATA.annual.map(x=>x.income), line:{{color:'#2ecc71'}}}},
    {{type:'scatter', mode:'lines+markers', name:'Expenses', x:DATA.annual.map(x=>x.label), y:DATA.annual.map(x=>x.expense), line:{{color:'#e74c3c'}}}},
    {{type:'scatter', mode:'lines+markers', name:'Balance', x:DATA.annual.map(x=>x.label), y:DATA.annual.map(x=>x.balance), line:{{color:'#f1c40f'}}}}
  ], {{paper_bgcolor:'#1e1e1e', plot_bgcolor:'#1e1e1e', font:{{color:'#eee'}}}});
}}
function show(tab) {{
  document.getElementById('dash').classList.toggle('hidden', tab!=='dash');
  document.getElementById('rep').classList.toggle('hidden', tab!=='rep');
  document.getElementById('tabDash').classList.toggle('active', tab==='dash');
  document.getElementById('tabRep').classList.toggle('active', tab==='rep');
  if (tab==='rep') drawRep();
}}
document.getElementById('tabDash').onclick = () => show('dash');
document.getElementById('tabRep').onclick = () => show('rep');
drawDash();
if (DATA.openReports) show('rep');
</script>
</body></html>"""


def main():
    open_reports = "--reports" in sys.argv
    seed()
    OUTPUT.mkdir(parents=True, exist_ok=True)
    html = build_html(snapshot(), open_reports)
    out = OUTPUT / "dashboard.html"
    out.write_text(html, encoding="utf-8")
    print(str(out))


if __name__ == "__main__":
    main()
