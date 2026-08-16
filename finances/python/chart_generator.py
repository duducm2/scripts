"""Build local Plotly cockpit dashboard (all charts on one page)."""

from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from data_aggregator import (  # noqa: E402
    OUTPUT,
    cockpit_raw,
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
    raw = cockpit_raw(data)
    cur = raw["currentMonth"]
    notes = data["notifications"] if widget_on(s, "ShowNotifications") else []
    note_html = (
        "".join(f'<div class="note">{n}</div>' for n in notes)
        or '<div class="note ok">No alerts</div>'
    )

    cards_html = ""
    if widget_on(s, "ShowBalance"):
        cards_html = f"""
        <div class="kpis" id="kpiRow">
          <div class="kpi"><div class="lbl">Balance</div><div class="val" id="kpiBalance">{format_brl(data['balance'])}</div></div>
          <div class="kpi"><div class="lbl">Incomes</div><div class="val pos" id="kpiIncome">{format_brl(data['totals']['income'])}</div></div>
          <div class="kpi"><div class="lbl">Expenses</div><div class="val neg" id="kpiExpense">{format_brl(data['totals']['expense'])}</div></div>
          <div class="kpi"><div class="lbl">Card avail.</div><div class="val" id="kpiCard">{format_brl(data['card_available'])} <span class="dim">/ {format_brl(data['card_limit'])}</span></div></div>
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
        multi = data.get("period_multi")
        prev_lbl = "Prior period" if multi else "Saved last month"
        cur_lbl = "This period" if multi else "This month"
        perf_html = f"""
        <div class="panel panel-slim">
          <div class="perf-line" id="perfLine">{prev_lbl} {format_brl(prev_b)} · {cur_lbl} {format_brl(cur_b)} ({vs:+.0f}%) · Kept {data['saved_pct']:.0f}% · Top: {top}</div>
        </div>"""

    pies_html = ""
    if widget_on(s, "ShowPies"):
        pies_html = """
          <div class="panel chart-cell"><h2>Expenses by category</h2><div id="pieExp" class="chart"></div></div>
          <div class="panel chart-cell"><h2>Incomes by category</h2><div id="pieInc" class="chart"></div></div>"""

    year_lbl = data.get("period_year") or cur[:4]
    reports_html = f"""
          <div class="panel chart-cell"><h2>Spent per main category</h2><div id="barCat" class="chart"></div></div>
          <div class="panel chart-cell"><h2>Monthly balance</h2><div id="barBal" class="chart"></div></div>
          <div class="panel chart-cell chart-span"><h2 id="annualTitle">Annual cash flow ({year_lbl})</h2><div id="lineYear" class="chart"></div></div>"""

    goals_html = ""
    if widget_on(s, "ShowGoals"):
        items = []
        for g in data["goals"]:
            name = g.get("name", "")
            gcur = parse_decimal(g.get("current_amount"))
            tgt = parse_decimal(g.get("target_amount"))
            pct = (gcur / tgt * 100) if tgt > 0 else (100.0 if gcur > 0 else 0.0)
            width = min(pct, 100.0)
            rem = tgt - gcur
            tdate = (g.get("target_date") or "").strip()
            today = __import__("datetime").datetime.now().strftime("%Y-%m-%d")
            if tgt > 0 and gcur >= tgt:
                fill = "#f1c40f"
                cap = "Reached"
            elif tdate and tdate < today and (tgt <= 0 or gcur < tgt):
                fill = "#7f8c8d"
                cap = f"Rem {format_brl(rem)}" if rem > 0 else "Reached"
            else:
                fill = "#3498db"
                cap = f"Rem {format_brl(rem)}" if rem > 0 else "Reached"
            items.append(
                f'<div class="bar-row">'
                f'<div class="bar-head"><span>{name}</span>'
                f"<span>{format_brl(gcur)} / {format_brl(tgt)} · {pct:.0f}%</span></div>"
                f'<div class="bar-track"><div class="bar-fill" style="width:{width:.1f}%;background:{fill}"></div></div>'
                f'<div class="bar-meta">{cap}</div>'
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
          <div id="budgetsBody">{''.join(items) or '<p class="empty">No budgets this period</p>'}</div></div>"""

    payload = {
        "expensePie": pie_spec(data["expense_pie"]),
        "incomePie": pie_spec(data["income_pie"]),
        "series": data["series"],
        "annual": data["annual"],
        "spentMain": [{"name": n, "value": v} for n, v, _ in data["expense_pie"]],
    }
    payload_json = json.dumps(payload, ensure_ascii=False)
    raw_json = json.dumps(raw, ensure_ascii=False)

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
      --ctrl-bg: #2c2c2c;
      --ctrl-fg: #eee;
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
      --ctrl-bg: #fff;
      --ctrl-fg: #1a1a1a;
    }}
    body {{ font-family: Segoe UI, sans-serif; background:var(--bg); color:var(--text); margin:0; font-size:13px; }}
    header {{ padding:8px 14px; background:var(--header); border-bottom:1px solid var(--border);
      display:flex; align-items:center; justify-content:space-between; gap:12px; flex-wrap:wrap; }}
    header h1 {{ margin:0; font-size:15px; font-weight:600; }}
    .period-controls {{ display:flex; align-items:center; gap:8px; flex-wrap:wrap; }}
    .period-controls label {{ display:flex; align-items:center; gap:4px; font-size:12px; color:var(--muted); }}
    .period-controls input[type="month"] {{
      background:var(--ctrl-bg); color:var(--ctrl-fg); border:1px solid var(--border);
      border-radius:6px; padding:4px 8px; font-size:12px;
    }}
    .period-controls button {{
      background:var(--toggle-bg); color:var(--toggle-fg); border:1px solid var(--border);
      border-radius:6px; padding:5px 10px; font-size:12px; cursor:pointer;
    }}
    .period-controls button:hover {{ filter:brightness(1.08); }}
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
    .chart-span {{ grid-column:1 / -1; }}
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
  <h1>Finance cockpit · <span id="periodTitle">{data['year_month']}</span></h1>
  <div class="period-controls">
    <label>From <input type="month" id="periodFrom" value="{cur}"/></label>
    <label>To <input type="month" id="periodTo" value="{cur}"/></label>
    <button type="button" id="periodApply">Apply</button>
    <button type="button" id="themeToggle" aria-label="Toggle theme">Light</button>
  </div>
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
const RAW = {raw_json};
const THEME_KEY = 'finance-cockpit-theme';
const MONTH_NAMES = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

function parseDecimal(value) {{
  if (value == null) return 0;
  let s = String(value).trim().replace('R$','').replace(/\\s/g,'');
  if (!s) return 0;
  let sign = 1;
  if (s[0] === '-' || s[0] === '+') {{
    if (s[0] === '-') sign = -1;
    s = s.slice(1);
  }}
  if (s.includes(',') && s.includes('.')) {{
    if (s.lastIndexOf(',') > s.lastIndexOf('.')) s = s.replace(/\\./g,'').replace(',', '.');
    else s = s.replace(/,/g,'');
  }} else if (s.includes(',')) {{
    const parts = s.split(',');
    if (parts.length === 2 && parts[1].length <= 2) s = s.replace(/\\./g,'').replace(',', '.');
    else s = s.replace(/,/g,'');
  }}
  const n = parseFloat(s);
  return Number.isFinite(n) ? sign * n : 0;
}}
function formatBrl(num) {{
  const neg = num < 0;
  const n = Math.abs(num);
  const formatted = n.toLocaleString('pt-BR', {{minimumFractionDigits:2, maximumFractionDigits:2}});
  return (neg ? '-R$ ' : 'R$ ') + formatted;
}}
function monthShift(ym, delta) {{
  let y = parseInt(ym.slice(0,4), 10);
  let m = parseInt(ym.slice(5,7), 10) + delta;
  while (m > 12) {{ m -= 12; y += 1; }}
  while (m < 1) {{ m += 12; y -= 1; }}
  return y.toString().padStart(4,'0') + '-' + String(m).padStart(2,'0');
}}
function monthsInRange(from, to) {{
  if (!from || !to) return [];
  let a = from, b = to;
  if (a.slice(0,4) !== b.slice(0,4)) {{
    a = b.slice(0,4) + '-' + a.slice(5);
  }}
  if (a > b) {{ const t = a; a = b; b = t; }}
  const out = [];
  let cur = a;
  while (cur <= b) {{
    out.push(cur);
    cur = monthShift(cur, 1);
  }}
  return out;
}}
function periodLabel(months) {{
  if (!months.length) return RAW.currentMonth;
  if (months.length === 1) return months[0];
  const year = months[0].slice(0,4);
  return year + ' (' + months.map(m => MONTH_NAMES[parseInt(m.slice(5,7),10)-1]).join(', ') + ')';
}}
function catById() {{
  const map = {{}};
  for (const c of RAW.categories) map[c.id] = c;
  return map;
}}
function mainCategoryId(cid, byId) {{
  const c = byId[cid];
  if (!c) return cid;
  const parent = (c.parent_id || '').trim();
  return parent || cid;
}}
function catLabel(row, fallback) {{
  if (!row) return fallback || 'Uncategorized';
  const icon = (row.icon || '').trim();
  const name = row.name || fallback || '';
  return icon && name ? icon + ' ' + name : (name || icon || fallback || 'Uncategorized');
}}
function monthTotals(months) {{
  let income = 0, expense = 0;
  const set = new Set(months);
  for (const t of RAW.transactions) {{
    const ym = String(t.date || '').slice(0,7);
    if (!set.has(ym)) continue;
    const amt = parseDecimal(t.amount);
    if (t.type === 'income') income += amt;
    else if (t.type === 'expense' || t.type === 'card_expense') expense += amt;
  }}
  return {{income, expense, balance: income - expense}};
}}
function byCategory(months, types) {{
  const byId = catById();
  const set = new Set(months);
  const totals = {{}};
  for (const t of RAW.transactions) {{
    const ym = String(t.date || '').slice(0,7);
    if (!set.has(ym)) continue;
    if (!types.has(t.type)) continue;
    const cid = mainCategoryId(t.category_id || '', byId);
    totals[cid] = (totals[cid] || 0) + parseDecimal(t.amount);
  }}
  const rows = Object.entries(totals).map(([cid, amt]) => {{
    const row = byId[cid];
    return {{name: catLabel(row, cid), value: amt, color: (row && row.color) || '#7F8C8D'}};
  }});
  rows.sort((a,b) => b.value - a.value);
  return rows;
}}
function pieFromRows(rows) {{
  return {{
    labels: rows.map(r => r.name),
    values: rows.map(r => Math.round(r.value * 100) / 100),
    colors: rows.map(r => r.color),
    custom: rows.map(r => formatBrl(r.value))
  }};
}}
function seriesFor(months) {{
  return months.map(ym => {{
    const tot = monthTotals([ym]);
    return {{month: ym, income: tot.income, expense: tot.expense, balance: tot.balance}};
  }});
}}
function annualFor(year) {{
  const out = [];
  for (let m = 1; m <= 12; m++) {{
    const ym = year + '-' + String(m).padStart(2,'0');
    const tot = monthTotals([ym]);
    out.push({{month: ym, label: MONTH_NAMES[m-1], income: tot.income, expense: tot.expense, balance: tot.balance}});
  }}
  return out;
}}
function priorMonths(months) {{
  if (!months.length) return [];
  const sorted = [...months].sort();
  const n = sorted.length;
  const earliest = sorted[0];
  const out = [];
  for (let i = n; i >= 1; i--) out.push(monthShift(earliest, -i));
  return out;
}}
function aggregateBudgets(months) {{
  const set = new Set(months);
  const planned = {{}}, spent = {{}};
  for (const b of RAW.budgets) {{
    if (!set.has(b.year_month)) continue;
    const cid = b.category_id || '';
    planned[cid] = (planned[cid] || 0) + parseDecimal(b.planned_amount);
    spent[cid] = (spent[cid] || 0) + parseDecimal(b.spent_amount);
  }}
  const ids = [...new Set([...Object.keys(planned), ...Object.keys(spent)])].sort();
  return ids.map(cid => ({{category_id: cid, planned: planned[cid] || 0, spent: spent[cid] || 0}}));
}}
function renderBudgets(rows) {{
  const el = document.getElementById('budgetsBody');
  if (!el) return;
  if (!rows.length) {{
    el.innerHTML = '<p class="empty">No budgets this period</p>';
    return;
  }}
  const byId = catById();
  el.innerHTML = rows.map(b => {{
    const name = catLabel(byId[b.category_id], b.category_id);
    const rem = b.planned - b.spent;
    const pct = b.planned > 0 ? (b.spent / b.planned * 100) : (b.spent > 0 ? 100 : 0);
    const width = Math.min(pct, 100);
    const over = b.spent > b.planned && b.planned > 0;
    const fill = over ? '#e74c3c' : '#2ecc71';
    const cap = over ? ('Exceeded ' + formatBrl(-rem)) : ('Remain ' + formatBrl(rem));
    return '<div class="bar-row">'
      + '<div class="bar-head"><span>' + name + '</span><span>'
      + formatBrl(b.spent) + ' / ' + formatBrl(b.planned) + ' · ' + pct.toFixed(0) + '%</span></div>'
      + '<div class="bar-track"><div class="bar-fill" style="width:' + width.toFixed(1) + '%;background:' + fill + '"></div></div>'
      + '<div class="bar-meta">' + cap + '</div></div>';
  }}).join('');
}}
function applyPeriod() {{
  const fromEl = document.getElementById('periodFrom');
  const toEl = document.getElementById('periodTo');
  let from = fromEl.value;
  let to = toEl.value;
  if (!from || !to) return;
  if (from.slice(0,4) !== to.slice(0,4)) {{
    from = to.slice(0,4) + '-' + from.slice(5);
    fromEl.value = from;
    if (from > to) {{ from = to; fromEl.value = from; }}
  }}
  if (from > to) {{
    const t = from; from = to; to = t;
    fromEl.value = from; toEl.value = to;
  }}
  const months = monthsInRange(from, to);
  const year = to.slice(0,4);
  const tot = monthTotals(months);
  const prev = monthTotals(priorMonths(months));
  const expRows = byCategory(months, new Set(['expense','card_expense']));
  const incRows = byCategory(months, new Set(['income']));
  const multi = months.length > 1;
  DATA.expensePie = pieFromRows(expRows);
  DATA.incomePie = pieFromRows(incRows);
  DATA.spentMain = expRows.map(r => ({{name: r.name, value: r.value}}));
  DATA.series = seriesFor(months);
  DATA.annual = annualFor(year);

  const title = document.getElementById('periodTitle');
  if (title) title.textContent = periodLabel(months);
  const annualTitle = document.getElementById('annualTitle');
  if (annualTitle) annualTitle.textContent = 'Annual cash flow (' + year + ')';
  const kpiIncome = document.getElementById('kpiIncome');
  const kpiExpense = document.getElementById('kpiExpense');
  if (kpiIncome) kpiIncome.textContent = formatBrl(tot.income);
  if (kpiExpense) kpiExpense.textContent = formatBrl(tot.expense);

  const perf = document.getElementById('perfLine');
  if (perf) {{
    const vs = prev.balance ? ((tot.balance - prev.balance) / Math.abs(prev.balance) * 100) : 0;
    const saved = tot.income ? (tot.balance / tot.income * 100) : 0;
    const top = expRows.slice(0,4).map(r => r.name + ' ' + formatBrl(r.value)).join(' · ') || 'No expenses';
    const prevLbl = multi ? 'Prior period' : 'Saved last month';
    const curLbl = multi ? 'This period' : 'This month';
    perf.textContent = prevLbl + ' ' + formatBrl(prev.balance) + ' · ' + curLbl + ' ' + formatBrl(tot.balance)
      + ' (' + (vs >= 0 ? '+' : '') + vs.toFixed(0) + '%) · Kept ' + saved.toFixed(0) + '% · Top: ' + top;
  }}
  renderBudgets(aggregateBudgets(months));
  drawAll();
}}
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
  const barCat = document.getElementById('barCat');
  if (barCat) {{
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
  }}
  const barBal = document.getElementById('barBal');
  if (barBal) {{
    const months = DATA.series.map(x => x.month);
    const bals = DATA.series.map(x => x.balance);
    const colors = bals.map(v => v >= 0 ? '#27ae60' : '#c0392b');
    Plotly.newPlot('barBal', [{{type:'bar', x:months, y:bals, marker:{{color:colors}},
      hovertemplate:'%{{x}}<br>R$ %{{y:.2f}}<extra></extra>'}}],
      Object.assign({{}}, L, {{showlegend:false}}), {{responsive:true, displayModeBar:false}});
  }}
  const lineYear = document.getElementById('lineYear');
  if (lineYear) {{
    Plotly.newPlot('lineYear', [
      {{type:'scatter', mode:'lines+markers', name:'In', x:DATA.annual.map(x=>x.label), y:DATA.annual.map(x=>x.income), line:{{color:'#2ecc71'}}}},
      {{type:'scatter', mode:'lines+markers', name:'Out', x:DATA.annual.map(x=>x.label), y:DATA.annual.map(x=>x.expense), line:{{color:'#e74c3c'}}}},
      {{type:'scatter', mode:'lines+markers', name:'Bal', x:DATA.annual.map(x=>x.label), y:DATA.annual.map(x=>x.balance), line:{{color:'#f1c40f'}}}}
    ], L, {{responsive:true, displayModeBar:false}});
  }}
}}
function applyTheme(theme) {{
  document.documentElement.setAttribute('data-theme', theme);
  localStorage.setItem(THEME_KEY, theme);
  const btn = document.getElementById('themeToggle');
  if (btn) btn.textContent = theme === 'dark' ? 'Light' : 'Dark';
  drawAll();
}}
(function init() {{
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
  const applyBtn = document.getElementById('periodApply');
  if (applyBtn) applyBtn.addEventListener('click', applyPeriod);
  const fromEl = document.getElementById('periodFrom');
  const toEl = document.getElementById('periodTo');
  if (fromEl) fromEl.value = RAW.currentMonth;
  if (toEl) toEl.value = RAW.currentMonth;
  if (fromEl) fromEl.addEventListener('change', () => {{
    if (toEl && fromEl.value.slice(0,4) !== toEl.value.slice(0,4))
      toEl.value = fromEl.value.slice(0,4) + '-' + toEl.value.slice(5);
  }});
  if (toEl) toEl.addEventListener('change', () => {{
    if (fromEl && fromEl.value.slice(0,4) !== toEl.value.slice(0,4))
      fromEl.value = toEl.value.slice(0,4) + '-' + fromEl.value.slice(5);
  }});
  drawAll();
}})();
</script>
</body></html>"""


def main():
    import argparse
    from data_aggregator import current_month

    seed()
    parser = argparse.ArgumentParser(description="Build finance cockpit dashboard")
    parser.add_argument("--year", default=None, help="Calendar year for annual chart")
    parser.add_argument(
        "--months",
        default=None,
        help="Comma-separated YYYY-MM months to aggregate",
    )
    args = parser.parse_args()
    year = args.year or current_month()[:4]
    if args.months:
        months = [m.strip() for m in args.months.split(",") if m.strip()]
    else:
        months = [current_month()]
    OUTPUT.mkdir(parents=True, exist_ok=True)
    html = build_html(snapshot(months=months, year=year))
    out = OUTPUT / "dashboard.html"
    out.write_text(html, encoding="utf-8")
    print(str(out))


if __name__ == "__main__":
    main()
