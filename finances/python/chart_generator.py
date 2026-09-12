"""Build local Plotly cockpit dashboard (all charts on one page)."""

from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from data_aggregator import (  # noqa: E402
    OUTPUT,
    configure_paths,
    cockpit_raw,
    format_brl,
    parse_decimal,
    snapshot,
    widget_on,
    cat_index,
    cat_label,
)
from seed_from_ini import seed  # noqa: E402
import data_aggregator as _agg  # noqa: E402


def pie_spec(rows):
    # rows: (name, amount, color) — category id comes from client-side rebuild
    return {
        "labels": [r[0] for r in rows],
        "values": [round(r[1], 2) for r in rows],
        "colors": [r[2] for r in rows],
        "custom": [format_brl(r[1]) for r in rows],
        "categoryIds": [""] * len(rows),
    }


def build_html(data: dict) -> str:
    s = data["settings"]
    raw = cockpit_raw(data)
    cur = raw["currentMonth"]
    date_from = raw.get("dateFrom") or raw.get("monthStart") or (cur + "-01")
    date_to = raw.get("dateTo") or raw.get("today") or date_from
    notes = data["notifications"] if widget_on(s, "ShowNotifications") else []
    notif_col = ""
    if widget_on(s, "ShowNotifications"):
        if notes:
            items = "".join(f'<div class="note">{n}</div>' for n in notes)
        else:
            items = '<div class="note ok">No alerts</div>'
        notif_col = f"""
          <div class="panel notif-col" id="notificationsPanel">
            <h2>Notifications</h2>
            <div id="notificationsBody">{items}</div>
          </div>"""

    cards_html = ""
    card_avail_block = ""
    card_lim = float(data.get("card_limit") or 0)
    card_sp = float(data.get("card_spent") or 0)
    util_pct = (
        (card_sp / card_lim * 100.0)
        if card_lim > 0
        else (100.0 if card_sp > 0 else 0.0)
    )
    util_width = min(util_pct, 100.0)
    util_over = card_lim > 0 and card_sp > card_lim
    util_fill = "#e74c3c" if util_over else ("#f39c12" if util_pct >= 80 else "#3498db")
    mini_rows = []
    for c in data.get("cards") or []:
        name = c.get("name") or c.get("id") or "Card"
        lim = parse_decimal(c.get("limit"))
        spent = parse_decimal(c.get("current_spent"))
        avail = lim - spent
        pct = (spent / lim * 100) if lim > 0 else (100.0 if spent > 0 else 0.0)
        width = min(pct, 100.0)
        over = lim > 0 and spent > lim
        fill = "#e74c3c" if over else ("#f39c12" if pct >= 80 else "#3498db")
        mini_rows.append(
            f'<div class="kpi-card-mini">'
            f'<div class="bar-head"><span>{name}</span>'
            f'<span><span class="kpi-card-avail-amt">{format_brl(avail)}</span>'
            f' <span class="dim">{pct:.0f}%</span></span></div>'
            f'<div class="bar-track"><div class="bar-fill" style="width:{width:.1f}%;'
            f'background:{fill}"></div></div></div>'
        )
    mini_html = "".join(mini_rows)
    card_avail_inner = f"""
            <div class="lbl">Card avail.</div>
            <div class="val" id="kpiCard">{format_brl(data['card_available'])} <span class="dim">/ {format_brl(data['card_limit'])}</span></div>
            <div class="kpi-util-row" title="Credit limit utilization">
              <div class="bar-track kpi-util-track"><div class="bar-fill" id="kpiCardUtilFill" style="width:{util_width:.1f}%;background:{util_fill}"></div></div>
              <span class="kpi-util-pct" id="kpiCardUtilPct">{util_pct:.0f}%</span>
            </div>
            <div class="kpi-card-minis" id="kpiCardMinis">{mini_html}</div>"""
    show_budgets = widget_on(s, "ShowBudgets")
    if show_budgets:
        card_avail_block = f"""
            <div class="funds-card-avail" id="fundsCardAvail">
              {card_avail_inner}
            </div>"""

    bal_chip = (
        f'<div class="stat-chip"><span class="lbl">Balance</span>'
        f'<span class="val" id="kpiBalance">{format_brl(data["balance"])}</span></div>'
    )
    inc_chip = (
        f'<div class="stat-chip"><span class="lbl">Incomes</span>'
        f'<span class="val pos" id="kpiIncome">{format_brl(data["totals"]["income"])}</span></div>'
    )
    exp_chip = (
        f'<div class="stat-chip"><span class="lbl">Expenses</span>'
        f'<span class="val neg" id="kpiExpense">{format_brl(data["totals"]["expense"])}</span></div>'
    )

    perf_html = ""
    show_perf = widget_on(s, "ShowPerformance")
    perf_body = ""
    if show_perf:
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
        perf_body = (
            f'<div class="perf-line" id="perfLine">{prev_lbl} {format_brl(prev_b)}'
            f" · {cur_lbl} {format_brl(cur_b)} ({vs:+.0f}%)"
            f" · Kept {data['saved_pct']:.0f}% · Top: {top}</div>"
        )

    pie_exp_html = ""
    pie_inc_html = ""
    show_pies = widget_on(s, "ShowPies")
    if show_pies:
        pie_exp_html = f"""
          <div class="chart-cell pie-exp-cell">
          <div class="panel pie-exp-panel">
            <div class="panel-head-row">
              <h2>Expenses by category</h2>
              {exp_chip}
            </div>
            <div id="pieExp" class="chart chart-pie"></div>
          </div>
          </div>"""
        pie_inc_html = f"""
          <div class="panel chart-cell pie-inc-cell">
            <div class="panel-head-row">
              <h2>Incomes by category</h2>
              {inc_chip}
            </div>
            <div id="pieInc" class="chart chart-pie"></div>
          </div>"""
    else:
        cards_html = f"""
        <div class="panel panel-slim period-stat-strip">
          <div class="perf-chips">{inc_chip}{exp_chip}</div>
        </div>"""

    year_lbl = data.get("period_year") or cur[:4]
    card_plan_html = ""
    if widget_on(s, "ShowCardPlan"):
        card_plan_html = """
          <div class="panel chart-cell chart-span card-plan-panel">
            <div class="panel-head-row card-plan-head">
              <h2>Card expenses by pay day</h2>
              <div class="card-plan-range period-controls">
                <label>From <input type="date" id="cardPlanFrom"/></label>
                <label>To <input type="date" id="cardPlanTo"/></label>
                <button type="button" id="cardPlanApply">Apply</button>
                <button type="button" id="cardPlanReset" title="Next 2 months from today">2 months</button>
              </div>
            </div>
            <div class="card-plan-summary" id="cardPlanSummary"></div>
            <div id="lineCardPlan" class="chart chart-card-plan"></div>
          </div>"""
    reports_html = f"""
          <div class="panel chart-cell chart-bal"><h2>Daily balance</h2><div id="barBal" class="chart"></div></div>
          <div class="panel chart-cell chart-invest"><h2>Income vs investments</h2><div id="incomeVsInvest" class="chart chart-treemap"></div></div>
          <div class="panel chart-cell chart-span"><h2 id="annualTitle">Annual cash flow ({year_lbl})</h2><div id="lineYear" class="chart"></div></div>{card_plan_html}"""

    goals_col = ""
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
        goals_col = f"""
          <div class="panel goals-col">
            <h2>Goals</h2>
            {''.join(items) or '<p class="empty">No goals</p>'}
          </div>"""

    perf_col = ""
    if show_perf and perf_body:
        perf_col = f"""
          <div class="panel perf-panel perf-col">
            <h2>Performance</h2>
            {perf_body}
          </div>"""

    perf_notif_stack = ""
    if perf_col or notif_col:
        perf_notif_stack = f"""
          <div class="perf-notif-stack">
            {perf_col}
            {notif_col}
          </div>"""

    # Split row (Goals | Perf/Notifications | Accounts) is assembled after accounts HTML.
    goals_for_split = goals_col
    perf_for_split = perf_notif_stack

    bud_html = ""
    if widget_on(s, "ShowBudgets"):
        by_id = cat_index(data["categories"])
        items = []
        total_planned = 0.0
        total_spent = 0.0
        for b in data["budgets"]:
            crow = by_id.get(b.get("category_id", ""))
            name = cat_label(crow, b.get("category_id", ""))
            p = parse_decimal(b.get("planned_amount"))
            sp = parse_decimal(b.get("spent_amount"))
            total_planned += p
            total_spent += sp
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
        tot_rem = total_planned - total_spent
        tot_pct = (
            (total_spent / total_planned * 100)
            if total_planned > 0
            else (100.0 if total_spent > 0 else 0.0)
        )
        tot_width = min(tot_pct, 100.0)
        tot_over = total_spent > total_planned and total_planned > 0
        tot_fill = "#e74c3c" if tot_over else "#f1c40f"
        if items:
            bud_hint = (
                f"Exceeded {format_brl(-tot_rem)}"
                if tot_over
                else f"Remain {format_brl(tot_rem)}"
            )
            bud_summary = (
                f'<div class="bar-head"><span>Total</span>'
                f"<span>{format_brl(total_spent)} / {format_brl(total_planned)}"
                f" · {tot_pct:.0f}%</span></div>"
                f'<div class="bar-track"><div class="bar-fill" style="width:{tot_width:.1f}%;'
                f'background:{tot_fill}"></div></div>'
                f'<div class="bar-meta">{bud_hint}</div>'
            )
        else:
            bud_summary = ""
        sum_style = "" if items else ' style="display:none"'
        bal_in_funds = ""
        if not widget_on(s, "ShowAccounts"):
            bal_in_funds = bal_chip
        bud_html = f"""
        <div class="panel chart-cell budget-panel">
          <div class="budget-head">
            <h2>Budgets</h2>
            <div class="budget-controls">
              <label class="budget-month-label">Month
                <select id="budgetMonth"></select>
              </label>
              <span id="budgetSaveStatus" class="budget-save-status" aria-live="polite"></span>
            </div>
          </div>
          <div class="funds-compare" id="fundsCompare" aria-live="polite">
            <div class="funds-compare-title-row">
              <div class="funds-compare-title">Available vs planned</div>
              <div class="funds-compare-actions">
                {bal_in_funds}
                <div class="budget-calc-wrap">
                  <button type="button" class="budget-calc-btn" id="budgetCalcBtn"
                    aria-describedby="budgetCalcTip">Calculations</button>
                  <div class="budget-calc-tip" id="budgetCalcTip" role="tooltip"></div>
                </div>
              </div>
            </div>
            <div class="funds-compare-head">
              <div class="funds-compare-legend">
                <span class="funds-legend-item">
                  <span class="funds-legend-swatch available"></span>
                  Available <span class="funds-compare-amounts" id="fundsAvailableVal">—</span>
                </span>
                <span class="funds-legend-item">
                  <span class="funds-legend-swatch planned"></span>
                  Planned <span class="funds-compare-amounts" id="fundsPlannedVal">—</span>
                </span>
              </div>
            </div>
            <div class="funds-compare-track">
              <div class="funds-bar funds-bar-available" id="fundsBarAvailable" style="width:0%"></div>
              <div class="funds-bar funds-bar-planned" id="fundsBarPlanned" style="width:0%"></div>
              <span class="funds-marker funds-marker-available" id="fundsMarkAvailable" style="left:0%"></span>
              <span class="funds-marker funds-marker-planned" id="fundsMarkPlanned" style="left:0%"></span>
            </div>
            <div class="funds-compare-meta" id="fundsCompareMeta"></div>
            {card_avail_block}
          </div>
          <div class="budget-categories">
            <div class="budget-categories-title">Categories · spent / planned</div>
            <div class="bar-row bar-row-total" id="budgetsSummary"{sum_style}>{bud_summary}</div>
            <div id="budgetsBody" class="budget-body">{''.join(items) or '<p class="empty">No budgets this month</p>'}</div>
          </div>
        </div>"""

    acc_html = ""
    if widget_on(s, "ShowAccounts"):
        acc_rows = []
        for a in data.get("accounts") or []:
            acc_rows.append(
                (
                    a.get("icon") or "🏦",
                    a.get("name") or a.get("id") or "Account",
                    parse_decimal(a.get("current_balance")),
                )
            )
        acc_rows.sort(key=lambda r: r[2], reverse=True)
        acc_total = sum(r[2] for r in acc_rows)
        max_abs = max((abs(r[2]) for r in acc_rows), default=0.0)
        acc_items = []
        for icon, name, bal in acc_rows:
            width = (abs(bal) / max_abs * 100.0) if max_abs > 0 else 0.0
            share = (bal / acc_total * 100.0) if acc_total else 0.0
            fill = "#e74c3c" if bal < 0 else "#3498db"
            acc_items.append(
                f'<div class="bar-row">'
                f'<div class="bar-head"><span>{icon} {name}</span>'
                f"<span>{format_brl(bal)}</span></div>"
                f'<div class="bar-track"><div class="bar-fill" style="width:{width:.1f}%;background:{fill}"></div></div>'
                f'<div class="bar-meta">{share:.0f}% of total</div>'
                f"</div>"
            )
        acc_card_fallback = ""
        if not show_budgets:
            acc_card_fallback = f"""
          <div class="funds-card-avail" id="fundsCardAvail">
            {card_avail_inner}
          </div>"""
        acc_html = f"""
        <div class="panel accounts-col">
          <div class="panel-head-row">
            <h2>Accounts</h2>
            {bal_chip}
          </div>
          {''.join(acc_items) or '<p class="empty">No accounts</p>'}
          {acc_card_fallback}
        </div>"""

    split_parts = [p for p in (goals_for_split, perf_for_split, acc_html) if p]
    split_n = len(split_parts)
    split_html = ""
    if split_parts:
        split_html = f"""
  <div class="split split-{split_n}">
    {''.join(split_parts)}
  </div>"""

    rec_html = ""
    if widget_on(s, "ShowRecurring"):
        rec_monthly = 0.0
        for r in data.get("recurring") or []:
            rec_monthly += parse_decimal(r.get("monthly_amount"))
        n_months = max(len(data.get("period_months") or []), 1)
        rec_period = rec_monthly * n_months
        rec_income = data["totals"]["income"]
        rec_pct = (rec_period / rec_income * 100.0) if rec_income else 0.0
        rec_html = f"""
          <div class="panel chart-cell recurring-cell" id="recurringPanel">
            <h2>Recurring bills</h2>
            <div class="bar-meta" id="recurringSummary" style="margin-bottom:6px">
              Period {format_brl(rec_period)} · Income {format_brl(rec_income)}
              · Recurring is {rec_pct:.0f}% of income
              · {n_months} month{"s" if n_months != 1 else ""}</div>
            <div id="recurringVsIncome" class="chart chart-treemap"></div>
          </div>"""

    payload = {
        "expensePie": pie_spec(data["expense_pie"]),
        "incomePie": pie_spec(data["income_pie"]),
        "series": data["series"],
        "annual": data["annual"],
        "cardInstallmentRemaining": data.get("cardInstallmentRemaining")
        or {"months": [], "series": []},
    }
    payload_json = json.dumps(payload, ensure_ascii=False)
    raw_json = json.dumps(raw, ensure_ascii=False)
    charts_class = "charts" if bud_html else "charts charts-no-budget"
    inc_rec_inner = ""
    if pie_inc_html or rec_html:
        if pie_inc_html and rec_html:
            inc_rec_inner = f"""
      <div class="inc-rec-row">
        {pie_inc_html}
        {rec_html}
      </div>"""
        else:
            inc_rec_inner = pie_inc_html or rec_html
    charts_main = ""
    if pie_exp_html or inc_rec_inner:
        charts_main = f"""
    <div class="charts-main">
      {pie_exp_html}
      {inc_rec_inner}
    </div>"""
    charts_body = f"""
  <div class="charts-top">
    {bud_html}
    {charts_main}
  </div>
  <div class="charts-reports">
    {reports_html}
  </div>"""

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
    body {{
      font-family: Segoe UI, sans-serif; background:var(--bg); color:var(--text);
      margin:0; font-size:13px; min-height:100vh;
    }}
    .period-controls {{ display:flex; align-items:center; gap:8px; flex-wrap:wrap; }}
    .period-controls label {{ display:flex; align-items:center; gap:4px; font-size:12px; color:var(--muted); }}
    .period-controls input[type="date"] {{
      background:var(--ctrl-bg); color:var(--ctrl-fg); border:1px solid var(--border);
      border-radius:6px; padding:4px 8px; font-size:12px;
    }}
    .period-controls button, #catViewBack {{
      background:var(--toggle-bg); color:var(--toggle-fg); border:1px solid var(--border);
      border-radius:6px; padding:5px 10px; font-size:12px; cursor:pointer;
    }}
    .period-controls button:hover, #catViewBack:hover {{ filter:brightness(1.08); }}
    #catViewBack {{ margin-top:8px; }}
    main {{ padding:12px 16px 72px; }}
    .cockpit-footer {{
      position:fixed; left:0; right:0; bottom:0; z-index:40;
      display:flex; align-items:center; justify-content:space-between; gap:12px; flex-wrap:wrap;
      padding:8px 16px; background:var(--header); border-top:1px solid var(--border);
      box-shadow:0 -4px 16px rgba(0,0,0,.18);
    }}
    .cockpit-footer .period-label {{
      font-size:12px; color:var(--muted); font-variant-numeric:tabular-nums;
    }}
    .panel-head-row {{
      display:flex; align-items:center; justify-content:space-between; gap:8px;
      flex-wrap:wrap; margin-bottom:6px;
    }}
    .panel-head-row h2 {{ margin:0; }}
    .panel-head-actions {{
      display:flex; align-items:center; gap:8px; flex-wrap:wrap;
    }}
    .stat-chip {{
      display:flex; flex-direction:column; align-items:flex-end; gap:1px;
      padding:4px 8px; border-radius:6px; border:1px solid var(--border);
      background:var(--bg); min-width:5.5rem;
    }}
    .stat-chip .lbl {{
      color:var(--muted); font-size:10px; text-transform:uppercase; letter-spacing:.03em;
    }}
    .stat-chip .val {{
      font-size:14px; font-weight:600; font-variant-numeric:tabular-nums; line-height:1.2;
    }}
    .perf-with-chips, .period-stat-strip {{
      display:flex; align-items:center; justify-content:space-between; gap:12px; flex-wrap:wrap;
    }}
    .perf-chips {{ display:flex; align-items:center; gap:8px; flex-wrap:wrap; }}
    .kpi-card-avail, .funds-card-avail {{ display:flex; flex-direction:column; gap:2px; }}
    .funds-card-avail {{
      margin-top:10px; padding-top:10px; border-top:1px solid var(--border);
    }}
    .funds-card-avail .lbl {{
      color:var(--muted); font-size:11px; text-transform:uppercase; letter-spacing:.03em;
    }}
    .funds-card-avail .val {{
      font-size:14px; margin-top:2px; font-weight:600; font-variant-numeric:tabular-nums;
    }}
    .kpi-util-row {{
      display:flex; align-items:center; gap:8px; margin-top:4px;
    }}
    .kpi-util-track {{ flex:1; height:6px; margin:0; }}
    .kpi-util-pct {{
      font-size:11px; font-weight:600; color:var(--muted);
      font-variant-numeric:tabular-nums; min-width:2.6em; text-align:right;
    }}
    .kpi-card-minis {{ margin-top:4px; display:flex; flex-direction:column; gap:4px; }}
    .kpi-card-mini .bar-head {{
      font-size:10px; color:var(--muted); margin-bottom:1px;
    }}
    .kpi-card-mini .kpi-card-avail-amt {{
      color:var(--text); font-weight:600; font-variant-numeric:tabular-nums;
    }}
    .kpi-card-mini .bar-track {{ height:4px; }}
    .dim {{ color:var(--muted2); font-size:12px; font-weight:400; }}
    .pos {{ color:#2ecc71; }} .neg {{ color:#e74c3c; }}
    .charts {{
      display:flex; flex-direction:column; gap:10px; margin-bottom:10px;
    }}
    .charts-top {{
      display:grid; grid-template-columns:minmax(0,1fr) minmax(0,1fr); gap:10px; align-items:stretch;
    }}
    .charts-no-budget .charts-top {{ grid-template-columns:minmax(0,1fr); }}
    .charts-top > * {{
      min-width:0; min-height:0; overflow:hidden;
    }}
    .charts-main {{
      display:flex; flex-direction:column; gap:10px; min-width:0;
      height:100%; min-height:0; align-self:stretch; overflow:hidden;
    }}
    .perf-panel-slim {{
      flex:0 0 auto; min-height:0;
    }}
    .perf-panel-slim .perf-line {{
      margin:0; white-space:normal; line-height:1.45;
    }}
    .perf-notif-stack {{
      display:flex; flex-direction:column; gap:10px; min-width:0; min-height:0;
      overflow:hidden; box-sizing:border-box;
    }}
    .perf-notif-stack > .perf-col,
    .perf-notif-stack > .notif-col {{
      flex:1 1 0; min-height:0; overflow:hidden;
      display:flex; flex-direction:column; box-sizing:border-box;
    }}
    .perf-notif-stack > .perf-col .perf-line {{
      flex:1 1 auto; min-height:0; overflow:auto; margin:0;
    }}
    .perf-notif-stack > .notif-col #notificationsBody {{
      flex:1 1 auto; min-height:0; overflow:auto;
    }}
    .pie-exp-cell {{
      display:flex; flex-direction:column; min-height:0; min-width:0;
      flex:1.6 1 0; height:auto;
    }}
    .pie-exp-cell .pie-exp-panel {{
      display:flex; flex-direction:column; min-width:0; min-height:0;
      flex:1 1 auto; height:100%; overflow:hidden;
    }}
    .inc-rec-row {{
      display:grid; grid-template-columns:minmax(0,1fr) minmax(0,1fr); gap:10px; min-width:0;
      flex:1 1 0; min-height:0; align-items:stretch;
      height:auto; overflow:hidden;
    }}
    .inc-rec-row > .panel {{
      min-width:0; min-height:0; height:100%; overflow:hidden;
      display:flex; flex-direction:column;
    }}
    .charts-reports {{
      display:grid; grid-template-columns:repeat(4, 1fr); gap:10px;
    }}
    /* Bottom row: Goals | Performance+Notifications | Accounts — equal height */
    .split {{
      display:flex; flex-direction:row; align-items:stretch; gap:10px;
      width:100%;
    }}
    .split > .panel,
    .split > .perf-notif-stack {{
      flex:1 1 0; min-width:0; min-height:0; align-self:stretch;
      box-sizing:border-box;
    }}
    .split > .goals-col,
    .split > .accounts-col {{
      display:flex; flex-direction:column; overflow:auto;
    }}
    .split > .perf-notif-stack {{
      overflow:hidden;
    }}
    .panel {{ background:var(--panel); padding:10px 12px; border-radius:6px; margin-bottom:0; border:1px solid var(--border); }}
    .panel-slim {{ margin-bottom:10px; }}
    .panel h2 {{ margin:0 0 4px; font-size:12px; color:var(--heading); font-weight:600; }}
    .split .note {{ margin-bottom:6px; }}
    .split .note:last-child {{ margin-bottom:0; }}
    .chart {{ height:320px; }}
    .chart-short {{ height:220px; }}
    .chart-treemap {{ height:280px; }}
    .chart-invest .chart-treemap {{ height:320px; }}
    .recurring-cell {{
      overflow:hidden; padding-bottom:12px;
    }}
    .recurring-cell .chart-treemap {{
      flex:1 1 0; min-height:0; height:auto; max-height:100%;
      overflow:hidden; width:100%;
    }}
    .recurring-cell h2,
    .recurring-cell #recurringSummary {{
      flex:0 0 auto;
    }}
    .chart-pie {{
      height:auto; min-height:0; overflow:hidden; width:100%;
      display:grid; grid-template-rows:minmax(0,1fr) auto;
      align-content:stretch;
    }}
    .chart-pie .pie-plot {{
      grid-row:1; min-height:0; width:100%; overflow:hidden; align-self:stretch;
    }}
    .chart-pie .pie-legend {{
      grid-row:2; display:flex; flex-wrap:wrap; justify-content:center;
      gap:6px 12px; padding:8px 4px 10px; font-size:11px; color:var(--plot-font);
      line-height:1.35; min-height:2em; box-sizing:border-box;
      overflow:visible; flex-shrink:0;
    }}
    .chart-pie .pie-legend-item {{
      display:inline-flex; align-items:center; gap:5px; white-space:nowrap;
    }}
    .chart-pie .pie-legend-item i {{
      width:10px; height:10px; border-radius:2px; flex:0 0 auto;
    }}
    .chart-cell {{ min-width:0; }}
    .pie-exp-cell, .pie-inc-cell, .recurring-cell, .exp-perf-cell {{
      display:flex; flex-direction:column; min-height:0; min-width:0;
    }}
    .exp-perf-cell {{
      display:grid; grid-template-columns:minmax(0,1.45fr) minmax(0,1fr); gap:10px; align-items:stretch;
      flex:1.45 1 0; min-height:0; height:100%;
    }}
    .exp-perf-cell > .panel {{
      min-height:0; height:100%; overflow:hidden;
    }}
    .exp-perf-cell .pie-exp-panel,
    .exp-perf-cell .perf-panel {{
      display:flex; flex-direction:column; min-width:0; min-height:0; height:100%;
      overflow:hidden;
    }}
    .pie-exp-panel, .pie-inc-cell {{
      overflow:hidden;
      padding-bottom:12px;
    }}
    .pie-inc-cell, .recurring-cell {{
      height:100%; min-height:0;
    }}
    .pie-exp-cell .chart-pie, .pie-inc-cell .chart-pie,
    .exp-perf-cell .chart-pie {{
      flex:1 1 auto; min-height:0; overflow:hidden; margin-bottom:0;
    }}
    .perf-panel .perf-line {{
      flex:1; margin:0; white-space:normal; line-height:1.45;
    }}
    .budget-panel {{
      display:flex; flex-direction:column; min-width:0; min-height:0;
      overflow:auto;
    }}
    .budget-panel .budget-body {{ flex:1; overflow:visible; }}
    .budget-head {{
      display:flex; align-items:center; justify-content:space-between; gap:8px;
      flex-wrap:wrap; margin-bottom:6px;
    }}
    .budget-head h2 {{ margin:0; }}
    .budget-controls {{
      display:flex; align-items:center; gap:10px; flex-wrap:wrap;
    }}
    .budget-month-label {{
      display:flex; align-items:center; gap:6px; font-size:11px; color:var(--muted);
    }}
    .budget-month-label select {{
      background:var(--toggle-bg); color:var(--text); border:1px solid var(--border);
      border-radius:4px; padding:2px 6px; font-size:12px;
    }}
    .budget-save-status {{
      font-size:11px; color:var(--muted); min-height:1.2em;
    }}
    .budget-save-status.saving {{ color:var(--heading); }}
    .budget-save-status.saved {{ color:#2ecc71; }}
    .budget-save-status.error {{ color:#e74c3c; }}
    .budget-planned-input {{
      width:5.5rem; background:var(--toggle-bg); color:var(--text);
      border:1px solid var(--border); border-radius:4px; padding:2px 6px;
      font-size:12px; text-align:right; font-variant-numeric:tabular-nums;
    }}
    .budget-planned-input:disabled {{ opacity:.55; }}
    .budget-edit-row {{
      display:flex; align-items:center; justify-content:flex-end; gap:6px;
      font-size:12px; margin-bottom:3px;
    }}
    .budget-edit-row .spent-label {{ color:var(--muted); }}
    .funds-compare {{
      margin:0 0 12px; padding:8px 10px 10px;
      border:1px solid var(--border); border-radius:6px;
      background:var(--bg);
    }}
    .funds-compare-title-row {{
      display:flex; align-items:center; justify-content:space-between;
      gap:8px; margin-bottom:6px;
    }}
    .funds-compare-actions {{
      display:flex; align-items:center; gap:8px; flex-shrink:0;
    }}
    .funds-compare-title {{
      font-size:11px; font-weight:600; color:var(--heading);
      text-transform:uppercase; letter-spacing:.03em; margin:0;
    }}
    .funds-compare-head {{
      display:flex; justify-content:space-between; gap:10px; flex-wrap:wrap;
      font-size:12px; margin-bottom:6px;
    }}
    .funds-compare-legend {{
      display:flex; align-items:center; gap:12px; flex-wrap:wrap; color:var(--muted);
    }}
    .funds-legend-item {{
      display:inline-flex; align-items:center; gap:5px;
      font-variant-numeric:tabular-nums;
    }}
    .funds-legend-swatch {{
      width:8px; height:8px; border-radius:2px; flex-shrink:0;
    }}
    .funds-legend-swatch.available {{ background:#3498db; }}
    .funds-legend-swatch.planned {{ background:#2ecc71; }}
    .funds-compare.deficit .funds-legend-swatch.planned {{ background:#e74c3c; }}
    .funds-compare-amounts {{
      font-weight:600; font-variant-numeric:tabular-nums; color:var(--text);
    }}
    .funds-compare-track {{
      position:relative; height:16px; background:var(--track);
      border-radius:4px; overflow:visible;
    }}
    .funds-bar {{
      position:absolute; left:0; top:0; height:100%; border-radius:4px;
      transition:width .15s ease, background-color .15s ease;
    }}
    .funds-bar-available {{
      top:3px; height:10px; background:#3498db; opacity:.85; z-index:1;
    }}
    .funds-bar-planned {{
      top:0; height:16px; background:#2ecc71; opacity:.55; z-index:2;
    }}
    .funds-compare.deficit .funds-bar-planned {{ background:#e74c3c; opacity:.65; }}
    .funds-marker {{
      position:absolute; top:-3px; width:0; height:0;
      border-left:5px solid transparent; border-right:5px solid transparent;
      border-top:7px solid #3498db;
      transform:translateX(-50%); z-index:3; pointer-events:none;
      transition:left .15s ease, border-top-color .15s ease;
    }}
    .funds-marker-available {{ border-top-color:#3498db; }}
    .funds-marker-planned {{
      top:auto; bottom:-3px; border-top:none;
      border-bottom:7px solid #2ecc71;
    }}
    .funds-compare.deficit .funds-marker-planned {{ border-bottom-color:#e74c3c; }}
    .funds-compare-meta {{
      color:var(--muted2); font-size:11px; margin-top:6px;
    }}
    .funds-compare.deficit .funds-compare-meta {{ color:#e74c3c; font-weight:600; }}
    .funds-compare.ok .funds-compare-meta {{ color:#2ecc71; }}
    .budget-calc-wrap {{
      position:relative; flex-shrink:0;
    }}
    .budget-calc-btn {{
      font:inherit; font-size:11px; font-weight:600; cursor:pointer;
      padding:3px 8px; border-radius:4px;
      border:1px solid var(--border); background:var(--panel); color:var(--heading);
    }}
    .budget-calc-btn:hover, .budget-calc-btn:focus-visible {{
      border-color:var(--heading); outline:none;
    }}
    .budget-calc-tip {{
      display:none; position:absolute; right:0; top:calc(100% + 6px); z-index:30;
      min-width:280px; max-width:380px; padding:8px 10px;
      border:1px solid var(--border); border-radius:6px;
      background:var(--panel); color:var(--text);
      box-shadow:0 6px 18px rgba(0,0,0,.28);
      font-size:11px; line-height:1.4; white-space:normal;
    }}
    .budget-calc-wrap:hover .budget-calc-tip,
    .budget-calc-wrap:focus-within .budget-calc-tip {{ display:block; }}
    .budget-calc-tip .calc-line {{ margin:0 0 4px; }}
    .budget-calc-tip .calc-line:last-child {{ margin-bottom:0; }}
    .budget-calc-tip .calc-muted {{ color:var(--muted2); }}
    .budget-calc-tip .calc-total {{
      margin-top:6px; padding-top:6px; border-top:1px solid var(--border);
      font-weight:600;
    }}
    .card-plan-panel {{ overflow:visible; min-width:0; }}
    .card-plan-head {{
      display:flex; align-items:center; justify-content:space-between; gap:10px; flex-wrap:wrap;
    }}
    .card-plan-head h2 {{ margin:0; }}
    .card-plan-range {{
      display:flex; align-items:center; gap:8px; flex-wrap:wrap;
    }}
    .card-plan-range button {{
      background:var(--toggle-bg); color:var(--toggle-fg); border:1px solid var(--border);
      border-radius:4px; padding:3px 10px; font-size:12px; cursor:pointer;
    }}
    .card-plan-summary {{
      display:flex; flex-wrap:wrap; gap:8px; margin:0 0 10px;
    }}
    .card-plan-box {{
      flex:1 1 140px; min-width:120px; max-width:220px;
      padding:8px 10px; border:1px solid var(--border); border-radius:6px;
      background:var(--bg);
    }}
    .card-plan-box.total {{
      border-color:var(--heading); background:var(--panel);
      max-width:200px;
    }}
    .card-plan-box .box-top {{
      display:flex; align-items:center; gap:6px; margin-bottom:4px;
    }}
    .card-plan-box .dot {{
      width:8px; height:8px; border-radius:50%; flex-shrink:0;
    }}
    .card-plan-box .box-name {{
      font-size:11px; font-weight:600; color:var(--heading);
      text-transform:uppercase; letter-spacing:.02em;
    }}
    .card-plan-box .box-total {{
      font-size:16px; font-weight:700; font-variant-numeric:tabular-nums;
      line-height:1.2; margin:2px 0 4px;
    }}
    .card-plan-box .box-split {{
      font-size:10px; color:var(--muted2); line-height:1.35;
    }}
    .chart-card-plan {{ height:300px; min-width:0; overflow:hidden; }}
    .budget-categories {{
      flex:1; display:flex; flex-direction:column;
      border:1px solid var(--border); border-radius:6px;
      padding:8px 10px 6px; background:var(--panel);
    }}
    .budget-categories-title {{
      font-size:11px; font-weight:600; color:var(--heading);
      text-transform:uppercase; letter-spacing:.03em; margin-bottom:6px;
    }}
    .budget-categories .bar-row-total {{
      margin:0 0 8px; padding:0 0 8px; border-bottom:1px solid var(--border);
    }}
    .budget-categories .budget-body {{ flex:1; overflow:visible; }}
    .chart-span {{ grid-column:1 / -1; }}
    .chart-bal {{ grid-column:span 3; }}
    .chart-invest {{ grid-column:span 1; }}
    .note {{ background:var(--note-bg); color:var(--note-fg); padding:6px 10px; border-radius:4px; margin-bottom:10px; font-size:12px; }}
    .note.ok {{ background:var(--note-ok-bg); color:var(--note-ok-fg); }}
    .perf-line {{ color:var(--perf); font-size:12px; line-height:1.4; }}
    .bar-row {{ margin:6px 0 8px; }}
    .bar-row-total {{
      margin:4px 0 12px; padding:8px 0 10px; border-bottom:1px solid var(--border);
    }}
    .bar-row-total .bar-head {{ font-weight:700; font-size:13px; }}
    .bar-row-total .bar-track {{ height:10px; }}
    .bar-head {{ display:flex; justify-content:space-between; gap:8px; font-size:12px; margin-bottom:3px; }}
    .bar-track {{ height:7px; background:var(--track); border-radius:4px; overflow:hidden; }}
    .bar-fill {{ height:100%; border-radius:4px; }}
    .bar-meta {{ color:var(--muted2); font-size:11px; margin-top:2px; }}
    .empty {{ color:var(--empty); margin:0; }}
    .chart-clickable {{ cursor:pointer; }}
    #categoryView {{ display:none; }}
    #categoryView.active {{ display:block; }}
    #cockpitView.hidden {{ display:none; }}
    .cat-view-head {{
      display:flex; align-items:center; justify-content:space-between; gap:12px;
      flex-wrap:wrap; margin-bottom:12px;
    }}
    .cat-view-head h2 {{ margin:0; font-size:16px; font-weight:600; }}
    .cat-view-meta {{ color:var(--muted); font-size:12px; margin-top:4px; }}
    .cat-view-total {{ font-size:15px; font-weight:600; }}
    .cat-view-total.neg {{ color:#e74c3c; }}
    .cat-view-total.pos {{ color:#2ecc71; }}
    .tx-table-wrap {{
      background:var(--panel); border:1px solid var(--border); border-radius:6px;
      overflow:auto;
    }}
    table.tx-table {{ width:100%; border-collapse:collapse; font-size:12px; }}
    table.tx-table th, table.tx-table td {{
      padding:8px 10px; text-align:left; border-bottom:1px solid var(--border);
      white-space:nowrap;
    }}
    table.tx-table th {{ color:var(--muted); font-weight:600; font-size:11px;
      text-transform:uppercase; letter-spacing:.03em; position:sticky; top:0;
      background:var(--panel); }}
    table.tx-table td.desc {{ white-space:normal; max-width:320px; }}
    table.tx-table td.amt {{ text-align:right; font-variant-numeric:tabular-nums; }}
    table.tx-table tr:last-child td {{ border-bottom:none; }}
    @media (max-width:900px) {{
      .charts-top, .inc-rec-row, .charts-reports {{ grid-template-columns:1fr; }}
      .split {{ flex-direction:column; }}
      .split > .panel,
      .split > .perf-notif-stack {{ flex:1 1 auto; }}
      .exp-perf-cell {{ grid-template-columns:1fr; }}
      .chart-bal, .chart-invest, .chart-span {{ grid-column:auto; }}
    }}
    @media (min-width:901px) and (max-width:1099px) {{
      .exp-perf-cell {{ grid-template-columns:1fr; }}
      .charts-reports {{ grid-template-columns:1fr 1fr; }}
      .chart-bal, .chart-invest, .chart-span {{ grid-column:1 / -1; }}
    }}
  </style>
</head>
<body>
<main>
  <div id="cockpitView">
  {cards_html}
  {perf_html}
  <div class="{charts_class}">
    {charts_body}
  </div>
  {split_html}
  </div>
  <div id="categoryView">
    <div class="panel">
      <div class="cat-view-head">
        <div>
          <h2 id="catViewTitle">Category</h2>
          <div class="cat-view-meta" id="catViewMeta"></div>
        </div>
        <div style="text-align:right">
          <div class="cat-view-total" id="catViewTotal"></div>
          <button type="button" id="catViewBack">Back to dashboard</button>
        </div>
      </div>
      <div class="tx-table-wrap">
        <table class="tx-table">
          <thead>
            <tr>
              <th>Date</th>
              <th>Description</th>
              <th>Amount</th>
              <th>Account</th>
              <th>Type</th>
            </tr>
          </thead>
          <tbody id="catViewBody"></tbody>
        </table>
      </div>
      <p class="empty" id="catViewEmpty" style="display:none;padding:12px 0">No transactions for this category in the selected period</p>
    </div>
  </div>
</main>
<footer class="cockpit-footer">
  <span class="period-label" id="periodTitle">{data['year_month']}</span>
  <div class="period-controls">
    <label>From <input type="date" id="periodFrom" value="{date_from}"/></label>
    <label>To <input type="date" id="periodTo" value="{date_to}"/></label>
    <button type="button" id="periodApply">Apply</button>
    <button type="button" id="themeToggle" aria-label="Toggle theme">Light</button>
  </div>
</footer>
<script>
const DATA = {payload_json};
const RAW = {raw_json};
const THEME_KEY = 'finance-cockpit-theme';
const MONTH_NAMES = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
let activeCategory = null; // {{ id, kind }} when detail view is open

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
function shiftDate(ymd, deltaDays) {{
  const d = new Date(ymd + 'T12:00:00');
  d.setDate(d.getDate() + deltaDays);
  return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
}}
function daysInclusive(from, to) {{
  const a = new Date(from + 'T12:00:00');
  const b = new Date(to + 'T12:00:00');
  return Math.round((b - a) / 86400000) + 1;
}}
function monthsSpanning(from, to) {{
  let a = from.slice(0,7), b = to.slice(0,7);
  if (a > b) {{ const t = a; a = b; b = t; }}
  const out = [];
  let cur = a;
  while (cur <= b) {{
    out.push(cur);
    cur = monthShift(cur, 1);
  }}
  return out;
}}
function periodLabel(from, to) {{
  if (!from || !to) return RAW.monthStart + ' – ' + RAW.today;
  return from === to ? from : (from + ' – ' + to);
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
function dateRangeTotals(from, to) {{
  let income = 0, expense = 0;
  for (const t of RAW.transactions) {{
    const d = String(t.date || '').slice(0,10);
    if (!d || d < from || d > to) continue;
    const amt = parseDecimal(t.amount);
    if (t.type === 'income') income += amt;
    else if (t.type === 'expense' || t.type === 'card_expense') expense += amt;
  }}
  return {{income, expense, balance: income - expense}};
}}
function byCategoryDates(from, to, types) {{
  const byId = catById();
  const totals = {{}};
  for (const t of RAW.transactions) {{
    const d = String(t.date || '').slice(0,10);
    if (!d || d < from || d > to) continue;
    if (!types.has(t.type)) continue;
    const cid = mainCategoryId(t.category_id || '', byId);
    totals[cid] = (totals[cid] || 0) + parseDecimal(t.amount);
  }}
  const rows = Object.entries(totals).map(([cid, amt]) => {{
    const row = byId[cid];
    return {{id: cid, name: catLabel(row, cid), value: amt, color: (row && row.color) || '#7F8C8D'}};
  }});
  rows.sort((a,b) => b.value - a.value);
  return rows;
}}
function incomeVsInvest(from, to) {{
  const byId = catById();
  let income = 0, invest = 0;
  for (const t of RAW.transactions) {{
    const d = String(t.date || '').slice(0,10);
    if (!d || d < from || d > to) continue;
    if (t.type !== 'income') continue;
    const amt = parseDecimal(t.amount);
    const cid = mainCategoryId(t.category_id || '', byId);
    if (cid === 'CAT_INVESTIM') invest += amt;
    else income += amt;
  }}
  return {{income, invest}};
}}
function pieFromRows(rows) {{
  return {{
    labels: rows.map(r => r.name),
    values: rows.map(r => Math.round(r.value * 100) / 100),
    colors: rows.map(r => r.color),
    custom: rows.map(r => formatBrl(r.value)),
    categoryIds: rows.map(r => r.id)
  }};
}}
function accountName(accountId) {{
  if (!accountId) return '';
  for (const a of RAW.accounts || []) {{
    if (a.id === accountId) return a.name || a.id;
  }}
  return accountId;
}}
function cardName(cardId) {{
  if (!cardId) return '';
  for (const c of RAW.cards || []) {{
    if (c.id === cardId) return c.name || c.id;
  }}
  return cardId;
}}
function typeLabel(t, cardId) {{
  if (t === 'transfer' && cardId) return 'Card payment';
  if (t === 'card_expense') return 'Credit card';
  if (t === 'expense') return 'Expense';
  if (t === 'income') return 'Income';
  if (t === 'transfer') return 'Transfer';
  return t || '';
}}
function currentPeriod() {{
  const fromEl = document.getElementById('periodFrom');
  const toEl = document.getElementById('periodTo');
  let from = fromEl ? fromEl.value : '';
  let to = toEl ? toEl.value : '';
  if (from && to && from > to) {{ const t = from; from = to; to = t; }}
  return {{ from, to }};
}}
function openCategoryView(categoryId, kind) {{
  if (!categoryId) return;
  activeCategory = {{ id: categoryId, kind: kind || 'expense' }};
  const cockpit = document.getElementById('cockpitView');
  const detail = document.getElementById('categoryView');
  if (cockpit) cockpit.classList.add('hidden');
  if (detail) detail.classList.add('active');
  renderCategoryView();
}}
function closeCategoryView() {{
  activeCategory = null;
  const cockpit = document.getElementById('cockpitView');
  const detail = document.getElementById('categoryView');
  if (detail) detail.classList.remove('active');
  if (cockpit) cockpit.classList.remove('hidden');
}}
function renderCategoryView() {{
  if (!activeCategory) return;
  const {{ from, to }} = currentPeriod();
  const byId = catById();
  const label = catLabel(byId[activeCategory.id], activeCategory.id);
  const types = activeCategory.kind === 'income'
    ? new Set(['income'])
    : new Set(['expense', 'card_expense']);
  const rows = [];
  let total = 0;
  for (const t of RAW.transactions) {{
    const d = String(t.date || '').slice(0,10);
    if (!d || (from && d < from) || (to && d > to)) continue;
    if (!types.has(t.type)) continue;
    if (mainCategoryId(t.category_id || '', byId) !== activeCategory.id) continue;
    const amt = parseDecimal(t.amount);
    total += amt;
    rows.push(t);
  }}
  rows.sort((a, b) => String(b.date || '').localeCompare(String(a.date || '')));
  const title = document.getElementById('catViewTitle');
  const meta = document.getElementById('catViewMeta');
  const totEl = document.getElementById('catViewTotal');
  const body = document.getElementById('catViewBody');
  const empty = document.getElementById('catViewEmpty');
  const wrap = document.querySelector('.tx-table-wrap');
  if (title) title.textContent = label;
  if (meta) {{
    const kindLbl = activeCategory.kind === 'income' ? 'Incomes' : 'Expenses';
    meta.textContent = kindLbl + ' · ' + periodLabel(from, to) + ' · ' + rows.length
      + (rows.length === 1 ? ' transaction' : ' transactions');
  }}
  if (totEl) {{
    totEl.textContent = formatBrl(total);
    totEl.className = 'cat-view-total ' + (activeCategory.kind === 'income' ? 'pos' : 'neg');
  }}
  if (!body) return;
  if (!rows.length) {{
    body.innerHTML = '';
    if (wrap) wrap.style.display = 'none';
    if (empty) empty.style.display = '';
    return;
  }}
  if (wrap) wrap.style.display = '';
  if (empty) empty.style.display = 'none';
  body.innerHTML = rows.map(t => {{
    const acc = t.type === 'card_expense'
      ? (cardName(t.card_id) || accountName(t.account_id) || '—')
      : (accountName(t.account_id) || '—');
    return '<tr>'
      + '<td>' + String(t.date || '').slice(0,10) + '</td>'
      + '<td class="desc">' + escapeHtml(t.description || '') + '</td>'
      + '<td class="amt">' + formatBrl(parseDecimal(t.amount)) + '</td>'
      + '<td>' + escapeHtml(acc) + '</td>'
      + '<td>' + escapeHtml(typeLabel(t.type, t.card_id)) + '</td>'
      + '</tr>';
  }}).join('');
}}
function escapeHtml(s) {{
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}}
function bindChartClick(elId, kind) {{
  const el = document.getElementById(elId);
  if (!el || !el.on) return;
  el.on('plotly_click', (ev) => {{
    if (!ev || !ev.points || !ev.points.length) return;
    const pt = ev.points[0];
    let cid = '';
    if (Array.isArray(pt.customdata)) cid = pt.customdata[1] || '';
    else if (pt.customdata && typeof pt.customdata === 'object') cid = pt.customdata.categoryId || '';
    if (!cid && typeof pt.pointNumber === 'number') {{
      const spec = kind === 'income' ? DATA.incomePie : DATA.expensePie;
      if (spec && spec.categoryIds)
        cid = spec.categoryIds[pt.pointNumber] || '';
    }}
    if (cid) openCategoryView(cid, kind);
  }});
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
function seriesFor(months) {{
  return months.map(ym => {{
    const tot = monthTotals([ym]);
    return {{month: ym, income: tot.income, expense: tot.expense, balance: tot.balance}};
  }});
}}
function daysSpanning(from, to) {{
  let a = from, b = to;
  if (a > b) {{ const t = a; a = b; b = t; }}
  const out = [];
  let cur = a;
  while (cur <= b) {{
    out.push(cur);
    cur = shiftDate(cur, 1);
  }}
  return out;
}}
function dayTotals(ymd) {{
  let income = 0, expense = 0;
  for (const t of RAW.transactions) {{
    const d = String(t.date || '').slice(0,10);
    if (d !== ymd) continue;
    const amt = parseDecimal(t.amount);
    if (t.type === 'income') income += amt;
    else if (t.type === 'expense' || t.type === 'card_expense') expense += amt;
  }}
  return {{income, expense, balance: income - expense}};
}}
function dailySeries(from, to) {{
  return daysSpanning(from, to).map(day => {{
    const tot = dayTotals(day);
    return {{day: day, income: tot.income, expense: tot.expense, balance: tot.balance}};
  }});
}}
function dayTickLabel(ymd) {{
  if (!ymd || ymd.length < 10) return ymd || '';
  return ymd.slice(8, 10) + '/' + ymd.slice(5, 7);
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
function formatCsvDecimal(num) {{
  return (Math.round(Number(num) * 100) / 100).toFixed(2).replace('.', ',');
}}
function budgetsForMonth(ym) {{
  if (!ym) return [];
  return RAW.budgets
    .filter(b => b.year_month === ym)
    .map(b => ({{
      category_id: b.category_id || '',
      planned: parseDecimal(b.planned_amount),
      spent: parseDecimal(b.spent_amount)
    }}))
    .sort((a, b) => a.category_id.localeCompare(b.category_id));
}}
function monthsWithBudgets(months) {{
  const have = new Set((RAW.budgets || []).map(b => b.year_month));
  const inPeriod = (months || []).filter(m => have.has(m));
  if (inPeriod.length) return inPeriod;
  if (RAW.currentMonth && have.has(RAW.currentMonth)) return [RAW.currentMonth];
  return [...have].sort();
}}
let budgetSaveTimer = null;
let budgetSaveClearTimer = null;
let budgetApiOk = null;
function setBudgetSaveStatus(text, cls) {{
  const el = document.getElementById('budgetSaveStatus');
  if (!el) return;
  el.textContent = text || '';
  el.className = 'budget-save-status' + (cls ? ' ' + cls : '');
  if (budgetSaveClearTimer) clearTimeout(budgetSaveClearTimer);
  if (cls === 'saved') {{
    budgetSaveClearTimer = setTimeout(() => {{
      if (el.textContent === text) {{
        el.textContent = '';
        el.className = 'budget-save-status';
      }}
    }}, 2000);
  }}
}}
async function checkBudgetApi() {{
  try {{
    const r = await fetch('/api/health', {{ cache: 'no-store' }});
    budgetApiOk = r.ok;
  }} catch (e) {{
    budgetApiOk = false;
  }}
  return budgetApiOk;
}}
async function refreshCardsFromApi() {{
  try {{
    const r = await fetch('/api/cards', {{ cache: 'no-store' }});
    if (!r.ok) return false;
    const data = await r.json();
    if (!data || !Array.isArray(data.cards)) return false;
    RAW.cards = data.cards;
    const byId = {{}};
    for (const c of data.cards) byId[c.id] = c;
    for (const lc of (RAW.liquidCards || [])) {{
      const src = byId[lc.id];
      if (!src) continue;
      const spent = parseDecimal(src.current_spent);
      if (Number.isFinite(spent)) lc.spent = spent;
      const lim = parseDecimal(src.limit);
      if (Number.isFinite(lim)) lc.limit = lim;
    }}
    return true;
  }} catch (e) {{
    return false;
  }}
}}
function syncBudgetMonthSelect(months) {{
  const sel = document.getElementById('budgetMonth');
  if (!sel) return '';
  const opts = monthsWithBudgets(months);
  const prev = sel.value;
  sel.innerHTML = opts.map(m => '<option value="' + m + '">' + m + '</option>').join('');
  let pick = prev;
  if (!opts.includes(pick)) {{
    pick = opts.length ? opts[opts.length - 1] : '';
  }}
  sel.value = pick;
  sel.disabled = !opts.length || budgetApiOk === false;
  return pick;
}}
function updateRawBudgetPlanned(ym, cid, plannedFmt) {{
  for (const b of RAW.budgets) {{
    if (b.year_month === ym && b.category_id === cid) {{
      b.planned_amount = plannedFmt;
      return true;
    }}
  }}
  return false;
}}
const lastSavedBudgetPlanned = Object.create(null);
function budgetSaveKey(ym, cid) {{
  return ym + '|' + cid;
}}
function applyLiveBudgetPlanned(ym, cid, inputEl) {{
  if (!ym || !cid || !inputEl) return;
  const raw = String(inputEl.value || '').trim();
  if (raw === '') return;
  const planned = parseDecimal(raw);
  if (!(planned >= 0) || Number.isNaN(planned)) return;
  updateRawBudgetPlanned(ym, cid, formatCsvDecimal(planned));
  refreshBudgetVisuals(ym);
}}
function scheduleBudgetSave(ym, cid, inputEl) {{
  if (budgetSaveTimer) clearTimeout(budgetSaveTimer);
  budgetSaveTimer = setTimeout(() => saveBudgetPlanned(ym, cid, inputEl), 400);
}}
function budgetRowStats(planned, spent) {{
  const rem = planned - spent;
  const pct = planned > 0 ? (spent / planned * 100) : (spent > 0 ? 100 : 0);
  const width = Math.min(pct, 100);
  const over = spent > planned && planned > 0;
  const fill = over ? '#e74c3c' : '#2ecc71';
  const cap = over ? ('Exceeded ' + formatBrl(-rem)) : ('Remain ' + formatBrl(rem));
  return {{ rem, pct, width, over, fill, cap }};
}}
function plannedTotalForCurrentMonth() {{
  const ym = RAW.currentMonth || '';
  let total = 0;
  for (const b of (RAW.budgets || [])) {{
    if (b.year_month !== ym) continue;
    total += parseDecimal(b.planned_amount);
  }}
  return total;
}}
function refreshFundsCompare() {{
  const root = document.getElementById('fundsCompare');
  if (!root) return;
  const available = Number(RAW.liquidAfterCard);
  const avail = Number.isFinite(available) ? available : 0;
  const planned = plannedTotalForCurrentMonth();
  const scale = Math.max(avail, planned, 0.01);
  const availPct = Math.max(0, (avail / scale) * 100);
  const planPct = Math.max(0, (planned / scale) * 100);
  const deficit = planned > avail + 0.001;
  const headroom = avail - planned;
  root.classList.toggle('deficit', deficit);
  root.classList.toggle('ok', !deficit);
  const availEl = document.getElementById('fundsAvailableVal');
  const planEl = document.getElementById('fundsPlannedVal');
  const barAvail = document.getElementById('fundsBarAvailable');
  const barPlan = document.getElementById('fundsBarPlanned');
  const markAvail = document.getElementById('fundsMarkAvailable');
  const markPlan = document.getElementById('fundsMarkPlanned');
  const meta = document.getElementById('fundsCompareMeta');
  const tip = document.getElementById('budgetCalcTip');
  if (availEl) availEl.textContent = formatBrl(avail);
  if (planEl) planEl.textContent = formatBrl(planned);
  if (barAvail) barAvail.style.width = availPct.toFixed(1) + '%';
  if (barPlan) barPlan.style.width = planPct.toFixed(1) + '%';
  if (markAvail) markAvail.style.left = availPct.toFixed(1) + '%';
  if (markPlan) markPlan.style.left = planPct.toFixed(1) + '%';
  if (meta) {{
    meta.textContent = deficit
      ? ('Over by ' + formatBrl(-headroom) + ' · reduce planned spending')
      : ('Headroom ' + formatBrl(headroom));
  }}
  if (tip) {{
    const accRows = Array.isArray(RAW.liquidAccounts) ? RAW.liquidAccounts : [];
    const cardRows = Array.isArray(RAW.liquidCards) ? RAW.liquidCards : [];
    const accBal = Number(RAW.liquidAccountBal);
    const cardSpent = Number(RAW.liquidCardSpent);
    const cardLimit = Number(RAW.liquidCardLimit);
    const bal = Number.isFinite(accBal) ? accBal : 0;
    const spent = Number.isFinite(cardSpent) ? cardSpent : 0;
    const limit = Number.isFinite(cardLimit) ? cardLimit : 0;
    const parts = [];
    if (accRows.length) {{
      parts.push('<div class="calc-line calc-muted">Checking</div>');
      for (const a of accRows) {{
        parts.push('<div class="calc-line">' + (a.name || 'Account')
          + ' · ' + formatBrl(Number(a.balance) || 0) + '</div>');
      }}
    }} else {{
      const accName = RAW.liquidAccountName || 'Checking';
      parts.push('<div class="calc-line">' + accName + ' · ' + formatBrl(bal) + '</div>');
    }}
    if (cardRows.length) {{
      parts.push('<div class="calc-line calc-muted">Credit cards</div>');
      for (const c of cardRows) {{
        const cSpent = Number(c.spent) || 0;
        const cLimit = Number(c.limit) || 0;
        const cAvail = cLimit - cSpent;
        parts.push('<div class="calc-line">' + (c.name || 'Card')
          + ' · spent ' + formatBrl(cSpent)
          + ' · limit ' + formatBrl(cLimit)
          + ' · avail ' + formatBrl(cAvail) + '</div>');
      }}
    }} else {{
      parts.push('<div class="calc-line">Cards · spent ' + formatBrl(spent)
        + ' · limit ' + formatBrl(limit) + '</div>');
    }}
    parts.push('<div class="calc-line calc-total">Available = checking '
      + formatBrl(bal) + ' − card spent ' + formatBrl(spent)
      + ' = ' + formatBrl(avail) + '</div>');
    tip.innerHTML = parts.join('');
  }}
}}
function refreshBudgetVisuals(ym) {{
  const rows = budgetsForMonth(ym);
  const sumEl = document.getElementById('budgetsSummary');
  let totalPlanned = 0, totalSpent = 0;
  for (const b of rows) {{
    totalPlanned += b.planned;
    totalSpent += b.spent;
  }}
  const totRem = totalPlanned - totalSpent;
  const totPct = totalPlanned > 0 ? (totalSpent / totalPlanned * 100) : (totalSpent > 0 ? 100 : 0);
  const totWidth = Math.min(totPct, 100);
  const totOver = totalSpent > totalPlanned && totalPlanned > 0;
  const totFill = totOver ? '#e74c3c' : '#f1c40f';
  const budHint = totOver ? ('Exceeded ' + formatBrl(-totRem)) : ('Remain ' + formatBrl(totRem));
  if (sumEl) {{
    sumEl.style.display = rows.length ? '' : 'none';
    if (rows.length) {{
      sumEl.innerHTML = '<div class="bar-head"><span>Total</span><span>'
        + formatBrl(totalSpent) + ' / ' + formatBrl(totalPlanned) + ' · ' + totPct.toFixed(0) + '%</span></div>'
        + '<div class="bar-track"><div class="bar-fill" style="width:' + totWidth.toFixed(1)
        + '%;background:' + totFill + '"></div></div>'
        + '<div class="bar-meta">' + budHint + '</div>';
    }} else {{
      sumEl.innerHTML = '';
    }}
  }}
  for (const b of rows) {{
    const row = document.querySelector('#budgetsBody .bar-row[data-category-id="' + b.category_id + '"]');
    if (!row) continue;
    const st = budgetRowStats(b.planned, b.spent);
    const pctEl = row.querySelector('.budget-pct');
    if (pctEl) pctEl.textContent = '· ' + st.pct.toFixed(0) + '%';
    const fillEl = row.querySelector('.bar-fill');
    if (fillEl) {{
      fillEl.style.width = st.width.toFixed(1) + '%';
      fillEl.style.background = st.fill;
    }}
    const meta = row.querySelector('.bar-meta');
    if (meta) meta.textContent = st.cap;
  }}
  refreshFundsCompare();
}}
async function saveBudgetPlanned(ym, cid, inputEl) {{
  if (!ym || !cid || !inputEl) return;
  if (budgetApiOk === false) {{
    setBudgetSaveStatus('Save unavailable (open via Finance dashboard)', 'error');
    refreshFundsCompare();
    return;
  }}
  const raw = String(inputEl.value || '').trim();
  if (raw === '') return;
  const planned = parseDecimal(raw);
  if (!(planned >= 0) || Number.isNaN(planned)) {{
    setBudgetSaveStatus('Invalid amount', 'error');
    return;
  }}
  const exists = (RAW.budgets || []).some(b => b.year_month === ym && b.category_id === cid);
  if (!exists) {{
    setBudgetSaveStatus('No budget row for this category', 'error');
    return;
  }}
  updateRawBudgetPlanned(ym, cid, formatCsvDecimal(planned));
  refreshFundsCompare();
  const key = budgetSaveKey(ym, cid);
  if (lastSavedBudgetPlanned[key] != null
      && Math.abs(lastSavedBudgetPlanned[key] - planned) < 0.001) {{
    return;
  }}
  setBudgetSaveStatus('Saving…', 'saving');
  try {{
    const r = await fetch('/api/budgets', {{
      method: 'PATCH',
      headers: {{ 'Content-Type': 'application/json' }},
      body: JSON.stringify({{ year_month: ym, category_id: cid, planned_amount: planned }})
    }});
    const data = await r.json().catch(() => ({{}}));
    if (!r.ok || !data.ok) {{
      setBudgetSaveStatus(data.error || ('Save failed (' + r.status + ')'), 'error');
      refreshFundsCompare();
      return;
    }}
    const fmt = data.planned_amount || formatCsvDecimal(planned);
    updateRawBudgetPlanned(ym, cid, fmt);
    lastSavedBudgetPlanned[key] = parseDecimal(fmt);
    if (inputEl.isConnected && document.activeElement !== inputEl) {{
      inputEl.value = parseDecimal(fmt).toFixed(2);
    }}
    setBudgetSaveStatus('Saved', 'saved');
    const sel = document.getElementById('budgetMonth');
    if (sel && sel.value === ym) refreshBudgetVisuals(ym);
    else refreshFundsCompare();
  }} catch (e) {{
    budgetApiOk = false;
    setBudgetSaveStatus('Save failed (server offline)', 'error');
    refreshFundsCompare();
  }}
}}
function onBudgetPlannedBlur(ym, inputEl) {{
  if (!inputEl) return;
  const planned = parseDecimal(inputEl.value);
  if (!(planned >= 0) || Number.isNaN(planned)) return;
  inputEl.value = planned.toFixed(2);
  applyLiveBudgetPlanned(ym, inputEl.dataset.categoryId, inputEl);
  scheduleBudgetSave(ym, inputEl.dataset.categoryId, inputEl);
}}
function renderBudgets(rows) {{
  const el = document.getElementById('budgetsBody');
  const sumEl = document.getElementById('budgetsSummary');
  const sel = document.getElementById('budgetMonth');
  const ym = sel ? sel.value : '';
  const canEdit = !!ym && budgetApiOk !== false;
  if (!el) return;
  if (!rows.length) {{
    if (sumEl) {{
      sumEl.innerHTML = '';
      sumEl.style.display = 'none';
    }}
    el.innerHTML = '<p class="empty">No budgets this month</p>';
    refreshFundsCompare();
    return;
  }}
  let totalPlanned = 0, totalSpent = 0;
  for (const b of rows) {{
    totalPlanned += b.planned;
    totalSpent += b.spent;
  }}
  const totRem = totalPlanned - totalSpent;
  const totPct = totalPlanned > 0 ? (totalSpent / totalPlanned * 100) : (totalSpent > 0 ? 100 : 0);
  const totWidth = Math.min(totPct, 100);
  const totOver = totalSpent > totalPlanned && totalPlanned > 0;
  const totFill = totOver ? '#e74c3c' : '#f1c40f';
  const budHint = totOver ? ('Exceeded ' + formatBrl(-totRem)) : ('Remain ' + formatBrl(totRem));
  if (sumEl) {{
    sumEl.style.display = '';
    sumEl.innerHTML = '<div class="bar-head"><span>Total</span><span>'
      + formatBrl(totalSpent) + ' / ' + formatBrl(totalPlanned) + ' · ' + totPct.toFixed(0) + '%</span></div>'
      + '<div class="bar-track"><div class="bar-fill" style="width:' + totWidth.toFixed(1)
      + '%;background:' + totFill + '"></div></div>'
      + '<div class="bar-meta">' + budHint + '</div>';
  }}
  const byId = catById();
  el.innerHTML = rows.map(b => {{
    const name = catLabel(byId[b.category_id], b.category_id);
    const st = budgetRowStats(b.planned, b.spent);
    const plannedVal = b.planned.toFixed(2);
    return '<div class="bar-row" data-category-id="' + b.category_id + '">'
      + '<div class="bar-head"><span>' + name + '</span></div>'
      + '<div class="budget-edit-row"><span class="spent-label">' + formatBrl(b.spent)
      + ' /</span><input type="text" inputmode="decimal" class="budget-planned-input" '
      + 'data-category-id="' + b.category_id + '" value="' + plannedVal + '"'
      + (canEdit ? '' : ' disabled') + '/>'
      + '<span class="budget-pct">· ' + st.pct.toFixed(0) + '%</span></div>'
      + '<div class="bar-track"><div class="bar-fill" style="width:' + st.width.toFixed(1) + '%;background:' + st.fill + '"></div></div>'
      + '<div class="bar-meta">' + st.cap + '</div></div>';
  }}).join('');
  el.querySelectorAll('.budget-planned-input').forEach(inp => {{
    inp.addEventListener('input', () => {{
      applyLiveBudgetPlanned(ym, inp.dataset.categoryId, inp);
      scheduleBudgetSave(ym, inp.dataset.categoryId, inp);
    }});
    inp.addEventListener('change', () => onBudgetPlannedBlur(ym, inp));
  }});
  refreshFundsCompare();
}}
function recurringMonthlyTotal() {{
  let s = 0;
  for (const r of RAW.recurring || []) s += parseDecimal(r.monthly_amount);
  return s;
}}
function renderRecurring(months, income) {{
  const sumEl = document.getElementById('recurringSummary');
  const rows = RAW.recurring || [];
  const n = (months && months.length) ? months.length : 1;
  const periodBills = recurringMonthlyTotal() * n;
  const pct = income > 0 ? (periodBills / income * 100) : 0;
  if (sumEl) {{
    sumEl.textContent = 'Period ' + formatBrl(periodBills) + ' · Income ' + formatBrl(income)
      + ' · Recurring is ' + pct.toFixed(0) + '% of income'
      + ' · ' + n + (n === 1 ? ' month' : ' months');
  }}
  const mapped = rows.map(r => ({{
    name: r.name || r.id || 'Bill',
    icon: r.icon || '🔁',
    periodAmt: parseDecimal(r.monthly_amount) * n
  }}));
  DATA.recurringVs = {{ bills: periodBills, income: income, rows: mapped }};
}}
function drawIncomeInvestTreemap() {{
  const el = document.getElementById('incomeVsInvest');
  if (!el) return;
  const vs = DATA.incomeVsInvest || {{ income: 0, invest: 0 }};
  const labels = [];
  const values = [];
  const colors = [];
  const custom = [];
  const total = (vs.income || 0) + (vs.invest || 0);
  const tiles = [
    {{ label: 'Income', value: vs.income || 0, color: '#2ecc71' }},
    {{ label: 'Investments', value: vs.invest || 0, color: '#e74c3c' }}
  ];
  for (const t of tiles) {{
    if (t.value <= 0) continue;
    labels.push(t.label);
    values.push(t.value);
    colors.push(t.color);
    const pct = total > 0 ? (t.value / total * 100) : 0;
    custom.push(formatBrl(t.value) + ' · ' + pct.toFixed(0) + '% of total');
  }}
  if (!values.length) {{
    el.innerHTML = '<p class="empty">No data</p>';
    return;
  }}
  const L = baseLayout();
  Plotly.newPlot('incomeVsInvest', [{{
    type: 'treemap',
    labels: labels,
    parents: labels.map(() => ''),
    values: values,
    marker: {{ colors: colors }},
    customdata: custom,
    texttemplate: '%{{label}}<br>%{{customdata}}',
    hovertemplate: '%{{label}}<br>%{{customdata}}<extra></extra>',
    textfont: {{ size: 12 }},
    pathbar: {{ visible: false }}
  }}], Object.assign({{}}, L, {{
    showlegend: false,
    height: 280,
    margin: {{ t: 8, b: 8, l: 8, r: 8 }}
  }}), {{ responsive: true, displayModeBar: false }});
}}
function drawRecurringTreemap() {{
  const el = document.getElementById('recurringVsIncome');
  if (!el) return;
  const vs = DATA.recurringVs || {{ bills: 0, income: 0, rows: [] }};
  const labels = [];
  const values = [];
  const colors = [];
  const custom = [];
  for (const r of vs.rows || []) {{
    if (r.periodAmt <= 0) continue;
    const icon = (r.icon || '').trim();
    const title = r.name || 'Bill';
    labels.push(icon ? (icon + ' ' + title) : title);
    values.push(r.periodAmt);
    colors.push('#9b59b6');
    const pInc = vs.income > 0 ? (r.periodAmt / vs.income * 100) : 0;
    custom.push(formatBrl(r.periodAmt) + ' · ' + pInc.toFixed(0) + '% of income');
  }}
  const leftover = vs.income - vs.bills;
  if (leftover > 0) {{
    labels.push('Remainder');
    values.push(leftover);
    colors.push('#2ecc71');
    const pInc = vs.income > 0 ? (leftover / vs.income * 100) : 0;
    custom.push(formatBrl(leftover) + ' · ' + pInc.toFixed(0) + '% of income');
  }}
  if (!values.length) {{
    el.innerHTML = '<p class="empty">No data</p>';
    return;
  }}
  const L = baseLayout();
  const panel = el.closest('.panel') || el.parentElement;
  const head = panel ? panel.querySelector('h2, .panel-head-row') : null;
  const summary = panel ? panel.querySelector('#recurringSummary') : null;
  const headH = head ? Math.ceil(head.getBoundingClientRect().height + 8) : 28;
  const summaryH = summary ? Math.ceil(summary.getBoundingClientRect().height + 8) : 0;
  let padY = 0;
  if (panel) {{
    const cs = getComputedStyle(panel);
    padY = (parseFloat(cs.paddingTop) || 0) + (parseFloat(cs.paddingBottom) || 0);
  }}
  const panelH = panel ? Math.floor(panel.getBoundingClientRect().height) : 0;
  const slack = 12;
  const chrome = headH + summaryH + padY + slack;
  const availH = panelH > chrome + 80 ? Math.floor(panelH - chrome) : 0;
  const h = Math.max(120, availH > 0 ? availH : Math.min(el.clientHeight || 220, 220));
  el.style.height = h + 'px';
  el.style.maxHeight = h + 'px';
  el.style.minHeight = '0';
  el.style.overflow = 'hidden';
  el.style.width = '100%';
  const plotW = Math.max(160, Math.floor(el.clientWidth || (panel && panel.clientWidth) || 280));
  Plotly.newPlot('recurringVsIncome', [{{
    type: 'treemap',
    labels: labels,
    parents: labels.map(() => ''),
    values: values,
    marker: {{ colors: colors }},
    customdata: custom,
    texttemplate: '%{{label}}<br>%{{customdata}}',
    hovertemplate: '%{{label}}<br>%{{customdata}}<extra></extra>',
    textfont: {{ size: 12 }},
    pathbar: {{ visible: false }}
  }}], Object.assign({{}}, L, {{
    showlegend: false,
    autosize: false,
    height: h,
    width: plotW,
    margin: {{ t: 4, b: 4, l: 4, r: 4 }}
  }}), {{ responsive: false, displayModeBar: false }}).then(() => {{
    el.style.height = h + 'px';
    el.style.maxHeight = h + 'px';
    const used = Math.ceil(el.getBoundingClientRect().height);
    if (used > h + 1) {{
      return Plotly.relayout(el, {{ height: h, width: plotW }});
    }}
  }});
}}
function applyPeriod() {{
  const fromEl = document.getElementById('periodFrom');
  const toEl = document.getElementById('periodTo');
  let from = fromEl.value;
  let to = toEl.value;
  if (!from || !to) return;
  if (from > to) {{
    const t = from; from = to; to = t;
    fromEl.value = from; toEl.value = to;
  }}
  const months = monthsSpanning(from, to);
  const year = to.slice(0,4);
  const tot = dateRangeTotals(from, to);
  const n = daysInclusive(from, to);
  const priorTo = shiftDate(from, -1);
  const priorFrom = shiftDate(from, -n);
  const prev = dateRangeTotals(priorFrom, priorTo);
  const expRows = byCategoryDates(from, to, new Set(['expense','card_expense']));
  const incRows = byCategoryDates(from, to, new Set(['income']));
  const multi = from !== to;
  DATA.expensePie = pieFromRows(expRows);
  DATA.incomePie = pieFromRows(incRows);
  DATA.series = dailySeries(from, to);
  DATA.annual = annualFor(year);
  DATA.incomeVsInvest = incomeVsInvest(from, to);

  const title = document.getElementById('periodTitle');
  if (title) title.textContent = periodLabel(from, to);
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
  renderBudgets(budgetsForMonth(syncBudgetMonthSelect(months)));
  renderRecurring(months, tot.income);
  if (activeCategory) renderCategoryView();
  else drawAll();
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
function pie(id, spec, kind) {{
  const host = document.getElementById(id);
  if (!host) return;
  if (!spec.values.length) {{ host.innerHTML = '<p class="empty">No data</p>'; return; }}
  const L = baseLayout();
  const custom = (spec.categoryIds || []).map((cid, i) => [spec.custom[i], cid]);
  const n = spec.labels.length;
  const panel = host.closest('.pie-exp-panel, .pie-inc-cell, .panel') || host.parentElement;
  const head = panel ? panel.querySelector('.panel-head-row') : null;
  const headH = head ? Math.ceil(head.getBoundingClientRect().height + 8) : 36;
  const panelH = panel ? Math.floor(panel.getBoundingClientRect().height) : 0;
  let padY = 0;
  if (panel) {{
    const cs = getComputedStyle(panel);
    padY = (parseFloat(cs.paddingTop) || 0) + (parseFloat(cs.paddingBottom) || 0);
  }}
  // Slack keeps the legend fully inside the panel (overflow:hidden).
  const slack = 16;
  const availH = panelH > headH + padY + 120
    ? Math.floor(panelH - headH - padY - slack)
    : 0;
  const legendDriven = Math.max(240, Math.min(460, 48 + n * 18));
  const height = availH > 0 ? availH : legendDriven;

  let plotEl = host.querySelector('.pie-plot');
  let legendEl = host.querySelector('.pie-legend');
  if (!plotEl) {{
    host.innerHTML = '';
    plotEl = document.createElement('div');
    plotEl.className = 'pie-plot';
    plotEl.id = id + 'Plot';
    legendEl = document.createElement('div');
    legendEl.className = 'pie-legend';
    host.appendChild(plotEl);
    host.appendChild(legendEl);
  }}

  host.style.height = height + 'px';
  host.style.maxHeight = height + 'px';
  host.style.minHeight = '0';
  host.style.width = '100%';
  host.style.overflow = 'hidden';
  host.style.display = '';
  host.classList.add('chart-clickable');

  legendEl.innerHTML = spec.labels.map((lab, i) =>
    '<span class="pie-legend-item"><i style="background:' + (spec.colors[i] || '#888') + '"></i>' +
    escapeHtml(lab) + '</span>'
  ).join('');

  // Collapse plot while measuring legend so height is not underestimated.
  plotEl.style.height = '0px';
  plotEl.style.maxHeight = '0px';
  plotEl.style.minHeight = '0';
  plotEl.style.overflow = 'hidden';
  const legendH = Math.max(
    36,
    Math.ceil(legendEl.getBoundingClientRect().height) || (16 + Math.ceil(n / 3) * 18)
  );
  const plotH = Math.max(100, height - legendH);
  plotEl.style.height = plotH + 'px';
  plotEl.style.maxHeight = plotH + 'px';
  plotEl.style.width = '100%';

  const plotW = Math.max(160, Math.floor(host.clientWidth || plotEl.clientWidth || 280));
  Plotly.newPlot(plotEl, [{{
    type:'pie', labels:spec.labels, values:spec.values, marker:{{colors:spec.colors}},
    customdata: custom, textfont:{{size:10}},
    domain: {{ x: [0.12, 0.88], y: [0.08, 0.92] }},
    hovertemplate: '%{{label}}<br>%{{percent}}<br>%{{customdata[0]}}<extra></extra>'
  }}], Object.assign({{}}, L, {{
    showlegend:false,
    autosize:false,
    height: plotH,
    width: plotW,
    margin: {{ t: 4, b: 4, l: 4, r: 4 }},
    legend: {{}}
  }}), {{responsive:false, displayModeBar:false}}).then(() => {{
    plotEl.style.height = plotH + 'px';
    plotEl.style.maxHeight = plotH + 'px';
    const used = Math.ceil(
      plotEl.getBoundingClientRect().height + legendEl.getBoundingClientRect().height
    );
    if (used > height) {{
      const fixH = Math.max(100, height - Math.ceil(legendEl.getBoundingClientRect().height));
      plotEl.style.height = fixH + 'px';
      plotEl.style.maxHeight = fixH + 'px';
      return Plotly.relayout(plotEl, {{ height: fixH, width: plotW }});
    }}
  }}).then(() => bindChartClick(plotEl.id, kind));
}}
function drawPies() {{
  pie('pieExp', DATA.expensePie, 'expense');
  pie('pieInc', DATA.incomePie, 'income');
}}
function equalizeSplitRowHeights() {{
  const row = document.querySelector('#cockpitView > .split');
  if (!row) return;
  // Stacked layout on narrow screens: let natural height win.
  if (window.matchMedia('(max-width: 900px)').matches) {{
    for (const child of row.children) {{
      child.style.minHeight = '';
      child.style.height = '';
    }}
    return;
  }}
  const kids = Array.from(row.children);
  if (kids.length < 2) return;
  for (const child of kids) {{
    child.style.minHeight = '';
    child.style.height = '';
  }}
  // Force reflow, then lock all columns to the tallest natural height.
  void row.offsetHeight;
  let maxH = 0;
  for (const child of kids) {{
    maxH = Math.max(maxH, Math.ceil(child.getBoundingClientRect().height));
  }}
  if (maxH < 40) return;
  for (const child of kids) {{
    child.style.minHeight = maxH + 'px';
    child.style.height = maxH + 'px';
  }}
}}
function drawAll() {{
  const L = baseLayout();
  drawPies();
  // Second pass after flex/grid settle so pie height fits the panel (avoids bottom overflow).
  requestAnimationFrame(() => requestAnimationFrame(() => {{
    drawPies();
    drawRecurringTreemap();
    equalizeSplitRowHeights();
  }}));
  const barBal = document.getElementById('barBal');
  if (barBal) {{
    const days = DATA.series.map(x => x.day);
    const labels = days.map(dayTickLabel);
    const bals = DATA.series.map(x => x.balance);
    const colors = bals.map(v => v >= 0 ? '#27ae60' : '#c0392b');
    const tickangle = days.length > 14 ? -45 : 0;
    Plotly.newPlot('barBal', [{{
      type:'bar',
      x: labels,
      y: bals,
      marker: {{color: colors}},
      customdata: DATA.series.map(x => [
        x.day, formatBrl(x.income), formatBrl(x.expense), formatBrl(x.balance)
      ]),
      hovertemplate:
        '%{{customdata[0]}}<br>Balance %{{customdata[3]}}'
        + '<br>In %{{customdata[1]}} · Out %{{customdata[2]}}<extra></extra>'
    }}], Object.assign({{}}, L, {{
      showlegend:false,
      margin: Object.assign({{}}, L.margin, {{b: days.length > 14 ? 64 : 48}}),
      xaxis: {{tickangle: tickangle, automargin:true, tickfont:{{size:10}}}}
    }}), {{responsive:true, displayModeBar:false}});
  }}
  const lineYear = document.getElementById('lineYear');
  if (lineYear) {{
    Plotly.newPlot('lineYear', [
      {{type:'scatter', mode:'lines+markers', name:'In', x:DATA.annual.map(x=>x.label), y:DATA.annual.map(x=>x.income), line:{{color:'#2ecc71'}}}},
      {{type:'scatter', mode:'lines+markers', name:'Out', x:DATA.annual.map(x=>x.label), y:DATA.annual.map(x=>x.expense), line:{{color:'#e74c3c'}}}},
      {{type:'scatter', mode:'lines+markers', name:'Bal', x:DATA.annual.map(x=>x.label), y:DATA.annual.map(x=>x.balance), line:{{color:'#f1c40f'}}}}
    ], L, {{responsive:true, displayModeBar:false}});
  }}
  drawIncomeInvestTreemap();
  drawRecurringTreemap();
  refreshCardsFromApi().then(() => drawCardInstallmentChart());
  equalizeSplitRowHeights();
}}
function rebuildCardInstallmentRemaining() {{
  const today = (RAW.today || '').slice(0, 10) || new Date().toISOString().slice(0, 10);
  function daysInMonth(y, m) {{
    if (m === 2) {{
      const leap = (y % 4 === 0 && (y % 100 !== 0 || y % 400 === 0));
      return leap ? 29 : 28;
    }}
    return [4, 6, 9, 11].includes(m) ? 30 : 31;
  }}
  function clampClosingDay(y, m, closingDay) {{
    let cd = parseInt(closingDay, 10);
    if (!Number.isFinite(cd) || cd < 1) cd = 1;
    if (cd > 31) cd = 31;
    return Math.min(cd, daysInMonth(y, m));
  }}
  function closingDateOn(y, m, closingDay) {{
    const cd = clampClosingDay(y, m, closingDay);
    return String(y).padStart(4, '0') + '-' + String(m).padStart(2, '0') + '-' + String(cd).padStart(2, '0');
  }}
  function billClosingDate(txDate, closingDay) {{
    let d = String(txDate || '').slice(0, 10);
    if (d.length < 10) d = today;
    let [y, m, day] = d.split('-').map(Number);
    const cd = clampClosingDay(y, m, closingDay);
    if (day <= cd) return closingDateOn(y, m, closingDay);
    m += 1;
    if (m > 12) {{ m = 1; y += 1; }}
    return closingDateOn(y, m, closingDay);
  }}
  function nextClosingOnOrAfter(dayStr, closingDay) {{
    let d = String(dayStr || '').slice(0, 10);
    if (d.length < 10) d = today;
    let [y, m, day] = d.split('-').map(Number);
    const cd = clampClosingDay(y, m, closingDay);
    if (day <= cd) return closingDateOn(y, m, closingDay);
    m += 1;
    if (m > 12) {{ m = 1; y += 1; }}
    return closingDateOn(y, m, closingDay);
  }}
  function previousClosingOnOrBefore(dayStr, closingDay) {{
    let d = String(dayStr || '').slice(0, 10);
    if (d.length < 10) d = today;
    let [y, m, day] = d.split('-').map(Number);
    const cd = clampClosingDay(y, m, closingDay);
    if (day >= cd) return closingDateOn(y, m, closingDay);
    m -= 1;
    if (m < 1) {{ m = 12; y -= 1; }}
    return closingDateOn(y, m, closingDay);
  }}
  function dueDateForClosing(closingDate, dueDay) {{
    let d = String(closingDate || '').slice(0, 10);
    if (d.length < 10) d = today;
    let [y, m, closeDom] = d.split('-').map(Number);
    let dd = parseInt(dueDay, 10);
    if (!Number.isFinite(dd) || dd < 1) dd = 1;
    if (dd > 31) dd = 31;
    if (dd > closeDom) return closingDateOn(y, m, dd);
    m += 1;
    if (m > 12) {{ m = 1; y += 1; }}
    return closingDateOn(y, m, dd);
  }}
  function cardDueDay(c, closingDay) {{
    let dd = parseInt(c && c.due_day, 10);
    if (Number.isFinite(dd) && dd >= 1 && dd <= 31) return dd;
    dd = (parseInt(closingDay, 10) || 1) + 7;
    if (dd > 31) dd -= 31;
    return Math.max(1, Math.min(31, dd));
  }}
  function nextPayDate(dayStr, closingDay, dueDay) {{
    const prev = previousClosingOnOrBefore(dayStr, closingDay);
    const due = dueDateForClosing(prev, dueDay);
    if (due >= String(dayStr || today).slice(0, 10)) return due;
    return dueDateForClosing(nextClosingOnOrAfter(dayStr, closingDay), dueDay);
  }}

  const closingByCard = {{}};
  const dueByCard = {{}};
  const spentByCard = {{}};
  const cardMeta = {{}};
  for (const c of RAW.cards || []) {{
    if (!c.id) continue;
    cardMeta[c.id] = c.name || c.id;
    spentByCard[c.id] = parseDecimal(c.current_spent);
    closingByCard[c.id] = parseInt(c.closing_day, 10) || 1;
    dueByCard[c.id] = cardDueDay(c, closingByCard[c.id]);
  }}
  for (const c of RAW.liquidCards || []) {{
    if (c.id && typeof c.spent === 'number') spentByCard[c.id] = c.spent;
  }}

  const unpaid = [];
  for (const t of RAW.transactions || []) {{
    if (t.type !== 'card_expense') continue;
    const p = String(t.paid || '0').trim().toLowerCase();
    if (p === '1' || p === 'true' || p === 'yes') continue;
    const cid = (t.card_id || '').trim();
    const d = String(t.date || '').slice(0, 10);
    if (!cid || d.length < 10) continue;
    const cd = closingByCard[cid] || 1;
    unpaid.push({{
      card_id: cid,
      date: d,
      closing: billClosingDate(d, cd),
      amount: parseDecimal(t.amount)
    }});
  }}

  const empty = {{
    months: [], closings: [], series: [], today, from_today: [], from_today_total: 0,
    open_total: 0, later_total: 0, last_date: ''
  }};
  const activeIds = Object.keys(spentByCard).filter(cid =>
    (spentByCard[cid] || 0) > 0.00001 || unpaid.some(u => u.card_id === cid)
  );
  if (!activeIds.length) {{
    DATA.cardInstallmentRemaining = empty;
    return;
  }}

  const palette = ['#e74c3c', '#3498db', '#9b59b6', '#f39c12', '#1abc9c', '#e67e22'];
  const series = [];
  const fromToday = [];
  let lastDate = '';
  activeIds.forEach((cid, i) => {{
    const totalAmt = Math.round((spentByCard[cid] || 0) * 100) / 100;
    const cd = closingByCard[cid] || 1;
    const openClose = nextClosingOnOrAfter(today, cd);
    const dueByClose = {{}};
    for (const u of unpaid) {{
      if (u.card_id !== cid) continue;
      dueByClose[u.closing] = (dueByClose[u.closing] || 0) + u.amount;
    }}
    if (dueByClose[openClose] == null) dueByClose[openClose] = 0;
    let closings = Object.keys(dueByClose).sort();
    let rawOpen = 0, rawLater = 0;
    for (const c of closings) {{
      if (c <= openClose) rawOpen += dueByClose[c];
      else rawLater += dueByClose[c];
    }}
    rawOpen = Math.round(rawOpen * 100) / 100;
    rawLater = Math.round(rawLater * 100) / 100;
    const rawSum = rawOpen + rawLater;
    let openAmt, laterAmt, dueVals = {{}};
    if (rawSum > 0.00001 && Math.abs(rawSum - totalAmt) > 0.02) {{
      const scale = totalAmt / rawSum;
      openAmt = Math.round(rawOpen * scale * 100) / 100;
      laterAmt = Math.round(Math.max(0, totalAmt - openAmt) * 100) / 100;
      for (const c of closings) dueVals[c] = Math.round(dueByClose[c] * scale * 100) / 100;
    }} else {{
      openAmt = totalAmt ? Math.min(rawOpen, totalAmt) : rawOpen;
      laterAmt = Math.round(Math.max(0, totalAmt - openAmt) * 100) / 100;
      for (const c of closings) dueVals[c] = Math.round(dueByClose[c] * 100) / 100;
      if (totalAmt > 0 && rawSum <= 0.00001) {{
        dueVals = {{}};
        dueVals[openClose] = totalAmt;
        closings = [openClose];
        openAmt = totalAmt;
        laterAmt = 0;
      }}
    }}
    const cardLast = closings.filter(c => (dueVals[c] || 0) > 0.00001).sort().slice(-1)[0] || '';
    if (cardLast && (!lastDate || cardLast > lastDate)) lastDate = cardLast;
    const color = palette[i % palette.length];
    const dd = dueByCard[cid] || cardDueDay(null, cd);
    const nextPay = nextPayDate(today, cd, dd);
    fromToday.push({{
      card_id: cid,
      name: cardMeta[cid] || cid,
      color,
      amount: totalAmt,
      open: openAmt,
      later: laterAmt,
      last_date: cardLast,
      next_closing: openClose,
      next_pay: nextPay,
      closing_day: cd,
      due_day: dd
    }});
    series.push({{
      card_id: cid,
      name: cardMeta[cid] || cid,
      color,
      dates: closings,
      values: closings.map(c => dueVals[c] || 0)
    }});
  }});

  const allClosings = Array.from(new Set(series.flatMap(s => s.dates || []))).sort();
  DATA.cardInstallmentRemaining = {{
    months: allClosings,
    closings: allClosings,
    series,
    today,
    from_today: fromToday,
    from_today_total: Math.round(fromToday.reduce((s, r) => s + r.amount, 0) * 100) / 100,
    open_total: Math.round(fromToday.reduce((s, r) => s + r.open, 0) * 100) / 100,
    later_total: Math.round(fromToday.reduce((s, r) => s + r.later, 0) * 100) / 100,
    last_date: lastDate
  }};
}}
function cardPlanDefaultRange() {{
  const today = (RAW.today || '').slice(0, 10) || new Date().toISOString().slice(0, 10);
  let [y, m, day] = today.split('-').map(Number);
  m += 2;
  while (m > 12) {{ m -= 12; y += 1; }}
  const dim = (yy, mm) => {{
    if (mm === 2) {{
      const leap = (yy % 4 === 0 && (yy % 100 !== 0 || yy % 400 === 0));
      return leap ? 29 : 28;
    }}
    return [4, 6, 9, 11].includes(mm) ? 30 : 31;
  }};
  const dd = Math.min(day, dim(y, m));
  const to = String(y).padStart(4, '0') + '-' + String(m).padStart(2, '0') + '-'
    + String(dd).padStart(2, '0');
  return {{ from: today, to }};
}}
function enumerateDays(from, to) {{
  const out = [];
  if (!from || !to || from > to) return out;
  let d = from;
  // Cap at ~400 days to avoid huge plots.
  for (let i = 0; i < 400 && d <= to; i++) {{
    out.push(d);
    d = shiftDate(d, 1);
  }}
  return out;
}}
function ensureCardPlanRangeInputs() {{
  const fromEl = document.getElementById('cardPlanFrom');
  const toEl = document.getElementById('cardPlanTo');
  if (!fromEl || !toEl) return cardPlanDefaultRange();
  if (!fromEl.value || !toEl.value) {{
    const def = cardPlanDefaultRange();
    fromEl.value = def.from;
    toEl.value = def.to;
  }}
  let from = fromEl.value;
  let to = toEl.value;
  if (from > to) {{
    const t = from; from = to; to = t;
    fromEl.value = from; toEl.value = to;
  }}
  return {{ from, to }};
}}
function drawCardInstallmentChart() {{
  const el = document.getElementById('lineCardPlan');
  const summaryEl = document.getElementById('cardPlanSummary');
  if (!el) return;
  rebuildCardInstallmentRemaining();
  const spec = DATA.cardInstallmentRemaining || {{
    months: [], closings: [], series: [], from_today: [], from_today_total: 0, open_total: 0, later_total: 0
  }};
  const rows = (spec.from_today || []).filter(r => r.amount > 0);
  if (summaryEl) {{
    if (!rows.length) {{
      summaryEl.innerHTML = '<div class="card-plan-box"><div class="box-name">All cards</div>'
        + '<div class="box-total">—</div>'
        + '<div class="box-split">Nothing to pay on cards</div></div>';
    }} else {{
      let html = rows.map(r =>
        '<div class="card-plan-box">'
        + '<div class="box-top"><span class="dot" style="background:' + (r.color || '#888') + '"></span>'
        + '<span class="box-name">' + r.name + '</span></div>'
        + '<div class="box-total">' + formatBrl(r.amount) + '</div>'
        + '<div class="box-split">Open ' + formatBrl(r.open || 0)
        + ' · Later ' + formatBrl(r.later || 0)
        + (r.next_closing ? '<br>Next close ' + r.next_closing : '')
        + (r.next_pay ? ' · Next pay ' + r.next_pay : '')
        + (r.last_date ? '<br>Until ' + r.last_date : '')
        + (r.closing_day ? ' · Closes day ' + r.closing_day : '')
        + (r.due_day ? ' · Pays day ' + r.due_day : '')
        + '</div></div>'
      ).join('');
      html += '<div class="card-plan-box total">'
        + '<div class="box-top"><span class="box-name">Total</span></div>'
        + '<div class="box-total">' + formatBrl(spec.from_today_total || 0) + '</div>'
        + '<div class="box-split">Open ' + formatBrl(spec.open_total || 0)
        + ' · Later ' + formatBrl(spec.later_total || 0) + '</div></div>';
      summaryEl.innerHTML = html;
    }}
  }}
  const range = ensureCardPlanRangeInputs();
  const dayAxisRequested = enumerateDays(range.from, range.to);
  if (!dayAxisRequested.length) {{
    el.innerHTML = '<p class="empty">Invalid card plan range</p>';
    return;
  }}
  // Extend the daily axis so every pay day with an amount in/near the selected
  // window is present (Plotly needs 2+ points to draw a line).
  function payDateForClosing(closingDate, dueDay) {{
    let d = String(closingDate || '').slice(0, 10);
    if (d.length < 10) return '';
    let [y, m, closeDom] = d.split('-').map(Number);
    let dd = parseInt(dueDay, 10);
    if (!Number.isFinite(dd) || dd < 1) dd = 1;
    if (dd > 31) dd = 31;
    const dim = (yy, mm) => {{
      if (mm === 2) {{
        const leap = (yy % 4 === 0 && (yy % 100 !== 0 || yy % 400 === 0));
        return leap ? 29 : 28;
      }}
      return [4, 6, 9, 11].includes(mm) ? 30 : 31;
    }};
    const clamp = (yy, mm, day) => Math.min(Math.max(1, day), dim(yy, mm));
    if (dd > closeDom) {{
      return String(y).padStart(4, '0') + '-' + String(m).padStart(2, '0') + '-'
        + String(clamp(y, m, dd)).padStart(2, '0');
    }}
    m += 1;
    if (m > 12) {{ m = 1; y += 1; }}
    return String(y).padStart(4, '0') + '-' + String(m).padStart(2, '0') + '-'
      + String(clamp(y, m, dd)).padStart(2, '0');
  }}
  const dueDayByCardEarly = {{}};
  for (const r of (spec.from_today || [])) {{
    if (r.card_id) dueDayByCardEarly[r.card_id] = r.due_day;
  }}
  let axisFrom = range.from;
  let axisTo = range.to;
  for (const s of (spec.series || [])) {{
    const dueDay = dueDayByCardEarly[s.card_id] || 1;
    const dates = s.dates || [];
    const values = s.values || [];
    for (let i = 0; i < dates.length; i++) {{
      const closeD = dates[i];
      if (!closeD) continue;
      const amt = values[i] || 0;
      if (amt <= 0.00001) continue;
      const payD = payDateForClosing(closeD, dueDay);
      if (!payD) continue;
      const inWindow = (payD >= range.from && payD <= range.to)
        || (closeD >= range.from && closeD <= range.to);
      if (!inWindow) continue;
      if (payD < axisFrom) axisFrom = payD;
      if (payD > axisTo) axisTo = payD;
    }}
  }}
  const dayAxis = enumerateDays(axisFrom, axisTo);
  if (!dayAxis.length) {{
    el.innerHTML = '<p class="empty">Invalid card plan range</p>';
    return;
  }}
  const L = baseLayout();
  let yMax = 0;
  for (const s of (spec.series || [])) {{
    for (const v of (s.values || [])) if (v > yMax) yMax = v;
  }}
  const tickCount = 5;
  const rough = yMax > 0 ? yMax / (tickCount - 1) : 1;
  const mag = Math.pow(10, Math.floor(Math.log10(rough || 1)));
  const step = Math.max(mag, Math.ceil(rough / mag) * mag);
  const tickvals = [];
  for (let v = 0; v <= yMax + step * 0.01; v += step) tickvals.push(v);
  if (!tickvals.length) tickvals.push(0);
  const today = (spec.today || RAW.today || '').slice(0, 10);
  const traces = (spec.series || []).map(s => {{
    const dueDay = dueDayByCardEarly[s.card_id] || 1;
    const byPay = {{}};
    const closeByPay = {{}};
    (s.dates || []).forEach((closeD, i) => {{
      const amt = s.values[i] || 0;
      if (amt <= 0.00001) return;
      const payD = payDateForClosing(closeD, dueDay);
      if (!payD) return;
      byPay[payD] = (byPay[payD] || 0) + amt;
      closeByPay[payD] = closeD;
    }});
    const ys = dayAxis.map(d => (Object.prototype.hasOwnProperty.call(byPay, d) ? byPay[d] : null));
    const custom = ys.map((v, i) => {{
      if (v == null) return '';
      const payD = dayAxis[i];
      const closeD = closeByPay[payD] || '';
      return formatBrl(v) + (closeD ? ' (close ' + closeD + ')' : '');
    }});
    return {{
      type: 'scatter',
      mode: 'lines+markers',
      name: s.name,
      x: dayAxis,
      y: ys,
      customdata: custom,
      connectgaps: true,
      line: {{ color: s.color, width: 2 }},
      marker: {{ color: s.color, size: 7 }},
      hovertemplate: '%{{x}}<br>%{{fullData.name}}: %{{customdata}}<extra></extra>'
    }};
  }});
  if (today && today >= axisFrom && today <= axisTo) {{
    traces.push({{
      type: 'scatter',
      mode: 'markers',
      name: 'Today',
      x: [today],
      y: [0],
      marker: {{
        symbol: 'triangle-up',
        size: 14,
        color: '#f1c40f',
        line: {{ width: 1, color: '#ffffff' }}
      }},
      hovertemplate: 'Today %{{x}}<extra></extra>',
      cliponaxis: false
    }});
  }}
  const dueDayByCard = dueDayByCardEarly;
  const shapes = [];
  const annotations = [];
  const seenPayLines = {{}};
  if (today && today >= axisFrom && today <= axisTo) {{
    shapes.push({{
      type: 'line',
      x0: today,
      x1: today,
      y0: 0,
      y1: 1,
      yref: 'paper',
      line: {{ color: '#f1c40f', width: 1.5, dash: 'dot' }}
    }});
    annotations.push({{
      x: today,
      y: 1,
      yref: 'paper',
      text: 'Today',
      showarrow: false,
      yanchor: 'bottom',
      font: {{ size: 10, color: '#f1c40f' }},
      xanchor: 'center'
    }});
  }}
  for (const s of spec.series || []) {{
    const color = s.color || '#888';
    const dates = s.dates || [];
    const values = s.values || [];
    const dueDay = dueDayByCard[s.card_id] || 1;
    for (let i = 0; i < dates.length; i++) {{
      const closeD = dates[i];
      if (!closeD) continue;
      const amt = values[i] || 0;
      if (amt <= 0.00001) continue;
      const payD = payDateForClosing(closeD, dueDay);
      if (!payD || payD < range.from || payD > range.to) continue;
      const key = payD + '|' + (s.card_id || s.name || color);
      if (seenPayLines[key]) {{
        seenPayLines[key].amt += amt;
        continue;
      }}
      seenPayLines[key] = {{ payD, color, amt, name: s.name || 'Card', cardId: s.card_id || s.name }};
    }}
  }}
  for (const key of Object.keys(seenPayLines)) {{
    const p = seenPayLines[key];
    shapes.push({{
      type: 'line',
      x0: p.payD,
      x1: p.payD,
      y0: 0,
      y1: 1,
      yref: 'paper',
      line: {{ color: p.color, width: 1.5, dash: 'dot' }}
    }});
    traces.push({{
      type: 'scatter',
      mode: 'markers',
      name: p.name + ' pay',
      x: [p.payD],
      y: [0],
      marker: {{
        symbol: 'triangle-down',
        size: 12,
        color: p.color,
        line: {{ width: 1, color: '#ffffff' }}
      }},
      hovertemplate: p.name + ' pay %{{x}}<br>' + formatBrl(p.amt) + '<extra></extra>',
      cliponaxis: false,
      legendgroup: p.cardId,
      showlegend: false
    }});
    annotations.push({{
      x: p.payD,
      y: 1,
      yref: 'paper',
      text: formatBrl(p.amt),
      showarrow: false,
      yanchor: 'bottom',
      xanchor: 'center',
      font: {{ size: 10, color: p.color }},
      bgcolor: 'rgba(0,0,0,0)',
      borderpad: 1
    }});
  }}
  // Tick every ~7 days (plus endpoints) so labels stay readable on a full daily axis.
  const tickEvery = dayAxis.length > 45 ? 7 : (dayAxis.length > 20 ? 3 : 1);
  const tickvalsX = dayAxis.filter((d, i) =>
    i === 0 || i === dayAxis.length - 1 || i % tickEvery === 0 || d === today
  );
  Plotly.newPlot('lineCardPlan', traces, Object.assign({{}}, L, {{
    showlegend: true,
    legend: {{ orientation: 'h', y: 1.14, x: 0, font: {{ size: 10 }} }},
    margin: {{ t: 64, b: 64, l: 88, r: 20 }},
    shapes: shapes,
    annotations: annotations,
    yaxis: {{
      title: {{ text: '' }},
      automargin: true,
      tickmode: 'array',
      tickvals: tickvals,
      ticktext: tickvals.map(v => formatBrl(v)),
      tickfont: {{ size: 10 }},
      rangemode: 'tozero'
    }},
    xaxis: {{
      type: 'category',
      categoryorder: 'array',
      categoryarray: dayAxis,
      title: {{ text: 'Day', font: {{ size: 11 }} }},
      automargin: true,
      tickmode: 'array',
      tickvals: tickvalsX,
      ticktext: tickvalsX.map(d => d.slice(5)),
      tickfont: {{ size: 9 }},
      tickangle: -45
    }}
  }}), {{ responsive: true, displayModeBar: false }});
}}
function applyTheme(theme) {{
  document.documentElement.setAttribute('data-theme', theme);
  localStorage.setItem(THEME_KEY, theme);
  const btn = document.getElementById('themeToggle');
  if (btn) btn.textContent = theme === 'dark' ? 'Light' : 'Dark';
  if (!activeCategory) drawAll();
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
  const backBtn = document.getElementById('catViewBack');
  function goBackToDashboard() {{
    if (!activeCategory) return;
    closeCategoryView();
    drawAll();
  }}
  if (backBtn) backBtn.addEventListener('click', goBackToDashboard);
  document.addEventListener('keydown', (e) => {{
    if (e.key !== 'Backspace') return;
    const tag = (e.target && e.target.tagName) ? e.target.tagName.toLowerCase() : '';
    if (tag === 'input' || tag === 'textarea' || tag === 'select' || (e.target && e.target.isContentEditable))
      return;
    if (!activeCategory) return;
    e.preventDefault();
    goBackToDashboard();
  }});
  const fromEl = document.getElementById('periodFrom');
  const toEl = document.getElementById('periodTo');
  if (fromEl) fromEl.value = RAW.monthStart || RAW.dateFrom;
  if (toEl) toEl.value = RAW.today || RAW.dateTo;
  const cardPlanFrom = document.getElementById('cardPlanFrom');
  const cardPlanTo = document.getElementById('cardPlanTo');
  const cardPlanApply = document.getElementById('cardPlanApply');
  const cardPlanReset = document.getElementById('cardPlanReset');
  if (cardPlanFrom && cardPlanTo) {{
    const def = cardPlanDefaultRange();
    if (!cardPlanFrom.value) cardPlanFrom.value = def.from;
    if (!cardPlanTo.value) cardPlanTo.value = def.to;
  }}
  if (cardPlanApply) {{
    cardPlanApply.addEventListener('click', () => {{
      refreshCardsFromApi().then(() => drawCardInstallmentChart());
    }});
  }}
  if (cardPlanReset) {{
    cardPlanReset.addEventListener('click', () => {{
      const def = cardPlanDefaultRange();
      if (cardPlanFrom) cardPlanFrom.value = def.from;
      if (cardPlanTo) cardPlanTo.value = def.to;
      refreshCardsFromApi().then(() => drawCardInstallmentChart());
    }});
  }}
  const budgetMonthSel = document.getElementById('budgetMonth');
  if (budgetMonthSel) {{
    budgetMonthSel.addEventListener('change', () => {{
      renderBudgets(budgetsForMonth(budgetMonthSel.value));
    }});
  }}
  let pieResizeTimer = 0;
  window.addEventListener('resize', () => {{
    if (pieResizeTimer) clearTimeout(pieResizeTimer);
    pieResizeTimer = setTimeout(() => {{
      if (!activeCategory) {{
        drawPies();
        drawRecurringTreemap();
        equalizeSplitRowHeights();
      }}
    }}, 120);
  }});
  checkBudgetApi().then(() => refreshCardsFromApi()).then(() => {{
    applyPeriod();
    if (budgetApiOk === false) {{
      setBudgetSaveStatus('Save unavailable (open via Finance dashboard)', 'error');
    }}
  }});
}})();
</script>
</body></html>"""


def main(argv: list[str] | None = None):
    import argparse

    parser = argparse.ArgumentParser(description="Build finance cockpit dashboard.html")
    parser.add_argument("--data-dir", default="", help="Absolute path to finances/data")
    parser.add_argument(
        "--output-dir", default="", help="Absolute path to finances/output"
    )
    args = parser.parse_args(argv)
    if args.data_dir or args.output_dir:
        configure_paths(
            data_dir=args.data_dir or None,
            output_dir=args.output_dir or None,
        )
    seed()
    out_dir = _agg.OUTPUT
    out_dir.mkdir(parents=True, exist_ok=True)
    html = build_html(snapshot())
    out = out_dir / "dashboard.html"
    out.write_text(html, encoding="utf-8")
    print(str(out.resolve()))


if __name__ == "__main__":
    main()
