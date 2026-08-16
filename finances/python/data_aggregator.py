"""Read finance CSVs with Brazilian comma decimals."""

from __future__ import annotations

import csv
import configparser
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
DATA = ROOT / "finances" / "data"
OUTPUT = ROOT / "finances" / "output"


def parse_decimal(value) -> float:
    if value is None:
        return 0.0
    s = str(value).strip().replace("R$", "").replace(" ", "")
    if not s:
        return 0.0
    sign = 1.0
    if s[0] in "+-":
        if s[0] == "-":
            sign = -1.0
        s = s[1:]
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        parts = s.split(",")
        if len(parts) == 2 and len(parts[1]) <= 2:
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    try:
        return sign * float(s)
    except ValueError:
        return 0.0


def format_brl(num: float) -> str:
    neg = num < 0
    n = abs(num)
    formatted = f"{n:,.2f}"
    formatted = formatted.replace(",", "X").replace(".", ",").replace("X", ".")
    return ("-R$ " if neg else "R$ ") + formatted


def read_csv(name: str) -> list[dict]:
    path = DATA / name
    if not path.exists():
        return []
    with path.open(encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def read_settings() -> dict:
    cfg = configparser.ConfigParser()
    path = DATA / "settings.ini"
    if path.exists():
        cfg.read(path, encoding="utf-8")
    dash = dict(cfg["Dashboard"]) if cfg.has_section("Dashboard") else {}
    gen = dict(cfg["General"]) if cfg.has_section("General") else {}
    return {"dashboard": dash, "general": gen}


def widget_on(settings: dict, key: str) -> bool:
    return settings.get("dashboard", {}).get(key, "1") != "0"


def current_month() -> str:
    return datetime.now().strftime("%Y-%m")


def month_shift(ym: str, delta: int) -> str:
    y, m = [int(x) for x in ym.split("-")]
    m += delta
    while m > 12:
        m -= 12
        y += 1
    while m < 1:
        m += 12
        y -= 1
    return f"{y:04d}-{m:02d}"


def cat_index(cats: list[dict]) -> dict[str, dict]:
    return {c.get("id", ""): c for c in cats}


def cat_label(row: dict | None, fallback: str = "") -> str:
    if not row:
        return fallback
    icon = (row.get("icon") or "").strip()
    name = row.get("name") or fallback
    if icon and name:
        return f"{icon} {name}"
    return name or icon or fallback


def main_category_id(cat_id: str, by_id: dict[str, dict]) -> str:
    c = by_id.get(cat_id)
    if not c:
        return cat_id
    parent = (c.get("parent_id") or "").strip()
    return parent or cat_id


def month_totals(txs: list[dict], ym: str) -> dict:
    income = expense = 0.0
    for t in txs:
        if not str(t.get("date", "")).startswith(ym):
            continue
        amt = parse_decimal(t.get("amount"))
        kind = t.get("type", "")
        if kind == "income":
            income += amt
        elif kind in ("expense", "card_expense"):
            expense += amt
    return {"income": income, "expense": expense, "balance": income - expense}


def period_totals(txs: list[dict], months: list[str]) -> dict:
    income = expense = 0.0
    for ym in months:
        tot = month_totals(txs, ym)
        income += tot["income"]
        expense += tot["expense"]
    return {"income": income, "expense": expense, "balance": income - expense}


def by_category(
    txs: list[dict], ym: str, types: set[str], cats: list[dict]
) -> list[tuple[str, float, str]]:
    return by_category_months(txs, [ym], types, cats)


def by_category_months(
    txs: list[dict], months: list[str], types: set[str], cats: list[dict]
) -> list[tuple[str, float, str]]:
    by_id = cat_index(cats)
    prefixes = tuple(months)
    totals: dict[str, float] = defaultdict(float)
    for t in txs:
        d = str(t.get("date", ""))
        if not any(d.startswith(ym) for ym in prefixes):
            continue
        if t.get("type") not in types:
            continue
        cid = main_category_id(t.get("category_id", ""), by_id)
        totals[cid] += parse_decimal(t.get("amount"))
    rows = []
    for cid, amt in totals.items():
        row = by_id.get(cid)
        name = cat_label(row, cid or "Uncategorized")
        color = (row or {}).get("color", "#7F8C8D")
        rows.append((name, amt, color))
    rows.sort(key=lambda r: r[1], reverse=True)
    return rows


def monthly_series(txs: list[dict], months: list[str] | None = None) -> list[dict]:
    if months is None:
        months = sorted({str(t.get("date", ""))[:7] for t in txs if t.get("date")})
    else:
        months = sorted(months)
    out = []
    for ym in months:
        tot = month_totals(txs, ym)
        tot["month"] = ym
        out.append(tot)
    return out


def annual_net(txs: list[dict], year: str) -> list[dict]:
    out = []
    for m in range(1, 13):
        ym = f"{year}-{m:02d}"
        tot = month_totals(txs, ym)
        tot["month"] = ym
        tot["label"] = datetime(int(year), m, 1).strftime("%b")
        out.append(tot)
    return out


def prior_period_months(months: list[str]) -> list[str]:
    months = sorted(months)
    if not months:
        return []
    n = len(months)
    earliest = months[0]
    return [month_shift(earliest, -i) for i in range(n, 0, -1)]


def period_label(months: list[str]) -> str:
    months = sorted(set(months))
    if not months:
        return current_month()
    if len(months) == 1:
        return months[0]
    year = months[0][:4]
    names = [datetime(2000, int(ym[5:7]), 1).strftime("%b") for ym in months]
    return f"{year} ({', '.join(names)})"


def aggregate_budgets(budgets: list[dict], months: list[str]) -> list[dict]:
    month_set = set(months)
    planned: dict[str, float] = defaultdict(float)
    spent: dict[str, float] = defaultdict(float)
    for b in budgets:
        if b.get("year_month") not in month_set:
            continue
        cid = b.get("category_id", "")
        planned[cid] += parse_decimal(b.get("planned_amount"))
        spent[cid] += parse_decimal(b.get("spent_amount"))
    label = period_label(months)
    rows = []
    for cid in sorted(planned.keys() | spent.keys()):
        rows.append(
            {
                "year_month": label,
                "category_id": cid,
                "planned_amount": f"{planned[cid]:.2f}".replace(".", ","),
                "spent_amount": f"{spent[cid]:.2f}".replace(".", ","),
            }
        )
    return rows


def collect_notifications(
    settings: dict,
    budgets: list[dict],
    cats: list[dict],
    cards: list[dict],
    goals: list[dict],
    months: list[str],
) -> list[str]:
    notes = []
    by_id = cat_index(cats)
    month_set = set(months)
    if settings.get("general", {}).get("NotifyBudgetExceeded", "1") != "0":
        for b in aggregate_budgets(budgets, list(month_set)):
            planned = parse_decimal(b.get("planned_amount"))
            spent = parse_decimal(b.get("spent_amount"))
            if planned > 0 and spent > planned:
                name = cat_label(
                    by_id.get(b.get("category_id", "")),
                    b.get("category_id", ""),
                )
                notes.append(
                    f"Budget exceeded: {name} ({format_brl(spent)} / {format_brl(planned)})"
                )
    if settings.get("general", {}).get("NotifyCardHighUsage", "1") != "0":
        warn = parse_decimal(settings.get("general", {}).get("CardUsageWarnPct", "80"))
        for c in cards:
            lim = parse_decimal(c.get("limit"))
            spent = parse_decimal(c.get("current_spent"))
            if lim > 0 and spent / lim * 100 >= warn:
                notes.append(
                    f"Card {c.get('name')} at {spent / lim * 100:.0f}% of limit"
                )
    today = datetime.now().strftime("%Y-%m-%d")
    for g in goals:
        tdate = (g.get("target_date") or "").strip()
        if tdate and tdate < today:
            cur = parse_decimal(g.get("current_amount"))
            tgt = parse_decimal(g.get("target_amount"))
            if tgt <= 0 or cur < tgt:
                notes.append(f"Goal past target date: {g.get('name')}")
    return notes


def snapshot(
    months: list[str] | None = None,
    year: str | None = None,
    date_from: str | None = None,
    date_to: str | None = None,
) -> dict:
    txs = read_csv("transactions.csv")
    accs = read_csv("accounts.csv")
    cats = read_csv("categories.csv")
    cards = read_csv("credit_cards.csv")
    goals = read_csv("goals.csv")
    budgets = read_csv("budgets.csv")
    settings = read_settings()
    today = datetime.now().strftime("%Y-%m-%d")
    month_start = datetime.now().strftime("%Y-%m-01")
    if date_from is None and date_to is None and months is None:
        date_from = month_start
        date_to = today
    if date_from and date_to:
        if date_from > date_to:
            date_from, date_to = date_to, date_from
        months = []
        cur = date_from[:7]
        end = date_to[:7]
        while cur <= end:
            months.append(cur)
            cur = month_shift(cur, 1)
        year = year or date_to[:4]
        tot = period_totals_dates(txs, date_from, date_to)
        # Prior window: same number of days ending the day before date_from
        n_days = (
            datetime.strptime(date_to, "%Y-%m-%d")
            - datetime.strptime(date_from, "%Y-%m-%d")
        ).days + 1
        prior_to_d = datetime.strptime(date_from, "%Y-%m-%d")
        prior_to = (prior_to_d - timedelta(days=1)).strftime("%Y-%m-%d")
        prior_from = (prior_to_d - timedelta(days=n_days)).strftime("%Y-%m-%d")
        prev_tot = period_totals_dates(txs, prior_from, prior_to)
        prev_label = f"{prior_from} – {prior_to}"
        label = date_from if date_from == date_to else f"{date_from} – {date_to}"
        exp_rows = by_category_dates(
            txs, date_from, date_to, {"expense", "card_expense"}, cats
        )
        inc_rows = by_category_dates(txs, date_from, date_to, {"income"}, cats)
        multi = date_from != date_to
    else:
        if not months:
            months = [current_month()]
        months = sorted({m.strip() for m in months if m and m.strip()})
        if not months:
            months = [current_month()]
        if not year:
            year = months[0][:4]
        prev_months = prior_period_months(months)
        tot = period_totals(txs, months)
        prev_tot = period_totals(txs, prev_months)
        prev_label = period_label(prev_months) if prev_months else ""
        label = period_label(months)
        exp_rows = by_category_months(txs, months, {"expense", "card_expense"}, cats)
        inc_rows = by_category_months(txs, months, {"income"}, cats)
        multi = len(months) > 1
        date_from = months[0] + "-01"
        date_to = today if months[-1] == current_month() else (months[-1] + "-28")

    balance = sum(parse_decimal(a.get("current_balance")) for a in accs)
    card_limit = sum(parse_decimal(c.get("limit")) for c in cards)
    card_spent = sum(parse_decimal(c.get("current_spent")) for c in cards)
    saved_pct = (tot["balance"] / tot["income"] * 100) if tot["income"] else 0.0
    return {
        "settings": settings,
        "year_month": label,
        "period_months": months,
        "period_year": year,
        "period_multi": multi,
        "date_from": date_from,
        "date_to": date_to,
        "prev_month": prev_label,
        "prev_months": [],
        "totals": tot,
        "prev_totals": prev_tot,
        "balance": balance,
        "card_limit": card_limit,
        "card_spent": card_spent,
        "card_available": card_limit - card_spent,
        "saved_pct": saved_pct,
        "expense_pie": exp_rows,
        "income_pie": inc_rows,
        "top_expenses": exp_rows[:5],
        "goals": goals,
        "budgets": aggregate_budgets(budgets, months),
        "categories": cats,
        "series": monthly_series(txs, months),
        "annual": annual_net(txs, year),
        "notifications": collect_notifications(
            settings, budgets, cats, cards, goals, months
        ),
        "accounts": accs,
        "cards": cards,
    }


def period_totals_dates(txs: list[dict], date_from: str, date_to: str) -> dict:
    income = expense = 0.0
    for t in txs:
        d = str(t.get("date", ""))[:10]
        if not d or d < date_from or d > date_to:
            continue
        amt = parse_decimal(t.get("amount"))
        kind = t.get("type", "")
        if kind == "income":
            income += amt
        elif kind in ("expense", "card_expense"):
            expense += amt
    return {"income": income, "expense": expense, "balance": income - expense}


def by_category_dates(
    txs: list[dict],
    date_from: str,
    date_to: str,
    types: set[str],
    cats: list[dict],
) -> list[tuple[str, float, str]]:
    by_id = cat_index(cats)
    totals: dict[str, float] = defaultdict(float)
    for t in txs:
        d = str(t.get("date", ""))[:10]
        if not d or d < date_from or d > date_to:
            continue
        if t.get("type") not in types:
            continue
        cid = main_category_id(t.get("category_id", ""), by_id)
        totals[cid] += parse_decimal(t.get("amount"))
    rows = []
    for cid, amt in totals.items():
        row = by_id.get(cid)
        name = cat_label(row, cid or "Uncategorized")
        color = (row or {}).get("color", "#7F8C8D")
        rows.append((name, amt, color))
    rows.sort(key=lambda r: r[1], reverse=True)
    return rows


def cockpit_raw(data: dict | None = None) -> dict:
    """Compact payload for client-side period filtering in the dashboard."""
    if data is None:
        data = snapshot()
    txs = read_csv("transactions.csv")
    cats = read_csv("categories.csv")
    budgets = read_csv("budgets.csv")
    today = datetime.now().strftime("%Y-%m-%d")
    month_start = datetime.now().strftime("%Y-%m-01")
    return {
        "currentMonth": current_month(),
        "today": today,
        "monthStart": month_start,
        "dateFrom": data.get("date_from") or month_start,
        "dateTo": data.get("date_to") or today,
        "balance": data["balance"],
        "cardAvailable": data["card_available"],
        "cardLimit": data["card_limit"],
        "cardSpent": data["card_spent"],
        "cards": [
            {
                "id": c.get("id", ""),
                "name": c.get("name", ""),
                "limit": c.get("limit", ""),
                "current_spent": c.get("current_spent", ""),
            }
            for c in data.get("cards") or []
        ],
        "widgets": {
            "ShowBalance": widget_on(data["settings"], "ShowBalance"),
            "ShowPies": widget_on(data["settings"], "ShowPies"),
            "ShowPerformance": widget_on(data["settings"], "ShowPerformance"),
            "ShowGoals": widget_on(data["settings"], "ShowGoals"),
            "ShowBudgets": widget_on(data["settings"], "ShowBudgets"),
            "ShowNotifications": widget_on(data["settings"], "ShowNotifications"),
        },
        "transactions": [
            {
                "date": t.get("date", ""),
                "amount": t.get("amount", ""),
                "type": t.get("type", ""),
                "category_id": t.get("category_id", ""),
            }
            for t in txs
        ],
        "categories": [
            {
                "id": c.get("id", ""),
                "name": c.get("name", ""),
                "parent_id": c.get("parent_id", ""),
                "color": c.get("color", "#7F8C8D"),
                "icon": c.get("icon", ""),
                "type": c.get("type", ""),
            }
            for c in cats
        ],
        "budgets": [
            {
                "year_month": b.get("year_month", ""),
                "category_id": b.get("category_id", ""),
                "planned_amount": b.get("planned_amount", ""),
                "spent_amount": b.get("spent_amount", ""),
            }
            for b in budgets
        ],
        "goals": data["goals"],
    }
