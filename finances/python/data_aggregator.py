"""Read finance CSVs with Brazilian comma decimals."""

from __future__ import annotations

import csv
import configparser
from collections import defaultdict
from datetime import datetime
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


def by_category(
    txs: list[dict], ym: str, types: set[str], cats: list[dict]
) -> list[tuple[str, float, str]]:
    by_id = cat_index(cats)
    totals: dict[str, float] = defaultdict(float)
    for t in txs:
        if not str(t.get("date", "")).startswith(ym):
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


def monthly_series(txs: list[dict]) -> list[dict]:
    months = sorted({str(t.get("date", ""))[:7] for t in txs if t.get("date")})
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


def collect_notifications(
    settings: dict,
    budgets: list[dict],
    cats: list[dict],
    cards: list[dict],
    goals: list[dict],
    ym: str,
) -> list[str]:
    notes = []
    by_id = cat_index(cats)
    if settings.get("general", {}).get("NotifyBudgetExceeded", "1") != "0":
        for b in budgets:
            if b.get("year_month") != ym:
                continue
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


def snapshot() -> dict:
    txs = read_csv("transactions.csv")
    accs = read_csv("accounts.csv")
    cats = read_csv("categories.csv")
    cards = read_csv("credit_cards.csv")
    goals = read_csv("goals.csv")
    budgets = read_csv("budgets.csv")
    settings = read_settings()
    ym = current_month()
    prev = month_shift(ym, -1)
    tot = month_totals(txs, ym)
    prev_tot = month_totals(txs, prev)
    balance = sum(parse_decimal(a.get("current_balance")) for a in accs)
    card_limit = sum(parse_decimal(c.get("limit")) for c in cards)
    card_spent = sum(parse_decimal(c.get("current_spent")) for c in cards)
    saved_pct = (tot["balance"] / tot["income"] * 100) if tot["income"] else 0.0
    exp_rows = by_category(txs, ym, {"expense", "card_expense"}, cats)
    inc_rows = by_category(txs, ym, {"income"}, cats)
    return {
        "settings": settings,
        "year_month": ym,
        "prev_month": prev,
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
        "budgets": [b for b in budgets if b.get("year_month") == ym],
        "categories": cats,
        "series": monthly_series(txs),
        "annual": annual_net(txs, ym[:4]),
        "notifications": collect_notifications(
            settings, budgets, cats, cards, goals, ym
        ),
        "accounts": accs,
        "cards": cards,
    }
