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


def configure_paths(
    data_dir: str | Path | None = None, output_dir: str | Path | None = None
) -> None:
    """Point DATA/OUTPUT at AHK's finance folders (absolute paths)."""
    global DATA, OUTPUT, ROOT
    if data_dir:
        DATA = Path(data_dir).expanduser().resolve()
        # finances/data -> repo root
        ROOT = DATA.parent.parent
    if output_dir:
        OUTPUT = Path(output_dir).expanduser().resolve()
    # Keep seed_from_ini in sync when it is loaded.
    try:
        import seed_from_ini as seed_mod

        seed_mod.DATA = DATA
        seed_mod.ROOT = ROOT
    except ImportError:
        pass


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


def tx_is_paid(t: dict) -> bool:
    p = str(t.get("paid") or "0").strip().lower()
    return p in ("1", "true", "yes")


def days_in_month(year: int, month: int) -> int:
    if month == 2:
        leap = year % 4 == 0 and (year % 100 != 0 or year % 400 == 0)
        return 29 if leap else 28
    if month in (4, 6, 9, 11):
        return 30
    return 31


def clamp_closing_day(year: int, month: int, closing_day: int) -> int:
    cd = int(closing_day or 1)
    if cd < 1:
        cd = 1
    if cd > 31:
        cd = 31
    return min(cd, days_in_month(year, month))


def closing_date_on(year: int, month: int, closing_day: int) -> str:
    cd = clamp_closing_day(year, month, closing_day)
    return f"{year:04d}-{month:02d}-{cd:02d}"


def bill_closing_date(tx_date: str, closing_day: int) -> str:
    """Statement close date that includes this purchase/parcel date.

    Purchases on or before closing_day land on that month's bill; later days
    roll to the next month's closing.
    """
    d = str(tx_date or "")[:10]
    if len(d) < 10:
        d = datetime.now().strftime("%Y-%m-%d")
    y, m, day = [int(x) for x in d.split("-")]
    cd = clamp_closing_day(y, m, closing_day)
    if day <= cd:
        return closing_date_on(y, m, closing_day)
    m += 1
    if m > 12:
        m = 1
        y += 1
    return closing_date_on(y, m, closing_day)


def next_closing_on_or_after(today: str, closing_day: int) -> str:
    d = str(today or "")[:10]
    if len(d) < 10:
        d = datetime.now().strftime("%Y-%m-%d")
    y, m, day = [int(x) for x in d.split("-")]
    cd = clamp_closing_day(y, m, closing_day)
    if day <= cd:
        return closing_date_on(y, m, closing_day)
    m += 1
    if m > 12:
        m = 1
        y += 1
    return closing_date_on(y, m, closing_day)


def previous_closing_on_or_before(today: str, closing_day: int) -> str:
    d = str(today or "")[:10]
    if len(d) < 10:
        d = datetime.now().strftime("%Y-%m-%d")
    y, m, day = [int(x) for x in d.split("-")]
    cd = clamp_closing_day(y, m, closing_day)
    if day >= cd:
        return closing_date_on(y, m, closing_day)
    m -= 1
    if m < 1:
        m = 12
        y -= 1
    return closing_date_on(y, m, closing_day)


def due_date_for_closing(closing_date: str, due_day: int) -> str:
    """Payment due date for the statement that closed on closing_date."""
    d = str(closing_date or "")[:10]
    if len(d) < 10:
        d = datetime.now().strftime("%Y-%m-%d")
    y, m, close_dom = [int(x) for x in d.split("-")]
    dd = int(due_day or 1)
    if dd < 1:
        dd = 1
    if dd > 31:
        dd = 31
    if dd > close_dom:
        return closing_date_on(y, m, dd)
    m += 1
    if m > 12:
        m = 1
        y += 1
    return closing_date_on(y, m, dd)


def next_pay_date(today: str, closing_day: int, due_day: int) -> str:
    """Next payment date on or after today for this card."""
    prev_close = previous_closing_on_or_before(today, closing_day)
    due = due_date_for_closing(prev_close, due_day)
    if due >= today[:10]:
        return due
    nxt_close = next_closing_on_or_after(today, closing_day)
    return due_date_for_closing(nxt_close, due_day)


def card_due_day(card: dict, closing_day: int) -> int:
    raw = str(card.get("due_day") or "").strip()
    if raw:
        try:
            dd = int(raw)
            if 1 <= dd <= 31:
                return dd
        except ValueError:
            pass
    dd = int(closing_day or 1) + 7
    if dd > 31:
        dd -= 31
    return max(1, min(31, dd))


def card_installment_remaining(
    txs: list[dict], cards: list[dict], today: str | None = None
) -> dict:
    """Per-card bill amounts on each closing day + open/later summary.

    Summary total = credit_cards.current_spent.
    Each unpaid card_expense is assigned to a statement closing date via
    closing_day. Open = dues on the next closing (and any overdue). Later =
    dues on later closings. Chart bars = amount due on each closing day.
    """
    if not today:
        today = datetime.now().strftime("%Y-%m-%d")
    closing_by_card = {}
    due_by_card = {}
    card_by_id = {c.get("id"): c for c in cards if c.get("id")}
    for c in cards:
        cid = c.get("id")
        if not cid:
            continue
        try:
            closing_by_card[cid] = int(str(c.get("closing_day") or "1").strip() or "1")
        except ValueError:
            closing_by_card[cid] = 1
        due_by_card[cid] = card_due_day(c, closing_by_card[cid])
    unpaid: list[dict] = []
    for t in txs:
        if t.get("type") != "card_expense":
            continue
        if tx_is_paid(t):
            continue
        cid = (t.get("card_id") or "").strip()
        d = str(t.get("date") or "")[:10]
        if not cid or len(d) < 10:
            continue
        cd = closing_by_card.get(cid, 1)
        unpaid.append(
            {
                "card_id": cid,
                "date": d,
                "closing": bill_closing_date(d, cd),
                "amount": parse_decimal(t.get("amount")),
                "installments": str(t.get("installments") or "1"),
                "installment_n": str(t.get("installment_n") or "1"),
            }
        )
    spent_by_card = {
        c.get("id"): parse_decimal(c.get("current_spent")) for c in cards if c.get("id")
    }
    card_meta = {
        c.get("id"): (c.get("name") or c.get("id") or "Card")
        for c in cards
        if c.get("id")
    }
    empty = {
        "months": [],
        "closings": [],
        "series": [],
        "today": today,
        "from_today": [],
        "from_today_total": 0.0,
        "open_total": 0.0,
        "later_total": 0.0,
        "last_date": "",
    }
    palette = ["#e74c3c", "#3498db", "#9b59b6", "#f39c12", "#1abc9c", "#e67e22"]

    active_ids = [
        cid
        for cid, spent in spent_by_card.items()
        if spent > 0.00001 or any(u["card_id"] == cid for u in unpaid)
    ]
    if not active_ids:
        return empty

    from_today_rows = []
    series = []
    last_date = ""
    for i, cid in enumerate(active_ids):
        total_amt = round(spent_by_card.get(cid, 0.0), 2)
        cd = closing_by_card.get(cid, 1)
        open_close = next_closing_on_or_after(today, cd)
        card_unpaid = [u for u in unpaid if u["card_id"] == cid]
        due_by_close: dict[str, float] = defaultdict(float)
        for u in card_unpaid:
            due_by_close[u["closing"]] += u["amount"]
        # Include the next closing even if empty so the open bill is visible.
        if open_close not in due_by_close:
            due_by_close[open_close] = 0.0
        closings = sorted(due_by_close.keys())
        raw_open = round(
            sum(amt for close, amt in due_by_close.items() if close <= open_close),
            2,
        )
        raw_later = round(
            sum(amt for close, amt in due_by_close.items() if close > open_close),
            2,
        )
        raw_sum = raw_open + raw_later
        # Align open/later split to current_spent when ledger and spent differ.
        if raw_sum > 0.00001 and abs(raw_sum - total_amt) > 0.02:
            scale = total_amt / raw_sum
            open_amt = round(raw_open * scale, 2)
            later_amt = round(max(0.0, total_amt - open_amt), 2)
            due_vals = {c: round(due_by_close[c] * scale, 2) for c in closings}
        else:
            open_amt = min(raw_open, total_amt) if total_amt else raw_open
            later_amt = round(max(0.0, total_amt - open_amt), 2)
            due_vals = {c: round(due_by_close[c], 2) for c in closings}
            if total_amt > 0 and raw_sum <= 0.00001:
                # Spent with no dated parcels: put it all on the next closing.
                due_vals = {open_close: total_amt}
                closings = [open_close]
                open_amt = total_amt
                later_amt = 0.0

        card_last = max(
            (c for c in closings if due_vals.get(c, 0) > 0.00001), default=""
        )
        if card_last and (not last_date or card_last > last_date):
            last_date = card_last
        color = palette[i % len(palette)]
        dd = due_by_card.get(cid, card_due_day(card_by_id.get(cid, {}), cd))
        next_pay = next_pay_date(today, cd, dd)
        from_today_rows.append(
            {
                "card_id": cid,
                "name": card_meta.get(cid, cid),
                "color": color,
                "amount": total_amt,
                "open": open_amt,
                "later": later_amt,
                "last_date": card_last,
                "next_closing": open_close,
                "next_pay": next_pay,
                "closing_day": cd,
                "due_day": dd,
            }
        )
        series.append(
            {
                "card_id": cid,
                "name": card_meta.get(cid, cid),
                "color": color,
                "dates": closings,
                "values": [due_vals[c] for c in closings],
            }
        )

    # Unified closing axis (all cards) for optional shared-x charts.
    all_closings = sorted({d for s in series for d in (s.get("dates") or [])})

    return {
        "months": all_closings,
        "closings": all_closings,
        "series": series,
        "today": today,
        "from_today": from_today_rows,
        "from_today_total": round(sum(r["amount"] for r in from_today_rows), 2),
        "open_total": round(sum(r["open"] for r in from_today_rows), 2),
        "later_total": round(sum(r["later"] for r in from_today_rows), 2),
        "last_date": last_date,
    }


def setting_get(section: dict, key: str, default: str = "") -> str:
    """ConfigParser lowercases option names; accept either casing."""
    if not section:
        return default
    if key in section:
        return section[key]
    lower = key.lower()
    for k, v in section.items():
        if str(k).lower() == lower:
            return v
    return default


def liquid_after_card(settings: dict, accs: list[dict], cards: list[dict]) -> dict:
    """Checking balances minus all cards' current spent (None if unresolved).

    Checking accounts = DefaultAccountId plus every card's linked_account_id.
    Available money uses every card's limit/spent, not only the primary card.
    """
    gen = settings.get("general") or {}
    default_acc_id = (setting_get(gen, "DefaultAccountId") or "").strip()
    empty = {
        "liquid_after_card": None,
        "liquid_account_bal": None,
        "liquid_card_spent": None,
        "liquid_card_limit": None,
        "liquid_account_name": "",
        "liquid_card_name": "",
        "liquid_accounts": [],
        "liquid_cards": [],
    }
    by_id = {a.get("id"): a for a in accs if a.get("id")}
    checking_ids: set[str] = set()
    if default_acc_id:
        checking_ids.add(default_acc_id)
    for c in cards:
        linked = (c.get("linked_account_id") or "").strip()
        if linked:
            checking_ids.add(linked)
    if not checking_ids:
        for a in accs:
            name = (a.get("name") or "").lower()
            aid = a.get("id") or ""
            if aid and "checking" in name:
                checking_ids.add(aid)
    account_rows: list[dict] = []
    total_bal = 0.0
    for aid in sorted(checking_ids):
        acc = by_id.get(aid)
        if not acc:
            continue
        bal = parse_decimal(acc.get("current_balance"))
        total_bal += bal
        account_rows.append({"id": aid, "name": acc.get("name") or aid, "balance": bal})
    card_rows: list[dict] = []
    total_spent = 0.0
    total_limit = 0.0
    for c in cards:
        spent = parse_decimal(c.get("current_spent"))
        limit = parse_decimal(c.get("limit"))
        total_spent += spent
        total_limit += limit
        cid = c.get("id") or ""
        card_rows.append(
            {
                "id": cid,
                "name": c.get("name") or cid or "Card",
                "spent": spent,
                "limit": limit,
            }
        )
    if not account_rows or not card_rows:
        return empty
    names = [r["name"] for r in account_rows]
    return {
        "liquid_after_card": total_bal - total_spent,
        "liquid_account_bal": total_bal,
        "liquid_card_spent": total_spent,
        "liquid_card_limit": total_limit,
        "liquid_account_name": " + ".join(names),
        "liquid_card_name": "all cards",
        "liquid_accounts": account_rows,
        "liquid_cards": card_rows,
    }


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
    recurring = read_csv("recurring_bills.csv")
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
    liquid = liquid_after_card(settings, accs, cards)
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
        "recurring": recurring,
        "budgets": aggregate_budgets(budgets, months),
        "categories": cats,
        "series": monthly_series(txs, months),
        "annual": annual_net(txs, year),
        "notifications": collect_notifications(
            settings, budgets, cats, cards, goals, months
        ),
        "cardInstallmentRemaining": card_installment_remaining(txs, cards),
        "accounts": accs,
        "cards": cards,
        **liquid,
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
        "liquidAfterCard": data.get("liquid_after_card"),
        "liquidAccountBal": data.get("liquid_account_bal"),
        "liquidCardSpent": data.get("liquid_card_spent"),
        "liquidCardLimit": data.get("liquid_card_limit"),
        "liquidAccountName": data.get("liquid_account_name") or "",
        "liquidCardName": data.get("liquid_card_name") or "",
        "liquidAccounts": data.get("liquid_accounts") or [],
        "liquidCards": data.get("liquid_cards") or [],
        "cards": [
            {
                "id": c.get("id", ""),
                "name": c.get("name", ""),
                "limit": c.get("limit", ""),
                "current_spent": c.get("current_spent", ""),
                "closing_day": c.get("closing_day", "1"),
                "due_day": c.get("due_day", ""),
            }
            for c in data.get("cards") or []
        ],
        "accounts": [
            {
                "id": a.get("id", ""),
                "name": a.get("name", ""),
                "icon": a.get("icon", ""),
                "current_balance": a.get("current_balance", ""),
            }
            for a in data.get("accounts") or []
        ],
        "widgets": {
            "ShowBalance": widget_on(data["settings"], "ShowBalance"),
            "ShowAccounts": widget_on(data["settings"], "ShowAccounts"),
            "ShowPies": widget_on(data["settings"], "ShowPies"),
            "ShowPerformance": widget_on(data["settings"], "ShowPerformance"),
            "ShowGoals": widget_on(data["settings"], "ShowGoals"),
            "ShowBudgets": widget_on(data["settings"], "ShowBudgets"),
            "ShowRecurring": widget_on(data["settings"], "ShowRecurring"),
            "ShowNotifications": widget_on(data["settings"], "ShowNotifications"),
            "ShowCardPlan": widget_on(data["settings"], "ShowCardPlan"),
        },
        "cardInstallmentRemaining": data.get("cardInstallmentRemaining")
        or {"months": [], "series": []},
        "transactions": [
            {
                "date": t.get("date", ""),
                "description": t.get("description", ""),
                "amount": t.get("amount", ""),
                "type": t.get("type", ""),
                "category_id": t.get("category_id", ""),
                "account_id": t.get("account_id", ""),
                "card_id": t.get("card_id", ""),
                "paid": t.get("paid", "0"),
                "installments": t.get("installments", "1"),
                "installment_n": t.get("installment_n", "1"),
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
        "recurring": [
            {
                "id": r.get("id", ""),
                "name": r.get("name", ""),
                "icon": r.get("icon", ""),
                "monthly_amount": r.get("monthly_amount", ""),
            }
            for r in data.get("recurring") or []
        ],
    }
