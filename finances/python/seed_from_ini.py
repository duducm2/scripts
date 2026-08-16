"""One-time seed of finance CSVs from repo INI catalogs (mirrors AHK Finance_EnsureData)."""

from __future__ import annotations

import configparser
import csv
import re
import unicodedata
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
DATA = ROOT / "finances" / "data"
PALETTE = [
    "#2ECC71",
    "#3498DB",
    "#9B59B6",
    "#E67E22",
    "#E74C3C",
    "#1ABC9C",
    "#F1C40F",
    "#34495E",
    "#16A085",
    "#27AE60",
    "#2980B9",
    "#8E44AD",
    "#D35400",
    "#C0392B",
    "#7F8C8D",
    "#2C3E50",
]


def slug(name: str) -> str:
    s = unicodedata.normalize("NFKD", name)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r"[^A-Za-z0-9]", "", s).upper()[:8] or "X"
    return s


def unique_id(prefix: str, name: str, existing: set[str]) -> str:
    base = prefix + slug(name)
    cand = base
    n = 2
    while cand in existing:
        cand = f"{base}{n}"
        n += 1
    existing.add(cand)
    return cand


def write_csv(name: str, headers: list[str], rows: list[dict]):
    DATA.mkdir(parents=True, exist_ok=True)
    path = DATA / name
    with path.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=headers, quoting=csv.QUOTE_MINIMAL)
        w.writeheader()
        for r in rows:
            w.writerow({h: r.get(h, "") for h in headers})


def default_cat_icon(name: str) -> str:
    icons = {
        "Adjustment": "⚖️",
        "Food": "🍽️",
        "Beauty": "💄",
        "Hairdresser": "💇",
        "Tattoo, piercing and earrings": "💉",
        "Car": "🚗",
        "Phone": "📱",
        "Shopping": "🛍️",
        "Household bills": "🏠",
        "Education": "📚",
        "Electronics": "💻",
        "Loan": "💳",
        "Humanitarian": "🤝",
        "Taxes": "🧾",
        "Investment income tax": "📉",
        "Games": "🎮",
        "Leisure": "🎉",
        "Bars and clubs": "🍸",
        "Drinks": "🥤",
        "Events": "🎫",
        "Restaurants": "🍴",
        "Travel": "✈️",
        "Misc materials": "🧰",
        "Home": "🏡",
        "Collectibles": "🧸",
        "Adult": "🔒",
        "Groceries": "🛒",
        "Food items": "🥫",
        "Grocery drinks": "🧃",
        "Frozen and deli": "🧊",
        "Personal care": "🧴",
        "Produce": "🥬",
        "Other": "✳️",
        "Bakery": "🥖",
        "Stationery": "✏️",
        "Cleaning products": "🧹",
        "Furniture": "🛋️",
        "Moving": "📦",
        "Banking": "🏦",
        "Pets": "🐾",
        "Clothing": "👕",
        "Costume": "🎭",
        "Health": "❤️",
        "Appointments": "🩺",
        "Medical supplies": "🩹",
        "Medicine": "💊",
        "Services": "🔧",
        "Transport": "🚌",
        "Subscriptions": "📺",
        "Bonus": "🎁",
        "Investments": "📈",
        "São Paulo tax rebate": "🏛️",
        "Prizes": "🏆",
        "Gift": "🎀",
        "Salary adjustment": "🔧",
        "Refund": "↩️",
        "Side income": "💡",
        "Income tax refund": "💰",
        "Salary": "💼",
        "Transfer": "🔄",
        "Sale": "🏷️",
    }
    return icons.get(name, "🏷️")


def seed():
    if (DATA / "categories.csv").exists() and (DATA / "accounts.csv").exists():
        return
    ids: set[str] = set()
    cats: list[dict] = []
    pal = 0
    exp = ROOT / "categories-expenses.ini"
    inc = ROOT / "categories-income.ini"
    if exp.exists():
        cfg = configparser.ConfigParser()
        cfg.optionxform = str
        cfg.read(exp, encoding="utf-8")
        for section in cfg.sections():
            mid = unique_id("CAT_", section, ids)
            cats.append(
                {
                    "id": mid,
                    "name": section,
                    "type": "expense",
                    "parent_id": "",
                    "color": PALETTE[pal % len(PALETTE)],
                    "icon": default_cat_icon(section),
                }
            )
            pal += 1
            for key in cfg.options(section):
                if key == "geral":
                    continue
                sid = unique_id("CAT_", key, ids)
                cats.append(
                    {
                        "id": sid,
                        "name": key,
                        "type": "expense",
                        "parent_id": mid,
                        "color": PALETTE[pal % len(PALETTE)],
                        "icon": default_cat_icon(key),
                    }
                )
                pal += 1
    if inc.exists():
        cfg = configparser.ConfigParser()
        cfg.optionxform = str
        cfg.read(inc, encoding="utf-8")
        for section in cfg.sections():
            mid = unique_id("CAT_", section, ids)
            cats.append(
                {
                    "id": mid,
                    "name": section,
                    "type": "income",
                    "parent_id": "",
                    "color": PALETTE[pal % len(PALETTE)],
                    "icon": default_cat_icon(section),
                }
            )
            pal += 1
    write_csv(
        "categories.csv", ["id", "name", "type", "parent_id", "color", "icon"], cats
    )

    acc_ini = ROOT / "accounts.ini"
    balances = {
        "BoschLife": "23481,62",
        "FGTS": "54104,01",
        "Meal voucher": "0,00",
        "Meli dólares": "0,00",
        "Mercado Pago long-term": "43265,02",
        "Mercado Pago main account": "48004,95",
        "Mercado Pago short-term": "5000,00",
        "Nubank Main": "5000,00",
        "Transition money": "0,00",
    }
    icons = {
        "BoschLife": "🏦",
        "FGTS": "🏛️",
        "Meal voucher": "🍽️",
        "Meli dólares": "💵",
        "Mercado Pago long-term": "📈",
        "Mercado Pago main account": "💳",
        "Mercado Pago short-term": "💰",
        "Nubank Main": "🟣",
        "Transition money": "🔄",
    }
    accs = []
    acc_ids: set[str] = set()
    names = []
    if acc_ini.exists():
        cfg = configparser.ConfigParser()
        cfg.optionxform = str
        cfg.read(acc_ini, encoding="utf-8")
        names = list(cfg.sections())
    for name in names:
        aid = unique_id("ACC_", name, acc_ids)
        bal = balances.get(name, "0,00")
        accs.append(
            {
                "id": aid,
                "name": name,
                "icon": icons.get(name, "🏦"),
                "initial_balance": bal,
                "current_balance": bal,
            }
        )
    write_csv(
        "accounts.csv",
        ["id", "name", "icon", "initial_balance", "current_balance"],
        accs,
    )

    mp = next(
        (a for a in accs if "Mercado Pago main" in a["name"]), accs[0] if accs else None
    )
    write_csv(
        "credit_cards.csv",
        ["id", "name", "limit", "current_spent", "linked_account_id", "closing_day"],
        [
            {
                "id": "CARD_MP",
                "name": "Mercado Pago",
                "limit": "12000,00",
                "current_spent": "2010,22",
                "linked_account_id": mp["id"] if mp else "",
                "closing_day": "9",
            }
        ],
    )

    write_csv(
        "goals.csv",
        ["id", "name", "current_amount", "target_amount", "target_date"],
        [
            {
                "id": "GOAL_PREV",
                "name": "Private pension",
                "current_amount": "16733,13",
                "target_amount": "926400,00",
                "target_date": "2062-01-01",
            },
            {
                "id": "GOAL_EMERG",
                "name": "Emergency fund",
                "current_amount": "2141,14",
                "target_amount": "18000,00",
                "target_date": "2099-11-01",
            },
            {
                "id": "GOAL_FATHER",
                "name": "Father's money",
                "current_amount": "10000,00",
                "target_amount": "10000,00",
                "target_date": "2026-09-15",
            },
            {
                "id": "GOAL_ALEM",
                "name": "Germany",
                "current_amount": "22314,00",
                "target_amount": "50000,00",
                "target_date": "",
            },
            {
                "id": "GOAL_ALEM2",
                "name": "Germany (Leonardo share)",
                "current_amount": "500,00",
                "target_amount": "15000,00",
                "target_date": "",
            },
            {
                "id": "GOAL_CARRO",
                "name": "New car",
                "current_amount": "600,00",
                "target_amount": "40000,00",
                "target_date": "2026-01-01",
            },
            {
                "id": "GOAL_TERR",
                "name": "Land down payment",
                "current_amount": "7218,73",
                "target_amount": "40000,00",
                "target_date": "2026-01-01",
            },
            {
                "id": "GOAL_NCAR",
                "name": "New car",
                "current_amount": "32000,00",
                "target_amount": "70000,00",
                "target_date": "2026-01-01",
            },
        ],
    )

    def cat_id(name: str) -> str:
        for c in cats:
            if c["name"] == name and not c["parent_id"]:
                return c["id"]
        return ""

    def acc_has(needle: str) -> str:
        for a in accs:
            if needle in a["name"]:
                return a["id"]
        return ""

    bud_pairs = [
        ("Groceries", "3000,00", "1039,26"),
        ("Food", "300,00", "603,43"),
        ("Beauty", "150,00", "100,00"),
        ("Leisure", "600,00", "167,28"),
        ("Health", "1000,00", "670,27"),
        ("Car", "400,00", "696,59"),
        ("Household bills", "2300,00", "0,00"),
        ("Education", "400,00", "139,77"),
        ("Electronics", "50,00", "733,44"),
        ("Humanitarian", "150,00", "350,90"),
    ]
    write_csv(
        "budgets.csv",
        ["year_month", "category_id", "planned_amount", "spent_amount"],
        [
            {
                "year_month": "2026-08",
                "category_id": cat_id(n),
                "planned_amount": p,
                "spent_amount": s,
            }
            for n, p, s in bud_pairs
            if cat_id(n)
        ],
    )

    acc_bl = acc_has("BoschLife")
    acc_mp = acc_has("Mercado Pago main") or acc_has("Mercado Pago")
    txs = [
        (
            "TX001",
            "2026-08-16",
            "Banana",
            "3,00",
            "expense",
            cat_id("Groceries"),
            "Produce",
            acc_bl,
            "",
            "",
        ),
        (
            "TX002",
            "2026-08-16",
            "Gift",
            "10,00",
            "income",
            cat_id("Bonus"),
            "",
            acc_bl,
            "",
            "",
        ),
        (
            "TX003",
            "2026-08-15",
            "Lunch",
            "31,24",
            "expense",
            cat_id("Food"),
            "",
            acc_mp,
            "",
            "",
        ),
        (
            "TX004",
            "2026-08-15",
            "Groceries",
            "317,04",
            "expense",
            cat_id("Groceries"),
            "",
            acc_mp,
            "",
            "",
        ),
        (
            "TX005",
            "2026-08-15",
            "Mic",
            "529,99",
            "card_expense",
            cat_id("Electronics"),
            "",
            acc_mp,
            "CARD_MP",
            "",
        ),
        (
            "TX006",
            "2026-08-03",
            "Salary and PLR",
            "4876,76",
            "income",
            cat_id("Salary"),
            "",
            acc_mp,
            "",
            "",
        ),
    ]
    headers = [
        "id",
        "date",
        "description",
        "amount",
        "type",
        "category_id",
        "subcategory",
        "account_id",
        "card_id",
        "transfer_account_id",
    ]
    write_csv("transactions.csv", headers, [dict(zip(headers, t)) for t in txs])

    settings = DATA / "settings.ini"
    if not settings.exists():
        settings.write_text(
            "[Dashboard]\nShowBalance=1\nShowPies=1\nShowPerformance=1\nShowGoals=1\nShowBudgets=1\nShowNotifications=1\n\n"
            "[General]\nDefaultAccountId="
            + acc_mp
            + "\nPrimaryCardId=CARD_MP\nNotifyBudgetExceeded=1\nNotifyCardHighUsage=1\nCardUsageWarnPct=80\n",
            encoding="utf-8",
        )


if __name__ == "__main__":
    seed()
    print("seeded", DATA)
