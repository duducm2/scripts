"""Fictional Jan–Aug 2026 transactions + budgets. Does not touch accounts or cards."""

from __future__ import annotations

import csv
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
DATA = ROOT / "finances" / "data"

TX_HEADERS = [
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
BUD_HEADERS = ["year_month", "category_id", "planned_amount", "spent_amount"]

PLANNED = [
    ("CAT_MERCADO", "3000,00"),
    ("CAT_ALIMENTA", "300,00"),
    ("CAT_BELEZA", "150,00"),
    ("CAT_LAZER", "600,00"),
    ("CAT_SAUDE", "1000,00"),
    ("CAT_CARRO", "400,00"),
    ("CAT_CONTASDE", "2300,00"),
    ("CAT_EDUCACAO", "400,00"),
    ("CAT_ELETRONI", "50,00"),
    ("CAT_HUMANITA", "150,00"),
]

MP = "ACC_MERCADOP"
MP_ST = "ACC_MERCADOP2"
MP_LT = "ACC_MERCADOP3"
NU = "ACC_NUBANKMA"
BL = "ACC_BOSCHLIF"
CARD = "CARD_MP"


def brl(n: float) -> str:
    return f"{n:.2f}".replace(".", ",")


def write_csv(name: str, headers: list[str], rows: list[dict]) -> None:
    path = DATA / name
    with path.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=headers, quoting=csv.QUOTE_MINIMAL)
        w.writeheader()
        for r in rows:
            w.writerow({h: r.get(h, "") for h in headers})


def tx(
    n: int,
    date: str,
    desc: str,
    amount: float,
    kind: str,
    cat: str,
    acc: str,
    sub: str = "",
    card: str = "",
    dest: str = "",
) -> dict:
    return {
        "id": f"TX{n:03d}",
        "date": date,
        "description": desc,
        "amount": brl(amount),
        "type": kind,
        "category_id": cat,
        "subcategory": sub,
        "account_id": acc,
        "card_id": card,
        "transfer_account_id": dest,
    }


def build_transactions() -> list[dict]:
    rows: list[dict] = []
    n = 1

    def add(*args, **kwargs):
        nonlocal n
        rows.append(tx(n, *args, **kwargs))
        n += 1

    salary = [4200.00, 4200.00, 4200.00, 4300.00, 4500.00, 4800.00, 4600.00, 4876.76]
    groceries = [890.40, 920.15, 1040.80, 980.00, 1102.33, 1250.40, 980.12, 1039.26]
    lunch = [210.50, 180.00, 245.90, 198.40, 260.10, 310.00, 175.20, 280.67]
    leisure = [120.00, 340.00, 80.00, 450.00, 167.00, 520.00, 90.00, 167.28]
    health = [80.00, 220.00, 450.00, 90.00, 670.00, 150.00, 310.00, 442.81]
    car = [180.00, 250.00, 400.00, 190.00, 320.00, 410.00, 200.00, 280.00]
    house = [2100.00, 2100.00, 2150.00, 2150.00, 2200.00, 2200.00, 2250.00, 2300.00]
    edu = [139.90, 139.90, 139.90, 200.00, 139.90, 400.00, 139.90, 139.77]
    human = [50.00, 80.00, 50.00, 200.00, 80.00, 350.00, 50.00, 150.00]
    beauty = [0.00, 150.00, 0.00, 80.00, 0.00, 120.00, 0.00, 100.00]
    # Past months: paid invoices (not in current_spent). August sums to 2010.22.
    card_past = [410.00, 890.50, 120.00, 650.00, 210.00, 1480.00, 330.00]
    extras = [
        (1, 18, "Pharmacy", 47.90, "expense", "CAT_SAUDE", MP, "Remédios"),
        (2, 8, "Uber", 36.40, "expense", "CAT_CARRO", NU, ""),
        (2, 22, "Cinema", 64.00, "expense", "CAT_LAZER", MP, "Eventos"),
        (3, 12, "Haircut", 80.00, "expense", "CAT_BELEZA", MP, "Cabeleireiro"),
        (3, 28, "Course installment", 200.00, "expense", "CAT_EDUCACAO", MP, ""),
        (4, 5, "Dog food", 129.90, "expense", "CAT_CACHORRO", MP, ""),
        (4, 19, "Clothes", 189.00, "expense", "CAT_ROUPA", MP, ""),
        (5, 7, "NF Paulista credit", 42.18, "income", "CAT_NOTAFISC", MP, ""),
        (5, 21, "Donation", 80.00, "expense", "CAT_HUMANITA", MP, ""),
        (6, 4, "PLR", 8500.00, "income", "CAT_BONIFICA", MP, ""),
        (6, 11, "Investment yield", 1523.45, "income", "CAT_INVESTIM", MP_LT, ""),
        (6, 18, "Headphones", 733.44, "card_expense", "CAT_ELETRONI", MP, "", CARD),
        (7, 9, "Refund", 120.00, "income", "CAT_REEMBOLS", MP, ""),
        (7, 26, "Restaurant", 98.70, "expense", "CAT_LAZER", MP, "Restaurantes"),
        (8, 5, "Side gig", 650.00, "income", "CAT_RENDAEXT", NU, ""),
        (8, 12, "Father's support", 10000.00, "income", "CAT_OUTROS2", MP, ""),
        (8, 16, "Banana", 3.00, "expense", "CAT_MERCADO", BL, "Hortifruti"),
        (8, 16, "Gift", 10.00, "income", "CAT_BONIFICA", BL, ""),
    ]

    for m in range(1, 9):
        ym = f"2026-{m:02d}"
        add(f"{ym}-03", "Salary", salary[m - 1], "income", "CAT_SALARIO", MP)
        add(
            f"{ym}-08",
            "Groceries",
            groceries[m - 1],
            "expense",
            "CAT_MERCADO",
            MP,
            "Alimentos",
        )
        add(f"{ym}-14", "Lunch out", lunch[m - 1], "expense", "CAT_ALIMENTA", MP)
        add(f"{ym}-10", "Household bills", house[m - 1], "expense", "CAT_CONTASDE", MP)
        add(f"{ym}-06", "Fuel", car[m - 1], "expense", "CAT_CARRO", NU)
        if leisure[m - 1]:
            add(f"{ym}-20", "Leisure", leisure[m - 1], "expense", "CAT_LAZER", MP)
        if health[m - 1]:
            add(
                f"{ym}-17",
                "Health",
                health[m - 1],
                "expense",
                "CAT_SAUDE",
                MP,
                "Consultas",
            )
        add(f"{ym}-11", "Course", edu[m - 1], "expense", "CAT_EDUCACAO", MP)
        add(f"{ym}-15", "Donation", human[m - 1], "expense", "CAT_HUMANITA", MP)
        if beauty[m - 1]:
            add(f"{ym}-22", "Beauty", beauty[m - 1], "expense", "CAT_BELEZA", MP)
        if m < 8:
            add(
                f"{ym}-25",
                "Card purchase",
                card_past[m - 1],
                "card_expense",
                "CAT_ELETRONI",
                MP,
                "",
                CARD,
            )
        if m in (2, 5, 8):
            add(
                f"{ym}-27",
                "Transfer to short-term",
                500.00,
                "transfer",
                "CAT_TRANSFER",
                MP,
                dest=MP_ST,
            )

    # August open invoice: 529.99 + 680.23 + 800.00 = 2010.22
    add("2026-08-15", "Mic", 529.99, "card_expense", "CAT_ELETRONI", MP, "", CARD)
    add("2026-08-07", "Keyboard", 680.23, "card_expense", "CAT_ELETRONI", MP, "", CARD)
    add(
        "2026-08-21",
        "Monitor stand",
        800.00,
        "card_expense",
        "CAT_ELETRONI",
        MP,
        "",
        CARD,
    )
    add("2026-08-15", "Groceries extra", 317.04, "expense", "CAT_MERCADO", MP)
    add("2026-08-15", "Lunch", 31.24, "expense", "CAT_ALIMENTA", MP)

    for item in extras:
        m, d, desc, amt, kind, cat, acc, *rest = item
        sub = rest[0] if rest else ""
        card = rest[1] if len(rest) > 1 else ""
        add(f"2026-{m:02d}-{d:02d}", desc, amt, kind, cat, acc, sub, card)

    rows.sort(key=lambda r: (r["date"], r["id"]))
    for i, r in enumerate(rows, 1):
        r["id"] = f"TX{i:03d}"
    return rows


def main_cat(cat_id: str, by_parent: dict[str, str]) -> str:
    return by_parent.get(cat_id, cat_id)


def load_parent_map() -> dict[str, str]:
    path = DATA / "categories.csv"
    by_parent: dict[str, str] = {}
    if not path.exists():
        return by_parent
    with path.open(encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            pid = (row.get("parent_id") or "").strip()
            cid = row.get("id", "")
            by_parent[cid] = pid or cid
    return by_parent


def parse_amt(s: str) -> float:
    return float(str(s).replace(".", "").replace(",", ".")) if s else 0.0


def build_budgets(txs: list[dict]) -> list[dict]:
    parents = load_parent_map()
    spent: dict[tuple[str, str], float] = defaultdict(float)
    for t in txs:
        if t["type"] not in ("expense", "card_expense"):
            continue
        ym = t["date"][:7]
        cid = main_cat(t["category_id"], parents)
        spent[(ym, cid)] += parse_amt(t["amount"])
    rows = []
    for m in range(1, 9):
        ym = f"2026-{m:02d}"
        for cid, planned in PLANNED:
            rows.append(
                {
                    "year_month": ym,
                    "category_id": cid,
                    "planned_amount": planned,
                    "spent_amount": brl(spent.get((ym, cid), 0.0)),
                }
            )
    return rows


def main() -> None:
    txs = build_transactions()
    write_csv("transactions.csv", TX_HEADERS, txs)
    write_csv("budgets.csv", BUD_HEADERS, build_budgets(txs))
    print(f"wrote {len(txs)} transactions and {8 * len(PLANNED)} budget rows")


if __name__ == "__main__":
    main()
