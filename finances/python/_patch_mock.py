from pathlib import Path

p = Path(__file__).with_name("seed_mock_history.py")
text = p.read_text(encoding="utf-8")
start = text.index("    extras = [")
end = text.index('    rows.sort(key=lambda r: (r["date"], r["id"]))')
new = """    extras = [
        (1, 18, "Pharmacy", 47.90, "expense", "CAT_SAUDE", MP),
        (2, 8, "Uber", 36.40, "expense", "CAT_CARRO", NU),
        (2, 22, "Cinema", 64.00, "expense", "CAT_LAZER", MP),
        (3, 12, "Haircut", 80.00, "expense", "CAT_BELEZA", MP),
        (3, 28, "Course installment", 200.00, "expense", "CAT_EDUCACAO", MP),
        (4, 5, "Dog food", 129.90, "expense", "CAT_CACHORRO", MP),
        (4, 19, "Clothes", 189.00, "expense", "CAT_ROUPA", MP),
        (5, 7, "NF Paulista credit", 42.18, "income", "CAT_NOTAFISC", MP),
        (5, 21, "Donation", 80.00, "expense", "CAT_HUMANITA", MP),
        (6, 4, "PLR", 8500.00, "income", "CAT_BONIFICA", MP),
        (6, 11, "Investment yield", 1523.45, "income", "CAT_INVESTIM", MP_LT),
        (6, 18, "Headphones", 733.44, "card_expense", "CAT_ELETRONI", MP, CARD),
        (7, 9, "Refund", 120.00, "income", "CAT_REEMBOLS", MP),
        (7, 26, "Restaurant", 98.70, "expense", "CAT_LAZER", MP),
        (8, 5, "Side gig", 650.00, "income", "CAT_RENDAEXT", NU),
        (8, 12, "Father's support", 10000.00, "income", "CAT_OUTROS2", MP),
        (8, 16, "Banana", 3.00, "expense", "CAT_MERCADO", BL),
        (8, 16, "Gift", 10.00, "income", "CAT_BONIFICA", BL),
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
    add("2026-08-15", "Mic", 529.99, "card_expense", "CAT_ELETRONI", MP, CARD)
    add("2026-08-07", "Keyboard", 680.23, "card_expense", "CAT_ELETRONI", MP, CARD)
    add(
        "2026-08-21",
        "Monitor stand",
        800.00,
        "card_expense",
        "CAT_ELETRONI",
        MP,
        CARD,
    )
    add("2026-08-15", "Groceries extra", 317.04, "expense", "CAT_MERCADO", MP)
    add("2026-08-15", "Lunch", 31.24, "expense", "CAT_ALIMENTA", MP)

    for item in extras:
        m, d, desc, amt, kind, cat, acc, *rest = item
        card = rest[0] if rest else ""
        add(f"2026-{m:02d}-{d:02d}", desc, amt, kind, cat, acc, card)

"""
p.write_text(text[:start] + new + text[end:], encoding="utf-8")
print("patched")
