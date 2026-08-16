"""Name maps from Mobills (PT / truncated) to local finance CSV ids."""

from __future__ import annotations

# Exact Mobills category Name: "…" → (category_id, subcategory_name_or_empty)
CATEGORY_EXACT: dict[str, tuple[str, str]] = {
    "Ajuste": ("CAT_AJUSTE", ""),
    "Alimentação": ("CAT_ALIMENTA", ""),
    "Beleza": ("CAT_BELEZA", ""),
    "Bonificação": ("CAT_BONIFICA", ""),
    "Cachorro": ("CAT_PETS", ""),
    "Carro": ("CAT_CARRO", ""),
    "Celular": ("CAT_CELULAR", ""),
    "Educação": ("CAT_EDUCACAO", ""),
    "Eletrônicos": ("CAT_ELETRONI", ""),
    "Hortifruti": ("CAT_HORTIFRU", ""),
    "Humanitário": ("CAT_HUMANITA", ""),
    "Materiais diversos": ("CAT_MATERIAI", ""),
    "Mercado": ("CAT_MERCADO", ""),
    "Outros": ("CAT_OUTROS", ""),
    "Salário": ("CAT_SALARIO", ""),
    "Saúde": ("CAT_SAUDE", ""),
    "Transporte": ("CAT_TRANSPOR", ""),
    "TR Transfer*": ("CAT_TRANSFER", ""),
    "Pagamento Recebido*": ("CAT_REEMBOLS", ""),
    "Reajuste*": ("CAT_REAJUSTE", ""),
    "Reajuste\\*": ("CAT_REAJUSTE", ""),
    # Truncated / odd spellings seen in dumps
    "Crédito_": ("CAT_OPERACAO", ""),
}

# Prefix → category_id (longest match wins); subcategory empty unless noted
CATEGORY_PREFIX: list[tuple[str, str, str]] = [
    ("TR Transfer", "CAT_TRANSFER", ""),
    ("Pagamento Recebido", "CAT_REEMBOLS", ""),
    ("Reajuste", "CAT_REAJUSTE", ""),
    ("Bonifica", "CAT_BONIFICA", ""),
    ("Hortifruti", "CAT_HORTIFRU", ""),
    ("Alimenta", "CAT_ALIMENTA", ""),
    ("Eletr", "CAT_ELETRONI", ""),
    ("Educa", "CAT_EDUCACAO", ""),
    ("Humanit", "CAT_HUMANITA", ""),
    ("Materiais", "CAT_MATERIAI", ""),
    ("Mercado", "CAT_MERCADO", ""),
    ("Saúde", "CAT_SAUDE", ""),
    ("Saude", "CAT_SAUDE", ""),
    ("Salário", "CAT_SALARIO", ""),
    ("Salario", "CAT_SALARIO", ""),
    ("Transporte", "CAT_TRANSPOR", ""),
    ("Cachorro", "CAT_PETS", ""),
    ("Beleza", "CAT_BELEZA", ""),
    ("Carro", "CAT_CARRO", ""),
    ("Celular", "CAT_CELULAR", ""),
    ("Ajuste", "CAT_AJUSTE", ""),
]

# Account Name: "…" (often truncated with …) → account_id
ACCOUNT_EXACT: dict[str, str] = {
    "BoschLife": "ACC_BOSCHLIF",
    "Nubank Main": "ACC_NUBANKMA",
    "Meli dólares": "ACC_MELIDOLA",
    "Meli dolares": "ACC_MELIDOLA",
    "FGTS": "ACC_FGTS",
    "Meal voucher": "ACC_MEALVOUC",
    "Transition money": "ACC_TRANSITI",
}

# Prefix of Mobills label → account_id (order: more specific first)
ACCOUNT_PREFIX: list[tuple[str, str]] = [
    ("Mercado Pago lo", "ACC_MERCADOP3"),  # long-term
    ("Mercado pago lo", "ACC_MERCADOP3"),
    ("Mercado Pago sh", "ACC_MERCADOP2"),  # short-term
    ("Mercado pago sh", "ACC_MERCADOP2"),
    ("Mercado pago ma", "ACC_MERCADOP"),  # main
    ("Mercado Pago ma", "ACC_MERCADOP"),
    ("Mercado Pago", "ACC_MERCADOP"),
    ("Mercado pago", "ACC_MERCADOP"),
    ("Nubank", "ACC_NUBANKMA"),
    ("Bosch", "ACC_BOSCHLIF"),
    ("Meli", "ACC_MELIDOLA"),
    ("FGTS", "ACC_FGTS"),
]

# Credit card names (exact or prefix)
CARD_PREFIX: list[tuple[str, str]] = [
    ("Mercado Pago", "CARD_MP"),
    ("Mercado pago", "CARD_MP"),
]
