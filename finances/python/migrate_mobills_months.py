"""Migrate Mobills UIA transaction dumps → finances/data/transactions.csv."""

from __future__ import annotations

import argparse
import csv
import re
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from data_aggregator import DATA, OUTPUT, parse_decimal, read_csv  # noqa: E402
from mobills_name_maps import (  # noqa: E402
    ACCOUNT_EXACT,
    ACCOUNT_PREFIX,
    CARD_PREFIX,
    CATEGORY_EXACT,
    CATEGORY_PREFIX,
)

ROOT = Path(__file__).resolve().parents[2]
MONTHS_DIR = ROOT / "finances" / "mobills" / "months"
TX_PATH = DATA / "transactions.csv"
REPORT_PATH = OUTPUT / "mobills_migration_report.txt"

MONTH_FILE_MAP = {
    "jan": "01",
    "fev": "02",
    "mar": "03",
    "apr": "04",
    "may": "05",
    "jun": "06",
    "jul": "07",
    "aug": "08",
}

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

NAME_ITEM_RE = re.compile(
    r'Type: 50029 \(DataItem\) Name: "([^"]+)" LocalizedType: "item"'
)
DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{4}$")
AMOUNT_RE = re.compile(r"^R\$\s*")
SKIP_DESC_PREFIXES = ("Expected end of day balance",)


def format_csv_decimal(num: float) -> str:
    neg = num < 0
    n = abs(num)
    s = f"{n:.2f}".replace(".", ",")
    return f"-{s}" if neg else s


def mobills_amount_to_float(raw: str) -> float:
    """Parse Mobills UI amounts like 'R$ 10,000.00' or 'R$ 3.00'."""
    s = raw.strip()
    s = AMOUNT_RE.sub("", s).strip()
    # US-style thousands + decimal point
    if "," in s and "." in s:
        s = s.replace(",", "")
    elif "," in s and "." not in s:
        # unlikely BR in this dump; treat comma as decimal
        s = s.replace(",", ".")
    try:
        return abs(float(s))
    except ValueError:
        return 0.0


def mmddyyyy_to_iso(d: str) -> str:
    mm, dd, yyyy = d.split("/")
    return f"{yyyy}-{mm}-{dd}"


def normalize_key(s: str) -> str:
    return s.replace("…", "").replace("...", "").replace("\\*", "*").strip().casefold()


def resolve_category(name: str, cat_by_id: dict) -> tuple[str, str, str | None]:
    """Return (category_id, subcategory, unmapped_name_or_None)."""
    if name in CATEGORY_EXACT:
        cid, sub = CATEGORY_EXACT[name]
        return cid, sub, None
    # Try exact after stripping trailing * _
    stripped = name.rstrip("*_\\ ").strip()
    if stripped in CATEGORY_EXACT:
        cid, sub = CATEGORY_EXACT[stripped]
        return cid, sub, None
    best: tuple[str, str, str] | None = None
    for prefix, cid, sub in CATEGORY_PREFIX:
        if name.startswith(prefix) or normalize_key(name).startswith(
            normalize_key(prefix)
        ):
            if best is None or len(prefix) > len(best[0]):
                best = (prefix, cid, sub)
    if best:
        return best[1], best[2], None
    # Match against English names in our CSV (casefold / startswith)
    nk = normalize_key(name)
    for cid, row in cat_by_id.items():
        en = normalize_key(row.get("name") or "")
        if en and (nk == en or nk.startswith(en) or en.startswith(nk)):
            parent = (row.get("parent_id") or "").strip()
            if parent:
                return parent, row.get("name") or "", None
            return cid, "", None
    return "", "", name


def resolve_account(name: str, acc_by_id: dict) -> tuple[str, str | None]:
    if name in ACCOUNT_EXACT:
        return ACCOUNT_EXACT[name], None
    for label, aid in ACCOUNT_EXACT.items():
        if normalize_key(name).startswith(normalize_key(label)[:8]):
            return aid, None
    for prefix, aid in ACCOUNT_PREFIX:
        if name.startswith(prefix) or normalize_key(name).startswith(
            normalize_key(prefix)
        ):
            return aid, None
    # Fuzzy against accounts.csv names
    nk = normalize_key(name).rstrip(".")
    best_id = ""
    best_len = 0
    for aid, row in acc_by_id.items():
        an = normalize_key(row.get("name") or "")
        # compare prefixes of truncated UI label
        for L in range(min(len(nk), len(an)), 5, -1):
            if an.startswith(nk[:L]) or nk.startswith(an[:L]):
                if L > best_len:
                    best_len = L
                    best_id = aid
                break
    if best_id:
        return best_id, None
    return "", name


def resolve_card(account_label: str, cards: list[dict]) -> tuple[str, str]:
    """If label matches a card, return (card_id, linked_account_id)."""
    for prefix, card_id in CARD_PREFIX:
        if account_label.startswith(prefix) or normalize_key(account_label).startswith(
            normalize_key(prefix)
        ):
            for c in cards:
                if c.get("id") == card_id:
                    return card_id, c.get("linked_account_id") or ""
            return card_id, ""
    for c in cards:
        cname = c.get("name") or ""
        if normalize_key(account_label).startswith(normalize_key(cname)[:8]):
            return c.get("id") or "", c.get("linked_account_id") or ""
    return "", ""


def infer_type(
    category_id: str,
    category_label: str,
    card_id: str,
    cat_by_id: dict,
) -> str:
    cl = category_label.casefold()
    if "transfer" in cl or category_id == "CAT_TRANSFER":
        return "transfer"
    if card_id:
        return "card_expense"
    row = cat_by_id.get(category_id) or {}
    if (row.get("type") or "") == "income":
        return "income"
    return "expense"


def parse_dump(text: str) -> list[dict]:
    names = NAME_ITEM_RE.findall(text)
    rows: list[dict] = []
    i = 0
    while i < len(names):
        n = names[i]
        if DATE_RE.fullmatch(n) and i + 4 < len(names):
            desc, cat, acc, amt = names[i + 1], names[i + 2], names[i + 3], names[i + 4]
            if amt.startswith("R$") and not any(
                desc.startswith(p) for p in SKIP_DESC_PREFIXES
            ):
                rows.append(
                    {
                        "date_raw": n,
                        "description": desc,
                        "category_label": cat,
                        "account_label": acc,
                        "amount_raw": amt,
                    }
                )
                i += 5
                continue
        i += 1
    return rows


def load_sources(extra_sample: Path | None) -> list[tuple[str, Path]]:
    sources: list[tuple[str, Path]] = []
    if MONTHS_DIR.is_dir():
        for path in sorted(MONTHS_DIR.glob("*.md")):
            if path.stat().st_size > 0:
                stem = path.stem.lower()
                sources.append((stem, path))
    if not sources and extra_sample and extra_sample.is_file():
        sources.append(("aug", extra_sample))
    return sources


def recompute_budgets(txs: list[dict]) -> None:
    budgets = read_csv("budgets.csv")
    cats = {c["id"]: c for c in read_csv("categories.csv")}
    by_ym: dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
    for tx in txs:
        if tx["type"] not in ("expense", "card_expense"):
            continue
        ym = tx["date"][:7]
        cid = tx["category_id"]
        cat = cats.get(cid)
        main = cid
        if cat and (cat.get("parent_id") or "").strip():
            main = cat["parent_id"]
        by_ym[ym][main] += parse_decimal(tx["amount"])
    for b in budgets:
        ym = b.get("year_month") or ""
        cid = b.get("category_id") or ""
        spent = by_ym.get(ym, {}).get(cid, 0.0)
        b["spent_amount"] = format_csv_decimal(spent)
    path = DATA / "budgets.csv"
    fields = ["year_month", "category_id", "planned_amount", "spent_amount"]
    with path.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields, quoting=csv.QUOTE_MINIMAL)
        w.writeheader()
        for b in budgets:
            w.writerow({k: b.get(k, "") for k in fields})


def migrate(sample: Path | None, write: bool) -> int:
    cats_list = read_csv("categories.csv")
    accs_list = read_csv("accounts.csv")
    cards = read_csv("credit_cards.csv")
    cat_by_id = {c["id"]: c for c in cats_list}
    acc_by_id = {a["id"]: a for a in accs_list}

    sources = load_sources(sample)
    report: list[str] = []
    if not sources:
        report.append("No non-empty month dumps found under finances/mobills/months/.")
        report.append(
            "Pass --sample finances/mobills/transactions.md to migrate August sample."
        )
        OUTPUT.mkdir(parents=True, exist_ok=True)
        REPORT_PATH.write_text("\n".join(report) + "\n", encoding="utf-8")
        print(REPORT_PATH)
        return 1

    all_rows: list[dict] = []
    unmapped_cats: set[str] = set()
    unmapped_accs: set[str] = set()
    per_file: dict[str, int] = {}

    for stem, path in sources:
        text = path.read_text(encoding="utf-8")
        raw_rows = parse_dump(text)
        per_file[stem] = len(raw_rows)
        for r in raw_rows:
            cid, sub, um_cat = resolve_category(r["category_label"], cat_by_id)
            if um_cat:
                unmapped_cats.add(um_cat)
            aid, um_acc = resolve_account(r["account_label"], acc_by_id)
            if um_acc:
                unmapped_accs.add(um_acc)
            card_id, linked = "", ""
            # Only treat as card when the label did not resolve to a bank account
            if not aid:
                card_id, linked = resolve_card(r["account_label"], cards)
                if card_id:
                    aid = linked
            t = infer_type(cid, r["category_label"], card_id, cat_by_id)
            amt = mobills_amount_to_float(r["amount_raw"])
            iso = mmddyyyy_to_iso(r["date_raw"])
            transfer_dest = ""
            if t == "transfer":
                # account_id = destination shown; leave transfer_account_id empty if unknown
                pass
            all_rows.append(
                {
                    "date": iso,
                    "description": r["description"],
                    "amount": format_csv_decimal(amt),
                    "type": t,
                    "category_id": cid,
                    "subcategory": sub,
                    "account_id": aid,
                    "card_id": card_id if t == "card_expense" else "",
                    "transfer_account_id": transfer_dest,
                    "_key": (
                        iso,
                        r["description"],
                        format_csv_decimal(amt),
                        aid,
                        cid,
                        r["category_label"],
                    ),
                    "_src": stem,
                }
            )

    # Deduplicate
    seen: set[tuple] = set()
    unique: list[dict] = []
    dupes = 0
    for row in all_rows:
        k = row["_key"]
        if k in seen:
            dupes += 1
            continue
        seen.add(k)
        unique.append(row)

    unique.sort(key=lambda r: (r["date"], r["description"], r["amount"]))
    out_rows: list[dict] = []
    for i, row in enumerate(unique, start=1):
        out_rows.append(
            {
                "id": f"TX{i:04d}",
                "date": row["date"],
                "description": row["description"],
                "amount": row["amount"],
                "type": row["type"],
                "category_id": row["category_id"],
                "subcategory": row["subcategory"],
                "account_id": row["account_id"],
                "card_id": row["card_id"],
                "transfer_account_id": row["transfer_account_id"],
            }
        )

    report.append(f"Sources: {len(sources)}")
    for stem, n in per_file.items():
        report.append(f"  {stem}: {n} parsed rows")
    report.append(f"Total parsed: {len(all_rows)}")
    report.append(f"Duplicates removed: {dupes}")
    report.append(f"Unique written: {len(out_rows)}")
    if unmapped_cats:
        report.append("UNMAPPED_CATEGORY:")
        for c in sorted(unmapped_cats):
            report.append(f"  {c}")
    else:
        report.append("UNMAPPED_CATEGORY: (none)")
    if unmapped_accs:
        report.append("UNMAPPED_ACCOUNT:")
        for a in sorted(unmapped_accs):
            report.append(f"  {a}")
    else:
        report.append("UNMAPPED_ACCOUNT: (none)")

    by_type: dict[str, int] = defaultdict(int)
    for r in out_rows:
        by_type[r["type"]] += 1
    report.append(
        "By type: " + ", ".join(f"{k}={v}" for k, v in sorted(by_type.items()))
    )

    OUTPUT.mkdir(parents=True, exist_ok=True)
    REPORT_PATH.write_text("\n".join(report) + "\n", encoding="utf-8")
    print("\n".join(report))
    print(f"Report: {REPORT_PATH}")

    if write:
        with TX_PATH.open("w", encoding="utf-8", newline="") as f:
            w = csv.DictWriter(f, fieldnames=TX_HEADERS, quoting=csv.QUOTE_MINIMAL)
            w.writeheader()
            for r in out_rows:
                w.writerow(r)
        recompute_budgets(out_rows)
        print(f"Wrote {TX_PATH} ({len(out_rows)} rows) and recomputed budgets.")
    else:
        print("Dry-run only (pass --write to save transactions.csv).")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Migrate Mobills UIA dumps to transactions.csv"
    )
    ap.add_argument(
        "--sample",
        type=Path,
        default=ROOT / "finances" / "mobills" / "transactions.md",
        help="Fallback dump when months/*.md are empty",
    )
    ap.add_argument(
        "--write",
        action="store_true",
        help="Write finances/data/transactions.csv and recompute budgets",
    )
    ap.add_argument(
        "--no-sample-fallback",
        action="store_true",
        help="Do not fall back to --sample when month files are empty",
    )
    args = ap.parse_args()
    sample = None if args.no_sample_fallback else args.sample
    return migrate(sample, args.write)


if __name__ == "__main__":
    raise SystemExit(main())
