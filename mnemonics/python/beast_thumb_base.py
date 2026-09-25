"""Canonical beast-thumb slug resolver (A–Z bases + Lynne Kelly uniques).

Rules:
- Custom 2-letter pegs (Bone goat, Fire goat, …) share the A–Z beast for the
  second letter → one icon per base animal.
- Lynne Kelly / single-letter codes keep a distinct icon per name.
- CSV names like "Bone goat" without metadata also collapse onto the A–Z base
  when the leading word is a known Custom adjective from bestiary.json.
"""

from __future__ import annotations

import json
import re
from functools import lru_cache
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
BESTIARY_PATH = ROOT / "technique" / "bestiary.json"


def slug(name: str) -> str:
    s = "".join(ch if ch.isalnum() else "_" for ch in name.lower())
    return re.sub(r"_+", "_", s).strip("_") or "beast"


def short_label(raw: str) -> str:
    """Beast CSV sometimes embeds Context/Quote into beast_name — keep the title."""
    name = (raw or "").strip()
    for sep in (" Context:", " Quote:", " Narrative:", "\n"):
        if sep in name:
            name = name.split(sep, 1)[0].strip()
    if name.startswith("[") and "]" in name:
        name = name[1 : name.index("]")].strip() or name
    return name[:80] if len(name) > 80 else name


@lru_cache(maxsize=1)
def load_bestiary() -> dict:
    data = json.loads(BESTIARY_PATH.read_text(encoding="utf-8"))
    items = data.get("items") or []
    by_code = {it["code"]: it for it in items if it.get("code")}
    az_by_letter = {
        it["code"].upper(): it["name"]
        for it in items
        if len(it.get("code") or "") == 1
    }
    az_name_to_slug = {name.lower(): slug(name) for name in az_by_letter.values()}
    custom_adjectives: set[str] = set()
    for it in items:
        code = it.get("code") or ""
        if it.get("source") == "Custom" and len(code) == 2:
            parts = (it.get("name") or "").split(" ", 1)
            if len(parts) == 2 and parts[0]:
                custom_adjectives.add(parts[0].lower())
    return {
        "items": items,
        "by_code": by_code,
        "az_by_letter": az_by_letter,
        "az_name_to_slug": az_name_to_slug,
        "custom_adjectives": frozenset(custom_adjectives),
        "total_count": data.get("total_count") or len(items),
    }


def _match_az_tail(label: str, data: dict) -> str | None:
    """If label is '{CustomAdj} {A-Z beast name}', return that beast's slug."""
    lower = label.lower().strip()
    az_name_to_slug = data["az_name_to_slug"]
    adjs = data["custom_adjectives"]
    best = None
    best_len = 0
    for name_l, s in az_name_to_slug.items():
        if lower == name_l:
            return s
        suffix = " " + name_l
        if not lower.endswith(suffix) or len(name_l) <= best_len:
            continue
        head = lower[: -len(suffix)].strip()
        if head in adjs:
            best = s
            best_len = len(name_l)
    return best


def canonical_slug(
    name: str,
    code: str | None = None,
    source: str | None = None,
) -> str:
    """Return the shared thumb slug for a bestiary/beast name."""
    data = load_bestiary()
    label = short_label(name)
    code = (code or "").strip()
    source = (source or "").strip()

    if not source and code and code in data["by_code"]:
        source = data["by_code"][code].get("source") or ""

    if source == "Custom" and len(code) == 2:
        letter = code[1].upper()
        az_name = data["az_by_letter"].get(letter)
        if az_name:
            return slug(az_name)

    if len(code) == 1:
        az_name = data["az_by_letter"].get(code.upper())
        if az_name:
            return slug(az_name)

    tail = _match_az_tail(label, data)
    if tail:
        return tail

    return slug(label)


def canonical_for_bestiary_item(item: dict) -> str:
    return canonical_slug(
        item.get("name") or "",
        code=item.get("code"),
        source=item.get("source"),
    )


def all_canonical_entries() -> list[dict]:
    """One row per bestiary item with its canonical slug."""
    data = load_bestiary()
    rows = []
    for it in data["items"]:
        s = canonical_for_bestiary_item(it)
        rows.append(
            {
                "code": it.get("code") or "",
                "name": it.get("name") or "",
                "source": it.get("source") or "",
                "order": it.get("order"),
                "canonical_slug": s,
                "filename": f"{s}.png",
                "url": f"/assets/beast-thumbs/{s}.png",
            }
        )
    return rows


def unique_canonical_targets() -> list[dict]:
    """Deduped list of icons that must exist (26 bases + Lynne Kelly uniques)."""
    seen: dict[str, dict] = {}
    for row in all_canonical_entries():
        s = row["canonical_slug"]
        if s not in seen:
            seen[s] = {
                "slug": s,
                "label": row["name"],
                "code": row["code"],
                "source": row["source"],
            }
    data = load_bestiary()
    for letter, name in data["az_by_letter"].items():
        s = slug(name)
        seen[s] = {
            "slug": s,
            "label": name,
            "code": letter,
            "source": data["by_code"][letter].get("source") or "Lynne Kelly",
        }
    return sorted(seen.values(), key=lambda x: x["slug"])
