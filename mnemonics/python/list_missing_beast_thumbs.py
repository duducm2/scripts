"""List beasts missing thumbs; coverage is by shared source_slug PNG."""

from __future__ import annotations

import csv
import json
import re
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "web" / "assets"
THUMBS = ASSETS / "beast-thumbs"
MANIFEST = ASSETS / "beast-thumb-manifest.json"
WORKLIST = ASSETS / "_beast_thumb_worklist.json"


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


def main() -> None:
    beasts = list(csv.DictReader((ROOT / "data" / "beasts.csv").open(encoding="utf-8")))
    have = {p.stem for p in THUMBS.glob("*.png")}
    missing = []
    for b in beasts:
        label = short_label(b.get("beast_name") or b["id"])
        if slug(label) not in have:
            missing.append(b)
    print(f"total={len(beasts)} have={len(have)} missing={len(missing)}")

    by_name: dict[str, list[str]] = defaultdict(list)
    for b in missing:
        label = short_label(b.get("beast_name") or b["id"])
        by_name[label].append(b["id"])
    items = [
        {"name": n, "ids": ids, "slug": slug(n)}
        for n, ids in sorted(by_name.items(), key=lambda x: x[0].lower())
    ]
    WORKLIST.write_text(
        json.dumps(items, indent=2, ensure_ascii=False) + "\n", encoding="utf-8"
    )
    print(f"missing_unique_names={len(items)}")
    for it in items[:25]:
        print(f"  {it['slug']}\tx{len(it['ids'])}\t{it['name']}")
    if len(items) > 25:
        print(f"  ... +{len(items) - 25} more")


if __name__ == "__main__":
    main()
