"""List beasts missing thumbs; coverage is by canonical source_slug PNG."""

from __future__ import annotations

import csv
import json
from collections import defaultdict
from pathlib import Path

from beast_thumb_base import canonical_slug, short_label

ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "web" / "assets"
THUMBS = ASSETS / "beast-thumbs"
WORKLIST = ASSETS / "_beast_thumb_worklist.json"


def main() -> None:
    beasts = list(csv.DictReader((ROOT / "data" / "beasts.csv").open(encoding="utf-8")))
    have = {p.stem for p in THUMBS.glob("*.png")}
    missing = []
    for b in beasts:
        label = short_label(b.get("beast_name") or b["id"])
        peg = b.get("peg_code") or ""
        s = canonical_slug(label, code=peg or None)
        if s not in have:
            missing.append({**b, "_slug": s, "_label": label})
    print(f"total={len(beasts)} have={len(have)} missing={len(missing)}")

    by_slug: dict[str, list[str]] = defaultdict(list)
    labels: dict[str, str] = {}
    for b in missing:
        s = b["_slug"]
        by_slug[s].append(b["id"])
        labels[s] = b["_label"]
    items = [
        {"name": labels[s], "ids": ids, "slug": s}
        for s, ids in sorted(by_slug.items(), key=lambda x: x[0])
    ]
    WORKLIST.write_text(json.dumps(items, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print(f"missing_unique_names={len(items)}")
    for it in items[:25]:
        print(f"  {it['slug']}\tx{len(it['ids'])}\t{it['name']}")
    if len(items) > 25:
        print(f"  ... +{len(items) - 25} more")


if __name__ == "__main__":
    main()
