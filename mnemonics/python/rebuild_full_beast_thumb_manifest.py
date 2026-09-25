"""Rebuild beast-thumb-manifest.json for ALL beasts; emit bestiary map + need_generate."""

from __future__ import annotations

import csv
import json
from collections import defaultdict
from datetime import date
from pathlib import Path

from beast_thumb_base import (
    all_canonical_entries,
    canonical_slug,
    load_bestiary,
    short_label,
    unique_canonical_targets,
)

ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "web" / "assets"
THUMBS = ASSETS / "beast-thumbs"
SOURCES = ASSETS / "_beast_thumb_sources"
MANIFEST = ASSETS / "beast-thumb-manifest.json"
BESTIARY_MAP = ASSETS / "bestiary-thumb-map.json"
NEED_GEN = ASSETS / "_beast_thumb_need_generate.json"


def main() -> None:
    beasts = list(csv.DictReader((ROOT / "data" / "beasts.csv").open(encoding="utf-8")))
    icons = {}
    by_slug: dict[str, list[str]] = defaultdict(list)
    for b in beasts:
        bid = b["id"]
        label = short_label(b.get("beast_name") or bid)
        peg = b.get("peg_code") or ""
        s = canonical_slug(label, code=peg or None)
        icons[bid] = {
            "label": label,
            "peg_code": peg,
            "filename": f"{s}.png",
            "url": f"/assets/beast-thumbs/{s}.png",
            "source_slug": s,
        }
        by_slug[s].append(bid)

    unique_slugs = set(by_slug)
    catalog = unique_canonical_targets()
    manifest = {
        "version": 3,
        "style": {
            "name": "Memory Quest Bestiary",
            "description": (
                "Cartoon RPG inventory icons for beast pegs — warm gold and navy "
                "accents, chunky readable silhouette."
            ),
            "format": "PNG RGBA",
            "dimensions_px": [256, 256],
            "background": "transparent",
            "generated_with": "Cursor GenerateImage",
            "generated_at": str(date.today()),
            "source_layout": (
                "Custom adjective variants share the A–Z second-letter base icon; "
                "Lynne Kelly / unique pegs each get their own canonical slug PNG"
            ),
            "generation_prompt_template": (
                "Isolated RPG inventory item icon of {beast_name}, polished mobile "
                "fantasy-game art, soft isometric 3D cartoon style, warm gold and "
                "dark navy accents, chunky readable silhouette, no text, no letters; "
                "solid chroma-key magenta background (#FF00FF)."
            ),
            "catalog_scope": (
                f"{len(beasts)} beasts / {len(unique_slugs)} in-use slugs / "
                f"{len(catalog)} bestiary canonical icons"
            ),
        },
        "icons": icons,
    }
    ASSETS.mkdir(parents=True, exist_ok=True)
    THUMBS.mkdir(parents=True, exist_ok=True)
    SOURCES.mkdir(parents=True, exist_ok=True)
    MANIFEST.write_text(
        json.dumps(manifest, indent=2, ensure_ascii=False) + "\n", encoding="utf-8"
    )

    # Full bestiary code → thumb map (702)
    bestiary_rows = all_canonical_entries()
    BESTIARY_MAP.write_text(
        json.dumps(
            {
                "version": 1,
                "total_count": load_bestiary()["total_count"],
                "canonical_icon_count": len(catalog),
                "by_code": {r["code"]: r for r in bestiary_rows},
            },
            indent=2,
            ensure_ascii=False,
        )
        + "\n",
        encoding="utf-8",
    )

    have_thumbs = {p.stem for p in THUMBS.glob("*.png")}
    have_sources = {p.stem for p in SOURCES.glob("*.png")}
    missing_ids = [
        bid for bid, meta in icons.items() if meta["source_slug"] not in have_thumbs
    ]

    # Need generate: Lynne Kelly / unique canonicals only (not Custom variants)
    need_gen = []
    for t in catalog:
        s = t["slug"]
        if s in have_sources or s in have_thumbs:
            continue
        need_gen.append(
            {
                "label": t["label"],
                "slug": s,
                "code": t["code"],
                "source": t["source"],
            }
        )

    NEED_GEN.write_text(
        json.dumps(need_gen, indent=2, ensure_ascii=False) + "\n", encoding="utf-8"
    )

    print(f"total_beasts={len(beasts)}")
    print(f"in_use_slugs={len(unique_slugs)}")
    print(f"bestiary_canonical={len(catalog)}")
    print(f"thumbs_present={len(have_thumbs)}")
    print(f"ids_missing_thumbs={len(missing_ids)}")
    print(f"sources_present={len(have_sources)}")
    print(f"need_generate={len(need_gen)}")
    for it in need_gen[:30]:
        print(f"NEED\t{it['slug']}\t{it['code']}\t{it['label']}")
    if len(need_gen) > 30:
        print(f"... +{len(need_gen) - 30} more")


if __name__ == "__main__":
    main()
