"""Rebuild beast-thumb-manifest.json for ALL beasts; report missing unique sources."""

from __future__ import annotations

import csv
import json
import re
from collections import defaultdict
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "web" / "assets"
THUMBS = ASSETS / "beast-thumbs"
SOURCES = ASSETS / "_beast_thumb_sources"
MANIFEST = ASSETS / "beast-thumb-manifest.json"


def slug(name: str) -> str:
    s = "".join(ch if ch.isalnum() else "_" for ch in name.lower())
    return re.sub(r"_+", "_", s).strip("_") or "beast"


def short_label(raw: str) -> str:
    """Beast CSV sometimes embeds Context/Quote into beast_name — keep the title."""
    name = (raw or "").strip()
    for sep in (" Context:", " Quote:", " Narrative:", "\n"):
        if sep in name:
            name = name.split(sep, 1)[0].strip()
    # Strip wrapping brackets from accidental exports like "[Agaric fungi]"
    if name.startswith("[") and "]" in name:
        name = name[1 : name.index("]")].strip() or name
    return name[:80] if len(name) > 80 else name


def main() -> None:
    beasts = list(csv.DictReader((ROOT / "data" / "beasts.csv").open(encoding="utf-8")))
    icons = {}
    by_label: dict[str, list[str]] = defaultdict(list)
    for b in beasts:
        bid = b["id"]
        label = short_label(b.get("beast_name") or bid)
        s = slug(label)
        icons[bid] = {
            "label": label,
            "peg_code": b.get("peg_code") or "",
            "filename": f"{s}.png",
            "url": f"/assets/beast-thumbs/{s}.png",
            "source_slug": s,
        }
        by_label[label].append(bid)

    unique_slugs = {meta["source_slug"] for meta in icons.values()}
    manifest = {
        "version": 2,
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
                "One chroma-keyed PNG per unique short-label slug; all beast ids "
                "sharing that label point at the same shared file"
            ),
            "generation_prompt_template": (
                "Isolated RPG inventory item icon of {beast_name}, polished mobile "
                "fantasy-game art, soft isometric 3D cartoon style, warm gold and "
                "dark navy accents, chunky readable silhouette, no text, no letters; "
                "solid chroma-key magenta background (#FF00FF)."
            ),
            "catalog_scope": (
                f"{len(beasts)} beasts / {len(by_label)} unique labels / "
                f"{len(unique_slugs)} shared thumb files"
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

    have_thumbs = {p.stem for p in THUMBS.glob("*.png")}
    have_sources = {p.stem for p in SOURCES.glob("*.png")}
    missing_ids = [
        bid for bid, meta in icons.items() if meta["source_slug"] not in have_thumbs
    ]
    need_gen = []
    for label, ids in sorted(by_label.items(), key=lambda x: x[0].lower()):
        s = slug(label)
        if s not in have_sources:
            need_gen.append({"label": label, "slug": s, "ids": ids})

    print(f"total_beasts={len(beasts)}")
    print(f"unique_labels={len(by_label)}")
    print(f"unique_slugs={len(unique_slugs)}")
    print(f"thumbs_present={len(have_thumbs)}")
    print(f"ids_missing_thumbs={len(missing_ids)}")
    print(f"sources_present={len(have_sources)}")
    print(f"unique_labels_need_generate={len(need_gen)}")
    for it in need_gen[:30]:
        print(f"NEED\t{it['slug']}\tx{len(it['ids'])}\t{it['label']}")
    if len(need_gen) > 30:
        print(f"... +{len(need_gen) - 30} more")
    (ASSETS / "_beast_thumb_need_generate.json").write_text(
        json.dumps(need_gen, indent=2, ensure_ascii=False) + "\n", encoding="utf-8"
    )


if __name__ == "__main__":
    main()
