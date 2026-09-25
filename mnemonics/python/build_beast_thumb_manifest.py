"""Build beast-thumb-manifest.json for Quick Recall–scoped beasts."""
from __future__ import annotations

import csv
import json
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def main() -> None:
    beasts = list(csv.DictReader((ROOT / "data" / "beasts.csv").open(encoding="utf-8")))
    qr = json.loads((ROOT / "data" / "quick_recall.json").read_text(encoding="utf-8"))
    palace_ids = set(qr.get("included_palace_ids") or [])
    qr_beasts = sorted(
        [b for b in beasts if b["palace_id"] in palace_ids],
        key=lambda b: b["id"],
    )

    icons = {}
    for b in qr_beasts:
        bid = b["id"]
        icons[bid] = {
            "label": b["beast_name"],
            "peg_code": b.get("peg_code") or "",
            "filename": f"{bid}.png",
            "url": f"/assets/beast-thumbs/{bid}.png",
        }

    manifest = {
        "version": 1,
        "style": {
            "name": "Memory Quest Bestiary",
            "description": (
                "Cartoon RPG inventory icons for Quick Recall beast pegs — "
                "warm gold and navy accents, chunky readable silhouette."
            ),
            "format": "PNG RGBA",
            "dimensions_px": [256, 256],
            "background": "transparent",
            "generated_with": "Cursor GenerateImage",
            "source_layout": (
                "Per unique beast_name icon, chroma-keyed magenta then copied "
                "to each beast id filename"
            ),
            "generation_prompt_template": (
                "Isolated RPG inventory item icon of {beast_name}, polished "
                "mobile fantasy-game art, soft isometric 3D cartoon style, "
                "warm gold and dark navy accents, chunky readable silhouette, "
                "no text, no letters; solid chroma-key magenta background (#FF00FF)."
            ),
            "catalog_scope": (
                f"{len(qr_beasts)} Quick Recall beasts across "
                f"{len(palace_ids)} included palaces"
            ),
        },
        "icons": icons,
    }

    assets = ROOT / "web" / "assets"
    assets.mkdir(parents=True, exist_ok=True)
    (assets / "beast-thumbs").mkdir(parents=True, exist_ok=True)
    out = assets / "beast-thumb-manifest.json"
    out.write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print(f"wrote {out} icons={len(icons)}")

    by_name: dict[str, list[str]] = defaultdict(list)
    for b in qr_beasts:
        by_name[b["beast_name"]].append(b["id"])
    print(f"unique_names={len(by_name)}")
    for name, bids in sorted(by_name.items(), key=lambda x: x[0].lower()):
        print(f"{name}\t{len(bids)}\t{','.join(bids)}")


if __name__ == "__main__":
    main()
