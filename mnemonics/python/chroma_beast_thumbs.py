"""Chroma-key magenta beast icon sources into one transparent 256x256 PNG per slug."""

from __future__ import annotations

import argparse
import json
from collections import defaultdict
from pathlib import Path

from PIL import Image

from beast_thumb_base import all_adjective_slugs, canonical_slug, slug

ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "web" / "assets"
THUMBS = ASSETS / "beast-thumbs"
MANIFEST = ASSETS / "beast-thumb-manifest.json"
SIZE = 256
TOL = 70


def _is_chroma(r: int, g: int, b: int, bg: tuple[int, int, int]) -> bool:
    br, bg_, bb = bg
    if abs(r - br) <= TOL and abs(g - bg_) <= TOL and abs(b - bb) <= TOL:
        return True
    if r >= 200 and g <= 60 and b >= 120:
        return True
    return False


def chroma_to_rgba(src: Path) -> Image.Image:
    im = Image.open(src).convert("RGBA")
    w, h = im.size
    samples = [
        im.getpixel((2, 2))[:3],
        im.getpixel((w - 3, 2))[:3],
        im.getpixel((2, h - 3))[:3],
        im.getpixel((w - 3, h - 3))[:3],
    ]
    bg = tuple(sum(c[i] for c in samples) // len(samples) for i in range(3))
    px = im.load()
    for y in range(h):
        for x in range(w):
            r, g, b, a = px[x, y]
            if _is_chroma(r, g, b, bg):
                px[x, y] = (0, 0, 0, 0)
    side = max(w, h)
    canvas = Image.new("RGBA", (side, side), (0, 0, 0, 0))
    canvas.paste(im, ((side - w) // 2, (side - h) // 2), im)
    return canvas.resize((SIZE, SIZE), Image.Resampling.LANCZOS)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument(
        "--sources",
        type=Path,
        default=ASSETS / "_beast_thumb_sources",
        help="Dir of source images named {slug}.png (canonical beast_name)",
    )
    ap.add_argument(
        "--slugs-from-bestiary",
        action="store_true",
        help="Also chroma every unique canonical target from bestiary.json",
    )
    ap.add_argument(
        "--adjectives-only",
        action="store_true",
        help="Only process adj_* adjective modifier sources",
    )
    args = ap.parse_args()
    sources = args.sources
    if not MANIFEST.is_file() and not args.adjectives_only:
        raise SystemExit(f"missing manifest: {MANIFEST}")

    by_slug: dict[str, list[tuple[str, str]]] = defaultdict(list)

    if not args.adjectives_only and MANIFEST.is_file():
        manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
        icons = manifest.get("icons") or {}
        for bid, meta in icons.items():
            label = meta.get("label") or bid
            peg = meta.get("peg_code") or ""
            s = meta.get("source_slug") or canonical_slug(label, code=peg or None)
            by_slug[s].append((bid, label))

    if args.slugs_from_bestiary:
        from beast_thumb_base import unique_canonical_targets

        for t in unique_canonical_targets():
            by_slug.setdefault(t["slug"], []).append(("", t["label"]))

    # Always include adjective modifier icons when their sources exist
    for adj_s in all_adjective_slugs():
        by_slug.setdefault(adj_s, []).append(("", adj_s))

    if args.adjectives_only:
        by_slug = {s: v for s, v in by_slug.items() if s.startswith("adj_")}

    THUMBS.mkdir(parents=True, exist_ok=True)
    done = 0
    missing = []
    for s, pairs in sorted(by_slug.items()):
        src = sources / f"{s}.png"
        if not src.is_file():
            # try raw slug of first label if different
            alt_label = pairs[0][1]
            alt = sources / f"{slug(alt_label)}.png"
            if alt.is_file():
                src = alt
            else:
                missing.append(pairs[0][1] or s)
                continue
        rgba = chroma_to_rgba(src)
        dest = THUMBS / f"{s}.png"
        rgba.save(dest, "PNG")
        done += 1
        print(f"ok {s} -> 1 file shared by {len(pairs)} id(s) ({pairs[0][1]})")

    print(f"wrote={done} missing_labels={len(missing)}")
    for m in missing[:40]:
        print(f"  MISSING {m}")
    if len(missing) > 40:
        print(f"  ... +{len(missing) - 40} more")


if __name__ == "__main__":
    main()
