# -*- coding: utf-8 -*-
"""Portrait trophy hall for the phone. Overwrites one PNG on every run."""
from __future__ import annotations

import argparse
import json
from pathlib import Path

from PIL import Image

PHONE_W = 1080
PHONE_H = 1920
BANDS = 3
TROPHY_LIMIT = 12
# Same pedestal centers and cap lines as the web hall.
XS = (0.174, 0.295, 0.427, 0.573, 0.706, 0.820)
UPPER_BOTTOM = 0.625
LOWER_BOTTOM = 0.375
SPOT_FRAC = 0.08
IMAGE_PREFIX = "/assets/trophies/"


def _session_items(path: Path) -> list[dict]:
    if not path.is_file():
        return []
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []
    items = payload.get("items") if isinstance(payload, dict) else None
    if not isinstance(items, list):
        return []
    return [row for row in items if isinstance(row, dict)]


def _trophy_file(web: Path, item: dict) -> Path | None:
    if item.get("equipped") is False:
        return None
    image = str(item.get("image") or "").strip()
    if not image.startswith(IMAGE_PREFIX) or not image.lower().endswith(".png"):
        return None
    name = Path(image).name
    if name != image[len(IMAGE_PREFIX) :]:
        return None
    path = web / "assets" / "trophies" / name
    if not path.is_file():
        return None
    return path


def _fit_bottom(image: Image.Image, box: int) -> Image.Image:
    image = image.convert("RGBA")
    scale = min(box / image.width, box / image.height)
    size = (max(1, round(image.width * scale)), max(1, round(image.height * scale)))
    return image.resize(size, Image.Resampling.LANCZOS)


def _paste_trophy(
    canvas: Image.Image, trophy: Image.Image, band_top: int, band_h: int, index: int
) -> None:
    col = index % 6
    bottom_frac = UPPER_BOTTOM if index < 6 else LOWER_BOTTOM
    spot = max(1, round(PHONE_W * SPOT_FRAC))
    center_x = round(XS[col] * PHONE_W)
    bottom_y = band_top + band_h - round(bottom_frac * band_h)
    art = _fit_bottom(trophy, spot)
    left = center_x - art.width // 2
    top = bottom_y - art.height
    canvas.alpha_composite(art, (left, top))


def render(scripts_root: Path) -> Path:
    root = Path(scripts_root)
    web = root / "mnemonics" / "web"
    hall_path = web / "assets" / "trophy-hall.png"
    if not hall_path.is_file():
        raise FileNotFoundError(f"trophy hall image missing: {hall_path}")
    session_path = root / "mnemonics" / "data" / "collectibles_session.json"
    out = web / "assets" / "phone" / "trophy-hall.png"

    hall = Image.open(hall_path).convert("RGBA")
    band_h = PHONE_H // BANDS
    band = hall.resize((PHONE_W, band_h), Image.Resampling.LANCZOS)
    canvas = Image.new("RGBA", (PHONE_W, PHONE_H))
    for band_index in range(BANDS):
        canvas.paste(band, (0, band_index * band_h))

    pictures: list[Path] = []
    for item in _session_items(session_path):
        path = _trophy_file(web, item)
        if path is None:
            continue
        pictures.append(path)
        if len(pictures) >= TROPHY_LIMIT:
            break
    for index, path in enumerate(pictures):
        with Image.open(path) as trophy:
            _paste_trophy(canvas, trophy, 0, band_h, index)

    out.parent.mkdir(parents=True, exist_ok=True)
    canvas.convert("RGB").save(out, "PNG")
    return out


def main() -> None:
    parser = argparse.ArgumentParser(description="Render the portrait trophy hall")
    parser.add_argument("--scripts-root", type=Path, required=True)
    args = parser.parse_args()
    render(args.scripts_root)


if __name__ == "__main__":
    main()
