"""Palace collectibles: schema, repair, Desktop pack import, relic prompt."""

from __future__ import annotations

import argparse
import csv
import json
import os
import random
import re
import uuid
from collections import deque
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

TROPHY_LIMIT = 12
SLOTS = ("trophy",)
SLOT_ORDER = ("trophy",)
SLOT_ANCHOR = {"trophy": "pedestal"}
STRIP_REGION: dict[str, str] = {}
SLOT_LAYER = {"trophy": "trophy"}
SLOT_OCCLUDES: dict[str, tuple[str, ...]] = {}
STRIP_SLOTS: tuple[str, ...] = ()
HELD_ICON_SLOTS: tuple[str, ...] = ()
ANCHORED_ICON_SLOTS = ("trophy",)
SIDEKICK_SLOTS: tuple[str, ...] = ()
_VIEWBOX = re.compile(r"""viewBox\s*=\s*["']([^"']+)["']""", re.IGNORECASE)
ANIMS = ("bob", "sway", "flicker", "orbit", "float")
ANCHORS = ("pedestal",)
# One look is chosen per prompt. The model never sees this list.
TROPHY_LOOKS = (
    "a woodcut print",
    "a stained-glass piece",
    "a hand-built clay figure",
    "an ink wash painting",
    "an enamel badge",
    "a cast bronze miniature",
    "a cut-paper collage",
    "a neon tube object",
    "an embroidered patch",
    "a blown-glass sculpture",
    "a tile mosaic",
    "a porcelain figurine",
    "a felt craft",
    "a linocut",
    "a watercolor study",
    "an oil painting of one object",
    "an origami model",
    "a charcoal sketch",
    "a tin toy",
    "a batik motif",
    "a gouache illustration",
    "a knitted charm",
    "a lacquer ornament",
    "a wax-seal impression",
    "a marble carving",
    "a cardboard model",
    "a ballpoint doodle",
    "a risograph print",
    "a shadow puppet",
    "a beadwork medallion",
    "a stop-motion puppet",
    "a glossy studio object",
    "a flat cartoon drawing",
    "a simple 3D render",
)
TROPHY_LOOK_MEMORY = 12
IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".webp"}
TROPHY_IMAGE_PREFIX = "/assets/trophies/"
TROPHY_IMAGE_MAX = 512

_FENCE = "```"
_TRAILING_COMMA = re.compile(r",(\s*[}\]])")
_FILE_BLOCK = re.compile(
    r"(?:===|---)FILE:\s*COLLECTIBLE\.json\s*(?:===|---)\s*(.*?)\s*(?:===|---)END_FILE(?:===|---)",
    re.IGNORECASE | re.DOTALL,
)
_UNSAFE_BLOCK = re.compile(
    r"<\s*(script|foreignObject|iframe|object|embed)\b[^>]*>.*?<\s*/\s*\1\s*>",
    re.IGNORECASE | re.DOTALL,
)
_UNSAFE_VOID = re.compile(
    r"<\s*(script|foreignObject|iframe|object|embed)\b[^>]*/\s*>",
    re.IGNORECASE,
)
_ON_ATTR = re.compile(
    r"\s+on[a-z]+\s*=\s*(?:\"[^\"]*\"|'[^']*'|[^\s>]+)",
    re.IGNORECASE,
)
_JS_URL = re.compile(r"javascript\s*:", re.IGNORECASE)


def repl_path(data_dir: Path) -> Path:
    """Saved collectibles. Nothing is written here until the user presses Save."""
    return Path(data_dir) / "repl" / "collectibles.json"


def session_path(data_dir: Path) -> Path:
    """Trophies on the hall shelves. They are never copied into the library."""
    return Path(data_dir) / "collectibles_session.json"


def requests_path(data_dir: Path) -> Path:
    """Category assigned to the relic request currently in flight."""
    return Path(data_dir) / "collectible_requests.json"


def _read_store(path: Path) -> list[dict[str, Any]]:
    if not path.is_file():
        return []
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []
    rows = raw.get("items") if isinstance(raw, dict) else raw
    if not isinstance(rows, list):
        return []
    items: list[dict[str, Any]] = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        cleaned = _stored_item(row)
        if cleaned:
            items.append(cleaned)
    return items


def _write_store(path: Path, items: list[dict[str, Any]]) -> None:
    if not items:
        if path.is_file():
            path.unlink()
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps({"items": items}, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def load_saved(data_dir: Path) -> list[dict[str, Any]]:
    data_dir = Path(data_dir)
    repl = repl_path(data_dir)
    legacy = data_dir / "collectibles.json"
    if not repl.is_file() and legacy.is_file():
        _write_store(repl, _read_store(legacy))
    return _read_store(repl)


def load_session(data_dir: Path) -> list[dict[str, Any]]:
    return _read_store(session_path(data_dir))


def load_items(data_dir: Path) -> list[dict[str, Any]]:
    """The hall is the session queue only, oldest first, at most twelve."""
    items: list[dict[str, Any]] = []
    for row in load_session(data_dir):
        public = _public_item(row, saved=False)
        if public:
            items.append(public)
    return items[:TROPHY_LIMIT]


def _stored_item(row: dict[str, Any]) -> dict[str, Any] | None:
    cleaned = _public_item(row, saved=False)
    if cleaned is None:
        return None
    cleaned.pop("saved", None)
    cleaned.pop("layer", None)
    cleaned.pop("occludes", None)
    cleaned.pop("misfit", None)
    created = str(row.get("created_at") or "").strip()
    if created:
        cleaned["created_at"] = created
    return cleaned


def _public_item(row: dict[str, Any], saved: bool = False) -> dict[str, Any] | None:
    slot = str(row.get("slot") or "").strip().lower()
    if slot not in SLOTS:
        return None
    image = _public_image(str(row.get("image") or ""))
    svg = sanitize_svg(str(row.get("svg") or ""))
    if not svg and not image:
        return None
    anim = str(row.get("anim") or "").strip().lower()
    anchor = str(row.get("anchor") or "").strip().lower()
    return {
        "id": str(row.get("id") or "").strip(),
        "palace_id": str(row.get("palace_id") or "").strip(),
        "slot": slot,
        "name": str(row.get("name") or "").strip()[:80],
        "blurb": str(row.get("blurb") or "").strip()[:240],
        "svg": svg,
        "image": image,
        "anim": anim if anim in ANIMS else "bob",
        "anchor": anchor if anchor in ANCHORS else "pedestal",
        "equipped": bool(row.get("equipped")),
        "saved": saved,
        "layer": SLOT_LAYER.get(slot, "gear_accessory"),
        "occludes": list(SLOT_OCCLUDES.get(slot, ())),
        "misfit": _strip_misfit(slot, svg),
    }


def sanitize_svg(svg: str) -> str:
    text = (svg or "").strip()
    if not text:
        return ""
    text = _UNSAFE_BLOCK.sub("", text)
    text = _UNSAFE_VOID.sub("", text)
    text = _ON_ATTR.sub("", text)
    text = _JS_URL.sub("", text)
    start = text.lower().find("<svg")
    if start < 0:
        return ""
    end = text.lower().rfind("</svg>")
    if end < start:
        return ""
    text = text[start : end + len("</svg>")]
    if len(text) > 12000:
        return ""
    if "<svg" not in text.lower():
        return ""
    return text


def svg_viewbox(svg: str) -> tuple[float, float, float, float] | None:
    match = _VIEWBOX.search(svg or "")
    if not match:
        return None
    parts = [part for part in re.split(r"[\s,]+", match.group(1).strip()) if part]
    if len(parts) != 4:
        return None
    try:
        return tuple(float(part) for part in parts)  # type: ignore[return-value]
    except ValueError:
        return None


def _near(value: float, target: float, tolerance: float = 2) -> bool:
    return abs(value - target) < tolerance


def validate_strip(slot: str, svg: str) -> tuple[str, str]:
    """Return (error, note). A non-empty error rejects the import."""
    box = svg_viewbox(svg)
    if slot in STRIP_SLOTS:
        if box is None or box[2] < 180 or box[3] < 180:
            return (
                'This slot needs a standing overlay. Use viewBox "0 0 305 424". '
                "A small icon was rejected.",
                "",
            )
        width, height = box[2], box[3]
        if (_near(width, 772) or _near(width, 193)) and _near(height, 424):
            return (
                'That frame size was rejected. Draw the standing T-pose with viewBox "0 0 305 424".',
                "",
            )
        if _near(width, 305) and _near(height, 424):
            return "", ""
        return (
            "",
            "Strip size does not match the standing walker (305x424). Marked as misfit.",
        )
    if slot in ANCHORED_ICON_SLOTS:
        if box is None or box[2] >= 300 or box[3] >= 400:
            return (
                'This slot needs a detailed icon. Use viewBox "0 0 256 256". '
                "A body strip was rejected.",
                "",
            )
        width, height = box[2], box[3]
        if _near(width, 256, 12) and _near(height, 256, 12):
            return "", ""
        return (
            "",
            "Icon grid is 256x256 so the drawing can hold detail. Marked as misfit.",
        )
    if slot in SIDEKICK_SLOTS:
        if box is None or (_near(box[2], 305, 8) and _near(box[3], 424, 8)):
            return (
                'A sidekick is a standing figure. Use viewBox "0 0 192 320". '
                "A body strip was rejected.",
                "",
            )
        if _near(box[2], 192, 12) and _near(box[3], 320, 16):
            return "", ""
        return (
            "",
            "Sidekick grid is 192x320 so the figure can hold a face and clothes. Marked as misfit.",
        )
    return "", ""


def _strip_misfit(slot: str, svg: str) -> bool:
    if (
        slot not in STRIP_SLOTS
        and slot not in ANCHORED_ICON_SLOTS
        and slot not in SIDEKICK_SLOTS
    ):
        return False
    error, note = validate_strip(slot, svg)
    return bool(error or note)


def _strip_fence(text: str) -> str:
    body = text.strip()
    if not body.startswith(_FENCE):
        return body
    rest = body[len(_FENCE) :]
    nl = rest.find("\n")
    if nl >= 0:
        rest = rest[nl + 1 :]
    if rest.rstrip().endswith(_FENCE):
        rest = rest.rstrip()[: -len(_FENCE)]
    return rest.strip()


def _strip_trailing_commas(text: str) -> str:
    prev = None
    while prev != text:
        prev = text
        text = _TRAILING_COMMA.sub(r"\1", text)
    return text


def extract_json_text(raw: str) -> str:
    text = (raw or "").replace("\r\n", "\n").replace("\r", "\n")
    match = _FILE_BLOCK.search(text)
    if match:
        text = match.group(1)
    text = _strip_fence(text)
    start = text.find("{")
    end = text.rfind("}")
    if start < 0 or end <= start:
        return ""
    return _strip_trailing_commas(text[start : end + 1])


def _as_item_dict(parsed: Any) -> dict[str, Any] | None:
    if not isinstance(parsed, dict):
        return None
    inner = parsed.get("item")
    if isinstance(inner, dict):
        return inner
    items = parsed.get("items")
    if isinstance(items, list) and items and isinstance(items[0], dict):
        return items[0]
    if "slot" in parsed or "svg" in parsed or "image" in parsed or "name" in parsed:
        return parsed
    return None


def _public_image(value: str) -> str:
    text = str(value or "").strip()
    if not text.startswith(TROPHY_IMAGE_PREFIX):
        return ""
    name = Path(text).name
    if name != text[len(TROPHY_IMAGE_PREFIX) :] or Path(name).suffix.lower() != ".png":
        return ""
    return TROPHY_IMAGE_PREFIX + name


def _trophy_dir(data_dir: Path) -> Path:
    web = Path(data_dir).parent / "web" / "assets"
    root = web if web.is_dir() else Path(data_dir) / "web-assets"
    return root / "trophies"


def _delete_trophy_file(data_dir: Path, row: dict[str, Any]) -> None:
    url = _public_image(str(row.get("image") or ""))
    if not url:
        return
    path = _trophy_dir(data_dir) / Path(url).name
    if path.is_file():
        path.unlink()


def _backdrop_kind(r: int, g: int, b: int) -> str:
    if min(r, g, b) >= 235 and (max(r, g, b) - min(r, g, b)) <= 12:
        return "white"
    if r >= 190 and b >= 140 and g <= 90 and (r - g) >= 70:
        return "magenta"
    return ""


def _is_backdrop_pixel(r: int, g: int, b: int, kind: str) -> bool:
    """Pale backdrop, including the off-white sheet a hard 235 cutoff used to leave."""
    if kind == "white":
        return min(r, g, b) >= 200 and (max(r, g, b) - min(r, g, b)) <= 28
    if kind == "magenta":
        return r >= 175 and b >= 120 and g <= 100 and (r - g) >= 55
    return False


def _is_fringe_pixel(r: int, g: int, b: int, kind: str) -> bool:
    if kind == "white":
        return min(r, g, b) >= 186 and (max(r, g, b) - min(r, g, b)) <= 22
    return _is_backdrop_pixel(r, g, b, kind)


def _is_pure_white(r: int, g: int, b: int) -> bool:
    """White the generator may use only for the backdrop and for see-through holes."""
    return min(r, g, b) >= 250 and (max(r, g, b) - min(r, g, b)) <= 5


def _clear_connected_backdrop(im: Any, kind: str) -> None:
    width, height = im.size
    px = im.load()
    seen = bytearray(width * height)
    queue: deque[tuple[int, int]] = deque()

    def push(x: int, y: int) -> None:
        index = y * width + x
        if seen[index]:
            return
        seen[index] = 1
        red, green, blue, alpha = px[x, y]
        if alpha >= 20 and not _is_backdrop_pixel(red, green, blue, kind):
            return
        queue.append((x, y))

    for x in range(width):
        push(x, 0)
        push(x, height - 1)
    for y in range(height):
        push(0, y)
        push(width - 1, y)
    while queue:
        x, y = queue.popleft()
        px[x, y] = (0, 0, 0, 0)
        if x > 0:
            push(x - 1, y)
        if x + 1 < width:
            push(x + 1, y)
        if y > 0:
            push(x, y - 1)
        if y + 1 < height:
            push(x, y + 1)


def _peel_fringe(im: Any, kind: str, passes: int = 2) -> None:
    """Drop a thin pale rim that is still stuck to the transparent edge."""
    width, height = im.size
    px = im.load()
    for _ in range(passes):
        doomed: list[tuple[int, int]] = []
        for y in range(height):
            for x in range(width):
                red, green, blue, alpha = px[x, y]
                if alpha < 20 or not _is_fringe_pixel(red, green, blue, kind):
                    continue
                if (
                    (x > 0 and px[x - 1, y][3] < 20)
                    or (x + 1 < width and px[x + 1, y][3] < 20)
                    or (y > 0 and px[x, y - 1][3] < 20)
                    or (y + 1 < height and px[x, y + 1][3] < 20)
                ):
                    doomed.append((x, y))
        if not doomed:
            return
        for x, y in doomed:
            px[x, y] = (0, 0, 0, 0)


def _crop_to_content(im: Any, pad: int = 1) -> Any:
    """Drop empty margin so the object fills the pedestal instead of floating inside it."""
    alpha = im.getchannel("A")
    mask = alpha.point(lambda value: 255 if value >= 20 else 0)
    box = mask.getbbox()
    if not box:
        return im
    left, top, right, bottom = box
    left = max(0, left - pad)
    top = max(0, top - pad)
    right = min(im.width, right + pad)
    bottom = min(im.height, bottom + pad)
    if (left, top, right, bottom) == (0, 0, im.width, im.height):
        return im
    return im.crop((left, top, right, bottom))


def _clear_pure_white(im: Any) -> None:
    """Drop pure white wherever it sits, including a hole the trophy frames."""
    width, height = im.size
    px = im.load()
    for y in range(height):
        for x in range(width):
            red, green, blue, alpha = px[x, y]
            if alpha < 20 or not _is_pure_white(red, green, blue):
                continue
            px[x, y] = (0, 0, 0, 0)


def knock_out_backdrop(image: Any) -> Any:
    """Clear a white or magenta backdrop, and pure white enclosed by the trophy.

    Off-white pixels stay. The generator is told to reserve pure white for the
    backdrop and for openings, and to paint trophy parts a step darker.
    """
    im = image.convert("RGBA")
    width, height = im.size
    if width < 2 or height < 2:
        return im
    px = im.load()
    corners = (
        px[0, 0],
        px[width - 1, 0],
        px[0, height - 1],
        px[width - 1, height - 1],
    )
    kinds = [
        "clear" if pixel[3] < 20 else _backdrop_kind(pixel[0], pixel[1], pixel[2])
        for pixel in corners
    ]
    if kinds.count("clear") == 4:
        # The flat field is already gone. Peel a pale mat that was left on the silhouette.
        for kind in ("white", "magenta"):
            _clear_connected_backdrop(im, kind)
            _peel_fringe(im, kind)
        _clear_pure_white(im)
        return _crop_to_content(im)
    kind = ""
    if kinds.count("white") >= 3:
        kind = "white"
    elif kinds.count("magenta") >= 3:
        kind = "magenta"
    if kind:
        _clear_connected_backdrop(im, kind)
        _peel_fringe(im, kind)
    _clear_pure_white(im)
    return _crop_to_content(im)


def prepare_trophy_png(src: Path) -> Any:
    from PIL import Image

    im = Image.open(src)
    im = knock_out_backdrop(im)
    width, height = im.size
    longest = max(width, height)
    if longest > TROPHY_IMAGE_MAX:
        scale = TROPHY_IMAGE_MAX / longest
        im = im.resize(
            (max(1, round(width * scale)), max(1, round(height * scale))),
            Image.Resampling.LANCZOS,
        )
    return im


def _newest_desktop_image(desktop: Path | None) -> Path | None:
    """The picture most recently added on the Desktop. The pack does not name it."""
    if desktop is None:
        return None
    root = Path(desktop)
    if not root.is_dir():
        return None
    newest: Path | None = None
    newest_stamp = -1.0
    for path in root.iterdir():
        if not path.is_file() or path.suffix.lower() not in IMAGE_SUFFIXES:
            continue
        try:
            stat = path.stat()
        except OSError:
            continue
        stamp = max(stat.st_mtime, getattr(stat, "st_ctime", 0))
        if stamp >= newest_stamp:
            newest_stamp = stamp
            newest = path
    return newest


def _store_trophy_png(data_dir: Path, item_id: str, image: Any) -> str:
    folder = _trophy_dir(data_dir)
    folder.mkdir(parents=True, exist_ok=True)
    filename = f"{item_id}.png"
    image.save(folder / filename, "PNG")
    return TROPHY_IMAGE_PREFIX + filename


def known_palace_ids(data_dir: Path) -> set[str]:
    path = Path(data_dir) / "palaces.csv"
    if not path.is_file():
        return set()
    with path.open(encoding="utf-8-sig", newline="") as handle:
        return {
            (row.get("id") or "").strip()
            for row in csv.DictReader(handle)
            if (row.get("id") or "").strip()
        }


def repair_item(
    raw: str, palace_ids: set[str], *, check_art: bool = True
) -> tuple[dict[str, Any] | None, list[str], str]:
    """Return (item, notes, error). Error is empty when the item can be saved."""
    notes: list[str] = []
    blob = extract_json_text(raw)
    if not blob:
        return (
            None,
            notes,
            "No JSON object found. Expected COLLECTIBLE.json inside COLLECTIBLE_PACK.txt.",
        )
    try:
        parsed = json.loads(blob)
    except json.JSONDecodeError as exc:
        return None, notes, f"JSON could not be parsed after repair: {exc.msg}"
    item = _as_item_dict(parsed)
    if item is None:
        return None, notes, "JSON has no collectible object (slot or name)."

    slot = str(item.get("slot") or "").strip().lower()
    if slot not in SLOTS:
        notes.append(f"Unknown slot {slot or '(blank)'} mapped to trophy.")
        slot = "trophy"
    anim = str(item.get("anim") or "").strip().lower()
    if anim not in ANIMS:
        notes.append(f"Unknown anim {anim or '(blank)'} mapped to bob.")
        anim = "bob"
    anchor = str(item.get("anchor") or "").strip().lower()
    if anchor not in ANCHORS:
        notes.append(f"Unknown anchor {anchor or '(blank)'} mapped to pedestal.")
        anchor = "pedestal"
    name = str(item.get("name") or "").strip()
    if not name:
        return None, notes, "Missing name."
    palace_id = str(item.get("palace_id") or item.get("palaceId") or "").strip()
    if not palace_id or palace_id not in palace_ids:
        return (
            None,
            notes,
            f"palace_id {palace_id or '(blank)'} is not a palace in this library.",
        )
    svg = sanitize_svg(str(item.get("svg") or ""))
    if svg and check_art:
        strip_error, strip_note = validate_strip(slot, svg)
        if strip_error:
            return None, notes, strip_error
        if strip_note:
            notes.append(strip_note)
    return (
        {
            "palace_id": palace_id,
            "slot": slot,
            "name": name[:80],
            "blurb": str(item.get("blurb") or "").strip()[:240],
            "svg": svg,
            "image": "",
            "anim": anim,
            "anchor": anchor,
        },
        notes,
        "",
    )


def _unequip_slot(rows: list[dict[str, Any]], slot: str, keep_id: str = "") -> None:
    for row in rows:
        if row.get("slot") == slot and row.get("id") != keep_id:
            row["equipped"] = False


def _load_requests(data_dir: Path) -> dict[str, Any]:
    path = requests_path(data_dir)
    empty = {
        "pending_slot": "",
        "pending_palace_id": "",
        "cycle": [],
        "recent_looks": [],
    }
    if not path.is_file():
        return empty
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return empty
    if not isinstance(raw, dict):
        return empty
    pending = str(raw.get("pending_slot") or "").strip().lower()
    cycle = [
        slot
        for slot in raw.get("cycle") or []
        if str(slot).strip().lower() in SLOT_ORDER
    ]
    recent = [
        str(look) for look in raw.get("recent_looks") or [] if str(look) in TROPHY_LOOKS
    ]
    return {
        "pending_slot": pending if pending in SLOT_ORDER else "",
        "pending_palace_id": str(raw.get("pending_palace_id") or "").strip(),
        "cycle": [str(slot).strip().lower() for slot in cycle],
        "recent_looks": recent[-TROPHY_LOOK_MEMORY:],
    }


def _save_requests(data_dir: Path, state: dict[str, Any]) -> None:
    path = requests_path(data_dir)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(
            {
                "pending_slot": state.get("pending_slot") or "",
                "pending_palace_id": state.get("pending_palace_id") or "",
                "cycle": list(state.get("cycle") or []),
                "recent_looks": [
                    look
                    for look in state.get("recent_looks") or []
                    if look in TROPHY_LOOKS
                ][-TROPHY_LOOK_MEMORY:],
            },
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )


def _occupied_slots(data_dir: Path) -> set[str]:
    occupied: set[str] = set()
    for row in load_session(data_dir) + load_saved(data_dir):
        slot = str(row.get("slot") or "").strip().lower()
        if slot in SLOT_ORDER:
            occupied.add(slot)
    return occupied


def _slot_candidates(occupied: set[str], cycle: set[str]) -> list[str]:
    """Next categories to ask for. Pets and accessories wait until the rest are covered."""
    early = [slot for slot in SLOT_ORDER if slot not in ("pet", "accessory")]
    pool = [slot for slot in SLOT_ORDER if slot not in occupied] or list(SLOT_ORDER)
    ready = all(slot in occupied or slot in cycle for slot in early)
    picks: list[str] = []
    for slot in pool:
        if slot in cycle:
            continue
        if slot in ("pet", "accessory") and not ready:
            continue
        picks.append(slot)
    return picks


def assign_slot(data_dir: Path, palace_id: str) -> str:
    """Remember the next relic category and return it."""
    data_dir = Path(data_dir)
    state = _load_requests(data_dir)
    occupied = _occupied_slots(data_dir)
    cycle = list(state["cycle"])
    picks = _slot_candidates(occupied, set(cycle))
    if not picks:
        cycle = []
        picks = _slot_candidates(occupied, set())
    slot = picks[0]
    if slot not in cycle:
        cycle.append(slot)
    _save_requests(
        data_dir,
        {
            "pending_slot": slot,
            "pending_palace_id": palace_id,
            "cycle": cycle,
            "recent_looks": state.get("recent_looks") or [],
        },
    )
    return slot


def _lock_imported_slot(
    data_dir: Path, item: dict[str, Any], notes: list[str]
) -> dict[str, Any]:
    """Store the relic in the category we asked for, even if the pack picked another."""
    state = _load_requests(data_dir)
    pending = state.get("pending_slot") or ""
    if pending not in SLOT_ORDER:
        return item
    anchor = SLOT_ANCHOR[pending]
    if item.get("slot") != pending or item.get("anchor") != anchor:
        notes.append(
            f"Requested {pending}; stored that slot instead of {item.get('slot') or '(blank)'}."
        )
        item["slot"] = pending
        item["anchor"] = anchor
    cycle = list(state.get("cycle") or [])
    if pending not in cycle:
        cycle.append(pending)
    _save_requests(
        data_dir,
        {
            "pending_slot": "",
            "pending_palace_id": "",
            "cycle": cycle,
            "recent_looks": state.get("recent_looks") or [],
        },
    )
    return item


def import_text(
    data_dir: Path, raw: str, *, desktop: Path | None = None
) -> dict[str, Any]:
    """Put a trophy on the newest end of the hall. The library is not written."""
    data_dir = Path(data_dir)
    item, notes, error = repair_item(raw, known_palace_ids(data_dir), check_art=False)
    if error or item is None:
        return {
            "ok": False,
            "error": error or "Could not repair collectible.",
            "notes": notes,
        }
    item = _lock_imported_slot(data_dir, item, notes)
    prepared = None
    if not item.get("svg"):
        src = _newest_desktop_image(desktop)
        if src is None:
            return {
                "ok": False,
                "error": "No trophy picture on the Desktop. Save the image there and import again.",
                "notes": notes,
            }
        try:
            prepared = prepare_trophy_png(src)
        except OSError as exc:
            return {
                "ok": False,
                "error": f"The newest Desktop picture could not be read: {exc}",
                "notes": notes,
            }
    else:
        strip_error, strip_note = validate_strip(
            str(item.get("slot") or ""), str(item.get("svg") or "")
        )
        if strip_error:
            return {"ok": False, "error": strip_error, "notes": notes}
        if strip_note and strip_note not in notes:
            notes.append(strip_note)
    session = load_session(data_dir)
    held = {
        "id": "col_" + uuid.uuid4().hex[:12],
        "created_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        **item,
        "equipped": True,
    }
    if prepared is not None:
        held["image"] = _store_trophy_png(data_dir, held["id"], prepared)
    session.append(held)
    if len(session) > TROPHY_LIMIT:
        dropped_rows = session[: len(session) - TROPHY_LIMIT]
        dropped = len(dropped_rows)
        for row in dropped_rows:
            _delete_trophy_file(data_dir, row)
        session = session[dropped:]
        notes.append(
            f"The hall holds {TROPHY_LIMIT}. Dropped the {dropped} oldest trophy"
            + ("s." if dropped != 1 else ".")
        )
    _write_store(session_path(data_dir), session)
    return {
        "ok": True,
        "item": _public_item(held, saved=False),
        "notes": notes,
        "count": len(load_items(data_dir)),
        "saved": False,
    }


def persist_item(data_dir: Path, item_id: str) -> dict[str, Any]:
    """Trophies stay in the session hall and are never written to the library."""
    del data_dir, item_id
    return {
        "ok": False,
        "error": "Trophies stay in the hall. They are not saved to the library.",
    }


def remove_item(data_dir: Path, item_id: str) -> dict[str, Any]:
    """Take an item off the avatar. An unsaved relic is deleted immediately."""
    data_dir = Path(data_dir)
    item_id = (item_id or "").strip()
    session = load_session(data_dir)
    saved = load_saved(data_dir)
    in_session = any(row.get("id") == item_id for row in session)
    in_saved = any(row.get("id") == item_id for row in saved)
    if not in_session and not in_saved:
        return {"ok": False, "error": "No collectible with that id."}
    if in_session and not in_saved:
        for row in session:
            if row.get("id") == item_id:
                _delete_trophy_file(data_dir, row)
        session = [row for row in session if row.get("id") != item_id]
        _write_store(session_path(data_dir), session)
        return {"ok": True, "items": load_items(data_dir), "deleted": True}
    for row in saved:
        if row.get("id") == item_id:
            row["equipped"] = False
    if in_session:
        session = [row for row in session if row.get("id") != item_id]
        _write_store(session_path(data_dir), session)
    _write_store(repl_path(data_dir), saved)
    return {"ok": True, "items": load_items(data_dir), "deleted": False}


def set_equipped(data_dir: Path, item_id: str, equipped: bool) -> dict[str, Any]:
    data_dir = Path(data_dir)
    item_id = (item_id or "").strip()
    session = load_session(data_dir)
    saved = load_saved(data_dir)
    target = next((row for row in session if row.get("id") == item_id), None)
    saved_target = next((row for row in saved if row.get("id") == item_id), None)
    if target is None and saved_target is None:
        return {"ok": False, "error": "No collectible with that id."}
    if not equipped and target is not None and saved_target is None:
        _delete_trophy_file(data_dir, target)
        session = [row for row in session if row.get("id") != item_id]
        _write_store(session_path(data_dir), session)
        return {"ok": True, "items": load_items(data_dir), "deleted": True}
    if equipped:
        if target is not None:
            target["equipped"] = True
        if saved_target is not None:
            saved_target["equipped"] = True
    elif saved_target is not None:
        saved_target["equipped"] = False
    _write_store(session_path(data_dir), session)
    _write_store(repl_path(data_dir), saved)
    return {"ok": True, "items": load_items(data_dir)}


def _csv_rows(data_dir: Path, name: str) -> list[dict[str, str]]:
    path = Path(data_dir) / name
    if not path.is_file():
        return []
    with path.open(encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def build_prompt(data_dir: Path, palace_id: str) -> dict[str, Any]:
    data_dir = Path(data_dir)
    palace_id = (palace_id or "").strip()
    palaces = _csv_rows(data_dir, "palaces.csv")
    palace = next(
        (row for row in palaces if (row.get("id") or "").strip() == palace_id), None
    )
    if palace is None:
        return {"ok": False, "error": "Open a palace before looking for a relic."}
    study_id = (palace.get("study_id") or "").strip()
    study = next(
        (
            row
            for row in _csv_rows(data_dir, "studies.csv")
            if (row.get("id") or "").strip() == study_id
        ),
        {},
    )
    beasts = [
        row
        for row in _csv_rows(data_dir, "beasts.csv")
        if (row.get("palace_id") or "").strip() == palace_id
    ]
    beast_ids = {(row.get("id") or "").strip() for row in beasts}
    keywords: list[str] = []
    for atom in _csv_rows(data_dir, "atoms.csv"):
        if (atom.get("beast_id") or "").strip() not in beast_ids:
            continue
        for part in re.split(r"[,;\n]", atom.get("keywords") or ""):
            bit = part.strip()
            if bit and bit not in keywords:
                keywords.append(bit)
            if len(keywords) >= 24:
                break
    beast_names = [
        ((row.get("beast_name") or "").strip())
        for row in beasts
        if (row.get("beast_name") or "").strip()
    ]
    study_title = (study.get("title") or study_id or "this study").strip()
    title = (palace.get("title") or palace_id).strip()
    character = (palace.get("character_name") or "").strip()
    slot = assign_slot(data_dir, palace_id)
    anchor = SLOT_ANCHOR[slot]
    look = pick_trophy_look(data_dir)
    prompt = _prompt_text(
        palace_id=palace_id,
        study_title=study_title,
        title=title,
        character=character,
        beasts=beast_names[:8],
        keywords=keywords[:24],
        slot=slot,
        anchor=anchor,
        look=look,
    )
    return {
        "ok": True,
        "prompt": prompt,
        "palace_id": palace_id,
        "title": title,
        "slot": slot,
        "anchor": anchor,
    }


def pick_trophy_look(data_dir: Path) -> str:
    """Choose one look the model has not used lately. The prompt names only that look."""
    data_dir = Path(data_dir)
    state = _load_requests(data_dir)
    recent = [look for look in state.get("recent_looks") or [] if look in TROPHY_LOOKS]
    pool = [look for look in TROPHY_LOOKS if look not in recent] or list(TROPHY_LOOKS)
    look = random.choice(pool)
    recent.append(look)
    state["recent_looks"] = recent[-TROPHY_LOOK_MEMORY:]
    _save_requests(data_dir, state)
    return look


def _prompt_text(
    *,
    palace_id: str,
    study_title: str,
    title: str,
    character: str,
    beasts: list[str],
    keywords: list[str],
    slot: str,
    anchor: str,
    look: str,
) -> str:
    beast_line = ", ".join(beasts) if beasts else "(none)"
    keyword_line = ", ".join(keywords) if keywords else "(none)"
    return (
        "Invent ONE trophy for a Memory Palace hall.\n"
        "Use your image generator. Do not draw an SVG.\n"
        "Make one object that belongs to this palace. It should feel like a prize worth keeping.\n"
        f"Render this trophy in exactly this look: {look}. "
        "Use that look for the whole object. Do not switch to another medium.\n"
        "One object only. Not a scene and not a portrait. "
        "The named look decides the form. Do not default to a cup.\n"
        "Show the whole object. Every edge must sit inside the picture, "
        "including the lowest part, so nothing is sliced off.\n"
        "Keep it large, with a margin of empty background on every side, including under it.\n"
        "Do not paint a floor, a shadow, a glow, or a pedestal under it. "
        "The hall already has the stone table.\n"
        "Background: transparent if you can. Otherwise one flat pure white, "
        "RGB 255,255,255, the same color in every background pixel, with no gradient.\n"
        "If the trophy has an opening and the hall should show through it, "
        "paint that opening the same pure white, even when the object surrounds it.\n"
        "I delete every pure-white pixel, including white enclosed inside the trophy. "
        "Pure white means red, green, and blue are all 250 or higher and almost the same.\n"
        "Do not paint any part of the trophy pure white. "
        "A part that should look white must be off-white, such as 245,242,230, so I keep it.\n"
        "Do not make the whole object white, off-white, or pale gray.\n"
        f'Copy slot "{slot}" and anchor "{anchor}" exactly. Do not change them.\n\n'
        f"Study: {study_title}\n"
        f"Palace id (copy exactly): {palace_id}\n"
        f"Palace title: {title}\n"
        f"Main topic, which the trophy must show: {title}\n"
        f"Beasts, for flavor only: {beast_line}\n"
        f"Keywords, for flavor only: {keyword_line}\n\n"
        "Deliver the picture as its own image file, and one file named exactly COLLECTIBLE_PACK.txt "
        "(download chip, or one marked fence).\n"
        "Do not put a picture filename in the JSON. I use the newest image on the Desktop.\n"
        "Never claim you saved it to Desktop. I save the files myself.\n"
        "The pack body is only these markers and one JSON object. No Markdown outside the pack.\n\n"
        "===FILE: COLLECTIBLE.json===\n"
        "{\n"
        f'  "palace_id": "{palace_id}",\n'
        f'  "slot": "{slot}",\n'
        '  "name": "Short trophy name",\n'
        '  "blurb": "One sentence on why this trophy belongs to this palace.",\n'
        '  "anim": "bob",\n'
        f'  "anchor": "{anchor}"\n'
        "}\n"
        "===END_FILE===\n\n"
        "Rules:\n"
        f'- slot must be exactly "{slot}".\n'
        f'- anchor must be exactly "{anchor}".\n'
        f"- anim is one of: {', '.join(ANIMS)}. It moves only the trophy.\n"
        "- Do not include an image filename or an svg. The picture is a separate file.\n"
        "- The picture is one object on a transparent background, or on flat RGB 255,255,255 if transparency is unavailable.\n"
        "- Openings the hall should show through are that same pure white, even inside the object. "
        "I delete every pure-white pixel. Trophy parts are never pure white; a white-looking part is off-white, such as 245,242,230.\n"
        "- Show the whole object. The picture edge must not cut off any part, especially the bottom.\n"
        "- No floor, shadow, glow, or pedestal in the picture. Leave empty background under the object too.\n"
        "- The named look decides the form. Do not default to a cup.\n"
        "- The whole object is not white, off-white, or pale gray. A white-looking part is off-white, never pure white.\n"
        "- The name names the palace topic. The blurb says why the trophy belongs to it.\n"
        "- No text in the picture.\n"
        "- Replace the example name, blurb, and anim. Keep palace_id, slot, and anchor exactly as given.\n"
        "- Re-deliver with the exact filename COLLECTIBLE_PACK.txt. "
        "Do not add updated, corrected, or v2 to the name.\n"
    )


def queue_prompt(data_dir: Path, prompt: str) -> None:
    path = Path(data_dir) / "collectible_prompt_pending.txt"
    path.write_text(prompt, encoding="utf-8")


def gemini_is_open() -> bool:
    if os.name != "nt":
        return False
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.windll.user32
    found = {"ok": False}

    @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    def visit(hwnd, _lparam):
        if not user32.IsWindowVisible(hwnd):
            return True
        length = user32.GetWindowTextLengthW(hwnd)
        if length <= 0:
            return True
        buf = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, buf, length + 1)
        if "gemini" in buf.value.lower():
            found["ok"] = True
            return False
        return True

    user32.EnumWindows(visit, 0)
    return found["ok"]


def set_clipboard(text: str) -> bool:
    if os.name != "nt":
        return False
    import ctypes

    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    if not user32.OpenClipboard(0):
        return False
    try:
        user32.EmptyClipboard()
        data = ctypes.create_unicode_buffer(text)
        size = ctypes.sizeof(data)
        handle = kernel32.GlobalAlloc(0x0002, size)
        if not handle:
            return False
        locked = kernel32.GlobalLock(handle)
        if not locked:
            kernel32.GlobalFree(handle)
            return False
        ctypes.memmove(locked, ctypes.addressof(data), size)
        kernel32.GlobalUnlock(handle)
        if not user32.SetClipboardData(13, handle):
            kernel32.GlobalFree(handle)
            return False
        return True
    finally:
        user32.CloseClipboard()


def prepare_prompt(data_dir: Path, palace_id: str) -> dict[str, Any]:
    built = build_prompt(data_dir, palace_id)
    if not built.get("ok"):
        return built
    prompt = str(built["prompt"])
    copied = set_clipboard(prompt)
    open_gemini = gemini_is_open()
    if open_gemini:
        queue_prompt(data_dir, prompt)
    return {
        "ok": True,
        "copied": copied,
        "gemini_open": open_gemini,
        "title": built.get("title") or "",
        "slot": built.get("slot") or "",
        "anchor": built.get("anchor") or "",
    }


def _main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Hold a COLLECTIBLE_PACK on the avatar until it is saved"
    )
    parser.add_argument("command", choices=["import-desktop"])
    parser.add_argument("--data-dir", type=Path, required=True)
    parser.add_argument("--pack", type=Path, required=True)
    parser.add_argument("--desktop", type=Path, default=None)
    args = parser.parse_args(argv)
    raw = args.pack.read_text(encoding="utf-8")
    desktop = args.desktop if args.desktop is not None else args.pack.parent
    result = import_text(args.data_dir, raw, desktop=desktop)
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(_main())
