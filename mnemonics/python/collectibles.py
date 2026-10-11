"""Palace collectibles: schema, repair, Desktop pack import, relic prompt."""

from __future__ import annotations

import argparse
import csv
import json
import os
import re
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

SLOTS = (
    "armor",
    "gloves",
    "pants",
    "shoes",
    "hat",
    "ring",
    "staff",
    "sword",
    "pet",
    "cape",
    "mount",
    "accessory",
)
# Ask for worn gear before companions fall back to pets and accessories.
SLOT_ORDER = (
    "armor",
    "gloves",
    "pants",
    "shoes",
    "hat",
    "ring",
    "staff",
    "sword",
    "cape",
    "mount",
    "pet",
    "accessory",
)
SLOT_ANCHOR = {
    "hat": "head",
    "armor": "shoulders",
    "gloves": "hands",
    "ring": "hand_left",
    "staff": "hand_right",
    "sword": "hand_right",
    "pants": "feet",
    "shoes": "feet",
    "cape": "back",
    "pet": "side",
    "mount": "below",
    "accessory": "side",
}
SLOT_LAYER = {
    "hat": "gear_head",
    "armor": "gear_torso",
    "gloves": "gear_hands_front",
    "pants": "gear_legs",
    "shoes": "gear_legs",
    "cape": "gear_back",
    "ring": "gear_hands_front",
    "staff": "gear_hands_front",
    "sword": "gear_hands_front",
    "pet": "gear_pet",
    "mount": "gear_back",
    "accessory": "gear_accessory",
}
# Armor hides the default shirt. Pants and shoes are standing overlays drawn over the legs.
SLOT_OCCLUDES = {
    "armor": ("body_torso",),
}
# Worn clothing is one still frame the same size as the standing figure.
STRIP_SLOTS = (
    "hat",
    "armor",
    "gloves",
    "pants",
    "shoes",
    "cape",
)
# Held items and companions are small pictures pinned to a named point.
HELD_ICON_SLOTS = (
    "ring",
    "staff",
    "sword",
)
ANCHORED_ICON_SLOTS = (
    "pet",
    "mount",
    "accessory",
    "ring",
    "staff",
    "sword",
)
_VIEWBOX = re.compile(r"""viewBox\s*=\s*["']([^"']+)["']""", re.IGNORECASE)
ANIMS = ("bob", "sway", "flicker", "orbit", "float")
ANCHORS = (
    "head",
    "shoulders",
    "hands",
    "feet",
    "side",
    "back",
    "below",
    "hand_left",
    "hand_right",
)

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
    """Relics currently on the avatar that have not been saved."""
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
    saved = load_saved(data_dir)
    session = load_session(data_dir)
    saved_ids = {row["id"] for row in saved}
    items: list[dict[str, Any]] = []
    for row in saved:
        public = _public_item(row, saved=True)
        if public:
            items.append(public)
    for row in session:
        if row["id"] in saved_ids:
            continue
        public = _public_item(row, saved=False)
        if public:
            items.append(public)
    return items


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
    svg = sanitize_svg(str(row.get("svg") or ""))
    if not svg:
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
        "anim": anim if anim in ANIMS else "bob",
        "anchor": anchor if anchor in ANCHORS else "side",
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
    if len(text) > 6000:
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
        if box is None or box[2] >= 180 or box[3] >= 180:
            return (
                'This slot needs a small icon. Use viewBox "0 0 64 64". '
                "A body strip was rejected.",
                "",
            )
        width, height = box[2], box[3]
        if _near(width, 64, 8) and _near(height, 64, 8):
            return "", ""
        return "", "Icon size does not match 64x64. Marked as misfit."
    return "", ""


def _strip_misfit(slot: str, svg: str) -> bool:
    if slot not in STRIP_SLOTS and slot not in ANCHORED_ICON_SLOTS:
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
    if "slot" in parsed or "svg" in parsed or "name" in parsed:
        return parsed
    return None


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
        return None, notes, "JSON has no collectible object (slot, name, svg)."

    slot = str(item.get("slot") or "").strip().lower()
    if slot not in SLOTS:
        notes.append(f"Unknown slot {slot or '(blank)'} mapped to accessory.")
        slot = "accessory"
    anim = str(item.get("anim") or "").strip().lower()
    if anim not in ANIMS:
        notes.append(f"Unknown anim {anim or '(blank)'} mapped to bob.")
        anim = "bob"
    anchor = str(item.get("anchor") or "").strip().lower()
    if anchor not in ANCHORS:
        notes.append(f"Unknown anchor {anchor or '(blank)'} mapped to side.")
        anchor = "side"
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
    if not svg:
        return None, notes, "svg is missing, unsafe, or not a small <svg> illustration."
    if check_art:
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
    empty = {"pending_slot": "", "pending_palace_id": "", "cycle": []}
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
    return {
        "pending_slot": pending if pending in SLOT_ORDER else "",
        "pending_palace_id": str(raw.get("pending_palace_id") or "").strip(),
        "cycle": [str(slot).strip().lower() for slot in cycle],
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
        {"pending_slot": "", "pending_palace_id": "", "cycle": cycle},
    )
    return item


def import_text(data_dir: Path, raw: str) -> dict[str, Any]:
    """Hold a generated relic on the avatar. It is not written to REPL storage."""
    data_dir = Path(data_dir)
    item, notes, error = repair_item(raw, known_palace_ids(data_dir), check_art=False)
    if error or item is None:
        return {
            "ok": False,
            "error": error or "Could not repair collectible.",
            "notes": notes,
        }
    item = _lock_imported_slot(data_dir, item, notes)
    strip_error, strip_note = validate_strip(
        str(item.get("slot") or ""), str(item.get("svg") or "")
    )
    if strip_error:
        return {"ok": False, "error": strip_error, "notes": notes}
    if strip_note and strip_note not in notes:
        notes.append(strip_note)
    session = load_session(data_dir)
    saved = load_saved(data_dir)
    slot = item["slot"]
    _unequip_slot(saved, slot)
    session = [row for row in session if row.get("slot") != slot]
    held = {
        "id": "col_" + uuid.uuid4().hex[:12],
        "created_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        **item,
        "equipped": True,
    }
    session.append(held)
    _write_store(session_path(data_dir), session)
    _write_store(repl_path(data_dir), saved)
    return {
        "ok": True,
        "item": _public_item(held, saved=False),
        "notes": notes,
        "count": len(load_items(data_dir)),
        "saved": False,
    }


def persist_item(data_dir: Path, item_id: str) -> dict[str, Any]:
    """Copy one held relic into REPL storage and drop the unsaved copy."""
    data_dir = Path(data_dir)
    item_id = (item_id or "").strip()
    session = load_session(data_dir)
    saved = load_saved(data_dir)
    target = next((row for row in session if row.get("id") == item_id), None)
    if target is None:
        if any(row.get("id") == item_id for row in saved):
            return {"ok": True, "items": load_items(data_dir), "saved": True}
        return {"ok": False, "error": "No collectible with that id."}
    if target.get("equipped"):
        _unequip_slot(saved, str(target.get("slot") or ""), keep_id=item_id)
    saved = [row for row in saved if row.get("id") != item_id]
    stored = _stored_item(target)
    if stored is None:
        return {"ok": False, "error": "That collectible could not be saved."}
    saved.append(stored)
    session = [row for row in session if row.get("id") != item_id]
    _write_store(repl_path(data_dir), saved)
    _write_store(session_path(data_dir), session)
    return {"ok": True, "items": load_items(data_dir), "saved": True}


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
        session = [row for row in session if row.get("id") != item_id]
        _write_store(session_path(data_dir), session)
        return {"ok": True, "items": load_items(data_dir), "deleted": True}
    slot = str((target or saved_target or {}).get("slot") or "")
    if equipped:
        session = [
            row
            for row in session
            if row.get("slot") != slot or row.get("id") == item_id
        ]
        _unequip_slot(saved, slot, keep_id=item_id)
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
    prompt = _prompt_text(
        palace_id=palace_id,
        study_title=study_title,
        title=title,
        character=character,
        beasts=beast_names[:8],
        keywords=keywords[:24],
        slot=slot,
        anchor=anchor,
    )
    return {
        "ok": True,
        "prompt": prompt,
        "palace_id": palace_id,
        "title": title,
        "slot": slot,
        "anchor": anchor,
    }


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
) -> str:
    beast_line = ", ".join(beasts) if beasts else "(none)"
    keyword_line = ", ".join(keywords) if keywords else "(none)"
    character_line = character or "(none)"
    if slot in STRIP_SLOTS:
        draw_rules = (
            "- The walker stands still in a three-quarter T-pose, turned slightly toward the viewer and facing right. Both eyes are visible. Arms are straight out. Do not move or animate the character.\n"
            "- anim moves only this relic. The body stays a fixed puzzle.\n"
            "- Draw a 305x424 transparent SVG of the item on that standing figure.\n"
            f"- Cover only the body region for the {slot}. Leave the rest of the canvas empty.\n"
            "- Head near the top, feet together at the bottom, arms extended horizontally.\n"
            '- svg viewBox must be "0 0 305 424". A 64x64 icon will be rejected. '
            "A walk strip will be rejected.\n"
        )
        svg_example = '<svg xmlns=\\"http://www.w3.org/2000/svg\\" viewBox=\\"0 0 305 424\\">...</svg>'
    elif slot in HELD_ICON_SLOTS:
        draw_rules = (
            "- The walker stands still. Do not move or animate the character.\n"
            "- anim moves only this icon, pinned to its anchor.\n"
            f"- Draw a 64x64 transparent SVG of the item. It will be pinned to {anchor}.\n"
            "- Keep the drawing inside the icon. Do not draw the character.\n"
            '- svg viewBox must be "0 0 64 64". A full-body strip will be rejected.\n'
        )
        svg_example = '<svg xmlns=\\"http://www.w3.org/2000/svg\\" viewBox=\\"0 0 64 64\\">...</svg>'
    else:
        draw_rules = (
            "- The walker stands still. Do not move or animate the character.\n"
            "- anim moves only this relic, pinned to its anchor.\n"
            f"- svg is one small illustration pinned to {anchor}, under 6000 characters.\n"
            "- Keep it tight to that point. Do not let it float away.\n"
            '- viewBox "0 0 64 64" is enough for a pet, a mount, or an accessory.\n'
        )
        svg_example = '<svg xmlns=\\"http://www.w3.org/2000/svg\\" viewBox=\\"0 0 64 64\\">...</svg>'
    return (
        "Invent ONE wearable collectible for my Memory Palace walker.\n"
        "It must be thematically tied to this palace. Do not invent a generic fantasy item.\n"
        f'The category is already chosen. Copy slot "{slot}" and anchor "{anchor}" exactly. '
        "Do not change them.\n\n"
        f"Study: {study_title}\n"
        f"Palace id (copy exactly): {palace_id}\n"
        f"Palace title: {title}\n"
        f"Character: {character_line}\n"
        f"Beasts: {beast_line}\n"
        f"Keywords: {keyword_line}\n\n"
        "Deliver one file named exactly COLLECTIBLE_PACK.txt (download chip, or one marked fence).\n"
        "Never claim you saved it to Desktop. I save the file myself.\n"
        "The pack body is only these markers and one JSON object. No Markdown outside the pack.\n\n"
        "===FILE: COLLECTIBLE.json===\n"
        "{\n"
        f'  "palace_id": "{palace_id}",\n'
        f'  "slot": "{slot}",\n'
        '  "name": "Short relic name",\n'
        '  "blurb": "One sentence on why it belongs to this palace.",\n'
        f'  "svg": "{svg_example}",\n'
        '  "anim": "bob",\n'
        f'  "anchor": "{anchor}"\n'
        "}\n"
        "===END_FILE===\n\n"
        "Rules:\n"
        f'- slot must be exactly "{slot}". Do not pick another category.\n'
        f'- anchor must be exactly "{anchor}". Do not change it.\n'
        f"- anim is one of: {', '.join(ANIMS)}\n"
        "- svg is under 6000 characters, no scripts, no external images.\n"
        f"{draw_rules}"
        "- Use the palace's own objects, colors, and names. Replace the example name, blurb, "
        "svg, and anim. Keep palace_id, slot, and anchor exactly as given.\n"
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
    args = parser.parse_args(argv)
    raw = args.pack.read_text(encoding="utf-8")
    result = import_text(args.data_dir, raw)
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(_main())
