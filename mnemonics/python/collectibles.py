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
ANIMS = ("bob", "sway", "flicker", "orbit", "float")
ANCHORS = ("head", "shoulders", "hands", "feet", "side", "back", "below")

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


def store_path(data_dir: Path) -> Path:
    return Path(data_dir) / "collectibles.json"


def load_items(data_dir: Path) -> list[dict[str, Any]]:
    path = store_path(data_dir)
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
        cleaned = _public_item(row)
        if cleaned:
            items.append(cleaned)
    return items


def save_items(data_dir: Path, items: list[dict[str, Any]]) -> None:
    path = store_path(data_dir)
    path.parent.mkdir(parents=True, exist_ok=True)
    body = {"items": items}
    path.write_text(
        json.dumps(body, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )


def _public_item(row: dict[str, Any]) -> dict[str, Any] | None:
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
    raw: str, palace_ids: set[str]
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


def import_text(data_dir: Path, raw: str) -> dict[str, Any]:
    data_dir = Path(data_dir)
    item, notes, error = repair_item(raw, known_palace_ids(data_dir))
    if error or item is None:
        return {
            "ok": False,
            "error": error or "Could not repair collectible.",
            "notes": notes,
        }
    items = load_items(data_dir)
    for row in items:
        if row.get("slot") == item["slot"]:
            row["equipped"] = False
    saved = {
        "id": "col_" + uuid.uuid4().hex[:12],
        "created_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        **item,
        "equipped": True,
    }
    items.append(saved)
    save_items(data_dir, items)
    return {
        "ok": True,
        "item": _public_item(saved),
        "notes": notes,
        "count": len(items),
    }


def set_equipped(data_dir: Path, item_id: str, equipped: bool) -> dict[str, Any]:
    items = load_items(data_dir)
    target = next((row for row in items if row.get("id") == item_id), None)
    if target is None:
        return {"ok": False, "error": "No collectible with that id."}
    if equipped:
        for row in items:
            if row.get("slot") == target.get("slot"):
                row["equipped"] = row.get("id") == item_id
    else:
        target["equipped"] = False
    save_items(data_dir, items)
    return {"ok": True, "items": items}


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
    prompt = _prompt_text(
        palace_id=palace_id,
        study_title=study_title,
        title=title,
        character=character,
        beasts=beast_names[:8],
        keywords=keywords[:24],
    )
    return {"ok": True, "prompt": prompt, "palace_id": palace_id, "title": title}


def _prompt_text(
    *,
    palace_id: str,
    study_title: str,
    title: str,
    character: str,
    beasts: list[str],
    keywords: list[str],
) -> str:
    beast_line = ", ".join(beasts) if beasts else "(none)"
    keyword_line = ", ".join(keywords) if keywords else "(none)"
    character_line = character or "(none)"
    return (
        "Invent ONE wearable collectible for my Memory Palace walker.\n"
        "It must be thematically tied to this palace, the way a brick pet or a logo shirt "
        "would belong to a Databricks palace. Do not invent a generic fantasy item.\n\n"
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
        '  "slot": "pet",\n'
        '  "name": "Short relic name",\n'
        '  "blurb": "One sentence on why it belongs to this palace.",\n'
        '  "svg": "<svg xmlns=\\"http://www.w3.org/2000/svg\\" viewBox=\\"0 0 64 64\\">...</svg>",\n'
        '  "anim": "bob",\n'
        '  "anchor": "side"\n'
        "}\n"
        "===END_FILE===\n\n"
        "Rules:\n"
        f"- slot is one of: {', '.join(SLOTS)}\n"
        f"- anim is one of: {', '.join(ANIMS)}\n"
        f"- anchor is one of: {', '.join(ANCHORS)} "
        "(head, shoulders, hands, feet, side for a pet, back for a cape, below for a mount)\n"
        "- svg is one small illustration, under 6000 characters, no scripts, no external images.\n"
        "- Use the palace's own objects, colors, and names. Replace the example values; "
        "keep palace_id exactly as given.\n"
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
    }


def _main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Import a COLLECTIBLE_PACK into collectibles.json"
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
