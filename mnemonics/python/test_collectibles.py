"""Repair and import rules for palace collectibles."""

from __future__ import annotations

import json
from pathlib import Path

from collectibles import import_text, repair_item, sanitize_svg

PALACE = "pal_brick"
SVG = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64"><rect width="64" height="64" fill="#c45"/></svg>'


def _library(tmp: Path) -> None:
    (tmp / "palaces.csv").write_text(
        "id,study_id,title\n" + PALACE + ",STUDY_DA,Databricks\n",
        encoding="utf-8",
    )


def _pack(body: str) -> str:
    return "===FILE: COLLECTIBLE.json===\n" + body + "\n===END_FILE===\n"


def test_maps_unknown_slot_and_strips_fence(tmp_path: Path) -> None:
    _library(tmp_path)
    raw = "```json\n" + _pack(
        json.dumps(
            {
                "palace_id": PALACE,
                "slot": "logo-shirt",
                "name": "Brick pup",
                "blurb": "A pet made of bricks.",
                "svg": SVG,
                "anim": "wiggle",
                "anchor": "orbit",
            }
        )
        + "\n```"
    )
    item, notes, error = repair_item(raw, {PALACE})
    assert error == ""
    assert item is not None
    assert item["slot"] == "accessory"
    assert item["anim"] == "bob"
    assert item["anchor"] == "side"
    assert any("accessory" in note for note in notes)


def test_trailing_comma_and_script_are_repaired() -> None:
    payload = {
        "palace_id": PALACE,
        "slot": "hat",
        "name": "Catalog cap",
        "svg": '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64"><script>alert(1)</script><rect onclick="alert(1)" width="10" height="10"/></svg>',
        "anim": "bob",
        "anchor": "head",
    }
    text = json.dumps(payload)
    text = text[:-1] + ",}"
    item, _notes, error = repair_item(_pack(text), {PALACE})
    assert error == ""
    assert item is not None
    assert "<script" not in item["svg"].lower()
    assert "onclick" not in item["svg"].lower()
    assert "<svg" in item["svg"].lower()


def test_missing_svg_is_an_error() -> None:
    item, _notes, error = repair_item(
        _pack(json.dumps({"palace_id": PALACE, "slot": "pet", "name": "Pup", "svg": ""})),
        {PALACE},
    )
    assert item is None
    assert "svg" in error


def test_unknown_palace_is_an_error() -> None:
    _item, _notes, error = repair_item(
        _pack(json.dumps({"palace_id": "nope", "slot": "pet", "name": "Pup", "svg": SVG})),
        {PALACE},
    )
    assert "palace_id" in error


def test_import_equips_new_item_and_clears_the_slot(tmp_path: Path) -> None:
    _library(tmp_path)
    first = import_text(
        tmp_path,
        _pack(
            json.dumps(
                {
                    "palace_id": PALACE,
                    "slot": "pet",
                    "name": "First pup",
                    "svg": SVG,
                    "anim": "bob",
                    "anchor": "side",
                }
            )
        ),
    )
    second = import_text(
        tmp_path,
        _pack(
            json.dumps(
                {
                    "palace_id": PALACE,
                    "slot": "pet",
                    "name": "Brick pup",
                    "svg": SVG,
                    "anim": "float",
                    "anchor": "side",
                }
            )
        ),
    )
    assert first["ok"] and second["ok"]
    saved = json.loads((tmp_path / "collectibles.json").read_text(encoding="utf-8"))
    pets = [row for row in saved["items"] if row["slot"] == "pet"]
    assert len(pets) == 2
    equipped = [row for row in pets if row["equipped"]]
    assert len(equipped) == 1
    assert equipped[0]["name"] == "Brick pup"


def test_sanitize_svg_drops_javascript_url() -> None:
    cleaned = sanitize_svg(
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10">'
        '<a href="javascript:alert(1)"><rect width="10" height="10"/></a></svg>'
    )
    assert "javascript" not in cleaned.lower()
    assert "<svg" in cleaned.lower()
