"""Repair and import rules for palace collectibles."""

from __future__ import annotations

import json
from pathlib import Path

from collectibles import (
    assign_slot,
    build_prompt,
    import_text,
    persist_item,
    remove_item,
    repair_item,
    sanitize_svg,
)

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
        "svg": '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 305 424"><script>alert(1)</script><rect onclick="alert(1)" width="10" height="10"/></svg>',
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
        _pack(
            json.dumps({"palace_id": PALACE, "slot": "pet", "name": "Pup", "svg": ""})
        ),
        {PALACE},
    )
    assert item is None
    assert "svg" in error


def test_unknown_palace_is_an_error() -> None:
    _item, _notes, error = repair_item(
        _pack(
            json.dumps({"palace_id": "nope", "slot": "pet", "name": "Pup", "svg": SVG})
        ),
        {PALACE},
    )
    assert "palace_id" in error


def test_import_holds_the_new_relic_without_writing_the_library(tmp_path: Path) -> None:
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
    assert first["saved"] is False
    assert not (tmp_path / "repl" / "collectibles.json").exists()
    assert not (tmp_path / "collectibles.json").exists()
    held = json.loads(
        (tmp_path / "collectibles_session.json").read_text(encoding="utf-8")
    )
    pets = [row for row in held["items"] if row["slot"] == "pet"]
    assert len(pets) == 1
    assert pets[0]["name"] == "Brick pup"
    assert pets[0]["equipped"] is True


def test_save_moves_a_relic_into_repl_storage(tmp_path: Path) -> None:
    _library(tmp_path)
    imported = import_text(
        tmp_path,
        _pack(
            json.dumps(
                {
                    "palace_id": PALACE,
                    "slot": "hat",
                    "name": "Catalog cap",
                    "svg": '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 305 424"><rect width="40" height="20" fill="#c45"/></svg>',
                    "anim": "bob",
                    "anchor": "head",
                }
            )
        ),
    )
    item_id = imported["item"]["id"]
    result = persist_item(tmp_path, item_id)
    assert result["ok"]
    assert not (tmp_path / "collectibles_session.json").exists()
    library = json.loads(
        (tmp_path / "repl" / "collectibles.json").read_text(encoding="utf-8")
    )
    assert library["items"][0]["id"] == item_id
    assert library["items"][0]["equipped"] is True
    listed = [row for row in result["items"] if row["id"] == item_id]
    assert listed[0]["saved"] is True


def test_removing_an_unsaved_relic_deletes_it(tmp_path: Path) -> None:
    _library(tmp_path)
    imported = import_text(
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
    result = remove_item(tmp_path, imported["item"]["id"])
    assert result["ok"] and result["deleted"] is True
    assert result["items"] == []
    assert not (tmp_path / "collectibles_session.json").exists()
    assert not (tmp_path / "repl" / "collectibles.json").exists()


def test_taking_off_a_saved_relic_keeps_it_in_the_library(tmp_path: Path) -> None:
    _library(tmp_path)
    imported = import_text(
        tmp_path,
        _pack(
            json.dumps(
                {
                    "palace_id": PALACE,
                    "slot": "cape",
                    "name": "Brick cape",
                    "svg": '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 305 424"><rect width="40" height="20" fill="#c45"/></svg>',
                    "anim": "sway",
                    "anchor": "back",
                }
            )
        ),
    )
    item_id = imported["item"]["id"]
    persist_item(tmp_path, item_id)
    result = remove_item(tmp_path, item_id)
    assert result["ok"] and result["deleted"] is False
    kept = json.loads(
        (tmp_path / "repl" / "collectibles.json").read_text(encoding="utf-8")
    )
    assert kept["items"][0]["id"] == item_id
    assert kept["items"][0]["equipped"] is False


def test_prompt_assigns_armor_before_pet_and_remembers_it(tmp_path: Path) -> None:
    _library(tmp_path)
    first = build_prompt(tmp_path, PALACE)
    second = build_prompt(tmp_path, PALACE)
    assert first["ok"] and first["slot"] == "armor" and first["anchor"] == "shoulders"
    assert '"slot": "armor"' in first["prompt"]
    assert '"slot": "pet"' not in first["prompt"]
    assert "brick pet" not in first["prompt"]
    assert second["slot"] == "gloves"
    saved = json.loads(
        (tmp_path / "collectible_requests.json").read_text(encoding="utf-8")
    )
    assert saved["pending_slot"] == "gloves"
    assert saved["cycle"] == ["armor", "gloves"]


def test_import_stores_the_requested_slot_when_the_pack_says_pet(
    tmp_path: Path,
) -> None:
    _library(tmp_path)
    assert assign_slot(tmp_path, PALACE) == "armor"
    wide = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 305 424"><rect width="40" height="80" fill="#c45"/></svg>'
    result = import_text(
        tmp_path,
        _pack(
            json.dumps(
                {
                    "palace_id": PALACE,
                    "slot": "pet",
                    "name": "Brick pup",
                    "svg": wide,
                    "anim": "bob",
                    "anchor": "side",
                }
            )
        ),
    )
    assert result["ok"]
    assert result["item"]["slot"] == "armor"
    assert result["item"]["anchor"] == "shoulders"
    saved = json.loads(
        (tmp_path / "collectible_requests.json").read_text(encoding="utf-8")
    )
    assert saved["pending_slot"] == ""
    assert "armor" in saved["cycle"]


def test_sanitize_svg_drops_javascript_url() -> None:
    cleaned = sanitize_svg(
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10">'
        '<a href="javascript:alert(1)"><rect width="10" height="10"/></a></svg>'
    )
    assert "javascript" not in cleaned.lower()
    assert "<svg" in cleaned.lower()


def test_small_wearable_icon_is_rejected() -> None:
    item, _notes, error = repair_item(
        _pack(
            json.dumps(
                {
                    "palace_id": PALACE,
                    "slot": "armor",
                    "name": "Tiny plate",
                    "svg": SVG,
                }
            )
        ),
        {PALACE},
    )
    assert item is None
    assert "rejected" in error


def test_walk_strip_is_rejected_for_a_standing_slot() -> None:
    item, _notes, error = repair_item(
        _pack(
            json.dumps(
                {
                    "palace_id": PALACE,
                    "slot": "armor",
                    "name": "Walk plate",
                    "svg": '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 772 424"></svg>',
                }
            )
        ),
        {PALACE},
    )
    assert item is None
    assert "rejected" in error


def test_held_icon_is_accepted_and_a_body_strip_is_not() -> None:
    icon, _notes, error = repair_item(
        _pack(
            json.dumps(
                {
                    "palace_id": PALACE,
                    "slot": "sword",
                    "name": "Desk blade",
                    "svg": '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 256 256"><rect width="256" height="256" fill="#c45"/></svg>',
                    "anchor": "hand_right",
                }
            )
        ),
        {PALACE},
    )
    assert error == ""
    assert icon is not None
    assert icon["anchor"] == "hand_right"
    from collectibles import validate_strip

    err, note = validate_strip(
        "sword",
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 256 256"></svg>',
    )
    assert err == "" and note == ""
    coarse_err, coarse_note = validate_strip("sword", SVG)
    assert coarse_err == "" and "256x256" in coarse_note
    strip, _notes, strip_error = repair_item(
        _pack(
            json.dumps(
                {
                    "palace_id": PALACE,
                    "slot": "sword",
                    "name": "Tall blade",
                    "svg": '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 305 424"></svg>',
                }
            )
        ),
        {PALACE},
    )
    assert strip is None
    assert "rejected" in strip_error


def test_odd_strip_size_is_kept_and_marked_misfit() -> None:
    from collectibles import validate_strip

    error, note = validate_strip(
        "hat",
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 200"></svg>',
    )
    assert error == ""
    assert "misfit" in note
