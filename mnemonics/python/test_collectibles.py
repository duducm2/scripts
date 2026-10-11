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
                "name": "Brick cup",
                "blurb": "A trophy made of bricks.",
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
    assert item["slot"] == "trophy"
    assert item["anim"] == "bob"
    assert item["anchor"] == "pedestal"
    assert any("trophy" in note for note in notes)


def test_trailing_comma_and_script_are_repaired() -> None:
    payload = {
        "palace_id": PALACE,
        "slot": "trophy",
        "name": "Catalog cup",
        "svg": '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 256 256"><script>alert(1)</script><rect onclick="alert(1)" width="10" height="10"/></svg>',
        "anim": "bob",
        "anchor": "pedestal",
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


def test_import_keeps_each_trophy_without_writing_the_library(tmp_path: Path) -> None:
    _library(tmp_path)
    first = import_text(
        tmp_path,
        _pack(
            json.dumps(
                {
                    "palace_id": PALACE,
                    "slot": "trophy",
                    "name": "First cup",
                    "svg": SVG,
                    "anim": "bob",
                    "anchor": "pedestal",
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
                    "slot": "trophy",
                    "name": "Brick cup",
                    "svg": SVG,
                    "anim": "float",
                    "anchor": "pedestal",
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
    names = [row["name"] for row in held["items"]]
    assert names == ["First cup", "Brick cup"]
    assert all(row["slot"] == "trophy" and row["equipped"] is True for row in held["items"])


def test_save_is_refused(tmp_path: Path) -> None:
    _library(tmp_path)
    imported = import_text(
        tmp_path,
        _pack(
            json.dumps(
                {
                    "palace_id": PALACE,
                    "slot": "trophy",
                    "name": "Catalog cup",
                    "svg": SVG,
                    "anim": "bob",
                    "anchor": "pedestal",
                }
            )
        ),
    )
    result = persist_item(tmp_path, imported["item"]["id"])
    assert result["ok"] is False
    assert (tmp_path / "collectibles_session.json").is_file()
    assert not (tmp_path / "repl" / "collectibles.json").exists()


def test_removing_a_trophy_deletes_it(tmp_path: Path) -> None:
    _library(tmp_path)
    imported = import_text(
        tmp_path,
        _pack(
            json.dumps(
                {
                    "palace_id": PALACE,
                    "slot": "trophy",
                    "name": "Brick cup",
                    "svg": SVG,
                    "anim": "float",
                    "anchor": "pedestal",
                }
            )
        ),
    )
    result = remove_item(tmp_path, imported["item"]["id"])
    assert result["ok"] and result["deleted"] is True
    assert result["items"] == []
    assert not (tmp_path / "collectibles_session.json").exists()
    assert not (tmp_path / "repl" / "collectibles.json").exists()


def test_a_thirteenth_trophy_drops_the_oldest(tmp_path: Path) -> None:
    _library(tmp_path)
    for index in range(13):
        result = import_text(
            tmp_path,
            _pack(
                json.dumps(
                    {
                        "palace_id": PALACE,
                        "slot": "trophy",
                        "name": f"Cup {index}",
                        "svg": SVG,
                        "anim": "bob",
                        "anchor": "pedestal",
                    }
                )
            ),
        )
        assert result["ok"]
    held = json.loads(
        (tmp_path / "collectibles_session.json").read_text(encoding="utf-8")
    )
    names = [row["name"] for row in held["items"]]
    assert names == [f"Cup {index}" for index in range(1, 13)]
    assert result["count"] == 12
    assert not (tmp_path / "repl" / "collectibles.json").exists()


def test_prompt_asks_for_one_trophy(tmp_path: Path) -> None:
    _library(tmp_path)
    first = build_prompt(tmp_path, PALACE)
    second = build_prompt(tmp_path, PALACE)
    assert first["ok"] and first["slot"] == "trophy" and first["anchor"] == "pedestal"
    assert '"slot": "trophy"' in first["prompt"]
    assert "Databricks" in first["prompt"]
    assert "garment" in first["prompt"]
    assert second["slot"] == "trophy"
    saved = json.loads(
        (tmp_path / "collectible_requests.json").read_text(encoding="utf-8")
    )
    assert saved["pending_slot"] == "trophy"
    assert saved["cycle"] == ["trophy"]


def test_import_stores_the_requested_trophy_when_the_pack_says_pet(
    tmp_path: Path,
) -> None:
    _library(tmp_path)
    assert assign_slot(tmp_path, PALACE) == "trophy"
    icon = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 256 256"><rect width="40" height="80" fill="#c45"/></svg>'
    result = import_text(
        tmp_path,
        _pack(
            json.dumps(
                {
                    "palace_id": PALACE,
                    "slot": "pet",
                    "name": "Brick cup",
                    "svg": icon,
                    "anim": "bob",
                    "anchor": "side",
                }
            )
        ),
    )
    assert result["ok"]
    assert result["item"]["slot"] == "trophy"
    assert result["item"]["anchor"] == "pedestal"
    saved = json.loads(
        (tmp_path / "collectible_requests.json").read_text(encoding="utf-8")
    )
    assert saved["pending_slot"] == ""
    assert "trophy" in saved["cycle"]


def test_sanitize_svg_drops_javascript_url() -> None:
    cleaned = sanitize_svg(
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10">'
        '<a href="javascript:alert(1)"><rect width="10" height="10"/></a></svg>'
    )
    assert "javascript" not in cleaned.lower()
    assert "<svg" in cleaned.lower()


def test_body_strip_is_rejected_for_a_trophy() -> None:
    item, _notes, error = repair_item(
        _pack(
            json.dumps(
                {
                    "palace_id": PALACE,
                    "slot": "trophy",
                    "name": "Walk plate",
                    "svg": '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 305 424"></svg>',
                }
            )
        ),
        {PALACE},
    )
    assert item is None
    assert "rejected" in error


def test_trophy_icon_is_accepted_and_a_coarse_grid_is_marked() -> None:
    icon, _notes, error = repair_item(
        _pack(
            json.dumps(
                {
                    "palace_id": PALACE,
                    "slot": "trophy",
                    "name": "Desk cup",
                    "svg": '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 256 256"><rect width="256" height="256" fill="#c45"/></svg>',
                    "anchor": "pedestal",
                }
            )
        ),
        {PALACE},
    )
    assert error == ""
    assert icon is not None
    assert icon["anchor"] == "pedestal"
    from collectibles import validate_strip

    err, note = validate_strip(
        "trophy",
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 256 256"></svg>',
    )
    assert err == "" and note == ""
    coarse_err, coarse_note = validate_strip("trophy", SVG)
    assert coarse_err == "" and "256x256" in coarse_note


def test_odd_icon_size_is_kept_and_marked_misfit() -> None:
    from collectibles import validate_strip

    error, note = validate_strip(
        "trophy",
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 200"></svg>',
    )
    assert error == ""
    assert "misfit" in note
