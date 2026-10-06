from __future__ import annotations

import keyword_images
import palace_practice_render
from palace_practice_render import render_atom_block_md
from study_quick_recall_md import render_atom_line
from palace_store import PalaceStore
from schemas import (
    concept_bracket_groups,
    embed_keyword_mnemonics,
    iter_keyword_pairs,
    normalize_atom_keywords,
    validate_atom_mnemonics,
)


JOBS_CONCEPT = (
    "[I use Databricks Jobs to convert] "
    "[my interactive notebooks into scheduled production pipelines.]"
)
JOBS_KEYWORDS = "calendar | Databricks Jobs || notebook | notebooks"


def _chip(phrase: str, mnemonic: str, *, lead: bool = False) -> str:
    from urllib.parse import quote

    message = quote(f"[{phrase}]", safe="")
    color = "9fd4ff" if lead else "ffd966"
    src = (
        "https://img.shields.io/static/v1?style=flat-square"
        f"&label=&message={message}&color={color}"
    )
    return f'<img alt="[{phrase}]" src="{src}" />({mnemonic})'


def test_two_part_contract_and_render_order() -> None:
    assert validate_atom_mnemonics(JOBS_CONCEPT, JOBS_KEYWORDS) is None
    assert (
        validate_atom_mnemonics(
            JOBS_CONCEPT + " — Note: supplemental nuance", JOBS_KEYWORDS
        )
        is None
    )
    assert embed_keyword_mnemonics(JOBS_CONCEPT, JOBS_KEYWORDS) == (
        f"[I use {_chip('Databricks Jobs', 'calendar', lead=True)} to convert] "
        "[my interactive "
        f"{_chip('notebooks', 'notebook')} "
        "into scheduled production pipelines.]"
    )


def test_inline_whole_group_reuses_brackets_and_substring_nests() -> None:
    concept = "[Parallel] [I draw a parallel axis] [as one line]"
    keywords = "fence | Parallel || rails | parallel || pen | line"
    assert embed_keyword_mnemonics(concept, keywords) == (
        f"{_chip('Parallel', 'fence', lead=True)} "
        f"[I draw a {_chip('parallel', 'rails')} axis] "
        f"[as one {_chip('line', 'pen')}]"
    )


def test_note_suffix_is_not_wrapped() -> None:
    concept = "[Name] [I define it] [clearly] — Note: supplemental nuance"
    keywords = "tag | Name || die | define || net | clearly"
    rendered = embed_keyword_mnemonics(concept, keywords)
    core, note = rendered.split(" — Note:", 1)
    assert note == " supplemental nuance"
    assert "<span" not in note
    assert _chip("clearly", "net") in core
    assert "nuance" not in core


def _openverse_row(image_id: str) -> dict:
    return {
        "id": image_id,
        "url": f"https://cdn.example/{image_id}.jpg",
        "thumbnail": f"https://cdn.example/{image_id}-thumb.jpg",
        "title": image_id,
        "foreign_landing_url": f"https://example.com/{image_id}",
    }


def test_search_uses_a_drawing_when_one_exists(monkeypatch) -> None:
    calls: list[str] = []

    def fake_get(url: str) -> dict:
        calls.append(url)
        return {"results": [_openverse_row("drawn")]}

    monkeypatch.setattr(keyword_images, "_get_json", fake_get)
    hits = keyword_images.search_images("pen", 1)
    assert hits[0]["id"] == "drawn"
    assert len(calls) == 1
    assert "category=illustration" in calls[0]


def test_search_falls_back_to_any_picture_of_the_keyword(monkeypatch) -> None:
    calls: list[str] = []

    def fake_get(url: str) -> dict:
        calls.append(url)
        if "category=illustration" in url:
            return {"results": []}
        return {"results": [_openverse_row("photo")]}

    monkeypatch.setattr(keyword_images, "_get_json", fake_get)
    hits = keyword_images.search_images("nail clipper", 1)
    assert [hit["id"] for hit in hits] == ["photo"]
    assert len(calls) == 2
    assert "category=" not in calls[1]


def test_search_fills_a_short_drawing_page_with_other_pictures(monkeypatch) -> None:
    def fake_get(url: str) -> dict:
        if "category=illustration" in url:
            return {"results": [_openverse_row("drawn")]}
        return {"results": [_openverse_row("drawn"), _openverse_row("photo")]}

    monkeypatch.setattr(keyword_images, "_get_json", fake_get)
    hits = keyword_images.search_images("pen", 2)
    assert [hit["id"] for hit in hits] == ["drawn", "photo"]


def test_search_offers_a_real_picture_instead_of_a_broken_title(monkeypatch) -> None:
    def fake_get(url: str) -> dict:
        return {
            "results": [
                {
                    "id": "svg",
                    "url": "https://upload.wikimedia.org/wikipedia/commons/8/84/Map.svg",
                    "thumbnail": "https://api.openverse.org/v1/images/svg/thumb/",
                    "title": "Miami-Dade County Florida Incorporated and Unincorporated areas",
                },
                {
                    "id": "gone",
                    "url": "https://example.com/notes.svg",
                    "thumbnail": "https://api.openverse.org/v1/images/gone/thumb/",
                    "title": "not a picture",
                },
                _openverse_row("photo"),
            ]
        }

    monkeypatch.setattr(keyword_images, "_get_json", fake_get)
    hits = keyword_images.search_images("hammock", 15)
    assert [hit["id"] for hit in hits] == ["svg", "photo"]
    assert hits[0]["thumbnail"].endswith("330px-Map.svg.png")
    assert "api.openverse.org" not in hits[0]["thumbnail"]
    assert (
        keyword_images._extension_for(
            b"<svg xmlns='http://www.w3.org/2000/svg'>", "image/svg+xml"
        )
        is None
    )
    assert (
        keyword_images._extension_for(b"\x89PNG\r\n\x1a\nrest", "image/png") == ".png"
    )


def test_wikimedia_original_uses_a_small_thumbnail() -> None:
    original = "https://upload.wikimedia.org/wikipedia/commons/9/9d/Niger_river_map.svg"
    urls = keyword_images._image_urls(
        {"image": original, "thumbnail": "https://api.openverse.org/v1/images/x/thumb/"}
    )
    assert urls[0].endswith("330px-Niger_river_map.svg.png")
    assert urls[-1] == original


def test_refresh_legacy_replaces_only_unstyled_pictures(monkeypatch, tmp_path) -> None:
    image_dir = tmp_path / "keyword-images"
    image_dir.mkdir()
    manifest = {
        "version": 1,
        "keywords": {
            "pen": {
                "key": "pen",
                "slug": "pen",
                "file": "pen.jpg",
                "version": 1,
                "empty": False,
            },
            "wand": {"key": "wand", "empty": True, "file": ""},
            "apple": {
                "key": "apple",
                "slug": "apple",
                "file": "apple.jpg",
                "version": 2,
                "empty": False,
                "style": "illustration",
            },
        },
    }
    (image_dir / "manifest.json").write_text(__import__("json").dumps(manifest))
    monkeypatch.setattr(keyword_images, "IMAGE_DIR", image_dir)
    monkeypatch.setattr(keyword_images, "MANIFEST_PATH", image_dir / "manifest.json")
    stored: list[str] = []

    monkeypatch.setattr(
        keyword_images,
        "search_images",
        lambda query, page_size=1: [
            {
                "id": "drawn",
                "style": "illustration",
                "image": "http://x",
                "thumbnail": "http://x",
            }
        ],
    )

    def fake_store(key, hit, *, bump):
        stored.append(key)
        data = keyword_images.load_manifest()
        data["keywords"][key]["style"] = hit["style"]
        data["keywords"][key]["version"] = (
            int(data["keywords"][key].get("version") or 1) + 1
        )
        keyword_images.save_manifest(data)
        return data["keywords"][key]

    monkeypatch.setattr(keyword_images, "_store_hit", fake_store)
    result = keyword_images.refresh_legacy()
    assert stored == ["pen"]
    assert result["updated"] == 1
    assert result["remaining"] == 0
    again = keyword_images.refresh_legacy()
    assert again["pending"] == 0


def test_refresh_legacy_keeps_the_file_when_search_is_empty(
    monkeypatch, tmp_path
) -> None:
    image_dir = tmp_path / "keyword-images"
    image_dir.mkdir()
    manifest = {
        "version": 1,
        "keywords": {
            "pen": {
                "key": "pen",
                "slug": "pen",
                "file": "pen.jpg",
                "version": 1,
                "empty": False,
            },
        },
    }
    (image_dir / "manifest.json").write_text(__import__("json").dumps(manifest))
    monkeypatch.setattr(keyword_images, "IMAGE_DIR", image_dir)
    monkeypatch.setattr(keyword_images, "MANIFEST_PATH", image_dir / "manifest.json")
    monkeypatch.setattr(keyword_images, "search_images", lambda query, page_size=1: [])
    monkeypatch.setattr(keyword_images, "_fallback_hits", lambda key: [])
    result = keyword_images.refresh_legacy()
    saved = keyword_images.load_manifest()["keywords"]["pen"]
    assert saved["file"] == "pen.jpg"
    assert saved["style"] == "any"
    assert result["kept"] == 1


def test_practice_markdown_row_uses_cached_files_in_order(
    monkeypatch, tmp_path
) -> None:
    image_dir = tmp_path / "keyword-images"
    image_dir.mkdir()
    (image_dir / "calendar.jpg").write_bytes(b"jpg")
    (image_dir / "notebook.jpg").write_bytes(b"jpg")
    manifest = {
        "version": 1,
        "keywords": {
            "calendar": {"file": "calendar.jpg", "empty": False},
            "notebook": {"file": "notebook.jpg", "empty": False},
            "wand": {"empty": True},
        },
    }
    (image_dir / "manifest.json").write_text(__import__("json").dumps(manifest))
    monkeypatch.setattr(keyword_images, "IMAGE_DIR", image_dir)
    monkeypatch.setattr(keyword_images, "MANIFEST_PATH", image_dir / "manifest.json")

    row = keyword_images.practice_markdown_row(
        "calendar | Databricks Jobs || notebook | notebooks || wand | magic"
    )
    calendar_at = row.index("keyword-images/calendar.jpg")
    notebook_at = row.index("keyword-images/notebook.jpg")
    assert calendar_at < notebook_at
    assert "wand" not in row
    assert 'width="72"' in row


def test_atom_block_places_keyword_pictures_under_the_concept(monkeypatch) -> None:
    monkeypatch.setattr(
        palace_practice_render,
        "practice_markdown_row",
        lambda keywords: '<img src="../../web/assets/keyword-images/calendar.jpg" alt="calendar" width="72" height="72" />',
    )
    text = "\n".join(
        render_atom_block_md(
            {
                "concept": JOBS_CONCEPT,
                "keywords": JOBS_KEYWORDS,
                "quote": "q",
            }
        )
    )
    concept_at = text.index("**Concept**")
    picture_at = text.index("keyword-images/calendar.jpg")
    quote_at = text.index("**Quote**")
    assert concept_at < picture_at < quote_at


def test_quick_recall_line_stays_without_keyword_pictures() -> None:
    line = render_atom_line(
        {"peg_code": "B", "beast_name": "Byron", "beast_source": ""},
        {"concept": JOBS_CONCEPT, "keywords": JOBS_KEYWORDS, "quote": "q"},
    )
    assert "keyword-images" not in line
    assert _chip("Databricks Jobs", "calendar", lead=True) in line


def test_atom_block_omits_keywords_section() -> None:
    text = "\n".join(
        render_atom_block_md(
            {
                "concept": JOBS_CONCEPT,
                "keywords": JOBS_KEYWORDS,
                "quote": "q",
                "story": "s",
            }
        )
    )
    assert "Keywords" not in text
    assert "No keywords yet" not in text
    assert _chip("Databricks Jobs", "calendar", lead=True) in text


def test_repeated_term_stays_separate_across_groups() -> None:
    concept = "[I use Parallel Coordinates to draw] [a parallel axis]"
    keywords = "fence | Parallel Coordinates || rails | parallel"
    assert validate_atom_mnemonics(concept, keywords) is None
    assert len(iter_keyword_pairs(keywords)) == 2


def test_second_pair_on_one_group_fails() -> None:
    extra_pair = (
        "calendar | Databricks Jobs || notebook | notebooks || bead | pipelines"
    )
    assert "exactly one keyword pair" in (
        validate_atom_mnemonics(JOBS_CONCEPT, extra_pair) or ""
    )


def test_pair_count_must_equal_group_count() -> None:
    missing_group = "calendar | Databricks Jobs"
    assert "exactly one keyword pair" in (
        validate_atom_mnemonics(JOBS_CONCEPT, missing_group) or ""
    )

    wrong_order = "notebook | notebooks || calendar | Databricks Jobs"
    assert "matching bracket group" in (
        validate_atom_mnemonics(JOBS_CONCEPT, wrong_order) or ""
    )


def test_pair_and_group_limits_are_enforced() -> None:
    bare_name = "[Name] [I define it] [clearly]"
    assert "exactly 2 bracket groups" in (
        validate_atom_mnemonics(
            bare_name, "tag | Name || die | define || net | clearly"
        )
        or ""
    )
    seven_groups = "[Name] [one] [two] [three] [four] [five] [six]"
    seven_pairs = (
        "tag | Name || 1 | one || 2 | two || 3 | three || "
        "4 | four || 5 | five || 6 | six"
    )
    assert "exactly 2 bracket groups" in (
        validate_atom_mnemonics(seven_groups, seven_pairs) or ""
    )
    assert "I use" in (
        validate_atom_mnemonics(
            "[Databricks Jobs] [I convert notebooks]",
            "calendar | Databricks Jobs || pen | convert",
        )
        or ""
    )
    assert "closing period" in (
        validate_atom_mnemonics(
            "[I use Databricks Jobs to convert.] [my notebooks]",
            "calendar | Databricks Jobs || notebook | notebooks",
        )
        or ""
    )


def test_note_is_not_a_keyword_source() -> None:
    concept = (
        "[I use Databricks Jobs to convert] "
        "[my interactive notebooks into scheduled production pipelines.] "
        "— Note: supplemental nuance"
    )
    keywords = "calendar | Databricks Jobs || net | nuance"
    assert "matching bracket group" in (
        validate_atom_mnemonics(concept, keywords) or ""
    )


def test_nested_citation_brackets_stay_inside_a_group() -> None:
    assert concept_bracket_groups("[Name] [claim [cite: 1]]") == [
        "Name",
        "claim [cite: 1]",
    ]


def test_legacy_keyword_reads_remain_unchanged() -> None:
    legacy = (
        "a | one || b | two || c | three || d | four || "
        "e | five || f | six || g | seven"
    )
    assert normalize_atom_keywords(legacy) == legacy


def test_unrelated_edit_preserves_legacy_keywords_byte_for_byte(tmp_path) -> None:
    legacy_keywords = "fence|Parallel||wire|polyline"
    existing = {
        "id": "ATOM_0001",
        "beast_id": "BEAST_0001",
        "kind": "single",
        "zone": "",
        "zone_label": "",
        "concept": "[Parallel Coordinates: I draw axes] [and connect tuples]",
        "keywords": legacy_keywords,
        "quote": "",
        "story": "old",
        "ipa": "",
        "sort_order": "1",
    }
    data = {
        "atoms": [existing],
        "beasts": [{"id": "BEAST_0001", "palace_id": ""}],
        "palaces": [],
    }
    store = PalaceStore(tmp_path)
    store._save_tree = lambda *_args, **_kwargs: None

    result = store._upsert_atom(
        data, data["atoms"], {"id": "ATOM_0001", "story": "new"}, existing
    )

    assert result["ok"] is True
    assert result["row"]["keywords"] == legacy_keywords
    assert validate_atom_mnemonics(existing["concept"], legacy_keywords) is not None
