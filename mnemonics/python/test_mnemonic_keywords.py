from __future__ import annotations

from palace_practice_render import format_keywords_lines
from palace_store import PalaceStore
from schemas import (
    concept_bracket_groups,
    iter_keyword_pairs,
    normalize_atom_keywords,
    validate_atom_mnemonics,
)


PARALLEL_CONCEPT = (
    "[Parallel Coordinates] [I draw each variable as a parallel axis] "
    "[and turn each tuple into a polyline]"
)
PARALLEL_KEYWORDS = (
    "fence | Parallel Coordinates || easel | draw || "
    "axis | parallel axis || bead | tuple"
)


def test_parallel_coordinates_contract_and_render_order() -> None:
    assert validate_atom_mnemonics(PARALLEL_CONCEPT, PARALLEL_KEYWORDS) is None
    assert format_keywords_lines(PARALLEL_KEYWORDS) == [
        "[**Parallel Coordinates**] → [fence]",
        "[**draw**] → [easel]",
        "[**parallel axis**] → [axis]",
        "[**tuple**] → [bead]",
    ]


def test_repeated_term_stays_separate_across_groups() -> None:
    concept = "[Parallel] [I draw a parallel axis]"
    keywords = "fence | Parallel || easel | draw || rails | parallel"
    assert validate_atom_mnemonics(concept, keywords) is None
    assert len(iter_keyword_pairs(keywords)) == 3


def test_every_group_needs_a_pair_in_source_order() -> None:
    missing_group = (
        "fence | Parallel Coordinates || axis | parallel axis || easel | draw"
    )
    assert "cover every bracket group" in (
        validate_atom_mnemonics(PARALLEL_CONCEPT, missing_group) or ""
    )

    wrong_order = "easel | draw || fence | Parallel Coordinates || bead | tuple"
    assert "source order" in (
        validate_atom_mnemonics(PARALLEL_CONCEPT, wrong_order) or ""
    )


def test_pair_and_group_limits_are_enforced() -> None:
    assert "3–6 pairs" in (
        validate_atom_mnemonics("[Name] [I define it]", "tag | Name || die | define")
        or ""
    )
    seven_groups = "[Name] [one] [two] [three] [four] [five] [six]"
    seven_pairs = (
        "tag | Name || 1 | one || 2 | two || 3 | three || "
        "4 | four || 5 | five || 6 | six"
    )
    assert "at most 6 bracket groups" in (
        validate_atom_mnemonics(seven_groups, seven_pairs) or ""
    )


def test_note_is_not_a_keyword_source() -> None:
    concept = "[Name] [I define it] — Note: supplemental nuance"
    keywords = "tag | Name || die | define || net | nuance"
    assert "cover every bracket group" in (
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
        "sensory": "visual",
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
