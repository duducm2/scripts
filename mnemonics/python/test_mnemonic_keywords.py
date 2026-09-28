from __future__ import annotations

from palace_practice_render import render_atom_block_md
from palace_store import PalaceStore
from schemas import (
    concept_bracket_groups,
    embed_keyword_mnemonics,
    iter_keyword_pairs,
    normalize_atom_keywords,
    validate_atom_mnemonics,
)


PARALLEL_CONCEPT = (
    "[Parallel Coordinates] [I draw each variable as a parallel axis] "
    "[and turn each tuple into a polyline]"
)
PARALLEL_KEYWORDS = "fence | Parallel Coordinates || easel | draw || bead | tuple"


def _chip(phrase: str, mnemonic: str) -> str:
    from urllib.parse import quote

    message = quote(f"[{phrase}]", safe="")
    src = (
        "https://img.shields.io/static/v1?style=flat-square"
        f"&label=&message={message}&color=9ecbff"
    )
    return f'<img alt="[{phrase}]" src="{src}" />({mnemonic})'


def test_parallel_coordinates_contract_and_render_order() -> None:
    assert validate_atom_mnemonics(PARALLEL_CONCEPT, PARALLEL_KEYWORDS) is None
    assert embed_keyword_mnemonics(PARALLEL_CONCEPT, PARALLEL_KEYWORDS) == (
        f"{_chip('Parallel Coordinates', 'fence')} "
        f"[I {_chip('draw', 'easel')} each variable as a parallel axis] "
        f"[and turn each {_chip('tuple', 'bead')} into a polyline]"
    )


def test_inline_whole_group_reuses_brackets_and_substring_nests() -> None:
    concept = "[Parallel] [I draw a parallel axis] [as one line]"
    keywords = "fence | Parallel || rails | parallel || pen | line"
    assert embed_keyword_mnemonics(concept, keywords) == (
        f"{_chip('Parallel', 'fence')} "
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


def test_atom_block_omits_keywords_section() -> None:
    text = "\n".join(
        render_atom_block_md(
            {
                "concept": PARALLEL_CONCEPT,
                "keywords": PARALLEL_KEYWORDS,
                "quote": "q",
                "story": "s",
            }
        )
    )
    assert "Keywords" not in text
    assert "No keywords yet" not in text
    assert _chip("Parallel Coordinates", "fence") in text


def test_repeated_term_stays_separate_across_groups() -> None:
    concept = "[Parallel] [I draw a parallel axis] [as one line]"
    keywords = "fence | Parallel || rails | parallel || pen | line"
    assert validate_atom_mnemonics(concept, keywords) is None
    assert len(iter_keyword_pairs(keywords)) == 3


def test_second_pair_on_one_group_fails() -> None:
    two_on_second = (
        "fence | Parallel Coordinates || easel | draw || "
        "axis | parallel axis || bead | tuple"
    )
    assert "exactly one keyword pair" in (
        validate_atom_mnemonics(PARALLEL_CONCEPT, two_on_second) or ""
    )


def test_pair_count_must_equal_group_count() -> None:
    missing_group = "fence | Parallel Coordinates || easel | draw"
    assert "exactly one keyword pair" in (
        validate_atom_mnemonics(PARALLEL_CONCEPT, missing_group) or ""
    )

    wrong_order = "easel | draw || fence | Parallel Coordinates || bead | tuple"
    assert "matching bracket group" in (
        validate_atom_mnemonics(PARALLEL_CONCEPT, wrong_order) or ""
    )


def test_pair_and_group_limits_are_enforced() -> None:
    assert "3–6 bracket groups" in (
        validate_atom_mnemonics("[Name] [I define it]", "tag | Name || die | define")
        or ""
    )
    seven_groups = "[Name] [one] [two] [three] [four] [five] [six]"
    seven_pairs = (
        "tag | Name || 1 | one || 2 | two || 3 | three || "
        "4 | four || 5 | five || 6 | six"
    )
    assert "3–6 bracket groups" in (
        validate_atom_mnemonics(seven_groups, seven_pairs) or ""
    )


def test_note_is_not_a_keyword_source() -> None:
    concept = "[Name] [I define it] [clearly] — Note: supplemental nuance"
    keywords = "tag | Name || die | define || net | nuance"
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
