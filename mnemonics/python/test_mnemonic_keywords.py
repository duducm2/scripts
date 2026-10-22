from __future__ import annotations

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
    "[my interactive notebooks] "
    "[into scheduled] "
    "[production pipelines.]"
)
JOBS_KEYWORDS = (
    "calendar | Databricks Jobs || wand | convert || notebook | notebooks || "
    "clock | scheduled || factory | pipelines"
)


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
        f"[I use {_chip('Databricks Jobs', 'calendar', lead=True)} to "
        f"{_chip('convert', 'wand')}] "
        f"[my interactive {_chip('notebooks', 'notebook')}] "
        f"[into {_chip('scheduled', 'clock')}] "
        f"[production {_chip('pipelines', 'factory')}.]"
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
    assert "keyword-images" not in text
    assert _chip("Databricks Jobs", "calendar", lead=True) in text


def test_repeated_term_stays_separate_across_groups() -> None:
    concept = (
        "[I use Parallel Coordinates to draw] [a parallel axis] "
        "[as one] [straight line]"
    )
    keywords = (
        "fence | Parallel Coordinates || pen | draw || rails | parallel || "
        "bead | one || wire | line"
    )
    assert validate_atom_mnemonics(concept, keywords) is None
    assert len(iter_keyword_pairs(keywords)) == 5


def test_extra_pair_beyond_group_contract_fails() -> None:
    extra_pair = JOBS_KEYWORDS + " || bead | pipelines"
    err = validate_atom_mnemonics(JOBS_CONCEPT, extra_pair) or ""
    assert "group 1 needs two keyword pairs" in err or "5–6 pairs" in err


def test_verb_keyword_is_required() -> None:
    missing_verb = (
        "calendar | Databricks Jobs || notebook | notebooks || "
        "clock | scheduled || factory | pipelines || bead | production"
    )
    assert "infinitive" in (validate_atom_mnemonics(JOBS_CONCEPT, missing_verb) or "")


def test_multi_word_recognizable_word_fails() -> None:
    concept = (
        "[I use a snowflake schema to center] [fact] [tables] [among] [dimensions.]"
    )
    abstract = (
        "flake | snowflake schema || hub | center || maze | complex query joins || "
        "die | fact || chest | tables || tag | dimensions"
    )
    assert "single concrete word" in (validate_atom_mnemonics(concept, abstract) or "")
    ok = (
        "flake | snowflake schema || hub | center || die | fact || "
        "chest | tables || ring | among || tag | dimensions"
    )
    assert validate_atom_mnemonics(concept, ok) is None


def test_pair_count_must_match_group_contract() -> None:
    too_few = "calendar | Databricks Jobs || wand | convert || notebook | notebooks"
    assert "5–6 pairs" in (validate_atom_mnemonics(JOBS_CONCEPT, too_few) or "")

    wrong_order = (
        "notebook | notebooks || calendar | Databricks Jobs || wand | convert || "
        "clock | scheduled || factory | pipelines"
    )
    err = validate_atom_mnemonics(JOBS_CONCEPT, wrong_order) or ""
    assert "formal name" in err or "infinitive" in err


def test_pair_and_group_limits_are_enforced() -> None:
    two_groups = (
        "[I use Databricks Jobs to convert] "
        "[my interactive notebooks into scheduled production pipelines.]"
    )
    assert "4–5 bracket groups" in (
        validate_atom_mnemonics(
            two_groups,
            "calendar | Databricks Jobs || wand | convert || notebook | notebooks || "
            "clock | scheduled || factory | pipelines",
        )
        or ""
    )
    bare_name = "[Name] [I define it] [clearly] [now] [here]"
    assert "I use" in (
        validate_atom_mnemonics(
            bare_name,
            "tag | Name || die | define || net | clearly || now | now || hat | here",
        )
        or ""
    )
    assert "I use" in (
        validate_atom_mnemonics(
            "[Databricks Jobs] [I convert notebooks] [into] [pipelines] [today]",
            "calendar | Databricks Jobs || pen | convert || box | into || "
            "factory | pipelines || sun | today",
        )
        or ""
    )
    assert "closing period" in (
        validate_atom_mnemonics(
            "[I use Databricks Jobs to convert.] [my notebooks] [into] [pipelines] "
            "[today]",
            "calendar | Databricks Jobs || wand | convert || notebook | notebooks || "
            "box | into || factory | pipelines",
        )
        or ""
    )


def test_note_is_not_a_keyword_source() -> None:
    concept = (
        "[I use Databricks Jobs to convert] "
        "[my interactive notebooks] "
        "[into scheduled] "
        "[production pipelines.] "
        "— Note: supplemental nuance"
    )
    keywords = (
        "calendar | Databricks Jobs || wand | convert || notebook | notebooks || "
        "net | nuance || factory | pipelines"
    )
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

