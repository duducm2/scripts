"""Unit tests for Knowledge Atom character attribution and justification."""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from palace_practice_render import render_atom_block_md
from study_quick_recall_md import render_atom_line
from schemas import ATOMS_HEADERS, extract_primary_keyword


def test_atoms_headers_contract():
    assert "character" in ATOMS_HEADERS, "character missing from ATOMS_HEADERS"
    assert "justification" in ATOMS_HEADERS, "justification missing from ATOMS_HEADERS"
    assert ATOMS_HEADERS.index("character") < ATOMS_HEADERS.index("justification")
    assert len(ATOMS_HEADERS) == 13, f"Expected 13 headers, got {len(ATOMS_HEADERS)}: {ATOMS_HEADERS}"


def test_extract_primary_keyword():
    kw1 = (
        "calendar | Databricks Jobs || wand | convert || notebook | notebooks || "
        "clock | scheduled || factory | pipelines"
    )
    assert extract_primary_keyword(kw1) == "calendar"

    kw2 = "sword | Excalibur || shield | defend"
    assert extract_primary_keyword(kw2) == "sword"

    assert extract_primary_keyword("") == ""
    assert extract_primary_keyword(None) == ""


def test_full_story_markdown_includes_character_and_justification():
    atom = {
        "concept": "[I use Databricks Jobs to convert] [notebooks] [into scheduled] [pipelines.]",
        "keywords": "calendar | Databricks Jobs || wand | convert || clock | scheduled || factory | pipelines",
        "quote": "Jobs allow scheduled pipelines.",
        "character": "Doctor Strange",
        "justification": "Doctor Strange manipulates timelines and cycles, embodying the scheduled calendar.",
    }
    lines = render_atom_block_md(atom)
    text = "\n".join(lines)

    assert "**Character**" in text, "Missing **Character** header in practice markdown"
    assert "Doctor Strange" in text, "Missing character name in practice markdown"
    assert "**Justification**" in text, "Missing **Justification** header in practice markdown"
    assert "Doctor Strange manipulates timelines" in text, "Missing justification body in practice markdown"


def test_full_story_markdown_omits_empty_character_and_justification():
    atom = {
        "concept": "[I use Databricks Jobs to convert] [notebooks] [into scheduled] [pipelines.]",
        "keywords": "calendar | Databricks Jobs || wand | convert || clock | scheduled || factory | pipelines",
        "quote": "Jobs allow scheduled pipelines.",
        "character": "",
        "justification": "",
    }
    lines = render_atom_block_md(atom)
    text = "\n".join(lines)

    assert "**Character**" not in text
    assert "**Justification**" not in text


def test_quick_recall_markdown_strictly_excludes_character_and_justification():
    beast = {
        "peg_code": "A",
        "beast_name": "Arachne",
        "beast_source": "Lynne Kelly",
    }
    atom = {
        "concept": "Jobs convert notebooks into scheduled pipelines.",
        "keywords": "calendar | Databricks Jobs || wand | convert",
        "character": "Gandalf",
        "justification": "Gandalf possesses deep temporal wisdom for scheduling spells.",
    }
    line = render_atom_line(beast, atom)

    assert "Gandalf" not in line, "Character leaked into quick recall markdown line!"
    assert "temporal wisdom" not in line, "Justification leaked into quick recall markdown line!"
    assert "Arachne" in line, "Beast name should appear in quick recall markdown line"


def test_ahk_headers_match_python():
    ahk_file = Path(__file__).resolve().parent.parent.parent / "Utils" / "mnemonic_palace_helpers.ahk"
    assert ahk_file.exists(), f"Could not find {ahk_file}"
    text = ahk_file.read_text(encoding="utf-8")
    assert '"character"' in text, "character missing from mnemonic_palace_helpers.ahk"
    assert '"justification"' in text, "justification missing from mnemonic_palace_helpers.ahk"


def test_prompts_contain_proximity_and_best_judgment_rules():
    prompts_dir = Path(__file__).resolve().parent.parent / "technique" / "prompts"
    story_p = prompts_dir / "story-prompt.txt"
    reduction_p = prompts_dir / "story-reduction-prompt.txt"

    for p in (story_p, reduction_p):
        assert p.exists(), f"Prompt file {p} does not exist"
        txt = p.read_text(encoding="utf-8")
        assert "±6 Memory Palaces" in txt or "at least six memory palaces" in txt or "6 surrounding memory palaces" in txt, (
            f"Proximity rule missing from {p.name}"
        )
        assert "best" in txt.lower() and "judgment" in txt.lower(), (
            f"Best judgment mandate missing from {p.name}"
        )


def test_proximity_calculation():
    from assign_atom_characters import get_forbidden_nearby_characters

    char_index = {
        "STUDY_A": {
            1: {"gandalf"},
            4: {"frodo"},
            7: {"aragorn"},
            15: {"legolas"},
        }
    }

    # Palace 3 in STUDY_A: within distance 6 of 1, 4, 7 -> gandalf, frodo, aragorn forbidden; legolas allowed (15 - 3 = 12 > 6)
    forbidden = get_forbidden_nearby_characters("STUDY_A", 3, char_index, min_distance=6)
    assert "gandalf" in forbidden
    assert "frodo" in forbidden
    assert "aragorn" in forbidden
    assert "legolas" not in forbidden

    # Palace 14 in STUDY_A: within distance 6 of 15 -> legolas forbidden; 1, 4, 7 allowed (14 - 7 = 7 > 6)
    forbidden_14 = get_forbidden_nearby_characters("STUDY_A", 14, char_index, min_distance=6)
    assert "legolas" in forbidden_14
    assert "gandalf" not in forbidden_14
    assert "frodo" not in forbidden_14


def run_all():
    test_atoms_headers_contract()
    test_extract_primary_keyword()
    test_full_story_markdown_includes_character_and_justification()
    test_full_story_markdown_omits_empty_character_and_justification()
    test_quick_recall_markdown_strictly_excludes_character_and_justification()
    test_ahk_headers_match_python()
    test_prompts_contain_proximity_and_best_judgment_rules()
    test_proximity_calculation()
    print("All 8 test cases passed successfully!")


if __name__ == "__main__":
    run_all()

