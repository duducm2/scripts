"""One knowledge atom per beast."""

from palace_store import PalaceStore
from schemas import collapse_beast_atoms, validate_beast_atoms


def test_collapse_folds_the_second_atom_into_notes() -> None:
    beasts = [{"id": "B1", "is_smashed": "1", "beast_notes": "Kept."}]
    atoms = [
        {
            "id": "A1",
            "beast_id": "B1",
            "kind": "zoned",
            "zone": "Z1",
            "zone_label": "Head",
            "concept": "one",
            "keywords": "k | w",
            "quote": "q1",
            "story": "s1",
            "sort_order": "1",
        },
        {
            "id": "A2",
            "beast_id": "B1",
            "kind": "zoned",
            "zone": "Z2",
            "zone_label": "Paw",
            "concept": "two",
            "keywords": "a | b",
            "quote": "q2",
            "story": "s2",
            "sort_order": "2",
        },
    ]
    beasts_out, atoms_out = collapse_beast_atoms(beasts, atoms)
    assert len(atoms_out) == 1
    assert atoms_out[0]["id"] == "A1"
    assert atoms_out[0]["kind"] == "single"
    assert atoms_out[0]["zone"] == ""
    assert atoms_out[0]["zone_label"] == ""
    assert atoms_out[0]["concept"] == "one"
    assert atoms_out[0]["quote"] == "q1"
    notes = beasts_out[0]["beast_notes"]
    assert notes.startswith("Kept.")
    assert "[Z2 · Paw]" in notes
    assert "Concept: two" in notes
    assert "Keywords: a | b" in notes
    assert "Quote: q2" in notes
    assert "Story: s2" in notes
    assert beasts_out[0]["is_smashed"] == "0"


def test_validate_rejects_a_second_or_zoned_atom() -> None:
    assert validate_beast_atoms([{"kind": "single"}, {"kind": "single"}])
    assert validate_beast_atoms([{"kind": "zoned"}])
    assert validate_beast_atoms([{"kind": "single"}]) is None


def test_second_atom_save_is_rejected(tmp_path) -> None:
    existing = {
        "id": "ATOM_0001",
        "beast_id": "BEAST_0001",
        "kind": "single",
        "zone": "",
        "zone_label": "",
        "concept": "one",
        "keywords": "",
        "quote": "",
        "story": "",
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
        data,
        data["atoms"],
        {"beast_id": "BEAST_0001", "concept": "two", "kind": "zoned"},
        None,
    )
    assert result["ok"] is False
    assert "one Knowledge Atom" in result["error"]
