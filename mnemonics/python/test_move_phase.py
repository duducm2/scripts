"""Phase reorder keeps task identity and renumbers Phase titles."""

from __future__ import annotations

import copy

from study_plans_save import apply_phase_move


def _item(iid, path, text, checked, sort_order, plan_id="PLAN_0001"):
    return {
        "id": iid,
        "plan_id": plan_id,
        "section_path": path,
        "text": text,
        "checked": checked,
        "sort_order": str(sort_order),
    }


def test_move_phase_up_renumbers_and_keeps_checks() -> None:
    items = [
        _item("B", "Backlog", "later idea", "0", 1),
        _item("A1", "Phase 0: First > Topic", "learn A", "1", 2),
        _item("C1", "Phase 1: Second > Topic", "learn C", "0", 3),
        _item("N1", "Progress Tracking", "log hours", "0", 4),
    ]
    resources = [
        {
            "id": "R1",
            "plan_id": "PLAN_0001",
            "section_path": "Phase 1: Second > Topic",
            "line": "https://example.test",
            "sort_order": "1",
        }
    ]
    other = _item("Z", "Phase 9: Untouched", "other plan", "1", 3, plan_id="PLAN_0009")
    items.append(other)
    snapshot = copy.deepcopy(other)

    error = apply_phase_move(items, resources, "PLAN_0001", "Phase 1: Second", "up")
    assert error is None, error

    by_id = {row["id"]: row for row in items}
    assert by_id["A1"]["checked"] == "1"
    assert by_id["C1"]["checked"] == "0"
    assert by_id["A1"]["id"] == "A1"
    assert by_id["C1"]["section_path"] == "Phase 0: Second > Topic"
    assert by_id["A1"]["section_path"] == "Phase 1: First > Topic"
    assert by_id["B"]["section_path"] == "Backlog"
    assert by_id["N1"]["section_path"] == "Progress Tracking"
    assert resources[0]["section_path"] == "Phase 0: Second > Topic"
    assert int(by_id["C1"]["sort_order"]) < int(by_id["A1"]["sort_order"])
    assert int(by_id["B"]["sort_order"]) < int(by_id["C1"]["sort_order"])
    assert int(by_id["A1"]["sort_order"]) < int(by_id["N1"]["sort_order"])
    assert other == snapshot


def test_non_phase_section_stays_between_phase_slots() -> None:
    items = [
        _item("A", "Phase 1: Alpha > T", "a", "0", 1),
        _item("E", "Extra reading", "e", "0", 2),
        _item("B", "Phase 2: Beta > T", "b", "0", 3),
    ]
    error = apply_phase_move(items, [], "PLAN_0001", "Phase 2: Beta", "up")
    assert error is None, error
    by_id = {row["id"]: row for row in items}
    assert by_id["B"]["section_path"] == "Phase 1: Beta > T"
    assert by_id["E"]["section_path"] == "Extra reading"
    assert by_id["A"]["section_path"] == "Phase 2: Alpha > T"
    assert (
        int(by_id["B"]["sort_order"])
        < int(by_id["E"]["sort_order"])
        < int(by_id["A"]["sort_order"])
    )


def test_edge_and_missing_phase() -> None:
    items = [_item("A", "Phase 1: Only", "a", "0", 1)]
    assert (
        apply_phase_move(items, [], "PLAN_0001", "Phase 1: Only", "up")
        == "already at edge"
    )
    assert (
        apply_phase_move(items, [], "PLAN_0001", "Phase 4: Missing", "down")
        == "phase not found"
    )
    assert items[0]["section_path"] == "Phase 1: Only"
