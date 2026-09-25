"""Generate consolidated Quick Recall.md (smartphone dashboard of recent atoms)."""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent))

from data_aggregator import load_all  # noqa: E402
from schemas import bold_keyword_terms, keyword_display_terms  # noqa: E402
from beast_thumb_base import beast_thumb_md_image, short_label  # noqa: E402

STATE_NAME = "quick_recall.json"
OUT_NAME = "Quick Recall.md"
ATOM_CAP = 100
SEED_STUDIES: list[tuple[str, int]] = [
    ("STUDY_DATAANALYST", 7),
    ("STUDY_PIANO", 4),
]

_WS_RE = re.compile(r"\s+")


def state_path(data_dir: Path) -> Path:
    return data_dir / STATE_NAME


def out_path(output_dir: Path) -> Path:
    return output_dir / OUT_NAME


def load_state(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {"included_palace_ids": [], "known_palace_ids": []}
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {"included_palace_ids": [], "known_palace_ids": []}
    included = [
        str(x).strip() for x in (raw.get("included_palace_ids") or []) if str(x).strip()
    ]
    known = [
        str(x).strip() for x in (raw.get("known_palace_ids") or []) if str(x).strip()
    ]
    return {"included_palace_ids": included, "known_palace_ids": known}


def save_state(path: Path, state: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "included_palace_ids": list(state.get("included_palace_ids") or []),
        "known_palace_ids": list(state.get("known_palace_ids") or []),
    }
    path.write_text(
        json.dumps(payload, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )


def _palace_num(p: dict[str, str]) -> int:
    try:
        return int(p.get("palace_number") or 0)
    except ValueError:
        return 0


def _sort_int(row: dict[str, str], key: str = "sort_order") -> int:
    try:
        return int(row.get(key) or 0)
    except ValueError:
        return 0


def collapse_concept(raw: str | None) -> str:
    return _WS_RE.sub(
        " ", (raw or "").replace("\r\n", "\n").replace("\r", "\n")
    ).strip()


def atom_count_for_palace(
    palace_id: str,
    beasts_by_palace: dict[str, list[dict[str, str]]],
    atoms_by_beast: dict[str, list[dict[str, str]]],
) -> int:
    n = 0
    for b in beasts_by_palace.get(palace_id, []):
        n += len(atoms_by_beast.get(b["id"], []))
    return n


def build_indexes(
    data: dict[str, list[dict[str, str]]],
) -> tuple[
    dict[str, dict[str, str]],
    dict[str, dict[str, str]],
    dict[str, list[dict[str, str]]],
    dict[str, list[dict[str, str]]],
]:
    studies = {s["id"]: s for s in data["studies"] if s.get("id")}
    palaces = {p["id"]: p for p in data["palaces"] if p.get("id")}
    beasts_by_palace: dict[str, list[dict[str, str]]] = {}
    for b in data["beasts"]:
        pid = (b.get("palace_id") or "").strip()
        if not pid:
            continue
        beasts_by_palace.setdefault(pid, []).append(b)
    for pid in beasts_by_palace:
        beasts_by_palace[pid].sort(key=_sort_int)

    atoms_by_beast: dict[str, list[dict[str, str]]] = {}
    for a in data["atoms"]:
        bid = (a.get("beast_id") or "").strip()
        if not bid:
            continue
        atoms_by_beast.setdefault(bid, []).append(a)
    for bid in atoms_by_beast:
        atoms_by_beast[bid].sort(key=_sort_int)

    return studies, palaces, beasts_by_palace, atoms_by_beast


def seed_included(palaces: dict[str, dict[str, str]]) -> list[str]:
    """Newest-first: DA top N, then Piano top M (interleaved by seed order)."""
    by_study: dict[str, list[dict[str, str]]] = {}
    for p in palaces.values():
        sid = (p.get("study_id") or "").strip()
        if not sid:
            continue
        by_study.setdefault(sid, []).append(p)
    for sid in by_study:
        by_study[sid].sort(key=_palace_num, reverse=True)

    included: list[str] = []
    seen: set[str] = set()
    for study_id, count in SEED_STUDIES:
        for p in by_study.get(study_id, [])[:count]:
            pid = p["id"]
            if pid in seen:
                continue
            seen.add(pid)
            included.append(pid)
    return included


def update_membership(
    state: dict[str, Any],
    palaces: dict[str, dict[str, str]],
    beasts_by_palace: dict[str, list[dict[str, str]]],
    atoms_by_beast: dict[str, list[dict[str, str]]],
) -> dict[str, Any]:
    csv_ids = list(palaces.keys())
    csv_set = set(csv_ids)
    included = [
        pid for pid in (state.get("included_palace_ids") or []) if pid in csv_set
    ]
    known = list(state.get("known_palace_ids") or [])
    known_set = set(known)

    first_run = not included and not known_set
    if first_run:
        included = seed_included(palaces)
        known = list(csv_ids)
        known_set = set(known)
    else:
        # Newly imported = present in CSV but never seen before
        new_ids = [pid for pid in csv_ids if pid not in known_set]
        new_ids.sort(
            key=lambda pid: (_palace_num(palaces[pid]), pid),
            reverse=True,
        )
        included_set = set(included)
        prepend = [pid for pid in new_ids if pid not in included_set]
        included = prepend + included
        for pid in csv_ids:
            if pid not in known_set:
                known.append(pid)
                known_set.add(pid)

    # Drop deleted-from-CSV ids already filtered; also drop orphans from known
    known = [pid for pid in known if pid in csv_set]

    def total_atoms(ids: list[str]) -> int:
        return sum(
            atom_count_for_palace(pid, beasts_by_palace, atoms_by_beast) for pid in ids
        )

    while included and total_atoms(included) > ATOM_CAP:
        included.pop()  # oldest = end of newest-first list

    return {"included_palace_ids": included, "known_palace_ids": known}


def render_atom_line(beast: dict[str, str], atom: dict[str, str]) -> str:
    peg = (beast.get("peg_code") or "").strip()
    name = short_label(beast.get("beast_name") or "")
    concept = collapse_concept(atom.get("concept"))
    if concept:
        concept = bold_keyword_terms(
            concept, keyword_display_terms(atom.get("keywords"))
        )
    parts: list[str] = []
    img = beast_thumb_md_image(name, code=peg or None, from_dir="output", width=44)
    if img:
        parts.append(img)
    else:
        parts.append("🟧")
    if peg:
        parts.append(f"[{peg}]")
    if name:
        parts.append(name)
    if concept:
        parts.append(concept)
    return " ".join(parts)


def render_markdown(
    included: list[str],
    studies: dict[str, dict[str, str]],
    palaces: dict[str, dict[str, str]],
    beasts_by_palace: dict[str, list[dict[str, str]]],
    atoms_by_beast: dict[str, list[dict[str, str]]],
) -> str:
    # Study order: first appearance in included (newest-import study first)
    study_order: list[str] = []
    by_study: dict[str, list[str]] = {}
    for pid in included:
        p = palaces.get(pid)
        if not p:
            continue
        sid = (p.get("study_id") or "").strip()
        if not sid:
            continue
        if sid not in by_study:
            by_study[sid] = []
            study_order.append(sid)
        by_study[sid].append(pid)

    lines: list[str] = ["# Quick Recall", ""]

    for sid in study_order:
        pids = by_study[sid]
        # Within study: newest palace_number first
        pids = sorted(pids, key=lambda x: _palace_num(palaces[x]), reverse=True)
        study = studies.get(sid) or {}
        title = (study.get("title") or sid).strip() or sid
        n_palaces = len(pids)
        n_atoms = sum(
            atom_count_for_palace(pid, beasts_by_palace, atoms_by_beast) for pid in pids
        )
        palace_label = "palace" if n_palaces == 1 else "palaces"
        atom_label = "atom" if n_atoms == 1 else "atoms"
        lines.append("<details open>")
        lines.append(
            f"<summary><strong>{title}</strong> · {n_palaces} {palace_label}"
            f" · {n_atoms} {atom_label}</summary>"
        )
        lines.append("")

        for i, pid in enumerate(pids):
            if i > 0:
                lines.append("")
            for beast in beasts_by_palace.get(pid, []):
                for atom in atoms_by_beast.get(beast["id"], []):
                    lines.append(render_atom_line(beast, atom))
        lines.append("")

        lines.append("</details>")
        lines.append("")

    return "\n".join(lines).rstrip() + "\n"


def sync_quick_recall(data_dir: Path, output_dir: Path, dry_run: bool = False) -> Path:
    data = load_all(data_dir)
    studies, palaces, beasts_by_palace, atoms_by_beast = build_indexes(data)
    st_path = state_path(data_dir)
    state = load_state(st_path)
    state = update_membership(state, palaces, beasts_by_palace, atoms_by_beast)
    md = render_markdown(
        state["included_palace_ids"],
        studies,
        palaces,
        beasts_by_palace,
        atoms_by_beast,
    )
    md_path = out_path(output_dir)
    if dry_run:
        print(f"[dry-run] would write {md_path}")
        print(f"[dry-run] would write {st_path}")
        print(
            f"[dry-run] included={len(state['included_palace_ids'])} "
            f"known={len(state['known_palace_ids'])}"
        )
        return md_path

    output_dir.mkdir(parents=True, exist_ok=True)
    md_path.write_text(md, encoding="utf-8")
    save_state(st_path, state)
    n_atoms = sum(
        atom_count_for_palace(pid, beasts_by_palace, atoms_by_beast)
        for pid in state["included_palace_ids"]
    )
    print(
        f"Wrote {md_path} "
        f"({len(state['included_palace_ids'])} palaces, {n_atoms} atoms)"
    )
    return md_path


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(
        description="Sync Quick Recall.md membership + Markdown"
    )
    p.add_argument("--data-dir", type=Path, required=True)
    p.add_argument("--output-dir", type=Path, required=True)
    p.add_argument("--dry-run", action="store_true")
    args = p.parse_args(argv)
    sync_quick_recall(
        args.data_dir.resolve(),
        args.output_dir.resolve(),
        dry_run=args.dry_run,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
