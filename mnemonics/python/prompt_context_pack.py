"""Build attachable CSV packs and Desktop import templates for Memory Palace prompts."""
from __future__ import annotations

import argparse
import csv
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from data_aggregator import load_all  # noqa: E402
from schemas import ATOMS_HEADERS, BEASTS_HEADERS, PALACES_HEADERS  # noqa: E402


def _write(path: Path, headers: list[str], rows: list[dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=headers, extrasaction="ignore")
        w.writeheader()
        for row in rows:
            w.writerow({h: row.get(h, "") for h in headers})


def pack_study(data_dir: Path, study_id: str, out_dir: Path) -> dict[str, str]:
    data = load_all(data_dir)
    palaces = [s for s in data["palaces"] if s.get("study_id") == study_id]
    palace_ids = {s["id"] for s in palaces}
    beasts = [b for b in data["beasts"] if b.get("palace_id") in palace_ids]
    beast_ids = {b["id"] for b in beasts}
    atoms = [a for a in data["atoms"] if a.get("beast_id") in beast_ids]

    out_dir.mkdir(parents=True, exist_ok=True)
    paths = {
        "palaces": out_dir / "palaces.csv",
        "beasts": out_dir / "beasts.csv",
        "atoms": out_dir / "atoms.csv",
    }
    _write(paths["palaces"], PALACES_HEADERS, palaces)
    _write(paths["beasts"], BEASTS_HEADERS, beasts)
    _write(paths["atoms"], ATOMS_HEADERS, atoms)

    # Empty import templates for Desktop workflow
    tpl = out_dir / "PALACE_ATOMS_TEMPLATE.csv"
    _write(tpl, ATOMS_HEADERS, [])
    return {k: str(v) for k, v in paths.items()} | {"template": str(tpl)}


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="Pack study CSV slices for prompt attach")
    p.add_argument("--data-dir", type=Path, required=True)
    p.add_argument("--study-id", type=str, required=True)
    p.add_argument(
        "--out-dir",
        type=Path,
        default=None,
        help="Default: mnemonics/python/packs/<study_id>",
    )
    args = p.parse_args(argv)
    out = args.out_dir or (Path(__file__).resolve().parent / "packs" / args.study_id)
    written = pack_study(args.data_dir.resolve(), args.study_id, out.resolve())
    for k, v in written.items():
        print(f"{k}: {v}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
