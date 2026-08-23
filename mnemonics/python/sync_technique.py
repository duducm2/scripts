"""One-way mirror of notes/studies/technique into mnemonics/technique/."""

from __future__ import annotations

import argparse
import json
import shutil
import sys
from pathlib import Path
from typing import Any


SKIP_NAMES = {".git", "__pycache__", ".DS_Store", "Thumbs.db"}
MANIFEST_NAME = ".sync_manifest.json"


def _should_skip_name(name: str) -> bool:
    return name in SKIP_NAMES or name.startswith(".")


def _file_mtime_ns(path: Path) -> int:
    return path.stat().st_mtime_ns


def load_manifest(dest: Path) -> dict[str, Any]:
    path = dest / MANIFEST_NAME
    if not path.exists():
        return {"source": "", "files": {}}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {"source": "", "files": {}}


def save_manifest(dest: Path, data: dict[str, Any]) -> None:
    dest.mkdir(parents=True, exist_ok=True)
    (dest / MANIFEST_NAME).write_text(
        json.dumps(data, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )


def iter_source_files(source: Path) -> list[Path]:
    files: list[Path] = []
    for path in source.rglob("*"):
        if not path.is_file():
            continue
        if any(_should_skip_name(p) for p in path.relative_to(source).parts):
            continue
        files.append(path)
    return sorted(files)


def resolve_technique_source(
    technique_source: Path | None,
    notes_root: Path | None,
) -> Path | None:
    if technique_source is not None:
        return technique_source.resolve()
    if notes_root is None:
        return None
    notes_root = notes_root.resolve()
    # notes_root is typically .../notes/studies
    cand = notes_root / "technique"
    if cand.is_dir():
        return cand
    # sibling of studies
    if notes_root.name.lower() == "studies":
        cand2 = notes_root.parent / "studies" / "technique"
        if cand2.is_dir():
            return cand2
    return None


def sync_technique(
    source: Path,
    dest: Path,
    force: bool = False,
) -> tuple[int, int]:
    """Copy source → dest. Returns (copied_count, skipped_count)."""
    if not source.is_dir():
        raise FileNotFoundError(f"Technique source not found: {source}")

    dest.mkdir(parents=True, exist_ok=True)
    manifest = load_manifest(dest)
    prev_files: dict[str, Any] = manifest.get("files") or {}
    new_files: dict[str, Any] = {}
    copied = 0
    skipped = 0

    for src in iter_source_files(source):
        rel = src.relative_to(source).as_posix()
        mtime = _file_mtime_ns(src)
        size = src.stat().st_size
        dst = dest / Path(rel)
        prev = prev_files.get(rel) or {}
        unchanged = (
            not force
            and dst.exists()
            and prev.get("mtime_ns") == mtime
            and prev.get("size") == size
        )
        if unchanged:
            skipped += 1
        else:
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dst)
            copied += 1
        new_files[rel] = {"mtime_ns": mtime, "size": size}

    # Remove dest files that no longer exist in source (except manifest)
    for orphan in list(dest.rglob("*")):
        if not orphan.is_file():
            continue
        if orphan.name == MANIFEST_NAME:
            continue
        rel = orphan.relative_to(dest).as_posix()
        if rel not in new_files:
            orphan.unlink(missing_ok=True)

    save_manifest(
        dest,
        {
            "source": str(source.resolve()),
            "files": new_files,
        },
    )
    return copied, skipped


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="Mirror notes technique docs into scripts repo")
    p.add_argument("--source", type=Path, required=True)
    p.add_argument("--dest", type=Path, required=True)
    p.add_argument("--force", action="store_true")
    args = p.parse_args(argv)

    try:
        copied, skipped = sync_technique(
            args.source.resolve(),
            args.dest.resolve(),
            force=args.force,
        )
    except FileNotFoundError as e:
        print(str(e), file=sys.stderr)
        return 1
    except OSError as e:
        print(f"Sync failed: {e}", file=sys.stderr)
        return 1

    print(f"Technique sync: {copied} copied, {skipped} skipped -> {args.dest.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
