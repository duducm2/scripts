"""Generate per-study practice Markdown files with copied palace images."""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent))

from data_aggregator import load_all, resolve_image, save_all, snapshot  # noqa: E402
from palace_practice_render import md_escape, render_palace_section_md  # noqa: E402

PRACTICE_PREFIX = "practice/images/"


def slug_filename(notes_rel_path: str) -> str:
    slug = notes_rel_path.replace("\\", "/").strip("/")
    if not slug or ".." in slug or ":" in slug:
        raise ValueError(f"Unsafe notes_rel_path: {notes_rel_path!r}")
    return slug


def practice_md_path(practice_dir: Path, notes_rel_path: str) -> Path:
    return practice_dir / f"{slug_filename(notes_rel_path)}.md"


def practice_image_dir(practice_dir: Path, notes_rel_path: str) -> Path:
    return practice_dir / "images" / slug_filename(notes_rel_path)


def copy_palace_image(
    palace: dict[str, Any],
    slug: str,
    practice_dir: Path,
    notes_root: Path | None,
    output_dir: Path | None,
) -> str:
    """Copy palace image into practice/images/{slug}/; return MD-relative path or empty."""
    num = str(palace.get("number") or "").strip()
    if not num:
        return ""
    image_rel = (palace.get("image_rel") or "").strip()
    src = resolve_image(notes_root, image_rel, output_dir)
    dest_dir = practice_image_dir(practice_dir, slug)
    dest_dir.mkdir(parents=True, exist_ok=True)

    ext = "png"
    if src and src.exists():
        ext = src.suffix.lstrip(".") or "png"
        dest = dest_dir / f"{num}.{ext}"
        try:
            same = src.resolve() == dest.resolve()
        except OSError:
            same = False
        if not same:
            try:
                shutil.copy2(src, dest)
            except PermissionError:
                # Destination locked (e.g. Drive sync) — keep existing dest if present
                if not dest.exists():
                    raise
        return f"images/{slug_filename(slug)}/{num}.{ext}"

    # Already under practice/images in CSV — file may exist from prior sync
    if image_rel.startswith(PRACTICE_PREFIX):
        rel_tail = image_rel[len(PRACTICE_PREFIX) :]
        existing = practice_dir / "images" / rel_tail.replace("/", "\\")
        if not existing.exists():
            existing = practice_dir / "images" / rel_tail
        if existing.exists():
            slug_part = slug_filename(slug)
            return f"images/{slug_part}/{num}.{existing.suffix.lstrip('.') or ext}"

    return ""


def build_study_markdown(
    study_card: dict[str, Any],
    image_md_paths: dict[str, str],
) -> str:
    title = md_escape(study_card.get("title") or study_card.get("id") or "Study")
    lines: list[str] = [
        f"# {title}",
        "",
    ]
    palaces = study_card.get("palaces") or []
    if not palaces:
        lines.append("_No Memory Palaces yet._")
        lines.append("")
        return "\n".join(lines)

    for i, palace in enumerate(palaces):
        pid = palace.get("id", "")
        lines.extend(
            render_palace_section_md(
                palace,
                image_md_paths.get(pid, ""),
                open_default=(i == 0),
            )
        )

    return "\n".join(lines).rstrip() + "\n"


def write_study(
    study_id: str,
    data_dir: Path,
    output_dir: Path,
    notes_root: Path | None,
    dry_run: bool = False,
) -> Path | None:
    practice_dir = output_dir / "practice"
    practice_dir.mkdir(parents=True, exist_ok=True)

    snap = snapshot(
        data_dir=data_dir,
        notes_root=notes_root,
        output_dir=output_dir,
        study_id=study_id,
    )
    cards = snap.get("studies") or []
    if not cards:
        return None
    study_card = cards[0]
    slug = slug_filename(study_card.get("notes_rel_path") or study_id)

    image_md_paths: dict[str, str] = {}
    for palace in study_card.get("palaces") or []:
        pid = palace.get("id", "")
        rel = copy_palace_image(palace, slug, practice_dir, notes_root, output_dir)
        if rel:
            image_md_paths[pid] = rel

    md_text = build_study_markdown(study_card, image_md_paths)
    out_path = practice_md_path(practice_dir, slug)
    if dry_run:
        print(f"[dry-run] would write {out_path}")
        return out_path
    out_path.write_text(md_text, encoding="utf-8")
    print(f"Wrote {out_path}")
    return out_path


def delete_study(
    notes_rel_path: str,
    output_dir: Path,
    dry_run: bool = False,
) -> None:
    practice_dir = output_dir / "practice"
    slug = slug_filename(notes_rel_path)
    md = practice_md_path(practice_dir, slug)
    img_dir = practice_image_dir(practice_dir, slug)
    if md.exists():
        if dry_run:
            print(f"[dry-run] would delete {md}")
        else:
            md.unlink()
            print(f"Deleted {md}")
    if img_dir.exists():
        if dry_run:
            print(f"[dry-run] would delete {img_dir}")
        else:
            shutil.rmtree(img_dir)
            print(f"Deleted {img_dir}")


def active_study_slugs(data_dir: Path) -> set[str]:
    data = load_all(data_dir)
    out: set[str] = set()
    for s in data["studies"]:
        if s.get("active", "1") == "0":
            continue
        slug = (s.get("notes_rel_path") or "").strip()
        if slug:
            out.add(slug_filename(slug))
    return out


def prune_orphans(output_dir: Path, data_dir: Path, dry_run: bool = False) -> None:
    practice_dir = output_dir / "practice"
    if not practice_dir.exists():
        return
    valid = active_study_slugs(data_dir)
    for md in practice_dir.glob("*.md"):
        slug = md.stem
        if slug not in valid:
            if dry_run:
                print(f"[dry-run] would prune orphan {md}")
            else:
                md.unlink()
                print(f"Pruned orphan {md}")
    images_root = practice_dir / "images"
    if images_root.exists():
        for sub in images_root.iterdir():
            if sub.is_dir() and sub.name not in valid:
                if dry_run:
                    print(f"[dry-run] would prune orphan {sub}")
                else:
                    shutil.rmtree(sub)
                    print(f"Pruned orphan {sub}")


def sync_studies(
    study_ids: list[str],
    data_dir: Path,
    output_dir: Path,
    notes_root: Path | None,
    dry_run: bool = False,
) -> int:
    count = 0
    for sid in study_ids:
        if write_study(sid, data_dir, output_dir, notes_root, dry_run=dry_run):
            count += 1
    return count


def sync_all(
    data_dir: Path,
    output_dir: Path,
    notes_root: Path | None,
    dry_run: bool = False,
) -> int:
    data = load_all(data_dir)
    ids = [
        s["id"] for s in data["studies"] if s.get("active", "1") != "0" and s.get("id")
    ]
    # Remove MD for inactive studies
    for s in data["studies"]:
        if s.get("active", "1") == "0":
            slug = (s.get("notes_rel_path") or "").strip()
            if slug:
                delete_study(slug, output_dir, dry_run=dry_run)
    count = sync_studies(ids, data_dir, output_dir, notes_root, dry_run=dry_run)
    prune_orphans(output_dir, data_dir, dry_run=dry_run)
    return count


def _files_match(a: Path, b: Path) -> bool:
    if not a.exists() or not b.exists():
        return False
    return a.stat().st_size == b.stat().st_size


def migrate_image_paths(
    data_dir: Path,
    output_dir: Path,
    notes_root: Path | None,
    dry_run: bool = False,
    delete_notes_copies: bool = False,
) -> int:
    data = load_all(data_dir)
    practice_dir = output_dir / "practice"
    updated = 0
    study_slug_by_id = {
        s["id"]: slug_filename(s.get("notes_rel_path") or "") for s in data["studies"]
    }

    for palace in data["palaces"]:
        image_rel = (palace.get("image_rel_path") or "").strip()
        if not image_rel or image_rel.startswith(PRACTICE_PREFIX):
            continue
        study_id = palace.get("study_id", "")
        slug = study_slug_by_id.get(study_id, "")
        if not slug:
            continue
        num = str(palace.get("palace_number") or "").strip()
        if not num:
            continue

        src = resolve_image(notes_root, image_rel, output_dir)
        if not src or not src.exists():
            continue

        ext = src.suffix.lstrip(".") or "png"
        dest_dir = practice_image_dir(practice_dir, slug)
        dest_dir.mkdir(parents=True, exist_ok=True)
        dest = dest_dir / f"{num}.{ext}"
        new_rel = f"{PRACTICE_PREFIX}{slug_filename(slug)}/{num}.{ext}"

        if dry_run:
            print(f"[dry-run] would migrate {palace['id']} → {new_rel}")
        else:
            shutil.copy2(src, dest)
            palace["image_rel_path"] = new_rel
            updated += 1
            if delete_notes_copies and _files_match(src, dest):
                try:
                    src.unlink()
                except OSError as e:
                    print(
                        f"Warning: could not delete notes copy {src}: {e}",
                        file=sys.stderr,
                    )

    if updated and not dry_run:
        save_all(data_dir, data)
        print(f"Migrated {updated} palace image path(s) in palaces.csv")
    return updated


def prune_legacy_notes_md(
    notes_root: Path,
    dry_run: bool = False,
) -> int:
    if not notes_root.exists():
        print(f"Notes root not found: {notes_root}", file=sys.stderr)
        return 0
    removed = 0
    skip_parts = {"technique", "portals"}
    for path in sorted(notes_root.rglob("mnemonics-*.md")):
        if skip_parts.intersection(path.parts):
            continue
        if dry_run:
            print(f"[dry-run] would delete legacy {path}")
        else:
            path.unlink()
            print(f"Deleted legacy {path}")
        removed += 1
    return removed


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(
        description="Sync Memory Palace practice Markdown files"
    )
    p.add_argument("--data-dir", type=Path, required=True)
    p.add_argument("--output-dir", type=Path, required=True)
    p.add_argument("--notes-root", type=Path, default=None)
    p.add_argument("--study-id", action="append", default=[], dest="study_ids")
    p.add_argument("--delete-slug", action="append", default=[], dest="delete_slugs")
    p.add_argument("--sync-all", action="store_true")
    p.add_argument("--migrate-image-paths", action="store_true")
    p.add_argument("--delete-notes-copies", action="store_true")
    p.add_argument("--prune-legacy-notes-md", action="store_true")
    p.add_argument("--dry-run", action="store_true")
    args = p.parse_args(argv)

    data_dir = args.data_dir.resolve()
    output_dir = args.output_dir.resolve()
    notes_root = args.notes_root.resolve() if args.notes_root else None

    for slug in args.delete_slugs:
        delete_study(slug, output_dir, dry_run=args.dry_run)

    if args.sync_all:
        n = sync_all(data_dir, output_dir, notes_root, dry_run=args.dry_run)
        print(f"Synced {n} active study/studies")
    elif args.study_ids:
        # Inactive study → delete practice file instead of writing
        data = load_all(data_dir)
        by_id = {s["id"]: s for s in data["studies"]}
        write_ids: list[str] = []
        for sid in args.study_ids:
            s = by_id.get(sid)
            if not s:
                continue
            if s.get("active", "1") == "0":
                slug = (s.get("notes_rel_path") or "").strip()
                if slug:
                    delete_study(slug, output_dir, dry_run=args.dry_run)
            else:
                write_ids.append(sid)
        if write_ids:
            sync_studies(
                write_ids, data_dir, output_dir, notes_root, dry_run=args.dry_run
            )

    if args.migrate_image_paths:
        migrate_image_paths(
            data_dir,
            output_dir,
            notes_root,
            dry_run=args.dry_run,
            delete_notes_copies=args.delete_notes_copies,
        )

    if args.prune_legacy_notes_md:
        if not notes_root:
            print("--prune-legacy-notes-md requires --notes-root", file=sys.stderr)
            return 1
        prune_legacy_notes_md(notes_root, dry_run=args.dry_run)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
