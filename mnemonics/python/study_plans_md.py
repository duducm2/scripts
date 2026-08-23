"""Sync study plan Markdown files to output/plans/ for mobile/GitHub access."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent))

from data_aggregator import load_all  # noqa: E402
from study_plan_parser import (  # noqa: E402
    default_studies_root,
    discover_plan_path,
    slug_filename,
)


def plans_dir(output_dir: Path) -> Path:
    return output_dir / "plans"


def plan_output_path(output_dir: Path, slug: str) -> Path:
    return plans_dir(output_dir) / f"{slug}.md"


def sync_header(source_rel: str) -> str:
    return f"<!-- synced from studies/{source_rel} -->\n\n"


def sync_one_plan(
    source: Path,
    slug: str,
    output_dir: Path,
    dry_run: bool = False,
) -> Path | None:
    dest = plan_output_path(output_dir, slug)
    text = source.read_text(encoding="utf-8")
    rel = f"{slug}/{source.name}"
    out_text = sync_header(rel) + text

    if dry_run:
        print(f"[dry-run] would write {dest}")
        return dest

    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_text(out_text, encoding="utf-8")
    print(f"Wrote {dest}")
    return dest


def build_readme_index(
    entries: list[dict[str, str]],
    dry_run: bool = False,
) -> None:
    lines = [
        "# Study Plans",
        "",
        "Synced from `mnemonics/studies/` for mobile access.",
        "",
        "| Study | Plan |",
        "| --- | --- |",
    ]
    for e in sorted(entries, key=lambda x: x.get("sort_order", "999")):
        title = e.get("title", e.get("slug", ""))
        slug = e.get("slug", "")
        lines.append(f"| {title} | [{slug}.md]({slug}.md) |")
    lines.append("")

    # Note: README path passed via caller
    return "\n".join(lines)


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
    pdir = plans_dir(output_dir)
    if not pdir.exists():
        return
    valid = active_study_slugs(data_dir)
    for md in pdir.glob("*.md"):
        if md.name.lower() == "readme.md":
            continue
        slug = md.stem
        if slug not in valid:
            if dry_run:
                print(f"[dry-run] would prune orphan {md}")
            else:
                md.unlink()
                print(f"Pruned orphan {md}")


def sync_study(
    study_id: str,
    data_dir: Path,
    studies_root: Path,
    output_dir: Path,
    dry_run: bool = False,
) -> dict[str, Any] | None:
    data = load_all(data_dir)
    study = next((s for s in data["studies"] if s["id"] == study_id), None)
    if not study:
        print(f"Study not found: {study_id}", file=sys.stderr)
        return None
    slug = slug_filename(study.get("notes_rel_path") or "")
    source = discover_plan_path(studies_root, slug)
    if not source:
        print(f"No plan file for {study_id} ({slug})", file=sys.stderr)
        return None
    sync_one_plan(source, slug, output_dir, dry_run=dry_run)
    return {
        "study_id": study_id,
        "title": study.get("title") or study_id,
        "slug": slug,
        "sort_order": study.get("sort_order") or "999",
    }


def sync_all(
    data_dir: Path,
    studies_root: Path,
    output_dir: Path,
    dry_run: bool = False,
) -> list[dict[str, str]]:
    data = load_all(data_dir)
    entries: list[dict[str, str]] = []
    for study in data["studies"]:
        if study.get("active", "1") == "0":
            continue
        slug = (study.get("notes_rel_path") or "").strip()
        if not slug:
            continue
        source = discover_plan_path(studies_root, slug)
        if not source:
            continue
        sync_one_plan(source, slug_filename(slug), output_dir, dry_run=dry_run)
        entries.append(
            {
                "study_id": study["id"],
                "title": study.get("title") or study["id"],
                "slug": slug_filename(slug),
                "sort_order": study.get("sort_order") or "999",
            }
        )

    readme_text = build_readme_index(entries, dry_run=dry_run)
    readme_path = plans_dir(output_dir) / "README.md"
    if dry_run:
        print(f"[dry-run] would write {readme_path}")
    else:
        readme_path.parent.mkdir(parents=True, exist_ok=True)
        readme_path.write_text(readme_text, encoding="utf-8")
        print(f"Wrote {readme_path}")

    prune_orphans(output_dir, data_dir, dry_run=dry_run)
    return entries


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="Sync study plans to output/plans/")
    p.add_argument("--data-dir", type=Path, required=True)
    p.add_argument("--output-dir", type=Path, required=True)
    p.add_argument(
        "--studies-root",
        type=Path,
        default=None,
        help="Root of mnemonics/studies (default: auto)",
    )
    p.add_argument("--study-id", action="append", default=[])
    p.add_argument("--sync-all", action="store_true")
    p.add_argument("--dry-run", action="store_true")
    args = p.parse_args(argv)

    studies_root = (
        args.studies_root.resolve()
        if args.studies_root
        else default_studies_root()
    )
    output_dir = args.output_dir.resolve()
    data_dir = args.data_dir.resolve()

    if args.sync_all:
        sync_all(data_dir, studies_root, output_dir, dry_run=args.dry_run)
        return 0

    if not args.study_id:
        p.error("Provide --sync-all or one or more --study-id")

    entries: list[dict[str, str]] = []
    for sid in args.study_id:
        entry = sync_study(sid, data_dir, studies_root, output_dir, dry_run=args.dry_run)
        if entry:
            entries.append(entry)

    if entries and not args.dry_run:
        readme_path = plans_dir(output_dir) / "README.md"
        # Rebuild index from all synced plans on disk + csv
        sync_all(data_dir, studies_root, output_dir, dry_run=False)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
