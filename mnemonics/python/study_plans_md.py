"""Export study plans from CSV to output/plans/ Markdown for mobile/GitHub."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent))

from data_aggregator import load_all  # noqa: E402
from plan_csv import (  # noqa: E402
    load_plan_tables,
    plan_row_to_payload,
    render_plan_markdown,
)
from study_plan_parser import default_studies_root, enrich_plan, slug_filename  # noqa: E402


def plans_dir(output_dir: Path) -> Path:
    return output_dir / "plans"


def plan_output_path(output_dir: Path, slug: str) -> Path:
    return plans_dir(output_dir) / f"{slug}.md"


def sync_header(plan_id: str, slug: str) -> str:
    return f"<!-- synced from plans.csv ({plan_id}) / study {slug} -->\n\n"


def sync_one_plan_from_csv(
    plan: dict[str, str],
    items: list[dict[str, str]],
    resources: list[dict[str, str]],
    slug: str,
    output_dir: Path,
    dry_run: bool = False,
) -> Path | None:
    payload = enrich_plan(plan_row_to_payload(plan, items, resources, slug))
    body = render_plan_markdown(payload)
    dest = plan_output_path(output_dir, slug)
    out_text = sync_header(plan.get("id") or "", slug) + body
    if dry_run:
        print(f"[dry-run] would write {dest}")
        return dest
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_text(out_text, encoding="utf-8")
    print(f"Wrote {dest}")
    return dest


# Back-compat name used by study_plans_save / AHK callers that pass a Path.
def sync_one_plan(
    source: Path | None,
    slug: str,
    output_dir: Path,
    dry_run: bool = False,
    *,
    data_dir: Path | None = None,
    study_id: str | None = None,
) -> Path | None:
    """Export plan MD for slug from CSV. `source` is ignored (legacy)."""
    if data_dir is None:
        data_dir = Path(__file__).resolve().parent.parent / "data"
    tables = load_plan_tables(data_dir)
    data = load_all(data_dir)
    sid = study_id
    if not sid:
        for s in data["studies"]:
            if slug_filename(s.get("notes_rel_path") or "") == slug_filename(slug):
                sid = s["id"]
                break
    if not sid:
        print(f"No study for slug {slug}", file=sys.stderr)
        return None
    plan = next(
        (
            r
            for r in tables["plans"]
            if r.get("study_id") == sid and r.get("active", "1") != "0"
        ),
        None,
    )
    if not plan:
        print(f"No CSV plan for {sid}", file=sys.stderr)
        return None
    return sync_one_plan_from_csv(
        plan,
        tables["plan_items"],
        tables["plan_resources"],
        slug_filename(slug),
        output_dir,
        dry_run=dry_run,
    )


def build_readme_index(entries: list[dict[str, str]]) -> str:
    lines = [
        "# Study Plans",
        "",
        "Synced from `plans.csv` / `plan_items.csv` for mobile access.",
        "",
        "| Study | Plan |",
        "| --- | --- |",
    ]
    for e in sorted(entries, key=lambda x: x.get("sort_order", "999")):
        title = e.get("title", e.get("slug", ""))
        slug = e.get("slug", "")
        lines.append(f"| {title} | [{slug}.md]({slug}.md) |")
    lines.append("")
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
        if md.stem not in valid:
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
    _ = studies_root
    data = load_all(data_dir)
    study = next((s for s in data["studies"] if s["id"] == study_id), None)
    if not study:
        print(f"Study not found: {study_id}", file=sys.stderr)
        return None
    slug = slug_filename(study.get("notes_rel_path") or "")
    path = sync_one_plan(
        None, slug, output_dir, dry_run=dry_run, data_dir=data_dir, study_id=study_id
    )
    if not path:
        return None
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
    _ = studies_root
    data = load_all(data_dir)
    tables = load_plan_tables(data_dir)
    entries: list[dict[str, str]] = []
    for study in data["studies"]:
        if study.get("active", "1") == "0":
            continue
        slug = (study.get("notes_rel_path") or "").strip()
        if not slug:
            continue
        plan = next(
            (
                r
                for r in tables["plans"]
                if r.get("study_id") == study["id"] and r.get("active", "1") != "0"
            ),
            None,
        )
        if not plan:
            continue
        sync_one_plan_from_csv(
            plan,
            tables["plan_items"],
            tables["plan_resources"],
            slug_filename(slug),
            output_dir,
            dry_run=dry_run,
        )
        entries.append(
            {
                "study_id": study["id"],
                "title": study.get("title") or study["id"],
                "slug": slug_filename(slug),
                "sort_order": study.get("sort_order") or "999",
            }
        )

    readme_text = build_readme_index(entries)
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
    p = argparse.ArgumentParser(description="Export study plans CSV → output/plans/")
    p.add_argument("--data-dir", type=Path, required=True)
    p.add_argument("--output-dir", type=Path, required=True)
    p.add_argument("--studies-root", type=Path, default=None)
    p.add_argument("--study-id", action="append", default=[])
    p.add_argument("--sync-all", action="store_true")
    p.add_argument("--dry-run", action="store_true")
    args = p.parse_args(argv)

    studies_root = (
        args.studies_root.resolve() if args.studies_root else default_studies_root()
    )
    output_dir = args.output_dir.resolve()
    data_dir = args.data_dir.resolve()

    if args.sync_all:
        sync_all(data_dir, studies_root, output_dir, dry_run=args.dry_run)
        return 0

    if not args.study_id:
        p.error("Provide --sync-all or one or more --study-id")

    for sid in args.study_id:
        sync_study(sid, data_dir, studies_root, output_dir, dry_run=args.dry_run)

    if args.study_id and not args.dry_run:
        sync_all(data_dir, studies_root, output_dir, dry_run=False)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
