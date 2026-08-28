"""Audit study plan sections missing learning resources."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from urllib.parse import quote_plus

sys.path.insert(0, str(Path(__file__).resolve().parent))

from data_aggregator import load_all  # noqa: E402
from plan_csv import load_plan_tables, next_id, save_plan_tables  # noqa: E402
from plan_resources import build_resource_line  # noqa: E402
from study_plan_parser import default_studies_root, load_all_plans  # noqa: E402


def leaf_gaps(plan: dict) -> list[tuple[str, int]]:
    gaps: list[tuple[str, int]] = []

    def walk(nodes: list, prefix: list[str]) -> None:
        for n in nodes:
            path = " > ".join(prefix + [n["title"]])
            todos = n.get("todos") or []
            resources = n.get("resources") or []
            children = n.get("children") or []
            if todos and not resources and not children:
                gaps.append((path, len(todos)))
            walk(children, prefix + [n["title"]])

    walk(plan.get("sections") or [], [])
    return gaps


def youtube_search(title: str, extra: str = "") -> str:
    q = quote_plus(f"{title} {extra}".strip())
    return f"https://www.youtube.com/results?search_query={q}"


def default_resource_title(section_title: str) -> str:
    return f"{section_title} — YouTube learning resources"


def fill_gaps(data_dir: Path, studies_root: Path, dry_run: bool = False) -> int:
    tables = load_plan_tables(data_dir)
    plans = load_all_plans(studies_root, data_dir)
    res_ids = {r.get("id") for r in tables["plan_resources"] if r.get("id")}
    added = 0

    study_hints: dict[str, str] = {
        "STUDY_AI": "machine learning data science tutorial",
        "STUDY_ENGLISH": "English grammar C1 tutorial",
        "STUDY_GERMAN": "German language tutorial",
        "STUDY_PIANO": "piano tutorial",
        "STUDY_SCIENCE": "scientific research tutorial",
        "STUDY_PRODUCTOWNER": "product owner agile tutorial",
    }

    for study_id, plan in plans.items():
        plan_id = plan.get("plan_id") or ""
        if not plan_id:
            continue
        hint = study_hints.get(study_id, "tutorial")
        for section_path, todo_count in leaf_gaps(plan):
            title = section_path.split(" > ")[-1]
            rid = next_id("PRES_", res_ids)
            res_ids.add(rid)
            max_sort = 0
            for r in tables["plan_resources"]:
                if r.get("plan_id") != plan_id:
                    continue
                if (r.get("section_path") or "") != section_path:
                    continue
                try:
                    max_sort = max(max_sort, int(r.get("sort_order") or "0"))
                except ValueError:
                    pass
            line = build_resource_line(
                default_resource_title(title),
                youtube_search(title, hint),
            )
            row = {
                "id": rid,
                "plan_id": plan_id,
                "section_path": section_path,
                "line": line,
                "sort_order": str(max_sort + 1),
            }
            if dry_run:
                print(f"[dry-run] {study_id}: {section_path}")
            else:
                tables["plan_resources"].append(row)
            added += 1

    if added and not dry_run:
        save_plan_tables(
            data_dir, tables["plans"], tables["plan_items"], tables["plan_resources"]
        )
    return added


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="Audit/fill plan resource gaps")
    p.add_argument("--data-dir", type=Path, required=True)
    p.add_argument("--studies-root", type=Path, default=None)
    p.add_argument(
        "--fill", action="store_true", help="Add YouTube search resources for gaps"
    )
    p.add_argument("--dry-run", action="store_true")
    args = p.parse_args(argv)

    data_dir = args.data_dir.resolve()
    studies_root = (
        args.studies_root.resolve() if args.studies_root else default_studies_root()
    )
    plans = load_all_plans(studies_root, data_dir)

    total_gaps = 0
    for study_id, plan in sorted(plans.items()):
        gaps = leaf_gaps(plan)
        total_gaps += len(gaps)
        print(f"\n{study_id} ({plan.get('title', '')}): {len(gaps)} gaps")
        for path, count in gaps:
            print(f"  [{count} todos] {path}")

    print(f"\nTotal gaps: {total_gaps}")

    if args.fill:
        n = fill_gaps(data_dir, studies_root, dry_run=args.dry_run)
        action = "would add" if args.dry_run else "added"
        print(f"\n{action} {n} resource row(s)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
