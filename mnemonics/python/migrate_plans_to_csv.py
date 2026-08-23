"""One-shot: migrate studies/*/ *-plan.md into plans / plan_items / plan_resources CSV."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from data_aggregator import load_all  # noqa: E402
from plan_csv import (  # noqa: E402
    load_plan_tables,
    migrate_parsed_plan_to_rows,
    next_id,
    save_plan_tables,
)
from study_plan_parser import (  # noqa: E402
    default_studies_root,
    discover_plan_path,
    parse_plan_file,
    slug_filename,
)


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="Migrate plan Markdown into CSV")
    p.add_argument("--data-dir", type=Path, required=True)
    p.add_argument("--studies-root", type=Path, default=None)
    p.add_argument("--force", action="store_true", help="Replace existing CSV plan rows")
    args = p.parse_args(argv)

    data_dir = args.data_dir.resolve()
    studies_root = (
        args.studies_root.resolve() if args.studies_root else default_studies_root()
    )
    data = load_all(data_dir)
    tables = load_plan_tables(data_dir)
    plans = list(tables["plans"])
    items = list(tables["plan_items"])
    resources = list(tables["plan_resources"])

    if args.force:
        plans, items, resources = [], [], []

    existing_study_plans = {r.get("study_id") for r in plans if r.get("study_id")}
    plan_ids = {r.get("id") for r in plans if r.get("id")}
    item_ids = {r.get("id") for r in items if r.get("id")}
    res_ids = {r.get("id") for r in resources if r.get("id")}

    migrated = 0
    for study in data["studies"]:
        if study.get("active", "1") == "0":
            continue
        study_id = study.get("id") or ""
        slug = (study.get("notes_rel_path") or "").strip()
        if not study_id or not slug:
            continue
        if study_id in existing_study_plans and not args.force:
            print(f"skip {study_id}: plan already in CSV")
            continue
        path = discover_plan_path(studies_root, slug)
        if not path:
            print(f"skip {study_id}: no *-plan.md")
            continue
        sn = slug_filename(slug)
        payload = parse_plan_file(path, sn)
        plan_id = next_id("PLAN_", plan_ids)
        plan_ids.add(plan_id)
        plan, new_items, new_res = migrate_parsed_plan_to_rows(
            payload,
            study_id,
            plan_id,
            item_ids=item_ids,
            res_ids=res_ids,
        )
        if args.force:
            # drop any leftover for this study
            plans = [r for r in plans if r.get("study_id") != study_id]
            keep_plan_ids = {r.get("id") for r in plans}
            items = [r for r in items if r.get("plan_id") in keep_plan_ids]
            resources = [r for r in resources if r.get("plan_id") in keep_plan_ids]
        plans.append(plan)
        items.extend(new_items)
        resources.extend(new_res)
        migrated += 1
        print(
            f"migrated {path.name} -> {plan_id} "
            f"({len(new_items)} items, {len(new_res)} resources)"
        )

    save_plan_tables(data_dir, plans, items, resources)
    print(f"done migrated={migrated} plans={len(plans)} items={len(items)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
