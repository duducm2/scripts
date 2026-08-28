"""Replace YouTube search links with curated direct watch URLs."""

from __future__ import annotations

import argparse
import csv
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from plan_csv import load_plan_tables, next_id, save_plan_tables  # noqa: E402
from plan_resources import (
    build_resource_line,
    classify_url,
    parse_resource_line,
)  # noqa: E402


def load_curated(path: Path) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    with path.open(encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            if not (row.get("plan_id") and row.get("section_path") and row.get("url")):
                continue
            rows.append(row)
    return rows


def section_urls(
    resources: list[dict[str, str]], plan_id: str, section_path: str
) -> set[str]:
    urls: set[str] = set()
    for r in resources:
        if r.get("plan_id") != plan_id:
            continue
        if (r.get("section_path") or "") != section_path:
            continue
        _, url = parse_resource_line(r.get("line") or "")
        if url:
            urls.add(url.strip())
    return urls


def is_search_resource(row: dict[str, str]) -> bool:
    _, url = parse_resource_line(row.get("line") or "")
    return classify_url(url) == "search"


def apply_curation(
    data_dir: Path,
    *,
    study_filter: str = "",
    dry_run: bool = False,
) -> dict[str, int]:
    curated_path = data_dir / "curated_youtube_resources.csv"
    if not curated_path.is_file():
        raise SystemExit(f"Missing curation file: {curated_path}")

    curated = load_curated(curated_path)
    if study_filter:
        curated = [r for r in curated if r.get("plan_id") == study_filter]

    tables = load_plan_tables(data_dir)
    resources = tables["plan_resources"]
    plan_ids = {p["id"]: p for p in tables["plans"]}
    res_ids = {r.get("id") for r in resources if r.get("id")}

    stats = {
        "sections": 0,
        "removed_search": 0,
        "added": 0,
        "skipped_existing": 0,
    }

    by_section: dict[tuple[str, str], list[dict[str, str]]] = {}
    for row in curated:
        key = (row["plan_id"], row["section_path"])
        by_section.setdefault(key, []).append(row)

    for (plan_id, section_path), entries in sorted(by_section.items()):
        if study_filter and plan_id != study_filter:
            continue
        if plan_id not in plan_ids:
            print(f"WARN: unknown plan_id {plan_id} for {section_path}")
            continue

        entries = sorted(entries, key=lambda r: int(r.get("sort_order") or "1"))
        existing_urls = section_urls(resources, plan_id, section_path)
        to_add = []
        for entry in entries:
            url = (entry.get("url") or "").strip()
            if url in existing_urls:
                stats["skipped_existing"] += 1
                continue
            to_add.append(entry)

        search_rows = [
            r
            for r in resources
            if r.get("plan_id") == plan_id
            and (r.get("section_path") or "") == section_path
            and is_search_resource(r)
        ]
        if not search_rows and not to_add:
            continue

        stats["sections"] += 1
        if dry_run:
            print(
                f"{plan_id} | {section_path.split(' > ')[-1]}: "
                f"remove {len(search_rows)} search, add {len(to_add)} curated"
            )
            stats["removed_search"] += len(search_rows)
            stats["added"] += len(to_add)
            continue

        resources[:] = [
            r
            for r in resources
            if not (
                r.get("plan_id") == plan_id
                and (r.get("section_path") or "") == section_path
                and is_search_resource(r)
            )
        ]
        stats["removed_search"] += len(search_rows)

        max_sort = 0
        for r in resources:
            if r.get("plan_id") != plan_id:
                continue
            if (r.get("section_path") or "") != section_path:
                continue
            try:
                max_sort = max(max_sort, int(r.get("sort_order") or "0"))
            except ValueError:
                pass

        for entry in to_add:
            rid = next_id("PRES_", res_ids)
            res_ids.add(rid)
            max_sort += 1
            title = (entry.get("title") or "").strip()
            url = (entry.get("url") or "").strip()
            resources.append(
                {
                    "id": rid,
                    "plan_id": plan_id,
                    "section_path": section_path,
                    "line": build_resource_line(title, url),
                    "sort_order": str(max_sort),
                }
            )
            stats["added"] += 1

    if not dry_run:
        save_plan_tables(
            data_dir,
            tables["plans"],
            tables["plan_items"],
            tables["plan_resources"],
        )

    return stats


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--data-dir",
        type=Path,
        default=Path(__file__).resolve().parent.parent / "data",
    )
    parser.add_argument(
        "--study",
        metavar="PLAN_ID",
        default="",
        help="Apply only for one plan (e.g. PLAN_0001)",
    )
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    stats = apply_curation(
        args.data_dir,
        study_filter=args.study.strip(),
        dry_run=args.dry_run,
    )
    mode = "DRY RUN" if args.dry_run else "APPLIED"
    print(
        f"{mode}: {stats['sections']} sections, "
        f"removed {stats['removed_search']} search rows, "
        f"added {stats['added']} curated rows, "
        f"skipped {stats['skipped_existing']} duplicates"
    )


if __name__ == "__main__":
    main()
