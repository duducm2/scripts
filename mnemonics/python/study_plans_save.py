"""Apply dashboard plan checkbox progress / backlog mutations to CSV, then export MD."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent))

from data_aggregator import load_all  # noqa: E402
from plan_csv import (  # noqa: E402
    BACKLOG_SECTION_PATH,
    load_plan_tables,
    next_id,
    save_plan_tables,
)
from study_plan_parser import (  # noqa: E402
    default_studies_root,
    enrich_plan,
    load_study_plan,
    slug_filename,
    strip_backlog_decor,
    strip_plan_topic_emoji,
)
from study_plans_md import sync_one_plan  # noqa: E402


def _plan_for_study(
    tables: dict[str, list[dict[str, str]]], study_id: str
) -> dict[str, str] | None:
    return next(
        (
            r
            for r in tables["plans"]
            if r.get("study_id") == study_id and r.get("active", "1") != "0"
        ),
        None,
    )


def apply_progress_to_csv(
    data_dir: Path,
    study_id: str,
    progress: dict[str, bool],
) -> tuple[int, dict[str, str] | None]:
    tables = load_plan_tables(data_dir)
    plan = _plan_for_study(tables, study_id)
    if not plan:
        return 0, None
    plan_id = plan["id"]
    changed = 0
    new_items: list[dict[str, str]] = []
    for row in tables["plan_items"]:
        if row.get("plan_id") != plan_id:
            new_items.append(row)
            continue
        iid = row.get("id") or ""
        if iid not in progress:
            new_items.append(row)
            continue
        want = "1" if progress[iid] else "0"
        if (row.get("checked") or "0") != want:
            row = dict(row)
            row["checked"] = want
            changed += 1
        new_items.append(row)
    if changed:
        save_plan_tables(data_dir, tables["plans"], new_items, tables["plan_resources"])
    return changed, plan


def add_backlog_item_csv(
    data_dir: Path, study_id: str, item_text: str
) -> dict[str, Any]:
    display = strip_backlog_decor(strip_plan_topic_emoji(item_text))
    if not display:
        return {"ok": False, "error": "empty backlog text", "study_id": study_id}

    tables = load_plan_tables(data_dir)
    plan = _plan_for_study(tables, study_id)
    if not plan:
        return {"ok": False, "error": "plan not found", "study_id": study_id}

    item_ids = {r.get("id") for r in tables["plan_items"] if r.get("id")}
    iid = next_id("PITEM_", item_ids)
    max_sort = 0
    for r in tables["plan_items"]:
        if r.get("plan_id") != plan["id"]:
            continue
        if (r.get("section_path") or "") != BACKLOG_SECTION_PATH:
            continue
        try:
            max_sort = max(max_sort, int(r.get("sort_order") or "0"))
        except ValueError:
            pass
    tables["plan_items"].append(
        {
            "id": iid,
            "plan_id": plan["id"],
            "section_path": BACKLOG_SECTION_PATH,
            "text": display,
            "checked": "0",
            "sort_order": str(max_sort + 1),
        }
    )
    save_plan_tables(
        data_dir, tables["plans"], tables["plan_items"], tables["plan_resources"]
    )
    return {"ok": True, "study_id": study_id, "item_id": iid}


def remove_backlog_item_csv(
    data_dir: Path, study_id: str, todo_id: str
) -> dict[str, Any]:
    tables = load_plan_tables(data_dir)
    plan = _plan_for_study(tables, study_id)
    if not plan:
        return {"ok": False, "error": "plan not found", "study_id": study_id}
    before = len(tables["plan_items"])
    tables["plan_items"] = [
        r
        for r in tables["plan_items"]
        if not (
            r.get("id") == todo_id
            and r.get("plan_id") == plan["id"]
            and (r.get("section_path") or "") == BACKLOG_SECTION_PATH
        )
    ]
    if len(tables["plan_items"]) == before:
        return {"ok": False, "error": "backlog item not found", "study_id": study_id}
    save_plan_tables(
        data_dir, tables["plans"], tables["plan_items"], tables["plan_resources"]
    )
    return {"ok": True, "study_id": study_id}


def refresh_plan_payload(
    study_id: str,
    studies_root: Path,
    data_dir: Path,
) -> dict[str, Any] | None:
    data = load_all(data_dir)
    study = next((s for s in data["studies"] if s["id"] == study_id), None)
    if not study:
        return None
    slug = (study.get("notes_rel_path") or "").strip()
    if not slug:
        return None
    plan = load_study_plan(studies_root, slug, data_dir=data_dir)
    if not plan:
        return None
    plan["study_id"] = study_id
    plan["study_title"] = study.get("title") or study_id
    return enrich_plan(plan)


def _sync_study_md(
    study_id: str, data_dir: Path, studies_root: Path, output_dir: Path
) -> None:
    data = load_all(data_dir)
    study = next((s for s in data["studies"] if s["id"] == study_id), None)
    if not study:
        return
    slug = slug_filename(study.get("notes_rel_path") or "")
    sync_one_plan(
        None, slug, output_dir, dry_run=False, data_dir=data_dir, study_id=study_id
    )


def save_study_progress(
    study_id: str,
    todos: list[dict[str, Any]],
    data_dir: Path,
    studies_root: Path,
    output_dir: Path,
    dry_run: bool = False,
) -> dict[str, Any]:
    progress = {str(t["id"]): bool(t.get("checked")) for t in todos if t.get("id")}
    if dry_run:
        return {"study_id": study_id, "ok": True, "changed": 0, "dry_run": True}
    changed, plan = apply_progress_to_csv(data_dir, study_id, progress)
    if not plan:
        return {"study_id": study_id, "ok": False, "error": "plan not found"}
    _sync_study_md(study_id, data_dir, studies_root, output_dir)
    return {
        "study_id": study_id,
        "ok": True,
        "changed": changed,
        "slug": plan.get("id"),
    }


def add_backlog_item(
    study_id: str,
    item_text: str,
    data_dir: Path,
    studies_root: Path,
    output_dir: Path,
    dry_run: bool = False,
) -> dict[str, Any]:
    if dry_run:
        return {"ok": True, "dry_run": True, "study_id": study_id}
    result = add_backlog_item_csv(data_dir, study_id, item_text)
    if not result.get("ok"):
        return result
    _sync_study_md(study_id, data_dir, studies_root, output_dir)
    plan = refresh_plan_payload(study_id, studies_root, data_dir)
    result["plan"] = plan
    return result


def remove_backlog_item(
    study_id: str,
    todo_id: str,
    data_dir: Path,
    studies_root: Path,
    output_dir: Path,
    dry_run: bool = False,
) -> dict[str, Any]:
    if dry_run:
        return {"ok": True, "dry_run": True, "study_id": study_id}
    result = remove_backlog_item_csv(data_dir, study_id, todo_id)
    if not result.get("ok"):
        return result
    _sync_study_md(study_id, data_dir, studies_root, output_dir)
    plan = refresh_plan_payload(study_id, studies_root, data_dir)
    result["plan"] = plan
    return result


def save_payload(
    payload: dict[str, Any],
    data_dir: Path,
    studies_root: Path,
    output_dir: Path,
    dry_run: bool = False,
) -> dict[str, Any]:
    if payload.get("action") == "add_backlog":
        return add_backlog_item(
            str(payload.get("study_id") or ""),
            str(payload.get("text") or ""),
            data_dir,
            studies_root,
            output_dir,
            dry_run=dry_run,
        )
    if payload.get("action") == "remove_backlog":
        return remove_backlog_item(
            str(payload.get("study_id") or ""),
            str(payload.get("todo_id") or ""),
            data_dir,
            studies_root,
            output_dir,
            dry_run=dry_run,
        )

    results: list[dict[str, Any]] = []
    plans = payload.get("plans")
    if isinstance(plans, dict):
        for study_id, entry in plans.items():
            todos = entry.get("todos") if isinstance(entry, dict) else None
            if not todos:
                continue
            results.append(
                save_study_progress(
                    study_id,
                    todos,
                    data_dir,
                    studies_root,
                    output_dir,
                    dry_run=dry_run,
                )
            )
    elif payload.get("study_id") and payload.get("todos"):
        results.append(
            save_study_progress(
                str(payload["study_id"]),
                payload["todos"],
                data_dir,
                studies_root,
                output_dir,
                dry_run=dry_run,
            )
        )

    ok = all(r.get("ok") for r in results) if results else False
    total_changed = sum(int(r.get("changed") or 0) for r in results)
    return {"ok": ok, "results": results, "total_changed": total_changed}


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="Save plan progress to CSV + export MD")
    p.add_argument("--data-dir", type=Path, required=True)
    p.add_argument("--output-dir", type=Path, required=True)
    p.add_argument("--studies-root", type=Path, default=None)
    p.add_argument("--payload-file", type=Path, default=None)
    p.add_argument("--dry-run", action="store_true")
    args = p.parse_args(argv)

    studies_root = (
        args.studies_root.resolve() if args.studies_root else default_studies_root()
    )
    data_dir = args.data_dir.resolve()
    output_dir = args.output_dir.resolve()

    if args.payload_file:
        payload = json.loads(args.payload_file.read_text(encoding="utf-8"))
    else:
        payload = json.loads(sys.stdin.read())

    result = save_payload(payload, data_dir, studies_root, output_dir, dry_run=args.dry_run)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())
