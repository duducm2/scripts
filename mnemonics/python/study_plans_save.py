"""Apply dashboard plan checkbox progress back to source Markdown files."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent))

from data_aggregator import load_all  # noqa: E402
from study_plan_parser import (  # noqa: E402
    HEADING_RE,
    TODO_RE,
    _stable_id,
    default_studies_root,
    discover_plan_path,
    enrich_plan,
    parse_plan_file,
    slug_filename,
)
from study_plans_md import sync_one_plan  # noqa: E402


def _format_mark(checked: bool) -> str:
    return "✅" if checked else " "


def apply_progress_to_text(text: str, slug: str, progress: dict[str, bool]) -> tuple[str, int]:
    """Return updated markdown and count of lines changed."""
    lines = text.replace("\r\n", "\n").split("\n")
    stack: list[tuple[int, str]] = []
    current_section: str | None = None
    in_backlog = False
    changed = 0

    for i, line in enumerate(lines):
        stripped = line.strip()
        hm = HEADING_RE.match(line)
        if hm:
            level = len(hm.group(1))
            heading_text = hm.group(2).strip()
            if level == 1:
                continue
            if level == 2 and heading_text == "📃 Backlog":
                in_backlog = True
                current_section = None
                continue
            in_backlog = False
            while stack and stack[-1][0] >= level:
                stack.pop()
            stack.append((level, heading_text))
            path_parts = [s[1] for s in stack]
            current_section = " > ".join(path_parts)
            continue

        if in_backlog or current_section is None:
            continue

        tm = TODO_RE.match(line)
        if not tm:
            continue

        text_part = tm.group("text").strip()
        todo_id = _stable_id(slug, current_section, text_part)
        if todo_id not in progress:
            continue

        checked = bool(progress[todo_id])
        current_checked = tm.group("mark").strip().lower() == "x" or tm.group("mark") == "✅"
        if current_checked == checked:
            continue

        mark = _format_mark(checked)
        lines[i] = f"- [{mark}] {text_part}"
        changed += 1

    return "\n".join(lines), changed


def save_plan_by_slug(
    slug: str,
    progress: dict[str, bool],
    studies_root: Path,
    output_dir: Path,
    dry_run: bool = False,
) -> dict[str, Any]:
    source = discover_plan_path(studies_root, slug)
    if not source:
        return {"slug": slug, "ok": False, "error": "plan file not found"}

    slug_norm = slug_filename(slug)
    text = source.read_text(encoding="utf-8")
    updated, changed = apply_progress_to_text(text, slug_norm, progress)
    if dry_run:
        return {
            "slug": slug_norm,
            "ok": True,
            "changed": changed,
            "source": str(source),
            "dry_run": True,
        }

    if changed:
        source.write_text(updated, encoding="utf-8")

    sync_one_plan(source, slug_norm, output_dir, dry_run=False)
    return {
        "slug": slug_norm,
        "ok": True,
        "changed": changed,
        "source": str(source),
    }


def save_study_progress(
    study_id: str,
    todos: list[dict[str, Any]],
    data_dir: Path,
    studies_root: Path,
    output_dir: Path,
    dry_run: bool = False,
) -> dict[str, Any]:
    data = load_all(data_dir)
    study = next((s for s in data["studies"] if s["id"] == study_id), None)
    if not study:
        return {"study_id": study_id, "ok": False, "error": "study not found"}

    slug = slug_filename(study.get("notes_rel_path") or "")
    progress = {str(t["id"]): bool(t.get("checked")) for t in todos if t.get("id")}
    result = save_plan_by_slug(slug, progress, studies_root, output_dir, dry_run=dry_run)
    result["study_id"] = study_id
    return result


def save_payload(
    payload: dict[str, Any],
    data_dir: Path,
    studies_root: Path,
    output_dir: Path,
    dry_run: bool = False,
) -> dict[str, Any]:
    """Save one or many studies from dashboard JSON payload."""
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
    elif payload.get("slug") and payload.get("progress"):
        slug = str(payload["slug"])
        progress = {str(k): bool(v) for k, v in payload["progress"].items()}
        results.append(
            save_plan_by_slug(slug, progress, studies_root, output_dir, dry_run=dry_run)
        )

    ok = all(r.get("ok") for r in results) if results else False
    total_changed = sum(int(r.get("changed") or 0) for r in results)
    return {"ok": ok, "results": results, "total_changed": total_changed}


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
    source = discover_plan_path(studies_root, slug)
    if not source:
        return None
    plan = enrich_plan(parse_plan_file(source, slug_filename(slug)))
    plan["study_id"] = study_id
    plan["study_title"] = study.get("title") or study_id
    return plan


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="Save plan checkbox progress to source Markdown")
    p.add_argument("--data-dir", type=Path, required=True)
    p.add_argument("--output-dir", type=Path, required=True)
    p.add_argument("--studies-root", type=Path, default=None)
    p.add_argument("--study-id", type=str, default=None)
    p.add_argument("--payload-file", type=Path, default=None)
    p.add_argument("--dry-run", action="store_true")
    args = p.parse_args(argv)

    studies_root = (
        args.studies_root.resolve()
        if args.studies_root
        else default_studies_root()
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
