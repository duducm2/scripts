"""Import TASK_PACK from Desktop into tasks data (Python port of task_import.ahk)."""

from __future__ import annotations

import csv
import io
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any

from task_store import (
    INBOX_TITLES,
    VALID_FILTERS,
    VALID_KINDS,
    VALID_RECURRENCE,
    STATUS_EMOJIS,
    TaskStore,
    next_id,
    next_sort,
    now_stamp,
)

FILE_MARKERS = [
    ("TASK_PROJECTS.csv", "TASK_PROJECTS"),
    ("TASK_TASKS.csv", "TASK_TASKS"),
    ("TASK_INFO.csv", "TASK_INFO"),
]


def desktop_dir() -> Path:
    return Path.home() / "Desktop"


def newest_match(patterns: list[str]) -> Path | None:
    desk = desktop_dir()
    best: Path | None = None
    best_mtime = 0.0
    for pat in patterns:
        for p in desk.glob(pat):
            if not p.is_file():
                continue
            m = p.stat().st_mtime
            if m > best_mtime:
                best_mtime = m
                best = p
    return best


def normalize_pack(src: Path) -> Path:
    dest = desktop_dir() / "TASK_PACK.txt"
    if src.resolve() != dest.resolve():
        shutil.copy2(src, dest)
        # remove variant if different name
        try:
            if src.name.lower() != "task_pack.txt" and src.parent == desktop_dir():
                # keep gemini dumps; still overwrite canonical
                pass
        except OSError:
            pass
    return dest


def extract_section(text: str, file_name: str) -> str:
    names = [file_name]
    if file_name.lower().endswith(".csv"):
        names.append(file_name[:-4])
    else:
        names.append(file_name + ".csv")
    for name in names:
        for style in ("===", "---"):
            needle = f"{style}FILE: {name}{style}"
            end_needle = f"{style}END_FILE{style}"
            pos = text.find(needle)
            if pos < 0:
                continue
            rest = text[pos + len(needle) :]
            end = rest.find(end_needle)
            if end >= 0:
                return rest[:end].strip("\r\n \t")
            # until next FILE
            next_pos = -1
            for style2 in ("===", "---"):
                n2 = rest.find(f"\n{style2}FILE:")
                if n2 >= 0 and (next_pos < 0 or n2 < next_pos):
                    next_pos = n2
            if next_pos >= 0:
                return rest[:next_pos].strip("\r\n \t")
            return rest.strip("\r\n \t")
    return ""


def parse_csv_body(body: str) -> list[dict[str, str]]:
    body = body.strip()
    if body.startswith("```"):
        nl = body.find("\n")
        body = body[nl + 1 :] if nl >= 0 else ""
        if body.rstrip().endswith("```"):
            body = body.rstrip()[:-3]
    body = body.strip()
    if not body:
        return []
    reader = csv.DictReader(io.StringIO(body))
    rows = []
    for r in reader:
        row = {k: (v or "").strip() if v is not None else "" for k, v in r.items() if k}
        # skip duplicate headers
        if row.get("title", "").lower() == "title":
            continue
        if row.get("project_title", "").lower() == "project_title":
            continue
        if row.get("attach_to", "").lower() == "attach_to":
            continue
        rows.append(row)
    return rows


def write_fix_file(error: str, extra: str = "") -> str:
    desk = desktop_dir() / "TASK_AI_FIX.txt"
    body = (
        "The Desktop Tasks importer rejected my last TASK_PACK.txt. Fix and re-deliver.\r\n\r\n"
        f"IMPORT ERROR\r\n{error}\r\n\r\n"
    )
    if extra.strip():
        body += f"EXTRA NOTES\r\n{extra.strip()}\r\n\r\n"
    body += (
        "WHAT YOU MUST DO\r\n"
        "- Read the IMPORT ERROR above and reframe as one complete, valid TASK_PACK.\r\n"
        "- Prefer download chip; else one marked fence. Never claim a disk save.\r\n\r\n"
        "DELIVERY RULES (mandatory)\r\n"
        "- Deliver one complete TASK_PACK.txt.\r\n"
        "- Include ===PREVIEW=== and FILE sections for TASK_PROJECTS / TASK_TASKS / TASK_INFO.\r\n"
    )
    desk.write_text(body, encoding="utf-8")
    return str(desk)


def preview_pack(store: TaskStore) -> dict[str, Any]:
    src = newest_match(["TASK_PACK*.txt", "TASK_PACK*.csv", "gemini-code*.txt"])
    if not src:
        write_fix_file("No TASK_PACK file on Desktop")
        return {"ok": False, "error": "No TASK_PACK file on Desktop"}
    path = normalize_pack(src)
    text = path.read_text(encoding="utf-8", errors="replace")
    if not text.strip():
        write_fix_file("TASK_PACK is empty")
        return {"ok": False, "error": "TASK_PACK is empty"}
    projects = parse_csv_body(extract_section(text, "TASK_PROJECTS.csv"))
    tasks = parse_csv_body(extract_section(text, "TASK_TASKS.csv"))
    infos = parse_csv_body(extract_section(text, "TASK_INFO.csv"))
    if not projects and not tasks and not infos:
        write_fix_file("No usable rows in TASK_PACK")
        return {"ok": False, "error": "No usable rows in TASK_PACK"}
    return {
        "ok": True,
        "path": str(path),
        "counts": {"projects": len(projects), "tasks": len(tasks), "info": len(infos)},
        "projects": projects,
        "tasks": tasks,
        "info": infos,
    }


def commit_pack(store: TaskStore, pack: dict | None = None) -> dict[str, Any]:
    if pack is None:
        prev = preview_pack(store)
        if not prev.get("ok"):
            return prev
        pack = prev
    errors: list[str] = []
    projects = store.load("projects")
    tasks = store.load("tasks")
    infos = store.load("info_points")
    staged: dict[str, str] = {}
    new_projects: list[dict] = []
    new_tasks: list[dict] = []
    new_infos: list[dict] = []

    def ensure_project(title: str, filt: str) -> str:
        title = (title or "").strip() or INBOX_TITLES.get(filt, "Inbox")
        key = f"{filt}|{title.lower()}"
        if key in staged:
            return staged[key]
        for p in projects:
            if p.get("filter") == filt and (p.get("title") or "").strip().lower() == title.lower():
                staged[key] = p["id"]
                store.ensure_general_section(p["id"])
                return p["id"]
        row = {
            "id": next_id("PROJ_", projects + new_projects),
            "title": title,
            "filter": filt,
            "section_path": "",
            "sort_order": next_sort(projects + new_projects),
            "active": "1",
            "created_at": now_stamp(),
        }
        new_projects.append(row)
        projects.append(row)
        staged[key] = row["id"]
        store.ensure_general_section(row["id"])
        return row["id"]

    def ensure_section(project_id: str, name: str) -> tuple[str, str]:
        """section_path column = section display name; return (section_id, mirrored path)."""
        name = (name or "").strip()
        sec = store.find_or_create_section(project_id, name or "General")
        path = store._section_path_for(sec)
        return sec["id"], path

    for r in pack.get("projects") or []:
        title = (r.get("title") or "").strip()
        filt = (r.get("filter") or "work").strip().lower()
        if not title:
            errors.append("PROJECT: missing title")
            continue
        if filt not in VALID_FILTERS:
            errors.append(f"PROJECT invalid filter: {title}")
            continue
        ensure_project(title, filt)
        # apply section_path if new
        for p in projects:
            if p["id"] == staged[f"{filt}|{title.lower()}"]:
                if (r.get("section_path") or "").strip() and not p.get("section_path"):
                    p["section_path"] = r["section_path"].strip()

    task_index: dict[str, str] = {}
    for t in tasks:
        if t.get("title") and t.get("filter"):
            task_index[f"{t['filter']}|{(t['title'] or '').strip().lower()}"] = t["id"]

    for r in pack.get("tasks") or []:
        title = (r.get("title") or "").strip()
        filt = (r.get("filter") or "work").strip().lower()
        kind = (r.get("kind") or "punctual").strip().lower()
        recurrence = (r.get("recurrence") or "").strip().lower()
        if not title:
            errors.append("TASK: missing title")
            continue
        if filt not in VALID_FILTERS:
            errors.append(f"TASK invalid filter: {title}")
            continue
        if kind not in VALID_KINDS:
            errors.append(f"TASK invalid kind: {title}")
            continue
        if recurrence not in VALID_RECURRENCE:
            errors.append(f"TASK invalid recurrence ({recurrence}): {title}")
            continue
        if kind == "punctual":
            recurrence = ""
        emoji = (r.get("emoji") or "").strip() or STATUS_EMOJIS["general"]
        proj_id = ensure_project(r.get("project_title") or "", filt)
        section_id, section_path = ensure_section(proj_id, r.get("section_path") or "")
        tid = next_id("TASK_", tasks + new_tasks)
        row = {
            "id": tid,
            "project_id": proj_id,
            "section_id": section_id,
            "title": title,
            "emoji": emoji,
            "kind": kind,
            "recurrence": recurrence,
            "due_date": (r.get("due_date") or "").strip(),
            "next_due": (r.get("next_due") or "").strip(),
            "section_path": section_path,
            "filter": filt,
            "sort_order": next_sort(tasks + new_tasks),
            "completed_at": "",
            "created_at": now_stamp(),
            "active": "1",
        }
        new_tasks.append(row)
        tasks.append(row)
        task_index[f"{filt}|{title.lower()}"] = tid

    for r in pack.get("info") or []:
        attach_to = (r.get("attach_to") or "task").strip().lower()
        parent_title = (r.get("parent_title") or "").strip()
        filt = (r.get("filter") or "work").strip().lower()
        title = (r.get("title") or "").strip()
        if not title:
            errors.append("INFO: missing title")
            continue
        if filt not in VALID_FILTERS:
            errors.append(f"INFO invalid filter: {title}")
            continue
        if attach_to == "project":
            parent_id = ensure_project(parent_title or INBOX_TITLES.get(filt, ""), filt)
            parent_type = "project"
        else:
            key = f"{filt}|{parent_title.lower()}"
            parent_id = task_index.get(key, "")
            if not parent_id:
                errors.append(f"INFO parent task not found: {parent_title}")
                continue
            parent_type = "task"
        row = {
            "id": next_id("INFO_", infos + new_infos),
            "parent_type": parent_type,
            "parent_id": parent_id,
            "title": title,
            "body": (r.get("body") or "").strip(),
            "emoji": "ℹ️",
            "section_path": (r.get("section_path") or "").strip(),
            "sort_order": next_sort(infos + new_infos),
            "created_at": now_stamp(),
        }
        new_infos.append(row)
        infos.append(row)

    if not new_projects and not new_tasks and not new_infos and errors:
        write_fix_file("Import produced no rows", "\n".join(errors))
        return {"ok": False, "error": "Import produced no rows", "errors": errors}

    store.save("projects", projects)
    store.save("tasks", tasks)
    store.save("info_points", infos)
    store.migrate_sections()

    # archive pack
    path = Path(pack.get("path") or desktop_dir() / "TASK_PACK.txt")
    if path.is_file():
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        dest = store.imported_dir / f"TASK_PACK-{stamp}.txt"
        try:
            shutil.move(str(path), str(dest))
        except OSError:
            try:
                shutil.copy2(path, dest)
                path.unlink()
            except OSError:
                pass

    if errors:
        write_fix_file("Partial import — some rows failed", "\n".join(errors))
    return {
        "ok": True,
        "added": {
            "projects": len(new_projects),
            "tasks": len(new_tasks),
            "info": len(new_infos),
        },
        "errors": errors,
    }


def preview_labels(pack: dict[str, Any]) -> list[str]:
    """Palace-style one-line labels for AHK confirm ListView."""
    labels: list[str] = []
    counts = pack.get("counts") or {}
    labels.append(
        f"--- Counts: {counts.get('projects', 0)} project(s) · "
        f"{counts.get('tasks', 0)} task(s) · {counts.get('info', 0)} info ---"
    )
    projects = pack.get("projects") or []
    if projects:
        labels.append(f"--- Projects ({len(projects)}) ---")
        for r in projects:
            filt = (r.get("filter") or "?").strip()
            title = (r.get("title") or "").strip()
            sec = (r.get("section_path") or "").strip()
            line = f"[PROJ] {filt} · {title}"
            if sec:
                line += f" · {sec}"
            labels.append(line)
    tasks = pack.get("tasks") or []
    if tasks:
        labels.append(f"--- Tasks ({len(tasks)}) ---")
        for r in tasks:
            filt = (r.get("filter") or "?").strip()
            title = (r.get("title") or "").strip()
            kind = (r.get("kind") or "punctual").strip()
            proj = (r.get("project_title") or "").strip()
            line = f"[TASK] {filt} · {title} ({kind})"
            if proj:
                line += f" @ {proj}"
            labels.append(line)
    infos = pack.get("info") or []
    if infos:
        labels.append(f"--- Info ({len(infos)}) ---")
        for r in infos:
            title = (r.get("title") or "").strip()
            parent = (r.get("parent_title") or "").strip()
            attach = (r.get("attach_to") or "task").strip()
            line = f"[INFO] → {parent or '?'} ({attach}) · {title}"
            labels.append(line)
    return labels


def _cli_main(argv: list[str] | None = None) -> int:
    import argparse
    import json
    import sys

    parser = argparse.ArgumentParser(description="TASK_PACK preview / commit for Import Management")
    parser.add_argument("--data-dir", required=True, help="Path to tasks/data")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_prev = sub.add_parser("preview", help="Discover Desktop TASK_PACK and write preview files")
    p_prev.add_argument("--json-out", required=True, help="Write full preview JSON here")
    p_prev.add_argument("--labels-out", required=True, help="Write one label per line for AHK ListView")

    p_commit = sub.add_parser("commit", help="Commit a preview JSON pack into tasks/data")
    p_commit.add_argument("--pack-json", required=True, help="Preview JSON from preview --json-out")

    args = parser.parse_args(argv)
    store = TaskStore(Path(args.data_dir))

    if args.cmd == "preview":
        result = preview_pack(store)
        json_path = Path(args.json_out)
        labels_path = Path(args.labels_out)
        json_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        if result.get("ok"):
            labels = preview_labels(result)
            labels_path.write_text("\n".join(labels) + ("\n" if labels else ""), encoding="utf-8")
            counts = result.get("counts") or {}
            print(
                f"OK {counts.get('projects', 0)} projects "
                f"{counts.get('tasks', 0)} tasks {counts.get('info', 0)} info"
            )
            return 0
        labels_path.write_text("", encoding="utf-8")
        err = result.get("error") or "preview failed"
        print(err, file=sys.stderr)
        return 1

    if args.cmd == "commit":
        pack_path = Path(args.pack_json)
        if not pack_path.is_file():
            print(f"pack-json missing: {pack_path}", file=sys.stderr)
            return 1
        pack = json.loads(pack_path.read_text(encoding="utf-8"))
        if not pack.get("ok"):
            print(pack.get("error") or "invalid pack json", file=sys.stderr)
            return 1
        result = commit_pack(store, pack)
        Path(args.pack_json).with_suffix(".commit.json").write_text(
            json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        if not result.get("ok"):
            print(result.get("error") or "commit failed", file=sys.stderr)
            return 1
        added = result.get("added") or {}
        print(
            f"OK added {added.get('projects', 0)} projects "
            f"{added.get('tasks', 0)} tasks {added.get('info', 0)} info"
        )
        # Partial row errors still ok=True but fix file may exist
        if result.get("errors"):
            return 2
        return 0

    return 1


if __name__ == "__main__":
    raise SystemExit(_cli_main())
