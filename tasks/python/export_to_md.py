"""Export tasks CSV data to notes Markdown (punctual.md = personal filter).

Human-readable mirror for the notes repo. Not required to round-trip import.
"""

from __future__ import annotations

import argparse
import csv
from collections import defaultdict
from pathlib import Path

RECURRENCE_LABELS = {
    "daily": "Daily",
    "weekly": "Weekly",
    "monthly": "Monthly",
    "quarterly": "Quarterly",
    "biannual": "Bi-yearly (Every 6 months)",
    "yearly": "Yearly",
    "every_2y": "Every 2 years",
    "every_3y": "Every 3 years",
    "every_5y": "Every 5 years",
    "every_10y": "Every 10 years",
}


def read_csv(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        return []
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        return [{k: (v or "") for k, v in row.items()} for row in csv.DictReader(f)]


def active_rows(rows: list[dict[str, str]]) -> list[dict[str, str]]:
    return [r for r in rows if r.get("active", "1") != "0"]


def sort_key(row: dict[str, str]) -> tuple:
    try:
        so = int(row.get("sort_order") or 0)
    except ValueError:
        so = 0
    return (so, (row.get("title") or "").lower())


def info_export_text(info: dict[str, str]) -> str:
    t = (info.get("title") or "").strip()
    b = (info.get("body") or "").strip()
    if t and b and t != b:
        return t + "\n" + b
    return t or b


def write_area(
    path: Path,
    heading: str,
    filt: str,
    projects: list[dict[str, str]],
    sections: list[dict[str, str]],
    tasks: list[dict[str, str]],
    infos: list[dict[str, str]],
) -> None:
    projs = sorted(
        [p for p in projects if p.get("filter") == filt],
        key=sort_key,
    )
    secs_by_proj: dict[str, list[dict[str, str]]] = defaultdict(list)
    for s in sections:
        secs_by_proj[s.get("project_id") or ""].append(s)
    for pid in secs_by_proj:
        secs_by_proj[pid].sort(key=sort_key)

    tasks_by_sec: dict[str, list[dict[str, str]]] = defaultdict(list)
    tasks_by_proj_orphan: dict[str, list[dict[str, str]]] = defaultdict(list)
    for t in tasks:
        if t.get("filter") != filt:
            continue
        sid = (t.get("section_id") or "").strip()
        pid = t.get("project_id") or ""
        if sid:
            tasks_by_sec[sid].append(t)
        else:
            tasks_by_proj_orphan[pid].append(t)
    for k in tasks_by_sec:
        tasks_by_sec[k].sort(key=sort_key)
    for k in tasks_by_proj_orphan:
        tasks_by_proj_orphan[k].sort(key=sort_key)

    info_by_parent: dict[tuple[str, str], list[dict[str, str]]] = defaultdict(list)
    for info in infos:
        pt = (info.get("parent_type") or "").strip()
        pid = (info.get("parent_id") or "").strip()
        if pt and pid:
            info_by_parent[(pt, pid)].append(info)
    for k in info_by_parent:
        info_by_parent[k].sort(key=sort_key)

    lines: list[str] = [f"# {heading}", ""]

    if filt == "habits":
        by_rec: dict[str, list[tuple[dict, dict]]] = defaultdict(list)
        other: list[tuple[dict, dict]] = []
        for p in projs:
            for s in secs_by_proj.get(p["id"], []):
                for t in tasks_by_sec.get(s["id"], []):
                    if t.get("kind") == "habitual" and (t.get("recurrence") or ""):
                        by_rec[t["recurrence"]].append((p, t))
                    else:
                        other.append((p, t))
            for t in tasks_by_proj_orphan.get(p["id"], []):
                if t.get("kind") == "habitual" and (t.get("recurrence") or ""):
                    by_rec[t["recurrence"]].append((p, t))
                else:
                    other.append((p, t))

        for rec in sorted(
            by_rec.keys(),
            key=lambda r: (
                list(RECURRENCE_LABELS).index(r) if r in RECURRENCE_LABELS else 99
            ),
        ):
            label = RECURRENCE_LABELS.get(rec, rec or "Other")
            lines.append(f"## {label}")
            lines.append("")
            for _p, t in by_rec[rec]:
                em = (t.get("emoji") or "🔲").strip() or "🔲"
                lines.append(f"{em} {t.get('title') or ''}".rstrip())
            lines.append("")

        if other:
            lines.append("## Projects")
            lines.append("")
            seen_proj: set[str] = set()
            for p, t in other:
                if p["id"] not in seen_proj:
                    seen_proj.add(p["id"])
                    lines.append(f"### {p.get('title') or 'Project'}")
                    lines.append("")
                em = (t.get("emoji") or "🔲").strip() or "🔲"
                lines.append(f"{em} {t.get('title') or ''}".rstrip())
            lines.append("")

        for p in projs:
            for info in info_by_parent.get(("project", p["id"]), []):
                body = info_export_text(info)
                if body:
                    lines.append(f"ℹ️ {body}")
                    lines.append("")
    else:
        for p in projs:
            lines.append(f"## {p.get('title') or 'Project'}")
            lines.append("")
            for info in info_by_parent.get(("project", p["id"]), []):
                body = info_export_text(info)
                if body:
                    lines.append(f"ℹ️ {body}")
                    lines.append("")

            secs = secs_by_proj.get(p["id"], [])
            for s in secs:
                title = (s.get("title") or "General").strip()
                is_general = title.lower() == "general"
                if not is_general:
                    lines.append(f"### {title}")
                    lines.append("")
                for t in tasks_by_sec.get(s["id"], []):
                    em = (t.get("emoji") or "🔲").strip() or "🔲"
                    due = (t.get("due_date") or t.get("next_due") or "").strip()
                    line = f"{em} {t.get('title') or ''}".rstrip()
                    if due:
                        line += f" ({due})"
                    lines.append(line)
                    for info in info_by_parent.get(("task", t["id"]), []):
                        body = info_export_text(info)
                        if body:
                            lines.append(f"ℹ️ {body}")
                    lines.append("")

            for t in tasks_by_proj_orphan.get(p["id"], []):
                em = (t.get("emoji") or "🔲").strip() or "🔲"
                lines.append(f"{em} {t.get('title') or ''}".rstrip())
                lines.append("")

    path.parent.mkdir(parents=True, exist_ok=True)
    text = "\n".join(lines).rstrip() + "\n"
    path.write_text(text, encoding="utf-8")


def main() -> int:
    ap = argparse.ArgumentParser(description="Export tasks CSV to notes Markdown")
    ap.add_argument("--data-dir", required=True, type=Path)
    ap.add_argument("--punctual", required=True, type=Path)
    ap.add_argument("--work", type=Path, default=None)
    ap.add_argument("--habits", type=Path, default=None)
    args = ap.parse_args()

    data = args.data_dir
    projects = active_rows(read_csv(data / "projects.csv"))
    sections = read_csv(data / "sections.csv")
    tasks = active_rows(read_csv(data / "tasks.csv"))
    infos = read_csv(data / "info_points.csv")

    write_area(args.punctual, "Personal", "personal", projects, sections, tasks, infos)
    if args.work:
        write_area(args.work, "Work", "work", projects, sections, tasks, infos)
    if args.habits:
        write_area(
            args.habits, "Habits & Health", "habits", projects, sections, tasks, infos
        )

    print(
        f"Exported punctual={args.punctual} ({len(projects)} projects, {len(tasks)} tasks)"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
