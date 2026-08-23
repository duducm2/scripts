"""Load study plans from CSV and render Markdown for mobile export."""

from __future__ import annotations

import csv
import hashlib
import re
from pathlib import Path
from typing import Any

from schemas import HEADERS, PLAN_ITEMS_HEADERS, PLAN_RESOURCES_HEADERS, PLANS_HEADERS

BACKLOG_SECTION_PATH = "Backlog"
RESOURCES_MARKER = "**🔗 Resources:**"


def _read_csv(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        return []
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        return [{k: (v or "").strip() for k, v in row.items()} for row in csv.DictReader(f)]


def _write_csv(path: Path, rows: list[dict[str, str]], headers: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=headers, extrasaction="ignore")
        w.writeheader()
        for row in rows:
            w.writerow({h: row.get(h, "") for h in headers})


def load_plan_tables(data_dir: Path) -> dict[str, list[dict[str, str]]]:
    return {
        "plans": _read_csv(data_dir / "plans.csv"),
        "plan_items": _read_csv(data_dir / "plan_items.csv"),
        "plan_resources": _read_csv(data_dir / "plan_resources.csv"),
    }


def save_plan_tables(
    data_dir: Path,
    plans: list[dict[str, str]],
    items: list[dict[str, str]],
    resources: list[dict[str, str]],
) -> None:
    _write_csv(data_dir / "plans.csv", plans, PLANS_HEADERS)
    _write_csv(data_dir / "plan_items.csv", items, PLAN_ITEMS_HEADERS)
    _write_csv(data_dir / "plan_resources.csv", resources, PLAN_RESOURCES_HEADERS)


def _checked_bool(val: str) -> bool:
    v = (val or "").strip().lower()
    return v in ("1", "true", "yes", "x", "✅")


def _sort_key(row: dict[str, str]) -> tuple[int, str]:
    try:
        n = int(row.get("sort_order") or "0")
    except ValueError:
        n = 0
    return (n, row.get("id") or "")


def _ensure_section_tree(
    sections: list[dict[str, Any]],
    path_parts: list[str],
) -> dict[str, Any]:
    """Return leaf section node, creating intermediates under sections root list."""
    nodes = sections
    parent: dict[str, Any] | None = None
    stack_path: list[str] = []
    for i, title in enumerate(path_parts):
        stack_path.append(title)
        found = None
        for n in nodes:
            if n.get("title") == title:
                found = n
                break
        if found is None:
            section_path = " > ".join(stack_path)
            found = {
                "level": i + 2,
                "title": title,
                "anchor": "sec-"
                + hashlib.sha256(section_path.encode("utf-8")).hexdigest()[:12],
                "todos": [],
                "resources": [],
                "children": [],
            }
            nodes.append(found)
        parent = found
        nodes = found["children"]
    assert parent is not None
    return parent


def plan_row_to_payload(
    plan: dict[str, str],
    items: list[dict[str, str]],
    resources: list[dict[str, str]],
    slug: str,
) -> dict[str, Any]:
    plan_id = plan.get("id") or ""
    plan_items = sorted(
        [r for r in items if r.get("plan_id") == plan_id],
        key=_sort_key,
    )
    plan_res = sorted(
        [r for r in resources if r.get("plan_id") == plan_id],
        key=_sort_key,
    )

    backlog: list[dict[str, Any]] = []
    sections: list[dict[str, Any]] = []
    flat_todos: list[dict[str, Any]] = []

    res_by_section: dict[str, list[str]] = {}
    for r in plan_res:
        sp = (r.get("section_path") or "").strip() or BACKLOG_SECTION_PATH
        res_by_section.setdefault(sp, []).append(r.get("line") or "")

    for it in plan_items:
        text = (it.get("text") or "").strip()
        if not text:
            continue
        sp = (it.get("section_path") or "").strip() or BACKLOG_SECTION_PATH
        checked = _checked_bool(it.get("checked") or "")
        todo = {
            "id": it.get("id") or "",
            "text": text,
            "checked": checked,
        }
        flat_todos.append({**todo, "section_path": sp})
        if sp == BACKLOG_SECTION_PATH or sp.lower() == "backlog":
            backlog.append(todo)
            continue
        parts = [p.strip() for p in sp.split(">") if p.strip()]
        if not parts:
            backlog.append(todo)
            continue
        node = _ensure_section_tree(sections, parts)
        node["todos"].append(todo)

    # Attach resources to matching section leaves
    def attach_resources(nodes: list[dict[str, Any]], prefix: list[str]) -> None:
        for n in nodes:
            path = " > ".join(prefix + [n["title"]])
            if path in res_by_section:
                n["resources"] = list(res_by_section[path])
            attach_resources(n.get("children") or [], prefix + [n["title"]])

    attach_resources(sections, [])

    return {
        "title": plan.get("title") or "Study Plan",
        "slug": slug,
        "plan_id": plan_id,
        "backlog": backlog,
        "sections": sections,
        "todos": flat_todos,
        "todo_count": len(flat_todos),
        "checked_count": sum(1 for t in flat_todos if t["checked"]),
        "source_rel": f"{slug}.csv",
        "source_path": "",
    }


def render_plan_markdown(payload: dict[str, Any]) -> str:
    """Render dashboard-shaped plan payload to Markdown."""
    lines: list[str] = [f"# {payload.get('title') or 'Study Plan'}", ""]
    lines.append("## 📃 Backlog")
    lines.append("")
    for t in payload.get("backlog") or []:
        mark = "✅" if t.get("checked") else " "
        lines.append(f"- [{mark}] {t.get('text') or ''}")
        lines.append("")

    def walk(nodes: list[dict[str, Any]]) -> None:
        for n in nodes:
            level = min(max(int(n.get("level") or 2), 2), 6)
            lines.append("#" * level + " " + (n.get("title") or ""))
            lines.append("")
            for t in n.get("todos") or []:
                mark = "✅" if t.get("checked") else " "
                lines.append(f"- [{mark}] {t.get('text') or ''}")
            if n.get("todos"):
                lines.append("")
            resources = n.get("resources") or []
            if resources:
                lines.append(RESOURCES_MARKER)
                lines.append("")
                for r in resources:
                    lines.append(r if r.startswith("-") else f"- {r}")
                lines.append("")
            walk(n.get("children") or [])

    walk(payload.get("sections") or [])
    return "\n".join(lines).rstrip() + "\n"


def next_id(prefix: str, existing: set[str], width: int = 4) -> str:
    n = 1
    while True:
        cand = f"{prefix}{n:0{width}d}"
        if cand not in existing:
            return cand
        n += 1


def migrate_parsed_plan_to_rows(
    payload: dict[str, Any],
    study_id: str,
    plan_id: str,
    *,
    item_ids: set[str],
    res_ids: set[str],
) -> tuple[dict[str, str], list[dict[str, str]], list[dict[str, str]]]:
    """Convert parse_plan_text payload into CSV rows."""
    plan = {
        "id": plan_id,
        "study_id": study_id,
        "title": payload.get("title") or "Study Plan",
        "sort_order": "1",
        "active": "1",
    }
    items: list[dict[str, str]] = []
    resources: list[dict[str, str]] = []
    sort_i = 0

    for t in payload.get("backlog") or []:
        sort_i += 1
        iid = next_id("PITEM_", item_ids)
        item_ids.add(iid)
        items.append(
            {
                "id": iid,
                "plan_id": plan_id,
                "section_path": BACKLOG_SECTION_PATH,
                "text": t.get("text") or "",
                "checked": "1" if t.get("checked") else "0",
                "sort_order": str(sort_i),
            }
        )

    def walk(nodes: list[dict[str, Any]], path_parts: list[str]) -> None:
        nonlocal sort_i
        for n in nodes:
            parts = path_parts + [n.get("title") or ""]
            sp = " > ".join(p for p in parts if p)
            for t in n.get("todos") or []:
                sort_i += 1
                iid = next_id("PITEM_", item_ids)
                item_ids.add(iid)
                items.append(
                    {
                        "id": iid,
                        "plan_id": plan_id,
                        "section_path": sp,
                        "text": t.get("text") or "",
                        "checked": "1" if t.get("checked") else "0",
                        "sort_order": str(sort_i),
                    }
                )
            for line in n.get("resources") or []:
                rid = next_id("PRES_", res_ids)
                res_ids.add(rid)
                resources.append(
                    {
                        "id": rid,
                        "plan_id": plan_id,
                        "section_path": sp,
                        "line": line,
                        "sort_order": str(len(resources) + 1),
                    }
                )
            walk(n.get("children") or [], parts)

    walk(payload.get("sections") or [], [])
    return plan, items, resources
