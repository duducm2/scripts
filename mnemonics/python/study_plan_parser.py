"""Parse study plan Markdown into structured JSON for dashboard rendering."""

from __future__ import annotations

import hashlib
import re
from pathlib import Path
from typing import Any

from data_aggregator import load_all

BACKLOG_HEADING = "## 📃 Backlog"
RESOURCES_MARKER = "**🔗 Resources:**"
TODO_RE = re.compile(
    r"^-\s*\[(?P<mark>[ xX✅])\]\s*(?P<text>.+)$",
    re.UNICODE,
)
HEADING_RE = re.compile(r"^(#{1,6})\s+(.+)$")


def slug_filename(notes_rel_path: str) -> str:
    slug = notes_rel_path.replace("\\", "/").strip("/")
    if not slug or ".." in slug or ":" in slug:
        raise ValueError(f"Unsafe notes_rel_path: {notes_rel_path!r}")
    return slug


def discover_plan_path(studies_root: Path, notes_rel_path: str) -> Path | None:
    """Return plan file for a study folder, or None."""
    slug = slug_filename(notes_rel_path)
    folder = studies_root / slug
    if not folder.is_dir():
        # Case-sensitive fallback (e.g. AI folder)
        for child in studies_root.iterdir():
            if child.is_dir() and child.name.lower() == slug.lower():
                folder = child
                break
        else:
            return None

    preferred = folder / f"{slug}-plan.md"
    if preferred.exists():
        return preferred

    matches = sorted(folder.glob("*-plan.md"))
    if matches:
        return matches[0]
    return None


def _todo_checked(mark: str) -> bool:
    m = mark.strip()
    return m.lower() == "x" or m == "✅"


def _stable_id(slug: str, section_path: str, text: str) -> str:
    norm = re.sub(r"\s+", " ", text.strip().lower())
    raw = f"{slug}:{section_path}:{norm}"
    digest = hashlib.sha256(raw.encode("utf-8")).hexdigest()[:16]
    return f"{slug}:{digest}"


def _parse_resources_block(lines: list[str], start: int) -> tuple[list[str], int]:
    """Collect resource lines until next heading or end."""
    resources: list[str] = []
    i = start
    while i < len(lines):
        line = lines[i]
        if HEADING_RE.match(line):
            break
        if line.strip():
            resources.append(line.rstrip())
        i += 1
    return resources, i


def parse_plan_text(text: str, slug: str) -> dict[str, Any]:
    """Parse plan markdown content into structured payload."""
    lines = text.replace("\r\n", "\n").split("\n")
    title = ""
    backlog: list[str] = []
    sections: list[dict[str, Any]] = []
    flat_todos: list[dict[str, Any]] = []

    # Title: first # heading
    for line in lines:
        m = HEADING_RE.match(line)
        if m and len(m.group(1)) == 1:
            title = m.group(2).strip()
            break

    # Stack for section tree: list of (level, node)
    stack: list[tuple[int, dict[str, Any]]] = []
    current_section: dict[str, Any] | None = None
    section_path = ""
    in_backlog = False

    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        hm = HEADING_RE.match(line)
        if hm:
            level = len(hm.group(1))
            heading_text = hm.group(2).strip()

            if level == 1:
                i += 1
                continue

            if level == 2 and heading_text == "📃 Backlog":
                in_backlog = True
                current_section = None
                i += 1
                continue

            in_backlog = False

            node: dict[str, Any] = {
                "level": level,
                "title": heading_text,
                "anchor": "",
                "todos": [],
                "resources": [],
                "children": [],
            }

            # Pop stack to find parent
            while stack and stack[-1][0] >= level:
                stack.pop()

            if stack:
                stack[-1][1]["children"].append(node)
            else:
                sections.append(node)

            stack.append((level, node))
            current_section = node
            path_parts = [s[1]["title"] for s in stack]
            section_path = " > ".join(path_parts)
            node["anchor"] = (
                "sec-" + hashlib.sha256(section_path.encode("utf-8")).hexdigest()[:12]
            )
            i += 1
            continue

        if in_backlog:
            if stripped:
                backlog.append(stripped)
            i += 1
            continue

        if stripped == RESOURCES_MARKER:
            if current_section is not None:
                res_lines, next_i = _parse_resources_block(lines, i + 1)
                current_section["resources"] = res_lines
                i = next_i
                continue
            i += 1
            continue

        tm = TODO_RE.match(line)
        if tm and current_section is not None:
            text_part = tm.group("text").strip()
            checked = _todo_checked(tm.group("mark"))
            todo_id = _stable_id(slug, section_path, text_part)
            todo = {
                "id": todo_id,
                "text": text_part,
                "checked": checked,
            }
            current_section["todos"].append(todo)
            flat_todos.append({**todo, "section_path": section_path})
            i += 1
            continue

        i += 1

    return {
        "title": title,
        "slug": slug,
        "backlog": backlog,
        "sections": sections,
        "todos": flat_todos,
        "todo_count": len(flat_todos),
        "checked_count": sum(1 for t in flat_todos if t["checked"]),
    }


def parse_plan_file(path: Path, slug: str) -> dict[str, Any]:
    text = path.read_text(encoding="utf-8")
    payload = parse_plan_text(text, slug)
    payload["source_rel"] = path.name
    payload["source_path"] = str(path)
    return payload


def build_toc(
    flat_sections: list[dict[str, Any]], backlog: list[str]
) -> list[dict[str, Any]]:
    """Flatten section tree into TOC entries."""
    toc: list[dict[str, Any]] = []
    if backlog:
        toc.append({"level": 2, "title": "Backlog", "anchor": "plan-backlog"})

    def walk(nodes: list[dict[str, Any]]) -> None:
        for node in nodes:
            toc.append(
                {
                    "level": node["level"],
                    "title": node["title"],
                    "anchor": node["anchor"],
                }
            )
            walk(node.get("children") or [])

    walk(flat_sections)
    return toc


def enrich_plan(payload: dict[str, Any]) -> dict[str, Any]:
    payload["toc"] = build_toc(
        payload.get("sections") or [], payload.get("backlog") or []
    )
    return payload


def load_study_plan(
    studies_root: Path,
    notes_rel_path: str,
) -> dict[str, Any] | None:
    path = discover_plan_path(studies_root, notes_rel_path)
    if not path:
        return None
    slug = slug_filename(notes_rel_path)
    return enrich_plan(parse_plan_file(path, slug))


def load_all_plans(
    studies_root: Path,
    data_dir: Path,
) -> dict[str, dict[str, Any]]:
    """Return plans keyed by study_id for active studies with plan files."""
    data = load_all(data_dir)
    out: dict[str, dict[str, Any]] = {}
    for study in data["studies"]:
        if study.get("active", "1") == "0":
            continue
        slug = (study.get("notes_rel_path") or "").strip()
        if not slug:
            continue
        plan = load_study_plan(studies_root, slug)
        if plan:
            plan["study_id"] = study["id"]
            plan["study_title"] = study.get("title") or study["id"]
            out[study["id"]] = plan
    return out


def default_studies_root() -> Path:
    return Path(__file__).resolve().parent.parent / "studies"
