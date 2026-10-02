"""CSV store for the Tasks web app (projects / sections / tasks / info / attachments)."""

from __future__ import annotations

import csv
import json
import re
import shutil
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

HEADERS = {
    "projects": [
        "id",
        "title",
        "filter",
        "section_path",
        "sort_order",
        "active",
        "created_at",
        "icon_ref",
        "icon_color",
        "icon_tint",
        "status",
    ],
    "sections": ["id", "project_id", "title", "sort_order", "status"],
    "tasks": [
        "id",
        "project_id",
        "section_id",
        "title",
        "emoji",
        "kind",
        "recurrence",
        "due_date",
        "next_due",
        "section_path",
        "filter",
        "sort_order",
        "completed_at",
        "created_at",
        "active",
        "import_batch",
    ],
    "info_points": [
        "id",
        "parent_type",
        "parent_id",
        "title",
        "body",
        "emoji",
        "section_path",
        "sort_order",
        "created_at",
        "edit_lock",
        "birth_date",
    ],
    "attachments": [
        "id",
        "parent_type",
        "parent_id",
        "kind",
        "ref",
        "description",
        "sort_order",
    ],
}

GENERAL_SECTION = "General"

STATUS_EMOJIS = {
    "general": "🔲",
    "waiting": "⏳",
    "important": "⚡",
    "done": "✅",
    "doubt": "❓",
    "none": "",
}

# Spoken / pack words that mean a status. Glyphs in STATUS_EMOJIS pass through as-is.
STATUS_EMOJI_ALIASES = {
    "general": "general",
    "default": "general",
    "normal": "general",
    "waiting": "waiting",
    "wait": "waiting",
    "blocked": "waiting",
    "important": "important",
    "priority": "important",
    "urgent": "important",
    "doubt": "doubt",
    "unsure": "doubt",
    "uncertain": "doubt",
    "done": "done",
    "complete": "done",
    "completed": "done",
    "none": "none",
}


def normalize_waiting_status(raw: str) -> str | None:
    """Project and section status is empty or waiting."""
    key = (raw or "").strip().lower()
    if key in {"", "waiting"}:
        return key
    return None


def resolve_status_emoji(raw: str) -> str:
    """Map a pack emoji cell to a stored glyph.

    Empty and the words general / waiting / important / doubt / done (plus a few
    spoken aliases) become the status glyphs. Any other value is kept, so a
    custom emoji such as 🩺 still imports.
    """
    text = (raw or "").strip()
    if not text:
        return STATUS_EMOJIS["general"]
    for glyph in STATUS_EMOJIS.values():
        if glyph and text == glyph:
            return glyph
    key = text.lower().strip(" .,;:!?")
    status = STATUS_EMOJI_ALIASES.get(key) or (key if key in STATUS_EMOJIS else "")
    if status:
        return STATUS_EMOJIS[status]
    return text


VALID_FILTERS = {"work", "personal", "habits"}
VALID_KINDS = {"punctual", "habitual"}
VALID_RECURRENCE = {
    "",
    "daily",
    "weekly",
    "monthly",
    "quarterly",
    "biannual",
    "yearly",
    "every_2y",
    "every_3y",
    "every_5y",
    "every_10y",
}

# Rollback: set False if Drive mtime is flaky (always read CSV from disk).
TASK_STORE_CACHE = True

INBOX_TITLES = {
    "work": "Work inbox",
    "personal": "Personal inbox",
    "habits": "Habits & Health",
}


def now_stamp() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _export_sort_key(row: dict) -> tuple:
    try:
        so = int(row.get("sort_order") or 0)
    except ValueError:
        so = 0
    return (so, (row.get("title") or "").lower())


def _status_name(emoji: str) -> str:
    em = (emoji or "").strip()
    for name, glyph in STATUS_EMOJIS.items():
        if em == glyph:
            return name
    key = em.lower().strip(" .,;:!?")
    aliased = STATUS_EMOJI_ALIASES.get(key)
    if aliased:
        return aliased
    if key in STATUS_EMOJIS:
        return key
    return "custom"


def _status_glyph(raw: str) -> tuple[str, str]:
    status = _status_name(raw)
    if status == "custom":
        return ((raw or "").strip(), "custom")
    return (STATUS_EMOJIS[status], status)


def _age_months(birth: str, as_of: date) -> int | None:
    raw = (birth or "").strip()[:10]
    try:
        bday = datetime.strptime(raw, "%Y-%m-%d").date()
    except ValueError:
        return None
    months = (as_of.year - bday.year) * 12 + (as_of.month - bday.month)
    if as_of.day < bday.day:
        months -= 1
    return max(0, months)


def project_export_filename(project: dict) -> str:
    title = (project.get("title") or "project").strip()
    slug = re.sub(r"[^A-Za-z0-9]+", "-", title).strip("-") or "project"
    pid = (
        re.sub(r"[^A-Za-z0-9_-]+", "", (project.get("id") or "project").strip())
        or "project"
    )
    return f"{slug}__{pid}.json"


def _export_attachments(rows: list[dict]) -> list[dict]:
    out: list[dict] = []
    for a in sorted(rows, key=_export_sort_key):
        doc: dict[str, Any] = {
            "id": a.get("id") or "",
            "kind": (a.get("kind") or "").strip(),
        }
        desc = (a.get("description") or "").strip()
        if desc:
            doc["description"] = desc
        ref = (a.get("ref") or "").replace("\\", "/").strip()
        if ref:
            doc["file"] = ref
        out.append(doc)
    return out


def _export_info(info: dict, atts_by_parent: dict, as_of: date) -> dict:
    doc: dict[str, Any] = {
        "id": info.get("id") or "",
        "title": (info.get("title") or "").strip(),
    }
    body = (info.get("body") or "").strip()
    if body and body != doc["title"]:
        doc["body"] = body
    category = (info.get("section_path") or "").strip()
    if category:
        doc["category"] = category
    emoji = (info.get("emoji") or "").strip()
    if emoji:
        doc["emoji"] = emoji
    birth = (info.get("birth_date") or "").strip()
    if birth:
        doc["birth_date"] = birth[:10]
        months = _age_months(birth, as_of)
        if months is not None:
            unit = "month" if months == 1 else "months"
            doc["age"] = f"{months} {unit}"
    if (info.get("edit_lock") or "").strip().lower() == "agent":
        doc["edit_lock"] = "agent"
    created = (info.get("created_at") or "").strip()
    if created:
        doc["created_at"] = created
    atts = _export_attachments(atts_by_parent.get(("info", info.get("id") or ""), []))
    if atts:
        doc["attachments"] = atts
    return doc


def _export_task(
    task: dict, infos_by_parent: dict, atts_by_parent: dict, as_of: date
) -> dict:
    kind = (task.get("kind") or "punctual").strip() or "punctual"
    active = (task.get("active") or "1") != "0"
    emoji, status = _status_glyph(str(task.get("emoji") or ""))
    doc: dict[str, Any] = {
        "id": task.get("id") or "",
        "title": (task.get("title") or "").strip(),
        "emoji": emoji,
        "status": status,
        "kind": kind,
        "active": active,
        "open": active and kind != "habitual" and status != "done",
    }
    rec = (task.get("recurrence") or "").strip()
    if rec:
        doc["recurrence"] = rec
    for key in ("due_date", "next_due", "completed_at", "created_at"):
        val = (task.get(key) or "").strip()
        if val:
            doc[key] = val
    if (task.get("import_batch") or "").strip():
        doc["newly_imported"] = True
    tid = task.get("id") or ""
    notes = [
        _export_info(i, atts_by_parent, as_of)
        for i in sorted(infos_by_parent.get(("task", tid), []), key=_export_sort_key)
    ]
    if notes:
        doc["info"] = notes
    atts = _export_attachments(atts_by_parent.get(("task", tid), []))
    if atts:
        doc["attachments"] = atts
    return doc


def today() -> str:
    return datetime.now().strftime("%Y-%m-%d")


def age_months(birth_date: str, as_of: str | None = None) -> int | None:
    """Completed months from YYYY-MM-DD birth_date to as_of (default today)."""
    raw = (birth_date or "").strip()
    if not raw:
        return None
    try:
        birth = datetime.strptime(raw[:10], "%Y-%m-%d").date()
        ref = datetime.strptime((as_of or today())[:10], "%Y-%m-%d").date()
    except ValueError:
        return None
    months = (ref.year - birth.year) * 12 + (ref.month - birth.month)
    if ref.day < birth.day:
        months -= 1
    return max(0, months)


def read_csv(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        return []
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        return [{k: (v or "") for k, v in row.items()} for row in csv.DictReader(f)]


def write_csv(path: Path, headers: list[str], rows: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=headers, lineterminator="\n")
        w.writeheader()
        for r in rows:
            w.writerow({h: r.get(h, "") for h in headers})


def next_id(prefix: str, rows: list[dict], pad: int = 4) -> str:
    mx = 0
    for r in rows:
        rid = r.get("id") or ""
        m = re.match(re.escape(prefix) + r"(\d+)$", rid)
        if m:
            mx = max(mx, int(m.group(1)))
    return f"{prefix}{mx + 1:0{pad}d}"


def next_sort(rows: list[dict]) -> str:
    mx = 0
    for r in rows:
        try:
            mx = max(mx, int(r.get("sort_order") or 0))
        except ValueError:
            pass
    return str(mx + 10)


def sort_key_row(r: dict) -> tuple:
    try:
        so = int(r.get("sort_order") or 0)
    except ValueError:
        so = 0
    return (so, r.get("id") or "")


def parse_move_dir(direction) -> int:
    try:
        d = int(direction)
    except (TypeError, ValueError):
        d = 0
    if d in (-1, 1):
        return d
    text = str(direction or "").strip().lower()
    if text in ("up", "-1"):
        return -1
    if text in ("down", "1"):
        return 1
    raise ValueError("dir must be -1 or 1")


def swap_sort_order(
    rows: list[dict],
    row_id: str,
    direction: int,
    sibling_predicate,
) -> tuple[list[dict], dict | None, bool]:
    """Swap sort_order with adjacent sibling. Returns (rows, target, moved)."""
    target = next((r for r in rows if r.get("id") == row_id), None)
    if not target:
        return rows, None, False
    siblings = sorted(
        [r for r in rows if sibling_predicate(r, target)],
        key=sort_key_row,
    )
    idx = next((i for i, r in enumerate(siblings) if r.get("id") == row_id), -1)
    if idx < 0:
        return rows, target, False
    j = idx + direction
    if j < 0 or j >= len(siblings):
        return rows, target, False
    a, b = siblings[idx], siblings[j]
    a_so, b_so = a.get("sort_order") or "0", b.get("sort_order") or "0"
    a["sort_order"], b["sort_order"] = b_so, a_so
    by_id = {r["id"]: r for r in rows}
    by_id[a["id"]] = a
    by_id[b["id"]] = b
    out = [by_id.get(r["id"], r) for r in rows]
    return out, by_id[row_id], True


class TaskStore:
    def __init__(self, data_dir: Path):
        self.data_dir = data_dir
        self.attach_dir = data_dir / "attachments"
        self.imported_dir = data_dir / "imported"
        self.attach_dir.mkdir(parents=True, exist_ok=True)
        self.imported_dir.mkdir(parents=True, exist_ok=True)
        self._row_cache: dict[str, list[dict[str, str]]] = {}
        self._mtime: dict[str, float] = {}
        self._export_follow = False
        self.ensure_files()
        self._export_follow = True

    def path(self, kind: str) -> Path:
        return self.data_dir / f"{kind}.csv"

    def ensure_files(self) -> None:
        for kind, headers in HEADERS.items():
            p = self.path(kind)
            if not p.exists():
                write_csv(p, headers, [])
        self.migrate_sections()
        self.migrate_task_images_to_info()
        self.migrate_info_point_columns()
        self.migrate_project_columns()
        (self.attach_dir / "icons").mkdir(parents=True, exist_ok=True)

    def migrate_sections(self) -> None:
        """Ensure General per project; backfill task.section_id from section_path."""
        projects = self.load("projects")
        sections = self.load("sections")
        tasks = self.load("tasks")
        changed_sec = False
        changed_tasks = False

        by_proj: dict[str, list[dict]] = {}
        for s in sections:
            by_proj.setdefault(s.get("project_id") or "", []).append(s)

        general_by_proj: dict[str, str] = {}
        for p in projects:
            pid = p.get("id") or ""
            if not pid:
                continue
            gen = next(
                (
                    s
                    for s in by_proj.get(pid, [])
                    if (s.get("title") or "").strip().lower() == GENERAL_SECTION.lower()
                ),
                None,
            )
            if not gen:
                gen = {
                    "id": next_id("SEC_", sections),
                    "project_id": pid,
                    "title": GENERAL_SECTION,
                    "sort_order": "0",
                }
                sections.append(gen)
                by_proj.setdefault(pid, []).append(gen)
                changed_sec = True
            general_by_proj[pid] = gen["id"]

        def find_or_add(pid: str, title: str) -> str:
            nonlocal changed_sec
            title = (title or "").strip() or GENERAL_SECTION
            if title.lower() == GENERAL_SECTION.lower():
                return general_by_proj[pid]
            for s in by_proj.get(pid, []):
                if (s.get("title") or "").strip().lower() == title.lower():
                    return s["id"]
            row = {
                "id": next_id("SEC_", sections),
                "project_id": pid,
                "title": title,
                "sort_order": next_sort(by_proj.get(pid, [])),
            }
            sections.append(row)
            by_proj.setdefault(pid, []).append(row)
            changed_sec = True
            return row["id"]

        for t in tasks:
            pid = t.get("project_id") or ""
            if not pid or pid not in general_by_proj:
                continue
            sid = (t.get("section_id") or "").strip()
            if sid and any(s.get("id") == sid for s in by_proj.get(pid, [])):
                # keep section_path in sync
                sec = next(s for s in by_proj[pid] if s["id"] == sid)
                title = sec.get("title") or GENERAL_SECTION
                path = "" if title.lower() == GENERAL_SECTION.lower() else title
                if (t.get("section_path") or "") != path:
                    t["section_path"] = path
                    changed_tasks = True
                continue
            path = (t.get("section_path") or "").strip()
            new_sid = find_or_add(pid, path) if path else general_by_proj[pid]
            if (t.get("section_id") or "") != new_sid:
                t["section_id"] = new_sid
                changed_tasks = True
            sec = next(s for s in by_proj[pid] if s["id"] == new_sid)
            title = sec.get("title") or GENERAL_SECTION
            want_path = "" if title.lower() == GENERAL_SECTION.lower() else title
            if (t.get("section_path") or "") != want_path:
                t["section_path"] = want_path
                changed_tasks = True

        if changed_sec:
            self.save("sections", sections)
        if changed_tasks:
            self.save("tasks", tasks)
        else:
            # Ensure section_id column exists on disk even when no row values changed
            path = self.path("tasks")
            if path.exists():
                with path.open("r", encoding="utf-8-sig", newline="") as f:
                    first = f.readline()
                if "section_id" not in first:
                    self.save("tasks", tasks)

    def migrate_task_images_to_info(self) -> None:
        """Reparent task/project image attachments onto new info points (once)."""
        atts = self.load("attachments")
        changed = False
        for a in atts:
            if (a.get("kind") or "").strip() != "image":
                continue
            pt = (a.get("parent_type") or "").strip()
            pid = (a.get("parent_id") or "").strip()
            if pt not in {"task", "project"} or not pid:
                continue
            title = (a.get("description") or "").strip() or "Image"
            created = self.upsert_info(
                {
                    "parent_type": pt,
                    "parent_id": pid,
                    "title": title,
                    "body": "",
                    "emoji": "ℹ️",
                }
            )
            info = created.get("info") or {}
            iid = info.get("id") or ""
            if not iid:
                continue
            a["parent_type"] = "info"
            a["parent_id"] = iid
            changed = True
        if changed:
            self.save("attachments", atts)

    def ensure_general_section(
        self, project_id: str, sections: list[dict] | None = None
    ) -> dict:
        rows = sections if sections is not None else self.load("sections")
        for s in rows:
            if (
                s.get("project_id") == project_id
                and (s.get("title") or "").strip().lower() == GENERAL_SECTION.lower()
            ):
                return s
        row = {
            "id": next_id("SEC_", rows),
            "project_id": project_id,
            "title": GENERAL_SECTION,
            "sort_order": "0",
        }
        rows.append(row)
        if sections is None:
            self.save("sections", rows)
        return row

    def find_or_create_section(self, project_id: str, title: str) -> dict:
        title = (title or "").strip() or GENERAL_SECTION
        rows = self.load("sections")
        for s in rows:
            if (
                s.get("project_id") == project_id
                and (s.get("title") or "").strip().lower() == title.lower()
            ):
                return s
        if title.lower() == GENERAL_SECTION.lower():
            row = {
                "id": next_id("SEC_", rows),
                "project_id": project_id,
                "title": GENERAL_SECTION,
                "sort_order": "0",
            }
        else:
            row = {
                "id": next_id("SEC_", rows),
                "project_id": project_id,
                "title": title,
                "sort_order": next_sort(
                    [s for s in rows if s.get("project_id") == project_id]
                ),
            }
        rows.append(row)
        self.save("sections", rows)
        return row

    def _section_path_for(self, section: dict) -> str:
        title = (section.get("title") or "").strip()
        if not title or title.lower() == GENERAL_SECTION.lower():
            return ""
        return title

    def resolve_task_section(self, project_id: str, payload: dict) -> tuple[str, str]:
        """Return (section_id, section_path) for a task payload."""
        sid = (payload.get("section_id") or "").strip()
        if sid:
            sec = self.find("sections", sid)
            if sec and sec.get("project_id") == project_id:
                return sid, self._section_path_for(sec)
        path = (payload.get("section_path") or "").strip()
        if path:
            sec = self.find_or_create_section(project_id, path)
            return sec["id"], self._section_path_for(sec)
        gen = self.ensure_general_section(project_id)
        return gen["id"], ""

    def load(self, kind: str) -> list[dict[str, str]]:
        path = self.path(kind)
        if not TASK_STORE_CACHE:
            return read_csv(path)
        try:
            mt = path.stat().st_mtime if path.exists() else -1.0
        except OSError:
            mt = -1.0
        cached = self._row_cache.get(kind)
        if cached is not None and self._mtime.get(kind) == mt:
            return [dict(r) for r in cached]
        rows = read_csv(path)
        self._row_cache[kind] = rows
        self._mtime[kind] = mt
        return [dict(r) for r in rows]

    def save(self, kind: str, rows: list[dict]) -> None:
        headers = HEADERS[kind]
        path = self.path(kind)
        old_rows = None
        if self._export_follow and kind != "projects":
            cached = self._row_cache.get(kind) if TASK_STORE_CACHE else None
            old_rows = (
                [dict(r) for r in cached] if cached is not None else read_csv(path)
            )
        write_csv(path, headers, rows)
        copied = [{h: str(r.get(h, "") or "") for h in headers} for r in rows]
        if TASK_STORE_CACHE:
            self._row_cache[kind] = copied
            try:
                self._mtime[kind] = path.stat().st_mtime
            except OSError:
                self._mtime[kind] = 0.0
        if old_rows is not None:
            self._refresh_exports_after_save(kind, old_rows, copied)

    def state(self) -> dict[str, Any]:
        projects = self.load("projects")
        sections = self.load("sections")
        tasks = self.load("tasks")
        infos = self.load("info_points")
        attachments = self.load("attachments")
        open_n = sum(
            1
            for t in tasks
            if t.get("active", "1") != "0"
            and t.get("kind") != "habitual"
            and t.get("emoji") != "✅"
        )
        return {
            "ok": True,
            "projects": projects,
            "sections": sections,
            "tasks": tasks,
            "info_points": infos,
            "attachments": attachments,
            "counts": {
                "projects": len(projects),
                "sections": len(sections),
                "tasks": len(tasks),
                "info": len(infos),
                "attachments": len(attachments),
                "open": open_n,
            },
            "status_emojis": STATUS_EMOJIS,
        }

    def find(self, kind: str, rid: str) -> dict | None:
        for r in self.load(kind):
            if r.get("id") == rid:
                return r
        return None

    # --- projects ---
    def upsert_project(self, payload: dict) -> dict:
        rows = self.load("projects")
        rid = (payload.get("id") or "").strip()
        title = (payload.get("title") or "").strip()
        filt = (payload.get("filter") or "work").strip().lower()
        if filt not in VALID_FILTERS:
            return {"ok": False, "error": "invalid filter"}
        if not title:
            return {"ok": False, "error": "title required"}
        if rid:
            out = []
            found = False
            for r in rows:
                if r["id"] == rid:
                    r = {
                        **r,
                        "title": title,
                        "filter": filt,
                        "section_path": (payload.get("section_path") or "").strip(),
                        "active": (payload.get("active") or r.get("active") or "1"),
                    }
                    found = True
                out.append(r)
            if not found:
                return {"ok": False, "error": "project not found"}
            self.save("projects", out)
            self.ensure_general_section(rid)
            self.sync_project_json_files()
            return {"ok": True, "project": next(x for x in out if x["id"] == rid)}
        row = {
            "id": next_id("PROJ_", rows),
            "title": title,
            "filter": filt,
            "section_path": (payload.get("section_path") or "").strip(),
            "sort_order": next_sort(rows),
            "active": "1",
            "created_at": now_stamp(),
            "icon_ref": "",
            "icon_color": "#ffffff",
            "icon_tint": "0",
        }
        rows.append(row)
        self.save("projects", rows)
        gen = self.ensure_general_section(row["id"])
        self.sync_project_json_files()
        return {"ok": True, "project": row, "section": gen}

    def _icon_file_path(self, icon_ref: str) -> Path | None:
        ref = (icon_ref or "").strip().replace("/", "\\")
        if not ref.lower().startswith("attachments\\icons\\"):
            return None
        path = (self.data_dir / ref).resolve()
        icons_root = (self.attach_dir / "icons").resolve()
        try:
            path.relative_to(icons_root)
        except ValueError:
            return None
        return path

    def _delete_icon_file(self, icon_ref: str) -> None:
        path = self._icon_file_path(icon_ref)
        if path and path.is_file():
            try:
                path.unlink()
            except OSError:
                pass

    @staticmethod
    def _icon_tint_for_ext(ext: str) -> str:
        """Never auto-flatten; user opts into flat tint via the Icon modal."""
        return "0"

    @staticmethod
    def _icon_tint_from_ref(icon_ref: str) -> str:
        """Legacy migration default: keep original colors (no auto white tint)."""
        return "0"

    def set_project_icon(self, project_id: str, image_bytes: bytes, ext: str) -> dict:
        """Save icon under attachments/icons/ and set projects.icon_ref (no info wrap)."""
        pid = (project_id or "").strip()
        if not pid:
            return {"ok": False, "error": "project id required"}
        rows = self.load("projects")
        target = next((r for r in rows if r.get("id") == pid), None)
        if not target:
            return {"ok": False, "error": "project not found"}
        safe_ext = (ext or ".png").lower()
        if safe_ext not in {".png", ".jpg", ".jpeg", ".webp", ".gif", ".svg"}:
            safe_ext = ".png"
        if safe_ext == ".jpeg":
            safe_ext = ".jpg"
        icons_dir = self.attach_dir / "icons"
        icons_dir.mkdir(parents=True, exist_ok=True)
        dest_name = f"{pid}{safe_ext}"
        dest = icons_dir / dest_name
        old_ref = (target.get("icon_ref") or "").strip()
        if old_ref:
            old_path = self._icon_file_path(old_ref)
            if old_path and old_path != dest.resolve() and old_path.is_file():
                try:
                    old_path.unlink()
                except OSError:
                    pass
        dest.write_bytes(image_bytes)
        ref = f"attachments\\icons\\{dest_name}"
        tint = self._icon_tint_for_ext(safe_ext)
        out = []
        for r in rows:
            if r["id"] == pid:
                r = {**r, "icon_ref": ref, "icon_tint": tint}
            out.append(r)
        self.save("projects", out)
        return {"ok": True, "project": next(x for x in out if x["id"] == pid)}

    def clear_project_icon(self, project_id: str) -> dict:
        pid = (project_id or "").strip()
        rows = self.load("projects")
        target = next((r for r in rows if r.get("id") == pid), None)
        if not target:
            return {"ok": False, "error": "project not found"}
        self._delete_icon_file(target.get("icon_ref") or "")
        out = []
        for r in rows:
            if r["id"] == pid:
                r = {**r, "icon_ref": ""}
            out.append(r)
        self.save("projects", out)
        return {"ok": True, "project": next(x for x in out if x["id"] == pid)}

    @staticmethod
    def _normalize_icon_color(color: str) -> str | None:
        """Return #rrggbb or None if invalid. Empty/default → #ffffff."""
        raw = (color or "").strip().lower()
        if not raw or raw in {"default", "none", "white"}:
            return "#ffffff"
        if not raw.startswith("#"):
            raw = "#" + raw
        if re.fullmatch(r"#[0-9a-f]{3}", raw):
            return "#" + "".join(c * 2 for c in raw[1:])
        if re.fullmatch(r"#[0-9a-f]{6}", raw):
            return raw
        return None

    @staticmethod
    def _normalize_icon_tint(tint: Any) -> str | None:
        if tint is None:
            return None
        if isinstance(tint, bool):
            return "1" if tint else "0"
        raw = str(tint).strip().lower()
        if raw in {"1", "true", "yes", "on"}:
            return "1"
        if raw in {"0", "false", "no", "off"}:
            return "0"
        return None

    def set_project_icon_color(
        self,
        project_id: str,
        color: str | None = None,
        tint: Any = None,
    ) -> dict:
        """Set icon_color and/or icon_tint; does not touch the icon file.

        Flat tint applies only when icon_tint=1 (SVG by default). Raster/3D
        icons keep native colors unless tint is explicitly enabled.
        """
        pid = (project_id or "").strip()
        if not pid:
            return {"ok": False, "error": "project id required"}
        rows = self.load("projects")
        target = next((r for r in rows if r.get("id") == pid), None)
        if not target:
            return {"ok": False, "error": "project not found"}
        updates: dict[str, str] = {}
        if tint is not None:
            normalized_tint = self._normalize_icon_tint(tint)
            if normalized_tint is None:
                return {"ok": False, "error": "invalid tint (use 0/1)"}
            updates["icon_tint"] = normalized_tint
        if color is not None:
            normalized = self._normalize_icon_color(color)
            if normalized is None:
                return {"ok": False, "error": "invalid color (use #rgb or #rrggbb)"}
            updates["icon_color"] = normalized
        if not updates:
            return {"ok": True, "project": target}
        out = []
        for r in rows:
            if r["id"] == pid:
                r = {**r, **updates}
            out.append(r)
        self.save("projects", out)
        return {"ok": True, "project": next(x for x in out if x["id"] == pid)}

    def set_project_status(self, project_id: str, status: str) -> dict:
        pid = (project_id or "").strip()
        if not pid:
            return {"ok": False, "error": "project id required"}
        normalized = normalize_waiting_status(status)
        if normalized is None:
            return {"ok": False, "error": "invalid status"}
        rows = self.load("projects")
        out = []
        found = None
        for r in rows:
            if r.get("id") == pid:
                r = {**r, "status": normalized}
                found = r
            out.append(r)
        if not found:
            return {"ok": False, "error": "project not found"}
        self.save("projects", out)
        self.write_project_json(pid)
        return {"ok": True, "project": found}

    def delete_project(self, project_id: str) -> dict:
        tasks = self.load("tasks")
        task_ids = {t["id"] for t in tasks if t.get("project_id") == project_id}
        infos = self.load("info_points")
        info_ids = set()
        info_out = []
        for i in infos:
            drop = (
                i.get("parent_type") == "project" and i.get("parent_id") == project_id
            ) or (i.get("parent_type") == "task" and i.get("parent_id") in task_ids)
            if drop:
                info_ids.add(i["id"])
            else:
                info_out.append(i)
        self.save("info_points", info_out)
        self._purge_attachments(
            lambda a: (
                a.get("parent_type") == "project" and a.get("parent_id") == project_id
            )
            or (a.get("parent_type") == "task" and a.get("parent_id") in task_ids)
            or (a.get("parent_type") == "info" and a.get("parent_id") in info_ids)
        )
        proj = self.find("projects", project_id)
        if proj:
            self._delete_icon_file(proj.get("icon_ref") or "")
        self.save("tasks", [t for t in tasks if t.get("project_id") != project_id])
        self.save(
            "sections",
            [s for s in self.load("sections") if s.get("project_id") != project_id],
        )
        self.save(
            "projects", [p for p in self.load("projects") if p.get("id") != project_id]
        )
        self.sync_project_json_files()
        return {"ok": True}

    # --- sections ---
    def upsert_section(self, payload: dict) -> dict:
        rows = self.load("sections")
        rid = (payload.get("id") or "").strip()
        title = (payload.get("title") or "").strip()
        project_id = (payload.get("project_id") or "").strip()
        if not title:
            return {"ok": False, "error": "title required"}
        if rid:
            out = []
            found = None
            for r in rows:
                if r["id"] == rid:
                    if (
                        r.get("title") or ""
                    ).strip().lower() == GENERAL_SECTION.lower():
                        return {"ok": False, "error": "cannot rename General"}
                    if title.lower() == GENERAL_SECTION.lower():
                        return {"ok": False, "error": "cannot rename to General"}
                    r = {**r, "title": title}
                    found = r
                out.append(r)
            if not found:
                return {"ok": False, "error": "section not found"}
            self.save("sections", out)
            # mirror title onto tasks.section_path
            tasks = self.load("tasks")
            path = self._section_path_for(found)
            changed = False
            for t in tasks:
                if t.get("section_id") == rid and (t.get("section_path") or "") != path:
                    t["section_path"] = path
                    changed = True
            if changed:
                self.save("tasks", tasks)
            return {"ok": True, "section": found}
        if not project_id:
            return {"ok": False, "error": "project_id required"}
        if not self.find("projects", project_id):
            return {"ok": False, "error": "project not found"}
        if title.lower() == GENERAL_SECTION.lower():
            return {"ok": True, "section": self.ensure_general_section(project_id)}
        for s in rows:
            if (
                s.get("project_id") == project_id
                and (s.get("title") or "").strip().lower() == title.lower()
            ):
                return {"ok": True, "section": s}
        row = {
            "id": next_id("SEC_", rows),
            "project_id": project_id,
            "title": title,
            "sort_order": next_sort(
                [s for s in rows if s.get("project_id") == project_id]
            ),
        }
        rows.append(row)
        self.save("sections", rows)
        return {"ok": True, "section": row}

    def delete_section(self, section_id: str) -> dict:
        rows = self.load("sections")
        sec = next((s for s in rows if s.get("id") == section_id), None)
        if not sec:
            return {"ok": False, "error": "section not found"}
        if (sec.get("title") or "").strip().lower() == GENERAL_SECTION.lower():
            return {"ok": False, "error": "cannot delete General"}
        gen = self.ensure_general_section(sec.get("project_id") or "")
        tasks = self.load("tasks")
        changed = False
        for t in tasks:
            if t.get("section_id") == section_id:
                t["section_id"] = gen["id"]
                t["section_path"] = ""
                changed = True
        if changed:
            self.save("tasks", tasks)
        self.save("sections", [s for s in rows if s.get("id") != section_id])
        return {"ok": True}

    def set_section_status(self, section_id: str, status: str) -> dict:
        sid = (section_id or "").strip()
        if not sid:
            return {"ok": False, "error": "section id required"}
        normalized = normalize_waiting_status(status)
        if normalized is None:
            return {"ok": False, "error": "invalid status"}
        rows = self.load("sections")
        out = []
        found = None
        for r in rows:
            if r.get("id") == sid:
                r = {**r, "status": normalized}
                found = r
            out.append(r)
        if not found:
            return {"ok": False, "error": "section not found"}
        self.save("sections", out)
        return {"ok": True, "section": found}

    # --- tasks ---
    def upsert_task(self, payload: dict) -> dict:
        rows = self.load("tasks")
        rid = (payload.get("id") or "").strip()
        title = (payload.get("title") or "").strip()
        if not title:
            return {"ok": False, "error": "title required"}
        filt = (payload.get("filter") or "work").strip().lower()
        kind = (payload.get("kind") or "punctual").strip().lower()
        recurrence = (payload.get("recurrence") or "").strip().lower()
        if filt not in VALID_FILTERS:
            return {"ok": False, "error": "invalid filter"}
        if kind not in VALID_KINDS:
            return {"ok": False, "error": "invalid kind"}
        if recurrence not in VALID_RECURRENCE:
            return {"ok": False, "error": "invalid recurrence"}
        if kind == "punctual":
            recurrence = ""
        emoji = (payload.get("emoji") or "").strip() or STATUS_EMOJIS["general"]
        if emoji.lower() == "none":
            emoji = ""
        project_id = (payload.get("project_id") or "").strip()
        if not project_id:
            return {"ok": False, "error": "project_id required"}
        section_id, section_path = self.resolve_task_section(project_id, payload)

        fields = {
            "project_id": project_id,
            "section_id": section_id,
            "title": title,
            "emoji": emoji,
            "kind": kind,
            "recurrence": recurrence,
            "due_date": (payload.get("due_date") or "").strip(),
            "next_due": (payload.get("next_due") or "").strip(),
            "section_path": section_path,
            "filter": filt,
            "active": (payload.get("active") or "1"),
        }
        if rid:
            out = []
            found = False
            for r in rows:
                if r["id"] == rid:
                    r = {**r, **fields}
                    if fields["emoji"] == "✅" and not r.get("completed_at"):
                        r["completed_at"] = now_stamp()
                    if fields["emoji"] != "✅":
                        r["completed_at"] = ""
                    # Editing a task clears the NEW import badge.
                    r["import_batch"] = ""
                    found = True
                out.append(r)
            if not found:
                return {"ok": False, "error": "task not found"}
            self.save("tasks", out)
            return {"ok": True, "task": next(x for x in out if x["id"] == rid)}
        row = {
            "id": next_id("TASK_", rows),
            **fields,
            "sort_order": next_sort(rows),
            "completed_at": "",
            "created_at": now_stamp(),
            "import_batch": "",
        }
        rows.append(row)
        self.save("tasks", rows)
        return {"ok": True, "task": row}

    def set_task_emoji(self, task_id: str, key_or_emoji: str) -> dict:
        key = (key_or_emoji or "").strip()
        if key in STATUS_EMOJIS:
            emoji = STATUS_EMOJIS[key]
        elif key.lower() == "none":
            emoji = ""
        else:
            emoji = key
        # Reject accidental literal "none" from older clients
        if (emoji or "").strip().lower() == "none":
            emoji = ""
        # "none" is an empty emoji; other empty keys are invalid
        if not emoji and key not in STATUS_EMOJIS and key.lower() != "none":
            return {"ok": False, "error": "emoji required"}
        rows = self.load("tasks")
        out = []
        found = None
        for r in rows:
            if r["id"] == task_id:
                if r.get("kind") == "habitual" and emoji == "✅":
                    # complete habit: advance next_due, reset to general
                    from_due = (
                        r.get("next_due") or r.get("due_date") or today()
                    ).strip()
                    r["next_due"] = self.advance_next_due(
                        r.get("recurrence") or "", from_due
                    )
                    r["emoji"] = STATUS_EMOJIS["general"]
                    r["completed_at"] = ""
                    r["import_batch"] = ""
                else:
                    r["emoji"] = emoji
                    r["completed_at"] = now_stamp() if emoji == "✅" else ""
                    if emoji == "✅":
                        r["import_batch"] = ""
                found = r
            out.append(r)
        if not found:
            return {"ok": False, "error": "task not found"}
        self.save("tasks", out)
        return {"ok": True, "task": found}

    def clear_task_import_batch(self, task_id: str) -> dict:
        rows = self.load("tasks")
        out = []
        found = None
        for r in rows:
            if r["id"] == task_id:
                r = {**r, "import_batch": ""}
                found = r
            out.append(r)
        if not found:
            return {"ok": False, "error": "task not found"}
        self.save("tasks", out)
        return {"ok": True, "task": found}

    def clear_all_import_batches(self) -> dict:
        rows = self.load("tasks")
        cleared = 0
        out = []
        for r in rows:
            if (r.get("import_batch") or "").strip():
                r = {**r, "import_batch": ""}
                cleared += 1
            out.append(r)
        if cleared:
            self.save("tasks", out)
        return {"ok": True, "cleared": cleared}

    def delete_task(self, task_id: str) -> dict:
        infos = self.load("info_points")
        info_ids = set()
        info_out = []
        for i in infos:
            if i.get("parent_type") == "task" and i.get("parent_id") == task_id:
                info_ids.add(i["id"])
            else:
                info_out.append(i)
        self.save("info_points", info_out)
        self._purge_attachments(
            lambda a: (a.get("parent_type") == "task" and a.get("parent_id") == task_id)
            or (a.get("parent_type") == "info" and a.get("parent_id") in info_ids)
        )
        self.save("tasks", [t for t in self.load("tasks") if t.get("id") != task_id])
        return {"ok": True}

    def migrate_info_point_columns(self) -> None:
        """Rewrite info_points.csv when edit_lock / birth_date headers are missing."""
        path = self.path("info_points")
        if not path.exists():
            return
        with path.open("r", encoding="utf-8-sig", newline="") as f:
            reader = csv.reader(f)
            try:
                header = next(reader)
            except StopIteration:
                return
        want = HEADERS["info_points"]
        if list(header) == want:
            return
        rows = self.load("info_points")
        self.save("info_points", rows)

    def migrate_project_columns(self) -> None:
        """Rewrite projects.csv when icon columns are missing."""
        path = self.path("projects")
        if not path.exists():
            return
        with path.open("r", encoding="utf-8-sig", newline="") as f:
            reader = csv.reader(f)
            try:
                header = next(reader)
            except StopIteration:
                return
        want = HEADERS["projects"]
        if list(header) == want:
            return
        rows = self.load("projects")
        for r in rows:
            if not (r.get("icon_color") or "").strip():
                r["icon_color"] = "#ffffff"
            if not (r.get("icon_tint") or "").strip():
                r["icon_tint"] = self._icon_tint_from_ref(r.get("icon_ref") or "")
        self.save("projects", rows)

    # --- info ---
    def upsert_info(self, payload: dict) -> dict:
        rows = self.load("info_points")
        rid = (payload.get("id") or "").strip()
        title = (
            payload.get("text") or payload.get("title") or payload.get("body") or ""
        ).strip()
        if not title:
            return {"ok": False, "error": "text required"}
        parent_type = (payload.get("parent_type") or "").strip()
        parent_id = (payload.get("parent_id") or "").strip()
        if parent_type not in {"project", "task"} or not parent_id:
            return {"ok": False, "error": "parent required"}
        fields = {
            "parent_type": parent_type,
            "parent_id": parent_id,
            "title": title,
            "body": (payload.get("body") if "body" in payload else None),
            "emoji": (payload.get("emoji") or "ℹ️").strip() or "ℹ️",
            "section_path": (payload.get("section_path") or "").strip(),
        }
        if rid:
            out = []
            found = False
            for r in rows:
                if r["id"] == rid:
                    merged = {**r, **{k: v for k, v in fields.items() if v is not None}}
                    if fields["body"] is None:
                        merged["body"] = r.get("body") or ""
                    if "edit_lock" in payload:
                        merged["edit_lock"] = str(
                            payload.get("edit_lock") or ""
                        ).strip()
                    if "birth_date" in payload:
                        merged["birth_date"] = str(
                            payload.get("birth_date") or ""
                        ).strip()
                    r = merged
                    found = True
                out.append(r)
            if not found:
                return {"ok": False, "error": "info not found"}
            self.save("info_points", out)
            return {"ok": True, "info": next(x for x in out if x["id"] == rid)}
        siblings = [
            i
            for i in rows
            if i.get("parent_type") == parent_type and i.get("parent_id") == parent_id
        ]
        row = {
            "id": next_id("INFO_", rows),
            "parent_type": parent_type,
            "parent_id": parent_id,
            "title": title,
            "body": (payload.get("body") or "").strip() if "body" in payload else "",
            "emoji": fields["emoji"],
            "section_path": fields["section_path"],
            "sort_order": next_sort(siblings),
            "created_at": now_stamp(),
            "edit_lock": str(payload.get("edit_lock") or "").strip(),
            "birth_date": str(payload.get("birth_date") or "").strip(),
        }
        rows.append(row)
        self.save("info_points", rows)
        return {"ok": True, "info": row}

    def move_project(self, project_id: str, direction: int) -> dict:
        try:
            direction = parse_move_dir(direction)
        except ValueError as e:
            return {"ok": False, "error": str(e)}
        rows = self.load("projects")
        target = next((r for r in rows if r.get("id") == project_id), None)
        if not target:
            return {"ok": False, "error": "project not found"}
        if target.get("active", "1") == "0":
            return {"ok": False, "error": "project not active"}
        filt = target.get("filter") or ""

        def pred(r: dict, _t: dict) -> bool:
            return r.get("filter") == filt and r.get("active", "1") != "0"

        out, row, moved = swap_sort_order(rows, project_id, direction, pred)
        if moved:
            self.save("projects", out)
        return {"ok": True, "project": row or target, "moved": moved}

    def move_section(self, section_id: str, direction: int) -> dict:
        try:
            direction = parse_move_dir(direction)
        except ValueError as e:
            return {"ok": False, "error": str(e)}
        rows = self.load("sections")
        target = next((r for r in rows if r.get("id") == section_id), None)
        if not target:
            return {"ok": False, "error": "section not found"}
        if (target.get("title") or "").strip().lower() == GENERAL_SECTION.lower():
            return {"ok": False, "error": "General cannot be moved"}
        pid = target.get("project_id") or ""

        def pred(r: dict, _t: dict) -> bool:
            title = (r.get("title") or "").strip().lower()
            return r.get("project_id") == pid and title != GENERAL_SECTION.lower()

        out, row, moved = swap_sort_order(rows, section_id, direction, pred)
        if moved:
            self.save("sections", out)
        return {"ok": True, "section": row or target, "moved": moved}

    def move_task(self, task_id: str, direction: int) -> dict:
        try:
            direction = parse_move_dir(direction)
        except ValueError as e:
            return {"ok": False, "error": str(e)}
        rows = self.load("tasks")
        target = next((r for r in rows if r.get("id") == task_id), None)
        if not target:
            return {"ok": False, "error": "task not found"}
        if target.get("active", "1") == "0":
            return {"ok": False, "error": "task not active"}
        pid = target.get("project_id") or ""
        sid = target.get("section_id") or ""
        kind = (target.get("kind") or "punctual").strip().lower()

        def pred(r: dict, _t: dict) -> bool:
            if r.get("project_id") != pid:
                return False
            if (r.get("section_id") or "") != sid:
                return False
            if r.get("active", "1") == "0":
                return False
            rk = (r.get("kind") or "punctual").strip().lower()
            if kind == "habitual":
                return rk == "habitual"
            return rk == "punctual" and (r.get("emoji") or "").strip() != "✅"

        out, row, moved = swap_sort_order(rows, task_id, direction, pred)
        if moved:
            self.save("tasks", out)
        return {"ok": True, "task": row or target, "moved": moved}

    def move_info(self, info_id: str, direction: int) -> dict:
        """Swap sort_order with adjacent sibling (same parent). direction: -1 up, +1 down."""
        rid = (info_id or "").strip()
        try:
            direction = parse_move_dir(direction)
        except ValueError as e:
            return {"ok": False, "error": str(e)}
        rows = self.load("info_points")
        target = next((r for r in rows if r.get("id") == rid), None)
        if not target:
            return {"ok": False, "error": "info not found"}
        parent_type = target.get("parent_type") or ""
        parent_id = target.get("parent_id") or ""

        def pred(r: dict, _t: dict) -> bool:
            return (
                r.get("parent_type") == parent_type and r.get("parent_id") == parent_id
            )

        out, row, moved = swap_sort_order(rows, rid, direction, pred)
        if moved:
            self.save("info_points", out)
        siblings = sorted(
            [
                r
                for r in out
                if r.get("parent_type") == parent_type
                and r.get("parent_id") == parent_id
            ],
            key=sort_key_row,
        )
        return {
            "ok": True,
            "info": row or target,
            "moved": moved,
            "infos": siblings,
        }

    def delete_info(self, info_id: str) -> dict:
        self.save(
            "info_points",
            [i for i in self.load("info_points") if i.get("id") != info_id],
        )
        self._purge_attachments(
            lambda a: a.get("parent_type") == "info" and a.get("parent_id") == info_id
        )
        return {"ok": True}

    # --- attachments ---
    def add_attachment(
        self,
        parent_type: str,
        parent_id: str,
        kind: str,
        ref: str,
        description: str = "",
    ) -> dict:
        rows = self.load("attachments")
        row = {
            "id": next_id("ATT_", rows),
            "parent_type": parent_type,
            "parent_id": parent_id,
            "kind": kind,
            "ref": ref,
            "description": description or ref,
            "sort_order": next_sort(rows),
        }
        rows.append(row)
        self.save("attachments", rows)
        return {"ok": True, "attachment": row}

    def save_image_bytes(
        self,
        parent_type: str,
        parent_id: str,
        data: bytes,
        filename: str,
        description: str = "",
    ) -> dict:
        safe = re.sub(r"[^\w.\-]+", "_", filename) or "image.png"
        desc = (description or "").strip() or safe
        pt = (parent_type or "").strip()
        pid = (parent_id or "").strip()
        if pt in {"task", "project"} and pid:
            created = self.upsert_info(
                {
                    "parent_type": pt,
                    "parent_id": pid,
                    "title": desc,
                    "body": "",
                    "emoji": "ℹ️",
                }
            )
            info = created.get("info") or {}
            if info.get("id"):
                pt = "info"
                pid = info["id"]
        dest_name = f"{datetime.now().strftime('%Y%m%d-%H%M%S')}-{pid}-{safe}"
        dest = self.attach_dir / dest_name
        dest.write_bytes(data)
        ref = f"attachments\\{dest_name}"
        result = self.add_attachment(pt, pid, "image", ref, desc)
        if pt == "info":
            result["info_id"] = pid
            info_row = self.find("info_points", pid)
            if info_row:
                result["info"] = info_row
        return result

    def delete_attachment(self, att_id: str) -> dict:
        rows = self.load("attachments")
        keep = []
        for a in rows:
            if a.get("id") == att_id:
                self._delete_managed_file(a)
            else:
                keep.append(a)
        self.save("attachments", keep)
        return {"ok": True}

    def _delete_managed_file(self, att: dict) -> None:
        kind = (att.get("kind") or "").strip()
        ref = (att.get("ref") or "").strip().replace("/", "\\")
        if kind not in {"image", "text"} or not ref:
            return
        if ref.lower().startswith("attachments\\"):
            path = self.data_dir / Path(ref)
        else:
            return
        if path.is_file() and self.attach_dir.resolve() in path.resolve().parents:
            try:
                path.unlink()
            except OSError:
                pass

    def _purge_attachments(self, should_drop) -> None:
        keep = []
        for a in self.load("attachments"):
            if should_drop(a):
                self._delete_managed_file(a)
            else:
                keep.append(a)
        self.save("attachments", keep)

    def advance_next_due(self, recurrence: str, from_date: str) -> str:
        try:
            base = datetime.strptime(from_date[:10], "%Y-%m-%d")
        except ValueError:
            base = datetime.now()
        rec = (recurrence or "").lower()
        if rec == "daily":
            base += timedelta(days=1)
        elif rec == "weekly":
            base += timedelta(days=7)
        elif rec == "monthly":
            base = _add_months(base, 1)
        elif rec == "quarterly":
            base = _add_months(base, 3)
        elif rec == "biannual":
            base = _add_months(base, 6)
        elif rec == "yearly":
            base = _add_months(base, 12)
        elif rec == "every_2y":
            base = _add_months(base, 24)
        elif rec == "every_3y":
            base = _add_months(base, 36)
        elif rec == "every_5y":
            base = _add_months(base, 60)
        elif rec == "every_10y":
            base = _add_months(base, 120)
        else:
            base += timedelta(days=1)
        return base.strftime("%Y-%m-%d")

    def export_project(self, project_id: str, *, as_of: date | None = None) -> dict:
        """Nested snapshot of one project for an AI companion (no image bytes)."""
        pid = (project_id or "").strip()
        project = self.find("projects", pid)
        if not project:
            return {"ok": False, "error": "project not found"}
        today = as_of or date.today()
        sections = [s for s in self.load("sections") if s.get("project_id") == pid]
        tasks = [t for t in self.load("tasks") if t.get("project_id") == pid]
        task_ids = {t.get("id") or "" for t in tasks}
        section_ids = {s.get("id") or "" for s in sections}
        infos = [
            i
            for i in self.load("info_points")
            if (i.get("parent_type") == "project" and i.get("parent_id") == pid)
            or (i.get("parent_type") == "task" and i.get("parent_id") in task_ids)
        ]
        info_ids = {i.get("id") or "" for i in infos}
        atts = [
            a
            for a in self.load("attachments")
            if (a.get("parent_type") == "project" and a.get("parent_id") == pid)
            or (a.get("parent_type") == "task" and a.get("parent_id") in task_ids)
            or (a.get("parent_type") == "info" and a.get("parent_id") in info_ids)
        ]
        infos_by_parent: dict[tuple[str, str], list[dict]] = {}
        for i in infos:
            infos_by_parent.setdefault(
                ((i.get("parent_type") or ""), (i.get("parent_id") or "")), []
            ).append(i)
        atts_by_parent: dict[tuple[str, str], list[dict]] = {}
        for a in atts:
            atts_by_parent.setdefault(
                ((a.get("parent_type") or ""), (a.get("parent_id") or "")), []
            ).append(a)

        section_docs: list[dict] = []
        for s in sorted(sections, key=_export_sort_key):
            sid = s.get("id") or ""
            owned = [t for t in tasks if (t.get("section_id") or "") == sid]
            sec_doc: dict[str, Any] = {
                "id": sid,
                "title": (s.get("title") or "").strip() or GENERAL_SECTION,
                "tasks": [
                    _export_task(t, infos_by_parent, atts_by_parent, today)
                    for t in sorted(owned, key=_export_sort_key)
                ],
            }
            if (s.get("status") or "").strip() == "waiting":
                sec_doc["status"] = "waiting"
            section_docs.append(sec_doc)
        orphans = [t for t in tasks if (t.get("section_id") or "") not in section_ids]
        if orphans:
            section_docs.append(
                {
                    "id": "",
                    "title": "Other",
                    "tasks": [
                        _export_task(t, infos_by_parent, atts_by_parent, today)
                        for t in sorted(orphans, key=_export_sort_key)
                    ],
                }
            )

        proj_doc: dict[str, Any] = {
            "id": project.get("id") or "",
            "title": (project.get("title") or "").strip(),
            "filter": (project.get("filter") or "").strip(),
        }
        if (project.get("status") or "").strip() == "waiting":
            proj_doc["status"] = "waiting"
        path = (project.get("section_path") or "").strip()
        if path:
            proj_doc["section_path"] = path
        created = (project.get("created_at") or "").strip()
        if created:
            proj_doc["created_at"] = created
        icon = (project.get("icon_ref") or "").replace("\\", "/").strip()
        if icon:
            proj_doc["icon_file"] = icon
        notes = [
            _export_info(i, atts_by_parent, today)
            for i in sorted(
                infos_by_parent.get(("project", pid), []), key=_export_sort_key
            )
        ]
        if notes:
            proj_doc["info"] = notes
        proj_atts = _export_attachments(atts_by_parent.get(("project", pid), []))
        if proj_atts:
            proj_doc["attachments"] = proj_atts
        proj_doc["sections"] = section_docs

        task_docs = [t for s in section_docs for t in s["tasks"]]
        info_n = len(notes) + sum(len(t.get("info") or []) for t in task_docs)
        att_n = len(proj_atts) + sum(len(t.get("attachments") or []) for t in task_docs)
        att_n += sum(len(i.get("attachments") or []) for i in notes)
        att_n += sum(
            len(i.get("attachments") or [])
            for t in task_docs
            for i in (t.get("info") or [])
        )
        document = {
            "schema": "tasks.project",
            "schema_version": 1,
            "exported_at": datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
            "about": (
                "Full snapshot of one project from the local Tasks app. "
                "Use this file as the whole context when discussing the project. "
                "sections[].tasks includes open, waiting, important, doubtful, and completed tasks. "
                "info points are notes on the project or on a task. "
                "attachments name image files; image bytes are not included."
            ),
            "status_legend": {
                glyph: name for name, glyph in STATUS_EMOJIS.items() if glyph
            },
            "project": proj_doc,
            "counts": {
                "sections": len(section_docs),
                "tasks": len(task_docs),
                "open_tasks": sum(1 for t in task_docs if t.get("open")),
                "completed_tasks": sum(
                    1 for t in task_docs if t.get("status") == "done"
                ),
                "info_points": info_n,
                "attachments": att_n,
            },
        }
        return {
            "ok": True,
            "filename": project_export_filename(project),
            "document": document,
        }

    def project_json_dir(self) -> Path:
        folder = self.data_dir / "exports"
        folder.mkdir(parents=True, exist_ok=True)
        return folder

    def write_project_json(self, project_id: str, *, update_index: bool = True) -> dict:
        """Write one project's companion JSON and drop a stale filename for that id."""
        result = self.export_project(project_id)
        if not result.get("ok"):
            return result
        folder = self.project_json_dir()
        filename = result["filename"]
        token = re.sub(r"[^A-Za-z0-9_-]+", "", (project_id or "").strip())
        if token:
            for old in folder.glob(f"*__{token}.json"):
                if old.name != filename:
                    old.unlink(missing_ok=True)
        text = json.dumps(result["document"], ensure_ascii=False, indent=2)
        (folder / filename).write_text(text + "\n", encoding="utf-8")
        if update_index:
            self.write_projects_index()
        return result

    def write_projects_index(self) -> None:
        """Catalog of current projects. Rewritten when projects are added or removed."""
        entries = []
        for project in sorted(self.load("projects"), key=_export_sort_key):
            pid = (project.get("id") or "").strip()
            if not pid:
                continue
            entries.append(
                {
                    "id": pid,
                    "title": (project.get("title") or "").strip(),
                    "filter": (project.get("filter") or "").strip(),
                    "file": project_export_filename(project),
                }
            )
        document = {
            "schema": "tasks.projects",
            "schema_version": 1,
            "projects": entries,
        }
        path = self.project_json_dir() / "projects.json"
        path.write_text(
            json.dumps(document, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )

    def sync_project_json_files(self) -> None:
        """Rewrite every project JSON and delete files for projects that are gone."""
        folder = self.project_json_dir()
        keep: set[str] = set()
        for project in self.load("projects"):
            pid = (project.get("id") or "").strip()
            if not pid:
                continue
            result = self.write_project_json(pid, update_index=False)
            if result.get("ok"):
                keep.add(result["filename"])
        for path in folder.glob("*.json"):
            if path.name == "projects.json" or path.name in keep:
                continue
            path.unlink(missing_ok=True)
        self.write_projects_index()

    def _refresh_exports_after_save(
        self, kind: str, old_rows: list[dict], new_rows: list[dict]
    ) -> None:
        for pid in self._changed_project_ids(kind, old_rows, new_rows):
            if self.find("projects", pid):
                self.write_project_json(pid)
            else:
                token = re.sub(r"[^A-Za-z0-9_-]+", "", pid)
                if not token:
                    continue
                for old in self.project_json_dir().glob(f"*__{token}.json"):
                    old.unlink(missing_ok=True)
                self.write_projects_index()

    def _changed_project_ids(
        self, kind: str, old_rows: list[dict], new_rows: list[dict]
    ) -> set[str]:
        def sig(row: dict) -> tuple:
            return tuple(sorted((str(k), str(v or "")) for k, v in row.items()))

        old_by = {r.get("id") or "": r for r in old_rows}
        new_by = {r.get("id") or "": r for r in new_rows}
        changed: set[str] = set()
        for rid in set(old_by) | set(new_by):
            if not rid or sig(old_by.get(rid) or {}) == sig(new_by.get(rid) or {}):
                continue
            row = new_by.get(rid) or old_by.get(rid) or {}
            pid = self._project_id_of_saved_row(kind, row)
            if pid:
                changed.add(pid)
        return changed

    def _project_id_of_saved_row(self, kind: str, row: dict) -> str:
        if kind in {"tasks", "sections"}:
            return (row.get("project_id") or "").strip()
        parent_type = (row.get("parent_type") or "").strip()
        parent_id = (row.get("parent_id") or "").strip()
        if parent_type == "project":
            return parent_id
        if parent_type == "task":
            task = self.find("tasks", parent_id)
            return (task or {}).get("project_id") or ""
        if parent_type == "info":
            info = self.find("info_points", parent_id)
            if not info:
                return ""
            return self._project_id_of_saved_row("info_points", info)
        return ""

    def ensure_inbox_project(self, filt: str) -> dict:
        title = INBOX_TITLES.get(filt, "Inbox")
        for p in self.load("projects"):
            if (
                p.get("filter") == filt
                and (p.get("title") or "").strip().lower() == title.lower()
            ):
                self.ensure_general_section(p["id"])
                return p
        return self.upsert_project({"title": title, "filter": filt})["project"]

    def spawn_personal_from_habit(self, task_id: str) -> dict:
        """Create a punctual Personal-inbox / General task from a missed habit."""
        habit = self.find("tasks", task_id)
        if not habit:
            return {"ok": False, "error": "task not found"}
        if habit.get("kind") != "habitual":
            return {"ok": False, "error": "not a habitual task"}
        title = (habit.get("title") or "").strip()
        if not title:
            return {"ok": False, "error": "habit has no title"}
        inbox = self.ensure_inbox_project("personal")
        gen = self.ensure_general_section(inbox["id"])
        return self.upsert_task(
            {
                "title": title,
                "project_id": inbox["id"],
                "section_id": gen["id"],
                "filter": "personal",
                "kind": "punctual",
                "emoji": STATUS_EMOJIS["general"],
            }
        )


def _add_months(dt: datetime, months: int) -> datetime:
    y = dt.year + (dt.month - 1 + months) // 12
    m = (dt.month - 1 + months) % 12 + 1
    d = min(
        dt.day,
        [
            31,
            29 if y % 4 == 0 and (y % 100 != 0 or y % 400 == 0) else 28,
            31,
            30,
            31,
            30,
            31,
            31,
            30,
            31,
            30,
            31,
        ][m - 1],
    )
    return dt.replace(year=y, month=m, day=d)
