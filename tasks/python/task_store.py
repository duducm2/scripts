"""CSV store for the Tasks web app (projects / tasks / info / attachments)."""

from __future__ import annotations

import csv
import re
import shutil
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

HEADERS = {
    "projects": ["id", "title", "filter", "section_path", "sort_order", "active", "created_at"],
    "tasks": [
        "id",
        "project_id",
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
    ],
    "attachments": ["id", "parent_type", "parent_id", "kind", "ref", "description", "sort_order"],
}

STATUS_EMOJIS = {
    "general": "🔲",
    "waiting": "⏳",
    "important": "⚡",
    "done": "✅",
    "doubt": "❓",
}

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

INBOX_TITLES = {
    "work": "Work inbox",
    "personal": "Personal inbox",
    "habits": "Habits & Health",
}


def now_stamp() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def today() -> str:
    return datetime.now().strftime("%Y-%m-%d")


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


class TaskStore:
    def __init__(self, data_dir: Path):
        self.data_dir = data_dir
        self.attach_dir = data_dir / "attachments"
        self.imported_dir = data_dir / "imported"
        self.attach_dir.mkdir(parents=True, exist_ok=True)
        self.imported_dir.mkdir(parents=True, exist_ok=True)
        self.ensure_files()

    def path(self, kind: str) -> Path:
        return self.data_dir / f"{kind}.csv"

    def ensure_files(self) -> None:
        for kind, headers in HEADERS.items():
            p = self.path(kind)
            if not p.exists():
                write_csv(p, headers, [])

    def load(self, kind: str) -> list[dict[str, str]]:
        return read_csv(self.path(kind))

    def save(self, kind: str, rows: list[dict]) -> None:
        write_csv(self.path(kind), HEADERS[kind], rows)

    def state(self) -> dict[str, Any]:
        projects = self.load("projects")
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
            "tasks": tasks,
            "info_points": infos,
            "attachments": attachments,
            "counts": {
                "projects": len(projects),
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
            return {"ok": True, "project": next(x for x in out if x["id"] == rid)}
        row = {
            "id": next_id("PROJ_", rows),
            "title": title,
            "filter": filt,
            "section_path": (payload.get("section_path") or "").strip(),
            "sort_order": next_sort(rows),
            "active": "1",
            "created_at": now_stamp(),
        }
        rows.append(row)
        self.save("projects", rows)
        return {"ok": True, "project": row}

    def delete_project(self, project_id: str) -> dict:
        tasks = self.load("tasks")
        task_ids = {t["id"] for t in tasks if t.get("project_id") == project_id}
        infos = self.load("info_points")
        info_ids = set()
        info_out = []
        for i in infos:
            drop = (i.get("parent_type") == "project" and i.get("parent_id") == project_id) or (
                i.get("parent_type") == "task" and i.get("parent_id") in task_ids
            )
            if drop:
                info_ids.add(i["id"])
            else:
                info_out.append(i)
        self.save("info_points", info_out)
        self._purge_attachments(
            lambda a: (a.get("parent_type") == "project" and a.get("parent_id") == project_id)
            or (a.get("parent_type") == "task" and a.get("parent_id") in task_ids)
            or (a.get("parent_type") == "info" and a.get("parent_id") in info_ids)
        )
        self.save("tasks", [t for t in tasks if t.get("project_id") != project_id])
        self.save("projects", [p for p in self.load("projects") if p.get("id") != project_id])
        return {"ok": True}

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
        project_id = (payload.get("project_id") or "").strip()
        if not project_id:
            return {"ok": False, "error": "project_id required"}

        fields = {
            "project_id": project_id,
            "title": title,
            "emoji": emoji,
            "kind": kind,
            "recurrence": recurrence,
            "due_date": (payload.get("due_date") or "").strip(),
            "next_due": (payload.get("next_due") or "").strip(),
            "section_path": (payload.get("section_path") or "").strip(),
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
        }
        rows.append(row)
        self.save("tasks", rows)
        return {"ok": True, "task": row}

    def set_task_emoji(self, task_id: str, key_or_emoji: str) -> dict:
        key = (key_or_emoji or "").strip()
        emoji = STATUS_EMOJIS.get(key, key)
        if not emoji:
            return {"ok": False, "error": "emoji required"}
        rows = self.load("tasks")
        out = []
        found = None
        for r in rows:
            if r["id"] == task_id:
                if r.get("kind") == "habitual" and emoji == "✅":
                    # complete habit: advance next_due, reset to general
                    from_due = (r.get("next_due") or r.get("due_date") or today()).strip()
                    r["next_due"] = self.advance_next_due(r.get("recurrence") or "", from_due)
                    r["emoji"] = STATUS_EMOJIS["general"]
                    r["completed_at"] = ""
                else:
                    r["emoji"] = emoji
                    r["completed_at"] = now_stamp() if emoji == "✅" else ""
                found = r
            out.append(r)
        if not found:
            return {"ok": False, "error": "task not found"}
        self.save("tasks", out)
        return {"ok": True, "task": found}

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

    # --- info ---
    def upsert_info(self, payload: dict) -> dict:
        rows = self.load("info_points")
        rid = (payload.get("id") or "").strip()
        title = (payload.get("title") or "").strip()
        if not title:
            return {"ok": False, "error": "title required"}
        parent_type = (payload.get("parent_type") or "").strip()
        parent_id = (payload.get("parent_id") or "").strip()
        if parent_type not in {"project", "task"} or not parent_id:
            return {"ok": False, "error": "parent required"}
        fields = {
            "parent_type": parent_type,
            "parent_id": parent_id,
            "title": title,
            "body": payload.get("body") if payload.get("body") is not None else "",
            "emoji": (payload.get("emoji") or "ℹ️").strip() or "ℹ️",
            "section_path": (payload.get("section_path") or "").strip(),
        }
        if rid:
            out = []
            found = False
            for r in rows:
                if r["id"] == rid:
                    # preserve body if not provided
                    body = fields["body"]
                    if "body" not in payload:
                        body = r.get("body") or ""
                    r = {**r, **fields, "body": body}
                    found = True
                out.append(r)
            if not found:
                return {"ok": False, "error": "info not found"}
            self.save("info_points", out)
            return {"ok": True, "info": next(x for x in out if x["id"] == rid)}
        row = {
            "id": next_id("INFO_", rows),
            **fields,
            "sort_order": next_sort(rows),
            "created_at": now_stamp(),
        }
        rows.append(row)
        self.save("info_points", rows)
        return {"ok": True, "info": row}

    def delete_info(self, info_id: str) -> dict:
        self.save("info_points", [i for i in self.load("info_points") if i.get("id") != info_id])
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
        self, parent_type: str, parent_id: str, data: bytes, filename: str, description: str = ""
    ) -> dict:
        safe = re.sub(r"[^\w.\-]+", "_", filename) or "image.png"
        dest_name = f"{datetime.now().strftime('%Y%m%d-%H%M%S')}-{parent_id}-{safe}"
        dest = self.attach_dir / dest_name
        dest.write_bytes(data)
        ref = f"attachments\\{dest_name}"
        return self.add_attachment(parent_type, parent_id, "image", ref, description or safe)

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

    def ensure_inbox_project(self, filt: str) -> dict:
        title = INBOX_TITLES.get(filt, "Inbox")
        for p in self.load("projects"):
            if p.get("filter") == filt and (p.get("title") or "").strip().lower() == title.lower():
                return p
        return self.upsert_project({"title": title, "filter": filt})["project"]


def _add_months(dt: datetime, months: int) -> datetime:
    y = dt.year + (dt.month - 1 + months) // 12
    m = (dt.month - 1 + months) % 12 + 1
    d = min(dt.day, [31, 29 if y % 4 == 0 and (y % 100 != 0 or y % 400 == 0) else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1])
    return dt.replace(year=y, month=m, day=d)
