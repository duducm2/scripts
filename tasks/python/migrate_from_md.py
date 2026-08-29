"""One-shot migrate work.md / punctual.md / habits.md into tasks CSV data."""

from __future__ import annotations

import argparse
import csv
import re
import shutil
from datetime import datetime
from pathlib import Path

INFO_EMOJI = "ℹ️"
TASK_EMOJIS = {"🔲", "⏳", "⚡", "✅", "❓"}
# Leading emoji / symbol at start of a content line
EMOJI_RE = re.compile(
    r"^("
    r"ℹ️|ℹ️️|"  # info variants
    r"🔲|⏳|⚡|✅|❓|🩺|"
    r"[\U0001F300-\U0001FAFF]"  # misc symbols/emoji
    r")\s*(.*)$"
)
HEADING_RE = re.compile(r"^(#{1,6})\s+(.*)$")
IMG_RE = re.compile(r"!\[([^\]]*)\]\(([^)]+)\)")
URL_RE = re.compile(r"^https?://\S+$", re.I)
PATH_RE = re.compile(r"^[A-Za-z]:\\|^\\\\")
LINK_STYLE = re.compile(r"^\[.+\]\(.+\)$")


def now_stamp() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def write_csv(path: Path, headers: list[str], rows: list[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=headers, lineterminator="\n")
        w.writeheader()
        for r in rows:
            w.writerow({h: r.get(h, "") for h in headers})


def next_id(prefix: str, n: int, pad: int = 4) -> str:
    return f"{prefix}{n:0{pad}d}"


def strip_md(line: str) -> str:
    return line.rstrip("\n").rstrip("\r")


RECURRENCE_MAP = {
    "daily": "daily",
    "weekly": "weekly",
    "monthly": "monthly",
    "quarterly": "quarterly",
    "bi-yearly": "biannual",
    "bi-yearly (every 6 months)": "biannual",
    "every 6 months": "biannual",
    "yearly": "yearly",
    "every 10 years": "every_10y",
    "every 10 years (documents)": "every_10y",
    "every 2 years": "every_2y",
    "every 3 years": "every_3y",
    "every 5 years": "every_5y",
}


def normalize_recurrence(heading: str) -> str:
    h = heading.strip().lower()
    for k, v in RECURRENCE_MAP.items():
        if h == k or h.startswith(k):
            return v
    if "daily" in h:
        return "daily"
    if "weekly" in h:
        return "weekly"
    if "monthly" in h:
        return "monthly"
    if "quarterly" in h:
        return "quarterly"
    if "bi-year" in h or "6 month" in h:
        return "biannual"
    if "every 10" in h:
        return "every_10y"
    if "every 2" in h:
        return "every_2y"
    if "every 3" in h:
        return "every_3y"
    if "every 5" in h:
        return "every_5y"
    if "yearly" in h or "year" in h:
        return "yearly"
    return ""


class Migrator:
    def __init__(self, data_dir: Path):
        self.data_dir = data_dir
        self.attach_dir = data_dir / "attachments"
        self.attach_dir.mkdir(parents=True, exist_ok=True)
        self.projects: list[dict] = []
        self.tasks: list[dict] = []
        self.infos: list[dict] = []
        self.attachments: list[dict] = []
        self._pid = 0
        self._tid = 0
        self._iid = 0
        self._aid = 0
        self._sort = 0

    def _so(self) -> str:
        self._sort += 10
        return str(self._sort)

    def add_project(self, title: str, filt: str, section_path: str = "") -> dict:
        self._pid += 1
        row = {
            "id": next_id("PROJ_", self._pid),
            "title": title.strip() or "Untitled",
            "filter": filt,
            "section_path": section_path,
            "sort_order": self._so(),
            "active": "1",
            "created_at": now_stamp(),
        }
        self.projects.append(row)
        return row

    def add_task(
        self,
        project_id: str,
        title: str,
        emoji: str,
        filt: str,
        kind: str,
        recurrence: str = "",
        section_path: str = "",
        due_date: str = "",
        next_due: str = "",
    ) -> dict:
        self._tid += 1
        if emoji == "✅":
            completed = now_stamp()
        else:
            completed = ""
        row = {
            "id": next_id("TASK_", self._tid),
            "project_id": project_id,
            "title": title.strip() or "(untitled)",
            "emoji": emoji or "🔲",
            "kind": kind,
            "recurrence": recurrence,
            "due_date": due_date,
            "next_due": next_due,
            "section_path": section_path,
            "filter": filt,
            "sort_order": self._so(),
            "completed_at": completed,
            "created_at": now_stamp(),
            "active": "1",
        }
        self.tasks.append(row)
        return row

    def add_info(
        self,
        parent_type: str,
        parent_id: str,
        title: str,
        body: str = "",
        section_path: str = "",
        emoji: str = INFO_EMOJI,
    ) -> dict:
        self._iid += 1
        row = {
            "id": next_id("INFO_", self._iid),
            "parent_type": parent_type,
            "parent_id": parent_id,
            "title": (title.strip() or "Note")[:200],
            "body": body,
            "emoji": emoji or INFO_EMOJI,
            "section_path": section_path,
            "sort_order": self._so(),
            "created_at": now_stamp(),
        }
        self.infos.append(row)
        return row

    def add_attachment(
        self, parent_type: str, parent_id: str, kind: str, ref: str, description: str
    ) -> dict:
        self._aid += 1
        row = {
            "id": next_id("ATT_", self._aid),
            "parent_type": parent_type,
            "parent_id": parent_id,
            "kind": kind,
            "ref": ref,
            "description": description,
            "sort_order": self._so(),
        }
        self.attachments.append(row)
        return row

    def copy_image(self, md_dir: Path, rel: str, parent_type: str, parent_id: str) -> None:
        rel = rel.strip().replace("/", "\\")
        src = (md_dir / rel).resolve() if not Path(rel).is_absolute() else Path(rel)
        if not src.exists():
            # try alongside md
            alt = md_dir / Path(rel).name
            if alt.exists():
                src = alt
            else:
                self.add_attachment(parent_type, parent_id, "file", str(src), f"missing image: {rel}")
                return
        dest_name = f"{datetime.now().strftime('%Y%m%d')}-{parent_id}-{src.name}"
        dest = self.attach_dir / dest_name
        try:
            shutil.copy2(src, dest)
            self.add_attachment(
                parent_type, parent_id, "image", f"attachments\\{dest_name}", src.name
            )
        except OSError:
            self.add_attachment(parent_type, parent_id, "file", str(src), rel)

    def attach_trailing(self, md_dir: Path, line: str, parent_type: str, parent_id: str) -> bool:
        """Handle image / url / path lines belonging to last entity. Returns True if consumed."""
        s = line.strip()
        if not s:
            return False
        m = IMG_RE.search(s)
        if m:
            self.copy_image(md_dir, m.group(2), parent_type, parent_id)
            return True
        if URL_RE.match(s):
            self.add_attachment(parent_type, parent_id, "url", s, s)
            return True
        if PATH_RE.match(s):
            self.add_attachment(parent_type, parent_id, "file", s, s)
            return True
        return False

    def parse_emoji_line(self, line: str) -> tuple[str, str] | None:
        s = line.strip()
        if not s or s.startswith("#") or s.startswith("<"):
            return None
        # skip pure markdown links without emoji
        m = EMOJI_RE.match(s)
        if not m:
            return None
        emoji, rest = m.group(1), m.group(2).strip()
        # normalize info emoji
        if emoji.startswith("ℹ"):
            emoji = INFO_EMOJI
        return emoji, rest

    def migrate_work_like(
        self,
        path: Path,
        filt: str,
        default_kind: str,
        seed_title: str,
    ) -> None:
        text = path.read_text(encoding="utf-8")
        lines = [strip_md(x) for x in text.splitlines()]
        md_dir = path.parent
        seed = self.add_project(seed_title, filt)
        current_project = seed
        section_parts: list[str] = []
        last_parent = ("project", seed["id"])
        prose_buf: list[str] = []

        def flush_prose():
            nonlocal prose_buf
            if not prose_buf:
                return
            body = "\n".join(prose_buf).strip()
            prose_buf = []
            if not body:
                return
            title = body.split("\n", 1)[0][:80]
            self.add_info(
                last_parent[0],
                last_parent[1],
                title,
                body,
                "/".join(section_parts),
            )

        top_sections = {"tasks", "projects", "ideas", "status", "activities", "backlog"}

        for raw in lines:
            if not raw.strip() or raw.strip().startswith("<link"):
                continue
            hm = HEADING_RE.match(raw.strip())
            if hm:
                flush_prose()
                level = len(hm.group(1))
                title = hm.group(2).strip()
                title_clean = re.sub(r"\*+", "", title).strip()
                low = title_clean.lower()
                if level == 1:
                    section_parts = []
                    if low in top_sections:
                        section_parts = [title_clean]
                    continue
                if level == 2:
                    # New project under current top section, or under seed
                    section_parts = section_parts[:1] + [title_clean] if section_parts else [title_clean]
                    current_project = self.add_project(
                        title_clean, filt, "/".join(section_parts[:-1])
                    )
                    last_parent = ("project", current_project["id"])
                    continue
                # deeper → section_path only
                depth = level - 1
                section_parts = section_parts[:depth]
                while len(section_parts) < depth:
                    section_parts.append(title_clean)
                if section_parts:
                    section_parts[-1] = title_clean
                else:
                    section_parts = [title_clean]
                continue

            parsed = self.parse_emoji_line(raw)
            if parsed:
                flush_prose()
                emoji, rest = parsed
                # strip trailing image markdown from rest
                imgs = IMG_RE.findall(rest)
                rest_clean = IMG_RE.sub("", rest).strip()
                sp = "/".join(section_parts)
                if emoji == INFO_EMOJI:
                    if last_parent[0] == "task":
                        ptype, pid = "task", last_parent[1]
                    else:
                        ptype, pid = "project", current_project["id"]
                    info = self.add_info(
                        ptype, pid, rest_clean or "Info", rest_clean, sp
                    )
                    for _alt, src in imgs:
                        self.copy_image(md_dir, src, "info", info["id"])
                    if URL_RE.match(rest_clean):
                        self.add_attachment("info", info["id"], "url", rest_clean, rest_clean)
                    # Keep last_parent as task/project so following lines attach correctly
                    last_parent = (ptype, pid)
                else:
                    kind = default_kind
                    task = self.add_task(
                        current_project["id"],
                        rest_clean or "(untitled)",
                        emoji,
                        filt,
                        kind,
                        section_path=sp,
                    )
                    last_parent = ("task", task["id"])
                    for _alt, src in imgs:
                        self.copy_image(md_dir, src, "task", task["id"])
                    if URL_RE.match(rest_clean):
                        # title is url — also attach
                        self.add_attachment("task", task["id"], "url", rest_clean, rest_clean)
                continue

            # continuation / trailing attachment for last entity
            parent_type, parent_id = last_parent
            if parent_type == "info":
                parent_type, parent_id = "project", current_project["id"]
            if self.attach_trailing(md_dir, raw, parent_type, parent_id):
                continue
            # prose continuation append to last info/task body if recent info
            if last_parent[0] == "task" and self.tasks and self.tasks[-1]["id"] == last_parent[1]:
                # append as info under task if looks like content
                if raw.strip() and not LINK_STYLE.match(raw.strip()):
                    prose_buf.append(raw.rstrip())
                continue
            if raw.strip():
                prose_buf.append(raw.rstrip())

        flush_prose()

    def migrate_habits(self, path: Path) -> None:
        text = path.read_text(encoding="utf-8")
        lines = [strip_md(x) for x in text.splitlines()]
        md_dir = path.parent
        seed = self.add_project("Habits & Health", "habits")
        current_project = seed
        section_parts: list[str] = []
        recurrence = ""
        last_parent = ("project", seed["id"])
        prose_buf: list[str] = []
        year_due = ""

        def flush_prose():
            nonlocal prose_buf
            if not prose_buf:
                return
            body = "\n".join(prose_buf).strip()
            prose_buf = []
            if not body:
                return
            title = body.split("\n", 1)[0][:80]
            self.add_info("project", current_project["id"], title, body, "/".join(section_parts))

        for raw in lines:
            if not raw.strip() or raw.strip().startswith("<link"):
                continue
            hm = HEADING_RE.match(raw.strip())
            if hm:
                flush_prose()
                level = len(hm.group(1))
                title = re.sub(r"\*+", "", hm.group(2)).strip()
                if level == 1:
                    section_parts = []
                    recurrence = ""
                    year_due = ""
                    continue
                if level == 2:
                    section_parts = [title]
                    recurrence = normalize_recurrence(title)
                    year_due = ""
                    # year headings like "2025 (Age 28)"
                    ym = re.match(r"^(\d{4})\b", title)
                    if ym and not recurrence:
                        year_due = f"{ym.group(1)}-01-01"
                        current_project = self.add_project(title, "habits", "timeline")
                        last_parent = ("project", current_project["id"])
                    elif title.lower() in {
                        "operating notes",
                        "recurring years lists (for quick reference)",
                        "age-specific screening add-ons",
                        "medication monitoring (fluvoxamina)",
                        "vaccination plan (follow-up, from 2025 onward)",
                        "notes from your records (prior vaccines)",
                        "year-by-year timeline (key dates with age)",
                    }:
                        current_project = self.add_project(title, "habits")
                        last_parent = ("project", current_project["id"])
                        recurrence = ""
                    else:
                        # cadence section stays under seed / current habits project
                        if current_project["id"] != seed["id"] and not year_due:
                            pass
                        current_project = seed
                        last_parent = ("project", seed["id"])
                    continue
                if level == 3:
                    section_parts = section_parts[:1] + [title]
                    ym = re.match(r"^(\d{4})\b", title)
                    if ym:
                        year_due = f"{ym.group(1)}-01-01"
                        current_project = self.add_project(title, "habits", "/".join(section_parts[:-1]))
                        last_parent = ("project", current_project["id"])
                    continue
                continue

            parsed = self.parse_emoji_line(raw)
            if parsed:
                flush_prose()
                emoji, rest = parsed
                imgs = IMG_RE.findall(rest)
                rest_clean = IMG_RE.sub("", rest).strip()
                sp = "/".join(section_parts)
                if emoji.startswith("ℹ"):
                    info = self.add_info(
                        "project", current_project["id"], rest_clean or "Info", rest_clean, sp
                    )
                    for _a, src in imgs:
                        self.copy_image(md_dir, src, "info", info["id"])
                    last_parent = ("project", current_project["id"])
                else:
                    kind = "habitual" if recurrence else ("punctual" if year_due else "habitual")
                    rec = recurrence if kind == "habitual" else ""
                    due = year_due if year_due else ""
                    next_due = due if kind == "habitual" and due else ""
                    task = self.add_task(
                        current_project["id"],
                        rest_clean or "(untitled)",
                        emoji,
                        "habits",
                        kind,
                        recurrence=rec,
                        section_path=sp,
                        due_date=due,
                        next_due=next_due,
                    )
                    last_parent = ("task", task["id"])
                    for _a, src in imgs:
                        self.copy_image(md_dir, src, "task", task["id"])
                continue

            parent_type, parent_id = last_parent
            if self.attach_trailing(md_dir, raw, parent_type, parent_id):
                continue
            if raw.strip():
                prose_buf.append(raw.rstrip())

        flush_prose()

    def save(self) -> None:
        write_csv(
            self.data_dir / "projects.csv",
            ["id", "title", "filter", "section_path", "sort_order", "active", "created_at"],
            self.projects,
        )
        write_csv(
            self.data_dir / "tasks.csv",
            [
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
            self.tasks,
        )
        write_csv(
            self.data_dir / "info_points.csv",
            [
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
            self.infos,
        )
        write_csv(
            self.data_dir / "attachments.csv",
            ["id", "parent_type", "parent_id", "kind", "ref", "description", "sort_order"],
            self.attachments,
        )


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--data-dir", required=True)
    ap.add_argument("--work", required=True)
    ap.add_argument("--punctual", required=True)
    ap.add_argument("--habits", required=True)
    args = ap.parse_args()

    data_dir = Path(args.data_dir)
    data_dir.mkdir(parents=True, exist_ok=True)

    m = Migrator(data_dir)
    m.migrate_work_like(Path(args.work), "work", "punctual", "Work inbox")
    m.migrate_work_like(Path(args.punctual), "personal", "punctual", "Personal inbox")
    m.migrate_habits(Path(args.habits))
    m.save()

    print(
        f"Migrated: {len(m.projects)} projects, {len(m.tasks)} tasks, "
        f"{len(m.infos)} info, {len(m.attachments)} attachments"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
