"""Local HTTP server for the Tasks web app (port 8766)."""

from __future__ import annotations

import argparse
import base64
import json
import subprocess
import sys
import traceback
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import unquote, urlparse

sys.path.insert(0, str(Path(__file__).resolve().parent))

from task_pack_import import commit_pack, preview_pack  # noqa: E402
from task_store import TaskStore  # noqa: E402

DEFAULT_PORT = 8766
WEB_DIR = Path(__file__).resolve().parent.parent / "web"


class TaskHandler(BaseHTTPRequestHandler):
    data_dir: Path
    scripts_root: Path
    notes_candidates: list[Path]

    def log_message(self, format: str, *args: Any) -> None:
        sys.stderr.write("[task_server] " + (format % args) + "\n")

    def _cors(self) -> None:
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, DELETE, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def _json(self, code: int, payload: dict[str, Any]) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self._cors()
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _bytes(self, code: int, data: bytes, content_type: str) -> None:
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        self._cors()
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _read_json(self) -> dict[str, Any]:
        length = int(self.headers.get("Content-Length", "0") or 0)
        raw = self.rfile.read(length) if length else b"{}"
        if not raw:
            return {}
        return json.loads(raw.decode("utf-8"))

    def do_OPTIONS(self) -> None:
        self.send_response(204)
        self._cors()
        self.end_headers()

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        path = unquote(parsed.path)

        if path.rstrip("/") == "/health":
            self._json(
                200,
                {
                    "ok": True,
                    "service": "task_server",
                    "port": DEFAULT_PORT,
                    "features": [
                        "state",
                        "crud",
                        "attachments",
                        "import",
                        "migrate",
                        "git_push",
                    ],
                },
            )
            return

        if path in ("/", "/index.html"):
            index = WEB_DIR / "index.html"
            if not index.is_file():
                self._json(500, {"ok": False, "error": "index.html missing"})
                return
            html = index.read_text(encoding="utf-8")
            # document title for Chrome HWND matching
            if "<title>Tasks</title>" not in html:
                html = html.replace("<title>", "<title>Tasks — ", 1)
            self._bytes(200, html.encode("utf-8"), "text/html; charset=utf-8")
            return

        if path.startswith("/data/attachments/"):
            rel = path[len("/data/") :]
            file_path = (self.data_dir / rel.replace("/", "\\")).resolve()
            attach_root = (self.data_dir / "attachments").resolve()
            try:
                file_path.relative_to(attach_root)
            except ValueError:
                self._json(403, {"ok": False, "error": "forbidden"})
                return
            if not file_path.is_file():
                self._json(404, {"ok": False, "error": "not found"})
                return
            data = file_path.read_bytes()
            ctype = "image/png"
            if file_path.suffix.lower() in {".jpg", ".jpeg"}:
                ctype = "image/jpeg"
            elif file_path.suffix.lower() == ".webp":
                ctype = "image/webp"
            elif file_path.suffix.lower() == ".gif":
                ctype = "image/gif"
            self._bytes(200, data, ctype)
            return

        if path == "/api/state":
            store = TaskStore(self.data_dir)
            self._json(200, store.state())
            return

        self._json(404, {"ok": False, "error": "not found"})

    def do_DELETE(self) -> None:
        parsed = urlparse(self.path)
        path = unquote(parsed.path)
        store = TaskStore(self.data_dir)
        try:
            if path.startswith("/api/projects/"):
                self._json(200, store.delete_project(path.split("/")[-1]))
                return
            if path.startswith("/api/tasks/"):
                self._json(200, store.delete_task(path.split("/")[-1]))
                return
            if path.startswith("/api/info/"):
                self._json(200, store.delete_info(path.split("/")[-1]))
                return
            if path.startswith("/api/attachments/"):
                self._json(200, store.delete_attachment(path.split("/")[-1]))
                return
            self._json(404, {"ok": False, "error": "not found"})
        except Exception as e:
            self._json(500, {"ok": False, "error": str(e), "trace": traceback.format_exc()})

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        path = unquote(parsed.path)
        store = TaskStore(self.data_dir)
        try:
            payload = self._read_json()
        except json.JSONDecodeError:
            self._json(400, {"ok": False, "error": "invalid JSON"})
            return
        try:
            if path == "/api/projects":
                self._json(200, store.upsert_project(payload))
                return
            if path == "/api/tasks":
                self._json(200, store.upsert_task(payload))
                return
            if path.startswith("/api/tasks/") and path.endswith("/emoji"):
                tid = path[len("/api/tasks/") : -len("/emoji")]
                key = payload.get("key") or payload.get("emoji") or ""
                self._json(200, store.set_task_emoji(tid, key))
                return
            if path == "/api/info":
                self._json(200, store.upsert_info(payload))
                return
            if path == "/api/attachments":
                parent_type = payload.get("parent_type") or ""
                parent_id = payload.get("parent_id") or ""
                kind = payload.get("kind") or "image"
                if kind == "image" and payload.get("data_base64"):
                    raw = base64.b64decode(payload["data_base64"])
                    self._json(
                        200,
                        store.save_image_bytes(
                            parent_type,
                            parent_id,
                            raw,
                            payload.get("filename") or "image.png",
                            payload.get("description") or "",
                        ),
                    )
                    return
                self._json(
                    200,
                    store.add_attachment(
                        parent_type,
                        parent_id,
                        kind,
                        payload.get("ref") or "",
                        payload.get("description") or "",
                    ),
                )
                return
            if path == "/api/import/preview":
                self._json(200, preview_pack(store))
                return
            if path == "/api/import/commit":
                # client may send preview payload; else re-read desktop
                pack = payload if payload.get("projects") is not None or payload.get("tasks") is not None else None
                if pack and pack.get("ok") is False:
                    self._json(400, pack)
                    return
                self._json(200, commit_pack(store, pack if pack and pack.get("ok") else None))
                return
            if path == "/api/migrate":
                self._json(200, self._migrate())
                return
            if path == "/api/git-push":
                self._json(200, self._git_push())
                return
            self._json(404, {"ok": False, "error": "not found"})
        except Exception as e:
            self._json(500, {"ok": False, "error": str(e), "trace": traceback.format_exc()})

    def _migrate(self) -> dict[str, Any]:
        py = Path(__file__).resolve().parent / "migrate_from_md.py"
        work = punctual = habits = None
        for root in self.notes_candidates:
            w = root / "work" / "work.md"
            p = root / "main" / "punctual.md"
            h = root / "main" / "habits.md"
            if w.is_file() and p.is_file() and h.is_file():
                work, punctual, habits = w, p, h
                break
        if not work:
            return {"ok": False, "error": "Could not find work.md / punctual.md / habits.md"}
        cmd = [
            sys.executable,
            str(py),
            "--data-dir",
            str(self.data_dir),
            "--work",
            str(work),
            "--punctual",
            str(punctual),
            "--habits",
            str(habits),
        ]
        r = subprocess.run(cmd, capture_output=True, text=True, cwd=str(py.parent))
        if r.returncode != 0:
            return {"ok": False, "error": (r.stderr or r.stdout or "migrate failed")[:500]}
        return {"ok": True, "message": (r.stdout or "Migrated").strip()}

    def _git_push(self) -> dict[str, Any]:
        root = self.scripts_root
        if not (root / ".git").exists():
            return {"ok": False, "error": "scripts repo not found"}
        msgs = []
        for args in (
            ["git", "add", "-A"],
            ["git", "diff", "--cached", "--quiet"],
        ):
            pass
        status = subprocess.run(
            ["git", "status", "--porcelain"],
            cwd=str(root),
            capture_output=True,
            text=True,
        )
        if status.returncode != 0:
            return {"ok": False, "error": status.stderr or "git status failed"}
        if not status.stdout.strip():
            # still try push
            push = subprocess.run(
                ["git", "push"],
                cwd=str(root),
                capture_output=True,
                text=True,
            )
            if push.returncode != 0:
                return {"ok": False, "error": (push.stderr or push.stdout or "push failed")[:400]}
            return {"ok": True, "message": "Nothing to commit; push ok"}
        commit = subprocess.run(
            ["git", "add", "-A"],
            cwd=str(root),
            capture_output=True,
            text=True,
        )
        if commit.returncode != 0:
            return {"ok": False, "error": commit.stderr or "git add failed"}
        commit = subprocess.run(
            ["git", "commit", "-m", "Tasks web sync"],
            cwd=str(root),
            capture_output=True,
            text=True,
        )
        # commit may fail if nothing staged after add — ok
        push = subprocess.run(
            ["git", "push"],
            cwd=str(root),
            capture_output=True,
            text=True,
        )
        if push.returncode != 0:
            return {"ok": False, "error": (push.stderr or push.stdout or "push failed")[:400]}
        msgs.append("pushed")
        return {"ok": True, "message": "Committed and pushed scripts repo"}


def make_handler(
    data_dir: Path, scripts_root: Path, notes_candidates: list[Path]
) -> type[TaskHandler]:
    class Handler(TaskHandler):
        pass

    Handler.data_dir = data_dir
    Handler.scripts_root = scripts_root
    Handler.notes_candidates = notes_candidates
    return Handler


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="Tasks web HTTP server")
    p.add_argument("--data-dir", type=Path, required=True)
    p.add_argument("--scripts-root", type=Path, required=True)
    p.add_argument("--host", default="127.0.0.1")
    p.add_argument("--port", type=int, default=DEFAULT_PORT)
    p.add_argument("--notes-root", type=Path, default=None)
    args = p.parse_args(argv)

    data_dir = args.data_dir.resolve()
    data_dir.mkdir(parents=True, exist_ok=True)
    (data_dir / "attachments").mkdir(exist_ok=True)

    notes_candidates: list[Path] = []
    if args.notes_root:
        notes_candidates.append(args.notes_root.resolve())
    # common sibling notes repo
    sibling = args.scripts_root.resolve().parent / "notes"
    if sibling.is_dir():
        notes_candidates.append(sibling)

    handler = make_handler(data_dir, args.scripts_root.resolve(), notes_candidates)
    server = ThreadingHTTPServer((args.host, args.port), handler)
    print(f"Tasks server listening on http://{args.host}:{args.port}", file=sys.stderr)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.shutdown()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
