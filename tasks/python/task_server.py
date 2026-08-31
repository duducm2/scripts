"""Local HTTP server for the Tasks web app (port 8766)."""

from __future__ import annotations

import argparse
import base64
import json
import os
import subprocess
import sys
import traceback
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import unquote, urlparse

sys.path.insert(0, str(Path(__file__).resolve().parent))

from task_store import TaskStore  # noqa: E402

DEFAULT_PORT = 8766
WEB_DIR = Path(__file__).resolve().parent.parent / "web"


def open_user_target(raw: str) -> dict[str, Any]:
    """Open an http(s) URL or a local file/folder with the OS handler."""
    target = (raw or "").strip().strip('"').strip("'")
    if not target or "\n" in target or "\r" in target:
        return {"ok": False, "error": "invalid target"}
    lower = target.lower()
    if lower.startswith(("javascript:", "vbscript:", "data:")):
        return {"ok": False, "error": "unsupported target"}
    if lower.startswith(("http://", "https://")):
        parsed = urlparse(target)
        if parsed.scheme not in ("http", "https") or not parsed.netloc:
            return {"ok": False, "error": "invalid url"}
        _os_open(target)
        return {"ok": True, "opened": target}
    path = target
    if lower.startswith("file:"):
        parsed = urlparse(target)
        path = unquote(parsed.path or "")
        if parsed.netloc and parsed.netloc.lower() not in (
            "",
            "localhost",
            "127.0.0.1",
        ):
            path = "\\\\" + parsed.netloc + path.replace("/", "\\")
        elif os.name == "nt" and len(path) >= 3 and path[0] == "/" and path[2] == ":":
            path = path[1:]
        path = path.replace("/", "\\") if os.name == "nt" else path
    p = Path(path)
    try:
        resolved = p.expanduser()
        if not resolved.exists():
            return {"ok": False, "error": "file not found"}
        _os_open(str(resolved))
        return {"ok": True, "opened": str(resolved)}
    except OSError as e:
        return {"ok": False, "error": str(e)}


def _os_open(target: str) -> None:
    if os.name == "nt":
        os.startfile(target)  # type: ignore[attr-defined]
        return
    opener = "open" if sys.platform == "darwin" else "xdg-open"
    subprocess.Popen(
        [opener, target], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
    )


class TaskHandler(BaseHTTPRequestHandler):
    data_dir: Path
    scripts_root: Path

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
                        "open",
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
            payload = store.state()
            env_path = self.data_dir / "environment.txt"
            env = "work"
            try:
                raw = env_path.read_text(encoding="utf-8").strip().lower()
                if raw in {"work", "personal", "habits"}:
                    env = raw
            except OSError:
                pass
            payload["environment"] = env
            payload["default_focus"] = env
            self._json(200, payload)
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
            if path.startswith("/api/sections/"):
                self._json(200, store.delete_section(path.split("/")[-1]))
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
            self._json(
                500, {"ok": False, "error": str(e), "trace": traceback.format_exc()}
            )

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
            if path == "/api/sections":
                self._json(200, store.upsert_section(payload))
                return
            if path == "/api/tasks":
                self._json(200, store.upsert_task(payload))
                return
            if path.startswith("/api/tasks/") and path.endswith("/emoji"):
                tid = path[len("/api/tasks/") : -len("/emoji")]
                key = payload.get("key") or payload.get("emoji") or ""
                self._json(200, store.set_task_emoji(tid, key))
                return
            if path.startswith("/api/tasks/") and path.endswith("/to-personal"):
                tid = path[len("/api/tasks/") : -len("/to-personal")]
                self._json(200, store.spawn_personal_from_habit(tid))
                return
            if path == "/api/info":
                self._json(200, store.upsert_info(payload))
                return
            if path == "/api/open":
                self._json(200, open_user_target(str(payload.get("target") or "")))
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
            self._json(404, {"ok": False, "error": "not found"})
        except Exception as e:
            self._json(
                500, {"ok": False, "error": str(e), "trace": traceback.format_exc()}
            )


def make_handler(data_dir: Path, scripts_root: Path) -> type[TaskHandler]:
    class Handler(TaskHandler):
        pass

    Handler.data_dir = data_dir
    Handler.scripts_root = scripts_root
    return Handler


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="Tasks web HTTP server")
    p.add_argument("--data-dir", type=Path, required=True)
    p.add_argument("--scripts-root", type=Path, required=True)
    p.add_argument("--host", default="127.0.0.1")
    p.add_argument("--port", type=int, default=DEFAULT_PORT)
    args = p.parse_args(argv)

    data_dir = args.data_dir.resolve()
    data_dir.mkdir(parents=True, exist_ok=True)
    (data_dir / "attachments").mkdir(exist_ok=True)

    handler = make_handler(data_dir, args.scripts_root.resolve())
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
