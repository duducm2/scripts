"""Local HTTP server for the Memory Palace web app (port 8767)."""

from __future__ import annotations

import argparse
import atexit
import json
import os
import shutil
import subprocess
import sys
import traceback
import urllib.error
import urllib.parse
import urllib.request
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import unquote, urlparse

sys.path.insert(0, str(Path(__file__).resolve().parent))

from palace_save import save_images, save_notes  # noqa: E402
from palace_store import PalaceStore  # noqa: E402
from study_plan_parser import default_studies_root  # noqa: E402
from study_plans_save import save_payload  # noqa: E402

DEFAULT_PORT = 8767
WEB_DIR = Path(__file__).resolve().parent.parent / "web"
MNEMONICS_ROOT = Path(__file__).resolve().parent.parent

STUDY_LINKS_API = (
    "https://script.google.com/macros/s/"
    "AKfycbzKDLbmzGF8iduyNpaUymONEkERi089rBjW0jrYUX4a8K9ornfGwYIOsgQP1K_dfaj5/exec"
)
LINK_KEYS = {
    "video": "subtopic",
    "article": "subtopic_article",
    "favorite": "subtopic_favorite",
}

GLOSSARY = [
    {"term": "Study", "def": "Domain owning Memory Palaces (CSV study_id)."},
    {
        "term": "Memory Palace",
        "def": "Location / former street; one character; up to 5 beasts.",
    },
    {"term": "Character", "def": "Exactly one canon character per palace."},
    {"term": "Beast", "def": "Bestiary peg holder for knowledge atoms."},
    {"term": "Knowledge Atom", "def": "Concept + Quote + Story + Sensory on a beast."},
    {"term": "Peg", "def": "Letter code from bestiary (never numeric)."},
    {"term": "Plan", "def": "Study checklist (plans / plan_items / plan_resources)."},
]


def _os_open(target: str) -> None:
    if os.name == "nt":
        os.startfile(target)  # type: ignore[attr-defined]
        return
    opener = "open" if sys.platform == "darwin" else "xdg-open"
    subprocess.Popen(
        [opener, target], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
    )


def _chrome_exe() -> str | None:
    """Resolve Google Chrome like AHK ``chrome.exe`` / common install paths."""
    which = shutil.which("chrome") or shutil.which("chrome.exe")
    if which:
        return which
    if os.name != "nt":
        return shutil.which("google-chrome") or shutil.which("chromium")
    candidates = [
        Path(os.environ.get("PROGRAMFILES", r"C:\Program Files"))
        / "Google"
        / "Chrome"
        / "Application"
        / "chrome.exe",
        Path(os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)"))
        / "Google"
        / "Chrome"
        / "Application"
        / "chrome.exe",
        Path(os.environ.get("LOCALAPPDATA", ""))
        / "Google"
        / "Chrome"
        / "Application"
        / "chrome.exe",
    ]
    for p in candidates:
        if p.is_file():
            return str(p)
    return None


def open_url_in_chrome(raw: str, *, new_window: bool = True) -> dict[str, Any]:
    """Open http(s) in Chrome (avoids Windows handing YouTube URLs to the app)."""
    target = (raw or "").strip().strip('"').strip("'")
    if not target or "\n" in target or "\r" in target:
        return {"ok": False, "error": "invalid target"}
    lower = target.lower()
    if lower.startswith(("javascript:", "vbscript:", "data:")):
        return {"ok": False, "error": "unsupported target"}
    if not lower.startswith(("http://", "https://")):
        return {"ok": False, "error": "expected http(s) url"}
    parsed = urlparse(target)
    if parsed.scheme not in ("http", "https") or not parsed.netloc:
        return {"ok": False, "error": "invalid url"}
    chrome = _chrome_exe()
    if not chrome:
        return {"ok": False, "error": "chrome not found"}
    cmd = [chrome]
    if new_window:
        cmd.append("--new-window")
    cmd.append(target)
    try:
        subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return {"ok": True, "opened": target, "via": "chrome"}
    except OSError as e:
        return {"ok": False, "error": str(e)}


def open_user_target(raw: str) -> dict[str, Any]:
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
        # Prefer Chrome for http(s) so protocol handlers (e.g. YouTube app) do not steal.
        chrome_result = open_url_in_chrome(target, new_window=True)
        if chrome_result.get("ok"):
            return chrome_result
        _os_open(target)
        return {"ok": True, "opened": target, "via": "os"}
    p = Path(target)
    try:
        resolved = p.expanduser()
        if not resolved.exists():
            return {"ok": False, "error": "file not found"}
        _os_open(str(resolved))
        return {"ok": True, "opened": str(resolved)}
    except OSError as e:
        return {"ok": False, "error": str(e)}


def study_link_get(key: str) -> dict[str, Any]:
    url = f"{STUDY_LINKS_API}?key={urllib.parse.quote(key)}"
    try:
        with urllib.request.urlopen(url, timeout=20) as resp:
            text = resp.read().decode("utf-8", errors="replace")
        return {"ok": True, "key": key, "raw": text, "url": _parse_link_url(text)}
    except Exception as e:
        return {"ok": False, "error": str(e), "key": key}


def study_link_set(key: str, link_url: str) -> dict[str, Any]:
    body = urllib.parse.urlencode({"key": key, "url": link_url}).encode("utf-8")
    req = urllib.request.Request(
        STUDY_LINKS_API,
        data=body,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=20) as resp:
            text = resp.read().decode("utf-8", errors="replace")
        return {"ok": True, "key": key, "raw": text}
    except Exception as e:
        return {"ok": False, "error": str(e), "key": key}


def _parse_link_url(raw: str) -> str:
    """Extract URL from Study Link API body (AHK-compatible form encoding).

    Apps Script returns a single line like ``key=subtopic&url=https://…``.
    ``url`` is always the last field and may itself contain ``&`` (e.g. ``&t=``),
    so we take everything after the first ``url=`` and URL-decode it.
    """
    text = (raw or "").replace("\r", "").strip()
    if not text:
        return ""
    lower = text.lower()
    if lower.startswith("http://") or lower.startswith("https://"):
        return text
    # Prefer last ``url=`` in case of duplicates; value may include ``&``.
    idx = lower.rfind("url=")
    if idx >= 0:
        return urllib.parse.unquote(text[idx + 4 :].strip())
    for part in text.split("\n"):
        part = part.strip()
        if part.lower().startswith("url="):
            return urllib.parse.unquote(part[4:].strip())
    return ""


class PalaceHandler(BaseHTTPRequestHandler):
    data_dir: Path
    output_dir: Path
    studies_root: Path
    scripts_root: Path
    # Drop stalled clients so worker threads cannot pile up forever.
    timeout = 60

    def log_message(self, format: str, *args: Any) -> None:
        sys.stderr.write("[palace_server] " + (format % args) + "\n")

    def handle_one_request(self) -> None:
        try:
            super().handle_one_request()
        except (ConnectionResetError, BrokenPipeError, TimeoutError, OSError) as e:
            sys.stderr.write(f"[palace_server] client connection dropped: {e}\n")
        except Exception:
            sys.stderr.write("[palace_server] request handler crashed:\n")
            traceback.print_exc(file=sys.stderr)

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

    def _store(self) -> PalaceStore:
        return PalaceStore(self.data_dir, self.output_dir, self.studies_root)

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
                    "service": "palace_server",
                    "port": DEFAULT_PORT,
                    "features": [
                        "state",
                        "crud",
                        "plans",
                        "notes",
                        "images",
                        "regen",
                        "quick-image",
                        "links",
                        "method",
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
            if "<title>Memory Palace</title>" not in html:
                html = html.replace("<title>", "<title>Memory Palace — ", 1)
            self._bytes(200, html.encode("utf-8"), "text/html; charset=utf-8")
            return

        # Static assets next to the SPA (dashboard.css, etc.)
        if path.startswith("/") and ".." not in path:
            rel = path.lstrip("/").replace("\\", "/")
            if rel and "/" not in rel and not rel.startswith("api"):
                asset = (WEB_DIR / rel).resolve()
                try:
                    asset.relative_to(WEB_DIR.resolve())
                except ValueError:
                    pass
                else:
                    if asset.is_file():
                        ctype = "application/octet-stream"
                        suf = asset.suffix.lower()
                        if suf == ".css":
                            ctype = "text/css; charset=utf-8"
                        elif suf == ".js":
                            ctype = "application/javascript; charset=utf-8"
                        elif suf == ".html":
                            ctype = "text/html; charset=utf-8"
                        elif suf == ".svg":
                            ctype = "image/svg+xml"
                        self._bytes(200, asset.read_bytes(), ctype)
                        return

        if path == "/api/state":
            self._json(200, self._store().state())
            return

        if path == "/api/plans/view":
            from plan_csv import load_plan_tables, plan_row_to_payload  # noqa: E402
            from study_plan_parser import enrich_plan, slug_filename  # noqa: E402

            qs = urllib.parse.parse_qs(parsed.query)
            study_id = (qs.get("study_id") or [""])[0].strip()
            if not study_id:
                self._json(400, {"ok": False, "error": "study_id required"})
                return
            store = self._store()
            data = store._load_tree()
            study = next(
                (s for s in data.get("studies", []) if s.get("id") == study_id),
                None,
            )
            if not study:
                self._json(404, {"ok": False, "error": "study not found"})
                return
            plan = next(
                (
                    p
                    for p in data.get("plans", [])
                    if p.get("study_id") == study_id and p.get("active", "1") != "0"
                ),
                None,
            )
            if not plan:
                plan = next(
                    (p for p in data.get("plans", []) if p.get("study_id") == study_id),
                    None,
                )
            if not plan:
                self._json(
                    200,
                    {
                        "ok": True,
                        "study_id": study_id,
                        "plan": None,
                        "github_url": store.state()["meta"]["plans_github"],
                    },
                )
                return
            slug = slug_filename(study.get("notes_rel_path") or study_id)
            tables = load_plan_tables(self.data_dir)
            payload = enrich_plan(
                plan_row_to_payload(
                    plan,
                    tables["plan_items"],
                    tables["plan_resources"],
                    slug,
                )
            )
            self._json(
                200,
                {
                    "ok": True,
                    "study_id": study_id,
                    "plan": payload,
                    "github_url": store.state()["meta"]["plans_github"],
                },
            )
            return

        if path == "/api/glossary":
            self._json(200, {"ok": True, "items": GLOSSARY})
            return

        if path == "/api/method":
            from technique_renderer import (  # noqa: E402
                build_method_panel,
                default_technique_dir,
            )

            technique_dir = default_technique_dir(Path(__file__).resolve().parent)
            if not technique_dir.is_dir():
                technique_dir = MNEMONICS_ROOT / "technique"
            html_body, canon = build_method_panel(technique_dir)
            self._json(
                200,
                {
                    "ok": True,
                    "html": html_body,
                    "canon": canon,
                    "technique_dir": str(technique_dir),
                },
            )
            return

        if path.startswith("/media/"):
            rel = path[len("/media/") :]
            # serve practice images from output/
            file_path = (self.output_dir / rel.replace("\\", "/")).resolve()
            root = self.output_dir.resolve()
            try:
                file_path.relative_to(root)
            except ValueError:
                self._json(403, {"ok": False, "error": "forbidden"})
                return
            if not file_path.is_file():
                self._json(404, {"ok": False, "error": "not found"})
                return
            data = file_path.read_bytes()
            ctype = "application/octet-stream"
            suf = file_path.suffix.lower()
            if suf in {".png"}:
                ctype = "image/png"
            elif suf in {".jpg", ".jpeg"}:
                ctype = "image/jpeg"
            elif suf == ".webp":
                ctype = "image/webp"
            elif suf == ".gif":
                ctype = "image/gif"
            elif suf == ".md":
                ctype = "text/markdown; charset=utf-8"
            self._bytes(200, data, ctype)
            return

        if path.startswith("/api/links/"):
            kind = path.split("/")[-1]
            key = LINK_KEYS.get(kind)
            if not key:
                self._json(404, {"ok": False, "error": "unknown link kind"})
                return
            self._json(200, study_link_get(key))
            return

        self._json(404, {"ok": False, "error": "not found"})

    def do_DELETE(self) -> None:
        parsed = urlparse(self.path)
        path = unquote(parsed.path)
        store = self._store()
        try:
            parts = [p for p in path.split("/") if p]
            # /api/{entity}/{id}
            if len(parts) == 3 and parts[0] == "api":
                entity, eid = parts[1], parts[2]
                self._json(200, store.delete(entity, eid))
                return
            self._json(404, {"ok": False, "error": "not found"})
        except Exception as e:
            self._json(
                500,
                {"ok": False, "error": str(e), "trace": traceback.format_exc()[-800:]},
            )

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        path = unquote(parsed.path).rstrip("/")
        try:
            payload = self._read_json()
        except json.JSONDecodeError:
            self._json(400, {"ok": False, "error": "invalid JSON"})
            return

        store = self._store()
        try:
            if path == "/api/open":
                self._json(200, open_user_target(str(payload.get("target") or "")))
                return

            if path == "/api/regen":
                self._json(200, store.regen_all())
                return

            if path == "/api/quick-image":
                desk = payload.get("desktop")
                desktop = Path(desk) if desk else None
                self._json(
                    200,
                    store.quick_image(str(payload.get("palace_id") or ""), desktop),
                )
                return

            if path in ("/api/plans/save", "/save"):
                result = save_payload(
                    payload, self.data_dir, self.studies_root, self.output_dir
                )
                self._json(200 if result.get("ok") else 400, result)
                return

            if path in ("/api/palace/notes", "/palace/notes"):
                result = save_notes(
                    str(payload.get("palace_id") or ""),
                    str(
                        payload.get("notes") if payload.get("notes") is not None else ""
                    ),
                    self.data_dir,
                    self.output_dir,
                    self.studies_root,
                )
                self._json(200 if result.get("ok") else 400, result)
                return

            if path in ("/api/palace/images", "/palace/images"):
                result = save_images(
                    payload, self.data_dir, self.output_dir, self.studies_root
                )
                self._json(200 if result.get("ok") else 400, result)
                return

            if path.startswith("/api/links/"):
                kind = path.split("/")[-1]
                key = LINK_KEYS.get(kind)
                if not key:
                    self._json(404, {"ok": False, "error": "unknown link kind"})
                    return
                action = str(payload.get("action") or "set").lower()
                if action == "open":
                    # Prefer URL already fetched by the SPA (avoids a second Apps Script round-trip).
                    url = str(payload.get("url") or "").strip()
                    got: dict[str, Any] = {"ok": True, "key": key}
                    if not url:
                        got = study_link_get(key)
                        url = str(got.get("url") or "")
                    if not url:
                        self._json(400, {"ok": False, "error": "no url stored", **got})
                        return
                    opened = open_url_in_chrome(url, new_window=True)
                    self._json(
                        200 if opened.get("ok") else 400,
                        {**got, "url": url, **opened},
                    )
                    return
                link_url = str(payload.get("url") or "").strip()
                if not link_url:
                    self._json(400, {"ok": False, "error": "url required"})
                    return
                self._json(200, study_link_set(key, link_url))
                return

            # /api/{entity} upsert
            parts = [p for p in path.split("/") if p]
            if len(parts) == 2 and parts[0] == "api":
                entity = parts[1]
                result = store.upsert(entity, payload)
                self._json(200 if result.get("ok") else 400, result)
                return

            self._json(404, {"ok": False, "error": "not found"})
        except Exception as e:
            self._json(
                500,
                {
                    "ok": False,
                    "error": str(e),
                    "trace": traceback.format_exc()[-800:],
                },
            )


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="Memory Palace web server")
    p.add_argument("--data-dir", type=Path, required=True)
    p.add_argument("--output-dir", type=Path, default=None)
    p.add_argument("--studies-root", type=Path, default=None)
    p.add_argument("--scripts-root", type=Path, default=None)
    p.add_argument("--host", default="127.0.0.1")
    p.add_argument("--port", type=int, default=DEFAULT_PORT)
    args = p.parse_args(argv)

    data_dir = args.data_dir.resolve()
    output_dir = (
        args.output_dir.resolve()
        if args.output_dir
        else (data_dir.parent / "output").resolve()
    )
    studies_root = (
        args.studies_root.resolve() if args.studies_root else default_studies_root()
    )
    scripts_root = (
        args.scripts_root.resolve()
        if args.scripts_root
        else Path(__file__).resolve().parents[2]
    )

    class Handler(PalaceHandler):
        pass

    Handler.data_dir = data_dir
    Handler.output_dir = output_dir
    Handler.studies_root = studies_root
    Handler.scripts_root = scripts_root

    class PalaceHTTPServer(ThreadingHTTPServer):
        allow_reuse_address = True
        daemon_threads = True

        def handle_error(self, request: Any, client_address: Any) -> None:
            sys.stderr.write(f"[palace_server] handle_error {client_address}:\n")
            traceback.print_exc(file=sys.stderr)

    server = PalaceHTTPServer((args.host, args.port), Handler)
    pid_path = data_dir / "palace_server.pid"
    try:
        pid_path.write_text(str(os.getpid()), encoding="utf-8")
    except OSError as e:
        sys.stderr.write(f"[palace_server] could not write pid file: {e}\n")

    def _clear_pid() -> None:
        try:
            if pid_path.is_file() and pid_path.read_text(
                encoding="utf-8"
            ).strip() == str(os.getpid()):
                pid_path.unlink(missing_ok=True)
        except OSError:
            pass

    atexit.register(_clear_pid)

    sys.stderr.write(
        f"[palace_server] listening on http://{args.host}:{args.port} "
        f"pid={os.getpid()} data={data_dir}\n"
    )
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        try:
            server.server_close()
        except OSError:
            pass
        _clear_pid()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
