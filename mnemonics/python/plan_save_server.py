"""Local HTTP server for saving plan checkbox progress from the dashboard."""

from __future__ import annotations

import argparse
import json
import sys
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent))

from study_plan_parser import default_studies_root  # noqa: E402
from study_plans_save import save_payload  # noqa: E402

DEFAULT_PORT = 8765


class PlanSaveHandler(BaseHTTPRequestHandler):
    data_dir: Path
    studies_root: Path
    output_dir: Path

    def log_message(self, format: str, *args: Any) -> None:
        sys.stderr.write("[plan_save_server] " + (format % args) + "\n")

    def _cors(self) -> None:
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def _json(self, code: int, payload: dict[str, Any]) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self._cors()
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self) -> None:
        self.send_response(204)
        self._cors()
        self.end_headers()

    def do_GET(self) -> None:
        if self.path.rstrip("/") == "/health":
            self._json(
                200,
                {
                    "ok": True,
                    "service": "plan_save_server",
                    "features": ["add_backlog", "remove_backlog"],
                },
            )
            return
        self._json(404, {"ok": False, "error": "not found"})

    def do_POST(self) -> None:
        if self.path.rstrip("/") != "/save":
            self._json(404, {"ok": False, "error": "not found"})
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            raw = self.rfile.read(length) if length else b"{}"
            payload = json.loads(raw.decode("utf-8"))
            result = save_payload(
                payload,
                self.data_dir,
                self.studies_root,
                self.output_dir,
            )
            code = 200 if result.get("ok") else 400
            self._json(code, result)
        except json.JSONDecodeError:
            self._json(400, {"ok": False, "error": "invalid JSON"})
        except OSError as e:
            self._json(500, {"ok": False, "error": str(e)})


def make_handler(
    data_dir: Path,
    studies_root: Path,
    output_dir: Path,
) -> type[PlanSaveHandler]:
    class Handler(PlanSaveHandler):
        pass

    Handler.data_dir = data_dir
    Handler.studies_root = studies_root
    Handler.output_dir = output_dir
    return Handler


def run_server(
    data_dir: Path,
    studies_root: Path,
    output_dir: Path,
    host: str = "127.0.0.1",
    port: int = DEFAULT_PORT,
) -> HTTPServer:
    handler = make_handler(data_dir, studies_root, output_dir)
    server = HTTPServer((host, port), handler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    print(f"Plan save server listening on http://{host}:{port}", file=sys.stderr)
    return server


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="Plan save HTTP server for dashboard")
    p.add_argument("--data-dir", type=Path, required=True)
    p.add_argument("--output-dir", type=Path, required=True)
    p.add_argument("--studies-root", type=Path, default=None)
    p.add_argument("--host", default="127.0.0.1")
    p.add_argument("--port", type=int, default=DEFAULT_PORT)
    args = p.parse_args(argv)

    studies_root = (
        args.studies_root.resolve()
        if args.studies_root
        else default_studies_root()
    )
    data_dir = args.data_dir.resolve()
    output_dir = args.output_dir.resolve()

    server = run_server(data_dir, studies_root, output_dir, args.host, args.port)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.shutdown()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
