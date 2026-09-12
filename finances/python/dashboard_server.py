"""Local finance cockpit server: static output/ + PATCH budgets.csv + notes."""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

sys.path.insert(0, str(Path(__file__).resolve().parent))
from data_aggregator import configure_paths, parse_decimal  # noqa: E402
import data_aggregator as _agg  # noqa: E402

PORT = 8765
BUDGET_HEADERS = ["year_month", "category_id", "planned_amount", "spent_amount"]
YEAR_MONTH_RE = re.compile(r"^\d{4}-\d{2}$")
NOTES_FILENAME = "general_notes.txt"
NOTES_MAX_BYTES = 200_000


def format_csv_decimal(num: float) -> str:
    return f"{float(num):.2f}".replace(".", ",")


def load_budgets(path: Path) -> list[dict]:
    if not path.exists():
        return []
    with path.open(encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def save_budgets(path: Path, rows: list[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(
            f,
            fieldnames=BUDGET_HEADERS,
            quoting=csv.QUOTE_MINIMAL,
            lineterminator="\n",
        )
        w.writeheader()
        for r in rows:
            w.writerow({h: r.get(h, "") for h in BUDGET_HEADERS})


def patch_budget(year_month: str, category_id: str, planned_raw) -> dict:
    if not YEAR_MONTH_RE.match(year_month or ""):
        return {"ok": False, "error": "year_month must be YYYY-MM", "status": 400}
    category_id = (category_id or "").strip()
    if not category_id:
        return {"ok": False, "error": "category_id required", "status": 400}

    planned = parse_decimal(planned_raw)
    if planned < 0:
        return {"ok": False, "error": "planned_amount must be >= 0", "status": 400}

    path = _agg.DATA / "budgets.csv"
    rows = load_budgets(path)
    found = False
    planned_fmt = format_csv_decimal(planned)
    for row in rows:
        if (
            row.get("year_month") == year_month
            and row.get("category_id") == category_id
        ):
            row["planned_amount"] = planned_fmt
            found = True
            break
    if not found:
        return {
            "ok": False,
            "error": f"No budget for {year_month} / {category_id}",
            "status": 404,
        }

    save_budgets(path, rows)
    return {"ok": True, "planned_amount": planned_fmt, "status": 200}


def load_cards() -> list[dict]:
    path = _agg.DATA / "credit_cards.csv"
    if not path.exists():
        return []
    with path.open(encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))
    out = []
    for c in rows:
        cid = (c.get("id") or "").strip()
        if not cid:
            continue
        out.append(
            {
                "id": cid,
                "name": c.get("name") or cid,
                "limit": c.get("limit") or "",
                "current_spent": c.get("current_spent") or "",
                "closing_day": str(c.get("closing_day") or "1").strip() or "1",
                "due_day": str(c.get("due_day") or "").strip(),
            }
        )
    return out


def notes_path() -> Path:
    return _agg.DATA / NOTES_FILENAME


def load_notes() -> str:
    path = notes_path()
    if not path.exists():
        return ""
    try:
        return path.read_text(encoding="utf-8")
    except OSError:
        return ""


def save_notes(text) -> dict:
    if text is None:
        text = ""
    if not isinstance(text, str):
        text = str(text)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    if len(text.encode("utf-8")) > NOTES_MAX_BYTES:
        return {"ok": False, "error": "Notes too large", "status": 400}
    path = notes_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\n")
    return {"ok": True, "status": 200}


class DashboardHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, directory: str | None = None, **kwargs):
        super().__init__(*args, directory=directory, **kwargs)

    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/api/health":
            self._json(200, {"ok": True})
            return
        if parsed.path == "/api/cards":
            self._json(200, {"ok": True, "cards": load_cards()})
            return
        if parsed.path == "/api/notes":
            self._json(200, {"ok": True, "text": load_notes()})
            return
        super().do_GET()

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, PUT, PATCH, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_PUT(self):
        parsed = urlparse(self.path)
        if parsed.path != "/api/notes":
            self._json(404, {"ok": False, "error": "Not found"})
            return
        length = int(self.headers.get("Content-Length") or 0)
        raw = self.rfile.read(length) if length else b"{}"
        try:
            body = json.loads(raw.decode("utf-8") or "{}")
        except json.JSONDecodeError:
            self._json(400, {"ok": False, "error": "Invalid JSON"})
            return
        result = save_notes(body.get("text", ""))
        status = int(result.pop("status", 200))
        self._json(status, result)

    def do_PATCH(self):
        parsed = urlparse(self.path)
        if parsed.path != "/api/budgets":
            self._json(404, {"ok": False, "error": "Not found"})
            return
        length = int(self.headers.get("Content-Length") or 0)
        raw = self.rfile.read(length) if length else b"{}"
        try:
            body = json.loads(raw.decode("utf-8") or "{}")
        except json.JSONDecodeError:
            self._json(400, {"ok": False, "error": "Invalid JSON"})
            return
        result = patch_budget(
            str(body.get("year_month", "")),
            str(body.get("category_id", "")),
            body.get("planned_amount"),
        )
        status = int(result.pop("status", 200))
        self._json(status, result)

    def _json(self, status: int, payload: dict):
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, fmt: str, *args):
        sys.stderr.write("%s - %s\n" % (self.address_string(), fmt % args))


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Serve finance cockpit + budget API")
    parser.add_argument("--data-dir", default="", help="Absolute path to finances/data")
    parser.add_argument(
        "--output-dir", default="", help="Absolute path to finances/output"
    )
    parser.add_argument("--port", type=int, default=PORT)
    args = parser.parse_args(argv)
    if args.data_dir or args.output_dir:
        configure_paths(
            data_dir=args.data_dir or None,
            output_dir=args.output_dir or None,
        )
    out_dir = _agg.OUTPUT
    out_dir.mkdir(parents=True, exist_ok=True)

    def factory(*a, **kw):
        return DashboardHandler(*a, directory=str(out_dir), **kw)

    try:
        httpd = ThreadingHTTPServer(("127.0.0.1", args.port), factory)
    except OSError:
        print(f"already running on 127.0.0.1:{args.port}", flush=True)
        return 0

    print(
        f"serving {out_dir} on http://127.0.0.1:{args.port}/ " f"(data={_agg.DATA})",
        flush=True,
    )
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        httpd.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
