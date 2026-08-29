"""Deprecated static snapshot — prefer the live web app (task_server.py :8766)."""

from __future__ import annotations

import argparse
import html
from pathlib import Path


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Optional static Tasks HTML stub (prefer task_server.py web app)"
    )
    ap.add_argument("--data-dir", required=True)
    ap.add_argument("--output-dir", required=True)
    args = ap.parse_args()
    data_dir = Path(args.data_dir)
    out_dir = Path(args.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    note = f"""<!DOCTYPE html><html><head><meta charset="utf-8"/><title>Tasks</title></head>
<body style="font-family:Segoe UI;background:#121212;color:#eee;padding:40px">
<h1>Tasks moved to the web app</h1>
<p>Open <a style="color:#f1c40f" href="http://127.0.0.1:8766/">http://127.0.0.1:8766/</a>
(Utility Shortcuts <b>[T]</b> starts the server).</p>
<p>Data dir: {html.escape(str(data_dir))}</p>
</body></html>
"""
    (out_dir / "dashboard.html").write_text(note, encoding="utf-8")
    print("Wrote", out_dir / "dashboard.html", "(redirect stub — use task_server.py)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
