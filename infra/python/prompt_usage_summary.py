#!/usr/bin/env python3
"""Summarize prompt_usage.log: most used, never used (vs prompts.ini), last used."""

from __future__ import annotations

import argparse
import configparser
import sys
from collections import Counter
from datetime import datetime, timedelta
from pathlib import Path


def repo_root() -> Path:
    return Path(__file__).resolve().parent.parent.parent


def read_ini(path: Path) -> configparser.ConfigParser:
    cp = configparser.ConfigParser()
    raw = path.read_bytes()
    if raw.startswith(b"\xff\xfe") or raw.startswith(b"\xfe\xff"):
        text = raw.decode("utf-16")
    else:
        text = raw.decode("utf-8-sig")
    cp.read_string(text)
    return cp


def load_prompt_chars(ini_path: Path) -> dict[str, str]:
    if not ini_path.is_file():
        return {}
    cp = read_ini(ini_path)
    out: dict[str, str] = {}
    for section in cp.sections():
        if not section.startswith("Prompt_"):
            continue
        char = cp.get(section, "Char", fallback="").strip().lower()
        name = cp.get(section, "Name", fallback="").strip()
        if char:
            out[char] = name
    return out


def parse_log(log_path: Path, since: datetime | None) -> list[dict]:
    if not log_path.is_file():
        return []
    rows: list[dict] = []
    for line in log_path.read_text(encoding="utf-8", errors="replace").splitlines():
        line = line.strip()
        if not line:
            continue
        parts = line.split("|")
        if len(parts) < 2:
            continue
        try:
            ts = datetime.strptime(parts[0], "%Y-%m-%d %H:%M:%S")
        except ValueError:
            continue
        if since and ts < since:
            continue
        char = parts[1].strip().lower()
        name = parts[2].strip() if len(parts) > 2 else ""
        rows.append({"ts": ts, "char": char, "name": name})
    return rows


def safe_print(text: str) -> None:
    enc = getattr(sys.stdout, "encoding", None) or "utf-8"
    try:
        print(text)
    except UnicodeEncodeError:
        print(text.encode(enc, errors="replace").decode(enc))


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Summarize prompt usage log")
    parser.add_argument(
        "--days", type=int, default=7, help="Window in days (0 = all time)"
    )
    parser.add_argument("--json", action="store_true", help="Emit JSON")
    args = parser.parse_args(argv)

    root = repo_root()
    log_path = root / "assets" / "data" / "prompt_usage.log"
    ini_path = root / "assets" / "data" / "prompts.ini"

    since = None
    if args.days > 0:
        since = datetime.now() - timedelta(days=args.days)

    registry = load_prompt_chars(ini_path)
    events = parse_log(log_path, since)
    counts = Counter(e["char"] for e in events if e["char"])
    last_used: dict[str, datetime] = {}
    for e in events:
        if e["char"]:
            last_used[e["char"]] = max(last_used.get(e["char"], e["ts"]), e["ts"])

    used_chars = set(counts.keys())
    never = sorted((ch, registry[ch]) for ch in registry if ch not in used_chars)
    top = counts.most_common(15)

    if args.json:
        import json

        print(
            json.dumps(
                {
                    "window_days": args.days,
                    "total_events": len(events),
                    "top": [
                        {"char": c, "name": registry.get(c, ""), "count": n}
                        for c, n in top
                    ],
                    "never_used": [{"char": c, "name": n} for c, n in never],
                    "last_used": {
                        c: last_used[c].isoformat(sep=" ") for c in sorted(last_used)
                    },
                },
                indent=2,
            )
        )
        return 0

    window = f"last {args.days} days" if args.days > 0 else "all time"
    safe_print(f"Prompt usage ({window})")
    safe_print(f"Log: {log_path}")
    safe_print(f"Events: {len(events)}")
    safe_print("")
    safe_print("Most used:")
    if not top:
        safe_print("  (none)")
    for char, n in top:
        safe_print(f"  [{char}] {registry.get(char, '?')} — {n}")
    safe_print("")
    safe_print("Never used in window:")
    if not never:
        safe_print("  (none — all registered prompts used)")
    for char, name in never:
        safe_print(f"  [{char}] {name}")
    safe_print("")
    safe_print("Last used:")
    for char in sorted(last_used, key=lambda c: last_used[c], reverse=True)[:10]:
        safe_print(f"  [{char}] {registry.get(char, '?')} — {last_used[char]}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
