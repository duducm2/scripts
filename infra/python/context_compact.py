"""Shrink a context file for AI attach: JSON minify, text collapse, CSV keep-range.

Originals are never modified. Callers pass --in / --out paths.
CSV keep-range is applied first; compact runs after when requested.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

OMITTED_DATA_URI = "[omitted data URI]"
OMITTED_URL = "[omitted URL]"
URL_MAX_LEN = 150
_HTTP_RE = re.compile(r"^https?://", re.IGNORECASE)


def _ext(path: Path) -> str:
    return path.suffix.lower().lstrip(".")


def compact_json_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, str):
        if value.lower().startswith("data:"):
            return OMITTED_DATA_URI
        if len(value) > URL_MAX_LEN and _HTTP_RE.match(value):
            return OMITTED_URL
        return value
    if isinstance(value, list):
        out: list[Any] = []
        for item in value:
            compacted = compact_json_value(item)
            if compacted is None:
                continue
            out.append(compacted)
        return out
    if isinstance(value, dict):
        out: dict[str, Any] = {}
        for key, item in value.items():
            compacted = compact_json_value(item)
            if compacted is None:
                continue
            out[key] = compacted
        return out
    return value


def compact_json_text(text: str) -> str:
    data = json.loads(text)
    compacted = compact_json_value(data)
    return json.dumps(compacted, ensure_ascii=False, separators=(",", ":"))


def compact_text(text: str) -> str:
    """Strip trailing whitespace; collapse 3+ consecutive blank lines to one."""
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    out: list[str] = []
    blank_run = 0
    for line in lines:
        stripped = line.rstrip()
        if stripped == "":
            blank_run += 1
            if blank_run < 3:
                out.append("")
            elif blank_run == 3:
                # Already kept two blanks; drop back to a single blank line.
                if out and out[-1] == "":
                    out.pop()
            continue
        blank_run = 0
        out.append(stripped)
    return "\n".join(out)


def compact_csv_text(text: str) -> str:
    """Strip trailing whitespace and drop fully empty rows; keep the header."""
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    if not lines:
        return ""
    header = lines[0].rstrip()
    body: list[str] = []
    for line in lines[1:]:
        stripped = line.rstrip()
        if stripped == "":
            continue
        body.append(stripped)
    if not body:
        return header
    return header + "\n" + "\n".join(body)


def csv_keep_range(text: str, keep_from: int, keep_to: int) -> str:
    """Keep header (line 1) plus inclusive 1-based FROM-TO; no duplicate header."""
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    if not lines:
        return ""
    if keep_from < 1 or keep_to < 1 or keep_from > keep_to:
        return text.replace("\r\n", "\n").replace("\r", "\n")
    kept: list[str] = [lines[0]]
    for idx, line in enumerate(lines, start=1):
        if keep_from <= idx <= keep_to and idx != 1:
            kept.append(line)
    return "\n".join(kept)


def transform_text(
    text: str, ext: str, compact: bool, csv_keep: tuple[int, int] | None
) -> str:
    result = text
    if csv_keep is not None and ext == "csv":
        result = csv_keep_range(result, csv_keep[0], csv_keep[1])
    if not compact:
        return result
    if ext == "json":
        try:
            return compact_json_text(result)
        except json.JSONDecodeError:
            return compact_text(result)
    if ext == "csv":
        return compact_csv_text(result)
    return compact_text(result)


def transform_file(
    src: Path,
    dst: Path,
    compact: bool = False,
    csv_keep: tuple[int, int] | None = None,
) -> None:
    text = src.read_text(encoding="utf-8")
    out = transform_text(text, _ext(src), compact, csv_keep)
    dst.parent.mkdir(parents=True, exist_ok=True)
    dst.write_text(out, encoding="utf-8", newline="\n")


def _parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Compact a context file into a temp copy."
    )
    parser.add_argument("--in", dest="src", required=True, help="Source file path")
    parser.add_argument(
        "--out", dest="dst", required=True, help="Destination file path"
    )
    parser.add_argument(
        "--compact", action="store_true", help="Apply type-specific compact"
    )
    parser.add_argument(
        "--csv-keep",
        nargs=2,
        type=int,
        metavar=("FROM", "TO"),
        help="1-based inclusive CSV line range to keep (header always kept)",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)
    src = Path(args.src)
    dst = Path(args.dst)
    if not src.is_file():
        print(f"source not found: {src}", file=sys.stderr)
        return 1
    csv_keep = None
    if args.csv_keep is not None:
        csv_keep = (int(args.csv_keep[0]), int(args.csv_keep[1]))
    try:
        transform_file(src, dst, compact=bool(args.compact), csv_keep=csv_keep)
    except OSError as exc:
        print(f"failed: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
