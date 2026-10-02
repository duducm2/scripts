# -*- coding: utf-8 -*-
"""Write entertainment lists as collapsed phone accordions."""
from __future__ import annotations

import argparse
import csv
import html
from pathlib import Path


def read_rows(path: Path) -> list[dict[str, str]]:
    with path.open(encoding="utf-8-sig", newline="") as handle:
        return [
            {k: (v or "") for k, v in row.items()} for row in csv.DictReader(handle)
        ]


def sort_key(row: dict[str, str]) -> tuple:
    try:
        order = int(row.get("sort_order") or 0)
    except ValueError:
        order = 0
    return (order, (row.get("title") or "").lower())


def children(rows: list[dict[str, str]], parent_id: str) -> list[dict[str, str]]:
    kids = [row for row in rows if (row.get("parent_id") or "") == parent_id]
    kids.sort(key=sort_key)
    return kids


def has_kids(rows: list[dict[str, str]], row_id: str) -> bool:
    return any((row.get("parent_id") or "") == row_id for row in rows)


def hidden(rows: list[dict[str, str]], row: dict[str, str]) -> bool:
    if (row.get("title") or "").strip().lower() != "backlog":
        return False
    return (
        not (row.get("url") or "").strip()
        and not (row.get("notes") or "").strip()
        and not has_kids(rows, row.get("id") or "")
    )


def item_html(row: dict[str, str]) -> str:
    title = html.escape((row.get("title") or "").strip())
    if (row.get("done") or "").strip() == "1":
        title = "✓ " + title
    image = (row.get("image") or "").strip()
    cover = ""
    if image:
        cover = f'<img src="{html.escape(image, quote=True)}" alt="" width="64"> '
    return f"<p>{cover}{title}</p>"


def details(summary: str, body: str) -> str:
    inner = body.strip()
    if inner:
        inner = "\n\n" + inner + "\n"
    else:
        inner = "\n"
    return f"<details>\n<summary>{html.escape(summary)}</summary>{inner}</details>"


def loose_label(shelf: str) -> str:
    name = shelf.strip().lower()
    if name == "tv shows":
        return "BBC"
    if name == "books":
        return "Big Read"
    return ""


def branch_html(rows: list[dict[str, str]], node: dict[str, str]) -> str:
    kids = [
        row for row in children(rows, node.get("id") or "") if not hidden(rows, row)
    ]
    label = loose_label(node.get("title") or "")
    parts: list[str] = []
    loose: list[dict[str, str]] = []

    def flush() -> None:
        if not loose:
            return
        body = "\n".join(item_html(row) for row in loose)
        parts.append(details(label, body) if label else body)
        loose.clear()

    for kid in kids:
        if has_kids(rows, kid.get("id") or ""):
            flush()
            parts.append(branch_html(rows, kid))
        else:
            loose.append(kid)
    flush()
    return details((node.get("title") or "").strip(), "\n\n".join(parts))


def render(rows: list[dict[str, str]]) -> str:
    shelves = [row for row in children(rows, "") if not hidden(rows, row)]
    blocks = [branch_html(rows, shelf) for shelf in shelves]
    return "# Entertainment\n\n" + "\n\n".join(blocks) + "\n"


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--csv", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(render(read_rows(Path(args.csv))), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
