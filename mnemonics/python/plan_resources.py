"""Parse, classify, and format study plan resource lines."""

from __future__ import annotations

import re
from typing import Any
from urllib.parse import urlparse

MD_LINK_RE = re.compile(r"^\s*[-*]?\s*\[([^\]]*)\]\(([^)]+)\)\s*$")
PLAIN_URL_RE = re.compile(r"^\s*[-*]?\s*(https?://\S+)\s*$")

KIND_EMOJI: dict[str, str] = {
    "youtube": "▶",
    "book": "📖",
    "docs": "📄",
    "article": "📝",
    "search": "🔍",
    "link": "🔗",
}


def parse_resource_line(line: str) -> tuple[str, str]:
    """Return (title, url) from a markdown list line or plain URL."""
    raw = (line or "").strip()
    if not raw:
        return "", ""
    m = MD_LINK_RE.match(raw)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    m = PLAIN_URL_RE.match(raw)
    if m:
        url = m.group(1).strip()
        return url, url
    return raw.lstrip("-* ").strip(), ""


def build_resource_line(title: str, url: str) -> str:
    title = (title or "").strip()
    url = (url or "").strip()
    if not url:
        return f"- {title}" if title else ""
    if not title:
        title = url
    return f"- [{title}]({url})"


def classify_url(url: str, title: str = "") -> str:
    """Infer resource kind from URL and optional title."""
    u = (url or "").lower()
    t = (title or "").lower()
    host = (urlparse(u).netloc or "").lower()

    if "youtube.com/results" in u or "search_query=" in u:
        return "search"
    if "youtube.com" in host or "youtu.be" in host:
        return "youtube"
    if "(book)" in t or "/books/" in u or "goodreads.com" in host or "amazon." in host:
        return "book"
    if (
        "scrumguides" in host
        or host.endswith(".org")
        or host.endswith(".dev")
        or "/docs/" in u
        or "learning.postman.com" in host
        or "developer.chrome.com" in host
        or "scrum.org" in host
    ):
        return "docs"
    if (
        "/blog/" in u
        or "medium.com" in host
        or "substack.com" in host
        or "intercom.com/blog" in u
        or "amplitude.com/blog" in u
    ):
        return "article"
    return "link"


def export_emoji(kind: str) -> str:
    return KIND_EMOJI.get(kind or "link", KIND_EMOJI["link"])


def format_export_line(line: str) -> str:
    """Markdown list line with kind emoji prefix for GitHub mobile."""
    title, url = parse_resource_line(line)
    if not url:
        return line if line.startswith("-") else f"- {line}"
    kind = classify_url(url, title)
    emoji = export_emoji(kind)
    return f"- {emoji} [{title}]({url})"


def enrich_resource_row(row: dict[str, str]) -> dict[str, Any]:
    """Turn a plan_resources CSV row into a structured dashboard payload."""
    line = row.get("line") or ""
    title, url = parse_resource_line(line)
    kind = classify_url(url, title)
    return {
        "id": row.get("id") or "",
        "line": line,
        "title": title,
        "url": url,
        "kind": kind,
        "icon": kind,
        "section_path": row.get("section_path") or "",
        "sort_order": row.get("sort_order") or "0",
    }


def enrich_resource_lines(lines: list[str]) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    for i, line in enumerate(lines):
        title, url = parse_resource_line(line)
        kind = classify_url(url, title)
        out.append(
            {
                "id": "",
                "line": line,
                "title": title,
                "url": url,
                "kind": kind,
                "icon": kind,
                "section_path": "",
                "sort_order": str(i + 1),
            }
        )
    return out
