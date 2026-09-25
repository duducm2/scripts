"""Wikimedia Commons search / download for Tasks project icons (no generative AI)."""

from __future__ import annotations

import json
import re
import urllib.error
import urllib.parse
import urllib.request
from collections import Counter
from typing import Any

COMMONS_API = "https://commons.wikimedia.org/w/api.php"
USER_AGENT = (
    "ScriptsTasksIcons/1.0 (local personal tasks dashboard; "
    "https://github.com/local; contact=local)"
)

STOPWORDS = {
    "a",
    "an",
    "the",
    "and",
    "or",
    "of",
    "to",
    "in",
    "on",
    "for",
    "with",
    "from",
    "by",
    "at",
    "is",
    "are",
    "was",
    "were",
    "be",
    "been",
    "this",
    "that",
    "these",
    "those",
    "it",
    "its",
    "as",
    "your",
    "you",
    "my",
    "our",
    "their",
    "his",
    "her",
    "into",
    "about",
    "over",
    "under",
    "after",
    "before",
    "when",
    "where",
    "what",
    "which",
    "who",
    "how",
    "not",
    "no",
    "yes",
    "all",
    "any",
    "each",
    "other",
    "more",
    "most",
    "some",
    "such",
    "than",
    "then",
    "also",
    "just",
    "only",
    "can",
    "will",
    "should",
    "would",
    "could",
    "may",
    "might",
    "need",
    "needs",
    "check",
    "make",
    "use",
    "used",
    "using",
    "new",
    "old",
    "day",
    "days",
    "month",
    "months",
    "year",
    "years",
    "info",
    "note",
    "notes",
    "task",
    "tasks",
    "project",
    "general",
    "daily",
    "weekly",
    "monthly",
    "schedule",
    "list",
    "plan",
    # Schedule / portion noise that crowds out identity terms (e.g. Dog → puppy feeding…)
    "meal",
    "meals",
    "feeding",
    "food",
    "per",
    "gram",
    "grams",
    "kg",
    "calorie",
    "calories",
    "adjust",
    "monitor",
    "support",
    "divided",
    "across",
    "current",
    "rapid",
    "growth",
    "adult",
}

MIME_PREF = {
    "image/svg+xml": 0,
    "image/png": 1,
    "image/webp": 2,
    "image/gif": 3,
    "image/jpeg": 8,
    "image/jpg": 8,
}

HINT_RE = re.compile(r"(icon|logo|symbol|emblem|badge|pictogram)", re.I)
TOKEN_RE = re.compile(r"[a-zA-Z][a-zA-Z0-9\-]{2,}")


def _http_get_json(url: str, timeout: float = 20.0) -> dict[str, Any]:
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8", errors="replace"))


def _http_get_bytes(url: str, timeout: float = 30.0) -> tuple[bytes, str]:
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        ctype = (resp.headers.get("Content-Type") or "").split(";")[0].strip().lower()
        return resp.read(), ctype


def tokenize(text: str) -> list[str]:
    return [
        t.lower() for t in TOKEN_RE.findall(text or "") if t.lower() not in STOPWORDS
    ]


def _stem_key(tok: str) -> str:
    """Collapse simple plurals so meal/meals don't both fill the keyword list."""
    t = (tok or "").lower()
    if len(t) > 4 and t.endswith("ies"):
        return t[:-3] + "y"
    if len(t) > 3 and t.endswith("s") and not t.endswith("ss"):
        return t[:-1]
    return t


def extract_keywords(
    project: dict[str, str],
    sections: list[dict[str, str]],
    infos: list[dict[str, str]],
    tasks: list[dict[str, str]],
    max_extra: int = 2,
) -> list[str]:
    """Title-first keywords: project title tokens, then a few content terms."""
    title_toks = tokenize(project.get("title") or "")
    weights: Counter[str] = Counter()
    for s in sections:
        for tok in tokenize(s.get("title") or ""):
            weights[tok] += 3
    for info in infos:
        for tok in tokenize(info.get("title") or ""):
            weights[tok] += 2
        for tok in tokenize((info.get("body") or "")[:400]):
            weights[tok] += 1
    for t in tasks:
        for tok in tokenize(t.get("title") or ""):
            weights[tok] += 1

    title_stems = {_stem_key(t) for t in title_toks}
    out: list[str] = []
    seen_stems: set[str] = set()
    for tok in title_toks:
        sk = _stem_key(tok)
        if sk in seen_stems:
            continue
        seen_stems.add(sk)
        out.append(tok)

    for tok, _ in weights.most_common(40):
        if len(out) >= len(title_toks) + max(0, max_extra):
            break
        sk = _stem_key(tok)
        if sk in seen_stems or sk in title_stems:
            continue
        seen_stems.add(sk)
        out.append(tok)
    return out


def _merge_results(
    existing: list[dict[str, Any]],
    incoming: list[dict[str, Any]],
    limit: int,
    seen: set[str],
) -> list[dict[str, Any]]:
    for item in incoming:
        if len(existing) >= limit:
            break
        t = item.get("title") or ""
        if not t or t in seen:
            continue
        seen.add(t)
        existing.append(item)
    return existing


def _candidate_score(item: dict[str, Any]) -> tuple:
    mime = (item.get("mime") or "").lower()
    title = item.get("title") or item.get("name") or ""
    mime_rank = MIME_PREF.get(mime, 6)
    hint = 0 if HINT_RE.search(title) else 1
    # Prefer smaller icons when size known
    size = item.get("size") or 10**12
    try:
        size_n = int(size)
    except (TypeError, ValueError):
        size_n = 10**12
    return (mime_rank, hint, size_n)


def _normalize_page(page: dict[str, Any]) -> dict[str, Any] | None:
    title = (page.get("title") or "").strip()
    if not title:
        return None
    infos = page.get("imageinfo") or []
    if not infos:
        return None
    ii = infos[0]
    mime = (ii.get("mime") or "").lower()
    if not mime.startswith("image/"):
        return None
    thumb = ii.get("thumburl") or ii.get("url") or ""
    url = ii.get("url") or thumb
    if not url:
        return None
    name = title[5:] if title.lower().startswith("file:") else title
    return {
        "title": title,
        "name": name,
        "mime": mime,
        "url": url,
        "thumburl": thumb or url,
        "width": ii.get("width"),
        "height": ii.get("height"),
        "size": ii.get("size"),
        "descriptionurl": ii.get("descriptionurl") or "",
    }


def search_commons(query: str, limit: int = 20) -> list[dict[str, Any]]:
    q = (query or "").strip()
    if not q:
        return []
    limit = max(1, min(int(limit or 20), 40))
    # Fetch extra so MIME ranking still yields enough after demoting JPEG photos / non-images.
    fetch_n = min(80, max(limit * 4, limit + 20))
    # Prefer file types Commons search understands; still filter MIME client-side.
    search_q = q if "filetype:" in q.lower() else f"{q} filetype:bitmap|drawing"
    params = {
        "action": "query",
        "format": "json",
        "generator": "search",
        "gsrsearch": search_q,
        "gsrnamespace": "6",
        "gsrlimit": str(fetch_n),
        "prop": "imageinfo",
        "iiprop": "url|mime|size|dimensions",
        "iiurlwidth": "128",
    }
    url = COMMONS_API + "?" + urllib.parse.urlencode(params)
    try:
        data = _http_get_json(url)
    except (
        urllib.error.URLError,
        urllib.error.HTTPError,
        TimeoutError,
        json.JSONDecodeError,
        OSError,
    ):
        return []
    pages = (data.get("query") or {}).get("pages") or {}
    items: list[dict[str, Any]] = []
    for page in pages.values():
        if not isinstance(page, dict):
            continue
        norm = _normalize_page(page)
        if norm:
            items.append(norm)
    items.sort(key=_candidate_score)
    preferred = [
        i
        for i in items
        if (i.get("mime") or "") in {"image/png", "image/svg+xml", "image/webp"}
    ]
    rest = [i for i in items if i not in preferred]
    ordered = preferred + rest
    return ordered[:limit]


def search_with_style_preference(query: str, limit: int = 20) -> list[dict[str, Any]]:
    """Prefer icon-like results; try a light 3D/game boost, then fall back to plain q."""
    q = (query or "").strip()
    if not q:
        return []
    variants: list[str] = []
    ql = q.lower()
    if "3d" not in ql and "game" not in ql:
        variants.append(f"{q} 3D game icon")
    if "icon" not in ql:
        variants.append(f"{q} icon")
    # Always try unmodified query last so long/noisy phrases can still hit.
    variants.append(q)
    seen: set[str] = set()
    merged: list[dict[str, Any]] = []
    for variant in variants:
        for item in search_commons(variant, limit=limit):
            t = item.get("title") or ""
            if t in seen:
                continue
            seen.add(t)
            merged.append(item)
            if len(merged) >= limit:
                return merged
    return merged


def suggest_for_project(
    project: dict[str, str],
    sections: list[dict[str, str]],
    infos: list[dict[str, str]],
    tasks: list[dict[str, str]],
    limit: int = 5,
) -> dict[str, Any]:
    """Cascade short title-first Commons queries until `limit` unique images."""
    limit = max(1, min(int(limit or 5), 10))
    title = (project.get("title") or "").strip()
    keywords = extract_keywords(project, sections, infos, tasks, max_extra=2)

    cascade: list[str] = []
    if title:
        cascade.append(title)
        if "icon" not in title.lower():
            cascade.append(f"{title} icon")
    # Short phrase: title tokens + up to 2 content terms (already title-first in keywords)
    if keywords:
        short = " ".join(keywords[: max(1, min(3, len(keywords)))])
        if short and short.lower() not in {c.lower() for c in cascade}:
            cascade.append(short)
        # Content-only fallback if title was empty/generic
        content_only = " ".join(k for k in keywords if k.lower() not in title.lower())
        if content_only and content_only.lower() not in {c.lower() for c in cascade}:
            cascade.append(content_only)

    seen: set[str] = set()
    results: list[dict[str, Any]] = []
    used_query = title or (keywords[0] if keywords else "")
    for step in cascade:
        if len(results) >= limit:
            break
        before = len(results)
        _merge_results(
            results, search_with_style_preference(step, limit=limit), limit, seen
        )
        if len(results) > before:
            used_query = step
        if len(results) >= limit:
            break
        # Plain Commons if style boost still thin for this step
        before = len(results)
        _merge_results(results, search_commons(step, limit=limit), limit, seen)
        if len(results) > before:
            used_query = step

    return {
        "ok": True,
        "query": used_query,
        "keywords": keywords,
        "results": results[:limit],
    }


def resolve_download_url(url: str = "", commons_title: str = "") -> tuple[str, str]:
    """Return (download_url, suggested_ext)."""
    title = (commons_title or "").strip()
    if title and not title.lower().startswith("file:"):
        title = "File:" + title
    if title:
        params = {
            "action": "query",
            "format": "json",
            "titles": title,
            "prop": "imageinfo",
            "iiprop": "url|mime",
        }
        api_url = COMMONS_API + "?" + urllib.parse.urlencode(params)
        data = _http_get_json(api_url)
        pages = (data.get("query") or {}).get("pages") or {}
        for page in pages.values():
            infos = (page or {}).get("imageinfo") or []
            if not infos:
                continue
            ii = infos[0]
            mime = (ii.get("mime") or "").lower()
            dl = ii.get("url") or ""
            if not dl:
                continue
            ext = _ext_from_mime(mime, dl)
            return dl, ext
        raise ValueError("commons title not found")
    raw = (url or "").strip()
    if not raw.startswith("http://") and not raw.startswith("https://"):
        raise ValueError("url or commons_title required")
    # Only allow Wikimedia hosts
    host = urllib.parse.urlparse(raw).netloc.lower()
    if not (
        host.endswith("wikimedia.org")
        or host.endswith("wikipedia.org")
        or host.endswith("wikimedia.org.")
    ):
        raise ValueError("only Wikimedia Commons URLs allowed")
    return raw, _ext_from_url(raw)


def _ext_from_mime(mime: str, url: str = "") -> str:
    mime = (mime or "").lower()
    if "svg" in mime:
        return ".svg"
    if "png" in mime:
        return ".png"
    if "webp" in mime:
        return ".webp"
    if "gif" in mime:
        return ".gif"
    if "jpeg" in mime or "jpg" in mime:
        return ".jpg"
    return _ext_from_url(url) or ".png"


def _ext_from_url(url: str) -> str:
    path = urllib.parse.urlparse(url).path.lower()
    for ext in (".svg", ".png", ".webp", ".gif", ".jpg", ".jpeg"):
        if path.endswith(ext):
            return ".jpg" if ext == ".jpeg" else ext
    return ".png"


def download_image(url: str = "", commons_title: str = "") -> tuple[bytes, str]:
    dl_url, ext = resolve_download_url(url=url, commons_title=commons_title)
    data, ctype = _http_get_bytes(dl_url)
    if not data:
        raise ValueError("empty download")
    if ctype:
        ext = _ext_from_mime(ctype, dl_url) or ext
    # Cap size ~4MB
    if len(data) > 4 * 1024 * 1024:
        raise ValueError("image too large")
    return data, ext
