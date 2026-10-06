"""One cached Openverse picture per tangible keyword.

The key is the keyword text folded to lower case with spaces collapsed.
A missing keyword is fetched once, preferring an Openverse illustration
and falling back to any picture of that word. An empty result is recorded
so later saves do not ask again. A click in the palace can still search
and replace the canonical file.
"""

from __future__ import annotations

import csv
import hashlib
import html
import json
import os
import re
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

from beast_thumb_base import slug
from schemas import iter_keyword_pairs

ROOT = Path(__file__).resolve().parents[1]
IMAGE_DIR = ROOT / "web" / "assets" / "keyword-images"
MANIFEST_PATH = IMAGE_DIR / "manifest.json"
ATOMS_CSV = ROOT / "data" / "atoms.csv"
OPENVERSE_SEARCH = "https://api.openverse.org/v1/images/"
USER_AGENT = "MemoryPalace/1.0 (local study tool; keyword icons)"
# Anonymous Openverse is about 20 requests per minute.
MIN_INTERVAL_S = 3.05
# Drawings first. A keyword with no illustration still gets a plain picture.
RASTER_SUFFIXES = (".jpg", ".jpeg", ".png", ".gif", ".webp")
# Drawings first. A keyword with no illustration still gets a plain picture.
PREFERRED_CATEGORY = "illustration"

_manifest_lock = threading.Lock()
_api_lock = threading.Lock()
_next_request_at = 0.0
_next_wiki_at = 0.0
WIKI_INTERVAL_S = 4.0


class OpenverseRateLimit(RuntimeError):
    pass


def keyword_key(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").strip().lower())


def _empty_manifest() -> dict:
    return {"version": 1, "keywords": {}}


def load_manifest() -> dict:
    if not MANIFEST_PATH.is_file():
        return _empty_manifest()
    try:
        data = json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return _empty_manifest()
    if not isinstance(data, dict):
        return _empty_manifest()
    data.setdefault("version", 1)
    keywords = data.get("keywords")
    if not isinstance(keywords, dict):
        data["keywords"] = {}
    return data


def save_manifest(data: dict) -> None:
    IMAGE_DIR.mkdir(parents=True, exist_ok=True)
    tmp = MANIFEST_PATH.with_suffix(".json.tmp")
    tmp.write_text(
        json.dumps(data, indent=2, ensure_ascii=False) + "\n", encoding="utf-8"
    )
    tmp.replace(MANIFEST_PATH)


def practice_markdown_row(keywords: str | None) -> str:
    """One picture per tangible keyword, in stored order, for study practice files.

    Paths are relative to mnemonics/output/practice/*.md. Words with no cached
    file are skipped. Quick Recall does not use this.
    """
    stored = load_manifest().get("keywords", {})
    parts: list[str] = []
    for left, _right in iter_keyword_pairs(keywords or ""):
        entry = stored.get(keyword_key(left)) or {}
        filename = str(entry.get("file") or "")
        if entry.get("empty") or not filename:
            continue
        if "/" in filename or "\\" in filename or filename.startswith("."):
            continue
        if not (IMAGE_DIR / filename).is_file():
            continue
        alt = html.escape(left, quote=True)
        src = f"../../web/assets/keyword-images/{filename}"
        parts.append(
            f'<img src="{src}" alt="{alt}" width="72" height="72" '
            f'style="vertical-align:middle;height:72px;width:72px;" />'
        )
    return " ".join(parts)


def public_map() -> dict[str, dict]:
    """Keyword key -> {url, version} for pictures the page can show."""
    out: dict[str, dict] = {}
    for key, entry in load_manifest().get("keywords", {}).items():
        filename = str(entry.get("file") or "")
        if entry.get("empty") or not filename:
            continue
        out[key] = {
            "url": f"/assets/keyword-images/{filename}",
            "version": int(entry.get("version") or 1),
        }
    return out


def _file_slug(key: str, keywords: dict) -> str:
    base = slug(key) or "keyword"
    for other_key, entry in keywords.items():
        if other_key != key and entry.get("slug") == base:
            digest = hashlib.sha1(key.encode("utf-8")).hexdigest()[:6]
            return f"{base}_{digest}"
    return base


def _throttle() -> None:
    global _next_request_at
    with _api_lock:
        now = time.monotonic()
        wait = _next_request_at - now
        if wait > 0:
            time.sleep(wait)
            now = time.monotonic()
        _next_request_at = now + MIN_INTERVAL_S


def _throttle_wikimedia() -> None:
    global _next_wiki_at
    with _api_lock:
        now = time.monotonic()
        wait = _next_wiki_at - now
        if wait > 0:
            time.sleep(wait)
            now = time.monotonic()
        _next_wiki_at = now + WIKI_INTERVAL_S


def _get_json(url: str) -> dict:
    _throttle()
    req = urllib.request.Request(
        url, headers={"User-Agent": USER_AGENT, "Accept": "application/json"}
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            if resp.status == 429:
                raise OpenverseRateLimit("Openverse rate limit")
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        if exc.code == 429:
            raise OpenverseRateLimit("Openverse rate limit") from exc
        raise


def _search_openverse(
    query: str,
    page_size: int,
    *,
    category: str | None,
    style: str,
    excluded_source: str | None = None,
) -> list[dict]:
    params = {
        "q": query,
        "page_size": "20",
        "mature": "false",
    }
    if category:
        params["category"] = category
    if excluded_source:
        params["excluded_source"] = excluded_source
    payload = _get_json(f"{OPENVERSE_SEARCH}?{urllib.parse.urlencode(params)}")
    hits: list[dict] = []
    for row in payload.get("results") or []:
        if not isinstance(row, dict):
            continue
        image_id = str(row.get("id") or "").strip()
        thumb = str(row.get("thumbnail") or "").strip()
        direct = str(row.get("url") or "").strip()
        display = _raster_display_url(direct, thumb)
        image = direct or display or ""
        if not image_id or not display:
            continue
        hits.append(
            {
                "id": image_id,
                "thumbnail": display,
                "image": image,
                "title": str(row.get("title") or ""),
                "source_page": str(row.get("foreign_landing_url") or ""),
                "style": style,
            }
        )
        if len(hits) >= page_size:
            break
    return hits


def search_images(query: str, page_size: int = 1) -> list[dict]:
    """Return Openverse hits, drawings first, then any picture of the keyword.

    Does not write the cache. A full page of illustrations stops there.
    Fewer than requested, including none, is filled from an ordinary search.
    """
    key = keyword_key(query)
    if not key:
        return []
    size = max(1, min(page_size, 20))
    preferred = _search_openverse(
        key, size, category=PREFERRED_CATEGORY, style=PREFERRED_CATEGORY
    )
    if len(preferred) >= size:
        return preferred
    plain = _search_openverse(key, size, category=None, style="any")
    if not preferred:
        return plain
    seen = {hit["id"] for hit in preferred}
    merged = list(preferred)
    for hit in plain:
        if hit["id"] in seen:
            continue
        merged.append(hit)
        if len(merged) >= size:
            break
    return merged


def _url_suffix(url: str) -> str:
    path = urllib.parse.urlparse(url).path.lower()
    return Path(path).suffix


def _raster_display_url(direct: str, thumb: str) -> str | None:
    """A picture an <img> can paint. API thumbs and raw SVG show the title instead."""
    wiki = _wikimedia_thumb(direct) or _wikimedia_thumb(thumb)
    if wiki:
        return wiki
    for url in (thumb, direct):
        if not url or "api.openverse.org" in url:
            continue
        if _url_suffix(url) in RASTER_SUFFIXES:
            return url
    return None


def _wikimedia_thumb(url: str, width: int = 330) -> str | None:
    """Commons original file URL to a small PNG or JPEG thumbnail."""
    if not url or "/thumb/" in url:
        return None
    match = re.match(
        r"^https://upload\.wikimedia\.org/wikipedia/([^/]+)/"
        r"([0-9a-f]/[0-9a-f]{2})/([^/?#]+)$",
        url.split("?", 1)[0],
        re.IGNORECASE,
    )
    if not match:
        return None
    project, shard, name = match.groups()
    thumb_name = f"{width}px-{name}"
    if name.lower().endswith(".svg"):
        thumb_name += ".png"
    return (
        f"https://upload.wikimedia.org/wikipedia/{project}/thumb/"
        f"{shard}/{name}/{thumb_name}"
    )


def _image_urls(hit: dict) -> list[str]:
    """Smallest usable file first. Wikimedia originals are easy to get blocked on."""
    direct = str(hit.get("image") or "").strip()
    thumb = str(hit.get("thumbnail") or "").strip()
    urls: list[str] = []
    for candidate in (
        _wikimedia_thumb(direct),
        _wikimedia_thumb(thumb),
        thumb if thumb and "api.openverse.org" not in thumb else "",
        direct if not direct.lower().split("?", 1)[0].endswith(".svg") else "",
        thumb,
        direct,
    ):
        if candidate and candidate not in urls:
            urls.append(candidate)
    return urls


def _extension_for(data: bytes, content_type: str) -> str | None:
    """Raster type from the file bytes. SVG and HTML are not pictures."""
    del content_type
    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        return ".png"
    if data.startswith(b"\xff\xd8"):
        return ".jpg"
    if len(data) >= 12 and data.startswith(b"RIFF") and data[8:12] == b"WEBP":
        return ".webp"
    if data.startswith(b"GIF8"):
        return ".gif"
    return None


def _download(url: str, stem: Path) -> str:
    if "wikimedia.org" in url:
        _throttle_wikimedia()
    elif "api.openverse.org" in url:
        _throttle()
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=30) as resp:
        if resp.status == 429:
            raise OpenverseRateLimit("Openverse rate limit")
        data = resp.read()
        ctype = resp.headers.get_content_type()
    if not data:
        raise RuntimeError("empty image")
    ext = _extension_for(data, ctype)
    if not ext:
        raise RuntimeError("not a raster image")
    IMAGE_DIR.mkdir(parents=True, exist_ok=True)
    for old in IMAGE_DIR.glob(stem.name + ".*"):
        if old.suffix.lower() != ".json":
            old.unlink()
    path = stem.with_suffix(ext)
    path.write_bytes(data)
    return path.name


def _download_hit(hit: dict, stem: Path) -> str:
    urls = [
        url
        for url in _image_urls(hit)
        if not ("wikimedia.org" in url and "/thumb/" not in url)
    ]
    if not urls:
        urls = _image_urls(hit)
    if not urls:
        raise RuntimeError("empty image")
    last: Exception | None = None
    for url in urls:
        try:
            return _download(url, stem)
        except OpenverseRateLimit:
            raise
        except urllib.error.HTTPError as exc:
            last = exc
            if exc.code == 429 and (
                "wikimedia.org" in url or "api.openverse.org" in url
            ):
                raise OpenverseRateLimit("image host rate limit") from exc
            continue
        except (OSError, urllib.error.URLError, TimeoutError, RuntimeError) as exc:
            last = exc
            continue
    if last:
        raise last
    raise RuntimeError("empty image")


def _store_hit(key: str, hit: dict, *, bump: bool) -> dict:
    with _manifest_lock:
        data = load_manifest()
        keywords = data["keywords"]
        previous = keywords.get(key) or {}
        file_slug = str(previous.get("slug") or "") or _file_slug(key, keywords)
        filename = _download_hit(hit, IMAGE_DIR / file_slug)
        version = int(previous.get("version") or 0)
        if bump or previous.get("file") != filename:
            version += 1
        if version < 1:
            version = 1
        entry = {
            "key": key,
            "slug": file_slug,
            "file": filename,
            "openverse_id": hit["id"],
            "source_page": hit.get("source_page") or "",
            "version": version,
            "empty": False,
            "style": str(hit.get("style") or previous.get("style") or ""),
        }
        keywords[key] = entry
        save_manifest(data)
        return entry


def _record_empty(key: str) -> dict:
    with _manifest_lock:
        data = load_manifest()
        entry = {
            "key": key,
            "slug": "",
            "file": "",
            "openverse_id": "",
            "source_page": "",
            "version": 0,
            "empty": True,
        }
        data["keywords"][key] = entry
        save_manifest(data)
        return entry


def _known(key: str) -> bool:
    with _manifest_lock:
        return key in load_manifest().get("keywords", {})


def ensure_keywords(words: list[str]) -> None:
    """Fetch the first Openverse hit for keywords that are not cached yet.

    Skipped under pytest so atom-save tests do not call the network.
    Network and rate-limit failures leave the atom save intact.
    """
    if os.environ.get("PYTEST_CURRENT_TEST"):
        return
    seen: set[str] = set()
    for word in words:
        key = keyword_key(word)
        if not key or key in seen or _known(key):
            continue
        seen.add(key)
        try:
            hits = search_images(key, 1)
        except OpenverseRateLimit:
            raise
        except (
            OSError,
            urllib.error.URLError,
            json.JSONDecodeError,
            TimeoutError,
            RuntimeError,
        ):
            continue
        if _known(key):
            continue
        try:
            if hits:
                _store_hit(key, hits[0], bump=False)
            else:
                _record_empty(key)
        except OpenverseRateLimit:
            raise
        except (OSError, urllib.error.URLError, TimeoutError, RuntimeError):
            continue


def choose_keyword_image(keyword: str, openverse_id: str) -> dict:
    """Download one of the current top results and make it canonical."""
    key = keyword_key(keyword)
    image_id = str(openverse_id or "").strip()
    if not key or not image_id:
        return {"ok": False, "error": "keyword and openverse_id required"}
    hits = search_images(key, 15)
    hit = next((row for row in hits if row["id"] == image_id), None)
    if not hit:
        return {"ok": False, "error": "that result is no longer in the top 15"}
    entry = _store_hit(key, hit, bump=True)
    return {
        "ok": True,
        "key": key,
        "image": {
            "url": f"/assets/keyword-images/{entry['file']}",
            "version": entry["version"],
        },
    }


def _store_some(key: str, hits: list[dict]) -> bool:
    for hit in hits:
        try:
            _store_hit(key, hit, bump=True)
            return True
        except OpenverseRateLimit:
            raise
        except (
            OSError,
            urllib.error.URLError,
            TimeoutError,
            RuntimeError,
        ):
            continue
    return False


def _fallback_hits(key: str) -> list[dict]:
    """Drawings from other hosts, then any non-Wikimedia picture of the word."""
    drawings = _search_openverse(
        key,
        5,
        category=PREFERRED_CATEGORY,
        style=PREFERRED_CATEGORY,
        excluded_source="wikimedia",
    )
    if drawings:
        return drawings
    return _search_openverse(
        key, 5, category=None, style="any", excluded_source="wikimedia"
    )


def _legacy_keys() -> list[str]:
    """Cached pictures from before the drawing preference."""
    pending: list[str] = []
    for key, entry in load_manifest().get("keywords", {}).items():
        if not isinstance(entry, dict) or entry.get("empty") or entry.get("style"):
            continue
        pending.append(str(key))
    return pending


def _mark_style(key: str, style: str) -> None:
    """Remember that a legacy picture was checked, and keep its file."""
    with _manifest_lock:
        data = load_manifest()
        entry = data["keywords"].get(key)
        if not isinstance(entry, dict):
            return
        entry["style"] = style
        save_manifest(data)


def refresh_legacy() -> dict:
    """Replace each cached picture with a drawing, or any picture of that word.

    Words with no Openverse hit keep the file they already have. A later run
    skips pictures that already record a style, so a rate limit can resume.
    """
    pending = _legacy_keys()
    total = len(pending)
    updated = 0
    kept = 0
    try:
        for index, key in enumerate(pending, start=1):
            if key not in _legacy_keys():
                continue
            print(f"{index}/{total} {key}", flush=True)
            try:
                hits = search_images(key, 1)
                saved = _store_some(key, hits) if hits else False
                if not saved:
                    saved = _store_some(key, _fallback_hits(key))
                if saved:
                    updated += 1
                elif not hits:
                    _mark_style(key, "any")
                    kept += 1
                else:
                    print(f"skip {key}: no downloadable picture", flush=True)
            except OpenverseRateLimit:
                try:
                    if _store_some(key, _fallback_hits(key)):
                        updated += 1
                        continue
                except OpenverseRateLimit:
                    raise
                print(f"skip {key}: rate limit", flush=True)
                continue
            except (
                OSError,
                urllib.error.URLError,
                json.JSONDecodeError,
                TimeoutError,
                RuntimeError,
            ) as exc:
                print(f"skip {key}: {exc}", flush=True)
                continue
    except OpenverseRateLimit:
        return {
            "ok": True,
            "pending": total,
            "updated": updated,
            "kept": kept,
            "remaining": len(_legacy_keys()),
            "stopped": "rate limit",
        }
    return {
        "ok": True,
        "pending": total,
        "updated": updated,
        "kept": kept,
        "remaining": len(_legacy_keys()),
        "stopped": "",
    }


def keywords_in_atoms(rows: list[dict]) -> list[str]:
    """Unique tangible keywords in first-seen order."""
    found: list[str] = []
    seen: set[str] = set()
    for row in rows:
        for left, _right in iter_keyword_pairs(row.get("keywords")):
            key = keyword_key(left)
            if key and key not in seen:
                seen.add(key)
                found.append(key)
    return found


def backfill_atoms(rows: list[dict]) -> dict:
    """Fetch every stored keyword that is not already in the manifest."""
    words = keywords_in_atoms(rows)
    missing = [word for word in words if not _known(word)]
    done = 0
    try:
        ensure_keywords(missing)
        done = sum(1 for word in missing if _known(word))
    except OpenverseRateLimit:
        done = sum(1 for word in missing if _known(word))
        return {
            "ok": True,
            "unique": len(words),
            "fetched": done,
            "remaining": len(words) - sum(1 for word in words if _known(word)),
            "stopped": "rate limit",
        }
    remaining = len(words) - sum(1 for word in words if _known(word))
    return {
        "ok": True,
        "unique": len(words),
        "fetched": done,
        "remaining": remaining,
        "stopped": "",
    }


def broken_cached_keys() -> list[str]:
    """Keywords whose saved file is missing or not a real picture."""
    pending: list[str] = []
    for key, entry in load_manifest().get("keywords", {}).items():
        if not isinstance(entry, dict) or entry.get("empty"):
            continue
        filename = str(entry.get("file") or "")
        path = IMAGE_DIR / filename if filename else None
        if path is None or not path.is_file():
            pending.append(str(key))
            continue
        if _extension_for(path.read_bytes(), "") is None:
            pending.append(str(key))
    return pending


def repair_broken() -> dict:
    """Replace saved SVG or HTML files with a real picture."""
    keys = broken_cached_keys()
    with _manifest_lock:
        data = load_manifest()
        for key in keys:
            entry = data["keywords"].get(key)
            if isinstance(entry, dict):
                entry.pop("style", None)
        save_manifest(data)
    return refresh_legacy()


def _load_atom_rows() -> list[dict]:
    if not ATOMS_CSV.is_file():
        return []
    with ATOMS_CSV.open(encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def main(argv: list[str] | None = None) -> None:
    import argparse

    parser = argparse.ArgumentParser(description="Cache one picture per keyword")
    parser.add_argument(
        "--refresh-legacy",
        action="store_true",
        help="Replace pictures saved before the drawing preference",
    )
    parser.add_argument(
        "--repair-broken",
        action="store_true",
        help="Replace saved files that are not real pictures",
    )
    args = parser.parse_args(argv)
    if args.repair_broken:
        result = repair_broken()
        print(
            f"pending={result['pending']} updated={result['updated']} "
            f"kept={result['kept']} remaining={result['remaining']} "
            f"{result['stopped']}".rstrip()
        )
        return
    if args.refresh_legacy:
        result = refresh_legacy()
        print(
            f"pending={result['pending']} updated={result['updated']} "
            f"kept={result['kept']} remaining={result['remaining']} "
            f"{result['stopped']}".rstrip()
        )
        return
    rows = _load_atom_rows()
    result = backfill_atoms(rows)
    print(
        f"unique={result['unique']} fetched={result['fetched']} "
        f"remaining={result['remaining']} {result['stopped']}".rstrip()
    )


if __name__ == "__main__":
    main()
