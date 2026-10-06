"""One cached Openverse picture per tangible keyword.

The key is the keyword text folded to lower case with spaces collapsed.
A missing keyword is fetched once. An empty Openverse result is recorded
so later saves do not ask again. A click in the palace can still search
and replace the canonical file.
"""

from __future__ import annotations

import csv
import hashlib
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

_manifest_lock = threading.Lock()
_api_lock = threading.Lock()
_next_request_at = 0.0


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


def search_images(query: str, page_size: int = 1) -> list[dict]:
    """Return simplified Openverse hits. Does not write the cache."""
    key = keyword_key(query)
    if not key:
        return []
    params = urllib.parse.urlencode(
        {
            "q": key,
            "page_size": str(max(1, min(page_size, 20))),
            "mature": "false",
        }
    )
    payload = _get_json(f"{OPENVERSE_SEARCH}?{params}")
    hits: list[dict] = []
    for row in payload.get("results") or []:
        if not isinstance(row, dict):
            continue
        image_id = str(row.get("id") or "").strip()
        thumb = str(row.get("thumbnail") or "").strip()
        direct = str(row.get("url") or "").strip()
        image = direct or thumb
        if thumb and "api.openverse.org" not in thumb:
            image = thumb
        if not image_id or not image:
            continue
        hits.append(
            {
                "id": image_id,
                "thumbnail": thumb or image,
                "image": image,
                "title": str(row.get("title") or ""),
                "source_page": str(row.get("foreign_landing_url") or ""),
            }
        )
    return hits


def _extension_for(data: bytes, content_type: str) -> str:
    ctype = (content_type or "").split(";", 1)[0].strip().lower()
    known = {
        "image/jpeg": ".jpg",
        "image/jpg": ".jpg",
        "image/png": ".png",
        "image/webp": ".webp",
        "image/gif": ".gif",
    }
    if ctype in known:
        return known[ctype]
    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        return ".png"
    if data.startswith(b"\xff\xd8"):
        return ".jpg"
    if data.startswith(b"RIFF") and data[8:12] == b"WEBP":
        return ".webp"
    if data.startswith(b"GIF8"):
        return ".gif"
    return ".jpg"


def _download(url: str, stem: Path) -> str:
    if "api.openverse.org" in url:
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
    IMAGE_DIR.mkdir(parents=True, exist_ok=True)
    for old in IMAGE_DIR.glob(stem.name + ".*"):
        if old.suffix.lower() != ".json":
            old.unlink()
    path = stem.with_suffix(ext)
    path.write_bytes(data)
    return path.name


def _store_hit(key: str, hit: dict, *, bump: bool) -> dict:
    with _manifest_lock:
        data = load_manifest()
        keywords = data["keywords"]
        previous = keywords.get(key) or {}
        file_slug = str(previous.get("slug") or "") or _file_slug(key, keywords)
        filename = _download(
            hit.get("image") or hit["thumbnail"], IMAGE_DIR / file_slug
        )
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


def _load_atom_rows() -> list[dict]:
    if not ATOMS_CSV.is_file():
        return []
    with ATOMS_CSV.open(encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def main() -> None:
    rows = _load_atom_rows()
    result = backfill_atoms(rows)
    print(
        f"unique={result['unique']} fetched={result['fetched']} "
        f"remaining={result['remaining']} {result['stopped']}".rstrip()
    )


if __name__ == "__main__":
    main()
