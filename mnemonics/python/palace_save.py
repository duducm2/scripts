"""Save palace notes and gallery images from the dashboard HTTP API."""

from __future__ import annotations

import base64
import sys
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent))

from data_aggregator import load_all, save_all  # noqa: E402
from plan_csv import next_id  # noqa: E402
from study_practice_md import slug_filename, write_study  # noqa: E402

PRACTICE_PREFIX = "practice/images/"


def _find_palace(data: dict[str, list[dict[str, str]]], palace_id: str) -> dict[str, str] | None:
    return next((p for p in data["palaces"] if p.get("id") == palace_id), None)


def _study_slug(data: dict[str, list[dict[str, str]]], study_id: str) -> str:
    study = next((s for s in data["studies"] if s.get("id") == study_id), None)
    if not study:
        return ""
    return slug_filename(study.get("notes_rel_path") or study_id)


def _sync_study(
    study_id: str,
    data_dir: Path,
    output_dir: Path,
    notes_root: Path | None,
) -> None:
    if study_id:
        write_study(study_id, data_dir, output_dir, notes_root)


def _next_sort_order(rows: list[dict[str, str]], palace_id: str) -> int:
    orders = [
        int(r.get("sort_order") or 0)
        for r in rows
        if r.get("palace_id") == palace_id
    ]
    return (max(orders) + 1) if orders else 1


def _decode_image_bytes(data_b64: str, mime: str = "") -> tuple[bytes, str]:
    raw = data_b64.strip()
    if raw.startswith("data:"):
        header, _, payload = raw.partition(",")
        mime = header.split(";")[0].replace("data:", "").strip().lower()
        raw = payload
    try:
        data = base64.b64decode(raw, validate=True)
    except Exception as e:
        raise ValueError(f"invalid base64 image data: {e}") from e
    if not data:
        raise ValueError("empty image data")
    ext = "png"
    if "jpeg" in mime or "jpg" in mime:
        ext = "jpg"
    elif "webp" in mime:
        ext = "webp"
    elif "gif" in mime:
        ext = "gif"
    return data, ext


def save_notes(
    palace_id: str,
    notes: str,
    data_dir: Path,
    output_dir: Path,
    notes_root: Path | None,
) -> dict[str, Any]:
    palace_id = (palace_id or "").strip()
    if not palace_id:
        return {"ok": False, "error": "palace_id required"}
    data = load_all(data_dir)
    palace = _find_palace(data, palace_id)
    if not palace:
        return {"ok": False, "error": f"unknown palace_id {palace_id}"}
    study_id = palace.get("study_id") or ""
    new_palaces: list[dict[str, str]] = []
    for row in data["palaces"]:
        if row.get("id") == palace_id:
            updated = dict(row)
            updated["palace_notes"] = notes if notes is not None else ""
            new_palaces.append(updated)
        else:
            new_palaces.append(row)
    data["palaces"] = new_palaces
    save_all(data_dir, data)
    _sync_study(study_id, data_dir, output_dir, notes_root)
    return {"ok": True, "palace_id": palace_id, "study_id": study_id}


def _gallery_dest_path(
    output_dir: Path,
    slug: str,
    palace_number: str,
    image_id: str,
    ext: str,
) -> tuple[Path, str]:
    slug_part = slug_filename(slug)
    filename = f"{palace_number}-gallery-{image_id.replace('PALIMG_', '')}.{ext}"
    dest_dir = output_dir / "practice" / "images" / slug_part.replace("/", "\\")
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / filename
    rel = f"{PRACTICE_PREFIX}{slug_part}/{filename}"
    return dest, rel


def _delete_image_file(output_dir: Path, image_rel: str, all_rels: set[str]) -> None:
    if not image_rel or image_rel in all_rels:
        return
    norm = image_rel.replace("\\", "/")
    if norm.startswith(PRACTICE_PREFIX):
        path = output_dir / norm.replace("/", "\\")
        if not path.exists():
            path = output_dir / norm
        if path.exists() and path.is_file():
            try:
                path.unlink()
            except OSError:
                pass


def save_images(
    payload: dict[str, Any],
    data_dir: Path,
    output_dir: Path,
    notes_root: Path | None,
) -> dict[str, Any]:
    action = (payload.get("action") or "").strip().lower()
    palace_id = (payload.get("palace_id") or "").strip()
    if action not in ("add", "update", "delete"):
        return {"ok": False, "error": "action must be add, update, or delete"}
    if not palace_id:
        return {"ok": False, "error": "palace_id required"}

    data = load_all(data_dir)
    palace = _find_palace(data, palace_id)
    if not palace:
        return {"ok": False, "error": f"unknown palace_id {palace_id}"}
    study_id = palace.get("study_id") or ""
    slug = _study_slug(data, study_id)
    palace_number = str(palace.get("palace_number") or "").strip()
    images = list(data.get("palace_images") or [])

    if action == "add":
        data_b64 = payload.get("data_b64") or payload.get("image_data") or ""
        mime = (payload.get("mime") or payload.get("content_type") or "").strip().lower()
        if not data_b64:
            return {"ok": False, "error": "data_b64 required for add"}
        try:
            raw_bytes, ext = _decode_image_bytes(str(data_b64), mime)
        except ValueError as e:
            return {"ok": False, "error": str(e)}
        existing_ids = {r.get("id", "") for r in images}
        image_id = next_id("PALIMG_", existing_ids)
        dest, rel = _gallery_dest_path(output_dir, slug, palace_number, image_id, ext)
        dest.write_bytes(raw_bytes)
        caption = str(payload.get("caption") or "").strip()
        sort_order = str(
            payload.get("sort_order") or _next_sort_order(images, palace_id)
        )
        row = {
            "id": image_id,
            "palace_id": palace_id,
            "image_rel_path": rel,
            "caption": caption,
            "sort_order": sort_order,
        }
        images.append(row)
        data["palace_images"] = images
        save_all(data_dir, data)
        _sync_study(study_id, data_dir, output_dir, notes_root)
        return {
            "ok": True,
            "action": "add",
            "palace_id": palace_id,
            "study_id": study_id,
            "image": row,
        }

    image_id = (payload.get("image_id") or payload.get("id") or "").strip()
    if not image_id:
        return {"ok": False, "error": "image_id required for update/delete"}

    idx = next(
        (i for i, r in enumerate(images) if r.get("id") == image_id),
        -1,
    )
    if idx < 0 or images[idx].get("palace_id") != palace_id:
        return {"ok": False, "error": f"unknown image_id {image_id} for palace"}

    if action == "delete":
        removed = images.pop(idx)
        data["palace_images"] = images
        save_all(data_dir, data)
        all_rels = {
            (r.get("image_rel_path") or "").strip()
            for r in images
            if (r.get("image_rel_path") or "").strip()
        }
        hero = (palace.get("image_rel_path") or "").strip()
        if hero:
            all_rels.add(hero)
        _delete_image_file(output_dir, removed.get("image_rel_path") or "", all_rels)
        _sync_study(study_id, data_dir, output_dir, notes_root)
        return {
            "ok": True,
            "action": "delete",
            "palace_id": palace_id,
            "study_id": study_id,
            "image_id": image_id,
        }

    # update
    row = dict(images[idx])
    if "caption" in payload:
        row["caption"] = str(payload.get("caption") or "").strip()
    if "sort_order" in payload and str(payload.get("sort_order") or "").strip():
        row["sort_order"] = str(payload.get("sort_order")).strip()
    if payload.get("data_b64") or payload.get("image_data"):
        mime = (payload.get("mime") or payload.get("content_type") or "").strip().lower()
        try:
            raw_bytes, ext = _decode_image_bytes(
                str(payload.get("data_b64") or payload.get("image_data")), mime
            )
        except ValueError as e:
            return {"ok": False, "error": str(e)}
        old_rel = row.get("image_rel_path") or ""
        dest, rel = _gallery_dest_path(
            output_dir, slug, palace_number, image_id, ext
        )
        dest.write_bytes(raw_bytes)
        row["image_rel_path"] = rel
        all_rels = {
            (r.get("image_rel_path") or "").strip()
            for r in images
            if (r.get("image_rel_path") or "").strip()
        }
        hero = (palace.get("image_rel_path") or "").strip()
        if hero:
            all_rels.add(hero)
        if rel:
            all_rels.add(rel)
        _delete_image_file(output_dir, old_rel, all_rels)
    images[idx] = row
    data["palace_images"] = images
    save_all(data_dir, data)
    _sync_study(study_id, data_dir, output_dir, notes_root)
    return {
        "ok": True,
        "action": "update",
        "palace_id": palace_id,
        "study_id": study_id,
        "image": row,
    }
