"""Tests for context_compact: CSV keep+header, JSON strip, markdown collapse."""

from __future__ import annotations

import json
from pathlib import Path

from context_compact import (
    OMITTED_DATA_URI,
    OMITTED_URL,
    compact_csv_text,
    compact_json_text,
    compact_text,
    csv_keep_range,
    transform_file,
    transform_text,
)


def test_csv_keep_header_plus_range():
    text = "h1,h2\na,1\nb,2\nc,3\nd,4\ne,5\n"
    out = csv_keep_range(text, 3, 5)
    assert out.split("\n") == ["h1,h2", "b,2", "c,3", "d,4"]


def test_csv_keep_from_one_does_not_duplicate_header():
    text = "h1,h2\na,1\nb,2\nc,3\n"
    out = csv_keep_range(text, 1, 2)
    assert out.split("\n") == ["h1,h2", "a,1"]


def test_json_drops_nulls_and_omits_data_uri_and_long_urls():
    long_url = "https://example.com/" + ("x" * 200)
    src = {
        "keep": "ok",
        "drop": None,
        "nested": {"a": 1, "b": None},
        "list": [1, None, 2],
        "img": "data:image/svg+xml,%3Csvg%3E",
        "preview": long_url,
        "short": "https://ok.example/x",
    }
    out = json.loads(compact_json_text(json.dumps(src)))
    assert out["keep"] == "ok"
    assert "drop" not in out
    assert out["nested"] == {"a": 1}
    assert out["list"] == [1, 2]
    assert out["img"] == OMITTED_DATA_URI
    assert out["preview"] == OMITTED_URL
    assert out["short"] == "https://ok.example/x"
    assert "\n" not in compact_json_text(json.dumps(src))


def test_markdown_collapses_three_plus_blank_lines_to_one():
    text = "a\n\n\n\nb\n\nc"
    out = compact_text(text)
    assert out == "a\n\nb\n\nc"


def test_csv_compact_drops_empty_rows_keeps_header():
    text = "h1,h2\n\na,1\n  \nb,2\n"
    out = compact_csv_text(text)
    assert out.split("\n") == ["h1,h2", "a,1", "b,2"]


def test_transform_csv_keep_then_compact(tmp_path: Path):
    src = tmp_path / "data.csv"
    dst = tmp_path / "data.compacted.csv"
    src.write_text("h1,h2\na,1\n\nb,2\nc,3\nd,4\n", encoding="utf-8")
    transform_file(src, dst, compact=True, csv_keep=(4, 6))
    # header + lines 4-6, then empty rows dropped by compact
    assert dst.read_text(encoding="utf-8").split("\n") == ["h1,h2", "b,2", "c,3", "d,4"]


def test_transform_text_json_without_compact_is_unchanged():
    raw = '{\n  "a": null\n}\n'
    assert transform_text(raw, "json", compact=False, csv_keep=None) == raw
