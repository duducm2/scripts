"""Tests for prompt_usage_summary.py."""

from __future__ import annotations

import configparser
from datetime import datetime, timedelta
from pathlib import Path

from infra.python.prompt_usage_summary import load_prompt_chars, parse_log


def test_load_prompt_chars(tmp_path: Path) -> None:
    ini = tmp_path / "prompts.ini"
    ini.write_text(
        "[Prompt_1]\nChar=a\nName=Alpha\n[Prompt_2]\nChar=b\nName=Beta\n",
        encoding="utf-8",
    )
    chars = load_prompt_chars(ini)
    assert chars == {"a": "Alpha", "b": "Beta"}


def test_parse_log_filters_by_days(tmp_path: Path) -> None:
    log = tmp_path / "prompt_usage.log"
    old = (datetime.now() - timedelta(days=10)).strftime("%Y-%m-%d %H:%M:%S")
    recent = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d %H:%M:%S")
    log.write_text(
        f"{old}|a|Alpha\n{recent}|b|Beta\n",
        encoding="utf-8",
    )
    since = datetime.now() - timedelta(days=7)
    rows = parse_log(log, since)
    assert len(rows) == 1
    assert rows[0]["char"] == "b"
    assert rows[0]["name"] == "Beta"
