"""Shared UTF-8 JSON and minimal request validation for daemon IPC (orjson + pydantic)."""

from __future__ import annotations

from typing import Any

import orjson
from pydantic import BaseModel, ConfigDict, ValidationError, field_validator


def json_dumps(obj: dict[str, Any]) -> bytes:
    """Serialize a dict to compact UTF-8 JSON (non-ASCII as UTF-8, not \\u escapes)."""
    return orjson.dumps(obj)


def json_loads_dict(raw: bytes | bytearray | memoryview) -> dict[str, Any] | None:
    """Parse JSON bytes; return a dict or None if invalid or root is not an object."""
    try:
        val = orjson.loads(raw)
    except (orjson.JSONDecodeError, TypeError, ValueError):
        return None
    return val if isinstance(val, dict) else None


class IpcRequestEnvelope(BaseModel):
    """Wire requests must include id and op (AHK may send numeric ids as JSON numbers)."""

    model_config = ConfigDict(extra="ignore")

    id: str
    op: str

    @field_validator("id", "op", mode="before")
    @classmethod
    def _coerce_id_op(cls, v: object) -> str:
        if v is None:
            raise ValueError("missing")
        return str(v)


def validate_ipc_request_envelope(obj: object) -> bool:
    if not isinstance(obj, dict):
        return False
    try:
        IpcRequestEnvelope.model_validate(obj)
        return True
    except ValidationError:
        return False
