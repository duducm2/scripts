# Gemini IPC protocol: message framing and validation.
# Used by the persistent Python daemon and the AHK client.
# Frame format: length (4 bytes, big-endian) + UTF-8 JSON payload.

import json
import struct
import time


def encode_message(obj: dict) -> bytes:
    payload = json.dumps(obj, ensure_ascii=False).encode("utf-8")
    return struct.pack(">I", len(payload)) + payload


def decode_message(data: bytes) -> dict | None:
    if len(data) < 4:
        return None
    (length,) = struct.unpack(">I", data[:4])
    if len(data) < 4 + length:
        return None
    payload = data[4 : 4 + length].decode("utf-8")
    try:
        return json.loads(payload)
    except json.JSONDecodeError:
        return None


def read_frame(stream):
    """Read one length-prefixed frame from stream. Returns decoded dict or None on EOF/error."""
    header = stream.buffer.read(4)
    if len(header) < 4:
        return None
    (length,) = struct.unpack(">I", header)
    if length > 1024 * 1024:
        return None
    payload = stream.buffer.read(length)
    if len(payload) < length:
        return None
    try:
        return json.loads(payload.decode("utf-8"))
    except (json.JSONDecodeError, UnicodeDecodeError):
        return None


def write_frame(stream, obj: dict) -> None:
    stream.buffer.write(encode_message(obj))
    stream.buffer.flush()


# Request envelope keys
REQ_ID = "id"
REQ_OP = "op"
REQ_CONTEXT = "context"
REQ_PAYLOAD = "payload"
REQ_TS = "ts"
REQ_DEADLINE_MS = "deadlineMs"

# Response envelope keys
RESP_ID = "id"
RESP_OK = "ok"
RESP_RESULT = "result"
RESP_ERROR = "error"

# Operations
OP_HEALTH_CHECK = "HealthCheck"
OP_PING = "Ping"
OP_QUEUE_TASK = "QueueTask"
OP_GET_TASK_STATUS = "GetTaskStatus"

REQUIRED_KEYS = (REQ_ID, REQ_OP)


def make_request(
    req_id: str,
    op: str,
    context: str = "",
    payload: dict | None = None,
    deadline_ms: int = 0,
) -> dict:
    return {
        REQ_ID: req_id,
        REQ_OP: op,
        REQ_CONTEXT: context or "",
        REQ_PAYLOAD: payload if payload is not None else {},
        REQ_TS: int(time.time() * 1000),
        REQ_DEADLINE_MS: deadline_ms,
    }


def make_response(
    req_id: str,
    ok: bool,
    result=None,
    error: str = "",
) -> dict:
    return {
        RESP_ID: req_id,
        RESP_OK: ok,
        RESP_RESULT: result if result is not None else {},
        RESP_ERROR: error or "",
    }


def validate_request(obj: dict) -> bool:
    if not isinstance(obj, dict):
        return False
    return all(k in obj for k in REQUIRED_KEYS)
