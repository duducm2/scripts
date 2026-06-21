# Gemini IPC protocol: message framing and validation.
# Used by the persistent Python daemon and the AHK client.
# Frame format: length (4 bytes, big-endian) + UTF-8 JSON payload.

import struct
import time

from ipc_wire import json_dumps, json_loads_dict, validate_ipc_request_envelope


def encode_message(obj: dict) -> bytes:
    payload = json_dumps(obj)
    return struct.pack(">I", len(payload)) + payload


def decode_message(data: bytes) -> dict | None:
    if len(data) < 4:
        return None
    (length,) = struct.unpack(">I", data[:4])
    if len(data) < 4 + length:
        return None
    return json_loads_dict(memoryview(data)[4 : 4 + length])


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
    return json_loads_dict(payload)


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
OP_DETECT_LANG = "DetectLang"

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
    return validate_ipc_request_envelope(obj)
