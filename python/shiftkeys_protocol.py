# ShiftKeys IPC protocol: message framing and validation.
# Used by the persistent Python daemon and the AHK client.
# Frame format: length (4 bytes, big-endian) + UTF-8 JSON payload.
# Envelope: id, op, context, payload, ts, deadlineMs (request);
#           id, ok, result, errorCode, errorMessage, telemetry (response).

import json
import struct
import time

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
RESP_ERROR_CODE = "errorCode"
RESP_ERROR_MESSAGE = "errorMessage"
RESP_TELEMETRY = "telemetry"

# Operations
OP_HEALTH_CHECK = "HealthCheck"
OP_WATCH_UI_STATE = "WatchUIState"
OP_GET_WATCH_STATUS = "GetWatchStatus"
OP_RESOLVE_CONTEXT = "ResolveContext"
OP_FIND_ELEMENT = "FindElement"
OP_BULK_FIND = "BulkFind"
OP_WAIT_ELEMENT_STATE = "WaitElementState"

REQUIRED_REQUEST_KEYS = (REQ_ID, REQ_OP)


def encode_message(obj: dict) -> bytes:
    payload = json.dumps(obj, ensure_ascii=False).encode("utf-8")
    return struct.pack(">I", len(payload)) + payload


def decode_message(data: bytes) -> dict | None:
    if len(data) < 4:
        return None
    (length,) = struct.unpack(">I", data[:4])
    if length > 1024 * 1024:
        return None
    if len(data) < 4 + length:
        return None
    payload = data[4 : 4 + length].decode("utf-8")
    try:
        return json.loads(payload)
    except json.JSONDecodeError:
        return None


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
    error_code: int = 0,
    error_message: str = "",
    telemetry: dict | None = None,
) -> dict:
    out = {
        RESP_ID: req_id,
        RESP_OK: ok,
        RESP_RESULT: result if result is not None else {},
        RESP_ERROR_CODE: error_code,
        RESP_ERROR_MESSAGE: error_message or "",
    }
    if telemetry:
        out[RESP_TELEMETRY] = telemetry
    return out


def validate_request(obj: dict) -> bool:
    if not isinstance(obj, dict):
        return False
    return all(k in obj for k in REQUIRED_REQUEST_KEYS)
