# WindowManagement IPC protocol: message framing and validation.
# Used by the persistent Python daemon and the AHK client.
# Frame format: length (4 bytes, big-endian) + UTF-8 JSON payload.
# Envelope: id, op, context, payload, ts, deadlineMs (request);
#           id, ok, result, errorCode, errorMessage, telemetry (response).

import struct
import time

from ipc_wire import json_dumps, json_loads_dict, validate_ipc_request_envelope

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
OP_PING = "Ping"
# Phase 2: cache API
OP_GET_FOREGROUND_WINDOW_STATE = "GetForegroundWindowState"
OP_GET_CURSOR_WINDOWS = "GetCursorWindows"
OP_GET_PREVIEW_WINDOWS = "GetPreviewWindows"
OP_GET_VISIBLE_WINDOWS_BY_MONITOR = "GetVisibleWindowsByMonitor"
OP_RESOLVE_PROJECT_WINDOW = "ResolveProjectWindow"
OP_GET_AUTOMATION_CONTEXT = "GetAutomationContext"
OP_BEGIN_AUTOMATION_SWITCH = "BeginAutomationSwitch"
OP_END_AUTOMATION_SWITCH = "EndAutomationSwitch"

REQUIRED_REQUEST_KEYS = (REQ_ID, REQ_OP)


def encode_message(obj: dict) -> bytes:
    payload = json_dumps(obj)
    return struct.pack(">I", len(payload)) + payload


def decode_message(data: bytes) -> dict | None:
    if len(data) < 4:
        return None
    (length,) = struct.unpack(">I", data[:4])
    if length > 1024 * 1024:
        return None
    if len(data) < 4 + length:
        return None
    return json_loads_dict(memoryview(data)[4 : 4 + length])


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
    return validate_ipc_request_envelope(obj)
