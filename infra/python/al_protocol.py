# AppLauncher IPC protocol: MMF message schema and operations.
# Used by applauncher_daemon.py and AppLauncherIPC.ahk.
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
OP_ENUMERATE_WINDOWS = "EnumerateWindows"
OP_RESOLVE_CURSOR_TARGETS = "ResolveCursorTargets"
OP_WIKIPEDIA_STATE_READ = "WikipediaStateRead"
OP_WIKIPEDIA_SCROLL_RESTORE = "WikipediaScrollRestore"
OP_EVENT_PUSH = "EventPush"

REQUIRED_REQUEST_KEYS = (REQ_ID, REQ_OP)

# MMF header layout: version(4), seq(4), payloadLen(4), status(4), errorCode(4) = 20 bytes
MMF_HEADER_SIZE = 20
MMF_VERSION = 1
STATUS_IDLE = 0
STATUS_REQUEST_READY = 1  # AHK wrote request
STATUS_RESPONSE_READY = 2  # Daemon wrote response


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


def encode_payload(obj: dict) -> bytes:
    return json_dumps(obj)


def decode_payload(data: bytes) -> dict | None:
    if not data:
        return None
    return json_loads_dict(data)


def pack_header(
    version: int, seq: int, payload_len: int, status: int, error_code: int
) -> bytes:
    return struct.pack("<IIIII", version, seq, payload_len, status, error_code)


def unpack_header(data: bytes) -> tuple[int, int, int, int, int] | None:
    if len(data) < MMF_HEADER_SIZE:
        return None
    return struct.unpack("<IIIII", data[:MMF_HEADER_SIZE])
