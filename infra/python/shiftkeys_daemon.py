#!/usr/bin/env python3
"""
ShiftKeys persistent IPC daemon. Listens on Named Pipe \\\\.\\pipe\\shiftkeys_automation
for length-prefixed JSON frames. AHK connects once; daemon handles HealthCheck,
WatchUIState, ResolveContext, FindElement, BulkFind. No RunWait from AHK per shortcut.
"""

import sys
import struct
import time
import os
import threading
import uuid

# Add script dir for imports
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from daemon_logging import configure_daemon_logging, get_logger
from shiftkeys_protocol import (
    encode_message,
    decode_message,
    validate_request,
    make_response,
    REQ_ID,
    REQ_OP,
    REQ_PAYLOAD,
    OP_HEALTH_CHECK,
    OP_WATCH_UI_STATE,
    OP_GET_WATCH_STATUS,
    OP_RESOLVE_CONTEXT,
    OP_FIND_ELEMENT,
    OP_BULK_FIND,
    OP_WAIT_ELEMENT_STATE,
)

try:
    from shiftkeys_context import start_context_hook, get_context_snapshot
except ImportError:
    start_context_hook = lambda: None
    get_context_snapshot = lambda: {}
try:
    from shiftkeys_uia import (
        find_element as uia_find_element,
        wait_element_state as uia_wait_element_state,
    )
except ImportError:
    uia_find_element = lambda h, n, t=50000: False
    uia_wait_element_state = lambda h, n, to, p=300: "timeout"

PIPE_NAME = r"\\.\pipe\shiftkeys_automation"
MAX_FRAME = 1024 * 1024

# Phase 3: watch state (watchId -> {status, ts}); real UIA polling in Phase 4
_watches_lock = threading.Lock()
_watches: dict = {}


def _watch_worker(watch_id: str, timeout_ms: int) -> None:
    time.sleep(min(timeout_ms / 1000.0, 3600))  # cap 1h
    with _watches_lock:
        if watch_id in _watches and _watches[watch_id].get("status") == "pending":
            _watches[watch_id] = {"status": "timeout", "ts": time.time()}


def _watch_start(context: str, timeout_ms: int) -> str:
    watch_id = str(uuid.uuid4())
    with _watches_lock:
        _watches[watch_id] = {"status": "pending", "ts": time.time()}
    t = threading.Thread(target=_watch_worker, args=(watch_id, timeout_ms), daemon=True)
    t.start()
    return watch_id


def _watch_status(watch_id: str) -> dict:
    with _watches_lock:
        return dict(_watches.get(watch_id, {"status": "unknown"}))


try:
    import win32pipe
    import win32file
    import pywintypes

    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False


def read_frame(handle):
    """Read one length-prefixed frame from pipe. Returns decoded dict or None on EOF/error."""
    err, header = win32file.ReadFile(handle, 4)
    if err or len(header) < 4:
        return None
    (length,) = struct.unpack(">I", header)
    if length > MAX_FRAME or length == 0:
        return None
    err, payload = win32file.ReadFile(handle, length)
    if err or len(payload) < length:
        return None
    return decode_message(header + payload)


def write_frame(handle, obj):
    data = encode_message(obj)
    win32file.WriteFile(handle, data)


def handle_request(req: dict) -> dict:
    """Dispatch by op; return response envelope."""
    if not validate_request(req):
        return make_response(
            req.get(REQ_ID, ""), False, error_code=1, error_message="invalid request"
        )
    req_id = req[REQ_ID]
    op = req.get(REQ_OP, "")
    payload = req.get(REQ_PAYLOAD, {})

    if op == OP_HEALTH_CHECK:
        return make_response(
            req_id, True, result={"alive": True, "ts": int(time.time() * 1000)}
        )

    if op == OP_RESOLVE_CONTEXT:
        return make_response(req_id, True, result=get_context_snapshot())

    if op == OP_FIND_ELEMENT:
        hwnd = int(payload.get("hwnd", 0))
        name = str(payload.get("name", ""))
        ctype = int(payload.get("controlTypeId", 50000))
        found = bool(hwnd and name and uia_find_element(hwnd, name, ctype))
        return make_response(req_id, True, result={"found": found, "element": None})

    if op == OP_BULK_FIND:
        return make_response(req_id, True, result={"elements": []})

    if op == OP_WAIT_ELEMENT_STATE:
        hwnd = int(payload.get("hwnd", 0))
        name = str(payload.get("name", ""))
        timeout_ms = int(payload.get("timeoutMs", 5000))
        poll_ms = int(payload.get("pollMs", 300))
        status = (
            uia_wait_element_state(hwnd, name, timeout_ms, poll_ms)
            if (hwnd and name)
            else "timeout"
        )
        return make_response(req_id, True, result={"status": status})

    if op == OP_WATCH_UI_STATE:
        timeout_ms = int(payload.get("timeoutMs", 300000))
        context = str(payload.get("context", ""))
        watch_id = _watch_start(context, timeout_ms)
        return make_response(
            req_id, True, result={"watchId": watch_id, "status": "pending"}
        )

    if op == OP_GET_WATCH_STATUS:
        watch_id = str(payload.get("watchId", ""))
        if not watch_id:
            return make_response(
                req_id, False, error_code=3, error_message="missing watchId"
            )
        st = _watch_status(watch_id)
        return make_response(req_id, True, result=st)

    return make_response(req_id, False, error_code=2, error_message=f"unknown op: {op}")


def serve_connection(handle) -> None:
    log = get_logger()
    try:
        while True:
            msg = read_frame(handle)
            if msg is None:
                break
            if not isinstance(msg, dict):
                break
            rid = str(msg.get(REQ_ID, ""))
            op = str(msg.get(REQ_OP, ""))
            t0 = time.perf_counter()
            try:
                resp = handle_request(msg)
                ms = (time.perf_counter() - t0) * 1000.0
                log.info(
                    "ipc_request",
                    req_id=rid,
                    op=op,
                    ok=bool(resp.get("ok", False)),
                    ms_ms=round(ms, 3),
                )
            except Exception as e:
                log.exception("ipc_request_error", req_id=rid, op=op)
                resp = make_response(
                    rid,
                    False,
                    error_code=99,
                    error_message=str(e)[:500],
                )
            write_frame(handle, resp)
    except (BrokenPipeError, OSError):
        pass
    finally:
        try:
            win32file.CloseHandle(handle)
        except OSError:
            pass


def run_pipe_server() -> None:
    if not HAS_WIN32:
        print(
            "shiftkeys_daemon: pywin32 required for Named Pipe. Install: pip install pywin32",
            file=sys.stderr,
        )
        sys.exit(1)
    while True:
        try:
            pipe = win32pipe.CreateNamedPipe(
                PIPE_NAME,
                win32pipe.PIPE_ACCESS_DUPLEX,
                win32pipe.PIPE_TYPE_BYTE
                | win32pipe.PIPE_READMODE_BYTE
                | win32pipe.PIPE_WAIT,
                1,
                65536,
                65536,
                0,
                None,
            )
        except pywintypes.error as e:
            print(f"shiftkeys_daemon: CreateNamedPipe failed: {e}", file=sys.stderr)
            sys.exit(1)
        try:
            win32pipe.ConnectNamedPipe(pipe, None)
        except pywintypes.error as e:
            win32file.CloseHandle(pipe)
            if e.winerror == 535:  # ERROR_PIPE_CONNECTED
                continue
            time.sleep(0.5)
            continue
        serve_connection(pipe)


if __name__ == "__main__":
    configure_daemon_logging("shiftkeys_daemon")
    start_context_hook()
    run_pipe_server()
