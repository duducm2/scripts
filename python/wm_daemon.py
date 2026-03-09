#!/usr/bin/env python3
"""
WindowManagement persistent IPC daemon. Listens on Named Pipe \\\\.\\pipe\\wm_automation
for length-prefixed JSON frames. AHK connects once; daemon handles HealthCheck and Ping
(Phase 1). No RunWait from AHK per shortcut. Lifecycle: startup, heartbeat-ready idle
loop, graceful shutdown.
"""

import sys
import struct
import time
import os
import signal
import threading

# Add script dir for imports
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from wm_protocol import (
    encode_message,
    decode_message,
    validate_request,
    make_response,
    REQ_ID,
    REQ_OP,
    REQ_PAYLOAD,
    OP_HEALTH_CHECK,
    OP_PING,
    OP_GET_FOREGROUND_WINDOW_STATE,
    OP_GET_CURSOR_WINDOWS,
    OP_GET_PREVIEW_WINDOWS,
    OP_GET_VISIBLE_WINDOWS_BY_MONITOR,
    OP_RESOLVE_PROJECT_WINDOW,
)

PIPE_NAME = r"\\.\pipe\wm_automation"
MAX_FRAME = 1024 * 1024

_shutdown = threading.Event()


def _on_signal(signum, frame):
    _shutdown.set()


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

    if op == OP_PING:
        # Echo payload for latency harness; add server ts.
        echo = dict(payload) if isinstance(payload, dict) else {}
        echo["serverTs"] = int(time.time() * 1000)
        return make_response(req_id, True, result=echo)

    # Phase 2: cache API (requires wm_hooks)
    try:
        import wm_hooks as hooks
    except ImportError:
        hooks = None

    if op == OP_GET_FOREGROUND_WINDOW_STATE:
        if hooks:
            return make_response(
                req_id, True, result=hooks.get_foreground_window_state()
            )
        return make_response(
            req_id,
            True,
            result={"hwnd": 0, "pid": 0, "title": "", "class": "", "exe": ""},
        )

    if op == OP_GET_CURSOR_WINDOWS:
        if hooks:
            return make_response(
                req_id, True, result={"windows": hooks.get_cursor_windows()}
            )
        return make_response(req_id, True, result={"windows": []})

    if op == OP_GET_PREVIEW_WINDOWS:
        if hooks:
            return make_response(
                req_id, True, result={"windows": hooks.get_preview_windows()}
            )
        return make_response(req_id, True, result={"windows": []})

    if op == OP_GET_VISIBLE_WINDOWS_BY_MONITOR:
        mon = int(payload.get("monitorIndex", 0))
        if hooks:
            return make_response(
                req_id,
                True,
                result={"windows": hooks.get_visible_windows_by_monitor(mon)},
            )
        return make_response(req_id, True, result={"windows": []})

    if op == OP_RESOLVE_PROJECT_WINDOW:
        project_path = str(payload.get("projectPath", ""))
        if hooks:
            return make_response(
                req_id, True, result=hooks.resolve_project_window(project_path)
            )
        return make_response(req_id, True, result={"hwnd": 0, "title": ""})

    return make_response(req_id, False, error_code=2, error_message=f"unknown op: {op}")


def serve_connection(handle) -> None:
    try:
        while not _shutdown.is_set():
            msg = read_frame(handle)
            if msg is None:
                break
            resp = handle_request(msg)
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
            "wm_daemon: pywin32 required for Named Pipe. Install: pip install pywin32",
            file=sys.stderr,
        )
        sys.exit(1)
    signal.signal(signal.SIGINT, _on_signal)
    signal.signal(signal.SIGTERM, _on_signal)
    while not _shutdown.is_set():
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
            print(f"wm_daemon: CreateNamedPipe failed: {e}", file=sys.stderr)
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
    try:
        import wm_hooks

        wm_hooks.start_wm_hooks()
    except ImportError:
        pass
    try:
        run_pipe_server()
    finally:
        try:
            import wm_hooks

            wm_hooks.stop_wm_hooks()
        except ImportError:
            pass
