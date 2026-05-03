#!/usr/bin/env python3
"""
Persistent Gemini IPC daemon. Listens on Named Pipe \\\\.\\pipe\\gemini_automation
for length-prefixed JSON frames. The daemon owns a lightweight background task
queue so AHK hotkeys can enqueue Gemini work and return immediately.
"""

import struct
import sys
import time
import uuid
import signal
import threading

from daemon_logging import configure_daemon_logging, get_logger
from protocol import (
    encode_message,
    decode_message,
    validate_request,
    make_response,
    REQ_ID,
    REQ_OP,
    REQ_PAYLOAD,
    OP_HEALTH_CHECK,
    OP_PING,
    OP_QUEUE_TASK,
    OP_GET_TASK_STATUS,
    OP_DETECT_LANG,
)

PIPE_NAME = r"\\.\pipe\gemini_automation"
MAX_FRAME = 1024 * 1024
DEFAULT_READY_DELAY_MS = 60
_shutdown = threading.Event()
_tasks: dict[str, dict] = {}
_tasks_lock = threading.Lock()
_lang_detector = None
_lang_detector_lock = threading.Lock()

try:
    import win32pipe
    import win32file
    import pywintypes

    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False


def _on_signal(signum, frame):
    _shutdown.set()


def _now_ms() -> int:
    return int(time.time() * 1000)


def _get_lang_detector():
    """Lazy singleton: Portuguese / English / German only (short-text friendly)."""
    global _lang_detector
    if _lang_detector is not None:
        return _lang_detector
    with _lang_detector_lock:
        if _lang_detector is not None:
            return _lang_detector
        from lingua import Language, LanguageDetectorBuilder

        det = (
            LanguageDetectorBuilder.from_languages(
                Language.PORTUGUESE,
                Language.ENGLISH,
                Language.GERMAN,
            )
            .with_preloaded_language_models()
            .build()
        )
        _lang_detector = det
        return _lang_detector


def _prune_finished_tasks() -> None:
    cutoff = _now_ms() - 5 * 60 * 1000
    with _tasks_lock:
        stale = [
            task_id
            for task_id, task in _tasks.items()
            if int(task.get("createdAt", 0)) < cutoff
        ]
        for task_id in stale:
            _tasks.pop(task_id, None)


def handle_request(req: dict) -> dict:
    """Dispatch by op; return response envelope."""
    if not validate_request(req):
        return make_response(req.get(REQ_ID, ""), False, error="invalid request")
    req_id = req[REQ_ID]
    op = req.get(REQ_OP, "")
    payload = req.get(REQ_PAYLOAD, {})

    if op == OP_HEALTH_CHECK:
        return make_response(req_id, True, result={"alive": True, "ts": _now_ms()})

    if op == OP_PING:
        echo = dict(payload) if isinstance(payload, dict) else {}
        echo["pong"] = True
        echo["serverTs"] = _now_ms()
        return make_response(req_id, True, result=echo)

    if op == OP_QUEUE_TASK:
        task_kind = str(payload.get("taskKind", "")).strip()
        if not task_kind:
            return make_response(req_id, False, error="missing taskKind")
        ready_delay_ms = max(
            0, int(payload.get("readyDelayMs", DEFAULT_READY_DELAY_MS))
        )
        created_at = _now_ms()
        task_id = uuid.uuid4().hex
        task = {
            "taskId": task_id,
            "taskKind": task_kind,
            "status": "pending",
            "createdAt": created_at,
            "readyAt": created_at + ready_delay_ms,
            "payload": dict(payload) if isinstance(payload, dict) else {},
        }
        with _tasks_lock:
            _tasks[task_id] = task
        return make_response(
            req_id,
            True,
            result={
                "taskId": task_id,
                "taskKind": task_kind,
                "status": "pending",
                "createdAt": created_at,
                "readyAt": task["readyAt"],
            },
        )

    if op == OP_GET_TASK_STATUS:
        task_id = str(payload.get("taskId", "")).strip()
        if not task_id:
            return make_response(req_id, False, error="missing taskId")
        with _tasks_lock:
            task = _tasks.get(task_id)
            if not task:
                return make_response(req_id, False, error="task not found")
            if task["status"] == "pending" and _now_ms() >= int(task.get("readyAt", 0)):
                task["status"] = "ready"
            snapshot = {
                "taskId": task["taskId"],
                "taskKind": task.get("taskKind", ""),
                "status": task.get("status", "pending"),
                "createdAt": int(task.get("createdAt", 0)),
                "readyAt": int(task.get("readyAt", 0)),
                "payload": task.get("payload", {}),
            }
        return make_response(req_id, True, result=snapshot)

    if op == OP_DETECT_LANG:
        text = str(payload.get("text", "")).strip()
        if not text:
            return make_response(
                req_id,
                True,
                result={"language": "en", "confidence": 0.0, "fallback": True},
            )
        try:
            detector = _get_lang_detector()
        except Exception as e:
            return make_response(req_id, False, error=f"detector unavailable: {e}")
        result = detector.detect_language_of(text)
        iso = result.iso_code_639_1.name.lower() if result is not None else "en"
        return make_response(
            req_id,
            True,
            result={"language": iso, "fallback": result is None},
        )

    return make_response(req_id, False, error=f"unknown op: {op}")


def read_frame(handle):
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


def serve_connection(handle) -> None:
    log = get_logger()
    try:
        while not _shutdown.is_set():
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
                resp = make_response(rid, False, error=str(e)[:500])
            write_frame(handle, resp)
            _prune_finished_tasks()
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
            "gemini_daemon: pywin32 required for Named Pipe. Install: pip install pywin32",
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
            print(f"gemini_daemon: CreateNamedPipe failed: {e}", file=sys.stderr)
            sys.exit(1)
        try:
            win32pipe.ConnectNamedPipe(pipe, None)
        except pywintypes.error as e:
            win32file.CloseHandle(pipe)
            if e.winerror == 535:
                continue
            time.sleep(0.5)
            continue
        serve_connection(pipe)


if __name__ == "__main__":
    configure_daemon_logging("gemini_daemon")
    run_pipe_server()
