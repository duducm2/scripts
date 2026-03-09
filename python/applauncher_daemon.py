#!/usr/bin/env python3
"""
AppLauncher persistent IPC daemon. Uses Memory-Mapped File (MMF) for shared memory
with Mutex and two EventWaitHandles for synchronization. AHK writes request into MMF,
signals request event; daemon processes and writes response, signals response event.
No RunWait from AHK per shortcut. Lifecycle: startup, health heartbeat, graceful shutdown.
"""

import sys
import os
import time
import signal
import threading
import ctypes
from ctypes import wintypes

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from al_protocol import (
    MMF_HEADER_SIZE,
    MMF_VERSION,
    STATUS_IDLE,
    STATUS_REQUEST_READY,
    STATUS_RESPONSE_READY,
    pack_header,
    unpack_header,
    encode_payload,
    decode_payload,
    make_response,
    validate_request,
    REQ_ID,
    REQ_OP,
    REQ_PAYLOAD,
    OP_HEALTH_CHECK,
    OP_PING,
    OP_ENUMERATE_WINDOWS,
    OP_RESOLVE_CURSOR_TARGETS,
    OP_WIKIPEDIA_STATE_READ,
    OP_WIKIPEDIA_SCROLL_RESTORE,
    OP_EVENT_PUSH,
)

# --- Constants ---
MMF_NAME = "Global\\AppLauncherIPC"
MMF_SIZE = 4096
MUTEX_NAME = "Global\\AppLauncherIPCLock"
EVENT_REQUEST_NAME = "Global\\AppLauncherRequestReady"
EVENT_RESPONSE_NAME = "Global\\AppLauncherResponseReady"
PAYLOAD_MAX = MMF_SIZE - MMF_HEADER_SIZE

INVALID_HANDLE_VALUE = 0xFFFFFFFF
PAGE_READWRITE = 0x04
FILE_MAP_ALL_ACCESS = 0xF001F
WAIT_OBJECT_0 = 0
WAIT_TIMEOUT = 258
WAIT_ABANDONED = 0x80
INFINITE = 0xFFFFFFFF
EVENT_MODIFY_STATE = 0x0002
MUTEX_MODIFY_STATE = 0x0001

_shutdown = threading.Event()


def _on_signal(signum, frame):
    _shutdown.set()


# --- Windows API via ctypes ---
kernel32 = ctypes.windll.kernel32  # noqa: E402

CreateFileMappingW = kernel32.CreateFileMappingW
CreateFileMappingW.argtypes = [
    wintypes.HANDLE,
    wintypes.LPVOID,
    wintypes.DWORD,
    wintypes.DWORD,
    wintypes.DWORD,
    wintypes.LPCWSTR,
]
CreateFileMappingW.restype = wintypes.HANDLE

OpenFileMappingW = kernel32.OpenFileMappingW
OpenFileMappingW.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.LPCWSTR]
OpenFileMappingW.restype = wintypes.HANDLE

MapViewOfFile = kernel32.MapViewOfFile
MapViewOfFile.argtypes = [
    wintypes.HANDLE,
    wintypes.DWORD,
    wintypes.DWORD,
    wintypes.DWORD,
    ctypes.c_size_t,
]
MapViewOfFile.restype = wintypes.LPVOID

UnmapViewOfFile = kernel32.UnmapViewOfFile
UnmapViewOfFile.argtypes = [wintypes.LPCVOID]
UnmapViewOfFile.restype = wintypes.BOOL

CloseHandle = kernel32.CloseHandle
CloseHandle.argtypes = [wintypes.HANDLE]
CloseHandle.restype = wintypes.BOOL

CreateMutexW = kernel32.CreateMutexW
CreateMutexW.argtypes = [wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR]
CreateMutexW.restype = wintypes.HANDLE

OpenMutexW = kernel32.OpenMutexW
OpenMutexW.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.LPCWSTR]
OpenMutexW.restype = wintypes.HANDLE

CreateEventW = kernel32.CreateEventW
CreateEventW.argtypes = [
    wintypes.LPVOID,
    wintypes.BOOL,
    wintypes.BOOL,
    wintypes.LPCWSTR,
]
CreateEventW.restype = wintypes.HANDLE

OpenEventW = kernel32.OpenEventW
OpenEventW.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.LPCWSTR]
OpenEventW.restype = wintypes.HANDLE

SetEvent = kernel32.SetEvent
SetEvent.argtypes = [wintypes.HANDLE]
SetEvent.restype = wintypes.BOOL

ResetEvent = kernel32.ResetEvent
ResetEvent.argtypes = [wintypes.HANDLE]
ResetEvent.restype = wintypes.BOOL

WaitForSingleObject = kernel32.WaitForSingleObject
WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
WaitForSingleObject.restype = wintypes.DWORD

ReleaseMutex = kernel32.ReleaseMutex
ReleaseMutex.argtypes = [wintypes.HANDLE]
ReleaseMutex.restype = wintypes.BOOL


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
        echo = dict(payload) if isinstance(payload, dict) else {}
        echo["serverTs"] = int(time.time() * 1000)
        return make_response(req_id, True, result=echo)

    if op == OP_ENUMERATE_WINDOWS:
        # Phase 2: implement EnumWindows in daemon; for now return empty
        return make_response(
            req_id, True, result={"hwnds": [], "primary": 0, "fallback": 0}
        )

    if op == OP_RESOLVE_CURSOR_TARGETS:
        try:
            from al_window_enum import resolve_cursor_targets

            return make_response(req_id, True, result=resolve_cursor_targets())
        except Exception as e:
            return make_response(
                req_id,
                False,
                error_code=3,
                error_message=str(e),
                result={"primaryHwnd": 0, "fallbackHwnd": 0, "reason": "error"},
            )

    if op == OP_WIKIPEDIA_STATE_READ:
        # Phase 4: UIA state read
        return make_response(req_id, True, result={"url": "", "scrollPercent": 0.0})

    if op == OP_WIKIPEDIA_SCROLL_RESTORE:
        # Phase 4: scroll restore
        return make_response(req_id, True, result={"ok": True})

    if op == OP_EVENT_PUSH:
        return make_response(req_id, True, result={"accepted": True})

    return make_response(req_id, False, error_code=2, error_message=f"unknown op: {op}")


def run_daemon() -> None:
    signal.signal(signal.SIGINT, _on_signal)
    signal.signal(signal.SIGTERM, _on_signal)

    # Create MMF (daemon creates; AHK opens with OpenFileMappingW)
    h_map = CreateFileMappingW(
        INVALID_HANDLE_VALUE,
        None,
        PAGE_READWRITE,
        0,
        MMF_SIZE,
        MMF_NAME,
    )
    if not h_map or h_map == INVALID_HANDLE_VALUE:
        err = ctypes.get_last_error()
        print(f"applauncher_daemon: CreateFileMappingW failed: {err}", file=sys.stderr)
        sys.exit(1)

    p_buf = MapViewOfFile(h_map, FILE_MAP_ALL_ACCESS, 0, 0, MMF_SIZE)
    if not p_buf:
        CloseHandle(h_map)
        print("applauncher_daemon: MapViewOfFile failed", file=sys.stderr)
        sys.exit(1)

    # Create Mutex and Events (daemon creates; AHK opens with Open*)
    h_mutex = CreateMutexW(None, False, MUTEX_NAME)
    if not h_mutex or h_mutex == INVALID_HANDLE_VALUE:
        UnmapViewOfFile(p_buf)
        CloseHandle(h_map)
        print("applauncher_daemon: CreateMutexW failed", file=sys.stderr)
        sys.exit(1)

    h_evt_request = CreateEventW(
        None, True, False, EVENT_REQUEST_NAME
    )  # manual reset so AHK can signal
    h_evt_response = CreateEventW(None, True, False, EVENT_RESPONSE_NAME)
    if not h_evt_request or not h_evt_response:
        CloseHandle(h_mutex)
        UnmapViewOfFile(p_buf)
        CloseHandle(h_map)
        print("applauncher_daemon: CreateEventW failed", file=sys.stderr)
        sys.exit(1)

    seq = 0
    try:
        while not _shutdown.is_set():
            # Wait for AHK to write request (request event signaled)
            r = WaitForSingleObject(h_evt_request, 500)
            if r == WAIT_TIMEOUT:
                continue
            if r not in (WAIT_OBJECT_0, WAIT_ABANDONED):
                break

            if WaitForSingleObject(h_mutex, 5000) != WAIT_OBJECT_0:
                ResetEvent(h_evt_request)
                continue

            try:
                # Read header
                raw = (ctypes.c_char * MMF_SIZE).from_address(p_buf)
                header = unpack_header(bytes(raw[:MMF_HEADER_SIZE]))
                if not header:
                    ReleaseMutex(h_mutex)
                    ResetEvent(h_evt_request)
                    continue
                version, seq_val, payload_len, status, _ = header
                if status != STATUS_REQUEST_READY or payload_len > PAYLOAD_MAX:
                    ReleaseMutex(h_mutex)
                    ResetEvent(h_evt_request)
                    continue
                req = decode_payload(
                    bytes(raw[MMF_HEADER_SIZE : MMF_HEADER_SIZE + payload_len])
                )
                if not req:
                    ReleaseMutex(h_mutex)
                    ResetEvent(h_evt_request)
                    continue
                resp = handle_request(req)
            finally:
                ResetEvent(h_evt_request)
                ReleaseMutex(h_mutex)

            # Write response
            if WaitForSingleObject(h_mutex, 5000) != WAIT_OBJECT_0:
                continue
            try:
                resp_bytes = encode_payload(resp)
                if len(resp_bytes) > PAYLOAD_MAX:
                    resp_bytes = resp_bytes[:PAYLOAD_MAX]
                seq += 1
                header_bytes = pack_header(
                    MMF_VERSION, seq, len(resp_bytes), STATUS_RESPONSE_READY, 0
                )
                raw = (ctypes.c_char * MMF_SIZE).from_address(p_buf)
                for i, b in enumerate(header_bytes):
                    raw[i] = b
                for i, b in enumerate(resp_bytes):
                    raw[MMF_HEADER_SIZE + i] = b
                SetEvent(h_evt_response)
            finally:
                ReleaseMutex(h_mutex)
    finally:
        CloseHandle(h_evt_response)
        CloseHandle(h_evt_request)
        CloseHandle(h_mutex)
        UnmapViewOfFile(p_buf)
        CloseHandle(h_map)


if __name__ == "__main__":
    run_daemon()
