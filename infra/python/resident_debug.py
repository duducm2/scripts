"""Memory and slow-request log for the resident web servers.

Palace, Tasks, and Finance stay running so they open without a cold start.
This writes one local file per server under %TEMP% (not the Google Drive
repo). A line every 30s records working set and private bytes. A request
line is written only when that request took at least 50ms, and /health is
logged only when it took at least 200ms.
"""

from __future__ import annotations

import ctypes
import os
import tempfile
import threading
import time
from ctypes import wintypes
from pathlib import Path
from typing import Any

_SLOW_MS = 50
_SLOW_HEALTH_MS = 200
_SAMPLE_SEC = 30
_MAX_BYTES = 400_000
_LOCK = threading.Lock()
_STARTED: set[str] = set()


class _MemoryCounters(ctypes.Structure):
    _fields_ = [
        ("cb", wintypes.DWORD),
        ("PageFaultCount", wintypes.DWORD),
        ("PeakWorkingSetSize", ctypes.c_size_t),
        ("WorkingSetSize", ctypes.c_size_t),
        ("QuotaPeakPagedPoolUsage", ctypes.c_size_t),
        ("QuotaPagedPoolUsage", ctypes.c_size_t),
        ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t),
        ("QuotaNonPagedPoolUsage", ctypes.c_size_t),
        ("PagefileUsage", ctypes.c_size_t),
        ("PeakPagefileUsage", ctypes.c_size_t),
        ("PrivateUsage", ctypes.c_size_t),
    ]


def log_path(name: str) -> Path:
    return Path(tempfile.gettempdir()) / f"ahk-{name}-debug.log"


_MEM_FN = None


def _memory() -> tuple[int, int]:
    global _MEM_FN
    if os.name != "nt":
        return 0, 0
    if _MEM_FN is None:
        kernel = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel.GetCurrentProcess.restype = wintypes.HANDLE
        fn = kernel.K32GetProcessMemoryInfo
        fn.argtypes = [
            wintypes.HANDLE,
            ctypes.POINTER(_MemoryCounters),
            wintypes.DWORD,
        ]
        fn.restype = wintypes.BOOL
        _MEM_FN = (kernel.GetCurrentProcess, fn)
    get_current, fn = _MEM_FN
    counters = _MemoryCounters()
    counters.cb = ctypes.sizeof(counters)
    if not fn(get_current(), ctypes.byref(counters), counters.cb):
        return 0, 0
    return int(counters.WorkingSetSize), int(counters.PrivateUsage)


def _trim(path: Path) -> None:
    try:
        if not path.is_file() or path.stat().st_size <= _MAX_BYTES:
            return
        lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
        path.write_text("\n".join(lines[-200:]) + "\n", encoding="utf-8")
    except OSError:
        return


def _append(name: str, line: str) -> None:
    path = log_path(name)
    stamp = time.strftime("%Y-%m-%d %H:%M:%S")
    try:
        with _LOCK:
            _trim(path)
            with path.open("a", encoding="utf-8") as f:
                f.write(f"{stamp}\t{line}\n")
    except OSError:
        return


def _sample_loop(name: str) -> None:
    while True:
        working, private = _memory()
        threads = threading.active_count()
        _append(
            name,
            "mem\tws=%dMB\tprivate=%dMB\tthreads=%d"
            % (working // (1024 * 1024), private // (1024 * 1024), threads),
        )
        time.sleep(_SAMPLE_SEC)


def _request_path(handler: Any) -> str:
    raw = getattr(handler, "path", "") or ""
    return raw.split("?", 1)[0]


def _note_request(name: str, handler: Any, elapsed_ms: float) -> None:
    path = _request_path(handler)
    health = path.rstrip("/") in ("/health", "/api/health")
    if health and elapsed_ms < _SLOW_HEALTH_MS:
        return
    if not health and elapsed_ms < _SLOW_MS:
        return
    command = getattr(handler, "command", "") or ""
    _append(name, "req\t%d ms\t%s %s" % (round(elapsed_ms), command, path))


def install(handler_cls: type, name: str) -> None:
    """Start the memory sample and time this handler's requests."""
    if name not in _STARTED:
        _STARTED.add(name)
        _append(name, "started\tpid=%d" % os.getpid())
        threading.Thread(
            target=_sample_loop,
            args=(name,),
            name=f"resident-debug-{name}",
            daemon=True,
        ).start()
    if getattr(handler_cls, "_resident_debug", False):
        return
    original = handler_cls.handle_one_request

    def handle_one_request(self: Any, *args: Any, **kwargs: Any) -> Any:
        started = time.perf_counter()
        try:
            return original(self, *args, **kwargs)
        finally:
            _note_request(name, self, (time.perf_counter() - started) * 1000)

    handler_cls.handle_one_request = handle_one_request  # type: ignore[method-assign]
    handler_cls._resident_debug = True  # type: ignore[attr-defined]
