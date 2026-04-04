# WindowManagement daemon: WinEvent hooks and O(1) window-state cache.
# SetWinEventHook(EVENT_OBJECT_CREATE, EVENT_OBJECT_DESTROY, EVENT_SYSTEM_FOREGROUND)
# maintains hwnd/pid maps and Cursor-filtered sets. Resync thread for integrity.

import threading
import time
import ctypes
from ctypes import wintypes

try:
    import win32gui
    import win32process
    import win32con
    import pythoncom

    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

# WinEvent constants
EVENT_OBJECT_CREATE = 0x8000
EVENT_OBJECT_DESTROY = 0x8001
EVENT_SYSTEM_FOREGROUND = 0x0003
OBJID_WINDOW = 0
WINEVENT_OUTOFCONTEXT = 0x0000

# Window style
GWL_EXSTYLE = -20
WS_EX_TOOLWINDOW = 0x00000080
MONITOR_DEFAULTTONEAREST = 2

# Cache: hwnd -> {pid, title, class, exe, monitor, rect, zHint, alive}
# pid -> [hwnd...]
# _foreground_hwnd, _cursor_windows (main), _preview_windows
_lock = threading.Lock()
_hwnd_to_info: dict = {}
_pid_to_hwnds: dict = {}
_foreground_hwnd: int = 0
_last_non_gemini_foreground_hwnd: int = 0
_z_sequence: int = 0
_hook_thread = None
_resync_thread = None
_shutdown = threading.Event()
_resync_interval_sec = 30.0
_last_resync_ts = 0.0
_event_sequence = 0
_cursor_suppressed_until_ms: int = 0
_cursor_suppression_reason: str = ""


def _get_process_exe(pid: int) -> str:
    if pid == 0:
        return ""
    try:
        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        kernel32 = ctypes.windll.kernel32
        h = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        if not h:
            return ""
        try:
            size = wintypes.DWORD(4096)
            buf = ctypes.create_unicode_buffer(4096)
            if kernel32.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(size)):
                path = buf.value
                return path.rsplit("\\", 1)[-1].lower() if path else ""
        finally:
            kernel32.CloseHandle(h)
    except Exception:
        pass
    return ""


def _get_window_info(hwnd: int) -> dict | None:
    if not hwnd or not HAS_WIN32:
        return None
    try:
        if not win32gui.IsWindow(hwnd):
            return None
        title = win32gui.GetWindowText(hwnd) or ""
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
        except Exception:
            pid = 0
        exe = _get_process_exe(pid) if pid else ""
        try:
            class_name = win32gui.GetClassName(hwnd) or ""
        except Exception:
            class_name = ""
        try:
            rect = win32gui.GetWindowRect(hwnd)
            left, top, right, bottom = rect
        except Exception:
            left = top = right = bottom = 0
        try:
            user32 = ctypes.windll.user32
            hmon = user32.MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST)
        except Exception:
            hmon = 0
        return {
            "pid": pid,
            "title": title,
            "class": class_name,
            "exe": exe,
            "monitor": hmon,
            "rect": {"left": left, "top": top, "right": right, "bottom": bottom},
            "alive": True,
        }
    except Exception:
        return None


def _now_ms() -> int:
    return int(time.time() * 1000)


def _looks_like_gemini_window(info: dict | None) -> bool:
    if not info:
        return False
    title = str(info.get("title", "") or "").lower()
    if "gemini" in title:
        return True
    exe = str(info.get("exe", "") or "").lower()
    return exe == "chrome.exe" and "gemini" in title


def _is_visible_top_level(hwnd: int) -> bool:
    if not hwnd or not HAS_WIN32:
        return False
    try:
        if not win32gui.IsWindow(hwnd) or not win32gui.IsWindowVisible(hwnd):
            return False
        ex_style = ctypes.windll.user32.GetWindowLongPtrW(hwnd, GWL_EXSTYLE)
        if ex_style & WS_EX_TOOLWINDOW:
            return False
        class_name = win32gui.GetClassName(hwnd) or ""
        if class_name in ("Progman", "WorkerW"):
            return False
        if not (win32gui.GetWindowText(hwnd) or "").strip():
            return False
        return True
    except Exception:
        return False


def _update_cache_add(hwnd: int) -> None:
    global _z_sequence, _hwnd_to_info, _pid_to_hwnds
    if not _is_visible_top_level(hwnd):
        return
    info = _get_window_info(hwnd)
    if not info:
        return
    with _lock:
        _z_sequence += 1
        info["zHint"] = _z_sequence
        _hwnd_to_info[hwnd] = info
        pid = info["pid"]
        _pid_to_hwnds.setdefault(pid, []).append(hwnd)


def _update_cache_remove(hwnd: int) -> None:
    global _hwnd_to_info, _pid_to_hwnds, _foreground_hwnd
    with _lock:
        if hwnd in _hwnd_to_info:
            info = _hwnd_to_info.pop(hwnd)
            pid = info.get("pid", 0)
            if pid in _pid_to_hwnds:
                lst = _pid_to_hwnds[pid]
                if hwnd in lst:
                    lst.remove(hwnd)
                if not lst:
                    del _pid_to_hwnds[pid]
        if _foreground_hwnd == hwnd:
            _foreground_hwnd = 0


def _update_foreground(hwnd: int) -> None:
    global _foreground_hwnd, _last_non_gemini_foreground_hwnd, _z_sequence, _hwnd_to_info
    if not hwnd:
        return
    if hwnd not in _hwnd_to_info:
        _update_cache_add(hwnd)
    with _lock:
        _foreground_hwnd = hwnd
        _z_sequence += 1
        if hwnd in _hwnd_to_info:
            _hwnd_to_info[hwnd]["zHint"] = _z_sequence
        info = _hwnd_to_info.get(hwnd)
        if info and not _looks_like_gemini_window(info):
            _last_non_gemini_foreground_hwnd = hwnd


def begin_automation_switch(duration_ms: int = 1500, reason: str = "") -> dict:
    global _cursor_suppressed_until_ms, _cursor_suppression_reason
    duration_ms = max(0, int(duration_ms or 0))
    until_ms = _now_ms() + duration_ms
    with _lock:
        _cursor_suppressed_until_ms = until_ms
        _cursor_suppression_reason = reason or ""
    return {
        "cursorSuppressed": duration_ms > 0,
        "cursorSuppressedUntilMs": until_ms,
        "cursorSuppressionReason": reason or "",
    }


def end_automation_switch() -> dict:
    global _cursor_suppressed_until_ms, _cursor_suppression_reason
    with _lock:
        _cursor_suppressed_until_ms = 0
        _cursor_suppression_reason = ""
    return {
        "cursorSuppressed": False,
        "cursorSuppressedUntilMs": 0,
        "cursorSuppressionReason": "",
    }


def get_automation_context() -> dict:
    with _lock:
        fg = _foreground_hwnd
        last_non_gemini = _last_non_gemini_foreground_hwnd
        fg_info = _hwnd_to_info.get(fg, {})
        last_info = _hwnd_to_info.get(last_non_gemini, {})
        suppressed = _cursor_suppressed_until_ms > _now_ms()
        return {
            "foregroundHwnd": fg,
            "foregroundTitle": fg_info.get("title", ""),
            "foregroundExe": fg_info.get("exe", ""),
            "lastNonGeminiHwnd": last_non_gemini,
            "lastNonGeminiTitle": last_info.get("title", ""),
            "cursorSuppressed": suppressed,
            "cursorSuppressedUntilMs": _cursor_suppressed_until_ms if suppressed else 0,
            "cursorSuppressionReason": _cursor_suppression_reason if suppressed else "",
        }


def _get_cursor_windows() -> list:
    """Active Cursor windows (exclude preview)."""
    with _lock:
        out = []
        for hwnd, info in list(_hwnd_to_info.items()):
            if not info.get("alive", True):
                continue
            if (info.get("exe") or "").lower() != "cursor.exe":
                continue
            title = (info.get("title") or "").lower()
            if "preview" in title:
                continue
            out.append({"hwnd": hwnd, "title": info.get("title", "")})
        return out


def _get_preview_windows() -> list:
    """Cursor windows with 'preview' in title."""
    with _lock:
        out = []
        for hwnd, info in list(_hwnd_to_info.items()):
            if not info.get("alive", True):
                continue
            if (info.get("exe") or "").lower() != "cursor.exe":
                continue
            title = (info.get("title") or "").lower()
            if "preview" not in title:
                continue
            out.append({"hwnd": hwnd, "title": info.get("title", "")})
        return out


def _get_monitor_handles() -> list:
    """List of monitor handles in display order (index 0 = first monitor)."""
    try:
        import win32api

        return [m[0] for m in win32api.EnumDisplayMonitors(None, None)]
    except Exception:
        return []


def _get_visible_windows_on_monitor(mon_index: int) -> list:
    """Visible windows on monitor by index (1-based like AHK). Returns list of {hwnd, left, top, right, bottom, z}."""
    monitors = _get_monitor_handles()
    if mon_index < 1 or mon_index > len(monitors):
        return []
    h_target = monitors[mon_index - 1]
    TOL = 40
    user32 = ctypes.windll.user32
    visible = []
    try:

        def enum_cb(hwnd, _):
            if not _is_visible_top_level(hwnd):
                return True
            hmon = user32.MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST)
            if hmon != h_target:
                return True
            try:
                left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            except Exception:
                return True
            center_x = (left + right) // 2
            center_y = (top + bottom) // 2
            for v in visible:
                if (
                    v["left"] <= center_x <= v["right"]
                    and v["top"] <= center_y <= v["bottom"]
                ):
                    return True
            visible.append(
                {
                    "hwnd": hwnd,
                    "left": left,
                    "top": top,
                    "right": right,
                    "bottom": bottom,
                    "z": 0,
                }
            )
            return True

        EnumWindowsProc = ctypes.WINFUNCTYPE(
            wintypes.BOOL, wintypes.HWND, wintypes.LPARAM
        )
        user32.EnumWindows(EnumWindowsProc(enum_cb), 0)
    except Exception:
        pass
    # Sort by Y then X (same as AHK)
    n = len(visible)
    for i in range(n - 1):
        for j in range(n - 1 - i):
            a, b = visible[j], visible[j + 1]
            row_diff = a["top"] - b["top"]
            if row_diff > TOL or (abs(row_diff) <= TOL and a["left"] > b["left"]):
                visible[j], visible[j + 1] = b, a
    return visible


def _resync_all() -> None:
    """Full enumeration to refresh cache and z-order."""
    global _hwnd_to_info, _pid_to_hwnds, _last_resync_ts
    with _lock:
        _hwnd_to_info.clear()
        _pid_to_hwnds.clear()
    try:
        z = 0

        def enum_cb(hwnd, _):
            nonlocal z
            if not _is_visible_top_level(hwnd):
                return True
            info = _get_window_info(hwnd)
            if not info:
                return True
            z += 1
            info["zHint"] = z
            with _lock:
                _hwnd_to_info[hwnd] = info
                pid = info["pid"]
                _pid_to_hwnds.setdefault(pid, []).append(hwnd)
            return True

        EnumWindowsProc = ctypes.WINFUNCTYPE(
            wintypes.BOOL, wintypes.HWND, wintypes.LPARAM
        )
        ctypes.windll.user32.EnumWindows(EnumWindowsProc(enum_cb), 0)
        with _lock:
            _last_resync_ts = time.monotonic()
    except Exception:
        pass


def _resync_loop() -> None:
    while not _shutdown.wait(timeout=_resync_interval_sec):
        _resync_all()


if HAS_WIN32:
    WinEventProc = ctypes.WINFUNCTYPE(
        None,
        wintypes.HANDLE,
        wintypes.DWORD,
        wintypes.HWND,
        wintypes.LONG,
        wintypes.LONG,
        wintypes.DWORD,
        wintypes.DWORD,
    )

    @WinEventProc
    def _win_event_callback(
        hWinEventHook, event, hwnd, idObject, idChild, idEventThread, dwmsEventTime
    ):
        if idObject != OBJID_WINDOW:
            return
        if event == EVENT_OBJECT_CREATE and hwnd:
            _update_cache_add(hwnd)
        elif event == EVENT_OBJECT_DESTROY and hwnd:
            _update_cache_remove(hwnd)
        elif event == EVENT_SYSTEM_FOREGROUND and hwnd:
            _update_foreground(hwnd)

    def _hook_thread_run() -> None:
        pythoncom.CoInitialize()
        try:
            user32 = ctypes.windll.user32
            user32.SetWinEventHook.restype = wintypes.HANDLE
            user32.SetWinEventHook.argtypes = (
                wintypes.DWORD,
                wintypes.DWORD,
                wintypes.HMODULE,
                WinEventProc,
                wintypes.DWORD,
                wintypes.DWORD,
                wintypes.DWORD,
            )
            h1 = user32.SetWinEventHook(
                EVENT_OBJECT_CREATE,
                EVENT_OBJECT_CREATE,
                0,
                _win_event_callback,
                0,
                0,
                WINEVENT_OUTOFCONTEXT,
            )
            h2 = user32.SetWinEventHook(
                EVENT_OBJECT_DESTROY,
                EVENT_OBJECT_DESTROY,
                0,
                _win_event_callback,
                0,
                0,
                WINEVENT_OUTOFCONTEXT,
            )
            h3 = user32.SetWinEventHook(
                EVENT_SYSTEM_FOREGROUND,
                EVENT_SYSTEM_FOREGROUND,
                0,
                _win_event_callback,
                0,
                0,
                WINEVENT_OUTOFCONTEXT,
            )
            if not (h1 or h2 or h3):
                return
            try:
                fg = win32gui.GetForegroundWindow()
                if fg:
                    _update_foreground(fg)
                while not _shutdown.is_set():
                    pythoncom.PumpWaitingMessages()
                    time.sleep(0.05)
            finally:
                for handle in (h1, h2, h3):
                    if handle:
                        user32.UnhookWinEvent(handle)
        finally:
            pythoncom.CoUninitialize()


def start_wm_hooks() -> None:
    """Start WinEvent hook and resync threads. Call once at daemon startup."""
    global _hook_thread, _resync_thread
    if not HAS_WIN32 or _hook_thread is not None:
        return
    _resync_all()
    _hook_thread = threading.Thread(target=_hook_thread_run, daemon=True)
    _hook_thread.start()
    _resync_thread = threading.Thread(target=_resync_loop, daemon=True)
    _resync_thread.start()


def stop_wm_hooks() -> None:
    _shutdown.set()


# --- Cache API for IPC ---
def get_foreground_window_state() -> dict:
    with _lock:
        fg = _foreground_hwnd
        if not fg or fg not in _hwnd_to_info:
            suppressed = _cursor_suppressed_until_ms > _now_ms()
            return {
                "hwnd": 0,
                "pid": 0,
                "title": "",
                "class": "",
                "exe": "",
                "suppressCursorCentering": suppressed,
                "cursorSuppressedUntilMs": (
                    _cursor_suppressed_until_ms if suppressed else 0
                ),
                "cursorSuppressionReason": (
                    _cursor_suppression_reason if suppressed else ""
                ),
                "lastNonGeminiHwnd": _last_non_gemini_foreground_hwnd,
            }
        info = _hwnd_to_info[fg]
        suppressed = _cursor_suppressed_until_ms > _now_ms()
        return {
            "hwnd": fg,
            "pid": info.get("pid", 0),
            "title": info.get("title", ""),
            "class": info.get("class", ""),
            "exe": info.get("exe", ""),
            "suppressCursorCentering": suppressed,
            "cursorSuppressedUntilMs": _cursor_suppressed_until_ms if suppressed else 0,
            "cursorSuppressionReason": _cursor_suppression_reason if suppressed else "",
            "lastNonGeminiHwnd": _last_non_gemini_foreground_hwnd,
        }


def get_cursor_windows() -> list:
    return _get_cursor_windows()


def get_preview_windows() -> list:
    return _get_preview_windows()


def get_visible_windows_by_monitor(mon_index: int) -> list:
    return _get_visible_windows_on_monitor(mon_index)


def resolve_project_window(project_path: str) -> dict:
    """Return {hwnd, title} for first Cursor window whose title matches project path segments, or hwnd 0."""
    segments = _extract_project_match_segments(project_path)
    for w in _get_cursor_windows():
        title = w.get("title", "")
        if not title:
            continue
        for seg in segments:
            if seg and seg in title:
                return {"hwnd": w["hwnd"], "title": title}
    return {"hwnd": 0, "title": ""}


def _extract_project_match_segments(project_path: str) -> list:
    """Mirror AHK ExtractProjectMatchSegments: last path component and optional parent."""
    if not project_path or not project_path.strip():
        return []
    path = project_path.replace("/", "\\").strip()
    parts = [p for p in path.split("\\") if p]
    if not parts:
        return []
    segments = [parts[-1]]
    if len(parts) >= 2:
        segments.append(parts[-2])
    return segments
