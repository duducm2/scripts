# ShiftKeys daemon: O(1) context cache driven by WinEvent hooks.
# SetWinEventHook(EVENT_SYSTEM_FOREGROUND, EVENT_OBJECT_NAMECHANGE) updates
# IsChatGPTActive, IsGeminiActive, IsPowerBIActive, IsOutlookMainActive, etc.

import threading
import time

try:
    import ctypes
    from ctypes import wintypes
    import win32gui
    import win32process
    import win32con
    import pythoncom

    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

# Cache keys returned to AHK (mirror as g_ShiftKeys_IsChatGPTActive etc.)
CONTEXT_KEYS = (
    "IsChatGPTActive",
    "IsGeminiActive",
    "IsPowerBIActive",
    "IsOutlookMainActive",
    "IsOutlookMessageActive",
    "IsOutlookAppointmentActive",
    "IsWikipediaActive",
    "IsChromePdfViewerActive",
)

_lock = threading.Lock()
_cache = {k: False for k in CONTEXT_KEYS}
_last_update_ts = 0.0
_hook_thread = None
_stale_timeout_sec = 5.0


def _get_process_name(pid: int) -> str:
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


def _update_context_from_hwnd(hwnd: int) -> None:
    global _cache, _last_update_ts
    if not hwnd or not win32gui.IsWindow(hwnd):
        return
    try:
        title = (win32gui.GetWindowText(hwnd) or "").lower()
        pid = 0
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
        except Exception:
            pass
        exe = _get_process_name(pid) if pid else ""

        with _lock:
            _cache["IsChatGPTActive"] = (
                "chrome" in exe or "msedge" in exe
            ) and "chatgpt" in title
            _cache["IsGeminiActive"] = (
                "chrome" in exe or "msedge" in exe
            ) and "gemini" in title
            _cache["IsPowerBIActive"] = "pbidesktop.exe" in exe or "powerbi" in title
            _cache["IsOutlookMainActive"] = (
                "outlook.exe" in exe and "reminder" not in title
            )
            _cache["IsOutlookMessageActive"] = (
                "outlook.exe" in exe and " - Message " in title
            )
            _cache["IsOutlookAppointmentActive"] = (
                "outlook.exe" in exe and " - Appointment " in title
            )
            _cache["IsWikipediaActive"] = (
                "chrome" in exe or "msedge" in exe
            ) and "wikipedia" in title
            _cache["IsChromePdfViewerActive"] = (
                "chrome" in exe or "msedge" in exe
            ) and "pdf" in title
            _last_update_ts = time.monotonic()
    except Exception:
        pass


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
        if hwnd:
            _update_context_from_hwnd(hwnd)

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
            EVENT_SYSTEM_FOREGROUND = 0x0003
            EVENT_OBJECT_NAMECHANGE = 0x800C
            WINEVENT_OUTOFCONTEXT = 0x0000
            h1 = user32.SetWinEventHook(
                EVENT_SYSTEM_FOREGROUND,
                EVENT_SYSTEM_FOREGROUND,
                0,
                _win_event_callback,
                0,
                0,
                WINEVENT_OUTOFCONTEXT,
            )
            h2 = user32.SetWinEventHook(
                EVENT_OBJECT_NAMECHANGE,
                EVENT_OBJECT_NAMECHANGE,
                0,
                _win_event_callback,
                0,
                0,
                WINEVENT_OUTOFCONTEXT,
            )
            if not h1 and not h2:
                return
            try:
                while True:
                    pythoncom.PumpWaitingMessages()
                    time.sleep(0.05)
            finally:
                if h1:
                    user32.UnhookWinEvent(h1)
                if h2:
                    user32.UnhookWinEvent(h2)
        finally:
            pythoncom.CoUninitialize()


def start_context_hook() -> None:
    """Start the WinEvent hook in a background thread. Call once at daemon startup."""
    global _hook_thread
    if not HAS_WIN32 or _hook_thread is not None:
        return
    _hook_thread = threading.Thread(target=_hook_thread_run, daemon=True)
    _hook_thread.start()
    # Initial update from current foreground
    try:
        fg = win32gui.GetForegroundWindow()
        if fg:
            _update_context_from_hwnd(fg)
    except Exception:
        pass


def get_context_snapshot() -> dict:
    """Return a copy of the context cache for ResolveContext. Stale timeout triggers re-query of foreground."""
    global _last_update_ts
    with _lock:
        if time.monotonic() - _last_update_ts > _stale_timeout_sec:
            try:
                if HAS_WIN32:
                    fg = win32gui.GetForegroundWindow()
                    if fg:
                        _update_context_from_hwnd(fg)
            except Exception:
                pass
        return dict(_cache)
