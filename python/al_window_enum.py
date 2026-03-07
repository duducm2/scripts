# AppLauncher daemon: EnumWindows + GetWindowThreadProcessId to resolve Cursor/Code targets.
# Returns primary (habits/home/punctual/work) and fallback (first non-preview) HWNDs.

import ctypes
from ctypes import wintypes

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
TARGET_KEYWORDS = ("habits", "home", "punctual", "work")
PREVIEW_KEYWORD = "preview"
CURSOR_EXES = ("cursor.exe", "code.exe")

# EnumWindows callback: WNDENUMPROC (HWND, LPARAM) -> BOOL
WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)


def _get_process_exe(pid: int) -> str:
    if pid == 0:
        return ""
    try:
        h = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        if not h or h == 0xFFFFFFFF:
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


def _get_window_text(hwnd: wintypes.HWND) -> str:
    buf = ctypes.create_unicode_buffer(256)
    if user32.GetWindowTextW(hwnd, buf, 256):
        return buf.value or ""
    return ""


def resolve_cursor_targets() -> dict:
    """Return {primaryHwnd: int, fallbackHwnd: int, reason: str}."""
    primary = 0
    fallback = 0
    collected = []

    def enum_cb(hwnd: wintypes.HWND, lparam: wintypes.LPARAM) -> wintypes.BOOL:
        nonlocal primary, fallback, collected
        if not user32.IsWindow(hwnd):
            return True
        tid, pid = wintypes.DWORD(), wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        exe = _get_process_exe(pid.value)
        if exe not in CURSOR_EXES:
            return True
        title = _get_window_text(hwnd)
        title_lower = title.lower()
        if PREVIEW_KEYWORD in title_lower:
            return True
        h = int(hwnd) if hasattr(hwnd, "__int__") else hwnd
        if not fallback:
            fallback = h
        if any(kw in title_lower for kw in TARGET_KEYWORDS):
            primary = h
            return False  # stop enumeration
        collected.append(h)
        return True

    callback = WNDENUMPROC(enum_cb)
    user32.EnumWindows(callback, 0)

    if primary:
        return {
            "primaryHwnd": primary,
            "fallbackHwnd": fallback or primary,
            "reason": "target",
        }
    if fallback:
        return {"primaryHwnd": 0, "fallbackHwnd": fallback, "reason": "fallback"}
    return {"primaryHwnd": 0, "fallbackHwnd": 0, "reason": "none"}
