# ShiftKeys daemon: UIA helpers for FindElement / WaitElementState (Phase 4).
# Uses pywinauto (UIA) + STA COM init for calls from the named-pipe worker thread.

from __future__ import annotations

import time

# Selector profiles per app (name/type fallbacks)
GEMINI_STOP_BUTTON_NAMES = ("Stop response",)
CHATGPT_BUTTON_NAMES = (
    "Stop streaming",
    "Interromper transmissão",
    "Stop",
    "Interromper",
)

_MAX_DESCENDANTS = 15000

try:
    import pythoncom
    import win32gui
    from pywinauto import Desktop

    _HAS_UIA = True
except ImportError:
    pythoncom = None  # type: ignore[assignment]
    win32gui = None  # type: ignore[assignment]
    Desktop = None  # type: ignore[assignment]
    _HAS_UIA = False


def _valid_hwnd(hwnd: int) -> bool:
    if not hwnd:
        return False
    if win32gui is None:
        return True
    try:
        return bool(win32gui.IsWindow(int(hwnd)))
    except Exception:
        return False


def _current_control_type(wrap) -> int | None:
    """IUIAutomationElement.CurrentControlType (int), e.g. 50000 = Button."""
    try:
        el = wrap.element_info.element
        return int(el.CurrentControlType)
    except Exception:
        return None


def _find_element_core(hwnd: int, name_substring: str, control_type_id: int) -> bool:
    """UIA walk under hwnd; caller must have COM initialized for this thread."""
    if not _HAS_UIA or Desktop is None:
        return False
    if not _valid_hwnd(hwnd):
        return False
    needle = (name_substring or "").strip().lower()
    if not needle:
        return False
    want = int(control_type_id)
    root = Desktop(backend="uia").window(handle=int(hwnd))
    count = 0
    for el in root.descendants():
        count += 1
        if count > _MAX_DESCENDANTS:
            break
        try:
            if _current_control_type(el) != want:
                continue
            name = (el.window_text() or "").lower()
            if needle in name:
                return True
        except Exception:
            continue
    return False


def find_element(hwnd: int, name_substring: str, control_type_id: int = 50000) -> bool:
    """Return True if an element with Name containing name_substring exists under hwnd."""
    if not _HAS_UIA or pythoncom is None:
        return False
    pythoncom.CoInitialize()
    try:
        return _find_element_core(hwnd, name_substring, control_type_id)
    except Exception:
        return False
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def find_gemini_stop_button(hwnd: int) -> bool:
    """True if a Gemini-style Stop control exists (name variants, Button)."""
    for part in GEMINI_STOP_BUTTON_NAMES:
        if find_element(hwnd, part, 50000):
            return True
    return False


def wait_element_state(
    hwnd: int,
    name_substring: str,
    timeout_ms: int,
    poll_ms: int = 300,
    control_type_id: int = 50000,
) -> str:
    """Poll until element appears or timeout. Returns 'found' or 'timeout'."""
    if not _HAS_UIA or pythoncom is None:
        return "timeout"
    if not _valid_hwnd(hwnd) or not (name_substring or "").strip():
        return "timeout"
    want = int(control_type_id)
    pythoncom.CoInitialize()
    try:
        deadline = time.monotonic() + max(0, int(timeout_ms)) / 1000.0
        step = max(1, int(poll_ms)) / 1000.0
        while time.monotonic() < deadline:
            try:
                if _find_element_core(hwnd, name_substring, want):
                    return "found"
            except Exception:
                pass
            time.sleep(step)
        return "timeout"
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
