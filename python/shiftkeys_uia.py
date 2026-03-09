# ShiftKeys daemon: UIA helpers for FindElement / WaitElementState (Phase 4).
# Real implementation can use comtypes + IUIAutomationClient or pywinauto; this is a stub.

# Selector profiles per app (name/type fallbacks)
GEMINI_STOP_BUTTON_NAMES = ("Stop response",)
CHATGPT_BUTTON_NAMES = (
    "Stop streaming",
    "Interromper transmissão",
    "Stop",
    "Interromper",
)


def find_element(hwnd: int, name_substring: str, control_type_id: int = 50000) -> bool:
    """Return True if an element with Name containing name_substring exists under hwnd. Stub: False."""
    return False


def find_gemini_stop_button(hwnd: int) -> bool:
    """True if 'Stop response' button exists. Stub: False."""
    return find_element(hwnd, "Stop response", 50000)


def wait_element_state(
    hwnd: int, name_substring: str, timeout_ms: int, poll_ms: int = 300
) -> str:
    """Poll until element appears or timeout. Returns 'found' or 'timeout'. Stub: always 'timeout'."""
    import time

    time.sleep(min(timeout_ms / 1000.0, 5))
    return "timeout"
