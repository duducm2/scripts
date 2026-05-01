# GeminiToCursorBridge.ahk — revision notes

## Purpose
Self-contained bridge module that performs the “**Gemini → Cursor**” transfer:
1) copy Gemini’s last response to clipboard, 2) activate a Cursor window, 3) focus Cursor AI input, 4) paste, 5) send, 6) preserve/restore context where possible.

This script is designed with **verification layers** to avoid silent failures (wrong window, stale clipboard, focus lost mid-sequence).

## Runtime model / methodology
- **AutoHotkey v2** module-style file (functions + constants).
- Strong emphasis on **fail-fast verification** and **instrumentation**:
  - NDJSON logging to `.cursor\debug.log` via `Bridge_Log(...)`.
  - Validation checkpoints: clipboard length, result file, active window identity.
- Prefers **HWND-based selection** over fuzzy title matching (when possible) and explicitly skips “preview” windows.

## Key entry points
- `CopyFromGeminiToCursor(...)` (documented as the single entry point in the header comment; defined later in the file).
- Cursor window selection helpers:
  - `Bridge_GetAnyCursorHwnd()`
  - `Bridge_GetCursorHwndForProject(projectPath)`
  - `Bridge_FindAndActivateCursorWindow(projectPath)`
- Cursor focus helper:
  - `Bridge_FocusCursorAITextField(targetHwnd := 0)` (uses `^i`, then `{Tab 2}`).
- Gemini copy helper:
  - `Bridge_CopyGeminiLastMessageToClipboard()`:
    - Locates the running `Gemini.ahk` script window (`AutoHotkey64.exe` / `AutoHotkey32.exe`) by title.
    - Uses `WM_COPY_LAST_GEMINI := 0x8001` (message contract) and verifies success.

## Internal structure
- **Verification layers (from file header)**:
  - After copy: re-check clipboard validity and the result file (`.cursor\gemini_copy_result.txt`).
  - After activation: ensure the active window is the intended Cursor target; re-activate if needed.
  - Before paste: ensure clipboard is still valid and Cursor is still active.
- **Non-invasive mouse handling**:
  - `Bridge_MoveMouseToCenter(hwnd)` and `Bridge_MaybeCenterMouse(...)` attempt to avoid “halo” / stray pointer artifacts; integrates with `WM_MaybeCenterMouse` when available.
- **Project targeting**:
  - `Bridge_ExtractProjectMatchSegments(projectPath)` derives title match segments (most specific first) to prefer the correct Cursor project.

## Dependencies
- Soft dependency on WindowManagement helpers (`WM_MaybeCenterMouse`, `WMAutomation_SuppressCursorCentering`) via try/catch.
- Assumes a separate `Gemini.ahk` process implements the `WM_COPY_LAST_GEMINI` message contract and writes `.cursor\gemini_copy_result.txt`.

## Notable patterns (as used here)
- **Determinism over convenience**: “active window must be X” checks before sending keys.
- **Observability**: structured logs with hypothesis IDs to debug field failures.
- **Graceful fallback**: if richer integrations are missing, still operates with minimal core actions.

## DIP investigation output
### Findings summary
(paste here)

### Recommended changes
(paste here)

### Risks / regressions to watch
(paste here)

## DIP investigation inquiry
You are investigating `GeminiToCursorBridge.ahk`, whose objective is to transfer text from Gemini to Cursor reliably. The hardest constraint is doing this in a way that can be **asynchronous/non-interruptive**, ideally not stealing focus or disrupting active typing.

Please research and report on:
- **Modern / alternative methods**:
  - Replacing UI-driven copy/paste with API-based transfer (Gemini API + Cursor integration, if possible).
  - Browser extension exporting the response via localhost/websocket to avoid UIA copy entirely.
  - Clipboard managers / custom clipboard formats as a transport.
- **Code efficiency optimizations**:
  - More robust window selection than title substring matching (process tree, window class, accessibility properties).
  - Faster activation/focus logic; better handling of “Cursor has multiple windows/projects”.
  - Reducing sleeps by using event-driven readiness checks.
- **Parallelism / async feasibility in AutoHotkey**:
  - Multi-process architecture: a background “automation daemon” performing UI operations while the main script remains responsive.
  - Named pipes / MMF / message queue patterns for sequencing.
- **Complex background macros without interrupting typing**:
  - Is it feasible to run “copy from Gemini → paste into Cursor → send” without bringing either window to foreground?
  - Which parts fundamentally require foreground focus (keyboard input injection) vs can be done via UIA Invoke, WM_COMMAND, accessibility patterns, or app-internal APIs.
  - Practical design recommendations for “non-interruptive transfer” (e.g., queue requests, run at idle, show passive overlay confirmation, return focus deterministically).

