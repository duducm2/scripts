# mousemaster.ahk — revision notes

## Purpose
Implements a **keyboard-driven mouse navigation overlay** (“hinting”):
- On activation, scans the active window’s UIA tree for interactive elements.
- Draws an always-on-top, click-through overlay with short hint labels.
- Captures user keystrokes via `InputHook` to select a hint.
- Executes a native mouse move + click at the target element’s center.

## Runtime model / methodology
- **AutoHotkey v2**, `#Warn`.
- Uses **UIA-v2** to enumerate UI elements and their bounding boxes (`Location`).
- Uses a **click-through overlay GUI** (`+E0x20`, `WinSetTransColor`) so it does not block mouse interaction.
- Uses `InputHook("T10")` to capture typed hints with timeout and Escape cancellation.
- Executes actions via **Win32 cursor + mouse injection**:
  - `SetCursorPos`
  - `mouse_event` left down/up

## Key entry points
- Hotkey `^!#c` toggles activation:
  - Captures active window HWND via `WinGetID("A")`.
  - Calls `Mousemaster_Activate(ActiveWinID)` / `Mousemaster_Deactivate()`.
- `Mousemaster_Activate(WinID)`:
  - UIA scan: `UIA.ElementFromHandle(WinID)` then `FindElements({IsOffscreen:false, IsEnabled:true})`.
  - Filters by ControlType (Button/Edit/Hyperlink/MenuItem/Checkbox/RadioButton).
  - Builds `MousemasterElements` with hint + rect + center coordinates.
  - Renders overlay labels at element positions.
  - Starts `InputHook` with handlers:
    - `Mousemaster_OnChar`
    - `Mousemaster_OnEnd`
- `Mousemaster_OnChar(...)`:
  - Maintains `UserInputBuffer`, searches for partial/exact matches.
  - On unique exact match: stops hook, calls `Mousemaster_PerformAction(...)`.
- `Mousemaster_PerformAction(elementObject)`:
  - Deactivates overlay, re-activates original window, moves mouse to center, clicks.

## Internal structure
- **Hint generation**:
  - `Mousemaster_GenerateHint(index)` maps element index to an A–Z style label (base-26-like).
- **Error handling**:
  - Uses tooltips + timers for temporary error/status messaging.
- **State**:
  - `MousemasterActive`, `MousemasterOverlayGui`, `MousemasterElements`, `UserInputBuffer`, `MM_InputHook`, `ActiveWinID`.

## Dependencies
- `UIA-v2\Lib\UIA.ahk`.

## Notable patterns (as used here)
- **NoActivate overlay**: overlay is shown with `"NoActivate"` so the foreground app remains active.
- **Single-threaded but reactive**: input captured through `InputHook` callbacks rather than busy loops.
- **Native injection**: click performed via Win32 APIs rather than AHK `Click`/`MouseMove`.

## DIP investigation output
### Findings summary
(paste here)

### Recommended changes
(paste here)

### Risks / regressions to watch
(paste here)

## DIP investigation inquiry
You are investigating `mousemaster.ahk`, a hint-overlay + input-capture system. Key constraints: fast UIA scanning, overlay that does not steal focus, and click execution that remains reliable across DPI/scaling and different apps.

Please research and report on:
- **Modern / alternative methods**:
  - Existing hinting tools (e.g., PowerToys, browser-only link hint extensions) and whether they can replace this workflow.
  - WebView2 overlay or other UI frameworks for rendering and hit testing.
- **Code efficiency optimizations**:
  - Faster UIA enumeration strategies (scoping, caching, filtering early).
  - Reducing overlay control count (draw-on-canvas vs many `Text` controls).
  - Better heuristics for “clickable” elements beyond ControlType.
- **Parallelism / async feasibility in AutoHotkey**:
  - Whether UIA scanning can be offloaded to a worker process while the main hotkey remains responsive.
- **Complex background macro feasibility without interrupting typing**:
  - Whether clicks can be dispatched without moving the system cursor (SendInput vs UIA Invoke).
  - Feasibility of “background click” while user continues typing in another app (and the OS constraints around global cursor movement).

