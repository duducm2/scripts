# WindowManagement.ahk — revision notes

## Purpose
Consolidates **window management** hotkeys: minimize/maximize, move windows between monitors, cycle windows per monitor, close/minimize windows on a given monitor, and related automation helpers (including coordination with a background “daemon” via IPC flags).

## Runtime model / methodology
- **AutoHotkey v2**, `#SingleInstance Force`, `#UseHook True`.
- Central model: **hotkey → window enumeration/activation → Win32/AutoHotkey window operations**.
- Uses a periodic timer to monitor active window state:
  - 100ms in legacy mode
  - 250ms when daemon cache is enabled
- Integrates an “automation suppression” mechanism to prevent cursor-centering or other intrusive behaviors during scripted switches.

## Key entry points
- **Auto-execute section**:
  - Includes `env.ahk`, `GeminiToCursorBridge.ahk`, `Utils.ahk`, `aux\WMIPC.ahk`.
  - Defines `WM_AUTOMATION_SWITCH_DEFAULT_MS` and sets up timers.
  - Adds tray menu item “Test Cycle M1”.
- **Hotkeys (representative, from scan)**:
  - `#!+6` minimize active window.
  - `#!+M` maximize active window.
  - `^!#a/^!#s/^!#d/^!#f` move active window to ordered monitor 1–4.
  - `^!+#a/...` close window on monitor 1–4 (plus fallbacks like `^!+g`, scan-code variants).
  - `^!#q/^!#w/^!#e/^!#r` cycle windows on monitor 1–4.
  - `^!+#q/^!+#w/^!+#e/^!+#r` minimize window on monitor 1–4.
- **Core helpers**:
  - `WM_MaximizeActiveWindow()`: uses `WinMaximize`, falls back to `WM_SYSCOMMAND` `SC_MAXIMIZE`.
  - `WM_IsExcludedIndicatorWindow(hwnd)`: avoids selecting helper windows (e.g., Handy, the WM script itself).
  - Daemon integration:
    - `WM_UsesAutomationDaemon()`
    - `WMAutomation_SuppressCursorCentering(reason, durationMs)`
    - `WMAutomation_ClearCursorSuppression(reason)`
    - `WMAutomation_CursorCenteringSuppressed(hwnd)`

## Internal structure
- **Monitor ordering logic**:
  - `GetMonitorIndexByOrder(order)` computes monitor centers and sorts left-to-right.
  - `MoveWinToOrderedMonitor(order)` maps order → monitor index → actual move.
- **Automation suppression**:
  - Tracks suppression window via `g_WMAutomationSuppressUntil`.
  - Optionally coordinates suppression with the IPC daemon (`WMIPC_BeginAutomationSwitch`, `WMIPC_EndAutomationSwitch`).
- **Activation safety**:
  - Helper `TryActivateWindow_WM(...)` checks existence and shows overlay on failure.

## Dependencies
- `env.ahk`
- `Utils.ahk` (overlay helpers, center mouse, etc.)
- `aux\WMIPC.ahk` (daemon/IPC integration; flags like `WM_USE_DAEMON`, `WM_USE_PIPE_IPC`, `WM_USE_EVENT_HOOK_CACHE`)
- `GeminiToCursorBridge.ahk` is included to enable certain cross-workflow hotkeys.

## Notable patterns (as used here)
- **Fallback hotkeys** when Electron/IDE windows swallow Win-key chords (multiple bindings for the same intent).
- **Daemon feature flags** allow progressively replacing polling with event/cache driven state.
- **“Do no harm” guards**: suppression window prevents cursor centering during automation steps.

## DIP investigation output
### Findings summary
(paste here)

### Recommended changes
(paste here)

### Risks / regressions to watch
(paste here)

## DIP investigation inquiry
You are investigating `WindowManagement.ahk`, which performs frequent window activation/movement actions with hotkeys. Key constraints: correctness (right window/monitor), reliability across apps that resist activation, and minimizing disruption to the user.

Please research and report on:
- **Modern / alternative methods**:
  - PowerToys FancyZones / Windows Snap / virtual desktop APIs and what they can replace.
  - Window management libraries (e.g., AutoHotkey community libs, native APIs via other languages).
- **Code efficiency optimizations**:
  - Event-driven foreground tracking (WinEvent hooks) vs polling timers.
  - More robust window lists and per-monitor indexing strategies.
- **Parallelism / async feasibility in AutoHotkey**:
  - Using a background daemon for window state caches and action execution.
- **Complex background macros without interrupting typing**:
  - Feasibility of moving/minimizing windows without stealing keyboard focus.
  - Strategies to avoid “focus flicker” during rapid window cycling/moves.

