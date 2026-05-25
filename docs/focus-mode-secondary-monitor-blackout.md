# Focus mode and secondary-monitor blackout

This document describes how **focus mode** (black overlays on every monitor except the one hosting the active window) is triggered, cancelled, and kept consistent across scripts. Implementation lives primarily in **`Utils.ahk`**; the **automatic dwell watcher** is started from **`Shift keys.ahk`**.

## What focus mode does

- **`EnableFocusMode(keepMonitorIndex)`** creates full-screen black **Gui** overlays on **all monitors except** the monitor index passed as `keepMonitorIndex` (the “keep clear” display). Overlays use `-DPIScale`, `+AlwaysOnTop`, and **`WS_EX_TRANSPARENT`** (`+E0x20`) so pointer events pass through to windows underneath.
- **`DisableFocusMode()`** destroys those overlays, clears globals, stops the focus-mode window timer, and **`StopPdfFocusMonitor()`** so auto monitoring does not keep running after a manual teardown.
- **`FocusMode_SetKeepMonitor(keepMonitorIndex)`** rebuilds overlays when the keep-clear display should move (same **anchor** window moved to another monitor). Does not stop timers or call **`DisableFocusMode()`**.

Related globals: `g_FocusModeOn`, `g_FocusModeOverlays`, `g_FocusModeTrackedWindow`, `g_FocusModeActiveMonitor`, `g_FocusModeAnchorHwnd`.

Cross-process state (Shift keys owns overlays; WindowManagement can request teardown):

- **`A_ScriptDir\.cursor\focus_mode_keep_monitor`** — one-line keep-monitor index while blackout is on
- **`A_ScriptDir\.cursor\focus_mode_disable_request`** — empty sentinel; **`FocusMode_RequestDisableCrossProcess()`** creates it, **`MonitorPdfFocus`** in Shift keys consumes it

## Monitor-scoped behavior (relocate vs disable)

When **`StartPdfFocusMonitor`** is active (Study Topic / dwell / Peek / **`#!+Y`** on multi-monitor):

| Situation | Action |
|-----------|--------|
| Another window on **same** keep monitor | Keep blackout |
| **Same anchor HWND** (`g_FocusModeAnchorHwnd`), foreground on **different** monitor | Relocate (`FocusMode_SetKeepMonitor`) — e.g. **`^!#a/s/d/f`** move |
| **Different HWND**, foreground on **different** monitor | **`DisableFocusMode()`** — e.g. alt-tab, click, **`^!#q/w/e/r`** cycle on another monitor |
| **`^!#q/w/e/r`** activates a window on monitor **other than** keep | **`FocusMode_RequestDisableCrossProcess()`** (file IPC) + monitor tick |

- **`Debounced`** mode (default on **`StudyTopic_ApplyBlackoutCountdownTimeout`** when not overridden): cross-monitor **relocate** waits **`g_PdfFocusDebounceMs`** (1200 ms) for move/maximize animation.
- **`Immediate`** mode (dwell, countdown, **`#!+Y`**): relocate on the next **200 ms** tick with no debounce.
- **Disable** on cross-monitor focus (non-anchor) is always immediate (no debounce).

## Entry points

### 1. Manual toggle (Win+Alt+Shift+Y)

- Hotkey **`#!+Y`** calls **`ToggleFocusMode()`**.
- **On:** `EnableFocusMode()` then **`StartPdfFocusMonitor(WinExist("A"), "Immediate")`** when **`MonitorGetCount() > 1`** — same relocate/disable rules as auto blackout; only the trigger is manual.
- **Off:** `DisableFocusMode()`.
- **`ToggleFocusMode()`** treats focus mode as **on** if **`g_FocusModeOn`** **or** **`g_FocusModeOverlays.Length > 0`**.

### 2. Study Topic / QuickLook (Win+Alt+Shift+X)

When QuickLook is active, **`#!+X`** runs **`StudyTopic_StartBlackoutCountdown(hwnd)`**, which shows the standard interactive banner (see **`docs/standard_information_display.md`**) and on timeout calls **`StudyTopic_ApplyBlackoutCountdownTimeout`**.

- During the **3 s** countdown the banner uses **`trackActiveMonitor`** and **`preserveUserFocus`**, so you can switch monitors/windows; the bar follows the foreground display.
- On timeout, **`keepIdx`** and the anchor HWND come from the **current foreground** (not a forced re-activate of the window that started the countdown).
- Countdown callers pass **`"Immediate"`** for relocate debounce.

### 3. Automatic dwell watcher (20 seconds)

Started once via **`FocusBlackoutWatcher_Start()`** from **`Shift keys.ahk`** immediately after **`#include Utils.ahk`**.

- **Timer:** **`FocusBlackoutWatcher_Tick`** every **200 ms**.
- **Debug:** set **`FOCUS_BLACKOUT_DEBUG_LOG := true`** in **`Utils.ahk`** for best-effort **`FocusBlackoutWatcher_DebugLog`** output; failures are swallowed.
- **Gate:** does nothing if **`MonitorGetCount() <= 1`**.
- **Dwell:** same foreground HWND for **`FOCUS_BLACKOUT_DWELL_MS`** (20000 ms); crossing triggers the same **3 s** countdown banner (foreground at timeout sets the keep-clear monitor).
- **Suppress (D on banner):** **`Blackout_Disable7Min()`** sets **`g_BlackoutSuppressedUntil`** for **`BLACKOUT_SUPPRESS_MS`** (7 minutes) and resets dwell; no banner until that expires, then a **new 20 s dwell** is required.
- While **`g_FocusModeOn`** and foreground monitor equals **`g_FocusModeActiveMonitor`**, no new dwell prompt.

## Multi-script behavior (critical)

AutoHotkey **does not share globals between processes**. **`FocusBlackoutWatcher_Start()`**, **`#!+Y`**, and **`#!+X`** run in **Shift keys.ahk** so they share overlay state. **`WindowManagement.ahk`** unregisters duplicate **`#!+Y`** / **`#!+X`** and uses file IPC to request disable when cycling windows on a non-keep monitor.

## Restoration and cleanup

- **`DisableFocusMode()`** stops timers, clears overlays, and **`FocusMode_ClearKeepMonitorState()`** (removes IPC files).
- **`MonitorPdfFocus`** calls **`StopPdfFocusMonitor()`** when **`g_FocusModeOn`** is false.

## Key symbols (quick reference)

| Item               | Location / value                                                                                                |
| ------------------ | --------------------------------------------------------------------------------------------------------------- |
| Anchor HWND        | `g_FocusModeAnchorHwnd`                                                                                         |
| IPC keep monitor   | `.cursor\focus_mode_keep_monitor`                                                                               |
| IPC disable request| `.cursor\focus_mode_disable_request`                                                                            |
| Dwell duration     | `FOCUS_BLACKOUT_DWELL_MS` (20000)                                                                               |
| Watcher API        | `FocusBlackoutWatcher_Start()`, `FocusBlackoutWatcher_Stop()`                                                   |
| Relocate API       | `FocusMode_SetKeepMonitor()`, `FocusMode_BuildOverlays()`                                                       |
| Cross-process      | `FocusMode_ReadKeepMonitorFromFile()`, `FocusMode_RequestDisableCrossProcess()`                                 |

## Related documentation

- Interactive banner styling: **`docs/standard_information_display.md`**

