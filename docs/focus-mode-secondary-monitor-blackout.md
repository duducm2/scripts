# Focus mode and secondary-monitor blackout

This document describes how **focus mode** (black overlays on every monitor except the one hosting the active window) is triggered, cancelled, and kept consistent across scripts. Implementation lives primarily in **`Utils.ahk`**. Blackout is **manual only** via **`#!+Y`**.

## What focus mode does

- **`EnableFocusMode(keepMonitorIndex)`** creates full-screen black **Gui** overlays on **all monitors except** the monitor index passed as `keepMonitorIndex` (the “keep clear” display). Overlays use `-DPIScale`, `+AlwaysOnTop`, and **`WS_EX_TRANSPARENT`** (`+E0x20`) so pointer events pass through to windows underneath.
- **`DisableFocusMode()`** destroys those overlays, clears globals, stops the focus-mode window timer, and **`StopPdfFocusMonitor()`** so monitoring does not keep running after a manual teardown.
- **`FocusMode_SetKeepMonitor(keepMonitorIndex)`** rebuilds overlays when the keep-clear display should move (same **anchor** window moved to another monitor). Does not stop timers or call **`DisableFocusMode()`**.

Related globals: `g_FocusModeOn`, `g_FocusModeOverlays`, `g_FocusModeTrackedWindow`, `g_FocusModeActiveMonitor`, `g_FocusModeAnchorHwnd`.

Cross-process state (Shift keys owns overlays; WindowManagement can request teardown):

- **`A_ScriptDir\.cursor\focus_mode_keep_monitor`** — one-line keep-monitor index while blackout is on
- **`A_ScriptDir\.cursor\focus_mode_disable_request`** — empty sentinel; **`FocusMode_RequestDisableCrossProcess()`** creates it, **`MonitorPdfFocus`** in Shift keys consumes it

## Monitor-scoped behavior (relocate vs disable)

When **`StartPdfFocusMonitor`** is active (`#!+Y` on multi-monitor):

| Situation                                                                           | Action                                                                                    |
| ----------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------- |
| Another window on **same** keep monitor                                             | Keep blackout                                                                             |
| **Same anchor HWND** (`g_FocusModeAnchorHwnd`), foreground on **different** monitor | Relocate (`FocusMode_SetKeepMonitor`) — e.g. **`^!#a/s/d/f`** move                        |
| **Different HWND**, foreground on **different** monitor                             | **`DisableFocusMode()`** — e.g. alt-tab, click, **`^!#q/w/e/r`** cycle on another monitor |
| **`^!#q/w/e/r`** activates a window on monitor **other than** keep                  | **`FocusMode_RequestDisableCrossProcess()`** (file IPC) + monitor tick                    |

- **`Immediate`** mode (`#!+Y`): relocate on the next **200 ms** tick with no debounce.
- **`Debounced`** mode: cross-monitor **relocate** waits **`g_PdfFocusDebounceMs`** (1200 ms) for move/maximize animation (available to callers of **`StartPdfFocusMonitor`**).
- **Disable** on cross-monitor focus (non-anchor) is always immediate (no debounce).

## Entry point: Manual toggle (Win+Alt+Shift+Y)

- Hotkey **`#!+Y`** calls **`ToggleFocusMode()`**.
- **On:** `EnableFocusMode()` then **`StartPdfFocusMonitor(WinExist("A"), "Immediate")`** when **`MonitorGetCount() > 1`**.
- **Off:** `DisableFocusMode()`.
- **`ToggleFocusMode()`** treats focus mode as **on** if **`g_FocusModeOn`** **or** **`g_FocusModeOverlays.Length > 0`**.

Study Topic / QuickLook (`#!+X`) and Peek open paths do **not** start blackout. There is no dwell-based auto prompt.

## Banner suppress (D)

Some other interactive banners (e.g. Outlook reminders) offer **`[D]`** via **`Blackout_Disable7Min()`**, which sets **`g_BlackoutSuppressedUntil`** for **`BLACKOUT_SUPPRESS_MS`** (7 minutes). That suppresses those banners’ blackout-related actions for the window; it does not start focus mode.

## Multi-script behavior (critical)

AutoHotkey **does not share globals between processes**. **`#!+Y`** and **`#!+X`** run in **Shift keys.ahk** so they share overlay state. **`WindowManagement.ahk`** unregisters duplicate **`#!+Y`** / **`#!+X`** and uses file IPC to request disable when cycling windows on a non-keep monitor.

## Restoration and cleanup

- **`DisableFocusMode()`** stops timers, clears overlays, and **`FocusMode_ClearKeepMonitorState()`** (removes IPC files).
- **`MonitorPdfFocus`** calls **`StopPdfFocusMonitor()`** when **`g_FocusModeOn`** is false.

## Key symbols (quick reference)

| Item                | Location / value                                                                |
| ------------------- | ------------------------------------------------------------------------------- |
| Anchor HWND         | `g_FocusModeAnchorHwnd`                                                         |
| IPC keep monitor    | `.cursor\focus_mode_keep_monitor`                                               |
| IPC disable request | `.cursor\focus_mode_disable_request`                                            |
| Relocate API        | `FocusMode_SetKeepMonitor()`, `FocusMode_BuildOverlays()`                       |
| Cross-process       | `FocusMode_ReadKeepMonitorFromFile()`, `FocusMode_RequestDisableCrossProcess()` |
| Suppress API        | `Blackout_IsSuppressed()`, `Blackout_Disable7Min()`                             |

## Related documentation

- Interactive banner styling: **`docs/standard_information_display.md`**
