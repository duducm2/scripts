# Focus mode and secondary-monitor blackout

This document describes how **focus mode** (black overlays on every monitor except the one hosting the active window) is triggered, cancelled, and kept consistent across scripts. Core APIs live in **`Utils.ahk`** (module **`Utils/focus_mode.ahk`**). Blackout is started by **`#!+Y`** (Shift keys process) or **Wikipedia open** (AppLaunchers process). Study Topic / QuickLook (`#!+X`), Peek, and the old dwell watcher do **not** start blackout.

## What focus mode does

- **`EnableFocusMode(keepMonitorIndex)`** creates full-screen black **Gui** overlays on **all monitors except** the monitor index passed as `keepMonitorIndex` (the “keep clear” display). Overlays use `-DPIScale`, `+AlwaysOnTop`, and **`WS_EX_TRANSPARENT`** (`+E0x20`) so pointer events pass through to windows underneath. It also starts **`StartFocusModeWindowMonitor`**, which keeps **`g_FocusModeTrackedWindow`** in sync while the foreground stays on the keep-clear monitor.
- **`DisableFocusMode()`** destroys those overlays, clears globals, stops the focus-mode window timer, and **`StopPdfFocusMonitor()`** so monitoring does not keep running after teardown.
- **`FocusMode_SetKeepMonitor(keepMonitorIndex)`** rebuilds overlays when the keep-clear display should move (same **anchor** window moved to another monitor). Does not stop timers or call **`DisableFocusMode()`**.

Related globals: `g_FocusModeOn`, `g_FocusModeOverlays`, `g_FocusModeTrackedWindow`, `g_FocusModeActiveMonitor`, `g_FocusModeAnchorHwnd`.

Cross-process state (IPC under `A_ScriptDir\.cursor`; WindowManagement can request teardown):

- **`focus_mode_keep_monitor`** — one-line keep-monitor index while blackout is on
- **`focus_mode_disable_request`** — empty sentinel; **`FocusMode_RequestDisableCrossProcess()`** creates it, **`MonitorPdfFocus`** (when running) consumes it via **`FocusMode_CheckCrossProcessRequests`**

## Monitor-scoped behavior (relocate vs disable)

When **`StartPdfFocusMonitor`** is active (`#!+Y` on multi-monitor):

| Situation                                                                           | Action                                                                                    |
| ----------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------- |
| Another window on **same** keep monitor                                             | Keep blackout                                                                             |
| **Same anchor HWND** (`g_FocusModeAnchorHwnd`), foreground on **different** monitor | Relocate (`FocusMode_SetKeepMonitor`) — e.g. **`^!#a/s/d/f`** move                        |
| **Different HWND**, foreground on **different** monitor                             | **`DisableFocusMode()`** — e.g. alt-tab, click, **`^!#q/w/e/r`** cycle on another monitor |
| **`^!#q/w/e/r`** activates a window on monitor **other than** keep                  | **`FocusMode_RequestDisableCrossProcess()`** (file IPC) + monitor tick                    |

- **`Immediate`** mode (what `#!+Y` passes): relocate on the next **200 ms** tick with no debounce.
- **`Debounced`** mode: API still accepts `"Debounced"` on **`StartPdfFocusMonitor`** (waits **`g_PdfFocusDebounceMs`**, 1200 ms, before relocate). No current caller passes it.
- **Disable** on cross-monitor focus (non-anchor) is always immediate (no debounce).

Wikipedia uses a separate focus guard (**`StartWikipediaFocusMonitor`** / **`MonitorWikipediaFocus`**) that disables focus mode when Wikipedia loses foreground; it does **not** use **`StartPdfFocusMonitor`** relocate rules.

## Entry points

### 1. Manual toggle (Win+Alt+Shift+Y)

- Hotkey **`#!+Y`** calls **`ToggleFocusMode()`** (owned by the Shift keys process after other scripts unregister duplicates).
- **On:** `EnableFocusMode()` then **`StartPdfFocusMonitor(WinExist("A"), "Immediate")`** when **`MonitorGetCount() > 1`**.
- **Off:** `DisableFocusMode()`.
- **`ToggleFocusMode()`** treats focus mode as **on** if **`g_FocusModeOn`** **or** **`g_FocusModeOverlays.Length > 0`**.

### 2. Wikipedia open (AppLaunchers)

- [`AppLaunchers/wikipedia_selector.ahk`](../AppLaunchers/wikipedia_selector.ahk) and [`AppLaunchers/wikipedia_entry.ahk`](../AppLaunchers/wikipedia_entry.ahk) call **`EnableFocusMode()`** then **`StartWikipediaFocusMonitor()`** after entering fullscreen.
- Teardown: [`AppLaunchers/wikipedia_focus_guard.ahk`](../AppLaunchers/wikipedia_focus_guard.ahk) — on Wikipedia focus loss, exit fullscreen (**F11**), **`DisableFocusMode()`**, **`StopWikipediaFocusMonitor()`**.
- Overlays for this path live in the **AppLaunchers** process (separate from Shift keys `#!+Y` overlays).

Study Topic / QuickLook (`#!+X`) and Peek open paths do **not** start blackout. There is no dwell-based auto prompt.

## Banner suppress (D)

Outlook reminder interactive banners register **`D`** → **`DisableBlackout7Min`** / **`Blackout_Disable7Min()`**, which sets **`g_BlackoutSuppressedUntil`** for **`BLACKOUT_SUPPRESS_MS`** (7 minutes). While suppressed, the Outlook reminder **`ShowModal`** path shows **nothing** (the whole keys banner is skipped). The footer text does not advertise `[D]`; pressing **D** only sets the suppress window. This does not start or stop focus mode.

## Multi-script behavior (critical)

AutoHotkey **does not share globals between processes**.

- **`#!+Y`** / **`#!+X`** run in **Shift keys.ahk** so they share overlay state with Utils included there. **`WindowManagement.ahk`** unregisters duplicate **`#!+Y`** / **`#!+X`** and uses file IPC to request disable when cycling windows on a non-keep monitor.
- **Wikipedia** blackout runs in **AppLaunchers.ahk** (own `EnableFocusMode` / overlay globals). Cross-process IPC files are shared on disk, but each process only tears down overlays it owns.

## Restoration and cleanup

- **`DisableFocusMode()`** stops timers, clears overlays, and **`FocusMode_ClearKeepMonitorState()`** (removes IPC files).
- **`MonitorPdfFocus`** calls **`StopPdfFocusMonitor()`** when **`g_FocusModeOn`** is false.

## Key symbols (quick reference)

| Item                | Location / value                                                                |
| ------------------- | ------------------------------------------------------------------------------- |
| Anchor HWND         | `g_FocusModeAnchorHwnd`                                                         |
| Same-monitor track  | `StartFocusModeWindowMonitor()`, `FocusModeWindowMonitor`                       |
| IPC keep monitor    | `.cursor\focus_mode_keep_monitor`                                               |
| IPC disable request | `.cursor\focus_mode_disable_request`                                            |
| Relocate API        | `FocusMode_SetKeepMonitor()`, `FocusMode_BuildOverlays()`                       |
| Cross-process       | `FocusMode_ReadKeepMonitorFromFile()`, `FocusMode_RequestDisableCrossProcess()` |
| Wikipedia guard     | `StartWikipediaFocusMonitor()`, `MonitorWikipediaFocus`                         |
| Suppress API        | `Blackout_IsSuppressed()`, `Blackout_Disable7Min()`                             |

## Related documentation

- Interactive banner styling: **`docs/standard_information_display.md`**
- Repo-local sentinels overview: **`docs/lightweight-api-sentinel-files.md`**
