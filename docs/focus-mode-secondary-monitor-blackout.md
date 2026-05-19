# Focus mode and secondary-monitor blackout

This document describes how **focus mode** (black overlays on every monitor except the one hosting the active window) is triggered, cancelled, and kept consistent across scripts. Implementation lives primarily in **`Utils.ahk`**; the **automatic dwell watcher** is started from **`Shift keys.ahk`**.

## What focus mode does

- **`EnableFocusMode(keepMonitorIndex)`** creates full-screen black **Gui** overlays on **all monitors except** the monitor index passed as `keepMonitorIndex` (the “keep clear” display). Overlays use `-DPIScale`, `+AlwaysOnTop`, and **`WS_EX_TRANSPARENT`** (`+E0x20`) so pointer events pass through to windows underneath.
- **`DisableFocusMode()`** destroys those overlays, clears globals, stops the focus-mode window timer, and **`StopPdfFocusMonitor()`** so auto-relocate monitoring does not keep running after a manual teardown.
- **`FocusMode_SetKeepMonitor(keepMonitorIndex)`** rebuilds overlays when the keep-clear display should move (used by **`MonitorPdfFocus`** on cross-monitor focus). Does not stop timers or call **`DisableFocusMode()`**.

Related globals: `g_FocusModeOn`, `g_FocusModeOverlays`, `g_FocusModeTrackedWindow`, `g_FocusModeActiveMonitor`.

## Monitor-scoped auto behavior

When **`StartPdfFocusMonitor`** is active (after Study Topic / dwell countdown, Peek, etc.):

- **Same keep monitor:** switching to another window on the same display **keeps** blackout; **`g_FocusModeTrackedWindow`** / **`g_PdfFocusTrackedHwnd`** follow the foreground HWND.
- **Different monitor:** **`MonitorPdfFocus`** calls **`FocusMode_SetKeepMonitor`** so the clear display moves to the monitor containing the foreground window; other monitors get black overlays again.
- **`Debounced`** mode (default on **`StudyTopic_ApplyBlackoutCountdownTimeout`** when not overridden): cross-monitor focus must persist for **`g_PdfFocusDebounceMs`** (1200 ms) before relocate (reduces flicker during move/maximize).
- **`Immediate`** mode (dwell / countdown paths): relocate on the next **200 ms** tick with no debounce.

**`^!#a/s/d/f`** in **`WindowManagement.ahk`** moves the focused window to another monitor; relocation is picked up by **`MonitorPdfFocus`** in **Shift keys** (~200 ms after the move) — no extra IPC in WindowManagement.

Manual **`#!+Y`** does **not** start **`StartPdfFocusMonitor`**; blackout stays on the monitor chosen at toggle until **`#!+Y`** again (does not follow focus).

## Entry points

### 1. Manual toggle (Win+Alt+Shift+Y)

- Hotkey **`#!+Y`** calls **`ToggleFocusMode()`**.
- **On:** `EnableFocusMode()` with no index uses **`GetActiveMonitorIndex()`** from the foreground window center vs **`MonitorGet`** bounds (same rule as “which display stays clear”).
- **Off:** `DisableFocusMode()`.
- **`ToggleFocusMode()`** treats focus mode as **on** if **`g_FocusModeOn`** **or** **`g_FocusModeOverlays.Length > 0`** (does not rely solely on **`WinExist`** on overlay HWNDs, which can fail on some DPI setups).

### 2. Study Topic / QuickLook (Win+Alt+Shift+X)

When QuickLook is active, **`#!+X`** runs **`StudyTopic_StartBlackoutCountdown(hwnd)`**, which shows the standard interactive banner (see **`docs/standard_information_display.md`**) and on timeout calls **`StudyTopic_ApplyBlackoutCountdownTimeout`**.

- **`StudyTopic_ApplyBlackoutCountdownTimeout`** activates the target window, computes **`keepIdx := GetActiveMonitorIndex()`** with fallback **`StudyTopic_GetBlackoutKeepMonitorIndex()`** (primary monitor), then **`EnableFocusMode(keepIdx)`** and **`StartPdfFocusMonitor(targetHwnd, pdfFocusLossMode)`**.
- Countdown callers pass **`pdfFocusLossMode := "Immediate"`** so cross-monitor relocate happens on the next **`MonitorPdfFocus`** tick without debounce.
- **`MonitorPdfFocus`** is monitor-scoped (see above); it does not disable blackout when another window on the same monitor gains focus.

### 3. Automatic dwell watcher (20 seconds)

Started once via **`FocusBlackoutWatcher_Start()`** from **`Shift keys.ahk`** immediately after **`#include Utils.ahk`**.

- **Timer:** **`FocusBlackoutWatcher_Tick`** every **200 ms**.
- **Gate:** does nothing if **`MonitorGetCount() <= 1`** (no secondary monitors).
- **Dwell:** same foreground HWND for **`FOCUS_BLACKOUT_DWELL_MS`** (20000 ms); crossing triggers **`FocusBlackoutWatcher_StartCountdown(hwnd)`**.
- **Banner:** same **`StandardLoadingBar_ShowWithKeys`** contract as **`StudyTopic_StartBlackoutCountdown`** (3 s, **[N] Cancel**, progress + track monitor + preserve focus).
- **Cancel (N):** sets a **deny** HWND so the 20 s dwell does not restart until the user focuses another window and returns.
- **Timeout:** **`StudyTopic_ApplyBlackoutCountdownTimeout(hwnd, "Immediate")`** so **`StartPdfFocusMonitor(..., "Immediate")`** relocates blackout when foreground moves to another monitor.
- While **`g_FocusModeOn`** and **`GetActiveMonitorIndex()`** equals **`g_FocusModeActiveMonitor`**, the watcher does not accumulate another dwell (avoids stacking prompts during an active blackout on that display).

## Multi-script behavior (critical)

AutoHotkey **does not share globals between processes**. Each running script that **`#include`s `Utils.ahk`** gets its **own** copy of `g_FocusModeOn`, `g_FocusModeOverlays`, etc.

Therefore:

- **`FocusBlackoutWatcher_Start()`** runs only in **`Shift keys.ahk`** so the **dwell timer**, **`EnableFocusMode`** from the watcher, **`#!+Y`**, and **`#!+X`** (Study Topic / QuickLook blackout) run in the **same process** and see the same state.
- Secondary launcher scripts unregister duplicate Utils hotkeys after including Utils:

  `try Hotkey("#!+Y", "Off")`  
  `try Hotkey("#!+X", "Off")`

  **Implemented in:** `AppLaunchers.ahk`, `Gemini.ahk`, `Outlook.ahk`, `Microsoft Teams.ahk`, `WindowManagement.ahk`.

If those hotkeys were left registered in a script that did not apply the current blackout (for example **AppLaunchers**), that instance would see **empty** globals: **`#!+Y`** could call **`EnableFocusMode()`** instead of **`DisableFocusMode()`**, and **`#!+X`** could start **`StudyTopic_StartBlackoutCountdown`** in the wrong process so **`#!+Y`** in **Shift keys** would still not clear overlays owned elsewhere.

## Restoration and cleanup

- **`DisableFocusMode()`** ends with **`StopPdfFocusMonitor()`** so manually clearing blackout (**`#!+Y`**) does not leave **`MonitorPdfFocus`** running.
- **`MonitorPdfFocus`** calls **`StopPdfFocusMonitor()`** when **`g_FocusModeOn`** is false (cleanup only).

## Key symbols (quick reference)

| Item               | Location / value                                                                                                |
| ------------------ | --------------------------------------------------------------------------------------------------------------- |
| Dwell duration     | `FOCUS_BLACKOUT_DWELL_MS` (20000) in `Utils.ahk`                                                                |
| Countdown duration | `STUDY_TOPIC_BLACKOUT_DELAY_MS` (3000)                                                                          |
| Watcher API        | `FocusBlackoutWatcher_Start()`, `FocusBlackoutWatcher_Stop()`                                                   |
| Relocate API       | `FocusMode_SetKeepMonitor()`, `FocusMode_BuildOverlays()`                                                       |
| Banner pattern     | `StudyTopic_StartBlackoutCountdown` / `FocusBlackoutWatcher_StartCountdown` → `StandardLoadingBar_ShowWithKeys` |

## Related documentation

- Interactive banner styling and **`StandardLoadingBar_ShowWithKeys`**: **`docs/standard_information_display.md`**
