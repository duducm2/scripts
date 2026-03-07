# WindowManagement Evaluation Report

**Purpose:** Read-only analysis of the Window Management AutoHotkey script to identify performance bottlenecks, redundant logic, structural issues, and optimization opportunities. No modifications were made to the original file.

**Scope:** [WindowManagement.ahk](../WindowManagement.ahk) (2,652 lines) — Window positioning (monitor move/cycle/minimize/close), project quick selector (Win+Alt+Shift+L), Cursor window and preview-window activation, Copy-from-Gemini-to-Cursor bridge integration, and related hotkeys/timers.

**Reference standard:** This report applies the same evaluation categories and style as [shiftkeys-evaluation-report.md](shiftkeys-evaluation-report.md).

---

## 1. Executive Summary

The script centralizes window-management hotkeys (monitor ordering, cycle/minimize/close per monitor, Alt+Tab variants) and a project quick-selector flow that opens or focuses Cursor by project path, with sub-modes for selection, preview windows, and Copy-from-Gemini. It delegates notifications and overlay UI to [Utils.ahk](../Utils.ahk) and includes [GeminiToCursorBridge.ahk](../GeminiToCursorBridge.ahk) for paste targeting.

**Strengths:** Clear sectioning and comments; reuse of `ShowNotification_WM` and `TryActivateWindow_WM`; consistent use of `GetMonitorIndexByOrder` and monitor work-area logic; keyboard-only focus flow for Cursor AI field. **Main improvement areas:** Full `WinGetList()` scans and O(n²)-style visibility/sorting in `GetVisibleWindowsOnMonitor`; repeated Cursor-window enumeration and title/segment matching across `GetCursorHwndForProject`, `FindAndActivateCursorWindow`, `FindAndActivatePreviewWindow`, and `HandlePreviewWindowSelection`; fixed `Sleep` chains in `ActivateCursorProject` and `FocusCursorAITextField`; duplicated character-to-project assignment and hotkey enable/disable logic across project selector, selection mode, Copy-from-Gemini mode, and Cursor window sub-menu; a 100 ms `MonitorActiveWindow` timer running for the full session; and many silent `catch` blocks that hide failures. Addressing these would reduce latency on selector and activation paths, lower background CPU use, and make debugging and future changes easier.

---

## 2. Performance Bottlenecks

### 2.1 GetVisibleWindowsOnMonitor — full enumeration and O(n²) visibility + sort

`GetVisibleWindowsOnMonitor(mon)` (lines 510–604) does the following on every call:

- Calls `WinGetList()` with no filter, enumerating all top-level windows.
- For each window: multiple Win32 calls (`WinGetMinMax`, `IsWindowVisible`, `GetWindowLongPtr`, `MonitorFromWindow`, `GetWindowRect`), plus `WinGetClass` and `WinGetTitle`.
- For each candidate, an inner loop over already-accepted `visible` entries checks whether the window center is covered by a higher z-order window (lines 563–571), yielding O(n²) behavior when many windows are on the same monitor.
- A bubble sort (lines 585–598) re-orders the visible list by Y then X.

**Impact: High.** Cost scales with total window count and with the number of visible windows on the target monitor. Called by `CycleWindowsOnMonitor`, `MinimizeWindowOnMonitor`, and `CloseWindowOnMonitor`, so repeated use (e.g. cycling or closing on a busy monitor) amplifies the cost.

**Recommendation:** Consider filtering by process or window class before building the full list where possible; cache the visible list per monitor with a short TTL or invalidate on focus/visibility changes; replace the inner “covered by” loop with a more efficient structure (e.g. sort by z-order once and accept in that order with a single pass); or use a lighter heuristic (e.g. topmost N windows on the monitor) if exact visibility is not required.

### 2.2 HandlePreviewWindowSelection — nested Cursor × projects × segments and fallback rescan

`HandlePreviewWindowSelection` (lines 1545–1731) runs a heavy matching pipeline:

- Loops over `WinGetList("ahk_exe Cursor.exe")` and for each window gets the title.
- For each preview window (title contains "preview"), loops over all `g_Projects`, resolving path and calling `ExtractProjectMatchSegments(projectPath)` and then matching segments (exact, "(Workspace)" suffix, and last-word variants) against the title.
- If no matches are found, a second full `WinGetList("ahk_exe Cursor.exe")` loop (lines 1666–1692) rescans to add any preview window with an extracted workspace name.

So in the worst case: (Cursor windows × projects × segments) plus a full second enumeration.

**Impact: High.** This is likely the slowest interactive path when many Cursor and preview windows and projects exist. User waits on the order of hundreds of milliseconds to over a second.

**Recommendation:** Reuse a single Cursor window list for both “match to projects” and “unmatched preview” passes; factor “preview window ↔ project” matching into a shared helper used also by `FindAndActivatePreviewWindow` and `GetCursorHwndForProject`; consider caching `ExtractProjectMatchSegments` per project path for the duration of the selector session.

### 2.3 ShowCursorWindowSelectorSubMenu — per-window project matching and duplicated assignment logic

`ShowCursorWindowSelectorSubMenu` (lines 1881–2224):

- Gets `WinGetList("ahk_exe Cursor.exe")` once (good).
- Builds `projectIndexToChar` and `projectIndexToCategory` with the same category/character-assignment logic as `ShowProjectSelector` (lines 1911–1964), i.e. duplicated logic.
- For each Cursor window, calls nested helper `GetMatchingProjectIndex(winTitle)` (lines 1980–2005), which loops all projects and for each calls `WindowMatchesProject(winTitle, projectPath)`, which in turn calls `ExtractProjectMatchSegments(projectPath)` and iterates segments. So cost is (windows × projects × segments).
- Two passes over `windowsWithKeys` assign keys to unmatched windows (lines 1944–2073).

**Impact: High.** With many Cursor windows and projects, the matching and two-pass assignment add noticeable delay when opening the sub-menu.

**Recommendation:** Extract the “project index to character” and “category ordering” logic into a single shared function used by both `ShowProjectSelector` and `ShowCursorWindowSelectorSubMenu`. Consider building a “window title → project index” map once from the Cursor window list and reusing it instead of calling `GetMatchingProjectIndex` per window.

### 2.4 ActivateCursorProject startup polling loop

`ActivateCursorProject(projectPath)` (lines 733–794) calls `FindAndActivateCursorWindow(projectPath)` first; if that returns 0, it launches Cursor with `Run`, then:

- Loops up to 30 times with `Sleep 200` per iteration (lines 756–761).
- Each iteration calls `GetCursorHwndForProject(projectPath)`, which does a full `WinGetList("ahk_exe Cursor.exe")` and for each hwnd gets the title and runs segment matching.

So after launch, the script can perform up to 30 full Cursor window enumerations and title/segment checks (up to 6 seconds of wall time), all on the hotkey thread.

**Impact: Medium–High.** Users perceive a fixed 200 ms × N delay before the new window is detected; on slow machines the window may appear after the loop has already given up.

**Recommendation:** Replace fixed 200 ms sleeps with a bounded wait that polls at a short interval (e.g. 100–150 ms) and exits as soon as `GetCursorHwndForProject` returns non-zero, with a clear maximum total wait (e.g. 8–10 s). Optionally, use a single shared “get Cursor windows + match to path” helper so both “find existing” and “wait for new” paths share one enumeration/matching implementation.

### 2.5 FocusCursorAITextField — fixed Sleep chain

`FocusCursorAITextField(targetHwnd)` (lines 1129–1150) does:

- `WinActivate` and `WinWaitActive(..., 2)` (acceptable).
- `Sleep 200` (line 1140).
- `Send "^i"` to open AI panel.
- `Sleep 1200` (line 1143).
- `Send "{Tab 2}"` and `Sleep 100` (lines 1144–1145).

Total fixed delay is 1.5 s plus the `WinWaitActive` timeout when the window is not yet active. The 1200 ms sleep is intended for the AI panel to open; it is not conditioned on UI state.

**Impact: Medium.** On fast machines the user waits longer than necessary; on slow ones the panel might not be ready before Tab is sent. Adds at least 1.5 s to every “activate project and focus AI” flow.

**Recommendation:** Replace the 1200 ms (and optionally the 200 ms) sleep with a short loop that checks for a simple “AI input visible” or “focus in input” condition with timeout and poll interval (e.g. 50–100 ms), or use a single longer timeout with early exit when a condition is met. Keep a capped maximum wait for reliability.

### 2.6 MonitorActiveWindow timer — 10×/s for session lifetime

`SetTimer MonitorActiveWindow, 100` (line 53) runs `MonitorActiveWindow` (lines 208–243) every 100 ms. The callback gets the current foreground hwnd, compares it to the previous one, and if it changed (and the change was not immediately after a mouse click), calls `MoveMouseToCenter(hwnd)` and optional cursor flash. So every 100 ms the script does at least `WinExist("A")`, process checks, and possibly `GetWindowRect` and mouse move.

**Impact: Medium.** Per tick the work is small, but the timer runs for the entire session. On battery or under load, reducing the frequency (e.g. 200–250 ms) or only running the timer when the project selector (or another dependent feature) is active would cut background CPU.

**Recommendation:** Consider increasing the interval to 200 ms, or enabling this timer only when a mode that needs “center cursor on focus change” is active, and turning it off when that mode is closed.

### 2.7 Repeated Cursor window enumeration and segment matching

The following functions each enumerate Cursor windows and match titles to project path or segments:

- `GetCursorHwndForProject(projectPath)` (966–985): `WinGetList("ahk_exe Cursor.exe")`, then for each window title, iterate `ExtractProjectMatchSegments(projectPath)` and segment match.
- `FindAndActivateCursorWindow(projectPath)` (989–1051): Same pattern; builds a list, then activates by z-order or active match.
- `FindAndActivatePreviewWindow(projectPath)` (874–963): `WinGetList("ahk_exe Cursor.exe")`, filter by "preview" in title, then for each window loop match segments (and workspace name) against projects.
- `HandlePreviewWindowSelection` (1545–1731): As in 2.2, Cursor windows × projects × segments, plus fallback rescan.

There is no shared “get Cursor windows + match to one project path” or “get preview windows + match to projects” helper. Sequential flows (e.g. find existing then activate) can each do a full enumeration and matching pass.

**Impact: High.** Redundant work when multiple of these run in one user action (e.g. project select that calls `FindAndActivateCursorWindow`, or preview that calls both matching and activation logic). Latency and CPU scale with window and project count.

**Recommendation:** Introduce shared helpers, e.g. `GetCursorWindowsForProject(projectPath)` returning a list of matching hwnds/titles, and `GetPreviewWindowsWithProjectMatches()` returning a list of preview windows with optional project index. Reuse in `GetCursorHwndForProject`, `FindAndActivateCursorWindow`, `FindAndActivatePreviewWindow`, and `HandlePreviewWindowSelection` so enumeration and matching run once per logical operation.

### 2.8 Debug logging in hot paths

`_DebugLog_WM(loc, msg, data, hypothesisId)` (lines 19–26) builds a JSON-like string and appends it to a file. It is called from:

- `ActivateCursorProject` (entry, after FindAndActivate, after FocusCursorAITextField, return true/false).

When debug is enabled, every project activation does file I/O. The JSON string is built with string concatenation and does not escape quotes or special characters in `data`, so malformed or truncated log lines are possible.

**Impact: Medium** when debugging is on; **Low** otherwise.

**Recommendation:** Gate all `_DebugLog_WM` calls behind a global (e.g. `DEBUG_WINDOWMANAGEMENT`) so production runs do no file I/O. Optionally escape or sanitize `data` (or use a proper JSON library) to keep logs parseable.

---

## 3. Redundant or Inefficient Logic

### 3.1 Character-to-project assignment duplicated in four places

The same pipeline appears in:

- **ShowProjectSelector** (lines 2311–2365): Build `projectIndexToCategory`, then for each category collect project indices, assign `g_ProjectCharSequence` chars (skipping "3"), fill `projectIndexToChar`.
- **HandleSelectionModeTrigger** (lines 1234–1287): Same: `projectIndexToCategory`, then by category assign chars (skip "3"), fill `projectIndexToChar`.
- **HandleCopyFromGeminiModeTrigger** (similar block): Same character-assignment pattern.
- **ShowCursorWindowSelectorSubMenu** (lines 1911–1964): Same again for `projectIndexToChar` and `projectIndexToCategory`.

Only the use of the map differs (display vs. hotkey binding vs. window-to-key mapping). The assignment rules (category order, skip "3", sequential char index) are duplicated in four places, which risks drift and complicates changes (e.g. adding another reserved key).

**Recommendation:** Extract a single function, e.g. `BuildProjectIndexToCharMap()` returning `projectIndexToChar` (and optionally `projectIndexToCategory`), and call it from all four call sites. Reserve character "3" (and any others) inside that function.

### 3.2 Hotkey enable/disable boilerplate repeated

Enabling and disabling hotkeys with the same rules (comma/period via vkBC/vkBE, uppercase for a–z) appears in:

- **CleanupProjectSelector** (lines 830–856): Loop `g_ProjectHotkeyHandlers`, disable by char (vkBC/vkBE, else Hotkey char and StrUpper).
- **CleanupSelectionMode** and **CleanupCopyFromGeminiMode**: Similar loops over their handler arrays with the same VK and uppercase logic.
- **ShowProjectSelector** (lines 2468–2491): Enable hotkeys for `projectIndexToChar` with the same comma/period/uppercase handling.
- **HandleSelectionModeTrigger** and **HandleCopyFromGeminiModeTrigger**: Enable mode-specific hotkeys with the same pattern.

So “turn off a set of char-based hotkeys” and “turn on char-based hotkeys from a map” are each implemented multiple times.

**Recommendation:** Add two helpers: e.g. `DisableHotkeysForHandlers(handlersArray)` and `EnableHotkeysFromProjectIndexToChar(projectIndexToChar, handlerFactory)` that encapsulate vkBC/vkBE and uppercase. Use them in cleanup and in each mode’s setup so registration logic lives in one place.

### 3.3 Monitor work area and “center on active window’s monitor” repeated

- **ShowProjectSelector** (lines 2256–2296): Gets active window, `GetWindowRect`, computes center, then loops `MonitorGetCount()` and `MonitorGetWorkArea` to find the monitor containing that center; sets `monitorLeft/Top/Right/Bottom` and dimensions.
- **ShowCursorWindowSelectorSubMenu** and other GUIs could use the same “which monitor is the active window on?” logic for positioning.

Currently only the project selector does this; if more GUIs are added, the pattern may be duplicated.

**Recommendation:** Extract a function, e.g. `GetMonitorWorkAreaForActiveWindow()` or `GetMonitorBoundsContainingWindow(hwnd)`, returning work area and dimensions, and use it in `ShowProjectSelector` and any future selector GUIs.

### 3.4 Activation + WinWaitActive + MoveMouseToCenter pattern

The sequence “activate window by hwnd, WinWaitActive, then MoveMouseToCenter” appears in:

- `FindAndActivateCursorWindow` (1043–1047).
- `FindAndActivatePreviewWindow` (872–878, 932–936).
- `HandlePreviewWindowSelection` (1724–1726, 1732–1735).
- Single-window branch in `ShowCursorWindowSelectorSubMenu` (1901–1903).

**Recommendation:** Add a small helper, e.g. `ActivateWindowAndCenterMouse(hwnd, timeoutSeconds := 2)` that does `WinActivate`, `WinWaitActive`, and `MoveMouseToCenter`, and use it everywhere this sequence appears to avoid drift and to centralize timeout handling.

---

## 4. Synchronous Blocking and UI Interaction Issues

### 4.1 Fixed Sleep in activation and focus paths

- **AltTab(count)** (line 181): `Sleep 250` after sending Alt+Tab to “allow the window to activate.” No condition check.
- **MoveWinToMonitor** (lines 382, 404): `Sleep 100` after DllCall move, then `Sleep 150` “to allow window animation to finish.”
- **CycleWindowsOnMonitor** (507): `Sleep 100` after activation “for animation/focus stability.”
- **MinimizeWindowOnMonitor** / **CloseWindowOnMonitor** (631, 664): `Sleep 100` after activate before minimize/close.
- **ActivateCursorProject** (756–761): Loop with `Sleep 200` per iteration (see 2.4).
- **FocusCursorAITextField** (1140, 1143, 1145): `Sleep 200`, `Sleep 1200`, `Sleep 100` (see 2.5).
- **ShowProjectSelector** (2247): `Sleep 50` after `CleanupProjectSelector` if reopening.
- **HandlePreviewWindowSelection** (1557): `Sleep 100` after cleanup.

**Impact: Medium.** Total blocking time adds up on the hotkey thread; fixed delays may be too long on fast machines or too short on slow ones.

**Recommendation:** Where the intent is “wait for window/UI state,” replace fixed sleeps with a short loop that checks a condition (e.g. window active, or a simple visibility/focus check) with timeout and poll interval (e.g. 50–100 ms). Keep a maximum total wait for safety. Document or name constants for any remaining unavoidable delays.

### 4.2 WinWaitActive usage

`WinWaitActive("ahk_id " hwnd, , timeout)` is used in several places (e.g. 504, 769, 953, 1045, 1133, 1138, 1760, 1904). Timeouts are short (0.3–3 s), which is good. If the window does not become active, execution continues after the timeout; in some paths the next step (e.g. Send or focus) may then apply to the wrong window.

**Recommendation:** After any `WinWaitActive` that times out, consider checking that the active window is still the expected one before sending keys or calling `FocusCursorAITextField`; optionally show a short notification if activation failed so the user is aware.

---

## 5. Ambiguous or Risky Patterns

### 5.1 Silent catch blocks

There are many `catch { }` or `catch { ; comment }` blocks (e.g. 43, 214, 235, 315, 366, 426, 439, 447, 577, 752, 770, 783, 846, 854, 865, 916, 921, 943, 956, 978, 982, 1014, 1019, 1038, 1048, 1147, 1229, 1312, 1341, 1368, 1409, 1481, 1539, 1653, 1658). Errors (e.g. window closed, access denied, hotkey already in use) are swallowed with no log or user feedback, making failures hard to diagnose.

**Recommendation:** Add a one-line comment in each catch stating the expected failure (e.g. “window no longer valid,” “hotkey conflict”). Optionally, when a debug flag is set, call a single lightweight logger so that during development these failures are visible without changing behavior in production.

### 5.2 HandleCursorWindowSelection closes all non-target Cursor windows

`HandleCursorWindowSelection(targetHwnd, allCursorWindows)` (lines 1738–1760) iterates `allCursorWindows` and calls `WinClose` on every hwnd that is not the target. So when the user picks one Cursor window from the sub-menu, all other Cursor windows are closed without confirmation.

**Impact: Medium.** If the user triggers the selector by accident or misunderstands the effect, they can lose many editor windows. The behavior is intentional but aggressive.

**Recommendation:** Document this clearly in the UI (e.g. “Selecting a window will close all other Cursor windows”) or consider adding a confirmation step (“Close N other Cursor windows?”) when more than one other window would be closed. Optionally make this behavior configurable.

### 5.3 Manual JSON in \_DebugLog_WM

`_DebugLog_WM` (19–26) builds a string with `'{"location":"' . loc . '"...'` and does not escape quotes or newlines in `loc`, `msg`, `data`, or `hypothesisId`. If any of these contain `"` or `\`, the written line is no longer valid JSON and may break line-based NDJSON parsing.

**Recommendation:** Escape double quotes and backslashes in string fields, or use a small JSON/NDJSON helper so each log line is valid. Gate logging behind a debug flag as in 2.8.

### 5.4 Magic numbers and literals

- Timer interval `100` (line 53) for `MonitorActiveWindow`.
- Sleep values: 50, 100, 150, 200, 250, 300, 1200 ms in various places.
- `WinWaitActive` timeouts: 0.3, 1, 2, 3 s.
- Loop cap 30 in `ActivateCursorProject` (line 756).
- Tolerance `TOL := 40` in `GetVisibleWindowsOnMonitor` (line 522).

**Recommendation:** Introduce a small config block or constants (e.g. `WM_TIMER_INTERVAL_MS`, `WM_ACTIVATION_WAIT_MS`, `WM_AI_PANEL_WAIT_MS`, `WM_LAUNCH_POLL_MAX`, `WM_VISIBILITY_ROW_TOLERANCE`) so behavior is auditable and tunable in one place.

### 5.5 Key binding overlap in project selector

Commands [c], [3], [L], [K] are explicitly bound (lines 2494–2531). The character-assignment loop can also assign 'c', '3', 'l', 'k' to projects (they appear in `g_ProjectCharSequence`). The order of registration (project hotkeys first, then command hotkeys) means the command keys take precedence for those characters, but the behavior is not obvious from the code and could change if registration order or `g_ProjectCharSequence` changes.

**Recommendation:** Document that [c], [3], [L], [K] are reserved and must not be assigned to projects, or exclude them in `BuildProjectIndexToCharMap` so the mapping and the UI stay in sync and the intent is explicit.

---

## 6. Positive Aspects

- **Delegation to Utils:** Notifications and overlay (e.g. `ShowCenteredOverlay_Utils`) are delegated; no ad-hoc GUI for toasts.
- **Consistent monitor model:** `GetMonitorIndexByOrder` and left-to-right ordering are used consistently for MEH+A/S/D/F and related hotkeys.
- **Structured sections:** Clear headers and comments (e.g. Project Quick Selector, Cursor Window Selection, Copy from Gemini mode) make the file navigable.
- **Environment awareness:** `IS_WORK_ENVIRONMENT` and project `path` vs. `workPath` are used consistently for paths and Cursor executable.
- **Self-contained focus path:** `FocusCursorAITextField` uses only keyboard (^i, Tab) and no UIA dependency, reducing coupling and failure modes.
- **GeminiToCursorBridge:** Copy-from-Gemini flow is isolated in an included module, keeping window-management logic separate.

---

## 7. Prioritized Recommendations (quick wins first)

1. **Gate debug logging:** Add a global (e.g. `DEBUG_WINDOWMANAGEMENT`) and wrap all `_DebugLog_WM` calls so production runs do no file I/O. **Low effort, medium impact when debug is on.**
2. **Name constants:** Introduce named constants for timer interval, Sleep durations, WinWait timeouts, and loop caps (see 5.4). **Low effort, improves clarity and tuning.**
3. **Extract project-index-to-character map:** Single `BuildProjectIndexToCharMap()` (and optional category map) used by ShowProjectSelector, HandleSelectionModeTrigger, HandleCopyFromGeminiModeTrigger, and ShowCursorWindowSelectorSubMenu. **Medium effort, reduces duplication and drift.**
4. **Extract hotkey enable/disable helpers:** `DisableHotkeysForHandlers` and `EnableHotkeysFromProjectIndexToChar` (or equivalent) used in all cleanup and setup paths. **Medium effort, centralizes VK and uppercase handling.**
5. **Shared Cursor window matching:** Introduce `GetCursorWindowsForProject(projectPath)` and reuse in `GetCursorHwndForProject`, `FindAndActivateCursorWindow`, and (where applicable) preview/path matching so enumeration and segment matching run once per logical operation. **Medium effort, high impact on latency.**
6. **Replace ActivateCursorProject polling:** Bounded wait with shorter poll interval and exit as soon as `GetCursorHwndForProject` returns non-zero; cap total wait. **Low–medium effort, improves perceived launch time.**
7. **Replace FocusCursorAITextField fixed sleeps:** Condition-based wait (e.g. “AI input visible” or focus check) with timeout and poll interval instead of 200 + 1200 + 100 ms. **Medium effort, improves responsiveness.**
8. **Reduce MonitorActiveWindow frequency or scope:** Increase interval to 200 ms or run timer only when a dependent mode is active. **Low effort, lowers background CPU.**
9. **GetVisibleWindowsOnMonitor:** Avoid O(n²) “covered by” pass (e.g. sort by z-order and accept in one pass); consider short-TTL cache or lighter heuristic. **Medium effort, high impact on cycle/minimize/close on busy monitors.**
10. **HandlePreviewWindowSelection:** Single Cursor list for both matched and unmatched passes; reuse shared “preview ↔ project” matching; consider caching segment extraction. **Medium effort, high impact on preview path.**
11. **Comment or log empty catches:** One-line comment per catch documenting expected failure; optional debug-only log. **Low effort, improves diagnosability.**
12. **ActivateWindowAndCenterMouse helper:** Centralize “WinActivate + WinWaitActive + MoveMouseToCenter” and use everywhere. **Low effort, avoids drift and clarifies timeout behavior.**

---

_Report generated from read-only analysis of WindowManagement.ahk. No changes were made to the script._
