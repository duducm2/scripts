# AppLaunchers Evaluation Report

**Purpose:** Read-only analysis of the AppLaunchers AutoHotkey script to identify performance bottlenecks, structural flaws, logic issues, and optimization opportunities. No modifications were made to the original file.

**Scope:** [AppLaunchers.ahk](../AppLaunchers.ahk) (2,024 lines) — Application and website launcher hotkeys (Cursor, Desktop/Explorer, Chrome, WhatsApp, YouTube, Gmail, Cursor activate, Wikipedia selector with scroll save/restore, Pomodoro timer, centered banner and key-combo helpers).

---

## 1. Executive Summary

The script provides launcher hotkeys and Wikipedia/Pomodoro flows. It uses UIA for Chrome address-bar focus and Wikipedia URL/scroll (via `UIA_Browser` and clipboard-based JS execution), delegates overlays and loading bar to [Utils.ahk](../Utils.ahk), and includes helpers for monitor detection and scroll-position persistence.

**Strengths:** Clear sectioning, use of Utils for banners/overlays, EN/PT support for Desktop title, and structured Wikipedia selector with completed-articles history. **Main improvement areas:** Repeated `WinGetList` + title scans in Cursor activation (#!+n); fixed polling loops (Desktop window wait, Wikipedia page-ready and UIA retries) with many `Sleep` calls; long synchronous Wikipedia scroll-restore path with `BlockInput` and heavy Sleep chains; timer-driven Wikipedia focus monitor at 200 ms; and duplicated activation + delay patterns across launchers. Addressing these would improve shortcut latency and make timing more robust across hardware.

---

## 2. Performance Bottlenecks

### 2.1 Cursor window enumeration and title scan (#!+n)

Lines 53–82: Two `WinGetList` calls (one for `ahk_exe Cursor.exe`, one for `ahk_exe Code.exe`). For each process, every HWND is iterated and `WinGetTitle("ahk_id " hwnd)` is called. Title is lowercased and checked with multiple `InStr` calls for "preview", "habits", "home", "punctual", "work". No shared cache; each hotkey press does full enumeration.

**Impact: High.** Cost scales with number of Cursor/Code windows. On hotkey evaluation this path runs synchronously and can add tens of milliseconds.

**Recommendation:** Reuse a single enumeration (e.g. one `WinGetList("ahk_exe Cursor.exe")` and optionally merge Code.exe if needed), or delegate to a cached list updated by a low-frequency timer or daemon (see [windowmanagement-improvements.md](windowmanagement-improvements.md)). Prefer one pass to build target + fallback.

### 2.2 Desktop/Explorer activation polling loop (+#e)

Lines 173–185: If no Desktop window exists, script runs `loop 40` with `Sleep 50` (up to 2 s), each iteration calling `WinExist("Área de Trabalho ahk_class CabinetWClass")` and `WinExist("Desktop ahk_class CabinetWClass")`. After that, activation uses `WinWaitActive(..., 0.2)` then DllCall escalation (SwitchToThisWindow, SetForegroundWindow). Fixed `Sleep 350` and `Sleep 100` follow (lines 209–211).

**Impact: Medium.** Polling is bounded but blocks for up to 2 s on cold start; duplicate `WinExist` checks (e.g. 164–167) repeat the same title/class checks.

**Recommendation:** Use a single `WinWait` with a short timeout and title match mode for either localized title, or a single loop that exits on first successful `WinExist` to avoid redundant calls. Replace fixed 350/100 ms sleeps with shorter delays or condition checks where possible.

### 2.3 Wikipedia HandleWikipediaChar: long Sleep and retry chains

Lines 756–1010: After launching Chrome for a Wikipedia URL, the script uses `WinWait("ahk_exe chrome.exe", , 5)`, `Sleep(500)`, `WinWait("Wikipedia", , 10)`, then a page-ready loop (8 retries) that each time can call `GetWikipediaURL()`, `UIA_Browser("ahk_exe chrome.exe")`, and `JSReturnThroughClipboard("document.documentElement.scrollHeight")`. Multiple fixed sleeps follow (500, 1000, 800, 300, 800, 2500, 600, 400, 300, 1000, 300 ms, etc.). Scroll restoration then runs another UIA retry loop (5–8 retries) with further Sleeps (800, 2500, 400, 600, 300, 1000, 300). `BlockInput("On")` is used during restore without a guaranteed single exit path in all error branches.

**Impact: High.** Total blocking time can exceed 15–20 seconds on a single shortcut. UIA and clipboard-based JS are invoked many times; any failure path that does not call `BlockInput("Off")` can leave input blocked.

**Recommendation:** Shorten or replace fixed sleeps with condition-based waits (e.g. “document ready” or “scrollHeight stable” with a single poll interval and timeout). Ensure every path that enables `BlockInput("On")` has a corresponding `BlockInput("Off")` (e.g. single exit or finally-like pattern). Consider reducing retry counts and consolidating “page ready” and “UIA ready” into one staged wait.

### 2.4 Wikipedia #!+k scroll-restore path (existing window)

Lines 1296–1402: When Wikipedia already exists, the script runs a similar UIA + scroll-restore sequence: `BlockInput("On")`, loop 3 for UIA init with `Sleep(200)`, loop 5 for doc height with `Sleep(200)`/`Sleep(500)`, verification loop with `Sleep(200)`, then `Sleep(300)`/`Sleep(200)`. Same risk of BlockInput not turned off on every path.

**Impact: Medium–High.** Shorter than the new-window path but still several seconds of blocking and multiple UIA/clipboard calls.

**Recommendation:** Same as 2.3: condition-based waits, single exit for BlockInput, and reduced redundant polls.

### 2.5 Wikipedia focus monitor timer (200 ms)

Lines 420–435, 428: `StartWikipediaFocusMonitor()` sets `SetTimer(g_WikipediaFocusMonitorTimer, 200)`. `MonitorWikipediaFocus()` (376–418) runs every 200 ms when the selector has opened a Wikipedia window; it checks `WinActive("Wikipedia")` and `WinExist("Wikipedia")`, and if focus was lost it activates the window, `Sleep(50)`, sends F11, and calls `DisableFocusMode()` / `StopWikipediaFocusMonitor()`.

**Impact: Medium.** When the Wikipedia flow is active, a 200 ms timer adds steady background work and can prevent low-power state on laptops.

**Recommendation:** Consider a longer interval (e.g. 400–500 ms) or event-driven focus detection (e.g. SetWinEventHook for foreground change) so the script only reacts when the active window actually changes.

### 2.6 GetWikipediaURL and UIA_Browser per call

Lines 490–512: `GetWikipediaURL()` uses `WinActive("ahk_exe chrome.exe")`, `WinGetTitle("A")`, then `UIA_Browser("ahk_exe chrome.exe")` and `uia.GetCurrentURL()`. Called from `HandleWikipediaChar` (e.g. in the page-ready loop), from #!+k (1292), and from scroll/save logic. Each call constructs UIA and reads URL; no caching.

**Impact: Medium.** Repeated UIA construction and URL read in tight loops multiplies cost.

**Recommendation:** Where the same Chrome/Wikipedia window is assumed for a multi-step flow, obtain UIA once and reuse for URL and scroll operations within that flow.

---

## 3. Redundant or Inefficient Logic

### 3.1 Duplicate WinExist checks for Desktop (+#e)

Lines 164–167: `if WinExist("Área de Trabalho ahk_class CabinetWClass")` then `targetHwnd := WinExist(...)` (same string again). Same for "Desktop ahk_class CabinetWClass". The same pattern repeats inside the loop (175–182). Each `WinExist` re-enumerates or re-evaluates.

**Impact: Low–Medium.** Unnecessary duplicate work on every run.

**Recommendation:** Store the result of one `WinExist` in a variable and reuse it for both the condition and assignment.

### 3.2 Cursor #!+n: repeated target/fallback activation block

Lines 84–107: The “target window” and “fallback window” branches are almost identical: `WinExist` check, `WinActivate`, `WinWaitActive(..., 2)`, `CenterMouse()`, `Sleep(100)`, `Send("^t")`. Only the window spec differs.

**Impact: Low.** Duplication increases maintenance and chance of drift.

**Recommendation:** Extract a single helper (e.g. `ActivateCursorWindowAndSendCtrlT(winSpec)`) and call it for both target and fallback.

### 3.3 IsWindowOnMonitor3: full monitor loop

Lines 446–476: `IsWindowOnMonitor3()` gets active window rect, computes center, then `loop monitorCount` with `MonitorGet(A_Index, ...)` until the point is inside a monitor, then returns whether that index is 3. Correct but does a full iteration over all monitors every call.

**Impact: Low.** Monitor count is small; cost is modest but could be cached if the same “active window” is checked repeatedly in one flow.

**Recommendation:** Optional: cache “last hwnd + monitor index” and invalidate on focus change; otherwise leave as-is.

### 3.4 SafeDebugLog defined but never used

Lines 19–37: `SafeDebugLog(text)` is implemented with retry and exponential backoff but is never called in AppLaunchers.ahk. Dead code; if ever adopted in hot paths without a global gate, it could add file I/O.

**Impact: Low.** No current cost; future risk if logging is added without a debug flag.

**Recommendation:** Either remove it or add a global (e.g. `DEBUG_APPLAUNCHERS`) and gate any future calls so production runs no file I/O.

---

## 4. Synchronous Blocking and UI Interaction Issues

### 4.1 Fixed Sleep chains in launcher hotkeys

- **#!+n** (91–94, 103–106): `WinWaitActive(..., 2)` then `Sleep(100)` before Send.
- **+#e** (209–211): `Sleep 350`, `Send "^{Up}"`, `Sleep 100`, `Send "{F5}"`.
- **#!+f** (226–227): `WinWait(..., 10)`, `Sleep(300)`.
- **#!+,** (330): `WinWaitActive("ahk_exe Cursor.exe", , 10)` (up to 10 s block).
- **#!+z, #!+h, #!+w**: `WinWaitActive` with no timeout (indefinite wait until window is active).

**Impact: Medium.** Fixed delays can be too long on fast machines or too short on slow ones; unbounded `WinWaitActive` can hang if the window never activates.

**Recommendation:** Use bounded timeouts for all `WinWait*` (e.g. 5–10 s max) and show user feedback on timeout. Replace fixed Sleeps that wait for “window ready” with short condition loops (e.g. “window visible and not minimized”) where feasible.

### 4.2 BlockInput and error paths in Wikipedia restore

In `HandleWikipediaChar` (e.g. 792–1009) and in #!+k scroll restore (1302–1395), `BlockInput("On")` is used. Several error paths call `BlockInput("Off")` and hide loading bar, but nested try/catch and early returns could leave BlockInput on if a rare error path is hit.

**Impact: High** (if input remains blocked); otherwise **Medium**.

**Recommendation:** Use a single exit pattern or ensure a single cleanup function runs on both success and failure (e.g. set a flag and call `BlockInput("Off")` and `StandardLoadingBar_Hide` at one place after try/catch).

### 4.3 CenterMouse and #!+. fixed delays

`CenterMouse()` (1949–1952): `Sleep(200)` then `Send("#!+q")`. #!+. (1958–1969): `Sleep(100)`, `Send("^c")`, `Sleep(200)`, `Send("!v")`, `Sleep(700)`, `Send("!q")`, `Sleep(200)`, `SendEscape()`. Total fixed delay over 1.4 s.

**Impact: Medium.** User waits for the full sequence regardless of UI responsiveness.

**Recommendation:** Document the intended UX (e.g. “allow paste dialog to open”); if any of these Sleeps are waiting for a visible state, consider a short condition wait with timeout instead of a fixed 700 ms.

---

## 5. Ambiguous or Risky Patterns

### 5.1 Silent catch blocks

Multiple `catch { }` or `catch { ; ... }` blocks (e.g. 76–78, 251–252, 298–299, 366–367, 718–724, 798–800, 865–866) swallow errors with no logging or comment. In UIA and URL/scroll logic this can hide real failures.

**Recommendation:** Add a one-line comment per catch describing the expected failure (e.g. “window closed”, “UIA not ready”). Optionally call a gated debug logger so failures are visible when debugging.

### 5.2 Unbounded WinWaitActive in launchers

#!+z (273), #!+h (291), #!+w (311): After `Run`, the script uses `WinWaitActive("WhatsApp")`, `WinWaitActive("YouTube")`, `WinWaitActive("Gmail ahk_exe chrome.exe")` with no timeout. If the window never becomes active (e.g. login required, crash), the script blocks indefinitely.

**Recommendation:** Use a timeout (e.g. 10–15 s) and on timeout show a non-modal message (e.g. overlay) so the user is informed and the script does not hang.

### 5.3 Magic numbers and literals

Sleep values (50, 100, 200, 300, 350, 500, 800, 1000, 2500, etc.) and retry counts (3, 5, 8, 40) are hardcoded throughout. UIA Type 50004 and AutomationId "view_1012" (238) are literals. Monitor index 3 is hardcoded in `IsWindowOnMonitor3()`.

**Recommendation:** Introduce named constants at the top (e.g. `WM_DESKTOP_POLL_MS`, `WIKI_PAGE_READY_RETRIES`, `WIKI_UIA_RETRIES`, `CHROME_ADDRBAR_AUTOMATION_ID`) for tuning and clarity.

### 5.4 Portrait/orientation and “new window” comments

Wikipedia scroll logic (e.g. 883–886, 891–892, 948) includes comments about “portrait orientation” and “new windows” requiring longer waits. Logic is tailored to Monitor 3 and new Chrome windows; behavior on other setups may differ.

**Recommendation:** Document assumptions (monitor index, orientation, new vs existing window) in one place so future changes do not break edge cases.

---

## 6. Positive Aspects

- **Utils.ahk delegation:** Overlays, loading bar, and centered notifications use [Utils.ahk](../Utils.ahk) (e.g. `ShowCenteredOverlay_Utils`, `StandardLoadingBar_Show`, `ClipAngelBanner_Show`), keeping GUI logic centralized.
- **EN/PT support:** Desktop launcher checks both "Área de Trabalho" and "Desktop" for Explorer window title.
- **Wikipedia structure:** Clear separation of selector GUI, character handlers, scroll save/restore, and focus monitor; completed-articles history and CSV/INI persistence are organized.
- **Pomodoro:** Timer and chime use `SetTimer` with negative for one-shot; logging is optional (work environment suppressed).
- **Section headers:** Comment blocks and hotkey labels make the file navigable.

---

## 7. Prioritized Recommendations (quick wins first)

1. **Bounded WinWaitActive:** Add timeouts (e.g. 10 s) to #!+z, #!+h, #!+w `WinWaitActive` and show a short overlay on timeout. **Low effort, prevents indefinite block.**
2. **BlockInput safety:** Ensure every Wikipedia path that calls `BlockInput("On")` has a single exit or finally-like path that calls `BlockInput("Off")` and hides the loading bar. **Low–medium effort, high impact on reliability.**
3. **Duplicate WinExist in +#e:** Store result of one `WinExist` per title and reuse; avoid calling `WinExist` twice for the same spec. **Low effort.**
4. **Extract Cursor activation helper:** Single function for “activate this Cursor window and Send ^t” used for both target and fallback in #!+n. **Low effort, reduces duplication.**
5. **Wikipedia: condition-based waits:** Replace long fixed Sleep chains in HandleWikipediaChar and #!+k scroll restore with short loops that check “page ready” or “UIA ready” with a single poll interval and max wait. **Medium effort, improves responsiveness.**
6. **Reduce Wikipedia timer frequency:** Consider 400–500 ms for `MonitorWikipediaFocus` when active, or document the 200 ms choice. **Low effort.**
7. **Reuse UIA in Wikipedia flow:** In multi-step Wikipedia operations, obtain `UIA_Browser` once per logical flow and reuse for URL and scroll instead of creating it in every helper call. **Medium effort.**
8. **Named constants:** Introduce constants for Sleep durations, retry counts, and UIA identifiers (and document monitor index 3 assumption). **Low effort.**
9. **Comment empty catches:** Add one-line comments in catch blocks describing expected failures. **Low effort.**
10. **SafeDebugLog:** Remove if unused, or add `DEBUG_APPLAUNCHERS` and gate any future logging. **Low effort.**

---

_Report generated from read-only analysis of AppLaunchers.ahk. No changes were made to the script._
