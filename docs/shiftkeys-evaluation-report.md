# ShiftKeys Evaluation Report

**Purpose:** Read-only analysis of the Shift keys AutoHotkey script to identify performance bottlenecks, redundant logic, and inefficient UI interaction sequences. No modifications were made to the original file.

**Scope:** [Shift keys.ahk](../Shift%20keys.ahk) (15,194 lines) — System-wide Win+Alt+Shift symbol-layer shortcuts, app-specific hotkeys (Chrome, Outlook, Teams, Power BI, ClipAngel, ChatGPT, Gemini, Wikipedia, Mercado Livre, Gmail, Excel, Spotify, etc.), and UIA-based automation.

---

## 1. Executive Summary

The script provides a large set of context-dependent hotkeys and automations across many applications. It uses UI Automation (UIA) for Chrome, Outlook, Teams, Power BI, ClipAngel, ChatGPT, Gemini, and others; delegates loading/overlay UI to [Utils.ahk](../Utils.ahk); and includes helpers for button/list waiting and discovery.

**Strengths:** Broad coverage, consistent use of `#HotIf` for context, reuse of `FocusOutlookField` and some shared UIA patterns, and integration with the standard loading bar. **Main improvement areas:** Synchronous blocking monitors on Enter/^Enter (Gemini) that tie up the hotkey thread; repeated full-tree UIA scans in tight polling (`WaitForButton`, Gemini/ChatGPT monitors); duplicated hotkey bodies (e.g. Outlook +S/+T/+B, Ctrl+1..5 ladders, Power BI FindFirst fallback chains); expensive `#HotIf` predicates (e.g. `GetChatGPTWindowHwnd()` doing full window list + title scan); and many fixed `Sleep` sequences where condition-based waits would improve responsiveness and reliability. Addressing these would improve shortcut latency, reduce CPU/UIA load, and make the script easier to maintain.

---

## 2. Performance Bottlenecks

### 2.1 Synchronous blocking monitors on Enter / Ctrl+Enter (Gemini)

`Enter::` and `^Enter::` (lines 13769–13793) call `WaitForStopResponseButton_Gemini()` directly. That function (lines 13798–13920+) runs nested `while` loops: it waits for the "Stop response" button to appear, then monitors until it disappears, then runs a 1.5 s confirmation loop with `Sleep 300` and repeated `FindFirst` checks. All of this runs on the hotkey thread.

**Impact: High.** The thread is blocked for the full duration of the AI response (up to `timeout` default 300,000 ms). Input can feel laggy or unresponsive; quick key repeats or other hotkeys may be delayed.

**Recommendation:** Move monitoring to a timer-driven pattern (e.g. `SetTimer` every 250–500 ms) that checks UIA state and plays the chime when the button disappears. The Enter hotkey would submit, schedule the timer, and return immediately.

### 2.2 WaitForButton full-tree scan in tight polling

`WaitForButton(root, pattern, timeout)` (lines 2373–2505) runs a loop until `deadline`, each iteration:

- Calling `root.FindAll({ Type: "Button" })` (full descendant tree),
- Iterating all buttons and running `RegExMatch(btn.Name, pattern)` plus property/pattern checks,
- Then `Sleep 50` if no match.

In applications with many buttons (e.g. Chrome document, Power BI), each cycle is expensive. Debug logging (`SafeDebugLog`) runs multiple times per iteration when enabled, adding file I/O to the hot path.

**Impact: High.** CPU and UIA load scale with tree size and polling frequency; with debug on, disk I/O compounds latency.

**Recommendation:** Increase poll interval (e.g. 100–150 ms) for this helper; consider a single combined condition or cached button list per context where safe. Gate `SafeDebugLog` behind a global `DEBUG_ENABLED` (or similar) and avoid logging in the innermost loop.

### 2.3 WaitForButtonAndShowSmallLoading_ChatGPT blocking loop

`WaitForButtonAndShowSmallLoading_ChatGPT` (lines 15123–15193) uses a `while` loop with `Sleep 250`, and for each tick tries multiple button names via `FindElement`. When a button is found, it enters another `while` that polls every 250 ms with the same `FindElement` attempts until the button disappears. All on the calling thread.

**Impact: High.** Same as 2.1: long-running synchronous wait on the hotkey/flow thread.

**Recommendation:** Prefer a timer-based monitor (like Gemini.ahk’s completion detection) so the hotkey returns after submitting and the UI/banner updates asynchronously.

### 2.4 Repeated full FindAll / FindFirst ladders in Gemini and Chrome flows

In the Gemini fullscreen and other Chrome/Gemini blocks (e.g. lines 13720–13765, 6428–6508), the same pattern appears repeatedly:

- `FindFirst` by Name + Type 50000,
- Fallback `FindFirst` with Type `"Button"`,
- Fallback `FindAll({ Type: 50000 })` then iterate and filter by ClassName/Name.

Multiple hotkeys in the same context each perform their own full tree walk and multi-step fallback. No shared cache or single discovery per flow.

**Impact: High.** Redundant UIA work when the user triggers several such shortcuts in sequence; latency scales with tree size.

**Recommendation:** Extract shared helpers (e.g. `FindGeminiButtonByNamesAndClass(uia, names, classNeedles)`) and, where appropriate, reuse one discovery per logical operation or cache by window/context with short TTL.

### 2.5 Power BI drawer discovery repeated per config

`PowerBI_FindDrawerButton(root, config)` (lines 8201–8254) is called in loops over drawer configs (e.g. lines 7250, 8033, 8079). For each config it may:

- Try several names × type variants (`"Button"`, `50000`) with `FindFirst`,
- Then try `classNames` with `FindFirst`,
- Then, for `classContains`, call `root.FindAll({ Type: "Button" })` or `FindAll({ Type: 50000 })` and iterate all buttons.

So for N configs, the script can do N full button scans in the worst case.

**Impact: Medium–High.** Cost multiplies with number of drawer configs and tree size.

**Recommendation:** Perform one `FindAll` for buttons (or by type) per scope (e.g. per hotkey or per Power BI action), pass the collection into a helper that filters by config (names/class/classContains), so each button scan is reused across configs.

### 2.6 Expensive #HotIf predicates

- `#HotIf (hwnd := GetChatGPTWindowHwnd()) && WinActive("ahk_id " hwnd)` (line 6429): `GetChatGPTWindowHwnd()` (lines 64–70) enumerates all Chrome windows with `WinGetList("ahk_exe chrome.exe")` and checks each title with `InStr(..., "chatgpt", false)`. This runs whenever the hotkey system evaluates this context.
- Similar cost applies to predicates that use `IsChromePdfViewerActive()`, `IsMercadoLivreActive()`, `IsTeamsMeetingActive()`, `IsOutlookMainActive()`, etc., if they perform UIA or multiple window calls.

**Impact: High** for perceived hotkey latency when many hotkeys are defined and the engine repeatedly evaluates these conditions.

**Recommendation:** Prefer cheap predicates (e.g. window class/title only). For expensive checks (UIA, full window list), cache result with a short TTL (e.g. 300–500 ms) updated by a low-frequency timer, or restrict the expensive predicate to a small set of hotkeys.

### 2.7 Debug logging in hot paths

`SafeDebugLog` (lines 31–60) uses `FileAppend` with retry and backoff. It is invoked from:

- `WaitForButton` (multiple times per iteration: entry, button count, name checks, match found, return, timeout),
- Spotify hotkey logic (multiple branches).

When debug is enabled, every call does file I/O on the hot path.

**Impact: High** when debugging is on; otherwise low.

**Recommendation:** Guard all `SafeDebugLog` calls with a global (e.g. `DEBUG_SHIFTKEYS`) so that in production no file I/O runs. Optionally sample (e.g. log at most once per 300 ms per flow) to reduce volume when debugging.

---

## 3. Redundant or Inefficient Logic

### 3.1 Outlook +S, +T, +B duplicated across main and message inspector

The same focus logic appears in two `#HotIf` blocks:

- **IsOutlookMainActive()** (lines 4517–4546): `+S::` / `+T::` / `+B::` call `FocusOutlookField` with the same AutomationIds and fallbacks (Subject 4101, To/Required 4109/4117, etc.).
- **IsOutlookMessageActive()** (lines 4891–4923): `+S::` / `+T::` / `+B::` repeat the same `FocusOutlookField` sequences.

**Recommendation:** Define shared handlers (e.g. `Outlook_FocusSubject()`, `Outlook_FocusTo()`, `Outlook_FocusBodyFromSubject()`) that call `FocusOutlookField` with the appropriate criteria; bind these in both `#HotIf` blocks so logic lives in one place.

### 3.2 Ctrl+1..5 ladders repeated in multiple contexts

The pattern “send N×{Down} then {Enter}” appears for Ctrl+1..5 in at least:

- ClipAngel (lines 1699–1750),
- Outlook message inspector Command Palette (lines 4925–4963),

and similar “move down then select” patterns for Alt+2..5 in ClipAngel. The bodies are almost identical except for the number of `{Down}` and `{Enter}`.

**Recommendation:** One helper, e.g. `SelectNthItemByDownEnter(n)` (and optionally a variant for Alt+number), called from small hotkey wrappers in each context to remove duplication and simplify tuning.

### 3.3 Power BI view/tab hotkeys repeated FindFirst fallback pattern

`+t`, `+u`, `+i`, `+o`, `+p`, `+h` (lines 6953–7127) each:

- Get `root := UIA.ElementFromHandle(win)` (same for all),
- Run a multi-step FindFirst chain (by Name, Type 50019/50000, matchmode Substring, ClassName, etc.),
- Click or send keys; on failure show `MsgBox`.

The structure is repeated with only the target name and type varying. Similar repetition in other Power BI shortcuts.

**Recommendation:** Extract a single helper, e.g. `PowerBI_FindAndClickTab(root, names, typeOptions)` or `PowerBI_FindViewByName(root, nameVariants)`, and use it from each hotkey to centralize fallback order and error handling.

### 3.4 ClipAngel filter selector and cleanup state duplication

`ShowClipAngelFilterSelector()` and `CleanupClipAngelFilterSelector()` (lines 1970+, 1917–1966) both manage:

- `g_ClipAngelFilterSelectorActive`,
- Dynamic hotkey registration/unregistration (including case variants and vkBC/vkBE),
- GUI lifecycle and map resets.

The attach/detach logic is mirrored in two places, which increases the risk of drift (e.g. missing a key in one path).

**Recommendation:** Centralize “attach ClipAngel filter hotkeys” and “detach ClipAngel filter hotkeys” in two helpers, and have Show/Cleanup call them plus set state so registration logic exists once.

### 3.5 PadShortcut does not pad

`PadShortcut(shortcut, targetWidth := 24)` (lines 79–83) returns `shortcut` unchanged; `targetWidth` is unused. Call sites (e.g. `ProcessCheatSheetText`) imply alignment/formatting intent.

**Recommendation:** Either implement padding/alignment so the cheat sheet output is consistent, or remove `PadShortcut` and the parameter from call sites to avoid misleading abstraction.

### 3.6 ChatGPT RenameChatGPTWindowToChatGPT long linear sequence

`RenameChatGPTWindowToChatGPT()` (lines 6056–6254) performs a long sequence of steps with many `FindElement`/`FindFirst` attempts and fixed `Sleep` (500, 500, 1000, 300, etc.). Sidebar, chat button, sibling, expand, conversation options, rename button, etc. are each resolved with nested try/catch and name lists. No shared “find button by names” helper.

**Recommendation:** Extract small helpers (e.g. “find sidebar close button”, “find chat list button”, “find conversation options button”) and use them in this flow. Replace fixed sleeps with short condition waits (e.g. “wait until element exists” with timeout and poll interval) where the sleep is clearly waiting for UI state.

---

## 4. Synchronous Blocking and UI Interaction Issues

### 4.1 Fixed Sleep chains in ClipAngel and Wikipedia

- **NavigateClipAngelComboBox** (lines 1756–1862): Long chain of `Sleep` (30, 50, 100, 150, 200 ms) after activation, UIA access, click, SetFocus, Home, navigation keys, Tab, F10, Ctrl+Home. Total delay is fixed regardless of actual UI speed.
- **Wikipedia restore/save** (e.g. lines 3148–3201): `BlockInput("On")`, then `Sleep(500)`, JS execution, `Sleep(500)`, banner update, `Sleep(1000)`, then cleanup. BlockInput plus long sleeps can make the session feel unresponsive.

**Impact: Medium.** On fast machines the user waits longer than necessary; on slow ones the script may still act before the UI is ready.

**Recommendation:** Where a Sleep is clearly waiting for UI state (e.g. menu open, element focusable), replace with a short loop that checks the condition (e.g. element exists or has focus) with a timeout and poll interval (e.g. 50–100 ms), and break early when the condition is met.

### 4.2 BlockInput and modal MsgBox in flows

- Wikipedia scroll restore (lines 3155–3188) turns `BlockInput("On")` for the whole operation and can leave it on if an error path is hit before `BlockInput("Off")` (some paths do call it).
- Many hotkeys use `MsgBox` on failure (e.g. “Could not find…”, “Error…”). Modal dialogs block the script until the user dismisses them and can interrupt quick repeated use of shortcuts.

**Recommendation:** Ensure every path that enables BlockInput has a corresponding Off (e.g. in a finally-like pattern or single exit path). Consider non-modal feedback (e.g. toast or banner) for “not found” or non-fatal errors so the user can keep working without dismissing a dialog.

### 4.3 WinWaitActive / WinWaitClose in automation paths

Several flows use `WinWaitActive(..., 1)` or similar (e.g. line 1761, 2538). If the window does not become active within the timeout, the script continues; elsewhere, synchronous `WinWait*` can hold the thread. Usage is acceptable for one-off automation but adds to total blocking time.

**Recommendation:** Keep timeouts short; where multiple steps depend on window state, consider documenting the assumed order and total wait so future changes do not inadvertently extend blocking.

---

## 5. Ambiguous or Risky Patterns

### 5.1 Empty or silent catch blocks

- `PowerBI_FindDrawerButton` (lines 8250–8251): `catch Error { }` with no comment or log.
- `PowerBI_AttemptCollapse` (lines 5285–5286): same.
- Many UIA `try/catch` blocks (e.g. in ChatGPT rename, Gemini fullscreen, Teams) swallow errors with no indication of why or whether it was expected.

**Recommendation:** Add a one-line comment per catch documenting the expected failure (e.g. “element not found”, “pattern not supported”). Optionally, in debug mode, call a single lightweight logger so failures are visible during development.

### 5.2 Overlapping #HotIf for same key in Chrome

Google Keep and other Chrome-scoped blocks define `+s` (and possibly others). If `#HotIf WinActive("... Keep ...")` and `#HotIf ... InStr(title, "Google")` both apply (e.g. Keep tab title contains “Google”), resolution depends on order of definition and can be confusing.

**Recommendation:** Make Chrome sub-contexts mutually exclusive (e.g. “Google” block explicitly excludes Keep/YouTube/Gemini by title or URL), or use a single predicate router (e.g. `GetChromePageType()`) that returns Keep | Search | Gemini | … and bind keys per page type.

### 5.3 Magic numbers and repeated literals

- Many `Sleep` values (50, 100, 150, 200, 250, 300, 500, 1000, 1500, 5000) appear without named constants.
- UIA AutomationIds (e.g. "8346", "4226", "4356", "4101", "4109", "4117") and key codes (e.g. "vkBC", "vkBE") are hardcoded in multiple places.
- Hold threshold 700 (e.g. for #!+a) is a magic number.

**Recommendation:** Group constants at the top or in a small config section (e.g. `TIMING_*`, `OUTLOOK_IDS_*`, `CLIPANGEL_KEYS_*`, `HOTKEY_HOLD_MS`) so behavior is auditable and tunable without searching the file.

### 5.4 UIA Type: number vs string

Script uses both numeric (e.g. `Type: 50000`, `Type: 50019`) and string (e.g. `Type: "Button"`, `Type: "50000"`) in different places. Consistency with [Gemini.ahk](../Gemini.ahk) and UIA docs would reduce ambiguity; prefer numeric and document at top (e.g. `; UIA ControlType: Button=50000, MenuItem=50011, TabItem=50019`).

---

## 6. Shortcut and Flow Efficiency

- Hotkeys are well scoped by `#HotIf` (app/context), so overlap is limited to the Chrome sub-contexts noted above.
- Efficiency gains will come mainly from: (1) making Enter/^Enter and ChatGPT “wait for button” flows non-blocking (timer-based), (2) reducing repeated UIA tree walks via shared helpers and one scan per scope where possible, (3) cheapening or caching `#HotIf` predicates, and (4) replacing fixed Sleeps with condition-based waits where the intent is “wait for UI state.”

---

## 7. Positive Aspects

- **Delegation to Utils:** Loading bar, overlays, and standard patterns are delegated to [Utils.ahk](../Utils.ahk); no ad-hoc GUI in Shift keys.ahk for those.
- **Consistent context use:** `#HotIf` is used systematically so hotkeys apply only in the right app/window.
- **Reuse of FocusOutlookField:** Outlook field focus is centralized in one function; only the binding is duplicated across main/message inspector.
- **EN/PT and locale:** Many flows support both English and Portuguese (e.g. button names, ChatGPT “Stop streaming” / “Interromper transmissão”), improving robustness across locales.
- **Structured sections:** Section headers and comments (e.g. ClipAngel, Power BI, Outlook, Gemini) make the 15k-line file navigable.

---

## 8. Prioritized Recommendations (quick wins first)

1. **Gate debug logging:** Add a global (e.g. `DEBUG_SHIFTKEYS`) and wrap all `SafeDebugLog` calls so that in production no file I/O runs in hot paths. **Low effort, high impact when debug is on.**
2. **Make Gemini Enter/^Enter non-blocking:** Replace direct `WaitForStopResponseButton_Gemini()` call with “submit + SetTimer-based monitor” so the hotkey returns immediately and the chime runs when the stop button disappears. **Medium effort, high impact on responsiveness.**
3. **Extract shared UIA helpers:** Introduce helpers for “find button by names/class” and “find Gemini/ChatGPT control” and reuse them in Chrome/Gemini/ChatGPT blocks to cut repeated FindAll/FindFirst ladders. **Medium effort, high impact on latency and maintainability.**
4. **Centralize Outlook +S/+T/+B and ClipAngel hotkey logic:** Single implementation for Outlook field focus and for Ctrl+1..5 / Alt+2..5 so both `#HotIf` blocks only bind to shared functions. **Low–medium effort, reduces duplication and drift.**
5. **Power BI: one FindAll per scope:** In drawer and view hotkeys, do one `FindAll` (or by type) per action and pass the collection into helpers that filter by config (names/classContains). **Medium effort, reduces redundant scans.**
6. **Cheapen or cache expensive #HotIf:** Cache `GetChatGPTWindowHwnd()` (and similar) with short TTL or restrict expensive predicates to fewer hotkeys so evaluation cost is lower. **Medium effort, improves hotkey evaluation time.**
7. **Replace fixed Sleep with condition waits:** In NavigateClipAngelComboBox, Wikipedia restore, and ChatGPT rename, replace Sleeps that clearly wait for UI state with a short “wait until condition or timeout” loop (e.g. 50–100 ms poll). **Medium effort, improves responsiveness and reliability.**
8. **Name constants:** Introduce named constants for Sleep durations, UIA IDs, and key codes (and document UIA types at top). **Low effort, improves clarity and tuning.**
9. **Document or fix PadShortcut:** Either implement alignment/padding or remove the helper and its parameter so call sites match behavior. **Low effort.**
10. **Comment empty catches and overlapping Chrome keys:** One-line comments in empty catch blocks; clarify Chrome sub-context order or make conditions mutually exclusive. **Low effort.**

---

_Report generated from read-only analysis of Shift keys.ahk. No changes were made to the script._
