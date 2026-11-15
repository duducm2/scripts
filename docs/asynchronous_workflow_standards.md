# Asynchronous Workflow Standards

This document defines the architectural pattern and implementation guidelines for workflows that submit tasks to an external process (e.g., Gemini in the browser) while keeping the user in their source application. It is intended for future AI maintenance and human developers extending or refactoring such flows.

---

## 1. Asynchronous Logic

### Architectural Pattern

The pattern has three distinct phases. Window focus changes only during **submission** and **retrieval**; the user remains in the source window for the entire **monitoring** phase.

| Phase                                  | Purpose                                                             | Focus behavior                                                                                                |
| -------------------------------------- | ------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------- |
| **Submission**                         | Send the task to the external process (e.g., paste prompt, submit). | Switch to external window only for the time required to submit; then return immediately to the source window. |
| **Background processing / Monitoring** | Wait for the external process to finish.                            | No window switching. User keeps working in the source application.                                            |
| **Retrieval**                          | Get the result (e.g., copy response).                               | Switch to external window once, perform all copy/verification steps, then return to the source window once.   |

**Principles:**

- **Single round-trip per phase.** Submission: one switch to external, one switch back. Retrieval: one switch to external, one switch back. No alternating focus (“ping-pong”) between source and external window.
- **Immediate return after submit.** As soon as the submit action (e.g., paste + Enter) is done, restore focus to the source window. Do not wait for the external process to complete before returning.
- **Completion-triggered retrieval.** Activate the external window again only after completion has been detected (e.g., via background monitoring). Never activate it “just to check” during the wait.

---

## 2. Context Retention (Window Handles)

### Storing Handles

- **Source (user) window:** Capture at the very start of the workflow, before any focus change:
  - `OriginalHwnd := WinExist("A")` (or equivalent).
  - Store it on a long-lived object (e.g., class instance) so it is available in submission, monitoring, and retrieval.
- **External process window:** Resolve once after you have switched to it (or identified it by title/process):
  - e.g. `GeminiHwnd := GetGeminiWindowHwnd()` and store on the same instance.
  - Use this handle for all later steps (monitoring and retrieval) so you never re-resolve by activating.

### Restoring Focus

- **After submission:** Restore to the source window using the stored handle:
  - `WinActivate("ahk_id " OriginalHwnd)` (and optionally `WinRestore` if minimized, `WinWaitActive` with a short timeout).
- **After retrieval:** Same: single `WinActivate("ahk_id " OriginalHwnd)` after all copy/verification is done.
- **No delayed “safety” restores.** Avoid timers that call `WinActivate(OriginalHwnd)` again (e.g. 400 ms later). They cause a second visible focus change and ping-pong.

---

## 3. State Monitoring (Non-Blocking Background Detection)

### Mechanism

- Use a **timer** (e.g. `SetTimer(callback, 500)`) to run a **completion check** periodically (e.g. every 500 ms). The main thread is not blocked; the user can keep using the source window.
- The callback should:
  - **Not activate any window.** It must only _observe_ the external process (e.g. via UI Automation or other APIs that work on a window by handle).
  - Determine “still running” vs “finished” (e.g. “Stop streaming” button present vs absent).
  - When “finished” is detected: stop the timer (`SetTimer(callback, 0)`), then run the **retrieval** phase (which is the only place that may activate the external window again).

### Avoiding Activation During Monitoring

- **Do not** use APIs that activate the window when they fail or when the window is in the background. For example, some browser automation libraries (e.g. `UIA_Browser`) call `WinActivate` internally when they cannot get the document or main pane.
- **Prefer** low-level UI Automation that works on a window by handle without changing focus:
  - e.g. `root := UIA.ElementFromHandle(ExternalHwnd)` then `root.FindElement(...)` / `root.ElementExist(...)`.
- If you must use a library that can activate, either:
  - Use a different code path that does not trigger activation (e.g. raw UIA from handle), or
  - Accept that monitoring may cause focus to jump and document it.

### Timeout and Cleanup

- Maintain a retry/timeout (e.g. max number of timer ticks). When exceeded, stop the timer and clean up (e.g. hide loading indicator, optionally notify the user). Do not leave the timer running indefinitely.

---

## 4. Implementation Guidelines (Avoiding Ping-Pong)

### Rules

1. **Submission phase**
   - Activate the external window only to submit (paste, send keys, etc.).
   - Restore focus to the source window **once** immediately after submit (plus optional `WinRestore` / `WinWaitActive`).
   - Do **not** start a timer that activates the source window again after a delay.

2. **Monitoring phase**
   - Do **not** call `WinActivate` (or equivalent) on either the source or the external window.
   - Do **not** use APIs that internally activate the external window (see State Monitoring above).
   - Use only observation APIs that work by window handle (e.g. `UIA.ElementFromHandle(hwnd)` and tree queries).

3. **Retrieval phase**
   - Activate the external window **once** at the start of retrieval.
   - Perform all copy and verification retries **while the external window remains active**; do not switch back to the source window between retries.
   - After all copy/verification is done, call `WinActivate("ahk_id " OriginalHwnd)` **once**, then show the result (e.g. banner).

4. **Shared helpers**
   - If a helper (e.g. “copy last message”) is used both from the hotkey and from the retrieval phase, support an option like `alreadyActive: true` so that when the caller has already activated the external window, the helper does **not** activate it again. This keeps retrieval to a single activation.

### Reference Implementation

The Win+Alt+Shift+8 pronunciation workflow in `Gemini.ahk` follows this pattern:

- **Start():** Stores `OriginalHwnd`, shows loading, activates Gemini, pastes prompt, restores focus to `OriginalHwnd` once, starts `SetTimer(CheckCompletion, 500)`.
- **CheckCompletion():** Uses `UIA.ElementFromHandle(this.GeminiHwnd)` and `root.FindElement` / `root.ElementExist` only; no `WinActivate`, no `UIA_Browser` (to avoid its internal activation fallbacks). On completion: stops timer, plays sound, calls `RetrieveResponse()`.
- **RetrieveResponse():** Activates Gemini once, runs copy (with optional retries using `alreadyActive: true`), then `WinActivate(OriginalHwnd)` once, hides loading, shows result banner.

The **Win+Alt+Shift+7** TTS-from-selection workflow (`GeminiAsyncTTS` in `Gemini.ahk`) follows the same submit → monitor → retrieve shape: **no** `WinActivate` during monitoring, and retrieval only after completion. The retrieval activation switches to Gemini for read aloud immediately after streaming finishes (plus a short DOM settle delay).

---

## 5. Suppressed Off-Screen UIA Automation (Desktop Apps & WebView2)

When automating desktop applications in the background while the user keeps working elsewhere (e.g. **Clip Angel** WinForms and **Handy** Tauri/WebView2), follow these 7 patterns to prevent focus stealing, UI stalls, false positives, and lost triggers:

1. **Show passive, non-activating feedback immediately (`0ms`)**:
   - Call `StandardLoadingBar_Show(..., { passive: true, noActivate: true })` at the very start of the hotkey handler—immediately after acquiring the automation lock, *before* DB checks, ingest waits, or window/process recovery.
   - Instant visual feedback prevents the user from assuming the shortcut was ignored and re-triggering stacked runs.
2. **Non-activating window suppression (`WS_EX_NOACTIVATE` + `ShowWindow`)**:
   - Never call `WinRestore` or `WinShow` on a suppressed window (both activate and steal focus).
   - Move the window off-screen (`-32000, -32000`), set `WinSetTransparent(0)`, apply `WS_EX_NOACTIVATE` (`+0x08000000`), and reveal with `DllCall("ShowWindow", "ptr", hwnd, "int", 4)` (`SW_SHOWNOACTIVATE`) or `8` (`SW_SHOWNA`). Remove `WS_EX_NOACTIVATE` when ending suppression.
   - Examples: `ClipAngel_BeginFavoriteSuppress` (`Utils/clip_angel_favorite.ahk`), `Handy_SuppressOffscreen` (`Utils/handy_uia_helpers.ahk`).
3. **Never yank focus when the user switches windows**:
   - Capture `priorHwnd` at start, but **only** restore `priorHwnd` if the target automation app itself currently holds foreground focus (`WinActive("ahk_exe ClipAngel.exe")` / `WinActive("ahk_exe handy.exe")`).
   - If the user intentionally Alt+Tabbed or clicked another window while the macro ran, leave them where they are (`ClipAngel_RestorePriorFocus`, `Handy_RestorePrevWindow`).
4. **WebView2 / Chromium immediate focus snap-back**:
   - Calling UIA `InvokePattern.Invoke()` or `LegacyIAccessiblePattern.DoDefaultAction()` inside WebView2/Tauri (`handy.exe`) can internally trigger `SetFocus()`.
   - Snapshot `prevActive := WinExist("A")` right before the UIA invoke and immediately restore it right after if the target app grabbed focus—do not wait until the end of a multi-second flow (`Handy_UiaInvokeBackground` in `Utils/handy_uia_helpers.ahk`).
   - Never fall back to `.Click()`, `.SetFocus()`, or foreground `Send` while suppressed; use `ValuePattern` or `ControlSend`.
5. **WinForms tray apps, popup menus, startup dialogs, and mid-flow recovery**:
   - **Hidden HWNDs:** WinForms tray apps like `ClipAngel.exe` may hide or recreate HWNDs; check with `DetectHiddenWindows true` and window identity validation (`ClipAngel_MainHwnd` / `ClipAngel_IsMainWindowIdentity` in `Utils/clip_angel_activate.ahk`) before falling back to `SendInput("!p")`.
   - **Popup menus:** `ToolStripDropDown` popup menus close immediately if the user switches windows. Never rely solely on a popup menu staying open in the background; keep toolbar / `ControlSend` fallbacks and retry with a refreshed UIA root.
   - **Modal startup dialogs:** Auto-dismiss modal `#32770` startup warnings (such as hotkey-registration dialogs for `Alt+P` / `Alt+B`) silently in the background via `WinSetTransparent(0)` + `ControlClick("Button1", ..., "NA")` / `BM_CLICK` (`ClipAngel_DismissStartupHotkeyDialogs`).
   - **Mid-flow recovery:** When recovering a stuck/windowless process mid-flow (`ClipAngel_RecoverProcessForAutomation`), pass `keepSuppressed := true` into the restart helper (`ClipAngel_RestartAttempt`) so the recovered window stays suppressed for the caller instead of being minimized in `finally`.
6. **Never trust stale pre-checks or ignore verification failures**:
   - When DB pre-checking `"already done"` (e.g. `ClipAngelDb_Newest()`), verify the DB record actually matches current state (such as `A_Clipboard`) before skipping.
   - Re-read the target clip ID *after* ingest/open, treat explicit DB verification `0` (`dbOk = false`) as an authoritative failure (`ok := false`), and never fall through to success sound/banner when `ok` is `false` (`MarkLastClipAsFavorite` in `Utils/clip_angel_favorite.ahk`).
7. **Keep background pipelines `passive: true`**:
   - Background file, HTTP, and CDP operations (`PromptPaste_BusyShow`, `pack_pipeline.ahk`, `task_import.ahk`, `import_watcher.ahk`, `toggle_outlook_teams.ahk`) must pass `passive: true` (and `noActivate: true` on non-interactive overlays) to `StandardLoadingBar_Show`.

---

## Summary

| Concern                  | Guideline                                                                                                                                                    |
| ------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| **Asynchronous logic**   | Submission → immediate return → background monitoring (timer) → retrieval only after completion.                                                             |
| **Context retention**    | Store source and external HWNDs at start; use them for all restore and targeting.                                                                            |
| **State monitoring**     | Timer-driven polling only; use APIs that do not activate the window (e.g. `UIA.ElementFromHandle` + tree search).                                            |
| **Ping-pong prevention** | One focus switch to external and one back per phase; no delayed restores; no activation during monitoring; single activation and single return in retrieval. |
| **Suppressed UIA**       | `0ms` passive/noActivate banner; `WS_EX_NOACTIVATE` + `ShowWindow(4/8)`; conditional focus restore; WebView2 immediate snap-back; authoritative DB checks.    |
