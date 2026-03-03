# Standard Loading Bar Rollout – Identified Locations

This document lists all locations across the main codebase that are suitable for the standard loading bar (`StandardLoadingBar_Show`, `StandardLoadingBar_Update`, `StandardLoadingBar_Hide` from `Utils.ahk`). It covers long-running shortcut flows and current banner/overlay usage.

**Canonical API (in Utils.ahk):**

- `StandardLoadingBar_Show(state := "Working...", barColor := "3772FF")`
- `StandardLoadingBar_Update(state := "", barColor := "")`
- `StandardLoadingBar_Hide(delayMs := 0)` — use `delayMs > 0` to show a final message briefly before hiding.

**Lifecycle:** Call `Show` at start, `Update` at milestones, `Hide` in all exit paths (including `try/finally` and error branches).

---

## 1. Shift keys.ahk

| Location                  | Entry point                                                                                  | Why long-running                                            | Current feedback                       | Loading bar insertion                                                          | Priority | Status                                                  |
| ------------------------- | -------------------------------------------------------------------------------------------- | ----------------------------------------------------------- | -------------------------------------- | ------------------------------------------------------------------------------ | -------- | ------------------------------------------------------- |
| Fold Explorer             | `FoldAllDirectoriesInExplorer()` (e.g. `^,`)                                                 | UIA tree traversal, collapse loops, retries                 | Progress overlay (text + bar)          | Show at start, Update "Directories folded", Hide(800)                          | High     | **Done** – uses `StandardLoadingBar_*`                  |
| Unfold Explorer           | `UnfoldAllDirectoriesInExplorer()` (e.g. `^q`)                                               | UIA tree traversal, expand loops, retries                   | Progress overlay                       | Show at start, Update "Directories unfolded", Hide(800)                        | High     | **Done** – uses `StandardLoadingBar_*`                  |
| Wikipedia restore (Shift) | `RestoreWikipediaScrollPosition(...)` (internal)                                             | UIA_Browser, JS clipboard, BlockInput, Sleep                | `CreateCenteredBanner_ChatGPT`         | Show(bannerText), Update "Scroll position restored!", Hide(500/0) on all paths | High     | **Candidate** – replace banner with standard bar        |
| Wikipedia save (Shift)    | `SaveWikipediaScrollPositionManually_ShiftKeys()` (e.g. `+p`)                                | Exit fullscreen, UIA, scroll read, INI write, re-fullscreen | `CreateCenteredBanner_ChatGPT`         | Show "Saving...", Update "Scroll position saved!" / error, Hide in finally     | High     | **Candidate** – replace banner with standard bar        |
| Wikipedia “no position”   | Multiple branches setting "No previous scroll position found"                                | Short message only                                          | `CreateCenteredBanner_ChatGPT` (red)   | Optional: Show + Hide(duration) for consistency                                | Low      | **Candidate**                                           |
| Marp export               | `^6::` (Marp export flow)                                                                    | Save dialog, filename read, replace prompt                  | `ShowSmallLoadingIndicator_ChatGPT`    | Show "Exporting with Marp...", Update milestones, Hide in finally              | High     | **Candidate** – replace small loading with standard bar |
| ChatGPT / OpenAI flows    | `ShowSmallLoadingIndicator_ChatGPT` / `HideSmallLoadingIndicator_ChatGPT` at many call sites | WaitForButton, UIA, API-style flows                         | Small loading indicator (banner-style) | Replace with StandardLoadingBar_Show/Update/Hide at each call site             | Medium   | **Candidate** – many call sites (see grep)              |
| WaitForButton call chains | Spotify connect, WhatsApp voice, etc.                                                        | Polling loop with timeout (e.g. 5s)                         | Often none or MsgBox on failure        | Show at start of wait, Update "Waiting...", Hide on success/failure            | Medium   | **Candidate** – add bar where currently no feedback     |
| Collapse/expand chat      | Collapse chat sections, add participants, etc.                                               | UIA traversal, retries                                      | `ShowSmallLoadingIndicator_ChatGPT`    | Standard bar show/update/hide                                                  | Low      | **Candidate**                                           |
| Model switch / settings   | Model switch, settings apply, rename window                                                  | UIA / window ops                                            | `ShowSmallLoadingIndicator_ChatGPT`    | Standard bar                                                                   | Low      | **Candidate**                                           |

**Migration notes:** Shift keys includes Utils.ahk; no new include needed. Ensure every path that shows the bar calls `StandardLoadingBar_Hide(0)` or `Hide(delayMs)` in `finally` or error handlers.

---

## 2. AppLaunchers.ahk

| Location                         | Entry point                                                    | Why long-running                                                 | Current feedback                                         | Loading bar insertion                                                                | Priority | Status                                       |
| -------------------------------- | -------------------------------------------------------------- | ---------------------------------------------------------------- | -------------------------------------------------------- | ------------------------------------------------------------------------------------ | -------- | -------------------------------------------- |
| Restore scroll (short path)      | `RestoreWikipediaScrollPosition(scrollPercentage, bannerText)` | UIA_Browser, JS, BlockInput, Sleep                               | `CreateCenteredBanner_Launchers`                         | Show(bannerText), Update "Scroll position restored!", Hide(500); Hide(0) on error    | High     | **Done** – uses `StandardLoadingBar_*`       |
| HandleWikipediaChar (new window) | `HandleWikipediaChar(char)` – URL branch, new window           | UIA retries, doc height stabilization, scroll verify, fullscreen | `CreateCenteredBanner_Launchers`                         | Show "Restoring scroll position...", Update at each error/success, Hide on all exits | High     | **Done** – uses `StandardLoadingBar_*`       |
| HandleWikipediaChar (existing)   | Same – existing window branch                                  | UIA retries, doc height, scroll verify                           | `CreateCenteredBanner_Launchers`                         | Same pattern                                                                         | High     | **Done** – uses `StandardLoadingBar_*`       |
| Save scroll (Monitor 3)          | `SaveWikipediaScrollPositionManually()` (e.g. hotkey)          | UIA_Browser, scroll/height read, INI write                       | `CreateCenteredBanner_Launchers`                         | Show "Saving...", Update "Scroll position saved!", Hide(1000); finally Hide(0)       | High     | **Done** – uses `StandardLoadingBar_*`       |
| Chrome open/activate             | `#!+f` (e.g.)                                                  | WinWait(10), UIA address bar                                     | `ClipAngelBanner_Show` "Checking search bar..." / "Done" | Optional: Standard bar "Opening Chrome..." / "Done"                                  | Medium   | **Candidate** – replace or keep short banner |
| Cursor open/activate             | `#!+,` (e.g.)                                                  | WinWaitActive(10)                                                | Error overlay only                                       | Show "Launching Cursor...", Hide on success/error                                    | Medium   | **Candidate** – add bar                      |
| Desktop Explorer                 | `+#e`                                                          | Loop 40 × Sleep 50, activation fallbacks                         | Error overlay only                                       | Show "Opening Explorer...", Hide on exit                                             | Low      | **Candidate**                                |

**Migration notes:** AppLaunchers includes Utils.ahk. All Wikipedia restore/save flows now use the standard loading bar. `CreateCenteredBanner_Launchers` remains for any non–loading-bar UI (e.g. Pomodoro overlay).

---

## 3. Act.ahk

| Location         | Entry point                 | Why long-running                                                                                         | Current feedback       | Loading bar insertion                                                                          | Priority | Status                                 |
| ---------------- | --------------------------- | -------------------------------------------------------------------------------------------------------- | ---------------------- | ---------------------------------------------------------------------------------------------- | -------- | -------------------------------------- |
| Startup sequence | Auto-execute (script start) | git fetch/pull scripts + Sleep 10s + git fetch/pull notes + Sleep 10s + Run multiple scripts + Run Excel | MsgBox (work env) only | Show "Updating scripts...", Update "Updating notes...", "Launching apps...", "Done", Hide(500) | High     | **Done** – uses `StandardLoadingBar_*` |

**Migration notes:** Act.ahk now includes `Utils.ahk` and calls the standard loading bar during startup. Ensure `StandardLoadingBar_Hide` runs before `Run(excelFile)` and on any early return if added.

---

## 4. Gemini.ahk

| Location           | Entry point                                      | Why long-running                                       | Current feedback                                          | Loading bar insertion                                                         | Priority | Status                                           |
| ------------------ | ------------------------------------------------ | ------------------------------------------------------ | --------------------------------------------------------- | ----------------------------------------------------------------------------- | -------- | ------------------------------------------------ |
| First-time init    | `InitializeGeminiFirstTime()` (`#!+i` first run) | Open Chrome, loop wait for window, send prompt to tabs | `ShowSmallLoadingIndicator` / `HideSmallLoadingIndicator` | Show "Opening Gemini (2 tabs)...", "Sending prompt...", Hide on all paths     | High     | **Done** – uses `StandardLoadingBar_*`           |
| Async lookup (TTS) | `#!+8` – Start/CheckCompletion/RetrieveResponse  | Async polling, retries, timeouts                       | Small loading indicator                                   | Show "Loading…", Update/Hide on completion/error                              | High     | **Done** – uses `StandardLoadingBar_*`           |
| Async TTS          | `#!+7` – Start/CheckCompletion                   | Same                                                   | Small loading indicator                                   | Same                                                                          | High     | **Done** – uses `StandardLoadingBar_*`           |
| Read aloud         | `GeminiTriggerReadAloud` / `#!+o`                | UIA search/click/retries, waits                        | `CreateCenteredBanner` + `ShowNotification`               | Show bar at start, Update "Reading aloud" / "Retrying...", Hide on done/error | Medium   | **Candidate** – replace banner with standard bar |
| Copy last message  | `CopyLastGeminiMessageToClipboard()` (`#!+p`)    | Activate, scroll, UIA, clipboard wait                  | "Copied!" / error notification                            | Show "Copying...", Hide on success/error                                      | Medium   | **Candidate** – add bar                          |
| Copy + read aloud  | Combined flow                                    | Same as read aloud + copy                              | Banners + notifications                                   | Standard bar for whole flow                                                   | Medium   | **Candidate**                                    |

**Migration notes:** Gemini includes Utils.ahk. All `ShowSmallLoadingIndicator` / `HideSmallLoadingIndicator` call sites that were used for long operations now use `StandardLoadingBar_Show` / `StandardLoadingBar_Hide`. `ShowSmallLoadingIndicator` / `HideSmallLoadingIndicator` and `CreateCenteredBanner` remain defined for any remaining short notifications.

---

## 5. WindowManagement.ahk

| Location                | Entry point                                                              | Why long-running                                       | Current feedback                     | Loading bar insertion                                                  | Priority | Status                          |
| ----------------------- | ------------------------------------------------------------------------ | ------------------------------------------------------ | ------------------------------------ | ---------------------------------------------------------------------- | -------- | ------------------------------- |
| Project selector        | `#!+l` → `ShowProjectSelector()` / `HandleSelectionModeTrigger`          | Build maps, GUI, hotkeys; can feel heavy on large sets | Selector GUI + `ShowNotification_WM` | Optional: Show "Loading projects..." until list ready                  | Medium   | **Candidate** – add bar if slow |
| Activate Cursor project | `HandleProjectSelection()` → `ActivateCursorProject()`                   | Launch Cursor, poll up to ~6s for AI field focus       | Error notification / success sound   | Show "Opening project...", Update "Focusing...", Hide on success/error | Medium   | **Candidate** – add bar         |
| Copy from Gemini        | `HandleCopyFromGeminiProjectSelection()` → `CopyFromGeminiToCursor(...)` | Cross-script, timeouts                                 | Failure notifications only           | Show "Copying from Gemini...", Hide on done/error                      | Medium   | **Candidate** – add bar         |
| Preview window          | `HandlePreviewWindowSelection()` (`3`)                                   | Scan Cursor windows/projects                           | Largely silent until done/fail       | Show "Finding preview...", Hide on exit                                | Low      | **Candidate**                   |

**Migration notes:** WindowManagement does **not** include Utils.ahk (it includes env and GeminiToCursorBridge). To use the standard bar here, add `#Include %A_ScriptDir%\Utils.ahk` and call `StandardLoadingBar_Show/Update/Hide`, or keep using `ShowNotification_WM` for short messages and add the bar only for the longest flows.

---

## 6. Microsoft Teams.ahk

| Location            | Entry point                                              | Why long-running                                  | Current feedback          | Loading bar insertion                                  | Priority | Status                  |
| ------------------- | -------------------------------------------------------- | ------------------------------------------------- | ------------------------- | ------------------------------------------------------ | -------- | ----------------------- |
| Activate with retry | `ActivateWindowWithRetry(...)` (used by meeting hotkeys) | Repeated activation strategies, waits             | Error overlays only       | Show "Activating Teams...", Hide on success/error      | Medium   | **Candidate** – add bar |
| New conversation    | `#!+r`                                                   | May launch Teams, wait up to 15s, clipboard retry | Error overlays on failure | Show "Opening new conversation...", Hide on exit       | Medium   | **Candidate** – add bar |
| Share toggle        | `#!+t`                                                   | UIA lookup, WaitListItem, sleeps                  | Completion overlay + beep | Optional: Standard bar "Toggling share..."             | Low      | **Candidate**           |
| Mute/Camera         | `#!+5` / `#!+4`                                          | State checks, retries                             | Status overlays + beep    | Optional: keep overlay for instant feedback or use bar | Low      | **Candidate**           |

**Migration notes:** Teams uses local `ShowCenteredOverlay(hwnd, text, duration)`. For long flows, add `StandardLoadingBar_Show/Update/Hide` (Teams includes Utils.ahk).

---

## 7. Outlook.ahk

| Location         | Entry point                                                                | Why long-running                  | Current feedback                         | Loading bar insertion                               | Priority | Status                  |
| ---------------- | -------------------------------------------------------------------------- | --------------------------------- | ---------------------------------------- | --------------------------------------------------- | -------- | ----------------------- |
| Mailbox/Calendar | `ActivateOutlookMailbox()` / `ActivateOutlookCalendar()` (`#!+b` / `#!+g`) | Window activation, may wait       | `ShowCenteredOverlay_Utils` on failure   | Show "Activating Outlook...", Hide on success/error | Medium   | **Candidate** – add bar |
| Reminders        | `#!+v`                                                                     | App open check, window activation | Error overlay if reminder window missing | Show "Opening reminders...", Hide on exit           | Medium   | **Candidate** – add bar |
| Voice aloud      | `#!+d` → `ExecuteVoiceAloudOption(...)`                                    | Sends, sleeps, context switches   | Selection GUI, error MsgBox              | Show "Running voice option...", Hide on done/error  | Medium   | **Candidate** – add bar |

**Migration notes:** Outlook includes Utils.ahk; can use `StandardLoadingBar_*` directly. Keep `ShowCenteredOverlay_Utils` for short result messages if desired.

---

## 8. Spotify.ahk

| Location      | Entry point                                  | Why long-running                 | Current feedback         | Loading bar insertion                            | Priority | Status                  |
| ------------- | -------------------------------------------- | -------------------------------- | ------------------------ | ------------------------------------------------ | -------- | ----------------------- |
| Open/activate | `#!+s` → `OpenSpotify()`                     | May launch app, WinWaitActive(2) | None / centering only    | Show "Opening Spotify...", Hide on success/error | Low      | **Candidate** – add bar |
| Volume + Ctrl | `*Volume_Down` / `*Volume_Up` with Ctrl path | `ActivateSpotify()` + waits      | ToolTip when not running | Optional: short bar "Activating Spotify..."      | Low      | **Candidate**           |

**Migration notes:** Spotify does **not** include Utils.ahk. To use the standard bar, add `#Include %A_ScriptDir%\Utils.ahk` and call `StandardLoadingBar_Show/Update/Hide`.

---

## 9. Utils.ahk

| Location                 | Entry point                                               | Why long-running                    | Current feedback                              | Loading bar insertion                                                 | Priority | Status                         |
| ------------------------ | --------------------------------------------------------- | ----------------------------------- | --------------------------------------------- | --------------------------------------------------------------------- | -------- | ------------------------------ |
| Quick update scripts     | Hotstring selector → `QuickUpdateScripts()`               | git pull, reload, verification      | Error/success overlays only                   | Show "Updating scripts...", Update "Reloading...", Hide on done/error | Medium   | **Candidate** – add bar        |
| Gemini focus + paste     | `GeminiNavigateFocusAndPasteFirstSnippet()` (e.g. `^!#4`) | Open Gemini, wait/load/focus, paste | Tab banner/sound                              | Show "Preparing Gemini...", Hide on exit                              | Low      | **Candidate** – add bar        |
| Delayed submit flow      | `GeminiDelayedSubmitFlow()` (`^!#L`)                      | Countdown, monitor handoff          | Countdown banner + cancel                     | Already has good UX; optional bar for "Submitting..." phase           | Low      | **Candidate**                  |
| Merge non-favorite clips | Macro `u` in selector – `MergeNonFavoriteClips()`         | Iterative UIA, many waits           | Persistent banner + completion/error overlays | Replace banner with Standard bar show/update/hide                     | Medium   | **Candidate** – replace banner |
| Dictation flows          | `~#!+0`, `#!+j`, etc.                                     | Long async lifecycle                | Indicator, banners, sounds                    | Standard bar for "Starting dictation..." etc. where appropriate       | Low      | **Candidate** – optional       |
| Desktop to Recycle Bin   | Desktop cleanup                                           | File ops                            | Banners + overlays                            | Optional: Show "Moving to Recycle Bin...", Hide on done               | Low      | **Candidate**                  |

**Migration notes:** Utils.ahk defines the standard loading bar. Use `StandardLoadingBar_Show/Update/Hide` in the same file for any long flow; ensure `finally` or every exit path calls `Hide`.

---

## Summary

- **Implemented (standard bar in use):** Shift keys (Fold/Unfold Explorer), AppLaunchers (all Wikipedia restore/save flows), Act.ahk (startup), Gemini.ahk (first-time init, async lookup/TTS loading).
- **High-priority candidates:** Shift keys Wikipedia restore/save and Marp export; optionally Chrome/Cursor launch in AppLaunchers.
- **Medium-priority candidates:** WindowManagement project/copy/preview flows; Teams activate/new conversation; Outlook activation and voice; Utils quick update and merge clips.
- **Low-priority candidates:** Short or instant-feedback flows (mute/camera, model switch, Spotify open, etc.); optional consistency replacements.

**Risks:** Always call `StandardLoadingBar_Hide(0)` (or timed) in `finally` or on every error path so the bar never stays on screen. When adding the bar to scripts that do not yet include Utils.ahk (WindowManagement, Spotify), add `#Include %A_ScriptDir%\Utils.ahk` and confirm no conflicting hotkeys or auto-execute issues.

---

## Verification (post-rollout)

1. **Shift keys – Fold/Unfold Explorer** (`^,` / `^q`): Focus Cursor with Explorer open; trigger fold then unfold. Bar should show "Folding directories..." / "Unfolding directories...", then "Directories folded" / "Directories unfolded", then disappear after ~800 ms. No bar left on screen on success or after cancel.
2. **AppLaunchers – Wikipedia**: On Monitor 3, open a Wikipedia article, scroll, use the save hotkey; then trigger restore (e.g. via `#!+k` and key). Bar should show "Saving..." / "Restoring scroll position...", update to success/error, then hide. Same for the "new window" restore path if used.
3. **Act.ahk**: Run Act (e.g. double-click or Run). Bar should show "Updating scripts...", then "Updating notes...", "Launching apps...", "Done", then hide before Excel opens. If you cancel (work env MsgBox No), bar should still hide if it was shown.
4. **Gemini.ahk**: First run init (`#!+i` when Gemini not open) or async lookup/TTS: bar should show "Opening Gemini..." / "Loading…" and hide on completion or error. No stuck bar.
5. **Linting**: No linter errors in `Utils.ahk`, `Shift keys.ahk`, `AppLaunchers.ahk`, `Act.ahk`, `Gemini.ahk` for the added/changed lines.
