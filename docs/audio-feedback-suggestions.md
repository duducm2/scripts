# Audio Feedback – Suggested Insertion Points

This document is the single source of truth for adding sound effects to user actions across the AHK scripts. When implementing, always guard playback with `IsSoundEnabled()` (from `Utils.ahk`) and use assets under `A_ScriptDir . "\sounds\"` (e.g. `sounds\success.wav`, `sounds\error.wav`).

---

## Suggested insertion points by file

### Utils.ahk

| File      | Function / line area | Action type                         | Suggested sound type/asset     | Notes                                                                 |
| --------- | -------------------- | ----------------------------------- | ------------------------------ | --------------------------------------------------------------------- |
| Utils.ahk | `DesktopToRecycle_Run()` ≈3333 | Process completion (destructive)    | Short “task done” chime        | After `exitCode = 0`, with green banner “Desktop items moved to Recycle Bin”. |
| Utils.ahk | `DesktopToRecycle_Run()` ≈3335–3337 | Error                                | Soft error tone                | In `else` / `catch` when path not found or error.                     |
| Utils.ahk | `QuickUpdateScripts()` ≈536 | Process completion                   | Reuse “success” chime          | When `failedScripts.Length = 0` and “All scripts updated successfully!”. |
| Utils.ahk | `QuickUpdateScripts()` ≈529–534 | Error                                | Same soft error tone           | When `failedScripts.Length > 0` and showing failed list.              |
| Utils.ahk | Handy Open AI model ≈1381 | _(existing)_                         | No change                      | Already uses `handy-model-chosen.mp3`.                                |
| Utils.ahk | MergeNonFavoriteClips ≈1023–1025 | Process completion                   | Very short “merge complete”    | When showing “Done” and hiding banner (Clip Angel).                   |
| Utils.ahk | `ToggleOutlookAndTeams` ≈1916 | Process completion                   | Short “toggle complete”        | When showing “Done”.                                                  |
| Utils.ahk | `CleanClipboardInternal()` success path | Process completion                   | Short “cleanup done” chime     | After cleanup algorithm finishes successfully (identify single success path). |
| Utils.ahk | `DictationStartWithClipboardOption` ≈2580 | Confirmation (optional)              | Very short “ack”               | When user clicks “Yes” before `CleanClipboardInternal()`; avoid double with cleanup-done. |
| Utils.ahk | Dictation stop / merge countdown | _(existing)_                         | No change                      | Already uses `speach-finished.wav` / `retro4.wav`.                    |

### AppLaunchers.ahk

| File            | Function / line area | Action type                    | Suggested sound type/asset     | Notes                                                                 |
| --------------- | -------------------- | ------------------------------ | ------------------------------ | --------------------------------------------------------------------- |
| AppLaunchers.ahk | `HandleWikipediaChar` ≈754–798 | Process completion             | Short “selection confirmed”     | After article chosen and page ready (e.g. after scroll restore / WinWaitActive). |
| AppLaunchers.ahk | `ShowCursorFallbackPanel()` ≈101 | Warning / info                  | Very short, low-priority cue   | When no Cursor window found.                                          |
| AppLaunchers.ahk | Pomodoro start ≈1908; completion | _(existing)_                    | Optional: `pomodo-complete.wav` | Start has `pomodo-start.wav`; completion uses SoundBeep + MessageBeep; optional dedicated completion asset. |

### WindowManagement.ahk

| File                 | Function / line area | Action type           | Suggested sound type/asset   | Notes                                                                 |
| -------------------- | -------------------- | --------------------- | ---------------------------- | --------------------------------------------------------------------- |
| WindowManagement.ahk | `MoveWinToMonitor()` ≈426–430 | Process completion    | Short “window moved” cue     | After successful WinMove/maximize, before/after `MoveMouseToCenter`.    |
| WindowManagement.ahk | `CloseWindowOnMonitor()` after ≈690 | Deletion / removal    | Short “closed” sound         | After successful `WinClose` (no exception).                           |
| WindowManagement.ahk | `MinimizeWindowOnMonitor()` ≈656 | Confirmation          | Very short “minimized” cue   | In try block success path.                                            |
| WindowManagement.ahk | `CycleWindowsOnMonitor()` ≈508+ | Confirmation (optional) | Subtle “cycle” tick          | When focus switched to next window; keep low volume/short.            |
| WindowManagement.ahk | `ActivateCursorProject` ≈801–804 | _(existing)_          | Consistency fix only         | Add `IsSoundEnabled()` check around `into-cursor-textfield.wav`.      |
| WindowManagement.ahk | `HandleProjectSelection` success | Process completion    | Already covered             | Covered by `into-cursor-textfield.wav` in ActivateCursorProject.     |
| WindowManagement.ahk | Copy-from-Gemini failure ≈1481–1493 | Error                 | One soft error tone          | For “Gemini.ahk not running”, “Open Gemini in Chrome first”, etc.     |

### Gemini.ahk

| File      | Function / line area | Action type        | Suggested sound type/asset | Notes                                                                 |
| --------- | -------------------- | ------------------ | -------------------------- | --------------------------------------------------------------------- |
| Gemini.ahk | PlayCopyCompletedChime / WaitForButtonAndShowSmallLoading | _(existing)_ | No change                  | Copy/response ready already has sound.                                |
| Gemini.ahk | `GeminiTriggerReadAloud` Pause/Resume ≈285, 312 | Confirmation (optional) | Very short “pause”/“resume” ack | Different from completion.                                            |
| Gemini.ahk | Read-aloud start (after clicking option) | Process completion | Short “playback started”    | When “Read aloud” actually starts; can share style with completion chimes. |

### Shift keys.ahk

| File         | Function / line area | Action type     | Suggested sound type/asset | Notes                                                                 |
| ------------ | -------------------- | --------------- | -------------------------- | --------------------------------------------------------------------- |
| Shift keys.ahk | Cursor git commit ≈8983, 9049 | _(existing)_    | `cursor-git-commit.wav`     | No change.                                                            |
| Shift keys.ahk | `PlayCompletionChime_Gemini` ≈13825 | _(existing)_    | `gemini-completion.wav`    | No change.                                                            |
| Shift keys.ahk | Abort / empty commit ≈9039, 9040 | Error / abort   | Short “aborted” or soft error | When commit is aborted (e.g. empty commit message).                   |

---

## Sound design recommendations (by context)

- **Confirmation** (user choice registered, e.g. Yes to cleanup, Pause/Resume): Short (50–150 ms), medium pitch, single tone or two-note “ack”. Avoid harsh or high volume.
- **Process completion** (task finished successfully): Clear “done” cue; 100–250 ms; can reuse one “success” asset for script update, merge clips, toggle apps, cleanup done, window move, project open (or use 2–3 variants: “file/task”, “window”, “clipboard”).
- **Deletion / close** (Recycle Bin, close window): Slightly lower or different timbre than generic success so it’s distinguishable; still neutral (not alarming).
- **Error / warning** (failed update, no Cursor window, no Gemini, abort commit): Soft error (e.g. short low tone or single “bonk”); avoid system-critical style.
- **Optional / low priority** (cycle window, fallback panel): Very short and quiet so they don’t annoy with repeated use.

**Implementation note:** Prefer reusing a small set of assets (e.g. `success.wav`, `error.wav`, `close.wav`) and calling them from multiple sites, with optional second set for “window” vs “clipboard” if you want more variety later.

---

## Already has sound (no change)

These locations already play a sound; no new insertion needed:

- **Utils.ahk:** Handy AI model chosen ≈1381 (`handy-model-chosen.mp3`); dictation stop / merge uses `speach-finished.wav` / `retro4.wav`.
- **AppLaunchers.ahk:** Pomodoro start ≈1908 (`pomodo-start.wav`); completion uses `PlayCompletionChime` (SoundBeep + MessageBeep).
- **WindowManagement.ahk:** `ActivateCursorProject` focus success ≈801–804 (`into-cursor-textfield.wav`); project selection success is covered by that same sound.
- **Gemini.ahk:** Copy completed / response ready (`copy.wav` via `PlayCopyCompletedChime`; `gemini-completion.wav` where used).
- **Shift keys.ahk:** Cursor git commit ≈8983, 9049 (`cursor-git-commit.wav`); `PlayCompletionChime_Gemini` ≈13825 (`gemini-completion.wav`).

---

## Consistency fix (code change, no new sound)

- **WindowManagement.ahk** ≈801–804: The existing `into-cursor-textfield.wav` is played without an `IsSoundEnabled()` check. Wrap the `SoundPlay` call in `if (IsSoundEnabled()) { ... }` so it respects the global sound toggle like all other play sites.
