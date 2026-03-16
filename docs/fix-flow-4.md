---
name: Fix Dictation Hotkey State Loss and Blocking
overview: Resolve bugs where the dictation start/stop actions are missed by eliminating thread-blocking commands (RunWait, KeyWait) and proactively validating the handy.exe process and window state.
todos:
  - id: remove_hotkey_thread_blockers
    content: In Utils.ahk, inside the `~#!+0::` hotkey, completely remove the `isProcessing` static variable logic and delete the `KeyWait("0", "L")` command. Change the PowerShell `RunWait` command for `Set-MicVolume.ps1` to `Run` so the thread executes instantly without blocking subsequent hotkey presses.
    status: pending
  - id: remove_timer_thread_blockers
    content: In Utils.ahk, inside the `CheckDictationRecordingWindow()` function, locate the start branch that calls `Set-MicVolume.ps1` and change the `RunWait` command to `Run`.
    status: pending
    dependencies: [remove_hotkey_thread_blockers]
  - id: implement_handy_process_validation
    content: In Utils.ahk, inside the `~#!+0::` hotkey (immediately after the debounce check), add a validation layer that checks `if (!ProcessExist("handy.exe"))`. If it does not exist, call `Run(GetHandyShortcutPath())` and use `ProcessWait("handy.exe", 3)` before proceeding.
    status: pending
    dependencies: [remove_hotkey_thread_blockers]
  - id: implement_robust_stop_state_resolution
    content: In Utils.ahk, inside the `~#!+0::` hotkey, replace the reliance on `dictationWasActiveOnKeyPress := g_DictationActive` with a direct window check: `recordingWindowExists := WinExist("Recording ahk_exe handy.exe")`. Update the stop condition to evaluate `if (recordingWindowExists || g_DictationActive)` to set `g_PendingGeminiPromptAfterDictation := true`.
    status: pending
    dependencies: [implement_handy_process_validation]
---

# Fix Dictation Hotkey State Loss and Blocking

## Analysis / Context

The implementation of the single-owner mutex successfully fixed the duplicate execution of the Dictation-to-Gemini flow, but two new critical bugs emerged: failure to trigger the start action, and failure to register the stop action (which skips the "Send to Gemini?" banner).

**Root Cause 1 (Stop Action Missed):**
The `~#!+0::` hotkey currently uses `KeyWait("0", "L")` and synchronously executes the `Set-MicVolume.ps1` script via `RunWait`. `RunWait` freezes the AutoHotkey thread for 1–2 seconds. Because AutoHotkey's default `MaxThreadsPerHotkey` is 1, if the user speaks a quick sentence and presses `~#!+0` a second time to stop _while the first thread is still frozen_, the second press is entirely ignored by AHK. The dictation window closes (handled by OS pass-through to `handy.exe`), but AHK never sets `g_PendingGeminiPromptAfterDictation := true`. The background timer then detects the window closure and assumes it was an abort, silently ending the flow.

**Root Cause 2 (Start Action Fails):**
AHK assumes `handy.exe` is running and relies on the OS pass-through to trigger it. If `handy.exe` is closed, the hotkey does nothing.

## Proposed Changes

1. **Eliminate Thread Starvation**: Remove `KeyWait` and `isProcessing` locks. Convert all `RunWait` calls for volume adjustment into asynchronous `Run` calls. This guarantees the hotkey thread completes in milliseconds and is immediately ready to accept the stop input.
2. **Proactive Process Validation**: Automatically launch `handy.exe` if the process is missing when the user presses the hotkey.
3. **Robust State Resolution**: Calculate the stop condition by directly querying the OS for the recording window (`WinExist("Recording ahk_exe handy.exe")`) rather than trusting only the async boolean state, ensuring the Gemini prompt is strictly queued if the dictation UI was visible.

## Files to Modify

- `c:\Users\eduev\Meu Drive\17 - Projects\scripts\Utils.ahk`

## Implementation Strategy

1. **Fix Thread Blockers in Hotkey**:
   - Locate `~#!+0::` in `Utils.ahk`.
   - Remove `static isProcessing := false` and all assignments/checks related to it.
   - Remove `KeyWait("0", "L")`.
   - Change `RunWait "powershell.exe...` to `Run "powershell.exe...`.

2. **Fix Thread Blockers in Timer**:
   - Locate `CheckDictationRecordingWindow()` in `Utils.ahk`.
   - Change its `RunWait "powershell.exe...` to `Run "powershell.exe...`.

3. **Handy Launch Guarantee**:
   - In `~#!+0::`, right after `lastHotkeyTick := currentTick`, insert:
     ```autohotkey
     if (!ProcessExist("handy.exe")) {
         handyPath := GetHandyShortcutPath()
         if (handyPath != "" && FileExist(handyPath)) {
             Run(handyPath)
             ProcessWait("handy.exe", 3)
         }
     }
     ```

4. **Reliable Stop Detection**:
   - In `~#!+0::`, replace:
     ```autohotkey
     dictationWasActiveOnKeyPress := g_DictationActive
     ```
     with:
     ```autohotkey
     recordingWindowExists := false
     try recordingWindowExists := WinExist("Recording ahk_exe handy.exe")
     ```
   - Update the conditional block at the bottom of the hotkey:
     ```autohotkey
     if (recordingWindowExists || g_DictationActive) {
         g_PendingGeminiPromptAfterDictation := true
         g_DictationGeminiConfirmBannerVisible := false
         ; ... (logging)
     }
     ```
