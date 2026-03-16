# Dictation Flow Multiple Trigger Bug: Diagnostic Hypothesis

## Overview

The Dictation-to-Gemini-Cursor flow is currently spawning multiple overlapping instances of the first `Send to Gemini?` banner immediately after transcription ends. Because the user is forced to press 'N' exactly seven or more times to clear them, this strongly indicates that **multiple independent script processes are executing the flow simultaneously**, rather than a loop bug within a single script.

Below are the potential causes, ordered from the most probable to the least probable, accompanied by a systematic testing and debugging guide.

---

## Hypothesis 1: `#include` Multi-Process Duplication (Highly Probable)

**The Theory:**
The file `Utils.ahk` contains the core logic for the dictation hotkey (`~#!+0::`), the state check timer (`CheckDictationRecordingWindow`), and the GUI banner generation.
However, `Utils.ahk` is designed as a library and is `#include`d in multiple independent scripts (e.g., `Gemini.ahk`, `WindowManagement.ahk`, `AppLaunchers.ahk`, `Spotify.ahk`, etc. — up to 8 scripts are listed in `QuickUpdateScripts`).

Because the dictation hotkey uses the `~` (pass-through) modifier, **every single script process that includes `Utils.ahk` simultaneously registers and fires the hotkey.**
When dictation stops, all 7–8 background scripts independently detect the stop, evaluate their own isolated state machines, and spawn their own GUI banners. They stack visually, forcing you to press 'N' 7–8 times to close each process's banner.

**Testing & Validation:**

1. Open `Utils.ahk` and locate the `~#!+0::` hotkey.
2. Add a temporary debug tooltip or log as the very first line:
   ```autohotkey
   ToolTip("Hotkey fired in: " . A_ScriptName)
   Sleep 1000
   ToolTip()
   ```

Run your standard script launch routine. Press Win+Alt+Shift+0.
If you see the tooltip rapidly flashing with different script names (e.g., Gemini.ahk, then WindowManagement.ahk), this hypothesis is confirmed.
Proposed Fix: Restrict the dictation hotkey and its associated timers so they only initialize in a single primary script. You can wrap the hotkey at the bottom of Utils.ahk in an #HotIf directive
#HotIf A_ScriptName == "Utils.ahk"
~#!+0::
{
; ... hotkey logic ...
}
#HotIf

Alternatively, extract the dictation trigger logic out of the shared library and put it directly into a single orchestrator script like Act.ahk or Utils.ahk running standalone).
Hypothesis 2: Timer Re-entrancy during Ultra-Fast Polling
The Theory: When the hotkey is pressed, ToggleDictationMode() switches the window validation timer to an ultra-fast polling rate of 25ms (SetTimer(CheckDictationRecordingWindow, 25)). If the handy.exe window takes longer than 25ms to close, or if the main thread is blocked (e.g., by the PowerShell script RunWait), multiple timer ticks could pile up. If the Critical "On" guard in CheckDictationRecordingWindow isn't strictly preventing queued executions, multiple ticks might evaluate the stop condition and schedule the chime multiple times.

Testing & Validation:

Comment out the 25ms polling acceleration in ToggleDictationMode(
; SetTimer(CheckDictationRecordingWindow, 25)
; SetTimer(RevertDictationPolling, -3000)

Test the dictation flow. If the multiple banner issue disappears, timer re-entrancy is the culprit.
Proposed Fix: Add a strict re-entrancy guard at the top of CheckDictationRecordingWindow:

static isChecking := false
if (isChecking)
return
isChecking := true
; ... existing logic ...
isChecking := false

Hypothesis 3: Clipboard Event Spam
The Theory: When the flow finishes and dictation text is pasted to the clipboard, DictationCompletionChimeOrWaitForClipboard binds the DictationClipboardHandler to OnClipboardChange. If handy.exe updates the clipboard in multiple chunks or formats, the clipboard event might fire 5+ times instantly.

Testing & Validation:

Inside DictationClipboardHandler(DataType), add a debug lo
DebugFlowLog("Utils.ahk:ClipboardHandler", "Clipboard event fired", "Type=" DataType)

Test the dictation flow. Check the .cursor/debug-7e3dd7.log file. If the "Clipboard event fired" log appears 5+ times with the exact same timestamp, the clipboard is spamming the completion trigger.
Proposed Fix: The current OnClipboardChange(DictationClipboardHandler, 0) is supposed to unregister the event, but if 7 events are already in the Windows message queue, they might all fire. A timestamp debounce inside DictationClipboardHandler would discard subsequent rapid fires within a 500ms window.

Hypothesis 4: KeyWait Thread Queuing (Debounce Bypass)
The Theory: In the ~#!+0:: hotkey, the debounce update (lastHotkeyTick := currentTick) occurs before KeyWait("0", "L"). If the user holds the key slightly too long, or the hardware bounces, AHK might queue additional hotkey threads that wait for the first one to finish. Once KeyWait releases, the queued threads bypass the debounce check instantly because currentTick reflects their past queue time.

Testing & Validation:

Press and intentionally hold Win+Alt+Shift+0 for 2 seconds before releasing.
Observe if the issue reliably scales with how long you hold the key.
Proposed Fix: Move the lastHotkeyTick := A_TickCount update to the very end of the hotkey definition, directly before isProcessing := false, to ensure the 200ms lock applies after the key is physically released.

<!--
[PROMPT_SUGGESTION]Apply the #HotIf A_ScriptName restrictor to the ~#!+0 hotkey in Utils.ahk to test Hypothesis 1.[/PROMPT_SUGGESTION]
[PROMPT_SUGGESTION]Extract the dictation window monitoring timers and hotkeys from Utils.ahk into a separate DictationController.ahk file to prevent multi-process duplication.[/PROMPT_SUGGESTION]
-->
