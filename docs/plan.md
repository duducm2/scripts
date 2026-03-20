---
name: Fix Dictation Flow Gemini Copy Race Condition
overview: Resolve the race condition in the dictation flow where the script fails to copy Gemini's response by changing the asynchronous IPC calls to synchronous.
todos:
  - id: update_ipc_to_sync
    content: Change PostMessage to SendMessage in D2C_FlowManager.DoCopyCore to prevent focus stealing during the UIA copy sequence.
    status: completed
---

# Fix Dictation Flow Gemini Copy Race Condition

## Analysis / Context
The voice dictation flow (`D2C_FlowManager` in `Utils.ahk`) relies on Inter-Process Communication (IPC) to trigger the working copy function (`CopyLastGeminiMessageToClipboard`) residing in `Gemini.ahk`. Currently, it uses `PostMessage`, which is asynchronous.

Because it does not wait for the message to be fully processed, `Utils.ahk` continues executing. Meanwhile, `Gemini.ahk` clears the clipboard (`A_Clipboard := ""`) at the start of its copy routine. `Utils.ahk` detects this initial clipboard clearing as a sequence change via `Clipboard_WaitForSequenceChange`, incorrectly assumes the copy is complete, and immediately proceeds to restore focus to the original window. 

This premature action steals focus *before* `Gemini.ahk` can execute its UI scroll (`^{End}`) and the UIA click on the copy button, causing the copy to fail or keystrokes to bleed into the wrong window.

## Proposed Changes
- Convert the asynchronous `PostMessage` calls in `D2C_FlowManager.DoCopyCore` to synchronous `SendMessage` calls. 
- Using `SendMessage` ensures `Utils.ahk` safely blocks and waits for `Gemini.ahk` to complete the full UIA interaction and `ClipWait` before attempting to restore focus to the original context.
- This approach satisfies all constraints: it relies on the already-working UIA logic inside `Gemini.ahk` (which explicitly targets the latest response via `br.t` coordinate checks), maintains code integrity, and leaves the generation state verification (`CheckGeminiCompletion`) perfectly intact.

## Files to Modify
- `Utils.ahk`

## Implementation Strategy
1. Open `Utils.ahk` and locate the `DoCopyCore` method inside the `D2C_FlowManager` class.
2. Find the `WM_COPY_LAST_GEMINI` IPC dispatch block (in `DoCopyCore`, search for `WM_COPY_LAST_GEMINI`).
3. Replace the asynchronous call with synchronous `SendMessage`:
   ```ahk2
   try SendMessage(WM_COPY_LAST_GEMINI, 0, 0, , "ahk_id " targetHwnd)
   ```

## Status
Done: `Utils.ahk` `D2C_FlowManager.DoCopyCore` now uses `SendMessage` for `WM_COPY_LAST_GEMINI`.
