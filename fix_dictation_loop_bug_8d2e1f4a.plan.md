---
name: Fix Dictation Loop Bug
overview: Debug and optimize the dictation loop macro in Utils.ahk. The current implementation uses periodic timers which cause the loop to malfunction after the first iteration. The fix involves switching to one-shot timers and increasing the processing buffer.
todos:
  - id: fix_timer_logic
    content: Update DictationLoopStart and DictationLoopStop functions to use negative timer periods (one-shot) instead of positive (periodic).
    status: pending
    dependencies: []
  - id: optimize_buffer
    content: Increase the inter-loop delay in DictationLoopStop from 2000ms to 4000ms to ensure "Handy" has sufficient time to process transcription.
    status: pending
    dependencies: [fix_timer_logic]
---

# Fix Dictation Loop Bug

## Analysis / Context
The user reported that the dictation loop works for the first iteration but fails subsequently, triggering the termination command every ~3 seconds.

**Root Cause Analysis:**
In AutoHotkey, `SetTimer(Function, Period)` creates a **periodic** timer if `Period` is positive.
1. `DictationLoopStart` sets `DictationLoopStop` to run every 40s.
2. At T=40s, `DictationLoopStop` runs. It sets `DictationLoopStart` to run every 2s.
3. At T=42s, `DictationLoopStart` runs (triggered by the 2s timer). It resets the 40s timer.
4. At T=44s, `DictationLoopStart` runs **AGAIN** because the 2s timer is still active and periodic.
This causes the "Start" command (`#!+0`) to fire repeatedly every 2 seconds, rapidly toggling the dictation state and ruining the workflow.

## Proposed Changes
1.  **Switch to One-Shot Timers:** Change the `SetTimer` period arguments to negative values (e.g., `-40000`). This tells AutoHotkey to run the timer exactly once and then delete it.
2.  **Increase Processing Buffer:** Increase the delay between "Stop" and the next "Start" from 2 seconds to 4 seconds. This provides a robust buffer for the "Handy" software to finalize transcription before the next cycle begins.

## Files to Modify
*   `c:\Users\fie7ca\Documents\scripts\Utils.ahk`

## Implementation Strategy
1.  Locate `DictationLoopStart()` in `Utils.ahk`.
    *   Change `SetTimer(DictationLoopStop, 40000)` to `SetTimer(DictationLoopStop, -40000)`.
2.  Locate `DictationLoopStop()` in `Utils.ahk`.
    *   Change `SetTimer(DictationLoopStart, 2000)` to `SetTimer(DictationLoopStart, -4000)`.
