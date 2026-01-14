# Debug Hypotheses: Handy Dictation Triple-Trigger

## Problem

Triple-trigger effect: `Win+Alt+Shift+0` causes three start/end sounds and three paste actions when using advanced hotkeys (`7` or `J`).

## Hypotheses

### Hypothesis A: Passthrough (~) Prefix Recursion

**Theory:** The `~` prefix on `#!+0` allows the key to pass through to both AHK and the OS/handy.exe. When `SendInput "#!+0"` is called from `7`/`J` hotkeys, it may trigger:

- The AHK hotkey handler (ToggleDictationMode)
- The OS/handy.exe directly
- A recursive loop if handy.exe sends the key back

**Expected Log Pattern:** Multiple "Hotkey 0 pressed" entries within milliseconds, or "SendInput #!+0" followed immediately by "Hotkey 0 pressed".

### Hypothesis B: Recursive SendInput Loop

**Theory:** `SendInput "#!+0"` in `7`/`J` hotkeys triggers the `~#!+0` hotkey handler, which may call `ToggleDictationMode()`, which could somehow trigger another state change detection.

**Expected Log Pattern:** "Hotkey 7/J pressed" → "SendInput #!+0" → "Hotkey 0 pressed" → "ToggleDictationMode called" → repeated multiple times.

### Hypothesis C: Race Condition Between Timer and Manual Check

**Theory:** `StartDictationCheckTimer` runs every 500ms, and `ToggleDictationMode()` manually calls `CheckDictationRecordingWindow()`. Both may detect the same state transition simultaneously, causing multiple chime schedules and paste actions.

**Expected Log Pattern:** Single "Hotkey 0 pressed" but multiple "State Change Detected" entries, or multiple "Chime Scheduled" entries.

### Hypothesis D: Multiple Chime Timer Fires

**Theory:** `PlayDictationCompletionChime` is scheduled multiple times (multiple `SetTimer` calls), or the test-and-set pattern fails, allowing multiple executions.

**Expected Log Pattern:** Multiple "Chime Scheduled" entries, or multiple "PlayDictationCompletionChime called" entries.

### Hypothesis E: Hotkey Handler Called Multiple Times

**Theory:** The hotkey itself is being triggered multiple times by the system (keyboard repeat, multiple key events, or hotkey registration issue).

**Expected Log Pattern:** Multiple "Hotkey 0 pressed" entries with timestamps showing they're from actual key presses (not SendInput).

### Hypothesis F: State Detection Window Exists Check Fails

**Theory:** `WinExist("Recording ahk_exe handy.exe")` may return inconsistent results, causing `CheckDictationRecordingWindow` to detect multiple start/stop transitions.

**Expected Log Pattern:** Multiple "Window exists check" entries with alternating true/false values in rapid succession.

## Logging Strategy

- Entry points: All hotkey handlers
- State changes: Window existence checks, state transitions
- Timer events: Chime scheduling, pulse timer
- Action execution: Paste actions, SendInput calls

## Analysis Results (from runtime logs)

### Hypothesis C: Race Condition Between Timer and Manual Check

**STATUS: CONFIRMED**

**Evidence:**

- Multiple "State Change Detected" entries within milliseconds (e.g., lines 4044-4049: two state changes at 22734.991 and 22735.200, only 209ms apart)
- Multiple "Chime Scheduled" entries for the same transition (lines 4045, 4049)
- No "State transition skipped" entries found, indicating the guard wasn't working

**Root Cause:** Multiple calls to `CheckDictationRecordingWindow()` (from timer and manual calls) were detecting the same state transition simultaneously. The guard check happened too late - multiple calls could all read `g_DictationCompletionChimeScheduled = false` before any of them set it to `true`.

### Hypothesis D: Multiple Chime Timer Fires

**STATUS: CONFIRMED**

**Evidence:**

- Multiple "PlayDictationCompletionChime called" entries (e.g., lines 4076, 4078, 4082 within 202ms)
- Multiple "Chime executing" entries, indicating the test-and-set pattern in `PlayDictationCompletionChime` wasn't preventing duplicates

**Root Cause:** Multiple chimes were scheduled due to multiple state transitions being detected (see Hypothesis C).

### Hypothesis E: Hotkey Handler Called Multiple Times

**STATUS: PARTIALLY CONFIRMED**

**Evidence:**

- Duplicate "Hotkey 7 pressed" entries with identical timestamps (lines 4029-4030: both at 22733.958)
- This suggests the hotkey handler itself may be called multiple times, possibly due to key repeat or system-level duplicate events

### Hypotheses A, B, F

**STATUS: INCONCLUSIVE**

- No clear evidence of passthrough recursion, SendInput loops, or window existence check failures
- These may contribute but are not the primary cause

## Fix Implemented

1. **Atomic Critical Section**: Made the entire stop transition block atomic using `Critical "On"` to prevent multiple simultaneous calls from processing the same transition
2. **Double-Check Pattern**: Re-check conditions inside the critical section after acquiring the lock
3. **Cooldown Period**: Added 1-second cooldown (`g_LastStateTransitionTick`) to prevent rapid re-detection of the same transition
4. **Atomic Flag Setting**: Set both `g_DictationCompletionChimeScheduled` and `g_LastStateTransitionTick` atomically within the critical section

**Expected Behavior After Fix:**

- Only one "State Change Detected" entry per actual transition
- Only one "Chime Scheduled" entry per transition
- "State transition skipped" entries should appear when duplicates are prevented
- Single paste action execution
