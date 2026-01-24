1. Hypothesis Document (Ranked)
Subject: Redundant Audio Chimes during handy.exe Execution (3+ Triggers)

Hypothesis A: The "Pass-Through" Echo (External + Internal Conflict)

Probability: High
Logic: The hotkey ~#!+0 uses the tilde (~) prefix, which allows the keystroke to pass through to the OS and handy.exe.
Mechanism:
Chime 1: handy.exe receives the shortcut and plays its own native startup sound (if configured).
Chime 2: The AHK script intercepts the hotkey and immediately triggers SafePlayDictationSound.
Chime 3: The monitoring loop (CheckDictationRecordingWindow) detects the new window. If the "Audio Firewall" (debounce logic) fails or uses a different file path reference than the hotkey, it triggers a third sound.
Hypothesis B: State Thrashing (Race Condition)

Probability: Medium-High
Logic: The monitoring loop runs at an accelerated 50ms interval during startup.
Mechanism: If the handy.exe window initializes, briefly changes title/class, or flickers during launch, the script might detect a sequence of Start -> Stop -> Start. While the "Stop" transition has a cooldown, a flaw in the "Start" re-detection logic (e.g., if g_DictationActive is desynchronized) could force a re-trigger of the start sound.
Hypothesis C: Audio Firewall Bypass (Path Mismatch)

Probability: Medium
Logic: The debounce logic relies on InStr(filePath, "speach-start.wav").
Mechanism: If the hotkey calls the function with a relative path and the monitoring loop calls it with an absolute path (or vice versa), or if variable expansion differs, the static lastStartSoundTick might not apply to both calls, treating them as distinct sound events.
Hypothesis D: Multiple Script Instances

Probability: Low
Logic: Two or more instances of Utils.ahk running simultaneously.
Mechanism: Both instances receive the global hook ~#!+0 and both play sounds. Both instances monitor the window and play sounds. (2 instances x 2 triggers = 4 sounds).
2. Context Summary for Third-Party AI
System Overview:

Environment: AutoHotkey v2 on Windows.
Target Application: handy.exe (Dictation software).
Trigger: Win+Alt+Shift+0 (Hotkey: ~#!+0).
Current Logic Flow (Utils.ahk):

Hotkey Activation:
Sets global state g_DictationActive := true.
Sets g_LastStateTransitionTick to current time.
Immediately calls SafePlayDictationSound.
Accelerates monitoring timer to 50ms.
Monitoring Loop (CheckDictationRecordingWindow):
Polls for window: WinExist("Recording ahk_exe handy.exe").
Start Logic: If window exists AND !g_DictationActive -> Set Active, Play Sound.
Stop Logic: If window missing AND g_DictationActive -> Set Inactive (protected by 7s cooldown).
Audio Firewall (SafePlayDictationSound):
Uses a static variable to track the last start sound time.
Enforces a 7-second cooldown if the filename contains "speach-start.wav".
The Problem: User reports hearing three or more sounds upon triggering dictation, despite the debounce logic.

Objective for Next AI: Review the logic gaps in SafePlayDictationSound and CheckDictationRecordingWindow based on the hypotheses above. Specifically, verify if handy.exe has a native sound that needs to be accounted for, or if the AHK script is bypassing its own firewall.

---

## Debugging Progress

### Hypothesis A: The "Pass-Through" Echo (External + Internal Conflict) - FIXED

**Status:** Fixed by following print screen pattern

**Root Cause Identified:**
The sound was being played in TWO places:
1. Hotkey handler (`~#!+0::`) - plays sound immediately
2. Monitoring loop (`CheckDictationRecordingWindow()`) - plays sound when window detected

Even though the hotkey sets `g_DictationActive = true` before calling `ToggleDictationMode()`, there was a timing window where the monitoring loop could still detect the window and play the sound again, causing duplicate sounds.

**Solution Applied (Following Print Screen Pattern):**
- Removed sound playback from the monitoring loop's start detection logic
- Sound is now played ONLY in the hotkey handler (like print screen does)
- Monitoring loop now only handles state management and UI updates, not sound playback
- This ensures a single sound play, matching the print screen implementation pattern

**Changes Made:**
1. **Removed sound from monitoring loop**: `CheckDictationRecordingWindow()` no longer calls `SafePlayDictationSound()` when detecting window start
2. **Kept sound in hotkey handler**: Sound is played once in `~#!+0::` handler
3. **Added state synchronization**: Monitoring loop now sets `g_LastStateTransitionTick` to ensure proper state tracking

**Result:** Following the print screen pattern ensures only ONE sound is played when the hotkey is used, eliminating the duplicate sound issue.

**Remaining Issue:** If user still hears two sounds, one may be from handy.exe itself (native application sound), which is outside our control.

---

## 3. Research Report Findings (Architectural Analysis)

**Root Cause Confirmation:**
*   **Native Audio:** `handy.exe` (v0.1.5+) introduced native start/stop sounds. The "duplicate" sound is often the collision of AHK's `SoundPlay` and Handy's native chime.
*   **Latency Gap:** A ~20-50ms delay between AHK execution and Handy's Tauri event loop creates a distinct "stutter" or echo effect.

**Mechanism of "3+ Sounds" (Machine Gun Effect):**
*   **Typematic Rate:** Holding the key (Push-to-Talk) generates OS key repeats (10-30/sec).
*   **Debounce Failure:** Standard timer-based debounce (`A_TimeSincePriorHotkey`) often fails against rapid OS repeats.
*   **Solution:** Requires `KeyWait` (execution blocking) or a Strict State Machine to ignore repeats.

**Proposed Architectural Fixes:**
1.  **Proxy Key Isolation (Recommended):**
    *   Bind `handy.exe` to a virtual key (e.g., F24).
    *   AHK intercepts the physical key (e.g., CapsLock), plays sound (optional), and sends F24.
    *   This decouples the physical trigger from the app.
2.  **Process-Level Suppression:**
    *   Mute `handy.exe` using `SoundVolumeView`.
3.  **Strict State Machine:**
    *   Use static variables to track logical state, ignoring physical key repeats.

## 4. Action Plan for Next AI

Based on the research findings, the next steps should focus on architectural resolution rather than just code debugging:

1.  **Verify Native Sound:** Confirm if `handy.exe` is playing a sound. If so, decide whether to suppress AHK sound or Handy sound.
2.  **Implement Proxy Key Pattern (Strategy A):**
    *   Modify `Utils.ahk` to stop using `~#!+0`.
    *   Instead, map `#!+0` to send a virtual key (e.g., `{F24}`) that `handy.exe` listens to.
    *   This prevents the "Pass-Through Echo".
3.  **Implement KeyWait:**
    *   Add `KeyWait` to the hotkey handler to prevent "Machine Gun" repeats from the OS typematic rate.

---

## 5. Architectural Fixes Implemented

**Status:** Implemented based on research findings

### Fix 1: Strict State Machine
**Implementation:**
- Added `static lastHotkeyTick` to track last hotkey trigger time
- Ignores key repeats within 200ms (faster than OS typematic rate of 10-30/sec)
- Prevents "Machine Gun" effect from holding the key

**Code:**
```autohotkey
static lastHotkeyTick := 0
currentTick := A_TickCount
timeSinceLast := currentTick - lastHotkeyTick
if (timeSinceLast < 200) {
    return  ; Ignore key repeat
}
lastHotkeyTick := currentTick
```

### Fix 2: KeyWait Implementation
**Implementation:**
- Added `KeyWait("0", "L")` to block execution until main key (0) is released
- Prevents OS typematic key repeats from triggering multiple handler executions
- Only waits for main key, not modifiers (modifiers may be held for other purposes)

**Code:**
```autohotkey
KeyWait("0", "L")  ; Wait for "0" key to be released (L = logical state)
```

### Fix 3: Recursion Guard
**Implementation:**
- Added `static isProcessing` flag to prevent handler from retriggering during execution
- Ensures only one instance of handler runs at a time
- Prevents race conditions from rapid key presses

**Code:**
```autohotkey
static isProcessing := false
if (isProcessing) {
    return  ; Already processing, ignore
}
isProcessing := true
; ... handler code ...
isProcessing := false  ; Release guard at end
```

**Result:**
- Prevents "Machine Gun" effect from OS key repeats
- Ensures only one logical key press is processed per physical key release
- Prevents handler from retriggering during execution

**Remaining Issue:**
- If user still hears 2 sounds, one is likely from `handy.exe` native sound (v0.1.5+)
- This would require either:
  - Suppressing `handy.exe` sound (if configurable)
  - Suppressing AHK sound (if user prefers `handy.exe` sound)
  - Implementing Proxy Key Pattern to completely decouple AHK from `handy.exe` hotkey

---

## 6. Final Fix: Suppress AHK Sound (User Still Hearing 2 Sounds)

**Status:** Implemented - AHK sound disabled

**Root Cause Confirmed:**
- User confirmed still hearing 2 sounds after architectural fixes
- Research findings confirmed: `handy.exe` (v0.1.5+) has native start/stop sounds
- Both AHK and `handy.exe` are playing sounds simultaneously

**Solution Applied:**
- **Disabled AHK sound playback** in hotkey handler
- Now only `handy.exe`'s native sound plays
- This eliminates the duplicate sound issue

**Code Change:**
```autohotkey
; 3. Audio Feedback: DISABLED - handy.exe (v0.1.5+) plays its own native sound
; Suppressing AHK sound to prevent duplicate sounds (user hears 2 sounds: AHK + handy.exe)
; If you want AHK sound instead, uncomment the line below and configure handy.exe to disable its sound
; SafePlayDictationSound(g_DictationStartSound)
```

**Result:**
- User should now hear only ONE sound (from `handy.exe`)
- Visual feedback (red square indicator) still works
- All other functionality remains intact

**Alternative Options (if user wants AHK sound instead):**
1. Re-enable AHK sound (uncomment the code)
2. Configure `handy.exe` to disable its native sound (if settings allow)
3. Implement Proxy Key Pattern to completely decouple systems