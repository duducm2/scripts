# Automated Dictation Cycle Management: Complete Implementation Plan

## Executive Summary

This plan addresses the critical failure mode where ChatGPT dictation breaks after ~30 seconds of continuous audio input, causing complete transcription loss. The solution implements an automated timer-based restart system that prevents transcription from exceeding the 30-second threshold while maintaining seamless user experience.

---

## 1. Critical Problem Analysis

### 1.1 Failure Points

**Primary Failure:**
- ChatGPT's web-based dictation system has an undocumented ~30-second limit
- Exceeding this limit causes complete transcription failure with no recovery
- All progress is lost, requiring manual restart and re-speaking

**Secondary Failure Modes:**
- User cognitive load: Manual time tracking is error-prone and distracting
- Inconsistent restart timing: Human judgment introduces variability
- Context switching: Breaking flow to monitor time disrupts natural speech patterns

### 1.2 UX Concerns

**Current Pain Points:**
- **Cognitive overhead**: User must maintain awareness of elapsed time while speaking
- **Interruption anxiety**: Fear of losing progress creates hesitancy in natural speech
- **Recovery friction**: Manual restart requires keyboard interaction mid-thought
- **Uncertainty**: No clear feedback on when the 30-second threshold approaches

**Design Requirements:**
- Zero cognitive load: System handles timing automatically
- Seamless transitions: Restarts should be imperceptible to user workflow
- Clear feedback: Audio/visual cues indicate system actions
- Reliability: Must work consistently across different speech patterns

### 1.3 Robustness Considerations

**Keyboard Hotkey Limitations:**
- **Race conditions**: Multiple hotkey presses can queue or conflict
- **Timing precision**: System sleep delays may cause slight drift
- **Window focus**: ChatGPT window must be accessible for automation
- **Browser state**: Page must be loaded and responsive
- **Network latency**: Transcription processing time varies

**Error Recovery Needs:**
- Handle ChatGPT window not found
- Recover from failed hotkey triggers
- Detect stuck states (dictation not starting/stopping)
- Graceful degradation if automation fails

### 1.4 Limitations of Keyboard-Hotkey Approach

**Inherent Constraints:**
1. **No direct API access**: Must simulate user interactions via keyboard
2. **Timing uncertainty**: Cannot precisely measure actual transcription duration
3. **State inference**: Must infer dictation state from UI elements
4. **Browser dependency**: Relies on ChatGPT UI structure remaining stable
5. **Single-instance risk**: Multiple script instances could conflict

**Mitigation Strategies:**
- Use `#SingleInstance Force` (already present)
- Implement state machine with explicit transitions
- Add debouncing and lock mechanisms
- Monitor UI elements to verify state changes
- Implement timeout fallbacks

---

## 2. Design Strategy Comparison

### 2.1 Strategy A: Time-Based Automation (RECOMMENDED)

**Approach:**
- Start 30-second countdown timer when dictation begins
- Automatically trigger restart sequence at timer expiration
- Reset timer after each restart
- Continue until manual stop

**Pros:**
- ✅ Simple, predictable, reliable
- ✅ No external dependencies
- ✅ Guarantees 30-second maximum
- ✅ Low computational overhead
- ✅ Easy to implement and debug

**Cons:**
- ⚠️ May restart during mid-sentence (acceptable trade-off)
- ⚠️ Fixed timing doesn't adapt to speech patterns
- ⚠️ Requires accurate timer management

**Complexity:** Low  
**Reliability:** High  
**User Experience:** Good

---

### 2.2 Strategy B: Speech-Activity Detection Windows

**Approach:**
- Monitor microphone input levels or speech detection APIs
- Detect natural pauses in speech
- Trigger restart during silence gaps
- Use adaptive window sizing based on speech patterns

**Pros:**
- ✅ More natural timing (restarts during pauses)
- ✅ Adapts to user's speech rhythm
- ✅ Potentially better UX

**Cons:**
- ❌ Requires microphone API access (complex in AutoHotkey)
- ❌ May miss opportunities if user speaks continuously
- ❌ Higher complexity and failure modes
- ❌ Platform-dependent (Windows audio APIs)
- ❌ Still needs 30-second hard limit as fallback

**Complexity:** High  
**Reliability:** Medium  
**User Experience:** Excellent (if working)

---

### 2.3 Strategy C: Streaming Audio Segmentation

**Approach:**
- Intercept or monitor audio stream to ChatGPT
- Detect audio chunk boundaries
- Trigger restart at natural segment boundaries
- Maintain buffer of recent audio for seamless continuation

**Pros:**
- ✅ Most sophisticated approach
- ✅ Could preserve audio continuity
- ✅ Natural segmentation points

**Cons:**
- ❌ Extremely complex implementation
- ❌ Requires low-level audio interception
- ❌ May violate ChatGPT's terms of service
- ❌ Platform-specific audio stack knowledge needed
- ❌ High risk of breaking with browser updates

**Complexity:** Very High  
**Reliability:** Low  
**User Experience:** Excellent (if working)

---

### 2.4 Strategy D: Hybrid Timing + Detection

**Approach:**
- Primary: 30-second timer (hard limit)
- Secondary: Monitor UI for transcription completion events
- Tertiary: Detect natural pauses via simple heuristics (e.g., UI state changes)
- Combine signals to optimize restart timing

**Pros:**
- ✅ Best of both worlds
- ✅ Hard limit ensures safety
- ✅ Can optimize timing when possible

**Cons:**
- ⚠️ Increased complexity
- ⚠️ More failure modes
- ⚠️ May be over-engineered for the problem

**Complexity:** Medium-High  
**Reliability:** Medium-High  
**User Experience:** Excellent

---

## 3. Chosen Architecture: Time-Based Automation with State Machine

### 3.1 Justification

**Why Strategy A (Time-Based):**
1. **Reliability**: Simple timer logic is less prone to failure
2. **Guarantee**: Hard 30-second limit ensures transcription never exceeds threshold
3. **Implementation speed**: Can be built quickly with existing codebase patterns
4. **Maintainability**: Easy to understand and modify
5. **User acceptance**: Brief interruption every 30 seconds is acceptable trade-off

**Enhancement: State Machine Pattern**
- Explicit state tracking prevents race conditions
- Clear transitions enable reliable error recovery
- Makes debugging and logging straightforward
- Supports future extensions (e.g., pause/resume)

### 3.2 Architecture Overview

```
┌─────────────────────────────────────────────────────────────┐
│                    Dictation Manager                        │
│  ┌──────────────┐      ┌──────────────┐      ┌──────────┐ │
│  │   State      │◄────►│    Timer     │◄────►│ Hotkey   │ │
│  │   Machine    │      │   Manager    │      │ Handler  │ │
│  └──────────────┘      └──────────────┘      └──────────┘ │
│         │                      │                   │        │
│         └──────────────────────┴───────────────────┘        │
│                            │                                 │
│                    ┌───────▼────────┐                        │
│                    │  UI Monitor    │                        │
│                    │  (Optional)    │                        │
│                    └────────────────┘                        │
└─────────────────────────────────────────────────────────────┘
```

**Core Components:**
1. **State Machine**: Tracks dictation lifecycle (idle → active → restarting → active)
2. **Timer Manager**: 30-second countdown with reset capability
3. **Hotkey Orchestrator**: Coordinates Win+Alt+Shift+0/7 triggers
4. **UI Monitor**: Optional verification of dictation state
5. **Error Handler**: Recovery and fallback mechanisms

---

## 4. Algorithmic Flow

### 4.1 State Machine Definition

```
States:
- IDLE: No dictation active
- ACTIVE: Dictation running, timer counting
- RESTARTING: Executing restart sequence (stop → wait → start)
- PAUSED: Temporarily stopped (future feature)
- ERROR: Recovery state

Transitions:
IDLE → ACTIVE: User presses Win+Alt+Shift+7
ACTIVE → RESTARTING: Timer reaches 30 seconds OR transcription finishes (autosend mode)
RESTARTING → ACTIVE: Restart sequence completes successfully
ACTIVE → IDLE: User manually stops (Win+Alt+Shift+0 or Win+Alt+Shift+7)
RESTARTING → IDLE: User manually stops during restart
ANY → ERROR: Automation failure detected
ERROR → IDLE: Recovery successful or user intervention
```

### 4.2 Pseudocode: Main Timer Loop

```pseudocode
GLOBAL STATE: currentState = IDLE
GLOBAL TIMER: dictationTimer = null
GLOBAL FLAG: isAutoRestartEnabled = false
GLOBAL FLAG: autoSendMode = false

FUNCTION OnWinAltShift7():
    IF currentState == IDLE:
        autoSendMode = true
        StartDictation()
    ELSE IF currentState == ACTIVE:
        autoSendMode = true
        StopDictation()  // Manual stop overrides auto-restart
    END IF
END FUNCTION

FUNCTION OnWinAltShift0():
    IF currentState == ACTIVE OR currentState == RESTARTING:
        isAutoRestartEnabled = false  // Disable auto-restart on manual stop
        StopDictation()
    END IF
END FUNCTION

FUNCTION StartDictation():
    currentState = ACTIVE
    isAutoRestartEnabled = true
    autoSendMode = (determined by hotkey)
    
    // Call existing ToggleDictation() function
    ToggleDictation(autoSendMode)
    
    // Start 30-second timer
    dictationTimer = SetTimer(OnTimerExpired, -30000)  // 30 seconds, one-shot
    dictationStartTime = A_TickCount
END FUNCTION

FUNCTION OnTimerExpired():
    IF currentState == ACTIVE AND isAutoRestartEnabled:
        currentState = RESTARTING
        ExecuteRestartSequence()
    END IF
END FUNCTION

FUNCTION ExecuteRestartSequence():
    // Step 1: Stop current dictation
    // Trigger Win+Alt+Shift+0 programmatically (without auto-send)
    SendHotkey("#!+0")
    Sleep(500)  // Wait for stop to process
    
    // Step 2: Wait for transcription to finish (if needed)
    IF WaitForComposerSubmitButton(5000):
        // Transcription completed
        // Don't send - instead restart
    END IF
    
    // Step 3: Small delay to ensure UI is ready
    Sleep(200)
    
    // Step 4: Start dictation again
    SendHotkey("#!+0")  // Second press starts new dictation
    Sleep(500)
    
    // Step 5: Verify restart succeeded
    IF VerifyDictationActive():
        currentState = ACTIVE
        // Reset timer for next 30-second cycle
        dictationTimer = SetTimer(OnTimerExpired, -30000)
        dictationStartTime = A_TickCount
    ELSE:
        currentState = ERROR
        HandleRestartFailure()
    END IF
END FUNCTION

FUNCTION StopDictation():
    isAutoRestartEnabled = false
    SetTimer(dictationTimer, 0)  // Cancel timer
    
    IF currentState == ACTIVE:
        ToggleDictation(autoSendMode)
        currentState = IDLE
    ELSE IF currentState == RESTARTING:
        // Interrupt restart sequence
        SetTimer(dictationTimer, 0)
        currentState = IDLE
    END IF
END FUNCTION

FUNCTION VerifyDictationActive():
    // Check if dictation indicator is visible
    // OR check if ChatGPT UI shows dictation is active
    RETURN (dictationIndicatorVisible OR UIElementExists("dictation_active"))
END FUNCTION
```

### 4.3 Autosend Mode Special Handling

```pseudocode
FUNCTION OnTranscriptionFinished():
    // Called when WaitForComposerSubmitButton() detects completion
    IF autoSendMode AND isAutoRestartEnabled:
        // Instead of sending, restart dictation
        currentState = RESTARTING
        ExecuteRestartSequence()
    ELSE IF autoSendMode AND NOT isAutoRestartEnabled:
        // Normal autosend behavior
        SendTranscription()
    END IF
END FUNCTION
```

---

## 5. Hotkey Orchestration Plan

### 5.1 Hotkey Conflict Prevention

**Problem:** Multiple hotkey triggers can cause race conditions or double-firing.

**Solution: Debouncing and State Locks**

```pseudocode
GLOBAL LOCK: hotkeyLock = false
GLOBAL TIMESTAMP: lastHotkeyTime = 0
CONSTANT: HOTKEY_DEBOUNCE_MS = 500

FUNCTION SendHotkey(keys):
    // Prevent rapid-fire hotkeys
    IF (A_TickCount - lastHotkeyTime) < HOTKEY_DEBOUNCE_MS:
        RETURN false
    END IF
    
    // Prevent concurrent execution
    IF hotkeyLock:
        RETURN false
    END IF
    
    hotkeyLock = true
    lastHotkeyTime = A_TickCount
    
    TRY:
        Send(keys)
        Sleep(100)  // Small delay for processing
    FINALLY:
        hotkeyLock = false
    END TRY
    
    RETURN true
END FUNCTION
```

### 5.2 Programmatic Hotkey Triggering

**Challenge:** AutoHotkey v2 doesn't directly support triggering hotkey labels from code.

**Solution:** Extract hotkey logic into reusable functions.

```autohotkey
; Current structure (problematic):
#!+0:: {
    ToggleDictation(false)
}

; Refactored structure (solution):
#!+0:: {
    HandleDictationToggle(false, "manual")
}

FUNCTION HandleDictationToggle(autoSend, source):
    // source: "manual" | "auto_restart" | "timer"
    IF source == "auto_restart":
        // Skip some checks or use different flow
    END IF
    
    ToggleDictation(autoSend)
END FUNCTION

FUNCTION TriggerDictationToggleProgrammatically(autoSend):
    HandleDictationToggle(autoSend, "auto_restart")
END FUNCTION
```

### 5.3 Restart Sequence Timing

**Critical Timing Windows:**
1. **Stop dictation**: ~200-500ms for UI to process
2. **Wait for transcription**: 0-5 seconds (variable)
3. **Start dictation**: ~200-500ms for UI to process
4. **Total restart time**: ~1-6 seconds

**Optimization:**
- Use `WaitForComposerSubmitButton()` with short timeout (2-3 seconds)
- If timeout, proceed anyway (transcription may be quick)
- Don't wait indefinitely - prioritize keeping cycle under 30 seconds

---

## 6. Reliability Mechanisms

### 6.1 Debouncing

**Purpose:** Prevent rapid-fire timer triggers or hotkey conflicts.

**Implementation:**
```pseudocode
GLOBAL: lastRestartTime = 0
CONSTANT: MIN_RESTART_INTERVAL = 2000  // 2 seconds minimum between restarts

FUNCTION ExecuteRestartSequence():
    IF (A_TickCount - lastRestartTime) < MIN_RESTART_INTERVAL:
        RETURN  // Ignore if too soon
    END IF
    
    lastRestartTime = A_TickCount
    // ... proceed with restart
END FUNCTION
```

### 6.2 Error Recovery

**Failure Scenarios:**
1. ChatGPT window not found
2. Dictation button not accessible
3. Restart sequence times out
4. UI state inconsistent

**Recovery Strategy:**
```pseudocode
FUNCTION ExecuteRestartSequence():
    maxRetries = 3
    retryCount = 0
    
    WHILE retryCount < maxRetries:
        TRY:
            IF AttemptRestart():
                RETURN true  // Success
            END IF
        CATCH error:
            LogError(error)
            retryCount++
            Sleep(1000 * retryCount)  // Exponential backoff
        END TRY
    END WHILE
    
    // All retries failed
    currentState = ERROR
    NotifyUser("Auto-restart failed. Please restart manually.")
    RETURN false
END FUNCTION

FUNCTION AttemptRestart():
    // Verify ChatGPT window exists
    IF NOT GetChatGPTWindowHwnd():
        RETURN false
    END IF
    
    // Execute restart steps with timeouts
    IF NOT StopDictationWithTimeout(3000):
        RETURN false
    END IF
    
    IF NOT StartDictationWithTimeout(3000):
        RETURN false
    END IF
    
    RETURN true
END FUNCTION
```

### 6.3 Timer Reset and Synchronization

**Problem:** Timer drift or missed triggers.

**Solution:**
```pseudocode
GLOBAL: dictationStartTime = 0
CONSTANT: MAX_DICTATION_TIME_MS = 30000

FUNCTION StartDictation():
    dictationStartTime = A_TickCount
    // Use one-shot timer
    SetTimer(CheckTimerExpired, 100)  // Check every 100ms for precision
END FUNCTION

FUNCTION CheckTimerExpired():
    elapsed = A_TickCount - dictationStartTime
    
    IF elapsed >= MAX_DICTATION_TIME_MS:
        SetTimer(CheckTimerExpired, 0)  // Stop checking
        OnTimerExpired()
    END IF
END FUNCTION
```

**Alternative: Precise Timer**
```pseudocode
FUNCTION StartDictation():
    dictationStartTime = A_TickCount
    // One-shot timer, more efficient
    SetTimer(OnTimerExpired, -30000)
END FUNCTION
```

### 6.4 State Verification

**Purpose:** Ensure UI state matches expected state before/after operations.

```pseudocode
FUNCTION VerifyDictationState(expectedState):
    actualState = DetectDictationStateFromUI()
    
    IF actualState != expectedState:
        LogWarning("State mismatch: expected " + expectedState + ", got " + actualState)
        RETURN false
    END IF
    
    RETURN true
END FUNCTION

FUNCTION DetectDictationStateFromUI():
    // Check if dictation indicator is visible
    IF dictationIndicatorVisible:
        RETURN "ACTIVE"
    END IF
    
    // Check ChatGPT UI for dictation button state
    // (would require UIA inspection)
    
    RETURN "UNKNOWN"
END FUNCTION
```

---

## 7. Extensibility Design

### 7.1 Configurable Parameters

**Future Customization Points:**
```pseudocode
GLOBAL CONFIG:
    MAX_DICTATION_SECONDS = 30  // User-configurable
    RESTART_TIMEOUT_MS = 5000
    ENABLE_AUTO_RESTART = true
    RESTART_DURING_PAUSES_ONLY = false  // Future: speech detection
    SHOW_RESTART_NOTIFICATIONS = false
```

### 7.2 Dynamic Window Size (Future)

**Concept:** Adapt timer based on speech patterns.

```pseudocode
FUNCTION AdaptiveTimerManager():
    recentDurations = []  // Track last N dictation segments
    
    FUNCTION CalculateOptimalWindow():
        IF recentDurations.Length < 3:
            RETURN 30  // Default
        END IF
        
        average = Average(recentDurations)
        // Use 90% of average to be safe
        RETURN Max(20, Min(30, average * 0.9))
    END FUNCTION
END FUNCTION
```

### 7.3 Floating UI Indicator (Future)

**Concept:** Visual countdown timer showing time remaining.

```pseudocode
FUNCTION ShowRestartCountdown():
    remainingSeconds = (MAX_DICTATION_TIME_MS - elapsed) / 1000
    UpdateIndicator("Dictation: " + Round(remainingSeconds) + "s remaining")
END FUNCTION
```

### 7.4 Pause/Resume Support (Future)

**Concept:** Allow user to pause auto-restart temporarily.

```pseudocode
FUNCTION PauseAutoRestart():
    isAutoRestartEnabled = false
    SetTimer(CheckTimerExpired, 0)  // Pause timer
    currentState = PAUSED
END FUNCTION

FUNCTION ResumeAutoRestart():
    isAutoRestartEnabled = true
    // Resume timer from where it left off
    remainingTime = MAX_DICTATION_TIME_MS - (A_TickCount - dictationStartTime)
    SetTimer(CheckTimerExpired, -remainingTime)
    currentState = ACTIVE
END FUNCTION
```

---

## 8. Implementation Blueprint

### 8.1 File Structure

```
ChatGPT.ahk (existing file)
├── Add new global variables (Section 8.2)
├── Modify ToggleDictation() function (Section 8.3)
├── Add new functions (Section 8.4)
└── Modify hotkey handlers (Section 8.5)
```

### 8.2 Global Variables

```autohotkey
; =============================================================================
; Auto-Restart Dictation Manager - Global State
; =============================================================================
global g_dictationState := "IDLE"  ; IDLE | ACTIVE | RESTARTING | ERROR
global g_dictationStartTime := 0
global g_autoRestartEnabled := false
global g_autoRestartTimer := ""
global g_autoSendMode := false
global g_hotkeyLock := false
global g_lastHotkeyTime := 0
global CONST_MAX_DICTATION_MS := 30000  ; 30 seconds
global CONST_HOTKEY_DEBOUNCE_MS := 500
```

### 8.3 Modified ToggleDictation() Function

**Changes Required:**
1. Add state machine integration
2. Start timer when dictation starts
3. Handle autosend mode differently (restart instead of send)
4. Cancel timer when dictation stops

**Key Modifications:**
```autohotkey
ToggleDictation(autoSend, source := "manual") {
    static isDictating := false
    global g_transcribeChimePending
    global g_dictationState
    global g_autoRestartEnabled
    global g_autoRestartTimer
    global g_autoSendMode
    global g_dictationStartTime

    ; ... existing window activation code ...

    action := !isDictating ? "start" : "stop"
    g_autoSendMode := autoSend

    if (action = "start") {
        ; ... existing start code ...
        
        isDictating := true
        g_dictationState := "ACTIVE"
        
        ; Start auto-restart timer if enabled
        if (g_autoRestartEnabled || source = "auto_restart") {
            StartAutoRestartTimer()
        }
        
        ; ... rest of start code ...
    } else if (action = "stop") {
        ; Cancel auto-restart timer
        StopAutoRestartTimer()
        
        ; ... existing stop code ...
        
        isDictating := false
        g_dictationState := "IDLE"
        
        ; Special handling for autosend mode with auto-restart
        if (autoSend && g_autoRestartEnabled && source != "manual") {
            ; Don't send - instead restart
            Sleep(500)
            ExecuteAutoRestartSequence()
            return
        }
        
        ; ... rest of stop code (normal autosend handling) ...
    }
}
```

### 8.4 New Functions to Add

```autohotkey
; =============================================================================
; Auto-Restart Timer Management
; =============================================================================
StartAutoRestartTimer() {
    global g_dictationStartTime
    global g_autoRestartTimer
    global CONST_MAX_DICTATION_MS
    
    g_dictationStartTime := A_TickCount
    ; One-shot timer
    g_autoRestartTimer := SetTimer(OnAutoRestartTimerExpired, -CONST_MAX_DICTATION_MS)
}

StopAutoRestartTimer() {
    global g_autoRestartTimer
    if (g_autoRestartTimer) {
        SetTimer(g_autoRestartTimer, 0)
        g_autoRestartTimer := ""
    }
}

OnAutoRestartTimerExpired() {
    global g_dictationState
    global g_autoRestartEnabled
    
    if (g_dictationState = "ACTIVE" && g_autoRestartEnabled) {
        g_dictationState := "RESTARTING"
        ExecuteAutoRestartSequence()
    }
}

; =============================================================================
; Auto-Restart Sequence Execution
; =============================================================================
ExecuteAutoRestartSequence() {
    global g_dictationState
    global g_autoRestartEnabled
    global g_hotkeyLock
    global g_lastHotkeyTime
    global CONST_HOTKEY_DEBOUNCE_MS
    
    ; Debounce check
    if ((A_TickCount - g_lastHotkeyTime) < CONST_HOTKEY_DEBOUNCE_MS) {
        return
    }
    
    ; Lock check
    if (g_hotkeyLock) {
        return
    }
    
    g_hotkeyLock := true
    g_lastHotkeyTime := A_TickCount
    
    try {
        ; Step 1: Stop current dictation (without autosend)
        ToggleDictation(false, "auto_restart")
        
        ; Step 2: Wait for transcription to finish (short timeout)
        if (WaitForComposerSubmitButton(3000)) {
            ; Transcription completed
        }
        
        ; Step 3: Small delay
        Sleep(200)
        
        ; Step 4: Start dictation again
        ToggleDictation(false, "auto_restart")
        
        ; Step 5: Verify and restart timer
        Sleep(500)
        if (VerifyDictationActive()) {
            g_dictationState := "ACTIVE"
            StartAutoRestartTimer()
        } else {
            g_dictationState := "ERROR"
            ShowNotification("Auto-restart failed", 2000, "DF2935", "FFFFFF")
        }
    } catch Error as e {
        g_dictationState := "ERROR"
        ShowNotification("Auto-restart error", 2000, "DF2935", "FFFFFF")
    } finally {
        g_hotkeyLock := false
    }
}

VerifyDictationActive() {
    ; Check if dictation indicator is visible
    global smallLoadingGuis
    return (smallLoadingGuis.Length > 0)
}

; =============================================================================
; Hotkey Handler Modifications
; =============================================================================
HandleDictationToggleWithAutoRestart(autoSend) {
    global g_autoRestartEnabled
    global g_dictationState
    
    if (g_dictationState = "IDLE") {
        ; Starting dictation - enable auto-restart
        g_autoRestartEnabled := true
        ToggleDictation(autoSend, "manual")
    } else {
        ; Stopping dictation - disable auto-restart
        g_autoRestartEnabled := false
        ToggleDictation(autoSend, "manual")
    }
}
```

### 8.5 Modified Hotkey Handlers

```autohotkey
; =============================================================================
; Toggle Dictation (No Auto-Send)
; Hotkey: Win+Alt+Shift+0
; =============================================================================
#!+0::
{
    global g_autoRestartEnabled
    g_autoRestartEnabled := false  ; Disable auto-restart on manual stop
    ToggleDictation(false, "manual")
}

; =============================================================================
; Toggle Dictation (with Auto-Send) - WITH AUTO-RESTART
; Hotkey: Win+Alt+Shift+7
; =============================================================================
#!+7::
{
    HandleDictationToggleWithAutoRestart(true)
}
```

### 8.6 Event Listeners

**AutoHotkey v2 Timer Pattern:**
```autohotkey
; Timer is set using SetTimer() with callback function
; No separate "event listener" needed - timer callback serves this purpose
```

**UI State Monitoring (Optional):**
```autohotkey
; Could add periodic UI check if needed
; SetTimer(MonitorDictationUI, 1000)  ; Check every second
; But this adds overhead - timer-based approach is sufficient
```

### 8.7 Integration Points

**Existing Functions to Leverage:**
- `ToggleDictation()` - Core dictation logic (modify)
- `WaitForComposerSubmitButton()` - Detect transcription completion (reuse)
- `ShowDictationIndicator()` / `HideDictationIndicator()` - State verification (reuse)
- `PlayDictationStartedChime()` - Audio feedback (reuse)

**New Dependencies:**
- None - uses existing AutoHotkey v2 features only

---

## 9. UX Evaluation: Moment-to-Moment Experience

### 9.1 Starting Dictation (Win+Alt+Shift+7)

**User Action:** Press Win+Alt+Shift+7

**System Response:**
1. ChatGPT window activates (if not already active)
2. Dictation starts (existing behavior)
3. **NEW:** 30-second timer begins silently
4. Dictation indicator appears ("Dictation ON")
5. Start chime plays
6. User returns to previous window
7. User begins speaking

**User Experience:**
- ✅ No change from current behavior
- ✅ Timer runs invisibly in background
- ✅ No additional cognitive load

---

### 9.2 During Active Dictation (0-30 seconds)

**User Action:** Speaking continuously

**System Response:**
1. Timer counts down silently
2. Dictation indicator remains visible
3. No interruptions or warnings

**User Experience:**
- ✅ Completely transparent
- ✅ User can speak naturally without time awareness
- ✅ No distractions

---

### 9.3 Automatic Restart (at 30 seconds)

**User Action:** Still speaking (may be mid-sentence)

**System Response:**
1. **T+30s:** Timer expires
2. System automatically:
   - Stops current dictation (Win+Alt+Shift+0)
   - Waits for transcription to finish (0-3 seconds)
   - Starts new dictation (Win+Alt+Shift+0 again)
   - Resets 30-second timer
3. Restart chime plays (from existing ToggleDictation function)
4. Dictation indicator remains visible (brief flicker possible)

**User Experience:**
- ⚠️ Brief interruption (1-3 seconds)
- ✅ Audio cue (chime) indicates restart occurred
- ✅ Can continue speaking immediately after chime
- ⚠️ May need to repeat last few words if caught mid-sentence
- ✅ No manual intervention required

**Mitigation for Mid-Sentence Interruption:**
- User can pause briefly before 30-second mark if aware
- System could be enhanced to detect natural pauses (future)
- Trade-off: Brief interruption vs. complete transcription loss

---

### 9.4 Multiple Restart Cycles

**User Action:** Speaking for 2+ minutes continuously

**System Response:**
- Automatic restart every 30 seconds
- Each restart follows same sequence
- Timer resets after each restart
- Dictation indicator stays visible throughout

**User Experience:**
- ✅ Predictable rhythm (restart every 30s)
- ✅ Can plan pauses around restart times (optional)
- ✅ No risk of transcription failure
- ⚠️ Multiple brief interruptions (acceptable trade-off)

---

### 9.5 Manual Stop (Win+Alt+Shift+0 or Win+Alt+Shift+7)

**User Action:** Press stop hotkey

**System Response:**
1. Auto-restart timer cancels immediately
2. Dictation stops (existing behavior)
3. Transcription finishes
4. If autosend was enabled: sends transcription
5. Dictation indicator hides
6. Completion chime plays

**User Experience:**
- ✅ Immediate response
- ✅ Auto-restart stops (no unwanted restarts)
- ✅ Normal dictation flow preserved

---

### 9.6 Error Scenarios

**Scenario A: ChatGPT Window Not Found**
- **System:** Shows error notification, enters ERROR state
- **User:** Sees brief error message, can retry manually
- **Recovery:** User opens ChatGPT and tries again

**Scenario B: Restart Sequence Fails**
- **System:** Shows error notification, stops auto-restart
- **User:** Sees error, can manually restart if needed
- **Recovery:** System attempts retry (3 attempts), then stops

**Scenario C: Browser Becomes Unresponsive**
- **System:** Timeout after 5 seconds, shows error
- **User:** Aware of issue, can refresh page
- **Recovery:** Manual intervention required

---

## 10. Testing Strategy

### 10.1 Unit Testing Scenarios

1. **Timer Accuracy**
   - Start timer, verify it fires at exactly 30 seconds
   - Test with system under load (CPU-intensive tasks)

2. **State Transitions**
   - Verify all state transitions are valid
   - Test invalid transitions are rejected

3. **Hotkey Debouncing**
   - Rapid-fire hotkey presses should be ignored
   - Verify lock mechanism works

### 10.2 Integration Testing Scenarios

1. **Normal Flow**
   - Start dictation → wait 30s → verify restart → continue
   - Multiple cycles (2+ minutes of dictation)

2. **Manual Interruption**
   - Start dictation → manually stop at 15s → verify timer cancels
   - Start dictation → manually stop during restart → verify clean stop

3. **Autosend Mode**
   - Start with autosend → wait for transcription → verify restart instead of send

4. **Error Recovery**
   - Close ChatGPT window → verify error handling
   - Simulate UI unresponsiveness → verify timeout behavior

### 10.3 Edge Cases

1. **Very Fast Speech**: User speaks rapidly, restart occurs mid-sentence
2. **Very Slow Speech**: User pauses frequently, restart may occur during pause
3. **Multiple Instances**: Verify #SingleInstance prevents conflicts
4. **System Sleep**: Computer sleeps during dictation → verify timer behavior
5. **Window Focus Loss**: ChatGPT loses focus → verify automation still works

---

## 11. Implementation Checklist

### Phase 1: Core Timer System
- [ ] Add global variables for state management
- [ ] Implement `StartAutoRestartTimer()` function
- [ ] Implement `StopAutoRestartTimer()` function
- [ ] Implement `OnAutoRestartTimerExpired()` callback
- [ ] Test timer accuracy and reliability

### Phase 2: Restart Sequence
- [ ] Implement `ExecuteAutoRestartSequence()` function
- [ ] Add debouncing and locking mechanisms
- [ ] Integrate with existing `ToggleDictation()` function
- [ ] Test restart sequence end-to-end

### Phase 3: Hotkey Integration
- [ ] Modify Win+Alt+Shift+7 handler to enable auto-restart
- [ ] Modify Win+Alt+Shift+0 handler to disable auto-restart
- [ ] Add `HandleDictationToggleWithAutoRestart()` function
- [ ] Test hotkey interactions

### Phase 4: Autosend Mode Handling
- [ ] Modify `ToggleDictation()` to detect autosend + auto-restart
- [ ] Implement restart-on-transcription-complete logic
- [ ] Test autosend flow with auto-restart

### Phase 5: Error Handling
- [ ] Add error recovery mechanisms
- [ ] Implement retry logic with exponential backoff
- [ ] Add user notifications for errors
- [ ] Test error scenarios

### Phase 6: Testing & Refinement
- [ ] Run all test scenarios
- [ ] Performance testing (timer accuracy under load)
- [ ] User acceptance testing
- [ ] Refine timing and error messages based on feedback

---

## 12. Risk Assessment & Mitigation

### 12.1 Technical Risks

| Risk | Probability | Impact | Mitigation |
|------|------------|--------|------------|
| Timer drift/inaccuracy | Low | Medium | Use `A_TickCount` for precise timing |
| Hotkey conflicts | Medium | High | Implement debouncing and locks |
| UI state detection fails | Medium | Medium | Add fallback verification methods |
| Browser updates break automation | Low | High | Use robust UIA selectors, test regularly |
| Restart sequence times out | Low | Medium | Implement retry logic with timeouts |

### 12.2 UX Risks

| Risk | Probability | Impact | Mitigation |
|------|------------|--------|------------|
| Mid-sentence interruption | High | Low | Acceptable trade-off; user can pause |
| Restart chime is confusing | Low | Low | Use existing chime (user already familiar) |
| User doesn't understand auto-restart | Medium | Low | System is transparent; no explanation needed |
| Too frequent restarts | Low | Low | 30-second window is reasonable |

### 12.3 Operational Risks

| Risk | Probability | Impact | Mitigation |
|------|------------|--------|------------|
| Script crashes | Low | High | Comprehensive error handling |
| Multiple script instances | Low | High | #SingleInstance Force (already present) |
| ChatGPT UI changes | Medium | High | Use flexible UIA selectors, monitor updates |

---

## 13. Success Criteria

### 13.1 Functional Requirements
- ✅ Dictation never exceeds 30 seconds
- ✅ Automatic restart occurs every 30 seconds
- ✅ Manual stop works immediately
- ✅ Autosend mode restarts instead of sending
- ✅ System recovers from errors gracefully

### 13.2 Performance Requirements
- ✅ Timer accuracy: ±100ms
- ✅ Restart sequence: < 5 seconds
- ✅ No noticeable performance impact on system
- ✅ Hotkey response: < 100ms

### 13.3 UX Requirements
- ✅ Zero cognitive load (automatic operation)
- ✅ Minimal interruption (1-3 seconds per restart)
- ✅ Clear audio feedback (existing chimes)
- ✅ Reliable operation (99%+ success rate)

---

## 14. Future Enhancements (Post-MVP)

1. **Adaptive Timing**: Adjust window size based on speech patterns
2. **Pause Detection**: Restart during natural speech pauses
3. **Visual Countdown**: Optional floating timer showing seconds remaining
4. **Statistics**: Track average segment length, restart frequency
5. **Configurable Threshold**: User-adjustable 30-second limit
6. **Speech Activity Detection**: Integrate microphone level monitoring
7. **Multi-language Support**: Adapt to different speech rhythms

---

## 15. Conclusion

This plan provides a **reliable, time-based automation system** that solves the 30-second transcription failure problem with minimal complexity and maximum reliability. The chosen architecture prioritizes:

1. **Simplicity**: Timer-based approach is easy to understand and maintain
2. **Reliability**: Hard 30-second limit guarantees transcription safety
3. **User Experience**: Transparent operation with minimal interruption
4. **Extensibility**: State machine design supports future enhancements

The implementation can be completed incrementally, with each phase building on the previous one. The system integrates seamlessly with existing code while adding robust automation capabilities.

**Recommended Implementation Order:**
1. Phase 1-2: Core timer and restart sequence (MVP)
2. Phase 3-4: Hotkey integration and autosend handling
3. Phase 5-6: Error handling and testing

This approach delivers immediate value while maintaining a clear path for future improvements.

