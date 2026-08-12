; =============================================================================
; Utils module: dictation_toggle.ahk
; Dictation indicator, ~#!+0 hotkey, ToggleDictationMode
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; Dictation Indicator - Opaque language flag (no red square)
; Top-center of the active window (clamped inside); top-center of every other
; monitor's work area. Follows focus/window moves via a short timer.
; Handy slot 1=EN, 2=PT, 3=EN+PT. Toggles with Win+Alt+Shift+0.
; =============================================================================

; Global variables for dictation indicator
global g_DictationActive := false
global g_DictationPulseTimer := false
global g_DictationCheckTimer := false  ; Timer to check if Recording window still exists
global g_DictationCompletionChimeScheduled := false  ; Flag to prevent multiple completion chimes
global g_LastDictationSoundTick := 0  ; Timestamp of last dictation sound to throttle audio output
global g_DictationStartSound := A_ScriptDir . "\assets\sounds\speach-start.wav"
global g_DictationStopSound := A_ScriptDir . "\assets\sounds\speach-finished.wav"
global g_PendingDictationAction := ""  ; Action to execute after transcription: "Paste" (reserved for future)
global g_PendingGeminiPromptAfterDictation := false  ; When set by ~#!+0 stop, show "Send dictation? Y (4s)" after completion
global g_D2C_DictationSubmitMenuCycleFinished := false  ; After V/W/E/N/timeout/F/O: block stray second StartFromDictation for this wave
global g_DictationGeminiConfirmBannerVisible := false  ; Guard: only one "Send dictation?" banner at a time
global g_KeepIndicatorVisible := false  ; Flag to keep indicator visible until paste action completes
global g_LastStateTransitionTick := 0  ; Timestamp of last state transition to prevent rapid re-detection
global g_DictationSoundPlayed := false  ; Atomic test-and-set: one start chime per session
global g_DictationStartClipboardText := "" ; Track clipboard content at start to detect changes
global g_DictationHotkeyOwnerHandle := 0 ; Named mutex handle for cross-process single-owner dictation hotkey
global g_DictationHotkeyIsOwner := false ; True only in the single process that owns dictation hotkey handling
global g_DictationFlagGuis := []  ; Recording language flags, one GUI per monitor
global g_DictationFlagSlot := -1  ; Slot last shown (-1 = none; 0 = unknown/? fallback)
global g_DictationFlagFollowCache := ""  ; Skip redundant flag Move when geometry/slot unchanged

; Ensure only one script process handles the dictation hotkey logic.
InitializeDictationHotkeyOwnership() {
    global g_DictationHotkeyOwnerHandle, g_DictationHotkeyIsOwner
    mutexName := "Local\D2C_Dictation_Hotkey_Owner"
    hMutex := DllCall("CreateMutex", "Ptr", 0, "Int", 0, "Str", mutexName, "Ptr")
    if (!hMutex) {
        g_DictationHotkeyIsOwner := false
        return
    }
    err := DllCall("GetLastError", "UInt")
    if (err = 183) { ; ERROR_ALREADY_EXISTS
        g_DictationHotkeyIsOwner := false
        DllCall("CloseHandle", "Ptr", hMutex)
        return
    }
    g_DictationHotkeyOwnerHandle := hMutex
    g_DictationHotkeyIsOwner := true
}

ReleaseDictationHotkeyOwnership(*) {
    global g_DictationHotkeyOwnerHandle
    if (g_DictationHotkeyOwnerHandle) {
        try DllCall("CloseHandle", "Ptr", g_DictationHotkeyOwnerHandle)
        g_DictationHotkeyOwnerHandle := 0
    }
}

InitializeDictationHotkeyOwnership()
OnExit(ReleaseDictationHotkeyOwnership)

; Debug logging helper for dictation workflow
LogDebug(sessionId, runId, hypothesisId, location, message, data := "") {
    logPath := A_ScriptDir "\.cursor\debug.log"
    timestamp := A_Now "." Format("{:03}", A_MSec)
    logEntry := Format(
        '{{"sessionId":"{}","runId":"{}","hypothesisId":"{}","location":"{}","message":"{}","timestamp":"{}","data":{}}}',
        sessionId, runId, hypothesisId, location, message, timestamp, data ? '"' . data . '"' : '""')
    try {
        FileAppend(logEntry . "`n", logPath)
    } catch {
        ; Silently ignore logging errors
    }
}

; Constants for dictation indicator
global DICTATION_SQUARE_SIZE := 105  ; Flag height in px (150 minus 30%; aspect preserved)
global DICTATION_PULSE_INTERVAL := 50 ; Timer interval in ms (follow active window)

; Get the monitor that contains the active window
; Returns monitor index (1-based) or 0 if not found
GetDictationActiveMonitor() {
    hwnd := WinExist("A")
    if (!hwnd) {
        return 1  ; Default to primary monitor
    }

    rect := Buffer(16, 0)
    if (!DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect)) {
        return 1  ; Default to primary monitor
    }

    left := NumGet(rect, 0, "int")
    top := NumGet(rect, 4, "int")
    right := NumGet(rect, 8, "int")
    bottom := NumGet(rect, 12, "int")

    centerX := left + (right - left) // 2
    centerY := top + (bottom - top) // 2

    monitorCount := MonitorGetCount()
    loop monitorCount {
        MonitorGet(A_Index, &ml, &mt, &mr, &mb)
        if (centerX >= ml && centerX <= mr && centerY >= mt && centerY <= mb) {
            return A_Index
        }
    }

    return 1  ; Default to primary monitor
}

; Active-window screen rect used to clamp the recording flag.
; Returns false when there is no usable window (caller should use the monitor work area).
GetDictationActiveWindowRect(&wl, &wt, &wr, &wb) {
    hwnd := WinExist("A")
    if (hwnd) {
        try {
            if (WinGetMinMax("ahk_id " hwnd) = -1)
                hwnd := 0
        } catch {
            hwnd := 0
        }
    }
    if (!hwnd)
        return false
    rect := Buffer(16, 0)
    if (!DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect))
        return false
    wl := NumGet(rect, 0, "int")
    wt := NumGet(rect, 4, "int")
    wr := NumGet(rect, 8, "int")
    wb := NumGet(rect, 12, "int")
    return (wr - wl >= 8 && wb - wt >= 8)
}

DictationFlag_SlotLabel(slot) {
    return (slot = 1) ? "EN" : (slot = 2) ? "PT" : (slot = 3) ? "EN+PT" : "?"
}

; Opaque recording flag. Do not use WS_EX_TRANSPARENT (+E0x20): it suppresses painting.
DictationFlag_CreateGui(slot, imagePath) {
    global DICTATION_SQUARE_SIZE

    flagGui := Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale")
    flagGui.BackColor := "313244"
    flagGui.MarginX := 0
    flagGui.MarginY := 0

    usedPicture := false
    if (imagePath != "") {
        try {
            flagGui.Add("Picture", "h" . DICTATION_SQUARE_SIZE . " w-1", imagePath)
            usedPicture := true
        } catch {
            usedPicture := false
        }
    }
    if !usedPicture {
        flagGui.SetFont("s20 cFFFFFF Bold", "Segoe UI")
        flagGui.Add("Text", "Center w" . DICTATION_SQUARE_SIZE . " h" . DICTATION_SQUARE_SIZE . " Background45475A",
            DictationFlag_SlotLabel(slot))
    }
    flagGui.Show("AutoSize Hide")
    return flagGui
}

DictationFlag_Hide() {
    global g_DictationFlagGuis, g_DictationFlagSlot, g_DictationFlagFollowCache
    for item in g_DictationFlagGuis {
        try {
            if IsObject(item.gui)
                item.gui.Destroy()
        } catch {
        }
    }
    g_DictationFlagGuis := []
    g_DictationFlagSlot := -1
    g_DictationFlagFollowCache := ""
}

DictationFlag_MoveGui(flagGui, guiX, guiY) {
    try {
        flagGui.Move(guiX, guiY)
        hwnd := flagGui.Hwnd
        if (hwnd) {
            ; SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE = 0x0001 | 0x0004 | 0x0010 = 0x0015
            DllCall("SetWindowPos", "Ptr", hwnd, "Ptr", 0, "Int", guiX, "Int", guiY, "Int", 0, "Int", 0,
                "UInt", 0x0015)
        }
        flagGui.Show("NA")
    } catch {
    }
}

; Top-center of the active window on its monitor; top-center of the work area on other monitors.
DictationFlag_RepositionAll() {
    global g_DictationFlagGuis, g_DictationFlagSlot, g_DictationFlagFollowCache
    if (!g_DictationFlagGuis.Length)
        return

    hasWin := GetDictationActiveWindowRect(&wl, &wt, &wr, &wb)
    activeMon := GetDictationActiveMonitor()
    monitorCount := MonitorGetCount()
    winKey := hasWin ? (wl . "," . wt . "," . wr . "," . wb) : "0"
    key := winKey . "|" . activeMon . "|" . g_DictationFlagSlot . "|" . monitorCount
    if (key = g_DictationFlagFollowCache)
        return
    g_DictationFlagFollowCache := key

    for item in g_DictationFlagGuis {
        flagGui := item.gui
        monitorIdx := item.monitor
        if !IsObject(flagGui)
            continue

        try {
            MonitorGetWorkArea(monitorIdx, &ml, &mt, &mr, &mb)
            flagGui.GetPos(, , &gw, &gh)
        } catch {
            continue
        }

        if (monitorIdx = activeMon && hasWin) {
            guiX := wl + ((wr - wl) - gw) // 2
            guiY := wt + 8
            if (guiX < wl + 2)
                guiX := wl + 2
            if (guiY < wt + 2)
                guiY := wt + 2
            if (guiX + gw > wr - 2)
                guiX := wr - gw - 2
            if (guiY + gh > wb - 2)
                guiY := wb - gh - 2
        } else {
            guiX := ml + ((mr - ml) - gw) // 2
            guiY := mt + 12
        }
        if (guiX < ml)
            guiX := ml
        if (guiY < mt)
            guiY := mt
        if (guiX + gw > mr)
            guiX := mr - gw
        if (guiY + gh > mb)
            guiY := mb - gh

        DictationFlag_MoveGui(flagGui, guiX, guiY)
    }
}

; Show or refresh recording flags on every monitor from the persisted Handy slot.
DictationFlag_ShowForRecording() {
    global g_DictationFlagGuis, g_DictationFlagSlot

    slot := 0
    try slot := Handy_ReadPersistedAiModelSlotFromIni()
    if (slot < 1 || slot > 3)
        slot := 0

    monitorCount := MonitorGetCount()
    if (monitorCount < 1) {
        DictationFlag_Hide()
        return
    }

    needRebuild := (slot != g_DictationFlagSlot) || (g_DictationFlagGuis.Length != monitorCount)
    if (!needRebuild) {
        for item in g_DictationFlagGuis {
            try {
                if !IsObject(item.gui) || !item.gui.Hwnd {
                    needRebuild := true
                    break
                }
            } catch {
                needRebuild := true
                break
            }
        }
    }

    if (needRebuild) {
        DictationFlag_Hide()
        g_DictationFlagSlot := slot
        imagePath := (slot >= 1) ? LanguageFlag_GetImagePath(slot) : ""
        loop monitorCount {
            flagGui := DictationFlag_CreateGui(slot, imagePath)
            g_DictationFlagGuis.Push({ monitor: A_Index, gui: flagGui })
        }
    }

    DictationFlag_RepositionAll()
}

DictationIndicator_SyncPosition() {
    DictationFlag_RepositionAll()
}

; Show or refresh the opaque language flag on every monitor.
ShowDictationIndicator() {
    DictationFlag_ShowForRecording()
}

; Kept for callers (e.g. "Pasting..."); the flag has no status text overlay.
UpdateDictationIndicatorText(message := "") {
}

; Hide and destroy the dictation indicator
HideDictationIndicator() {
    DictationFlag_Hide()
}

; Follow the active window so the flag stays top-centered.
UpdateDictationIndicatorPulse() {
    global g_DictationFlagGuis
    if (!g_DictationFlagGuis.Length)
        return
    DictationIndicator_SyncPosition()
}

; Start the follow timer
StartDictationPulseTimer() {
    global g_DictationPulseTimer, DICTATION_PULSE_INTERVAL

    StopDictationPulseTimer()

    g_DictationPulseTimer := UpdateDictationIndicatorPulse
    SetTimer(g_DictationPulseTimer, DICTATION_PULSE_INTERVAL)
}

; Stop the pulse animation timer
StopDictationPulseTimer() {
    global g_DictationPulseTimer

    if (g_DictationPulseTimer) {
        try {
            SetTimer(g_DictationPulseTimer, 0)
        } catch {
            ; Ignore errors
        }
        g_DictationPulseTimer := false
    }
}

; Audio firewall: Throttle dictation sounds to prevent duplicates
; Enforces a minimum 1000ms gap between sounds regardless of how many times logic fires
SafePlayDictationSound(filePath) {
    Critical  ; Prevents thread interruption - ensures atomic check-and-update sequence
    global g_LastDictationSoundTick, g_DictationStartSound
    static lastStartSoundTick := 0

    ; Special handling for start sound: 7 second cooldown to prevent duplicates
    if (InStr(filePath, "speach-start.wav")) {
        if (A_TickCount - lastStartSoundTick < 7000) {
            return
        }
        lastStartSoundTick := A_TickCount
    } else {
        ; Standard 1 second cooldown for other sounds
        if (A_TickCount - g_LastDictationSoundTick < 1000) {
            return
        }
    }

    ; Update timestamp and play sound (if enabled)
    g_LastDictationSoundTick := A_TickCount
    if (FileExist(filePath)) {
        try {
            ScriptSoundPlay(filePath)
        } catch {
            ; Silently ignore playback failures (missing file, sync placeholder, format, etc.)
        }
    }
}

; Handler for clipboard changes during dictation completion
DictationClipboardHandler(DataType) {
    ; Remove handler immediately to prevent multiple triggers
    OnClipboardChange(DictationClipboardHandler, 0)

    ; Trigger completion logic immediately
    PlayDictationCompletionChime()
}

; Play completion chime after transcription finishes
PlayDictationCompletionChime(*) {
    global g_DictationCompletionChimeScheduled, g_PendingDictationAction,
        g_KeepIndicatorVisible, g_PendingGeminiPromptAfterDictation, g_D2C_DictationSubmitMenuCycleFinished

    ; Ensure clipboard handler is removed (safe to call even if already removed)
    try {
        OnClipboardChange(DictationClipboardHandler, 0)
    }

    ; Cancel fallback timer to prevent redundant calls
    SetTimer(PlayDictationCompletionChime, 0)

    ; CRITICAL: Test-and-set pattern - clear flag IMMEDIATELY to prevent duplicates
    ; Use Critical to ensure atomicity
    Critical "On"
    chimeShouldPlay := g_DictationCompletionChimeScheduled
    g_DictationCompletionChimeScheduled := false  ; Clear IMMEDIATELY to prevent other calls
    Critical "Off"

    ; Only play if flag was set (prevent duplicate execution)
    if (chimeShouldPlay) {
        g_D2C_DictationSubmitMenuCycleFinished := false
        SafePlayDictationSound(g_DictationStopSound)

        ; Execute pending action if one was set (reserved for future use).
        pendingAction := g_PendingDictationAction
        g_PendingDictationAction := ""  ; Clear immediately after reading

        if (pendingAction = "Paste") {
            ; Update indicator text to show status
            UpdateDictationIndicatorText("Pasting...")
            ; Execute paste command
            Send "^v"
            ; Wait for paste to complete before hiding indicator
            Sleep 100  ; Small delay to ensure paste completes
            ; Hide indicator only after paste completes
            HideDictationIndicator()
            g_KeepIndicatorVisible := false
        }

        ; If user stopped dictation with Win+Alt+Shift+0 (no pending action), show Gemini confirm banner (once only).
        Critical "On"
        pendingGemini := g_PendingGeminiPromptAfterDictation
        g_PendingGeminiPromptAfterDictation := false  ; Claim atomically so only one invocation shows the banner
        Critical "Off"
        if (pendingGemini && pendingAction = "") {
            D2C_FlowManager.GetInstance().StartFromDictation()
        }
    }
}

; Called when dictation stop detected: play chime now if clipboard already changed, else wait for change
DictationCompletionChimeOrWaitForClipboard() {
    global g_DictationStartClipboardText
    currentClip := ""
    try {
        currentClip := A_Clipboard
    }
    if (currentClip != g_DictationStartClipboardText) {
        PlayDictationCompletionChime()
    } else {
        OnClipboardChange(DictationClipboardHandler)
        SetTimer(PlayDictationCompletionChime, -1500)
    }
}

CheckDictationRecordingWindow() {
    global g_DictationActive, g_LastStateTransitionTick, g_DictationStartClipboardText
    global g_DictationSoundPlayed, g_DictationCompletionChimeScheduled, g_DictationPulseTimer, g_KeepIndicatorVisible
    ; Check if the "Recording" window exists
    windowExists := false
    try {
        windowExists := WinExist("Recording ahk_exe handy.exe")
    } catch {
        windowExists := false
    }

    ; Handle Start: window exists
    if (windowExists) {
        if (!g_DictationActive) {
            g_DictationActive := true
            g_LastStateTransitionTick := A_TickCount

            ; Capture current clipboard content to detect changes later
            try {
                g_DictationStartClipboardText := A_Clipboard
            } catch {
                g_DictationStartClipboardText := ""
            }

            try {
                RunSetMicVolumeScript()
            } catch Error as e {
                ; Silently handle errors - don't interrupt dictation if script fails
            }

            ShowDictationIndicator()
            StartDictationPulseTimer()
        }

        ; Atomic test-and-set: one sound per session when window first detected
        Critical "On"
        if (!g_DictationSoundPlayed) {
            g_DictationSoundPlayed := true
            Critical "Off"
            SafePlayDictationSound(g_DictationStartSound)
        } else {
            Critical "Off"
        }
    }
    ; Handle Stop: window gone and was active
    else if (!windowExists && g_DictationActive) {
        Critical "On"
        if (!g_DictationActive || g_DictationCompletionChimeScheduled) {
            Critical "Off"
            return
        }

        if (g_LastStateTransitionTick && (A_TickCount - g_LastStateTransitionTick < 500)) {
            Critical "Off"
            return
        }

        g_DictationCompletionChimeScheduled := true
        g_LastStateTransitionTick := A_TickCount
        g_DictationActive := false
        Critical "Off"
        g_DictationSoundPlayed := false

        StopDictationPulseTimer()
        HideDictationIndicator()
        DictationCompletionChimeOrWaitForClipboard()
    } else if (g_DictationActive && windowExists) {
        ShowDictationIndicator()
        if (!g_DictationPulseTimer) {
            StartDictationPulseTimer()
        }
    }
}

; Start timer to periodically check Recording window state
StartDictationCheckTimer() {
    global g_DictationCheckTimer

    ; Stop any existing timer
    StopDictationCheckTimer()

    ; Check every 500ms
    g_DictationCheckTimer := CheckDictationRecordingWindow
    SetTimer(g_DictationCheckTimer, 500)
}

; Stop the check timer
StopDictationCheckTimer() {
    global g_DictationCheckTimer

    if (g_DictationCheckTimer) {
        try {
            SetTimer(g_DictationCheckTimer, 0)
        } catch {
            ; Ignore errors
        }
        g_DictationCheckTimer := false
    }
}

; Toggle dictation mode on/off
; The check timer handles everything automatically, this just triggers an immediate check
ToggleDictationMode() {
    ; Trigger immediate check (the timer will handle showing/hiding)
    ; This provides instant detection if window already exists
    CheckDictationRecordingWindow()

    ; OPTIMIZED: Ultra-fast polling for instant window detection and audio feedback
    ; Start with 25ms polling (4x faster than normal) for ultra-responsive detection
    ; This ensures zero-delay audio feedback when handy.exe launches
    SetTimer(CheckDictationRecordingWindow, 25)
    ; Revert to normal 500ms polling after 3 seconds (window should be detected by then)
    SetTimer(RevertDictationPolling, -3000)
}

RevertDictationPolling() {
    SetTimer(CheckDictationRecordingWindow, 500)
}

; Force end dictation immediately (e.g., when Ask action is triggered)
; This immediately removes Esc restriction and hides the indicator
EndDictation() {
    global g_DictationActive, g_DictationSoundPlayed

    g_DictationActive := false
    g_DictationSoundPlayed := false

    StopDictationPulseTimer()
    HideDictationIndicator()
}

; Cleanup dictation indicator resources
CleanupDictationIndicator(*) {
    StopDictationPulseTimer()
    StopDictationCheckTimer()
    HideDictationIndicator()
}

; Register cleanup on script exit
OnExit(CleanupDictationIndicator)

; Toggle dictation mode with Win+Alt+Shift+0
; ~ prefix: key passes through to handy.exe. First press starts dictation, second stops and copies.
; Uses KeyWait + state machine + recursion guard to prevent duplicate triggers (typematic repeats).
~#!+0::
{
    global g_DictationActive, g_LastStateTransitionTick, g_DictationStartSound
    global g_ProgrammaticDictationStop, g_PendingGeminiPromptAfterDictation, g_D2C_DictationSubmitMenuCycleFinished
    global g_DictationHotkeyIsOwner
    static lastHotkeyTick := 0
    static isProcessing := false

    ; Defensive: if Utils is included after a script-level auto-execute return, globals may be uninitialized.
    ; Default to "not owner" to avoid double-handling dictation across processes.
    if (!IsSet(g_DictationHotkeyIsOwner))
        g_DictationHotkeyIsOwner := false

    if (!g_DictationHotkeyIsOwner) {
        return
    }

    ; Skip when script sends #!+0 programmatically
    if (g_ProgrammaticDictationStop) {
        g_ProgrammaticDictationStop := false
        return
    }

    if (isProcessing)
        return

    currentTick := A_TickCount
    if (currentTick - lastHotkeyTick < 200)
        return
    lastHotkeyTick := currentTick
    isProcessing := true
    ; Capture before KeyWait: check timer may clear g_DictationActive when Recording window closes,
    ; so by the time we reach if/else it can be false even when user intended to stop.
    dictationWasActiveOnKeyPress := g_DictationActive

    keyWaitStart := A_TickCount
    KeyWait("0", "L")

    if (!g_DictationActive) {
        g_DictationActive := true
        g_LastStateTransitionTick := A_TickCount
        ShowDictationIndicator()
        StartDictationPulseTimer()
        ; Sound: monitoring loop plays when window detected (zero latency)

        try {
            RunSetMicVolumeScript()
        } catch {
        }
    }

    ; User was stopping dictation (had been active when they pressed key) -> show Gemini confirm after completion
    if (dictationWasActiveOnKeyPress) {
        g_PendingGeminiPromptAfterDictation := true
        g_D2C_DictationSubmitMenuCycleFinished := false
        g_DictationGeminiConfirmBannerVisible := false  ; Allow 5s banner to show for this cycle (reset from previous N cancel)
    } else {
    }

    ToggleDictationMode()
    isProcessing := false
}
