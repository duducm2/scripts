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
; Multi-lang (Handy slot 3) start cue. CC0: Robin Lamb, "UI Sound Effects",
; https://opengameart.org/content/ui-sound-effects-button-clicks-user-feedback-notifications
; Pack file chimes.wav, raised 4.5 dB. Other pack chimes are much louder.
global g_DictationStartSoundMultilang := A_ScriptDir . "\assets\sounds\dictation-start-multilang.wav"
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
; Aborted EN <-> multi-lang take: monitor may hide the flag, but must not chime, hook clipboard, or banner.
global g_DictationSuppressCompletion := false
; After a language-switch restart, ignore "overlay gone" until the new Recording window has been seen.
global g_DictationAwaitRecordingWindow := false
; A Recording overlay counts as the restarted take only after the previous one has been absent.
global g_DictationRestartSawOverlayAbsent := false
; One-shot: the language-swap start chime ignores the shared 7s start-cue cooldown.
global g_DictationBypassStartSoundCooldown := false
; True while a K/L swap has muted the capture endpoint and not yet restored it.
global g_DictationSwapMicMuted := false
; True only during Utility Shortcuts K/L. Await and suppress must not affect a later take.
global g_DictationLanguageSwapActive := false
; Text Handy copied when the last user stop finished. V pastes this, not an older clip.
global g_DictationCompletedText := ""

; AppLaunchers.ahk is the single long-lived owner for ~#!+0, Recording flag, and
; Send dictation? banner — same script-name pin as HandyAi_IsOwnerProcess().
; Act / Shift keys / Gemini / WM / Teams / Outlook include Utils but must not own.
Dictation_IsOwnerProcess() {
    return A_ScriptName = "AppLaunchers.ahk"
}

; Primary: script-name pin. Secondary: named mutex guards against a second AppLaunchers.
InitializeDictationHotkeyOwnership() {
    global g_DictationHotkeyOwnerHandle, g_DictationHotkeyIsOwner
    g_DictationHotkeyIsOwner := false
    g_DictationHotkeyOwnerHandle := 0
    if (!Dictation_IsOwnerProcess())
        return

    mutexName := "Local\D2C_Dictation_Hotkey_Owner"
    hMutex := DllCall("CreateMutex", "Ptr", 0, "Int", 0, "Str", mutexName, "Ptr")
    if (!hMutex)
        return
    err := DllCall("GetLastError", "UInt")
    if (err = 183) { ; ERROR_ALREADY_EXISTS — another AppLaunchers already owns
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
global DICTATION_SQUARE_SIZE := 50  ; Flag height in px (150 minus 30%; aspect preserved)
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
; This is the only language flag. It is destroyed when the take ends.
DictationFlag_ShowForRecording() {
    global g_DictationFlagGuis, g_DictationFlagSlot, g_HandyAiPersistedSlot

    slot := 0
    try slot := Handy_ReadPersistedAiModelSlotFromIni()
    if (slot < 1 || slot > 3)
        slot := 0
    else
        g_HandyAiPersistedSlot := slot

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

; English / Portuguese use speach-start.wav. Multi-lang (slot 3) uses a different chime.
Dictation_StartSoundForCurrentModel() {
    global g_DictationStartSound, g_DictationStartSoundMultilang, HANDY_AI_SLOT_MULTILANG
    slot := 0
    try slot := Handy_GetPersistedAiModelSlot()
    catch
        slot := 0
    if (slot = HANDY_AI_SLOT_MULTILANG)
        return g_DictationStartSoundMultilang
    return g_DictationStartSound
}

; Audio firewall: Throttle dictation sounds to prevent duplicates
; Enforces a minimum 1000ms gap between sounds regardless of how many times logic fires
SafePlayDictationSound(filePath) {
    Critical  ; Prevents thread interruption - ensures atomic check-and-update sequence
    global g_LastDictationSoundTick, g_DictationStartSound, g_DictationStartSoundMultilang
    global g_DictationBypassStartSoundCooldown
    static lastStartSoundTick := 0

    ; Start cues: 7 second cooldown to prevent duplicates (English and multi-lang).
    ; A language-swap restart sets the bypass so the new language chime still plays.
    if (filePath = g_DictationStartSound || filePath = g_DictationStartSoundMultilang) {
        bypass := g_DictationBypassStartSoundCooldown
        if (bypass)
            g_DictationBypassStartSoundCooldown := false
        if (!bypass && A_TickCount - lastStartSoundTick < 7000) {
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

; Remember Handy’s new clipboard text so V can paste it after the stop.
Dictation_NoteCompletedClipboard() {
    global g_DictationCompletedText, g_DictationStartClipboardText
    clip := ""
    try clip := A_Clipboard
    catch
        clip := ""
    if (clip != "" && clip != g_DictationStartClipboardText)
        g_DictationCompletedText := clip
    return g_DictationCompletedText
}

; Wait until Handy’s stop has copied new text. Do not return the pre-take clipboard.
Dictation_WaitForCompletedText(timeoutMs := 1500) {
    global g_DictationCompletedText
    start := A_TickCount
    while (A_TickCount - start < timeoutMs) {
        noted := Dictation_NoteCompletedClipboard()
        if (noted != "")
            return noted
        Sleep 50
    }
    return g_DictationCompletedText
}

; Handler for clipboard changes during dictation completion
DictationClipboardHandler(DataType) {
    global g_DictationSuppressCompletion, g_PendingGeminiPromptAfterDictation
    ; Remove handler immediately to prevent multiple triggers
    OnClipboardChange(DictationClipboardHandler, 0)

    ; Aborted wrong-language take: its transcription must not complete the next session.
    if (g_DictationSuppressCompletion && !g_PendingGeminiPromptAfterDictation)
        return

    Dictation_NoteCompletedClipboard()
    ; Trigger completion logic immediately
    PlayDictationCompletionChime()
}

; Drop a queued stop chime and clipboard hook (aborted take, or the moment the new overlay appears).
Dictation_DisarmCompletionHooks() {
    global g_DictationCompletionChimeScheduled
    g_DictationCompletionChimeScheduled := false
    SetTimer(PlayDictationCompletionChime, 0)
    try OnClipboardChange(DictationClipboardHandler, 0)
    catch {
    }
}

; Play completion chime after transcription finishes
PlayDictationCompletionChime(*) {
    global g_DictationCompletionChimeScheduled, g_PendingDictationAction,
        g_KeepIndicatorVisible, g_PendingGeminiPromptAfterDictation, g_D2C_DictationSubmitMenuCycleFinished
    global g_DictationSuppressCompletion

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
        ; Aborted take: consume the flag so it cannot steal the next stop's chime, and do not banner.
        if (g_DictationSuppressCompletion && !g_PendingGeminiPromptAfterDictation)
            return
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
        Dictation_NoteCompletedClipboard()
        PlayDictationCompletionChime()
    } else {
        OnClipboardChange(DictationClipboardHandler)
        SetTimer(PlayDictationCompletionChime, -1500)
    }
}

; True when Handy Recording overlay is present (EN + common PT titles).
Dictation_RecordingWindowExists() {
    static titles := [
        "Recording ahk_exe handy.exe",
        "Recording Overlay ahk_exe handy.exe",
        "Gravação ahk_exe handy.exe",
        "Gravacao ahk_exe handy.exe",
        "Gravando ahk_exe handy.exe"
    ]
    for t in titles {
        try {
            if WinExist(t)
                return true
        } catch {
        }
    }
    return false
}

; English (slot 1) <-> multi-lang (slot 3) while Handy is recording: stop the session.
; The synthetic chord is ignored by ~#!+0 so "Send dictation?" does not open.
; Only the dictation owner may inject the chord; other hosts would look like a user stop.
; Returns true when a recording was stopped.
Handy_StopDictationIfEnglishMultilangSwitch(targetSlot) {
    global g_ProgrammaticDictationStop, g_DictationSuppressCompletion, HANDY_AI_SLOT_ENGLISH, HANDY_AI_SLOT_MULTILANG
    if (!Dictation_IsOwnerProcess())
        return false
    current := Handy_GetPersistedAiModelSlot()
    if (current = targetSlot)
        return false
    crossing := (current = HANDY_AI_SLOT_ENGLISH && targetSlot = HANDY_AI_SLOT_MULTILANG)
    || (current = HANDY_AI_SLOT_MULTILANG && targetSlot = HANDY_AI_SLOT_ENGLISH)
    if (!crossing || !Dictation_RecordingWindowExists())
        return false
    ; Aborted take must not arm a chime, clipboard hook, or Send dictation? banner.
    g_DictationSuppressCompletion := true
    Dictation_DisarmCompletionHooks()
    g_ProgrammaticDictationStop := true
    Send "#!+0"
    loop 30 {
        if (!Dictation_RecordingWindowExists())
            break
        Sleep 50
    }
    return true
}

; Release stuck modifiers, then send #!+0. forStop arms the stop flag; otherwise the start flag.
; The flag stays set until ~#!+0 consumes it.
Dictation_SendProgrammaticChord(forStop) {
    global g_ProgrammaticDictationStop, g_ProgrammaticDictationStart
    try SetTimer(Dictation_ExpireProgrammaticStart, 0)
    catch {
    }
    try UtilityShortcuts_ReleaseStuckModifiers()
    catch {
    }
    if (forStop) {
        g_ProgrammaticDictationStart := false
        g_ProgrammaticDictationStop := true
    } else {
        g_ProgrammaticDictationStop := false
        g_ProgrammaticDictationStart := true
    }
    Send "#!+0"
}

; True once ~#!+0 has cleared the programmatic stop or start flag.
Dictation_WaitProgrammaticChordConsumed(forStop, timeoutMs) {
    global g_ProgrammaticDictationStop, g_ProgrammaticDictationStart
    start := A_TickCount
    while (A_TickCount - start < timeoutMs) {
        if (forStop) {
            if (!g_ProgrammaticDictationStop)
                return true
        } else if (!g_ProgrammaticDictationStart) {
            return true
        }
        Sleep 50
    }
    return forStop ? !g_ProgrammaticDictationStop : !g_ProgrammaticDictationStart
}

; K/L finished or failed. Drop swap-only state so the next user stop can open Send dictation.
Dictation_EndLanguageSwap() {
    global g_DictationLanguageSwapActive, g_DictationSuppressCompletion
    global g_DictationAwaitRecordingWindow, g_DictationRestartSawOverlayAbsent
    g_DictationLanguageSwapActive := false
    g_DictationSuppressCompletion := false
    g_DictationAwaitRecordingWindow := false
    g_DictationRestartSawOverlayAbsent := false
}

; Hide the recording flag for an aborted take. Does not chime or open Send dictation?.
Dictation_FinishAbortedTake() {
    global g_DictationActive, g_DictationSoundPlayed, g_LastStateTransitionTick
    g_DictationActive := false
    g_DictationSoundPlayed := false
    g_LastStateTransitionTick := A_TickCount
    try StopDictationPulseTimer()
    catch {
    }
    try HideDictationIndicator()
    catch {
    }
}

; K/L language swap while a take is up. Stops whatever is recording, including a same-slot
; or Portuguese take. Returns "none" (nothing to stop), "stopped", or "failed".
; On failure the capture endpoint is unmuted and no further chord is sent.
Handy_StopDictationForLanguageSwap() {
    global g_DictationSuppressCompletion, g_ProgrammaticDictationStop, g_DictationRestartSawOverlayAbsent
    if (!Dictation_IsOwnerProcess())
        return "failed"
    if (!Dictation_RecordingWindowExists())
        return "none"

    Dictation_MuteCaptureForSwap()
    g_DictationSuppressCompletion := true
    g_DictationRestartSawOverlayAbsent := false
    Dictation_DisarmCompletionHooks()

    loop 3 {
        Dictation_SendProgrammaticChord(true)
        start := A_TickCount
        while (A_TickCount - start < 2000) {
            overlayGone := !Dictation_RecordingWindowExists()
            if (overlayGone && !g_ProgrammaticDictationStop) {
                g_DictationRestartSawOverlayAbsent := true
                Dictation_FinishAbortedTake()
                return "stopped"
            }
            Sleep 50
        }
        ; Overlay already gone, but the hotkey has not consumed the flag yet.
        if (!Dictation_RecordingWindowExists()) {
            g_ProgrammaticDictationStop := false
            g_DictationRestartSawOverlayAbsent := true
            Dictation_FinishAbortedTake()
            return "stopped"
        }
    }

    g_ProgrammaticDictationStop := false
    Dictation_UnmuteCaptureAfterSwap()
    AiModelBanner_Show("❌ Could not stop dictation", "E74C3C")
    Sleep 2000
    AiModelBanner_Hide()
    return "failed"
}

; After a mid-dictation language swap: start a normal take again.
; The start flag stays set until ~#!+0 consumes it. Returns true once the new overlay is up.
Handy_RestartDictationAfterLanguageSwitch() {
    global g_ProgrammaticDictationStart, g_ProgrammaticDictationStop, g_DictationActive, g_DictationSoundPlayed
    global g_PendingGeminiPromptAfterDictation
    global g_DictationSuppressCompletion, g_DictationAwaitRecordingWindow, g_DictationRestartSawOverlayAbsent
    global g_DictationBypassStartSoundCooldown
    if (!Dictation_IsOwnerProcess())
        return false
    ; A leftover stop flag would swallow this chord and the user's later stop.
    g_ProgrammaticDictationStop := false
    g_ProgrammaticDictationStart := false
    try SetTimer(Dictation_ExpireProgrammaticStart, 0)
    catch {
    }
    g_DictationActive := false
    g_DictationSoundPlayed := false
    g_PendingGeminiPromptAfterDictation := false
    g_DictationSuppressCompletion := true
    g_DictationAwaitRecordingWindow := true
    g_DictationRestartSawOverlayAbsent := !Dictation_RecordingWindowExists()
    g_DictationBypassStartSoundCooldown := true
    Dictation_DisarmCompletionHooks()

    loop 3 {
        ; A previous chord may have opened the overlay just after the last poll.
        if (A_Index > 1 && g_DictationRestartSawOverlayAbsent && Dictation_RecordingWindowExists()) {
            if (Dictation_WaitRestartedOverlay(200))
                return true
        }
        Dictation_SendProgrammaticChord(false)
        Dictation_WaitProgrammaticChordConsumed(false, 1500)
        if (Dictation_WaitRestartedOverlay(2000))
            return true
        ; This attempt did not land. Drop a stuck start flag before the next chord.
        g_ProgrammaticDictationStart := false
    }

    g_ProgrammaticDictationStart := false
    g_DictationAwaitRecordingWindow := false
    g_DictationRestartSawOverlayAbsent := false
    g_DictationBypassStartSoundCooldown := false
    g_DictationSuppressCompletion := false
    Dictation_UnmuteCaptureAfterSwap()
    return false
}

; Poll until a Recording overlay appears after the previous one was absent.
Dictation_WaitRestartedOverlay(timeoutMs) {
    global g_DictationRestartSawOverlayAbsent, g_DictationSoundPlayed, g_DictationBypassStartSoundCooldown
    global g_DictationAwaitRecordingWindow, g_DictationSuppressCompletion, g_DictationActive
    start := A_TickCount
    while (A_TickCount - start < timeoutMs) {
        exists := Dictation_RecordingWindowExists()
        if (!exists) {
            g_DictationRestartSawOverlayAbsent := true
            Sleep 50
            continue
        }
        ; Monitor already accepted this overlay, or it appeared only after the previous one was gone.
        if (g_DictationRestartSawOverlayAbsent || !g_DictationAwaitRecordingWindow) {
            CheckDictationRecordingWindow()
            Dictation_UnmuteCaptureAfterSwap()
            if (!g_DictationSoundPlayed) {
                g_DictationSoundPlayed := true
                g_DictationActive := true
                g_DictationBypassStartSoundCooldown := true
                g_DictationAwaitRecordingWindow := false
                g_DictationSuppressCompletion := false
                SafePlayDictationSound(Dictation_StartSoundForCurrentModel())
                try ShowDictationIndicator()
                catch {
                }
                try StartDictationPulseTimer()
                catch {
                }
            }
            return true
        }
        Sleep 50
    }
    return false
}

Dictation_ExpireProgrammaticStart(*) {
    global g_ProgrammaticDictationStart
    g_ProgrammaticDictationStart := false
}

; First sight of the Recording overlay after the previous overlay has been absent.
Dictation_AcceptRestartedRecordingWindow() {
    global g_DictationAwaitRecordingWindow, g_DictationSuppressCompletion, g_DictationStartClipboardText
    global g_DictationRestartSawOverlayAbsent, g_DictationCompletedText
    g_DictationAwaitRecordingWindow := false
    g_DictationRestartSawOverlayAbsent := false
    g_DictationSuppressCompletion := false
    g_DictationCompletedText := ""
    Dictation_DisarmCompletionHooks()
    Dictation_UnmuteCaptureAfterSwap()
    try g_DictationStartClipboardText := A_Clipboard
    catch {
        g_DictationStartClipboardText := ""
    }
}

CheckDictationRecordingWindow() {
    global g_DictationActive, g_LastStateTransitionTick, g_DictationStartClipboardText
    global g_DictationSoundPlayed, g_DictationCompletionChimeScheduled, g_DictationPulseTimer, g_KeepIndicatorVisible
    global g_DictationHotkeyIsOwner, g_DictationAwaitRecordingWindow, g_DictationSuppressCompletion
    global g_DictationRestartSawOverlayAbsent, g_DictationLanguageSwapActive
    ; Non-owners must never drive Recording flag / chime / banner (Act used to steal this).
    if (!IsSet(g_DictationHotkeyIsOwner) || !g_DictationHotkeyIsOwner)
        return

    windowExists := Dictation_RecordingWindowExists()
    ; Await and suppress apply only while a K/L swap is running.
    swapActive := g_DictationLanguageSwapActive

    ; The pre-stop overlay must disappear before any window counts as the restarted take.
    if (swapActive && g_DictationAwaitRecordingWindow && !windowExists) {
        g_DictationRestartSawOverlayAbsent := true
        if (g_DictationSuppressCompletion) {
            Dictation_FinishAbortedTake()
            return
        }
        return
    }
    if (swapActive && g_DictationAwaitRecordingWindow && windowExists && !g_DictationRestartSawOverlayAbsent)
        return

    ; New take's overlay: drop the aborted take's completion so it cannot swallow this stop.
    if (swapActive && windowExists && g_DictationAwaitRecordingWindow && g_DictationRestartSawOverlayAbsent)
        Dictation_AcceptRestartedRecordingWindow()

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
            SafePlayDictationSound(Dictation_StartSoundForCurrentModel())
        } else {
            Critical "Off"
        }
    }
    ; Handle Stop: window gone and was active
    else if (!windowExists && g_DictationActive) {
        ; Gap before the restarted overlay appears. Not a user stop.
        if (swapActive && g_DictationAwaitRecordingWindow)
            return

        ; Aborted wrong-language take closed. Hide the flag; do not chime or banner.
        if (swapActive && g_DictationSuppressCompletion) {
            g_DictationActive := false
            g_DictationSoundPlayed := false
            g_LastStateTransitionTick := A_TickCount
            StopDictationPulseTimer()
            try HideDictationIndicator()
            catch {
            }
            return
        }

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
    global g_DictationHotkeyIsOwner
    if (!IsSet(g_DictationHotkeyIsOwner) || !g_DictationHotkeyIsOwner)
        return
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
; Only AppLaunchers owns this path (Dictation_IsOwnerProcess / g_DictationHotkeyIsOwner).
~#!+0::
{
    global g_DictationActive, g_LastStateTransitionTick, g_DictationStartSound
    global g_ProgrammaticDictationStop, g_ProgrammaticDictationStart, g_PendingGeminiPromptAfterDictation,
        g_D2C_DictationSubmitMenuCycleFinished
    global g_DictationHotkeyIsOwner, g_DictationCompletionChimeScheduled
    global g_DictationSuppressCompletion, g_DictationAwaitRecordingWindow, g_DictationRestartSawOverlayAbsent
    global g_DictationCompletedText
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

    ; Language switch while a take was running: this chord is a new start.
    forcedStart := false
    if (g_ProgrammaticDictationStart) {
        g_ProgrammaticDictationStart := false
        SetTimer(Dictation_ExpireProgrammaticStart, 0)
        g_DictationActive := false
        g_PendingGeminiPromptAfterDictation := false
        forcedStart := true
    }

    if (!forcedStart && isProcessing)
        return

    currentTick := A_TickCount
    if (!forcedStart && currentTick - lastHotkeyTick < 200)
        return
    lastHotkeyTick := currentTick
    isProcessing := true
    try {
        ; A leftover swap must not swallow this user chord or the Send dictation banner.
        if (!forcedStart) {
            g_DictationSuppressCompletion := false
            g_DictationAwaitRecordingWindow := false
            g_DictationRestartSawOverlayAbsent := false
        }
        ; Overlay can close during KeyWait. Remember it now so a cleared recording flag still counts as a stop.
        overlayWasUp := false
        if (!forcedStart)
            overlayWasUp := Dictation_RecordingWindowExists()

        ; Capture before KeyWait: check timer may clear g_DictationActive when Recording window closes,
        ; so by the time we reach if/else it can be false even when user intended to stop.
        dictationWasActiveOnKeyPress := g_DictationActive

        ; Synthetic restart: Send has not released 0 yet, so KeyWait would never return.
        if (!forcedStart)
            KeyWait("0", "L")

        if (dictationWasActiveOnKeyPress || overlayWasUp) {
            ; Explicit STOP. Never fall into the start branch if Recording closed during KeyWait
            ; (that re-showed the recording flag and skipped/queued the command banner wrongly).
            ; A language-switch restart must not keep this stop suppressed.
            g_DictationSuppressCompletion := false
            g_DictationAwaitRecordingWindow := false
            g_PendingGeminiPromptAfterDictation := true
            g_D2C_DictationSubmitMenuCycleFinished := false
            g_DictationGeminiConfirmBannerVisible := false

            if (!g_DictationActive) {
                ; Monitor already ended the session during KeyWait — chime may have run before
                ; pendingGemini was set. Force banner path once.
                try HideDictationIndicator()
                catch {
                }
                try StopDictationPulseTimer()
                catch {
                }
                if (!g_DictationCompletionChimeScheduled)
                    SetTimer(Dictation_ForceSubmitMenuAfterStop, -50)
            } else {
                ToggleDictationMode()
            }
        } else if (!g_DictationActive) {
            ; START
            g_DictationActive := true
            g_DictationCompletedText := ""
            g_LastStateTransitionTick := A_TickCount
            ShowDictationIndicator()
            StartDictationPulseTimer()
            try {
                RunSetMicVolumeScript()
            } catch {
            }
            ToggleDictationMode()
        } else {
            ToggleDictationMode()
        }
    } finally {
        isProcessing := false
    }
}

; After an explicit stop where the monitor already cleared g_DictationActive during KeyWait.
Dictation_ForceSubmitMenuAfterStop(*) {
    global g_PendingGeminiPromptAfterDictation, g_D2C_DictationSubmitMenuCycleFinished
    try HideDictationIndicator()
    catch {
    }
    try StopDictationPulseTimer()
    catch {
    }
    if (g_D2C_DictationSubmitMenuCycleFinished)
        return
    Critical "On"
    pending := g_PendingGeminiPromptAfterDictation
    g_PendingGeminiPromptAfterDictation := false
    Critical "Off"
    if (!pending)
        return
    try {
        D2C_FlowManager.GetInstance().StartFromDictation()
    } catch {
    }
}
