; =============================================================================
; Utils module: print_screen_escape.ahk
; Print Screen chime, global Escape hotkey
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; Print Screen with Chime
; Hotkey: Alt+PrintScreen
; Intercepts the hotkey to prevent other apps from handling it,
; manually triggers the screenshot, and plays a single chime
; =============================================================================
global g_LastPrintScreenSound := 0  ; Track last sound time for debouncing
global g_PrintScreenInProgress := false  ; Prevents recursion from Send

; Audio firewall for PrintScreen: Throttle sounds to prevent duplicates
; Uses Critical section to ensure atomic check-and-update (like dictation mode)
SafePlayPrintScreenSound() {
    Critical  ; Prevents thread interruption - ensures atomic check-and-update sequence
    global g_LastPrintScreenSound

    ; If less than 1000ms has passed since last sound, ignore this call
    if (A_TickCount - g_LastPrintScreenSound < 1000) {
        return
    }

    ; Update timestamp and play sound (if enabled)
    g_LastPrintScreenSound := A_TickCount
    ScriptSoundPlay(A_ScriptDir . "\assets\sounds\print-screen.wav")
}

; Set higher InputLevel to ensure our handler processes before others
#InputLevel 10
!PrintScreen::  ; Removed ~ prefix to CONSUME the hotkey (prevents other apps from receiving it)
{
    global g_PrintScreenInProgress

    ; Prevent recursion: if we're already processing, skip (this handles Send retriggering)
    if (g_PrintScreenInProgress) {
        return
    }

    ; Set flag to prevent recursion from the Send below
    g_PrintScreenInProgress := true

    ; Manually send Alt+PrintScreen to Windows to trigger the screenshot
    ; SendInput is more reliable and won't retrigger our hotkey due to SendLevel
    SendInput("!{PrintScreen}")

    ; Brief delay to ensure screenshot is captured, then play single chime
    ; Uses Critical section to prevent duplicate sounds from concurrent handlers
    Sleep 10
    SafePlayPrintScreenSound()

    ; Reset flag after a brief delay to allow normal operation
    Sleep 100
    g_PrintScreenInProgress := false
}
#InputLevel 0

; True when SendEscape() must not inject Escape (Handy dictation/recording UI). Physical Escape uses Escape:: below.
IsHandyDictationEscapeSuppressed() {
    global g_DictationActive
    return g_DictationActive || WinActive("Recording ahk_exe handy.exe") || WinActive(
        "Recording Overlay ahk_exe handy.exe") || WinExist("Recording ahk_exe handy.exe") || WinExist(
            "Recording Overlay ahk_exe handy.exe")
}

; Helper to send Escape while respecting dictation state (no-op when suppressed; see IsHandyDictationEscapeSuppressed).
SendEscape(count := 1) {
    if (IsHandyDictationEscapeSuppressed()) {
        return
    }
    if (count > 1)
        Send("{Escape " . count . "}")
    else
        Send("{Escape}")
}

; =============================================================================
; Global Escape hotkey (registered via Hotkey(), not Escape:: label)
; Modals use Hotkey("Escape", modalFn, "On") which replaces this binding; Hotkey("Escape", "Off") leaves Escape
; unhandled until Utils_EnsureGlobalEscapeHotkey() runs (see AiModel cleanup, square selector, hotstring cleanup, etc.).
; =============================================================================

; Optional escape callback: when set (e.g. by WindowManagement for project selector), Utils runs it.
; Callbacks return true when they handled Escape; false clears a stale callback and forwards to the app.
global g_OnEscapePressed := ""

; Cross-process WM selector IPC: write close request, wait for live selector to react, else purge stale sentinel.
Utils_TryCloseViaSentinel(sentinelPath, closeReqPath) {
    if (!FileExist(sentinelPath))
        return false
    try FileAppend("", closeReqPath)
    catch {
    }
    ; WM_Check*CloseRequest polls every 120ms — brief wait for a live selector to close.
    Sleep 160
    if (!FileExist(sentinelPath))
        return true
    try FileDelete(sentinelPath)
    catch {
    }
    try FileDelete(closeReqPath)
    catch {
    }
    return false
}

Utils_GlobalEscapeHandler(*) {
    global g_OnEscapePressed

    if (g_OnEscapePressed) {
        handled := false
        try {
            handled := !!g_OnEscapePressed.Call()
        } catch {
            handled := false
        }
        if (handled)
            return
        g_OnEscapePressed := ""
    }

    try {
        if (Utils_TryCloseViaSentinel(A_ScriptDir "\.cursor\wm_selector_open",
            A_ScriptDir "\.cursor\wm_selector_close_request"))
            return
        if (Utils_TryCloseViaSentinel(A_ScriptDir "\.cursor\wm_minimized_list_open",
            A_ScriptDir "\.cursor\wm_minimized_list_close_request"))
            return
        if (Utils_TryCloseViaSentinel(A_ScriptDir "\.cursor\wm_audio_bt_open",
            A_ScriptDir "\.cursor\wm_audio_bt_close_request"))
            return
    } catch {
    }

    ; I10 hotkey: forward at send level 0 so this handler is not re-triggered by SendInput (SendLevel / #InputLevel).
    SendLevel 0
    SendInput "{Escape}"
}

Utils_EnsureGlobalEscapeHotkey() {
    ; I10: distinct from modals that use Hotkey("Escape", fn, "On") at default level; avoids replace ambiguity.
    ; Forward path uses SendLevel 0 so synthetic Escape does not re-enter this handler.
    try {
        Hotkey("Escape", Utils_GlobalEscapeHandler, "I10 On")
    } catch {
    }
}

; Initial registration (replaces legacy Escape:: label)
Utils_EnsureGlobalEscapeHotkey()