; =============================================================================
; AppLaunchers module: hotkey_clipangel.ahk
; #!+. Clip Angel paste and favorite flow; #!+7 tap/hold edit / paste-file
; Extracted verbatim from AppLaunchers.ahk; loaded via #include into the
; AppLaunchers.ahk process, which remains the entry point / source of truth.
; =============================================================================

; Hold threshold matches ZMK hold-tap tapping-term-ms = 200.
CLIPANGEL_WAS7_HOLD_MS := 200

; =============================================================================
; Send specific key combinations
; Hotkey: Win+Alt+Shift+.
; =============================================================================
#!+.::
{
    Sleep(100)
    Send("^c")
    Sleep(200)
    if hwnd := ClipAngel_MainHwnd()
        ClipAngel_ShowWindow(hwnd)
    Sleep(700)
    Send("!q")
    Sleep(200)
    ClipAngel_CloseAndRestoreFocus(0)
}

; =============================================================================
; Win+Alt+Shift+7 — tap: open Clip Angel + Edit (F4); hold 200ms+: Paste file then hide
; =============================================================================
#!+7::
{
    priorHwnd := 0
    try priorHwnd := WinGetID("A")
    catch {
    }

    pressTime := A_TickCount
    KeyWait "7", "T" . (CLIPANGEL_WAS7_HOLD_MS / 1000)
    isHold := (A_TickCount - pressTime) >= CLIPANGEL_WAS7_HOLD_MS
    if isHold
        try ScriptSoundPlay(A_ScriptDir . "\assets\sounds\copy.wav")

    ClipAngel_WaitChordModifiersReleased()
    ClipAngel_ReleaseChordModifiersForSend()

    if !ActivateClipAngelWithFocusCorrection(true) {
        ShowCenteredOverlay_Utils("❌ Clip Angel is not running.", 2000, BANNER_ACCENT_ERROR)
        return
    }

    hwnd := ClipAngel_MainHwnd()
    if !hwnd {
        ShowCenteredOverlay_Utils("❌ Clip Angel window not found.", 2000, BANNER_ACCENT_ERROR)
        return
    }

    if !isHold {
        SendInput "{F4}"
        return
    }

    try {
        if !ClipAngel_InvokePasteEnter(hwnd)
            ShowCenteredOverlay_Utils("❌ Clip Angel Paste file failed", 1500, BANNER_ACCENT_ERROR)
    } finally {
        ClipAngel_CloseAndRestoreFocus(priorHwnd)
    }
}

; =============================================================================
; Initialize Wikipedia scroll position auto-save timer - REMOVED
; =============================================================================
; Auto-save timer removed - now using manual save via Shift keys.ahk shortcut
