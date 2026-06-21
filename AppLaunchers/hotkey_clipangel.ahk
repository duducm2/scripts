; =============================================================================
; AppLaunchers module: hotkey_clipangel.ahk
; #!+. Clip Angel paste and favorite flow
; Extracted verbatim from AppLaunchers.ahk; loaded via #include into the
; AppLaunchers.ahk process, which remains the entry point / source of truth.
; =============================================================================

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
    SendEscape()
}

; =============================================================================
; Initialize Wikipedia scroll position auto-save timer - REMOVED
; =============================================================================
; Auto-save timer removed - now using manual save via Shift keys.ahk shortcut
