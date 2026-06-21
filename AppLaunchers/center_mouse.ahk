; =============================================================================
; AppLaunchers module: center_mouse.ahk
; CenterMouse helper on active window
; Extracted verbatim from AppLaunchers.ahk; loaded via #include into the
; AppLaunchers.ahk process, which remains the entry point / source of truth.
; =============================================================================

; =============================================================================
; Helper function to center mouse on the active window
; =============================================================================
CenterMouse() {
    hwnd := WinExist("A")
    if hwnd
        AL_CenterMouseOnHwnd(hwnd)
}
