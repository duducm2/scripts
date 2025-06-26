#Requires AutoHotkey v2.0+
#SingleInstance Force

; -----------------------------------------------------------------------------
; This script consolidates all Window Management hotkeys.
; -----------------------------------------------------------------------------

; --- Hotkeys & Functions -----------------------------------------------------

; =============================================================================
; Minimize Active Window
; Hotkey: Win+Alt+Shift+6
; Original File: Minimize.ahk
; =============================================================================
#!+6::
{
    WinMinimize "A"
}

; =============================================================================
; Maximize Active Window
; Hotkey: Win+Alt+Shift+M
; Original File: Maximize window.ahk
; =============================================================================
#!+M::
{
    WinMaximize "A"
}

; =============================================================================
; Move Active Window to Specific Monitor and Maximize
; Hotkeys: Ctrl+Alt+Shift+A/S/D/F (MEH) -> monitors 1/2/3/4
; =============================================================================
^!+a:: MoveWinToMonitor(1)
^!+s:: MoveWinToMonitor(2)
^!+d:: MoveWinToMonitor(3)
^!+f:: MoveWinToMonitor(4)

MoveWinToMonitor(mon) {
    ; Validate monitor index
    if (mon > MonitorGetCount() || mon < 1) {
        MsgBox "Invalid monitor index: " mon
        return
    }

    hwnd := WinExist("A")
    if !hwnd {
        MsgBox "No active window."
        return
    }

    ; Obtain monitor work area
    MonitorGet mon, &left, &top, &right, &bottom

    ; Ensure window can be moved (restore if maximised/minimised)
    state := WinGetMinMax(hwnd) ; 1=min,2=max,0=normal
    if (state != 0) {
        WinRestore hwnd
        Sleep 100
    }

    width := right - left
    height := bottom - top

    ; First try the native WinMove (returns 1 on success, 0 on failure)
    ok := 0
    try ok := WinMove(hwnd, left, top, width, height)
    catch {
        ok := 0
    }

    ; Fallback to MoveWindow API if WinMove fails
    if !ok {
        DllCall("MoveWindow", "ptr", hwnd, "int", left, "int", top, "int", width, "int", height, "int", true)
    }

    ; Finally maximise so Windows treats it as maximised on that monitor
    WinMaximize hwnd
}

; =============================================================================
; Switch to Previous Window
; Hotkey: Ctrl+Alt+Shift+B (MEH+B)
; =============================================================================
^!+b:: AltTab(1)

; =============================================================================
; Switch to Second Previous Window
; Hotkey: Ctrl+Alt+Shift+C (MEH+C)
; =============================================================================
^!+c:: AltTab(2)

AltTab(count := 1) {
    ; Send the Alt+Tab combination the specified number of times.
    ; Using "!{Tab}" presses Alt+Tab in one shot, which Windows accepts reliably.
    loop count {
        SendEvent "!{Tab}"
        Sleep 60 ; brief delay to allow window switch animation
    }
}
