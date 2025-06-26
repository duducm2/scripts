#Requires AutoHotkey v2.0+
#SingleInstance Force

; -----------------------------------------------------------------------------
; This script consolidates all Window Management hotkeys.
; -----------------------------------------------------------------------------

; --- Includes ----------------------------------------------------------------
#include %A_ScriptDir%\env.ahk

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
; Hotkeys: Ctrl+Alt+Shift+A/S/D/F (MEH)
; =============================================================================
^!+a:: MoveWinToMonitor(IS_WORK_ENVIRONMENT ? 1 : 1)
^!+s:: MoveWinToMonitor(IS_WORK_ENVIRONMENT ? 4 : 2)
^!+d:: MoveWinToMonitor(IS_WORK_ENVIRONMENT ? 2 : 3)
^!+f:: MoveWinToMonitor(IS_WORK_ENVIRONMENT ? 3 : 4)

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

    ; Give the OS a moment to finish activating the target window, then flash it.
    Sleep 80
    FlashActiveWindow()
}

; ----------------------------------------------------------------------------
; FlashActiveWindow  – draws an always-on-top yellow overlay covering
;                     the currently active window for ~500 ms so it	stands out.
; ----------------------------------------------------------------------------
FlashActiveWindow() {
    hwnd := WinExist("A")
    if (!hwnd) {
        return
    }

    ; Get active window position/size
    WinGetPos &wx, &wy, &ww, &wh, hwnd

    overlayW := 300
    overlayH := 200
    overlayX := wx + ((ww - overlayW) // 2)
    overlayY := wy + ((wh - overlayH) // 2)

    ; Create a borderless, click-through, opaque yellow overlay
    overlay := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x20") ; WS_EX_TRANSPARENT for click-through
    overlay.BackColor := "Yellow"
    overlay.Show("x" overlayX " y" overlayY " w" overlayW " h" overlayH)

    ; Ensure fully opaque (0=fully transparent, 255=opaque)
    WinSetTransparent(255, overlay)

    ; Destroy after 500 ms without blocking the script
    SetTimer (() => overlay.Destroy()), -500
}
