#Requires AutoHotkey v2.0+
#SingleInstance Force

; -----------------------------------------------------------------------------
; This script consolidates all Window Management hotkeys.
; -----------------------------------------------------------------------------

; --- Includes ----------------------------------------------------------------
#include %A_ScriptDir%\env.ahk

; --- Globals & Timers --------------------------------------------------------
global g_LastActiveHwnd := 0
global g_OverlayHwnd := 0

; Start a timer that checks for active-window changes every 150 ms.
SetTimer ActiveWindowMonitor, 150

ActiveWindowMonitor() {
    global g_LastActiveHwnd, g_OverlayHwnd
    curr := WinExist("A")
    if (!curr) {
        return
    }

    ; Ignore transitions to our own overlay window
    if (curr = g_OverlayHwnd)
        return

    if (curr != g_LastActiveHwnd) {
        g_LastActiveHwnd := curr
        if (ShouldFlashWindow(curr)) {
            FlashActiveWindow()
        }
    }
}

; Returns true if the given hwnd represents a normal application window we want to flash.
ShouldFlashWindow(hwnd) {
    ; Skip if minimized
    if (WinGetMinMax(hwnd) = 1)
        return false

    ; Skip tiny windows (taskbar buttons, tooltips, etc.)
    WinGetPos &wx, &wy, &ww, &wh, hwnd
    if (ww < 300 || wh < 200)
        return false

    ; Skip the Windows taskbar itself
    class := WinGetClass(hwnd)
    if (class = "Shell_TrayWnd" || class = "Button" || class = "MSTaskListWClass")
        return false

    return true
}

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
^!+s:: MoveWinToMonitor(IS_WORK_ENVIRONMENT ? 2 : 2)
^!+d:: MoveWinToMonitor(IS_WORK_ENVIRONMENT ? 3 : 3)
^!+f:: MoveWinToMonitor(IS_WORK_ENVIRONMENT ? 4 : 4)

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
    if (count < 1)
        return

    ; Temporarily release Ctrl/Shift so they don't interfere (Ctrl+Alt+Tab or Shift+Alt+Tab).
    ctrlHeld := GetKeyState("Ctrl", "P")
    shiftHeld := GetKeyState("Shift", "P")

    if (ctrlHeld)
        SendEvent "{Ctrl Up}"
    if (shiftHeld)
        SendEvent "{Shift Up}"

    ; Perform Alt+Tab sequence
    SendEvent "{Alt Down}"
    SendEvent Format("{Tab %d}", count)
    SendEvent "{Alt Up}"

    ; Restore original modifier state
    if (shiftHeld)
        SendEvent "{Shift Down}"
    if (ctrlHeld)
        SendEvent "{Ctrl Down}"

    ; Wait for the window to activate, then flash.
    Sleep 250
    FlashActiveWindow()
}

; ----------------------------------------------------------------------------
; FlashActiveWindow  – draws an always-on-top yellow overlay covering
;                     the currently active window for ~500 ms so it	stands out.
; ----------------------------------------------------------------------------
FlashActiveWindow() {
    global g_OverlayHwnd

    hwnd := WinExist("A")
    if (!hwnd) {
        return
    }

    ; If a previous overlay is still shown, close it now to avoid multiple squares.
    if (g_OverlayHwnd && WinExist("ahk_id " g_OverlayHwnd)) {
        try WinClose "ahk_id " g_OverlayHwnd
        g_OverlayHwnd := 0
        Sleep 10
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
    overlay.Show("NA x" overlayX " y" overlayY " w" overlayW " h" overlayH)

    ; Store overlay handle to avoid self-triggering in monitor
    g_OverlayHwnd := overlay.Hwnd

    ; Ensure fully opaque (0=fully transparent, 255=opaque)
    WinSetTransparent(255, overlay)

    ; Destroy after 500 ms without blocking the script and clear handle reference
    SetTimer (() => (overlay.Destroy(), g_OverlayHwnd := 0)), -500
}
