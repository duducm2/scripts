#Requires AutoHotkey v2.0+
#SingleInstance Force

; -----------------------------------------------------------------------------
; This script consolidates all Window Management hotkeys.
; -----------------------------------------------------------------------------

; --- Includes ----------------------------------------------------------------
#include %A_ScriptDir%\env.ahk

; --- Globals & Timers --------------------------------------------------------
global g_LastActiveHwnd := 0
global g_LastMouseClickTick := 0   ; Timestamp of the most recent mouse click (A_TickCount)
SetTimer MonitorActiveWindow, 100  ; Check 10× per second

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
^!+d:: MoveWinToMonitor(IS_WORK_ENVIRONMENT ? 2 : 4)
^!+f:: MoveWinToMonitor(IS_WORK_ENVIRONMENT ? 3 : 3)

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

    ; Wait briefly to allow the window to activate
    Sleep 250
}

; ----------------------------------------------------------------------------
; Mouse click hooks (update g_LastMouseClickTick)
; ----------------------------------------------------------------------------
~*LButton:: g_LastMouseClickTick := A_TickCount
~*RButton:: g_LastMouseClickTick := A_TickCount
~*MButton:: g_LastMouseClickTick := A_TickCount

; ----------------------------------------------------------------------------
; Set a timer that monitors active-window changes and, when they are triggered
; by keyboard activity (i.e. not immediately after a mouse click), moves the
; cursor to the centre of the newly-activated window.
; ----------------------------------------------------------------------------
MonitorActiveWindow() {
    static lastHwnd := 0
    hwnd := WinExist("A")
    if (!hwnd || hwnd = lastHwnd)
        return

    lastHwnd := hwnd

    ; If the window became active shortly after a mouse click, assume the user
    ; activated it with the mouse and skip moving the pointer.
    if (A_TickCount - g_LastMouseClickTick < 400)  ; 400 ms threshold
        return

    MoveMouseToCenter(hwnd)
}

MoveMouseToCenter(hwnd) {
    if !hwnd
        return

    rect := Buffer(16, 0)
    if !DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect)
        return

    left := NumGet(rect, 0, "int")
    top := NumGet(rect, 4, "int")
    right := NumGet(rect, 8, "int")
    bottom := NumGet(rect, 12, "int")

    centerX := left + (right - left) // 2
    centerY := top + (bottom - top) // 2

    DllCall("SetCursorPos", "int", centerX, "int", centerY)

    ; Show a halo highlight around the cursor
    ShowCursorHalo(centerX, centerY)
}

; ---------------------------------------------------------------------------
; Shows a temporary yellow halo (circle) centred at the given screen position.
; The halo lasts 500 ms, radius 40 px, thickness 8 px, semi-transparent.
; ---------------------------------------------------------------------------
ShowCursorHalo(cx, cy, duration := 500, alpha := 200) {
    static haloGuis := [], destroyTimer := 0

    ; Clean up any previous halos that might still be displayed
    if (haloGuis.Length > 0) {
        for eachGui in haloGuis {
            try eachGui.Destroy()
        }
        haloGuis := []
    }

    ; --- Configuration for the new multi-colored, thicker, larger halo ---
    colors := ["FF0000", "FFBF00", "FFFF00", "80FF00", "00FFFF", "0080FF", "8000FF"] ; Rainbow
    outermostRadius := 70  ; pixels
    bandThickness := 5     ; pixels

    currentRadius := outermostRadius
    for color in colors {
        ; Create a new GUI for this color band
        newGui := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x20 -DPIScale")
        newGui.BackColor := color
        haloGuis.Push(newGui)

        ; Build ring region for this band
        outerD := currentRadius * 2
        offset := bandThickness
        hOuter := DllCall("gdi32\CreateEllipticRgn", "int", 0, "int", 0, "int", outerD, "int", outerD, "ptr")
        hInner := DllCall("gdi32\CreateEllipticRgn", "int", offset, "int", offset, "int", outerD - offset, "int",
            outerD - offset, "ptr")
        RGN_DIFF := 3
        DllCall("gdi32\CombineRgn", "ptr", hOuter, "ptr", hOuter, "ptr", hInner, "int", RGN_DIFF)
        DllCall("user32\SetWindowRgn", "ptr", newGui.Hwnd, "ptr", hOuter, "int", true)
        DllCall("gdi32\DeleteObject", "ptr", hInner)

        ; Show the GUI at the correct position (top-left offset)
        newGui.Show("NA x" (cx - currentRadius) " y" (cy - currentRadius) " w" outerD " h" outerD)
        WinSetTransparent(alpha, newGui.Hwnd)

        currentRadius -= bandThickness ; Shrink radius for the next color band
    }

    ; Schedule destruction of all halo GUIs after the specified duration
    if IsObject(destroyTimer) {
        SetTimer(destroyTimer, 0)  ; cancel previous timer if it exists
    }
    destroyTimer := DestroyHalos.Bind(haloGuis)
    SetTimer(destroyTimer, -duration)
}

DestroyHalos(guisArray) {
    for eachGui in guisArray {
        try eachGui.Destroy()
    }
    guisArray.Length := 0
}
