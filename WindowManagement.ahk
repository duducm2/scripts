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
global g_WindowCycleIndices := Map()  ; Keeps per-monitor cycling position
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
; Move Active Window to Monitor by POSITION (left-to-right order)
; Hotkeys: Ctrl+Alt+Shift+A/S/D/F correspond to 1st/2nd/3rd/4th monitors
; =============================================================================
^!+a:: MoveWinToOrderedMonitor(1)  ; Left-most
^!+s:: MoveWinToOrderedMonitor(2)  ; 2nd from the left
^!+d:: MoveWinToOrderedMonitor(3)  ; 3rd from the left
^!+f:: MoveWinToOrderedMonitor(4)  ; 4th from the left

^!+q:: CycleWindowsOnMonitor(1)  ; Cycle windows on monitor 1
^!+w:: CycleWindowsOnMonitor(2)  ; Cycle windows on monitor 2
^!+e:: CycleWindowsOnMonitor(3)  ; Cycle windows on monitor 3
^!+r:: CycleWindowsOnMonitor(4)  ; Cycle windows on monitor 4

MoveWinToOrderedMonitor(order) {
    idx := GetMonitorIndexByOrder(order)
    if (idx)
        MoveWinToMonitor(idx)
    else
        MsgBox "Monitor " order " not available (only " MonitorGetCount() " detected)."
}

GetMonitorIndexByOrder(order) {
    count := MonitorGetCount()
    if (order < 1 || order > count)
        return 0

    monitors := []
    loop count {
        i := A_Index
        MonitorGet i, &l, &t, &r, &b
        cx := (l + r) // 2  ; centre-X for ordering
        cy := (t + b) // 2  ; centre-Y (tie-breaker)
        monitors.Push({ idx: i, cx: cx, cy: cy })
    }

    ; Simple left-to-right ordering (with small vertical offset tolerance)
    ; This is what the user expects for the MEH hotkeys.
    n := monitors.Length
    loop n - 1 {
        i := A_Index
        loop n - i {
            j := A_Index
            a := monitors[j]
            b := monitors[j + 1]
            if (a.cx > b.cx || (a.cx == b.cx && a.cy > b.cy)) {
                monitors[j] := b
                monitors[j + 1] := a
            }
        }
    }

    return monitors[order].idx
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
~*LButton::
{
    global g_LastMouseClickTick
    g_LastMouseClickTick := A_TickCount
}
~*RButton::
{
    global g_LastMouseClickTick
    g_LastMouseClickTick := A_TickCount
}
~*MButton::
{
    global g_LastMouseClickTick
    g_LastMouseClickTick := A_TickCount
}

; ----------------------------------------------------------------------------
; Set a timer that monitors active-window changes and, when they are triggered
; by keyboard activity (i.e. not immediately after a mouse click), moves the
; cursor to the centre of the newly-activated window.
; ----------------------------------------------------------------------------
MonitorActiveWindow() {
    global g_LastMouseClickTick
    static lastHwnd := 0
    hwnd := WinExist("A")
    if (!hwnd || hwnd = lastHwnd)
        return

    lastHwnd := hwnd

    ; If the window became active shortly after a mouse click, assume the user
    ; activated it with the mouse and skip moving the pointer.
    if (A_TickCount - g_LastMouseClickTick < 1000)  ; 1000 ms threshold
        return

    ; --- Exclude specific applications (e.g., Snipping Tool) ---
    ; Attempt to retrieve the process name; some system-level or UWP windows may
    ; deny access, which would normally raise an exception and stop the script.
    ; By catching the error we keep the timer running and simply ignore that
    ; particular window.
    try {
        processName := WinGetProcessName("ahk_id " hwnd)
    } catch {
        return  ; Could not retrieve process name (e.g., access denied)
    }
    if (processName = "ScreenClippingHost.exe" || processName = "SnippingTool.exe") {
        return
    }

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

    ; Move the mouse cursor to the calculated centre point
    DllCall("SetCursorPos", "int", centerX, "int", centerY)

    ; Show a halo highlight around the cursor (shorter 250 ms for faster cycling)
    ShowCursorHalo(centerX, centerY, 40)
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
    ; Use a single, high-contrast colour that remains visible for most forms of colour-blindness
    ; and avoids multi-colour symbolism.
    colors := [
        "FFFFFF",  ; white
        "000000",  ; black
        "FFB000",  ; vivid orange–yellow
        "FFFF00",  ; bright yellow
        "3772FF",  ; strong blue
        "DF2935",  ; magenta-red
        "248A3D"   ; bold green
    ]
    ; Increased size for better visibility across multiple monitors
    outermostRadius := 140  ; pixels (was 70)
    bandThickness := 10     ; pixels (was 5)

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

; -----------------------------------------------------------------------------
; Moves the active window to the specified monitor index and maximises it.
; Re-added because it was inadvertently removed during refactor.
; -----------------------------------------------------------------------------
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

    ; Move mouse to the center of the window after the move
    Sleep 150 ; allow window animation to finish
    MoveMouseToCenter(hwnd)
}

; =============================================================================
; Cycle through visible windows on a monitor (top-to-bottom rows, left-to-right)
; Hotkeys: Ctrl+Alt+Shift+Q/W/E/R map to monitors 1-4 (left-to-right order)
; =============================================================================
CycleWindowsOnMonitor(order) {
    global g_WindowCycleIndices
    idx := GetMonitorIndexByOrder(order)
    if (!idx) {
        MsgBox "Monitor " order " not available (only " MonitorGetCount() " detected)."
        return
    }

    windows := GetVisibleWindowsOnMonitor(idx)
    if (windows.Length = 0) {
        ; Nothing to cycle on this monitor – re-centre cursor on current active window
        hwndCur := WinExist("A")
        if (hwndCur)
            MoveMouseToCenter(hwndCur)
        return
    }

    ; Determine starting position for this monitor
    pos := g_WindowCycleIndices.Has(idx) ? g_WindowCycleIndices.Get(idx) + 1 : 1
    if (pos > windows.Length)
        pos := 1

    ; Ensure we don't stay on the same window if hotkey is pressed rapidly.
    hwndCur := WinExist("A")
    startPos := pos
    loop windows.Length {
        target := windows[pos]
        if (target.hwnd != hwndCur)  ; found the next different window
            break
        ; Otherwise advance to next and wrap
        pos++
        if (pos > windows.Length)
            pos := 1
        ; If we've come full circle, all windows are the same – just break
        if (pos = startPos)
            break
    }

    g_WindowCycleIndices.Set(idx, pos)

    target := windows[pos]
    try WinActivate "ahk_id " target.hwnd
    catch {
        return
    }
    ; Wait until the window is active to avoid race conditions during rapid cycling
    WinWaitActive "ahk_id " target.hwnd, , 0.3
    Sleep 100  ; small delay for animation/focus stability
    MoveMouseToCenter(target.hwnd)
}

GetVisibleWindowsOnMonitor(mon) {
    ; Determine the monitor handle for the requested monitor index.
    ; We take a point near the centre of that monitor and ask Windows which
    ; monitor it belongs to.  This avoids any guessing about geometry and
    ; matches the result returned by MonitorFromWindow used below.
    MonitorGet mon, &ml, &mt, &mr, &mb
    cx := (ml + mr) // 2
    cy := (mt + mb) // 2
    ; POINT structure is passed **by value**, so combine X and Y into an int64.
    point64 := (cy & 0xFFFFFFFF) << 32 | (cx & 0xFFFFFFFF)
    hTarget := DllCall("MonitorFromPoint", "int64", point64, "uint", 2, "ptr") ; MONITOR_DEFAULTTONEAREST

    result := []
    hwnds := WinGetList()

    GWL_EXSTYLE := -20
    WS_EX_TOOLWINDOW := 0x00000080
    TOL := 40  ; tolerance for considering two rows the same when sorting

    for hwnd in hwnds {
        zIdx := A_Index  ; preserves original Z-order (lower = bottom, higher = top)
        try {
            ; Skip minimised windows (-1). Keep normal (0) and maximised (+1).
            if (WinGetMinMax(hwnd) = -1)
                continue

            ; Skip invisible windows
            if !DllCall("IsWindowVisible", "ptr", hwnd)
                continue

            ; Skip toolwindows (notification icons, etc.)
            exStyle := DllCall("GetWindowLongPtr", "ptr", hwnd, "int", GWL_EXSTYLE, "ptr")
            if (exStyle & WS_EX_TOOLWINDOW)
                continue

            ; Verify window belongs to the target monitor according to Windows.
            hMon := DllCall("MonitorFromWindow", "ptr", hwnd, "uint", 2, "ptr") ; nearest monitor
            if (hMon != hTarget)
                continue

            ; Exclude desktop/worker windows or those without title
            class := WinGetClass(hwnd)
            if (class = "Progman" || class = "WorkerW")
                continue

            title := WinGetTitle(hwnd)
            if (title = "")
                continue

            ; Cache position for sorting
            rect := Buffer(16, 0)
            if DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect) {
                left := NumGet(rect, 0, "int")
                top := NumGet(rect, 4, "int")
            } else {
                left := 0, top := 0
            }

            result.Push({ hwnd: hwnd, left: left, top: top, z: zIdx })
        } catch {
            ; Ignore any errors for this window and continue enumerating.
            continue
        }
    }

    ; Manual bubble-sort: first by top (row), then by left (column).
    n := result.Length
    loop n - 1 {
        i := A_Index
        loop n - i {
            j := A_Index
            a := result[j]
            b := result[j + 1]
            needSwap := false
            if (a.top > b.top + TOL) {
                needSwap := true
            } else if (Abs(a.top - b.top) <= TOL) {
                if (a.left > b.left) {
                    needSwap := true
                } else if (a.left = b.left) {
                    ; Same position – fall back to Z-order (prefer topmost window first).
                    if (a.z < b.z)
                        needSwap := true
                }
            }
            if (needSwap) {
                result[j] := b
                result[j + 1] := a
            }
        }
    }
    return result
}
