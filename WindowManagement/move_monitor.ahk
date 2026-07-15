; =============================================================================
; WindowManagement module: move_monitor.ahk
; Move window to ordered monitor and MEH Alt+Tab
; Extracted verbatim from WindowManagement.ahk; loaded via #include into the
; WindowManagement.ahk process, which remains the entry point / source of truth.
; =============================================================================

MoveWinToOrderedMonitor(order) {
    idx := GetMonitorIndexByOrder(order)
    if (idx)
        MoveWinToMonitor(idx)
    else
        ShowNotification_WM("Monitor " order " not available (only " MonitorGetCount() " detected).")
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
    hwnd := 0
    state := ""
    if (WM_UsesAutomationDaemon()) {
        try {
            state := WMIPC_GetForegroundWindowState()
            if (state.Has("hwnd"))
                hwnd := Integer(state["hwnd"])
        } catch {
        }
    }
    if (!hwnd) {
        try {
            hwnd := WinExist("A")
        } catch {
            return
        }
    }
    if (!hwnd || hwnd = lastHwnd)
        return

    lastHwnd := hwnd

    if (A_TickCount - g_LastMouseClickTick < 1000)
        return

    if (WM_IsExcludedIndicatorWindow(hwnd))
        return

    if (WMAutomation_CursorCenteringSuppressed(hwnd))
        return

    MoveMouseToCenter(hwnd)
}

MoveMouseToCenter(hwnd) {
    static lastCenterTick := 0, lastCenterHwnd := 0
    ; Avoid showing two halos for the same window in rapid succession.
    if (hwnd = lastCenterHwnd && A_TickCount - lastCenterTick < 500)
        return
    lastCenterHwnd := hwnd
    lastCenterTick := A_TickCount

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

    ; Show a flash highlight around the cursor (lightweight indicator)
    ShowCursorFlash(centerX, centerY)
}

; ---------------------------------------------------------------------------
; Shows a lightweight flashing indicator at the cursor position.
; Flashes twice (150ms on, 100ms off, 150ms on) with a large red square.
; Uses size and motion for attention capture, minimizing GPU usage.
; ---------------------------------------------------------------------------
ShowCursorFlash(cx, cy) {
    static flashGui := 0, lastFlashTick := 0
    ; Prevent duplicate flashes in quick succession
    if (A_TickCount - lastFlashTick < 300)
        return
    lastFlashTick := A_TickCount

    ; Clean up any previous flash that might still be displayed
    if (flashGui && IsObject(flashGui)) {
        try flashGui.Destroy()
        flashGui := 0
    }

    ; Configuration: Large red square with border for visibility
    size := 250             ; 120×120 pixel square
    borderWidth := 3        ; 3-pixel border for enhanced visibility
    bgColor := "DF2935"     ; Bright red (colorblind-friendly)
    borderColor := "FFFFFF" ; White border
    alpha := 220            ; Semi-transparent

    ; Create the flash indicator GUI (fully guarded so errors never surface to user)
    try {
        flashGui := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x20 -DPIScale")
        flashGui.BackColor := bgColor

        ; Add border by creating a slightly larger outer GUI
        flashGui.Add("Text", "x0 y0 w" size " h" size " Background" bgColor)

        ; Position centered on cursor
        x := cx - (size // 2)
        y := cy - (size // 2)

        ; Show first flash
        flashGui.Show("NA x" x " y" y " w" size " h" size)
        WinSetTransparent(alpha, flashGui.Hwnd)
    } catch {
        ; Best-effort cleanup; avoid throwing from visual-only helper
        try {
            if (flashGui && IsObject(flashGui))
                flashGui.Destroy()
        }
        flashGui := 0
        return
    }

    ; Schedule flash animation: hide after 150ms, show again after 250ms, destroy after 400ms
    SetTimer(() => HideFlash(flashGui), -150)
    SetTimer(() => ShowFlash(flashGui, alpha), -250)
    SetTimer(() => DestroyFlash(flashGui), -400)
}

HideFlash(gui) {
    if (gui && IsObject(gui)) {
        try gui.Hide()
    }
}

ShowFlash(gui, alpha) {
    if (gui && IsObject(gui)) {
        try {
            gui.Show("NA")
            WinSetTransparent(alpha, gui.Hwnd)
        }
    }
}

DestroyFlash(gui) {
    if (gui && IsObject(gui)) {
        try gui.Destroy()
    }
}

; Move cursor to work-area center of AHK monitor index (spatial navigation when no window to move/cycle).
_WM_MoveCursorToMonitorWorkCenter(ahkMonIdx) {
    if (ahkMonIdx < 1 || ahkMonIdx > MonitorGetCount())
        return
    MonitorGet ahkMonIdx, &l, &t, &r, &b
    cx := (l + r) // 2
    cy := (t + b) // 2
    DllCall("user32\SetCursorPos", "int", cx, "int", cy)
    ShowCursorFlash(cx, cy)
}

WM_IsDesktopOrTaskbarClass(cls) {
    return cls = "Progman" || cls = "WorkerW" || cls = "Shell_TrayWnd" || cls = "Shell_SecondaryTrayWnd"
}

; -----------------------------------------------------------------------------
; Moves the active window to the specified monitor index and maximises it.
; Re-added because it was inadvertently removed during refactor.
; -----------------------------------------------------------------------------
MoveWinToMonitor(mon) {
    ; Validate monitor index
    if (mon > MonitorGetCount() || mon < 1) {
        ShowNotification_WM("Invalid monitor index: " mon)
        return
    }

    hwnd := 0
    try {
        hwnd := WinExist("A")
    } catch {
        hwnd := 0
    }
    if !hwnd {
        _WM_MoveCursorToMonitorWorkCenter(mon)
        return
    }

    if (WM_IsExcludedIndicatorWindow(hwnd)) {
        ShowNotification_WM("Cannot move this window (indicator / overlay).")
        return
    }

    try {
        activeClass := WinGetClass(hwnd)
    } catch {
        activeClass := ""
    }
    if (WM_IsDesktopOrTaskbarClass(activeClass)) {
        _WM_MoveCursorToMonitorWorkCenter(mon)
        return
    }

    sourceMon := WM_GetHwndMonitorIndex(hwnd)
    companionHwnd := sourceMon >= 1 ? WM_FindStrictSnapCompanion(hwnd, sourceMon) : 0

    ; AutoSlot: exchange whole-monitor FG layouts (full↔pair / half↔full / full↔full).
    if (sourceMon >= 1 && sourceMon != mon) {
        swapped := false
        try swapped := !!AutoSlot_TryForegroundSwap(hwnd, sourceMon, mon)
        catch
            swapped := false
        if (swapped)
            return
    }

    WMAutomation_SuppressCursorCentering("move_window_to_monitor", 800)

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

    ; Heal source companion and ensure moved window is foreground + cursor centered.
    WM_AfterLeavingMonitor(hwnd, sourceMon, companionHwnd, mon, -1)
    try AutoSlot_ScheduleRearrange(hwnd)
    catch {
    }
}
