; =============================================================================
; WindowManagement module: window_cycle.ahk
; Cycle / minimize / close visible windows on a monitor (by left-to-right order)
; Extracted verbatim from WindowManagement.ahk; loaded via #include into the
; WindowManagement.ahk process, which remains the entry point / source of truth.
; =============================================================================

; =============================================================================
; Cycle through visible windows on a monitor (top-to-bottom rows, left-to-right)
; Hotkeys: Ctrl+Alt+Win+Q/W/E/R map to monitors 1-4 (left-to-right order)
; =============================================================================
CycleWindowsOnMonitor(order) {
    global g_WindowCycleIndices
    idx := GetMonitorIndexByOrder(order)
    if (!idx) {
        ShowNotification_WM("Monitor " order " not available (only " MonitorGetCount() " detected).")
        return
    }

    windows := GetVisibleWindowsOnMonitor(idx)
    if (windows.Length = 0) {
        ; Empty monitor (or only excluded overlays): jump pointer to that screen instead of trapping on the old one.
        _WM_MoveCursorToMonitorWorkCenter(idx)
        return
    }

    ; If the currently active window is on a **different** monitor, reset the cycle
    ; so we start from the topmost visible window instead of cycling to the next.
    hwndCur := 0
    try {
        hwndCur := WinExist("A")
    } catch {
        ; No active window available, will reset cycle
        hwndCur := 0
    }
    hMonCur := 0
    if (hwndCur) {
        try {
            hMonCur := DllCall("MonitorFromWindow", "ptr", hwndCur, "uint", 2, "ptr") ; nearest monitor
        } catch {
            hMonCur := 0
        }
    }

    ; Get handle for the target monitor.
    MonitorGet idx, &l, &t, &r, &b
    cx := (l + r) // 2, cy := (t + b) // 2
    point64 := (cy & 0xFFFFFFFF) << 32 | (cx & 0xFFFFFFFF)
    hMonTarget := DllCall("MonitorFromPoint", "int64", point64, "uint", 2, "ptr")

    if (hMonCur != hMonTarget) {
        ; Coming from another monitor – reset cycle index to 0 so first pick is topmost
        if (g_WindowCycleIndices.Has(idx))
            g_WindowCycleIndices.Delete(idx)
    }

    ; Determine starting position: 1 past the currently-active window (if it belongs to this
    ; monitor) or the very first window otherwise.  This avoids stale indices and always bases
    ; cycling on the window that the user is actually looking at.
    activeIdx := 0
    loop windows.Length {
        if (windows[A_Index].hwnd = hwndCur) {
            activeIdx := A_Index
            break
        }
    }

    pos := activeIdx ? activeIdx + 1 : 1
    if (pos > windows.Length)
        pos := 1

    ; Remember the new position for subsequent cycles (only if we stayed on the same monitor).
    g_WindowCycleIndices.Set(idx, pos)

    ; Ensure we don't stay on the same window if hotkey is pressed rapidly.
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

    target := windows[pos]
    try WinActivate "ahk_id " target.hwnd
    catch {
        ShowNotification_WM("Error: Target window not found.")
        return
    }
    ; Wait until the window is active to avoid race conditions during rapid cycling
    WinWaitActive "ahk_id " target.hwnd, , 0.3
    ; The MonitorActiveWindow timer will centre the cursor automatically, so avoid
    ; calling it here to prevent duplicate halo flashes.
    Sleep 100  ; small delay for animation/focus stability

    keepMon := FocusMode_ReadKeepMonitorFromFile()
    if (keepMon && idx != keepMon)
        FocusMode_RequestDisableCrossProcess()
}

GetVisibleWindowsOnMonitor(mon, skipDaemon := false) {
    ; Daemon path: use O(1) cache when flags enabled (Phase 3)
    daemonFallback := ""
    if (WM_UsesAutomationDaemon() && !skipDaemon) {
        try {
            winList := WMIPC_GetVisibleWindowsByMonitor(mon)
            if (winList.Length > 0) {
                visible := []
                for w in winList {
                    h := Integer(w["hwnd"])
                    if (WM_IsExcludedIndicatorWindow(h))
                        continue
                    visible.Push({ hwnd: h, left: Integer(w["left"]), top: Integer(w["top"]), right: Integer(
                        w["right"]), bottom: Integer(w["bottom"]), z: Integer(w["z"]) })
                }
                ; Daemon uses EnumDisplayMonitors slot (mon); AHK uses MonitorGet(mon). If they diverge,
                ; every HWND can sit on a different HMONITOR than the work area center expects — fall back to legacy.
                MonitorGet mon, &dml, &dmt, &dmr, &dmb
                dcx := (dml + dmr) // 2
                dcy := (dmt + dmb) // 2
                dpoint64 := (dcy & 0xFFFFFFFF) << 32 | (dcx & 0xFFFFFFFF)
                hExpected := DllCall("MonitorFromPoint", "int64", dpoint64, "uint", 2, "ptr")
                onMonitor := []
                for v in visible {
                    try {
                        hMon := DllCall("MonitorFromWindow", "ptr", v.hwnd, "uint", 2, "ptr")
                        if (Integer(hMon) = Integer(hExpected))
                            onMonitor.Push(v)
                    } catch {
                    }
                }
                if (onMonitor.Length > 0) {
                    nSort := onMonitor.Length
                    if (nSort > 1) {
                        loop nSort - 1 {
                            i := A_Index
                            loop nSort - i {
                                j := A_Index
                                rowDiff := onMonitor[j].top - onMonitor[j + 1].top
                                if (rowDiff > 40 || (Abs(rowDiff) <= 40 && onMonitor[j].left > onMonitor[j + 1].left)) {
                                    tmp := onMonitor[j]
                                    onMonitor[j] := onMonitor[j + 1]
                                    onMonitor[j + 1] := tmp
                                }
                            }
                        }
                    }
                    return onMonitor
                }
                daemonFallback := "daemon_hmon_mismatch"
            }
            if (daemonFallback = "")
                daemonFallback := "daemon_empty_list"
        } catch {
            daemonFallback := "daemon_exception"
        }
    }
    ; Step-1: determine target monitor handle --------------------------------
    MonitorGet mon, &ml, &mt, &mr, &mb
    cx := (ml + mr) // 2
    cy := (mt + mb) // 2
    point64 := (cy & 0xFFFFFFFF) << 32 | (cx & 0xFFFFFFFF)
    hTarget := DllCall("MonitorFromPoint", "int64", point64, "uint", 2, "ptr")

    ; Enumerate all windows – WinGetList() returns them in top-to-bottom z-order
    hwnds := WinGetList()

    GWL_EXSTYLE := -20
    WS_EX_TOOLWINDOW := 0x00000080
    TOL := 40  ; tolerance when deciding if two windows share a “row”

    visible := []      ; windows that remain at least PARTIALLY visible

    for hwnd in hwnds {
        zIdx := hwnds.Length - A_Index  ; 0 = topmost, grows toward bottom

        try {
            ; --- basic eligibility checks (unchanged) ----------------------
            if (WinGetMinMax(hwnd) = -1)
                continue            ; minimised
            if !DllCall("IsWindowVisible", "ptr", hwnd)
                continue
            exStyle := DllCall("GetWindowLongPtr", "ptr", hwnd, "int", GWL_EXSTYLE, "ptr")
            if (exStyle & WS_EX_TOOLWINDOW)
                continue            ; skip tool windows (e.g., floating toolbars)
            hMon := DllCall("MonitorFromWindow", "ptr", hwnd, "uint", 2, "ptr")
            if (Integer(hMon) != Integer(hTarget))
                continue            ; not on the requested monitor
            class := WinGetClass(hwnd)
            if (class = "Progman" || class = "WorkerW")
                continue            ; desktop / worker windows
            title := WinGetTitle(hwnd)
            if (title = "")
                continue            ; unnamed (often invisible) windows
            if (WM_IsExcludedIndicatorWindow(hwnd))
                continue

            ; --- geometry --------------------------------------------------
            rect := Buffer(16, 0)
            if !DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect)
                continue

            left := NumGet(rect, 0, "int")
            top := NumGet(rect, 4, "int")
            right := NumGet(rect, 8, "int")
            bottom := NumGet(rect, 12, "int")

            ; --- visibility heuristic -------------------------------------
            centerX := (left + right) // 2
            centerY := (top + bottom) // 2

            covered := false
            for win in visible {
                if (centerX >= win.left && centerX <= win.right
                    && centerY >= win.top && centerY <= win.bottom) {
                    covered := true
                    break
                }
            }
            if (covered)
                continue            ; completely concealed by a higher window

            ; Otherwise, accept it as visible
            visible.Push({ hwnd: hwnd, left: left, top: top, right: right,
                bottom: bottom, z: zIdx })
        } catch {
            continue                ; ignore windows that throw on inspection
        }
    }

    ; ──────────────────────────────────────────────────────────────
    ; Re-order accepted windows: by Y (top→bottom), then X (left→right)
    ; ──────────────────────────────────────────────────────────────
    n := visible.Length
    if (n > 1) {
        loop n - 1 {
            i := A_Index
            loop n - i {
                j := A_Index
                rowDiff := visible[j].top - visible[j + 1].top
                if (rowDiff > TOL)                         ; lower row → move down
                || (Abs(rowDiff) <= TOL                  ; same “row”
                && visible[j].left > visible[j + 1].left) {
                    temp := visible[j]
                    visible[j] := visible[j + 1]
                    visible[j + 1] := temp
                }
            }
        }
    }

    return visible
}

; =============================================================================
; Minimize the active window on the specified monitor
; Function: MinimizeWindowOnMonitor(order)
; =============================================================================
MinimizeWindowOnMonitor(order) {
    idx := GetMonitorIndexByOrder(order)
    if (!idx) {
        ShowNotification_WM("Monitor " order " not available (only " MonitorGetCount() " detected).")
        return
    }

    ; Get the active window on the target monitor
    windows := GetVisibleWindowsOnMonitor(idx)
    if (windows.Length = 0) {
        ShowNotification_WM("No windows found on monitor " order)
        return
    }

    ; Get the topmost window on the monitor (first in the list)
    targetWindow := windows[1]

    try {
        ; Activate the window first
        WinActivate "ahk_id " targetWindow.hwnd
        ; Wait briefly for activation
        Sleep 100
        ; Then minimize it
        WinMinimize "ahk_id " targetWindow.hwnd
    } catch Error as e {
        ShowNotification_WM("Failed to minimize window on monitor " order ": " e.Message)
    }
}

; =============================================================================
; Close the active window on the specified monitor
; Function: CloseWindowOnMonitor(order)
; =============================================================================
CloseWindowOnMonitor(order) {
    idx := GetMonitorIndexByOrder(order)
    if (!idx) {
        ShowNotification_WM("Monitor " order " not available (only " MonitorGetCount() " detected).")
        return
    }

    ; Close always uses legacy WinGetList enumeration so the list matches MonitorGet(idx); daemon IPC can
    ; disagree with AHK monitor numbering when focus is on other displays.
    windows := GetVisibleWindowsOnMonitor(idx, true)
    if (windows.Length = 0) {
        ShowNotification_WM("No windows found on monitor " order)
        return
    }

    ; Always close spatial [1] (Y then X sort). Foreground-based picking broke when a focused HWND on an
    ; adjacent monitor (esp. M2 next to M1) still matched MonitorFromWindow to M1 or appeared in the list.
    targetWindow := windows[1]

    try {
        th := targetWindow.hwnd
        ; Close without stealing focus first (works better when foreground is on another monitor); retry with
        ; activate if the window ignores background WM_CLOSE.
        WinClose "ahk_id " th
        if !WinWaitClose("ahk_id " th, , 0.2) {
            try WinShow "ahk_id " th
            WinActivate "ahk_id " th
            WinWaitActive "ahk_id " th, , 1.2
            WinClose "ahk_id " th
            if !WinWaitClose("ahk_id " th, , 0.35) {
                PostMessage 0x0010, 0, 0, , "ahk_id " th  ; WM_CLOSE — some apps only honor async close
                WinWaitClose "ahk_id " th, , 1.5
            }
        }
    } catch Error as e {
        ShowNotification_WM("Failed to close window on monitor " order ": " e.Message)
    }
}
