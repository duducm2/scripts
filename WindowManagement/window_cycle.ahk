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

    hwndCur := 0
    try hwndCur := WinExist("A")
    catch hwndCur := 0
        ; Next window is the one after whatever is actually in front, when that
        ; window is on this monitor. If it is not (focus still on another monitor,
        ; or the last target never took focus), advance the stored index. Deleting
        ; that index and always restarting at windows[1] left the cycle stuck on
        ; the first window until a trip to another monitor made activation stick.
        activeIdx := 0
    loop windows.Length {
        if (windows[A_Index].hwnd = hwndCur) {
            activeIdx := A_Index
            break
        }
    }

    if (activeIdx)
        pos := activeIdx + 1
    else if (g_WindowCycleIndices.Has(idx))
        pos := g_WindowCycleIndices[idx] + 1
    else
        pos := 1
    if (pos > windows.Length || pos < 1)
        pos := 1

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

    g_WindowCycleIndices.Set(idx, pos)
    target := windows[pos]
    try WinActivate "ahk_id " target.hwnd
    catch {
        ShowNotification_WM("Error: Target window not found.")
        return
    }
    ; Short wait, then one foreground nudge if the target did not take focus
    ; (common when the previous foreground is on another monitor).
    if !WinWaitActive("ahk_id " target.hwnd, , 0.05)
        Cycle_ForceForeground(target.hwnd)
    WinWaitActive "ahk_id " target.hwnd, , 0.2
    ; Draw the square now. MoveMouseToCenter ignores a second call for this hwnd
    ; within 500ms, so the foreground timer does not flash again.
    MoveMouseToCenter(target.hwnd)

    keepMon := Cycle_CachedKeepMonitor()
    if (keepMon && idx != keepMon)
        FocusMode_RequestDisableCrossProcess()
}

; Focus-mode sentinel lives under A_ScriptDir (often Google Drive). Repeat
; presses of the cycle hotkey reuse the last read for about one second.
Cycle_CachedKeepMonitor() {
    static lastTick := 0, lastMon := 0
    if (lastTick && (A_TickCount - lastTick) < 1000)
        return lastMon
    lastMon := FocusMode_ReadKeepMonitorFromFile()
    lastTick := A_TickCount
    return lastMon
}

; WinActivate often fails while the foreground window is on another monitor.
; Attach this thread to that window's thread for one SetForegroundWindow call.
Cycle_ForceForeground(hwnd) {
    try {
        fg := DllCall("GetForegroundWindow", "ptr")
        if (!hwnd || fg = hwnd)
            return
        curTid := DllCall("GetCurrentThreadId", "UInt")
        fgTid := DllCall("GetWindowThreadProcessId", "ptr", fg, "ptr", 0, "UInt")
        tgtTid := DllCall("GetWindowThreadProcessId", "ptr", hwnd, "ptr", 0, "UInt")
        if (fgTid && fgTid != curTid)
            DllCall("AttachThreadInput", "UInt", curTid, "UInt", fgTid, "Int", 1)
        if (tgtTid && tgtTid != curTid)
            DllCall("AttachThreadInput", "UInt", curTid, "UInt", tgtTid, "Int", 1)
        DllCall("SetForegroundWindow", "ptr", hwnd)
        DllCall("BringWindowToTop", "ptr", hwnd)
        if (fgTid && fgTid != curTid)
            DllCall("AttachThreadInput", "UInt", curTid, "UInt", fgTid, "Int", 0)
        if (tgtTid && tgtTid != curTid)
            DllCall("AttachThreadInput", "UInt", curTid, "UInt", tgtTid, "Int", 0)
    } catch {
    }
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

            ; Skip only a window fully inside one already accepted (higher z-order).
            ; A center-point test dropped windows that still stuck out, so they
            ; never appeared in the cycle until z-order changed.
            covered := false
            for win in visible {
                if (left >= win.left && right <= win.right
                    && top >= win.top && bottom <= win.bottom) {
                    covered := true
                    break
                }
            }
            if (covered)
                continue            ; fully inside a higher window

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
