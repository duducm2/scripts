; =============================================================================
; Utils module: f11_fullscreen.ahk
; F11 fullscreen detection and toggle
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; F11 fullscreen detection / toggle (shared by WindowManagement + Chrome detach)
; =============================================================================

_WMF11_GetWindowRectHwnd(hwnd, &left, &top, &right, &bottom) {
    rect := Buffer(16, 0)
    if !DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect)
        return false
    left := NumGet(rect, 0, "int")
    top := NumGet(rect, 4, "int")
    right := NumGet(rect, 8, "int")
    bottom := NumGet(rect, 12, "int")
    return true
}

_WMF11_GetHwndMonitorIndex(hwnd) {
    if (!hwnd)
        return 0
    try {
        hMon := DllCall("MonitorFromWindow", "ptr", hwnd, "uint", 2, "ptr")
        loop MonitorGetCount() {
            MonitorGet A_Index, &ml, &mt, &mr, &mb
            cx := (ml + mr) // 2
            cy := (mt + mb) // 2
            point64 := (cy & 0xFFFFFFFF) << 32 | (cx & 0xFFFFFFFF)
            if (Integer(DllCall("MonitorFromPoint", "int64", point64, "uint", 2, "ptr")) = Integer(hMon))
                return A_Index
        }
    } catch {
    }
    return 0
}

_WMF11_IsDesktopOrTaskbarClass(cls) {
    return cls = "Progman" || cls = "WorkerW" || cls = "Shell_TrayWnd" || cls = "Shell_SecondaryTrayWnd"
}

_WMF11_IsExcludedIndicatorWindow(hwnd) {
    if (!hwnd)
        return false
    try {
        exe := StrLower(WinGetProcessName("ahk_id " hwnd))
    } catch {
        return false
    }
    if (exe = "handy.exe")
        return true
    try {
        title := WinGetTitle(hwnd)
    } catch {
        return false
    }
    if (InStr(StrLower(title), "windowmanagement.ahk"))
        return true
    return false
}

_WMF11_BackgroundHwndOnAnyScriptMonitor(hwnd) {
    try {
        hMon := DllCall("MonitorFromWindow", "ptr", hwnd, "uint", 2, "ptr")
        loop MonitorGetCount() {
            MonitorGet A_Index, &ml, &mt, &mr, &mb
            cx := (ml + mr) // 2
            cy := (mt + mb) // 2
            point64 := (cy & 0xFFFFFFFF) << 32 | (cx & 0xFFFFFFFF)
            if (Integer(DllCall("MonitorFromPoint", "int64", point64, "uint", 2, "ptr")) = Integer(hMon))
                return true
        }
    } catch {
    }
    return false
}

; F11 fullscreen: fills monitor (often past work area) or work area with no caption; not ordinary Win-maximize.
WM_WindowIsF11FullscreenRejectReason(hwnd) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return "no_hwnd"
    try {
        minMax := WinGetMinMax(hwnd)
        if (minMax = -1)
            return "minimized"
        if !DllCall("IsWindowVisible", "ptr", hwnd)
            return "not_visible"
        exStyle := DllCall("GetWindowLongPtr", "ptr", hwnd, "int", -20, "ptr")
        if (exStyle & 0x00000080)
            return "toolwindow"
        class := WinGetClass(hwnd)
        if (class = "Progman" || class = "WorkerW")
            return "desktop_class"
        if (_WMF11_IsDesktopOrTaskbarClass(class))
            return "taskbar_class"
        if (WinGetTitle(hwnd) = "")
            return "empty_title"
        if (_WMF11_IsExcludedIndicatorWindow(hwnd))
            return "excluded_indicator"
        if (!_WMF11_BackgroundHwndOnAnyScriptMonitor(hwnd))
            return "not_script_monitor"
        mon := _WMF11_GetHwndMonitorIndex(hwnd)
        if (!mon)
            return "no_monitor"
        MonitorGet mon, &ml, &mt, &mr, &mb
        MonitorGetWorkArea mon, &wl, &wt, &wr, &wb
        if (!_WMF11_GetWindowRectHwnd(hwnd, &left, &top, &right, &bottom))
            return "no_rect"
        tol := 24
        fillsMonitor := (Abs(left - ml) <= tol && Abs(top - mt) <= tol && Abs(right - mr) <= tol && Abs(bottom - mb) <=
        tol)
        fillsWorkArea := (Abs(left - wl) <= tol && Abs(top - wt) <= tol && Abs(right - wr) <= tol && Abs(bottom - wb) <=
        tol)
        extendsPastWorkArea := (bottom > wb + tol || right > wr + tol || left < wl - tol || top < wt - tol)
        style := DllCall("GetWindowLongPtr", "ptr", hwnd, "int", -16, "ptr")
        hasCaption := !!(style & 0x00C00000)
        if (minMax = 1 && fillsWorkArea && !extendsPastWorkArea && hasCaption)
            return "win_maximized"
        if (fillsMonitor && extendsPastWorkArea)
            return ""
        if (fillsWorkArea && !hasCaption)
            return ""
        if (!fillsMonitor && !fillsWorkArea)
            return "not_monitor_fill"
        if (fillsMonitor && !extendsPastWorkArea)
            return "within_work_area"
        if (fillsWorkArea && hasCaption)
            return "work_area_with_caption"
        return "no_match"
    } catch as err {
        return "exception:" . err.Message
    }
}

WM_WindowIsF11Fullscreen(hwnd) {
    return WM_WindowIsF11FullscreenRejectReason(hwnd) = ""
}

WM_WaitForF11State(hwnd, wantFullscreen, timeoutMs := 500) {
    deadline := A_TickCount + timeoutMs
    loop {
        if (WM_WindowIsF11Fullscreen(hwnd) = wantFullscreen)
            return true
        if (A_TickCount >= deadline)
            break
        Sleep 50
    }
    return WM_WindowIsF11Fullscreen(hwnd) = wantFullscreen
}

WM_EnsureForegroundForSend(hwnd, timeoutMs := 2000) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        try WinActivate "ahk_id " hwnd
        catch {
            return false
        }
        if WinActive("ahk_id " hwnd)
            return true
        Sleep 50
    }
    return WinActive("ahk_id " hwnd)
}

WM_ExitF11FullscreenForHwnd(hwnd, settleMs := 1200) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    if (!WM_WindowIsF11Fullscreen(hwnd))
        return true
    try {
        if !WM_EnsureForegroundForSend(hwnd)
            return false
        Sleep 80
        ClipAngel_ReleaseChordModifiersForSend()
        SendInput "{F11}"
        return WM_WaitForF11State(hwnd, false, settleMs)
    } catch {
        return false
    }
}

WM_EnterF11FullscreenForHwnd(hwnd, settleMs := 1200) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    if (WM_WindowIsF11Fullscreen(hwnd))
        return true
    try {
        if !WM_EnsureForegroundForSend(hwnd)
            return false
        Sleep 80
        ClipAngel_ReleaseChordModifiersForSend()
        SendInput "{F11}"
        return WM_WaitForF11State(hwnd, true, settleMs)
    } catch {
        return false
    }
}

