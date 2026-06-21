; =============================================================================
; WindowManagement module: helpers.ahk
; Small helper functions: notifications, window activation, automation cursor
; suppression, mouse centering. Pure functions (no top-level side effects).
; Extracted verbatim from WindowManagement.ahk (the entry point / source of truth).
; =============================================================================

; --- Helper Functions --------------------------------------------------------
ShowNotification_WM(message, durationMs := 1500) {
    ShowCenteredOverlay_Utils(message, durationMs, BANNER_ACCENT_ERROR)
}

; Activate window by winSpec; show graceful error and return false if not found.
TryActivateWindow_WM(winSpec, errorMessage := "❌ Error: Target window not found.") {
    if (!WinExist(winSpec)) {
        ShowNotification_WM(errorMessage)
        return false
    }
    try {
        WinActivate(winSpec)
        return true
    } catch {
        ShowNotification_WM(errorMessage)
        return false
    }
}

; Handy, Clip Angel, and WindowManagement identity: skip for per-monitor cycling, move-to-monitor, tile, and auto-cursor.
WM_IsExcludedIndicatorWindow(hwnd) {
    if (!hwnd)
        return false
    try {
        exe := StrLower(WinGetProcessName("ahk_id " hwnd))
    } catch {
        return false
    }
    if (exe = "handy.exe" || exe = "clipangel.exe")
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

WM_UsesAutomationDaemon() {
    global WM_USE_DAEMON := false, WM_USE_PIPE_IPC := false, WM_USE_EVENT_HOOK_CACHE := false
    return WM_USE_DAEMON && WM_USE_PIPE_IPC && WM_USE_EVENT_HOOK_CACHE
}

WMAutomation_SuppressCursorCentering(reason := "", durationMs := 0) {
    global g_WMAutomationSuppressUntil, g_WMAutomationSuppressReason
    durationMs := durationMs > 0 ? durationMs : WM_AUTOMATION_SWITCH_DEFAULT_MS
    g_WMAutomationSuppressUntil := A_TickCount + durationMs
    g_WMAutomationSuppressReason := reason
    if (WM_UsesAutomationDaemon()) {
        try WMIPC_BeginAutomationSwitch(reason, durationMs)
    }
    return g_WMAutomationSuppressUntil
}

WMAutomation_ClearCursorSuppression(reason := "") {
    global g_WMAutomationSuppressUntil, g_WMAutomationSuppressReason
    g_WMAutomationSuppressUntil := 0
    g_WMAutomationSuppressReason := ""
    if (WM_UsesAutomationDaemon()) {
        try WMIPC_EndAutomationSwitch(reason)
    }
}

WMAutomation_CursorCenteringSuppressed(hwnd := 0) {
    global g_WMAutomationSuppressUntil
    if (A_TickCount < g_WMAutomationSuppressUntil)
        return true
    if (WM_UsesAutomationDaemon()) {
        try {
            state := WMIPC_GetForegroundWindowState()
            if (state.Has("suppressCursorCentering") && state["suppressCursorCentering"])
                return true
        } catch {
        }
    }
    return false
}

WM_MaybeCenterMouse(hwnd, reason := "") {
    if (!hwnd || WMAutomation_CursorCenteringSuppressed(hwnd))
        return false
    MoveMouseToCenter(hwnd)
    return true
}
