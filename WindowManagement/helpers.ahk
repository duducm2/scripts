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

; Teams chrome hard-exclude: sharing control bar + meeting compact view only.
; Chat (teams-chat.md) and full meeting (teams-meeting.md) MUST stay eligible.
; Compact often shares a meeting-like Win32 title — distinguish via UIA
; "Maximize meeting window" only (never short-height on titled meetings).
WM_IsTeamsChromeHwnd(hwnd) {
    if (!hwnd)
        return false
    exe := ""
    class := ""
    title := ""
    try exe := StrLower(WinGetProcessName("ahk_id " hwnd))
    catch {
    }
    try class := WinGetClass("ahk_id " hwnd)
    catch {
    }
    isTeamsExe := (exe = "ms-teams.exe" || exe = "teams.exe" || exe = "msteams.exe")
    if (!isTeamsExe && class != "TeamsWebView")
        return false
    try title := WinGetTitle("ahk_id " hwnd)
    catch {
    }
    t := StrLower(title)

    ; --- Hard chrome by Win32 title ---
    if (InStr(t, "sharing control bar")
    || InStr(t, "barra de controle de compartilhamento")
    || InStr(t, "screen sharing toolbar"))
        return true
    if (InStr(t, "meeting compact view")
    || InStr(t, "modo de exibição compacto da reunião")
    || InStr(t, "modo de exibicao compacto da reuniao"))
        return true

    ; --- Eligible surfaces: never height/UIA-exclude ---
    if (InStr(t, "chat |") || InStr(t, "bate-papo |"))
        return false

    ; Normal meeting title ({name} | Microsoft Teams) without chrome prefix.
    ; Compact may look identical — only then use UIA Maximize button.
    if (InStr(t, "| microsoft teams")) {
        try {
            root := UIA.ElementFromHandle(hwnd)
            if (root) {
                uName := ""
                try uName := StrLower(root.Name)
                catch {
                }
                if (InStr(uName, "meeting compact view")
                || InStr(uName, "modo de exibição compacto da reunião")
                || InStr(uName, "modo de exibicao compacto da reuniao"))
                    return true
                if (root.FindFirst({ Name: "Maximize meeting window" })
                || root.FindFirst({ Name: "Maximizar janela da reunião" })
                || root.FindFirst({ Name: "Maximizar janela da reuniao" }))
                    return true
            }
        } catch {
        }
        return false  ; full meeting — participate in rearrange
    }

    ; Untitled / nonstandard Teams HWND: share bar is a short strip or has share UIA.
    h := 0
    try {
        rect := Buffer(16, 0)
        if DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect)
            h := NumGet(rect, 12, "int") - NumGet(rect, 4, "int")
    } catch {
        h := 0
    }
    if (h > 0 && h <= 280)
        return true
    try {
        root := UIA.ElementFromHandle(hwnd)
        if (root) {
            uName := ""
            try uName := StrLower(root.Name)
            catch {
            }
            if (InStr(uName, "sharing control bar")
            || InStr(uName, "barra de controle de compartilhamento")
            || InStr(uName, "meeting compact view")
            || InStr(uName, "modo de exibição compacto da reunião")
            || InStr(uName, "modo de exibicao compacto da reuniao"))
                return true
            if (root.FindFirst({ Name: "You're sharing your screen" })
            || root.FindFirst({ Name: "Você está compartilhando a tela" })
            || root.FindFirst({ Name: "Voce esta compartilhando a tela" }))
                return true
            if (root.FindFirst({ Name: "Maximize meeting window" })
            || root.FindFirst({ Name: "Maximizar janela da reunião" })
            || root.FindFirst({ Name: "Maximizar janela da reuniao" }))
                return true
        }
    } catch {
    }
    return false
}

; Handy, Clip Angel, Win+Shift+S clip UI, Teams chrome, and WindowManagement identity:
; skip for per-monitor cycling, move-to-monitor, tile, and auto-cursor.
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
    if (exe = "screenclippinghost.exe" || exe = "snippingtool.exe" || exe = "screensketch.exe")
        return true
    if (WM_IsTeamsChromeHwnd(hwnd))
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

WM_MaybeCenterMouse(hwnd, reason := "", force := false) {
    if (!hwnd)
        return false
    if (!force && WMAutomation_CursorCenteringSuppressed(hwnd))
        return false
    MoveMouseToCenter(hwnd)
    return true
}
