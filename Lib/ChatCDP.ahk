; =============================================================================
; Lib: ChatCDP.ahk
; Background CDP client for WhatsApp Web and Microsoft Teams Web.
; Reuses the persistent CDP bridge (127.0.0.1:8768) from Lib\AiCompanionCDP.ahk.
; =============================================================================

ChatCDP_IniPath() {
    return _AiCDP_RepoRoot() . "\assets\data\chat_cdp_mode.ini"
}

ChatCDP_IsEnabled(app := "") {
    appNorm := StrLower(Trim(app))
    ; Teams uses the native desktop app (ms-teams.exe) by default, not teams.microsoft.com in CDP browser.
    defaultVal := (appNorm = "teams") ? "off" : "cdp"
    path := ChatCDP_IniPath()
    if FileExist(path) {
        try {
            val := StrLower(Trim(IniRead(path, "Mode", appNorm != "" ? appNorm : "Enabled", defaultVal)))
            if (val = "off" || val = "0" || val = "ahk" || val = "false")
                return false
            return true
        } catch {
        }
    }
    return defaultVal != "off"
}

ChatCDP_AppMeta(appNorm) {
    if (appNorm = "teams")
        return { url: "https://teams.microsoft.com", needle: "teams.microsoft.com" }
    return { url: "https://web.whatsapp.com", needle: "web.whatsapp.com" }
}

ChatCDP_IsAppTabOpen(app := "whatsapp", timeoutMs := 600) {
    global AI_COMPANION_CDP_CHROME_PORT
    appNorm := StrLower(Trim(app))
    if (appNorm != "whatsapp" && appNorm != "teams")
        return false
    if !AiCompanionCDP_IsChromeCdpUp(timeoutMs)
        return false
    meta := ChatCDP_AppMeta(appNorm)
    needleLow := StrLower(meta.needle)
    try {
        whr := ComObject("WinHttp.WinHttpRequest.5.1")
        try whr.SetProxy(1)
        whr.Open("GET", "http://127.0.0.1:" . AI_COMPANION_CDP_CHROME_PORT . "/json/list", false)
        whr.SetTimeouts(timeoutMs, timeoutMs, 1500, 1500)
        whr.Send()
        if (whr.Status != 200)
            return false
        raw := whr.ResponseText
        for obj in _AiCDP_JsonArrayObjects(raw) {
            if (_AiCDP_JsonGetStr(obj, "type", "") != "page")
                continue
            tabUrl := StrLower(_AiCDP_JsonGetStr(obj, "url", ""))
            if InStr(tabUrl, needleLow)
                return true
        }
    } catch {
    }
    return false
}

global g_ChatCDP_LastEnsureAsyncTick := Map()

ChatCDP_EnsureAppOpenAsync(app := "whatsapp") {
    global g_ChatCDP_LastEnsureAsyncTick
    appNorm := StrLower(Trim(app))
    if (appNorm != "whatsapp" && appNorm != "teams")
        return false
    if ChatCDP_IsAppTabOpen(appNorm)
        return true
    now := A_TickCount
    if (g_ChatCDP_LastEnsureAsyncTick.Has(appNorm) && (now - g_ChatCDP_LastEnsureAsyncTick[appNorm]) < 8000)
        return true
    meta := ChatCDP_AppMeta(appNorm)
    if AiCompanionCDP_IsChromeCdpUp(900) {
        direct := AiCompanionCDP_DirectOpenUrl(meta.url, meta.needle, false, appNorm)
        if (direct.ok)
            g_ChatCDP_LastEnsureAsyncTick[appNorm] := now
        return direct.ok
    }
    ; If GeminiCDP Chrome is still starting up and port 9222 isn't listening yet,
    ; do not spawn a duplicate window and do not lock out retries for 8 seconds.
    if (AiCompanionCDP_GetProcessPids().Count > 0)
        return false
    g_ChatCDP_LastEnsureAsyncTick[appNorm] := now
    if (appNorm = "whatsapp") {
        launcher := _AiCDP_RepoRoot() . "\AppLaunchers\launch_whatsapp_cdp.ahk"
        if FileExist(launcher) {
            try {
                Run('"' . A_AhkPath . '" "' . launcher . '"')
                return true
            } catch {
            }
        }
    }
    launched := AiCompanionCDP_DirectLaunchChrome(meta.url, true, 500, appNorm, meta.needle)
    return launched.ok
}

ChatCDP_OpenOrFocus(app := "whatsapp", activate := true) {
    appNorm := StrLower(Trim(app))
    if (appNorm != "whatsapp" && appNorm != "teams")
        return { ok: false, hwnd: 0, error: "Unsupported chat app" }

    ; Layer 1: Python CDP bridge (/ensure) with built-in retry & self-heal.
    res := AiCompanionCDP_Post("/ensure", Map(
        "companion", appNorm,
        "activate", activate ? true : false,
        "new_tab", false,
        "launch_if_down", true
    ), 60000, true)
    hwnd := _AiCDP_JsonGetInt(res.raw, "hwnd", 0)
    if (res.ok) {
        if (activate)
            hwnd := AiCompanionCDP_ActivateWindow(hwnd, appNorm)
        return { ok: true, hwnd: hwnd, error: "" }
    }

    ; Layer 2: Silent direct CDP HTTP fallback on port 9222 (reuses existing Gemini CDP window).
    meta := ChatCDP_AppMeta(appNorm)
    if AiCompanionCDP_IsChromeCdpUp(800) {
        direct := AiCompanionCDP_DirectOpenUrl(meta.url, meta.needle, activate, appNorm)
        if (direct.ok)
            return direct
    }

    ; Layer 3: Silent direct CDP Chrome cold-launch fallback.
    launched := AiCompanionCDP_DirectLaunchChrome(meta.url, !activate, 15000, appNorm, meta.needle)
    if (launched.ok)
        return launched

    return { ok: false, hwnd: hwnd, error: res.error }
}

ChatCDP_Jump(app, contact, text := "", files := "", submit := false, activate := false, readyTimeoutSec := 90) {
    appNorm := StrLower(Trim(app))
    if (appNorm != "whatsapp" && appNorm != "teams")
        return { ok: false, hwnd: 0, error: "Unsupported chat app" }
    fileArr := IsObject(files) ? files : []
    res := AiCompanionCDP_Post("/chat/jump", Map(
        "app", appNorm,
        "contact", Trim(contact),
        "text", text,
        "files", fileArr,
        "submit", submit ? true : false,
        "activate", activate ? true : false,
        "launch_if_down", true,
        "ready_timeout_s", readyTimeoutSec
    ), 95000, true)
    hwnd := _AiCDP_JsonGetInt(res.raw, "hwnd", 0)
    if (res.ok && activate)
        hwnd := AiCompanionCDP_ActivateWindow(hwnd, appNorm)
    return { ok: res.ok, hwnd: hwnd, error: res.error }
}

ChatCDP_WhatsAppReady(timeoutMs := 3000) {
    res := AiCompanionCDP_Post("/whatsapp/ready", Map("launch_if_down", false), timeoutMs, false)
    if (!res.ok) {
        return {
            ok: false,
            ready: false,
            authRequired: false,
            stage: (res.error = "CDP_NOT_RUNNING") ? "cdp_down" : "error",
            error: res.error
        }
    }
    return {
        ok: true,
        ready: _AiCDP_JsonGetBool(res.raw, "ready", false),
        authRequired: _AiCDP_JsonGetBool(res.raw, "auth_required", false),
        stage: _AiCDP_JsonGetStr(res.raw, "stage", ""),
        error: ""
    }
}

ChatCDP_NotifyError(app, rawErr := "") {
    label := (StrLower(Trim(app)) = "teams") ? "Teams" : "WhatsApp"
    err := Trim(rawErr)
    low := StrLower(err)
    if (InStr(low, "auth_required"))
        msg := "❌ " . label . " (CDP): Sign in required in CDP Chrome — press Win+Alt+Shift+I or open " . label . " Web in CDP Chrome"
    else if (InStr(low, "cdp_not_running") || InStr(low, "cdp_unreachable") || InStr(low, "cdp_launch_failed") || InStr(low, "econnrefused"))
        msg := "❌ " . label . " (CDP): Could not open CDP Chrome — press Win+Alt+Shift+I to open it"
    else if (InStr(low, "0x80072efd") || InStr(low, "winhttp") || InStr(low, "conexão com o servidor") || InStr(low, "connection with the server") || InStr(low, "cdp server unavailable"))
        msg := "❌ " . label . " (CDP): Local CDP bridge unreachable after retries — try Win+Alt+Shift+I"
    else if (InStr(low, "contact_not_found:"))
        msg := "❌ " . label . " (CDP): Contact not found — " . Trim(SubStr(err, 19))
    else
        msg := "❌ " . label . " (CDP): " . (err != "" ? err : "Failed")
    try ShowCenteredOverlay_Utils(msg, 2600, BANNER_ACCENT_ERROR)
    catch {
    }
}

ChatCDP_WhatsAppVoiceToggle() {
    res := AiCompanionCDP_Post("/whatsapp/voice", Map("launch_if_down", false), 8000, false)
    state := _AiCDP_JsonGetStr(res.raw, "state", "")
    return { ok: res.ok, state: state, error: res.error }
}
