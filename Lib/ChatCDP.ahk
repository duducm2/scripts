; =============================================================================
; Lib: ChatCDP.ahk
; Background CDP client for WhatsApp Web and Microsoft Teams Web.
; Reuses the persistent CDP bridge (127.0.0.1:8768) from Lib\AiCompanionCDP.ahk.
; =============================================================================

ChatCDP_IniPath() {
    return A_ScriptDir . "\assets\data\chat_cdp_mode.ini"
}

ChatCDP_IsEnabled(app := "") {
    if (AiCompanionMode_Get() != "cdp")
        return false
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

ChatCDP_Jump(app, contact, text := "", files := "", submit := false, activate := false) {
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
        "activate", activate ? true : false
    ), 30000)
    hwnd := _AiCDP_JsonGetInt(res.raw, "hwnd", 0)
    if (res.ok && activate)
        hwnd := AiCompanionCDP_ActivateWindow(hwnd, appNorm)
    return { ok: res.ok, hwnd: hwnd, error: res.error }
}

ChatCDP_NotifyError(app, rawErr := "") {
    label := (StrLower(Trim(app)) = "teams") ? "Teams" : "WhatsApp"
    err := Trim(rawErr)
    low := StrLower(err)
    if (InStr(low, "auth_required"))
        msg := "❌ " . label . " (CDP): Sign in required in CDP Chrome — press Win+Alt+Shift+I or open " . label . " Web in CDP Chrome"
    else if (InStr(low, "contact_not_found:"))
        msg := "❌ " . label . " (CDP): Contact not found — " . Trim(SubStr(err, 19))
    else
        msg := "❌ " . label . " (CDP): " . (err != "" ? err : "Failed")
    try ShowCenteredOverlay_Utils(msg, 2600, BANNER_ACCENT_ERROR)
    catch {
    }
}
