; =============================================================================
; AppLaunchers/launch_whatsapp_cdp.ahk
; Detached one-shot launcher for #!+z (WhatsApp CDP) so AppLaunchers.ahk never
; blocks on HTTP/CDP retries and runs in parallel without stealing focus.
; =============================================================================
#Requires AutoHotkey v2.0+
#SingleInstance Force

#Include %A_ScriptDir%\..\env.ahk
#Include %A_ScriptDir%\..\Utils\standard_loading_bar.ahk

Utils_EnsureGlobalEscapeHotkey() {
}

ScriptSoundPlay(path, wait := false) {
    return false
}

ShowCenteredOverlay_Utils(text, duration := 1500, bgColor := BANNER_ACCENT_INTERMEDIATE) {
    if (duration < 1)
        duration := 1
    StandardLoadingBar_Show(text, bgColor, { passive: true, noActivate: true, centerOnHwnd: 0, textWidth: 500,
        fontSize: 17, passiveBgColor: bgColor, maxDurationMs: duration + 2000 })
    StandardLoadingBar_Hide(duration)
}

ResolveGlobalAICompanion() {
    global IS_WORK_ENVIRONMENT
    return (IsSet(IS_WORK_ENVIRONMENT) && IS_WORK_ENVIRONMENT) ? "enterprise" : "gemini"
}

#Include %A_ScriptDir%\..\Lib\AiCompanionCDP.ahk
#Include %A_ScriptDir%\..\Lib\ChatCDP.ahk

StandardLoadingBar_Show("⏳ Opening WhatsApp (CDP)…", BANNER_ACCENT_INTERMEDIATE, { passive: true, noActivate: true, centerOnHwnd: 0 })
res := ChatCDP_OpenOrFocus("whatsapp", false)
StandardLoadingBar_Hide(0)
if (!res.ok) {
    ChatCDP_NotifyError("whatsapp", res.error)
    Sleep 2800
    ExitApp
}
ShowCenteredOverlay_Utils("✅ WhatsApp ready (CDP)", 1200, BANNER_ACCENT_SUCCESS)
Sleep 1400
ExitApp
