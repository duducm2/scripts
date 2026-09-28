; =============================================================================
; Shift keys module: hotif_gemini_enterprise.ahk
; Gemini Enterprise (Vertex AI Search / AskBosch) Chrome hotkeys
; Loaded via #include into the Shift keys.ahk process.
; =============================================================================

#HotIf IsGeminiEnterpriseChromeActiveForHotkey()

$*d:: {
    if !ShiftLetterHotkey_IsBareShift() {
        ShiftLetterHotkey_Relay()
        return
    }
    ShiftLetterHotkey_Consume()
    try {
        GeminiEnterprise_ToggleNavDrawer()
        GeminiEnterprise_ReturnToComposer()
    } catch {
    }
}

$*n:: {
    if !ShiftLetterHotkey_IsBareShift() {
        ShiftLetterHotkey_Relay()
        return
    }
    ShiftLetterHotkey_Consume()
    try AiCompanion_StartNewChat(AI_COMPANION_ENTERPRISE)
    catch {
    }
}

$*s:: {
    if !ShiftLetterHotkey_IsBareShift() {
        ShiftLetterHotkey_Relay()
        return
    }
    ShiftLetterHotkey_Consume()
    try GeminiEnterprise_ClickNavSearch()
    catch {
    }
}

$*m:: {
    if !ShiftLetterHotkey_IsBareShift() {
        ShiftLetterHotkey_Relay()
        return
    }
    ShiftLetterHotkey_Consume()
    try AiCompanionModels_SelectRole(AI_COMPANION_ENTERPRISE, "deep")
    catch {
    }
}

$*q:: {
    if !ShiftLetterHotkey_IsBareShift() {
        ShiftLetterHotkey_Relay()
        return
    }
    ShiftLetterHotkey_Consume()
    try AiCompanionModels_SelectRole(AI_COMPANION_ENTERPRISE, "fast")
    catch {
    }
}

$*l:: {
    if !ShiftLetterHotkey_IsBareShift() {
        ShiftLetterHotkey_Relay()
        return
    }
    ShiftLetterHotkey_Consume()
    try ShowAiCompanionModelSelector(AI_COMPANION_ENTERPRISE)
    catch {
    }
}

$*a:: {
    if !ShiftLetterHotkey_IsBareShift() {
        ShiftLetterHotkey_Relay()
        return
    }
    ShiftLetterHotkey_Consume()
    try {
        ok := GeminiEnterprise_RunWithBusyBanner(
            "⏳ 3.1 Pro + Create images + Bosch prompt… Don't move the mouse", GeminiEnterprise_ShiftArt)
        if !ok
            ShowCenteredOverlay_Utils("Create images / art prompt failed", 2200, BANNER_ACCENT_ERROR)
    } catch {
    }
}

$*t:: {
    if !ShiftLetterHotkey_IsBareShift() {
        ShiftLetterHotkey_Relay()
        return
    }
    ShiftLetterHotkey_Consume()
    try {
        uia := GeminiEnterprise_GetActiveUia()
        if GeminiEnterprise_OpenToolsMenu(uia)
            Sleep 100
    } catch {
    }
}

$*i:: {
    if !ShiftLetterHotkey_IsBareShift() {
        ShiftLetterHotkey_Relay()
        return
    }
    ShiftLetterHotkey_Consume()
    try {
        ok := GeminiEnterprise_RunWithBusyBanner("⏳ Create images… Don't move the mouse", (*) =>
            GeminiEnterprise_ClickCreateImages())
        if (ok)
            GeminiEnterprise_ReturnToComposer()
        else
            ShowCenteredOverlay_Utils("Create images failed", 2200, BANNER_ACCENT_ERROR)
    } catch {
    }
}

$*e:: {
    if !ShiftLetterHotkey_IsBareShift() {
        ShiftLetterHotkey_Relay()
        return
    }
    ShiftLetterHotkey_Consume()
    try {
        ok := GeminiEnterprise_RunWithBusyBanner("⏳ Deep Research… Don't move the mouse", (*) =>
            GeminiEnterprise_ClickDeepResearch())
        if (ok)
            GeminiEnterprise_ReturnToComposer()
        else
            ShowCenteredOverlay_Utils("Deep Research failed", 2200, BANNER_ACCENT_ERROR)
    } catch {
    }
}

$*p:: {
    if !ShiftLetterHotkey_IsBareShift() {
        ShiftLetterHotkey_Relay()
        return
    }
    ShiftLetterHotkey_Consume()
    try {
        uia := GeminiEnterprise_GetActiveUia()
        if !GeminiEnterprise_FocusComposer(uia, true)
            ShowCenteredOverlay_Utils("Prompt field not found", 2200, BANNER_ACCENT_ERROR)
    } catch {
    }
}

; Shift+H: strip human reminders after last --- (keep divider + blank lines)
$*h:: {
    if !ShiftLetterHotkey_IsBareShift() {
        ShiftLetterHotkey_Relay()
        return
    }
    ShiftLetterHotkey_Consume()
    try {
        reason := ""
        if GeminiEnterprise_StripComposerHumanReminders(&reason)
            return
        if (reason = "empty")
            ShowCenteredOverlay_Utils("Could not read prompt text", 2200, BANNER_ACCENT_ERROR)
        else
            ShowCenteredOverlay_Utils("No --- human-reminder divider found", 2200, BANNER_ACCENT_ERROR)
    } catch {
    }
}

$Enter:: {
    if (GetKeyState("Shift", "P") || GetKeyState("Ctrl", "P")) {
        Send "{Enter}"
        return
    }
    try {
        uia := GeminiEnterprise_GetActiveUia()
        GeminiEnterprise_TrySubmit(uia)
    } catch {
        Send "{Enter}"
    }
    SetTimer(() => GeminiEnterprise_WaitForGenerationComplete(300000), -1)
}

$^Enter:: {
    try {
        uia := GeminiEnterprise_GetActiveUia()
        GeminiEnterprise_TrySubmit(uia)
    } catch {
        Send "{Enter}"
    }
    SetTimer(() => GeminiEnterprise_WaitForGenerationComplete(300000), -1)
}

#HotIf