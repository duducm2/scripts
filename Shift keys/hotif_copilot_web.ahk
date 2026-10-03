; =============================================================================
; Shift keys module: hotif_copilot_web.ahk
; M365 Copilot web Chrome hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

#HotIf IsCopilotWebChromeActiveForHotkey()

+d:: {
    try {
        CopilotWeb_ToggleNavDrawer()
        CopilotWeb_ReturnToComposer()
    } catch {
    }
}

+n:: {
    try AiCompanion_StartNewChat(AI_COMPANION_COPILOT)
    catch {
    }
}

+s:: {
    try CopilotWeb_ClickNavSearch()
    catch {
    }
}

+m:: {
    try AiCompanionModels_SelectRole(AI_COMPANION_COPILOT, "deep")
    catch {
    }
}

+q:: {
    try AiCompanionModels_SelectRole(AI_COMPANION_COPILOT, "fast")
    catch {
    }
}

+l:: {
    try ShowAiCompanionModelSelector(AI_COMPANION_COPILOT)
    catch {
    }
}

+a:: {
    try {
        ok := CopilotWeb_RunWithBusyBanner(
            "⏳ Think deeper + Generate image + Bosch prompt… Don't move the mouse", CopilotWeb_ShiftArt)
        if !ok
            ShowCenteredOverlay_Utils("Generate an image failed", 2200, BANNER_ACCENT_ERROR)
    } catch {
    }
}

+t:: {
    try {
        uia := CopilotWeb_GetActiveUia()
        if CopilotWeb_OpenSourcesMenu(uia) {
            Sleep 100
            Send "{Tab}"
        }
    } catch {
    }
}

+i:: {
    try {
        ok := CopilotWeb_RunWithBusyBanner("⏳ Generate an image… Don't move the mouse", (*) =>
            CopilotWeb_ClickAddCapability(
                COPILOT_CAPABILITY_IMAGE_NAMES))
        if (ok)
            CopilotWeb_ReturnToComposer()
        else
            ShowCenteredOverlay_Utils("Generate an image failed", 2200, BANNER_ACCENT_ERROR)
    } catch {
    }
}

+e:: {
    try {
        ok := CopilotWeb_RunWithBusyBanner("⏳ Research a topic… Don't move the mouse", (*) =>
            CopilotWeb_ClickAddCapability(
                COPILOT_CAPABILITY_RESEARCH_NAMES))
        if (ok)
            CopilotWeb_ReturnToComposer()
        else
            ShowCenteredOverlay_Utils("Research a topic failed", 2200, BANNER_ACCENT_ERROR)
    } catch {
    }
}

+c:: {
    try {
        CopilotWeb_ShiftCopyLastMessage()
        CopilotWeb_ReturnToComposer()
    } catch {
    }
}

+r:: {
    try {
        CopilotWeb_ShiftReadAloud()
        CopilotWeb_ReturnToComposer()
    } catch {
    }
}

+v:: {
    try {
        if !CopilotWeb_ToggleVoiceChat()
            ShowCenteredOverlay_Utils("Voice chat control not found", 2200, BANNER_ACCENT_ERROR)
    } catch {
    }
}

+f:: {
    try {
        if CopilotWeb_ToggleComposerFullscreen()
            CopilotWeb_ReturnToComposer()
        else
            ShowCenteredOverlay_Utils("Fullscreen input button not found", 2200, BANNER_ACCENT_ERROR)
    } catch {
    }
}

; Shift+H: strip human reminders after last --- (keep divider + blank lines for comments)
+h:: {
    try {
        if !CopilotWeb_StripComposerHumanReminders()
            ShowCenteredOverlay_Utils("No --- human-reminder divider found", 2200, BANNER_ACCENT_ERROR)
    } catch {
    }
}

$Enter:: {
    if (GetKeyState("Shift", "P") || GetKeyState("Ctrl", "P")) {
        Send "{Enter}"
        return
    }
    hwnd := WinExist("A")
    global AI_COMPANION_ENTER_SEND_FIRST
    if (AI_COMPANION_ENTER_SEND_FIRST) {
        AiCompanion_SendEnterFirst(hwnd, "copilot", false)
        return
    }
    if (AiCompanion_SendAndConfirm(hwnd, "copilot", CopilotWeb_HotkeySendFn(hwnd))) {
        if (!AiCompanion_FinishConfirmedSubmit(hwnd, "copilot", hwnd))
            SetTimer(() => CopilotWeb_WaitForGenerationComplete(300000), -1)
    } else {
        AiCompanion_TakeCatalogSendClick()
    }
}

$^Enter:: {
    hwnd := WinExist("A")
    global AI_COMPANION_ENTER_SEND_FIRST
    if (AI_COMPANION_ENTER_SEND_FIRST) {
        AiCompanion_SendEnterFirst(hwnd, "copilot", true)
        return
    }
    if (AiCompanion_SendAndConfirm(hwnd, "copilot", CopilotWeb_HotkeySendFn(hwnd))) {
        if (!AiCompanion_FinishConfirmedSubmit(hwnd, "copilot", hwnd))
            SetTimer(() => CopilotWeb_WaitForGenerationComplete(300000), -1)
    } else {
        AiCompanion_TakeCatalogSendClick()
    }
}

#HotIf