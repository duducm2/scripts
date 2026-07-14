; =============================================================================
; Shift keys module: hotif_copilot_web.ahk
; M365 Copilot web Chrome hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

#HotIf IsCopilotWebChromeActiveForHotkey()

$+d:: {
    try {
        CopilotWeb_ToggleNavDrawer()
        Sleep 200
    } catch {
    }
}

$+n:: {
    try CopilotWeb_ClickNewChat()
    catch {
    }
}

$+s:: {
    try CopilotWeb_ClickNavSearch()
    catch {
    }
}

$+m:: {
    try CopilotWeb_OpenModelSelector()
    catch {
    }
}

$+a:: {
    try {
        CopilotWeb_OpenModelSelector()
        Sleep 250
        CopilotWeb_ClickAddCapability(COPILOT_CAPABILITY_IMAGE_NAMES)
    } catch {
    }
}

$+t:: {
    try {
        uia := CopilotWeb_GetActiveUia()
        if CopilotWeb_OpenSourcesMenu(uia) {
            Sleep 100
            Send "{Tab}"
        }
    } catch {
    }
}

$+i:: {
    try CopilotWeb_ClickAddCapability(COPILOT_CAPABILITY_IMAGE_NAMES)
    catch {
    }
}

$+e:: {
    try CopilotWeb_ClickAddCapability(COPILOT_CAPABILITY_RESEARCH_NAMES)
    catch {
    }
}

$+c:: {
    try CopilotWeb_ShiftCopyLastMessage()
    catch {
    }
}

$+r:: {
    try CopilotWeb_ShiftReadAloud()
    catch {
    }
}

$+v:: {
    try {
        if !CopilotWeb_ToggleVoiceChat()
            ShowCenteredOverlay_Utils("Voice chat control not found", 2200, BANNER_ACCENT_ERROR)
    } catch {
    }
}

$+f:: {
    try {
        if !CopilotWeb_ToggleComposerFullscreen()
            ShowCenteredOverlay_Utils("Fullscreen input button not found", 2200, BANNER_ACCENT_ERROR)
    } catch {
    }
}

$Enter:: {
    if (GetKeyState("Shift", "P") || GetKeyState("Ctrl", "P")) {
        Send "{Enter}"
        return
    }
    Send "{Enter}"
    SetTimer(() => CopilotWeb_WaitForGenerationComplete(300000), -1)
}

$^Enter:: {
    Send "{Enter}"
    SetTimer(() => CopilotWeb_WaitForGenerationComplete(300000), -1)
}

#HotIf