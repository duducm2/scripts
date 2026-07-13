; =============================================================================
; Shift keys module: hotif_copilot_web.ahk
; M365 Copilot web Chrome hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

#HotIf IsCopilotWebChromeActiveForHotkey()

; KeyWait before UIA/Send so a held Shift+letter cannot leak into the composer.

$+d:: {
    KeyWait "d", "T1"
    try CopilotWeb_ToggleNavDrawer()
    catch {
    }
}

$+n:: {
    KeyWait "n", "T1"
    try CopilotWeb_ClickNewChat()
    catch {
    }
}

$+s:: {
    KeyWait "s", "T1"
    try CopilotWeb_ClickNavSearch()
    catch {
    }
}

$+m:: {
    KeyWait "m", "T1"
    try CopilotWeb_OpenModelSelector()
    catch {
    }
}

$+a:: {
    KeyWait "a", "T1"
    try {
        CopilotWeb_OpenModelSelector()
        Sleep 250
        CopilotWeb_ClickSourcesCapability("capability-id-imageGeneration", ["Designer", "Criar imagem"])
    } catch {
    }
}

$+t:: {
    KeyWait "t", "T1"
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
    KeyWait "i", "T1"
    try CopilotWeb_ClickSourcesCapability("capability-id-imageGeneration", ["Designer", "Criar imagem"])
    catch {
    }
}

$+e:: {
    KeyWait "e", "T1"
    try CopilotWeb_ClickSourcesCapability("capability-id-researcher", ["Researcher", "Pesquisador",
        "Pesquisa aprofundada"])
    catch {
    }
}

$+p:: {
    KeyWait "p", "T1"
    try CopilotWeb_FocusComposer(CopilotWeb_GetActiveUia(), false)
    catch {
    }
}

$+c:: {
    KeyWait "c", "T1"
    try CopilotWeb_ShiftCopyLastMessage()
    catch {
    }
}

$+r:: {
    KeyWait "r", "T1"
    try CopilotWeb_ShiftReadAloud()
    catch {
    }
}

$+v:: {
    KeyWait "v", "T1"
    try {
        if !CopilotWeb_ToggleVoiceChat()
            ShowCenteredOverlay_Utils("Voice chat control not found", 2200, BANNER_ACCENT_ERROR)
    } catch {
    }
}

$+g:: {
    KeyWait "g", "T1"
    try CopilotWeb_SendPromptFromFile()
    catch {
    }
}

$+f:: {
    KeyWait "f", "T1"
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