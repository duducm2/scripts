; =============================================================================
; Shift keys module: hotif_copilot_web.ahk
; M365 Copilot web Chrome hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

#HotIf IsCopilotWebChromeActiveForHotkey()

; ConsumeShiftLetter waits for release + Blind up; EndChord clears HotIf lock.

$+d:: {
    try {
        CopilotWeb_ConsumeShiftLetter("d")
        try CopilotWeb_ToggleNavDrawer()
        catch {
        }
    } finally {
        CopilotWeb_EndChord()
    }
}

$+n:: {
    try {
        CopilotWeb_ConsumeShiftLetter("n")
        try CopilotWeb_ClickNewChat()
        catch {
        }
    } finally {
        CopilotWeb_EndChord()
    }
}

$+s:: {
    try {
        CopilotWeb_ConsumeShiftLetter("s")
        try CopilotWeb_ClickNavSearch()
        catch {
        }
    } finally {
        CopilotWeb_EndChord()
    }
}

$+m:: {
    try {
        CopilotWeb_ConsumeShiftLetter("m")
        try CopilotWeb_OpenModelSelector()
        catch {
        }
    } finally {
        CopilotWeb_EndChord()
    }
}

$+a:: {
    try {
        CopilotWeb_ConsumeShiftLetter("a")
        try {
            CopilotWeb_OpenModelSelector()
            Sleep 250
            CopilotWeb_ClickSourcesCapability("capability-id-imageGeneration", ["Designer", "Criar imagem"])
        } catch {
        }
    } finally {
        CopilotWeb_EndChord()
    }
}

$+t:: {
    try {
        CopilotWeb_ConsumeShiftLetter("t")
        try {
            uia := CopilotWeb_GetActiveUia()
            if CopilotWeb_OpenSourcesMenu(uia) {
                Sleep 100
                Send "{Tab}"
            }
        } catch {
        }
    } finally {
        CopilotWeb_EndChord()
    }
}

$+i:: {
    try {
        CopilotWeb_ConsumeShiftLetter("i")
        try CopilotWeb_ClickSourcesCapability("capability-id-imageGeneration", ["Designer", "Criar imagem"])
        catch {
        }
    } finally {
        CopilotWeb_EndChord()
    }
}

$+e:: {
    try {
        CopilotWeb_ConsumeShiftLetter("e")
        try CopilotWeb_ClickSourcesCapability("capability-id-researcher", ["Researcher", "Pesquisador",
            "Pesquisa aprofundada"])
        catch {
        }
    } finally {
        CopilotWeb_EndChord()
    }
}

$+p:: {
    try {
        CopilotWeb_ConsumeShiftLetter("p")
        try CopilotWeb_FocusComposer(CopilotWeb_GetActiveUia(), false)
        catch {
        }
    } finally {
        CopilotWeb_EndChord()
    }
}

$+c:: {
    try {
        CopilotWeb_ConsumeShiftLetter("c")
        try CopilotWeb_ShiftCopyLastMessage()
        catch {
        }
    } finally {
        CopilotWeb_EndChord()
    }
}

$+r:: {
    try {
        CopilotWeb_ConsumeShiftLetter("r")
        try CopilotWeb_ShiftReadAloud()
        catch {
        }
    } finally {
        CopilotWeb_EndChord()
    }
}

$+v:: {
    try {
        CopilotWeb_ConsumeShiftLetter("v")
        try {
            if !CopilotWeb_ToggleVoiceChat()
                ShowCenteredOverlay_Utils("Voice chat control not found", 2200, BANNER_ACCENT_ERROR)
        } catch {
        }
    } finally {
        CopilotWeb_EndChord()
    }
}

$+g:: {
    try {
        CopilotWeb_ConsumeShiftLetter("g")
        try CopilotWeb_SendPromptFromFile()
        catch {
        }
    } finally {
        CopilotWeb_EndChord()
    }
}

$+f:: {
    try {
        CopilotWeb_ConsumeShiftLetter("f")
        try {
            if !CopilotWeb_ToggleComposerFullscreen()
                ShowCenteredOverlay_Utils("Fullscreen input button not found", 2200, BANNER_ACCENT_ERROR)
        } catch {
        }
    } finally {
        CopilotWeb_EndChord()
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