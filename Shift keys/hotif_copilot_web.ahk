; =============================================================================
; Shift keys module: hotif_copilot_web.ahk
; M365 Copilot web Chrome hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

#HotIf IsCopilotWebChromeActiveForHotkey()

; BeginChord + KeyWait keep the letter suppressed until release (composer refocus safe).

CopilotWeb_ShiftChord(letter, actionCallback) {
    CopilotWeb_BeginChord()
    try {
        if (actionCallback is Func)
            actionCallback.Call()
        else
            actionCallback()
        KeyWait letter
    } finally {
        CopilotWeb_EndChord()
    }
}

CopilotWeb_ShiftA(*) {
    CopilotWeb_OpenModelSelector()
    Sleep 250
    CopilotWeb_ClickSourcesCapability("capability-id-imageGeneration", ["Designer", "Criar imagem"])
}

CopilotWeb_ShiftT(*) {
    uia := CopilotWeb_GetActiveUia()
    if CopilotWeb_OpenSourcesMenu(uia) {
        Sleep 100
        Send "{Tab}"
    }
}

CopilotWeb_ShiftI(*) {
    CopilotWeb_ClickSourcesCapability("capability-id-imageGeneration", ["Designer", "Criar imagem"])
}

CopilotWeb_ShiftE(*) {
    CopilotWeb_ClickSourcesCapability("capability-id-researcher", ["Researcher", "Pesquisador",
        "Pesquisa aprofundada"])
}

CopilotWeb_ShiftP(*) {
    CopilotWeb_FocusComposer(CopilotWeb_GetActiveUia(), false)
}

CopilotWeb_ShiftV(*) {
    if !CopilotWeb_ToggleVoiceChat()
        ShowCenteredOverlay_Utils("Voice chat control not found", 2200, BANNER_ACCENT_ERROR)
}

CopilotWeb_ShiftF(*) {
    if !CopilotWeb_ToggleComposerFullscreen()
        ShowCenteredOverlay_Utils("Fullscreen input button not found", 2200, BANNER_ACCENT_ERROR)
}

$+d:: CopilotWeb_ShiftChord("d", CopilotWeb_ToggleNavDrawer)
$+n:: CopilotWeb_ShiftChord("n", CopilotWeb_ClickNewChat)
$+s:: CopilotWeb_ShiftChord("s", CopilotWeb_ClickNavSearch)
$+m:: CopilotWeb_ShiftChord("m", CopilotWeb_OpenModelSelector)
$+a:: CopilotWeb_ShiftChord("a", CopilotWeb_ShiftA)
$+t:: CopilotWeb_ShiftChord("t", CopilotWeb_ShiftT)
$+i:: CopilotWeb_ShiftChord("i", CopilotWeb_ShiftI)
$+e:: CopilotWeb_ShiftChord("e", CopilotWeb_ShiftE)
$+p:: CopilotWeb_ShiftChord("p", CopilotWeb_ShiftP)
$+c:: CopilotWeb_ShiftChord("c", CopilotWeb_ShiftCopyLastMessage)
$+r:: CopilotWeb_ShiftChord("r", CopilotWeb_ShiftReadAloud)
$+v:: CopilotWeb_ShiftChord("v", CopilotWeb_ShiftV)
$+g:: CopilotWeb_ShiftChord("g", CopilotWeb_SendPromptFromFile)
$+f:: CopilotWeb_ShiftChord("f", CopilotWeb_ShiftF)

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