; =============================================================================
; Shift keys module: hotif_copilot_web.ahk
; M365 Copilot web Chrome hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

#HotIf IsCopilotWebChromeActiveForHotkey()

; RunShiftLetterAction: suppress letter during UIA, undo stray char if composer absorbed it.

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

CopilotWeb_ShiftD(*) {
    if !CopilotWeb_ToggleNavDrawer()
        ShowCenteredOverlay_Utils("Copilot nav drawer button not found", 2000, BANNER_ACCENT_ERROR)
}

$+d:: CopilotWeb_RunShiftLetterAction("d", CopilotWeb_ShiftD)
$+n:: CopilotWeb_RunShiftLetterAction("n", CopilotWeb_ClickNewChat)
$+s:: CopilotWeb_RunShiftLetterAction("s", CopilotWeb_ClickNavSearch)
$+m:: CopilotWeb_RunShiftLetterAction("m", CopilotWeb_OpenModelSelector)
$+a:: CopilotWeb_RunShiftLetterAction("a", CopilotWeb_ShiftA)
$+t:: CopilotWeb_RunShiftLetterAction("t", CopilotWeb_ShiftT)
$+i:: CopilotWeb_RunShiftLetterAction("i", CopilotWeb_ShiftI)
$+e:: CopilotWeb_RunShiftLetterAction("e", CopilotWeb_ShiftE)
$+p:: CopilotWeb_RunShiftLetterAction("p", CopilotWeb_ShiftP)
$+c:: CopilotWeb_RunShiftLetterAction("c", CopilotWeb_ShiftCopyLastMessage)
$+r:: CopilotWeb_RunShiftLetterAction("r", CopilotWeb_ShiftReadAloud)
$+v:: CopilotWeb_RunShiftLetterAction("v", CopilotWeb_ShiftV)
$+g:: CopilotWeb_RunShiftLetterAction("g", CopilotWeb_SendPromptFromFile)
$+f:: CopilotWeb_RunShiftLetterAction("f", CopilotWeb_ShiftF)

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