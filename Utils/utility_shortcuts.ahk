; =============================================================================
; Utils module: utility_shortcuts.ahk
; Utility shortcuts #!+U, #!+L and ^!# secondary triggers
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; Hotkey Handler: Windows + Alt + Shift + U (#!+U)
; =============================================================================
#!+U::
{
    global g_HotstringSelectorActive, g_HotstringSelectorGui

    if (g_HotstringSelectorActive && IsObject(g_HotstringSelectorGui)) {
        CleanupHotstringSelector()
    } else {
        ShowHotstringSelector()
    }
}

; Win+Alt+Shift+L — paste current clipboard to a picked visible window (same as D2C menu [W])
#!+l:: {
    mgr := D2C_FlowManager.GetInstance()
    if (mgr.CurrentPhase = "PromptingSubmit") {
        mgr.OnSubmitW()
        return
    }
    mgr.PasteClipboardToVisibleWindow()
}

; Ctrl+Alt+Win+L - direct D2C submit path (paste + Enter, then monitor)
^!#L:: D2C_FlowManager.GetInstance().StartFromHotstring()

; Ctrl+Alt+Win+4 - AI Text Optimizer (same as Win+Alt+Shift+U then L, 4)
; Routes via ResolveGlobalAICompanion (Enterprise / Copilot / consumer Gemini).
^!#4:: {
    prompt := GetAioptPromptText()
    companion := ResolveGlobalAICompanion()
    if (companion = "enterprise")
        GeminiEnterprise_NavigateFocusAndPaste(prompt, true)
    else if (companion = "copilot")
        CopilotWeb_NavigateFocusAndPaste(prompt, true)
    else
        GeminiNavigateFocusAndPasteFirstSnippet(prompt, true)
}

; Ctrl+Alt+Win+2..8 - same macros as HotStrings panel (Win+Alt+Shift+U); secondary triggers only
^!#2:: QuickUpdateScripts()
^!#3:: ToggleOutlookAndTeams()
; InputLevel 10 + hook so ZMK / firmware chords win over other low-level handlers.
#InputLevel 10
#UseHook
^!#5:: CleanClipboard()
^!#7:: MarkLastClipAsFavorite()
^!#j:: MarkLastClipAsFavorite()
#UseHook False
#InputLevel 0
^!#8:: DesktopToRecycle_Trigger()
; Ctrl+Alt+Win+9 / +B - Handy Nemotron Portuguese / Parakeet Unified English (slots 2 and 1)
^!#9:: ExecuteHandyAiModelSelection(HANDY_AI_SLOT_PORTUGUESE)
^!#b:: ExecuteHandyAiModelSelection(HANDY_AI_SLOT_ENGLISH)

!+W::
{
    Sleep 50
    ; Send Win+Ctrl+Alt+Y using SendInput for better reliability
    ; SendInput is more reliable for complex modifier combinations
    SendInput "#^!y"

}

#^!m::
{
    ; Send Alt+Shift+W again
    ; Small delay to ensure previous key release is complete
    Sleep 50

    ; Send Win+Ctrl+Alt+Y using SendInput for better reliability
    ; SendInput is more reliable for complex modifier combinations
    SendInput "#^!m"

    ; Show message box

    Sleep 50

    Send '""'

    Sleep 50

    Send "{Left}"

}
