; =============================================================================
; Utils module: utility_shortcuts.ahk
; Utility shortcuts #!+U and ^!# secondary triggers
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

; Ctrl+Alt+Win+L - direct D2C submit path (paste + Enter, then monitor)
^!#L:: D2C_FlowManager.GetInstance().StartFromHotstring()

; Ctrl+Alt+Win+4 - AI Text Optimizer (same as Win+Alt+Shift+U then L, 4)
^!#4:: GeminiNavigateFocusAndPasteFirstSnippet(GetAioptPromptText(), true)

; Ctrl+Alt+Win+2..8 - same macros as HotStrings panel (Win+Alt+Shift+U); secondary triggers only
^!#2:: QuickUpdateScripts()
^!#3:: ToggleOutlookAndTeams()
^!#5:: CleanClipboard()
; Same macro; InputLevel 10 + hook so chord wins over other low-level handlers (optional ghosting fallback: ^!#j).
#InputLevel 10
#UseHook
^!#7:: MarkLastClipAsFavorite()
^!#j:: MarkLastClipAsFavorite()
#UseHook False
#InputLevel 0
^!#8:: DesktopToRecycle_Trigger()
; Ctrl+Alt+Win+9 / +B - Handy Cohere Portuguese / English (g_HandyAiModels slots 4 and 3)
^!#9:: ExecuteHandyAiModelSelection(HANDY_AI_SLOT_COHERE_PORTUGUESE)
^!#b:: ExecuteHandyAiModelSelection(HANDY_AI_SLOT_COHERE_ENGLISH)

; =============================================================================
; Alt+Shift+W Shortcut
; Hotkey: Alt+Shift+W
; Sends Alt+Shift+W again, then shows a message box
; =============================================================================
!+W::
{
    ; Send Alt+Shift+W again
    ; Small delay to ensure previous key release is complete
    Sleep 50

    ; Send Win+Ctrl+Alt+Y using SendInput for better reliability
    ; SendInput is more reliable for complex modifier combinations
    SendInput "#^!y"

    ; Show message box

    Sleep 50

    Send '""'

    Sleep 50

    Send "{Left}"

}
