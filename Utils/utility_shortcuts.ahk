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

; Win+Alt+Shift+L — paste OS clipboard (^v) to a picked visible window (same as D2C menu [W]).
; If a main text field is saved for that exe+title (assets/data/paste_field_mappings.ini),
; focus it via UIA before paste; if unknown, after paste ask [Y]/[N] to persist the focused field.
; In the picker: slot key = paste; [R] then slot = ignore that process (exe) for AutoSlot;
; [I] = manage/remove ignore entries (assets/data/autoslot_user_excludes.ini);
; [M] = manage/remove main text-field mappings (assets/data/paste_field_mappings.ini).
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

; Ctrl+Alt+Win+7 - toggle Chrome tab 1 <-> 2 on the resolved AI companion
; (Enterprise / Copilot / consumer Gemini via ResolveGlobalAICompanion).
^!#7:: ToggleAICompanionChromeTab()

ToggleAICompanionChromeTab() {
    global g_GeminiToggleTab
    companion := ResolveGlobalAICompanion()
    hwnd := 0
    label := "Gemini"
    switch companion {
        case "enterprise":
            hwnd := GetGeminiEnterpriseWindowHwnd()
            label := "Gemini Enterprise"
        case "copilot":
            hwnd := GetCopilotWebWindowHwnd()
            label := "Copilot"
        default:
            hwnd := FindGeminiChromeHwnd()
            label := "Gemini"
    }
    if (!hwnd) {
        ShowCenteredOverlay_Utils("❌ " . label . " is not open.", 1800, BANNER_ACCENT_ERROR)
        return
    }
    WinActivate("ahk_id " hwnd)
    if !WinWaitActive("ahk_id " hwnd, , 2) {
        ShowCenteredOverlay_Utils("❌ Could not activate " . label . ".", 1800, BANNER_ACCENT_ERROR)
        return
    }

    ; Prefer UIA: if on tab 1 go to 2, otherwise go to 1. Fall back to remembered flip.
    targetTab := (g_GeminiToggleTab = 1) ? 2 : 1
    try {
        uia := UIA_Browser("ahk_id " hwnd)
        tabInfo := GetChromeActiveTabIndex(uia)
        if (tabInfo && tabInfo.index)
            targetTab := (tabInfo.index = 1) ? 2 : 1
    } catch {
    }

    Send("^" . targetTab)
    Sleep 120
    g_GeminiToggleTab := targetTab
    ShowSingleCharTabBanner_Utils(targetTab)
}

; Ctrl+Alt+Win+2..8 - same macros as HotStrings panel (Win+Alt+Shift+U); secondary triggers only
^!#2:: QuickUpdateScripts()
^!#3:: ToggleOutlookAndTeams()
; InputLevel 10 + hook so ZMK / firmware chords win over other low-level handlers.
#InputLevel 10
#UseHook
^!#5:: CleanClipboard()
^!#4:: MarkLastClipAsFavorite()
^!#j:: MarkLastClipAsFavorite()
#UseHook False
#InputLevel 0
^!#8:: DesktopToRecycle_Trigger()
; Ctrl+Alt+Win+9 / +B - Handy Nemotron Portuguese / Parakeet Unified English (slots 2 and 1)
^!#9:: ExecuteHandyAiModelSelection(HANDY_AI_SLOT_PORTUGUESE)
^!#b:: ExecuteHandyAiModelSelection(HANDY_AI_SLOT_ENGLISH)

#^!m::
{
    ; Small delay to ensure previous key release is complete
    Sleep 50

    ; Send Win+Ctrl+Alt+M using SendInput for better reliability
    SendInput "#^!m"

    ; Show message box

    Sleep 50

    Send '""'

    Sleep 50

    Send "{Left}"

}
