; =============================================================================
; Utils module: utility_shortcuts.ahk
; Utility shortcuts ^!# secondary triggers
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; Ctrl+Alt+Win+L - direct D2C submit path (paste + Enter, then monitor)
^!#L:: D2C_FlowManager.GetInstance().StartFromHotstring()

; Ctrl+Alt+Win+4 - Gemini tab 1/2 toggle + banner
^!#4::
{
    global g_GeminiToggleTab

    geminiHwnd := 0
    try {
        for hwnd in WinGetList("ahk_exe chrome.exe") {
            try {
                if InStr(WinGetTitle("ahk_id " hwnd), "gemini", false) {
                    geminiHwnd := hwnd
                    break
                }
            } catch {
            }
        }
    } catch {
    }

    if (!geminiHwnd)
        return

    WinActivate("ahk_id " geminiHwnd)
    if (!WinWaitActive("ahk_id " geminiHwnd, , 2))
        return

    Sleep(120)
    uia := UIA_Browser("ahk_id " geminiHwnd)
    tabInfo := GetChromeActiveTabIndex(uia)
    if (!tabInfo) {
        Sleep(150)
        tabInfo := GetChromeActiveTabIndex(uia)
    }
    targetTab := (tabInfo && tabInfo.index == 1) ? 2 : 1
    g_GeminiToggleTab := targetTab
    Send("^" . targetTab)
    ShowSingleCharTabBanner_Utils(targetTab)
    Sleep(200)

    tabInfoAfter := GetChromeActiveTabIndex(uia)
    if (!tabInfoAfter) {
        Sleep(100)
        tabInfoAfter := GetChromeActiveTabIndex(uia)
    }
    tabOk := tabInfoAfter && tabInfoAfter.index == targetTab
    if (!tabOk)
        ShowCenteredOverlay_Utils("❌ Shortcut execution failed", 2000, BANNER_ACCENT_ERROR)
}

; Ctrl+Alt+Win+2..8 - direct macro hotkeys (secondary triggers)
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
