; =============================================================================
; Utils module: gemini_paste_helpers.ahk
; Gemini navigate/focus/paste helpers used by D2C and legacy dictation flows.
; =============================================================================

global g_HotstringGeminiAutoSubmit := true
global g_HotstringGeminiRestoreHwnd := 0
global g_GeminiDelayedSubmit_PreEnterDelayMs := 1000
global g_GeminiDelayedSubmit_WaitContentMaxMs := 5000

GeminiNavigateFocusAndPasteFirstSnippet(optionalPromptText := "", switchToFirstTab := true) {
    SetTitleMatchMode(2)
    geminiHwnd := 0
    if (!switchToFirstTab) {
        try {
            activeHwnd := WinExist("A")
            if (activeHwnd && WinGetProcessName("ahk_id " activeHwnd) = "chrome.exe" && InStr(WinGetTitle("ahk_id " activeHwnd
            ), "gemini", false))
                geminiHwnd := activeHwnd
        } catch {
        }
    }
    if (!geminiHwnd) {
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
    }

    if (!geminiHwnd) {
        Run "chrome.exe --new-window https://gemini.google.com/"
        if !WinWaitActive("ahk_exe chrome.exe", , 5)
            return
        Sleep 2500
        geminiHwnd := WinExist("A")
    }

    if (geminiHwnd) {
        WinActivate("ahk_id " geminiHwnd)
        WinWaitActive("ahk_id " geminiHwnd, , 2)
        Sleep 350
    } else {
        WinActivate("ahk_exe chrome.exe")
        WinWaitActive("ahk_exe chrome.exe", , 2)
        Sleep 350
    }

    if (switchToFirstTab) {
        Send("^1")
        Sleep 280
        ShowSingleCharTabBanner_Utils(1)
    }

    try {
        uia := UIA_Browser()
        Sleep 80
        anchorButton := 0
        try {
            anchorButton := uia.FindFirst({ Type: "50000", Name: "Open upload file menu", ControlType: "Button" })
            if (!anchorButton)
                anchorButton := uia.FindFirst({ Type: "50000", Name: "Open upload file menu", cs: false })
        } catch {
        }
        if (!anchorButton) {
            try {
                allButtons := uia.FindAll({ Type: "50000" })
                for button in allButtons {
                    try {
                        if (InStr(button.Name, "Open upload file menu", false)) {
                            anchorButton := button
                            break
                        }
                    } catch {
                        continue
                    }
                }
            } catch {
            }
        }
        if (anchorButton) {
            try {
                anchorButton.SetFocus()
                Sleep 25
                SendInput "+{Tab}"
                Sleep 15
            } catch {
                try {
                    promptField := FindGeminiPromptField(uia)
                    if (promptField)
                        promptField.SetFocus()
                } catch {
                }
            }
        } else {
            try {
                promptField := FindGeminiPromptField(uia)
                if (promptField)
                    promptField.SetFocus()
            } catch {
            }
        }
    } catch {
    }

    if (geminiHwnd && WinExist("ahk_id " geminiHwnd)) {
        WinActivate("ahk_id " geminiHwnd)
        WinWaitActive("ahk_id " geminiHwnd, , 2)
        Sleep 150
    }

    if (optionalPromptText != "") {
        InsertText(optionalPromptText)
    } else {
        ClipAngel_SendTopListItem(geminiHwnd)
    }
    Sleep 250
    ScriptSoundPlay(A_ScriptDir . "\assets\sounds\gemini-focused.wav")
}

GetGrammarPromptText() {
    return GetPromptText("grammar")
}

GetAioptPromptText() {
    return GetPromptText("aiopt")
}

GeminiDelayedSubmitMonitorStartFromUtils(originalHwnd, geminiChromeHwnd) {
    WM_START_DELAYED_SUBMIT_MONITOR := 0x8002
    targetHwnd := GetGeminiScriptMsgTargetHwnd()
    if (targetHwnd) {
        try SendMessage(WM_START_DELAYED_SUBMIT_MONITOR, originalHwnd, geminiChromeHwnd, , "ahk_id " targetHwnd)
    }
}

GeminiDelayedSubmitMonitorStopFromUtils() {
    WM_STOP_DELAYED_SUBMIT_MONITOR := 0x8003
    targetHwnd := GetGeminiScriptMsgTargetHwnd()
    if (targetHwnd) {
        try SendMessage(WM_STOP_DELAYED_SUBMIT_MONITOR, 0, 0, , "ahk_id " targetHwnd)
    }
}

DEPRECATED_GeminiDictationPasteOnlyFlow() {
    restoreHwnd := WinExist("A")
    GeminiNavigateFocusAndPasteFirstSnippet("", false)
    if (restoreHwnd && WinExist("ahk_id " restoreHwnd))
        WinActivate("ahk_id " restoreHwnd)
}

DEPRECATED_GeminiDelayedSubmitFlow() {
    global g_HotstringGeminiAutoSubmit, g_HotstringGeminiRestoreHwnd
    g_HotstringGeminiRestoreHwnd := WinExist("A")
    g_HotstringGeminiAutoSubmit := true
    GeminiFinalizeSubmit()
}

GeminiPromptFieldGetText() {
    try {
        uia := UIA_Browser()
        pf := FindGeminiPromptField(uia)
        if (!pf)
            return ""
        try {
            text := Trim(pf.Value)
            if (text != "")
                return text
        } catch {
        }
        try {
            text := Trim(pf.TextPattern.DocumentRange.GetText(-1))
            if (text != "")
                return text
        } catch {
        }
    } catch {
    }
    return ""
}

GeminiFinalizeSubmit() {
    global g_HotstringGeminiAutoSubmit, g_HotstringGeminiRestoreHwnd, g_GeminiDelayedSubmit_PreEnterDelayMs,
        g_GeminiDelayedSubmit_WaitContentMaxMs

    try Hotkey("n", "Off")
    try Hotkey("N", "Off")
    try Hotkey("y", "Off")
    try Hotkey("Y", "Off")
    HotstringGeminiBanner_Hide()

    GeminiNavigateFocusAndPasteFirstSnippet("", false)

    didAutoSubmit := false
    geminiChromeHwnd := 0
    if (g_HotstringGeminiAutoSubmit) {
        Sleep (g_GeminiDelayedSubmit_PreEnterDelayMs)
        pollIntervalMs := 200
        endTick := A_TickCount + g_GeminiDelayedSubmit_WaitContentMaxMs
        while (A_TickCount < endTick) {
            if (GeminiPromptFieldGetText() != "")
                break
            Sleep pollIntervalMs
        }
        Send("{Enter}")
        geminiChromeHwnd := WinExist("A")
        didAutoSubmit := true
    }

    g_HotstringGeminiAutoSubmit := true

    if (g_HotstringGeminiRestoreHwnd && WinExist("ahk_id " g_HotstringGeminiRestoreHwnd)) {
        WinActivate("ahk_id " g_HotstringGeminiRestoreHwnd)
    }

    if (didAutoSubmit && geminiChromeHwnd)
        GeminiDelayedSubmitMonitorStartFromUtils(g_HotstringGeminiRestoreHwnd, geminiChromeHwnd)
}
