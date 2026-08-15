; =============================================================================
; Utils module: gemini_paste_helpers.ahk
; Gemini navigate/focus/paste helpers used by D2C and legacy dictation flows.
; =============================================================================

global g_HotstringGeminiAutoSubmit := true
global g_HotstringGeminiRestoreHwnd := 0
global g_GeminiDelayedSubmit_PreEnterDelayMs := 1000
global g_GeminiDelayedSubmit_WaitContentMaxMs := 5000

Gemini_GetSearchRoot(uia) {
    if (!IsObject(uia))
        return 0
    try {
        root := uia.GetCurrentMainPaneElement()
        if (root)
            return root
    } catch {
    }
    return uia
}

GeminiNavigateFocusAndPasteFirstSnippet(optionalPromptText := "", switchToFirstTab := true) {
    prevTitleMatchMode := A_TitleMatchMode
    SetTitleMatchMode(2)
    geminiHwnd := 0
    try {
        if (!switchToFirstTab) {
            try {
                activeHwnd := WinExist("A")
                if (activeHwnd && WinGetProcessName("ahk_id " activeHwnd) = "chrome.exe" && InStr(WinGetTitle(
                    "ahk_id " activeHwnd), "gemini", false))
                    geminiHwnd := activeHwnd
            } catch {
            }
        }
        if (!geminiHwnd) {
            try {
                for hwnd in WinGetList("ahk_exe chrome.exe") {
                    try {
                        if IsConsumerGeminiChromeTitle(WinGetTitle("ahk_id " hwnd)) {
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
                return 0
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
            uia := geminiHwnd ? UIA_Browser("ahk_id " geminiHwnd) : UIA_Browser()
            Sleep 80
            Gemini_FocusPromptWithChime(uia, { playChime: false, useAnchorFallback: true })
        } catch {
        }

        if (optionalPromptText != "") {
            InsertText(optionalPromptText)
        } else {
            ClipAngel_SendTopListItem(geminiHwnd)
        }
        Sleep 250
        ScriptSoundPlay(A_ScriptDir . "\assets\sounds\gemini-focused.wav")
        return geminiHwnd ? geminiHwnd : WinExist("A")
    } finally {
        SetTitleMatchMode(prevTitleMatchMode)
    }
}

Gemini_IsExcludedSendButtonName(name) {
    if (!name)
        return true
    if (InStr(name, "Stop response", false) || InStr(name, "Interromper", false) || InStr(name, "Stop streaming",
        false))
        return true
    if (InStr(name, "Open upload file menu", false) || name = "Tools")
        return true
    return false
}

Gemini_IsSendButtonCandidate(btn) {
    try {
        name := btn.Name
        className := btn.ClassName
        if (Gemini_IsExcludedSendButtonName(name))
            return false
        if (InStr(className, "send-button", false) && !InStr(className, "stop", false))
            return true
        if (InStr(name, "Send", false) || InStr(name, "Enviar", false))
            return true
    } catch {
    }
    return false
}

; Send/stop discovery: FindFirstBuildCache on main pane only; no FindAll fallback (canon §3).
Gemini_FindSendButton(uia) {
    if (!IsObject(uia))
        return 0
    static cacheRequest := 0
    if (!cacheRequest)
        cacheRequest := UIA.CreateCacheRequest(["Name", "ClassName"], , 5)
    root := Gemini_GetSearchRoot(uia)
    if (!root)
        return 0
    try {
        el := root.FindFirstBuildCache(cacheRequest, { Type: UIA_ControlType_Button, ClassName: "send-button",
            matchmode: "Substring" }, 4)
        if (el && Gemini_IsSendButtonCandidate(el))
            return el
    } catch {
    }
    for nameSpec in ["Send", "Enviar"] {
        try {
            el := root.FindFirstBuildCache(cacheRequest, { Type: UIA_ControlType_Button, Name: nameSpec,
                matchmode: "Substring" }, 4)
            if (el && Gemini_IsSendButtonCandidate(el))
                return el
        } catch {
        }
    }
    return 0
}

Gemini_HasGeneratingStopButtonForUia(uia) {
    if (!IsObject(uia))
        return false
    static cacheRequest := 0
    if (!cacheRequest)
        cacheRequest := UIA.CreateCacheRequest(["Name", "ClassName"], , 5)
    root := Gemini_GetSearchRoot(uia)
    if (!root)
        return false
    for n in ["Stop response", "Stop streaming", "Interromper transmissão"] {
        try {
            if (root.FindFirstBuildCache(cacheRequest, { Type: UIA_ControlType_Button, Name: n }, 4))
                return true
        } catch {
        }
    }
    return false
}

Gemini_SubmitAttemptSucceeded(uia) {
    if (Gemini_HasGeneratingStopButtonForUia(uia))
        return true
    return GeminiPromptFieldGetTextFromUia(uia) = ""
}

GeminiPromptFieldGetTextFromUia(uia) {
    if (!IsObject(uia))
        return ""
    pf := 0
    try pf := FindGeminiPromptField(uia)
    catch {
    }
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
    return ""
}

Gemini_StripComposerHumanReminders() {
    try {
        uia := UIA_Browser()
        text := GeminiPromptFieldGetTextFromUia(uia)
        if (text = "")
            return false
        newText := StripPromptHumanReminders(text)
        if (newText = "")
            return false
        phase := ""
        if !Gemini_FocusPromptWithChime(uia, { playChime: false }, &phase) {
            pf := FindGeminiPromptField(uia)
            if pf {
                try pf.SetFocus()
                catch {
                }
            }
        }
        return ReplaceFocusedEditWithText(newText)
    } catch {
        return false
    }
}

Gemini_TrySubmitOnce(uia, fallback := "enter") {
    if (!IsObject(uia))
        return false
    if (Gemini_HasGeneratingStopButtonForUia(uia))
        return true
    sendBtn := Gemini_FindSendButton(uia)
    if (sendBtn) {
        try {
            if (sendBtn.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)) {
                sendBtn.InvokePattern.Invoke()
                return true
            }
        } catch {
        }
        try {
            sendBtn.Click()
            return true
        } catch {
        }
    }
    if (fallback = "ctrlEnter")
        SendInput "^{Enter}"
    else
        SendInput "{Enter}"
    return true
}

; Bounded content wait: minSettleMs floor for Clip Angel paste settle; 25 ms poll for early exit (canon §13).
Gemini_WaitForPromptContent(uia, maxMs, minSettleMs := "") {
    global g_GeminiDelayedSubmit_PreEnterDelayMs
    if (minSettleMs = "")
        minSettleMs := g_GeminiDelayedSubmit_PreEnterDelayMs
    if (!IsObject(uia))
        return false
    t0 := A_TickCount
    loop {
        if (GeminiPromptFieldGetTextFromUia(uia) != "" && (A_TickCount - t0) >= minSettleMs)
            return true
        if (A_TickCount - t0 >= maxMs)
            break
        Sleep GEMINI_PROMPT_FOCUS_POLL_MS
    }
    return GeminiPromptFieldGetTextFromUia(uia) != ""
}

Gemini_TrySubmit(geminiHwnd, uia := 0) {
    if (!geminiHwnd || !WinExist("ahk_id " geminiHwnd))
        return false
    if (!WinActive("ahk_id " geminiHwnd)) {
        try WinActivate("ahk_id " geminiHwnd)
        catch {
            return false
        }
        if (!WinWaitActive("ahk_id " geminiHwnd, , 1))
            return false
    }
    if (!IsObject(uia)) {
        try uia := UIA_Browser("ahk_id " geminiHwnd)
        catch {
            return false
        }
    }
    if (!IsObject(uia))
        return false
    for fallback in ["enter", "ctrlEnter"] {
        Gemini_FocusPromptWithChime(uia, { playChime: false, useAnchorFallback: true })
        Gemini_TrySubmitOnce(uia, fallback)
        endTick := A_TickCount + 2000
        while (A_TickCount < endTick) {
            if (Gemini_SubmitAttemptSucceeded(uia))
                return true
            Sleep 200
        }
    }
    return false
}

Gemini_WaitForPromptContentAndSubmit(geminiHwnd) {
    global g_GeminiDelayedSubmit_WaitContentMaxMs
    if (!geminiHwnd || !WinExist("ahk_id " geminiHwnd))
        return false
    uia := 0
    try uia := UIA_Browser("ahk_id " geminiHwnd)
    catch {
        return false
    }
    if (!IsObject(uia))
        return false
    if (!Gemini_WaitForPromptContent(uia, g_GeminiDelayedSubmit_WaitContentMaxMs))
        return false
    return Gemini_TrySubmit(geminiHwnd, uia)
}

GetGrammarPromptText() {
    return GetPromptText("grammar")
}

GetAioptPromptText() {
    return GetPromptText("aiopt")
}

GetMtaskPromptText() {
    return GetPromptText("mtask")
}

; Runtime prompt: instructions + live INI catalogs (accounts, expense/income categories).
GetMobillsPromptText() {
    body := GetPromptText("mobills-daily")
    accounts := MobillsDaily_RenderIniCatalog(A_ScriptDir "\accounts.ini", "Accounts")
    expenses := MobillsDaily_RenderIniCatalog(A_ScriptDir "\categories-expenses.ini", "Expense categories")
    incomes := MobillsDaily_RenderIniCatalog(A_ScriptDir "\categories-income.ini", "Income categories")
    extra := ""
    if (accounts != "")
        extra .= "`n`n" . accounts
    if (expenses != "")
        extra .= "`n`n" . expenses
    if (incomes != "")
        extra .= "`n`n" . incomes
    extra .= "`n`nCredit card (always): Mercado Pago"
    extra .= "`n`nDictation follows."
    return body . extra
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

GeminiPromptFieldGetText(geminiHwnd := 0) {
    try {
        uia := geminiHwnd ? UIA_Browser("ahk_id " geminiHwnd) : UIA_Browser()
        return GeminiPromptFieldGetTextFromUia(uia)
    } catch {
    }
    return ""
}

GeminiFinalizeSubmit() {
    global g_HotstringGeminiAutoSubmit, g_HotstringGeminiRestoreHwnd

    try Hotkey("n", "Off")
    try Hotkey("N", "Off")
    try Hotkey("y", "Off")
    try Hotkey("Y", "Off")
    HotstringGeminiBanner_Hide()

    geminiChromeHwnd := GeminiNavigateFocusAndPasteFirstSnippet("", false)
    if (!geminiChromeHwnd)
        geminiChromeHwnd := WinExist("A")
    didAutoSubmit := false
    if (g_HotstringGeminiAutoSubmit && geminiChromeHwnd)
        didAutoSubmit := Gemini_WaitForPromptContentAndSubmit(geminiChromeHwnd)

    g_HotstringGeminiAutoSubmit := true

    if (g_HotstringGeminiRestoreHwnd && WinExist("ahk_id " g_HotstringGeminiRestoreHwnd)) {
        WinActivate("ahk_id " g_HotstringGeminiRestoreHwnd)
    }

    if (didAutoSubmit && geminiChromeHwnd)
        GeminiDelayedSubmitMonitorStartFromUtils(g_HotstringGeminiRestoreHwnd, geminiChromeHwnd)
}
