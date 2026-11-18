; =============================================================================
; Gemini module: gemini_open.ahk
; InitializeGeminiFirstTime and #!+I open/focus hotkey
;   1× tap  = open/focus the resolved companion
;   2× tap  = new tab in that window with a fresh companion page
; Double-tap window: AI_QD_DOUBLE_TAP_MS (400) — matches #!+F / ZMK tap-dance.
; Extracted verbatim from Gemini.ahk; loaded via #include into the
; Gemini.ahk process, which remains the entry point / source of truth.
; =============================================================================

; =============================================================================
; Initialize Gemini window on first-time opening
; =============================================================================
InitializeGeminiFirstTime() {
    try {
        ; Show banner to inform user
        StandardLoadingBar_Show("📤 Opening Gemini...", BANNER_ACCENT_INTERMEDIATE)

        ; Remember existing Chrome windows so we can find the one we're about to create
        existingChromeHwnds := []
        try {
            for hwnd in WinGetList("ahk_exe chrome.exe")
                existingChromeHwnds.Push(hwnd)
        } catch {
        }

        ; Run Chrome with a new window and one Gemini tab
        Run "chrome.exe --new-window https://gemini.google.com/"

        ; Find the newly created Chrome window (WinGetList + optional CREATE hook)
        geminiHwnd := WaitForNewChromeWindow(existingChromeHwnds, GEMINI_FIRST_LAUNCH_MAX_LOOPS *
            GEMINI_FIRST_LAUNCH_POLL_MS)
        if !geminiHwnd {
            StandardLoadingBar_Hide(0)
            return
        }

        ; Activate the new Gemini window and wait until it is actually active
        try {
            WinActivate("ahk_id " geminiHwnd)
        } catch {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        if !WinWaitActive("ahk_id " geminiHwnd, , GEMINI_ACTIVATE_WAIT_MS // 1000) {
            StandardLoadingBar_Hide(0)
            return
        }

        try {
            uia := UIA_Browser("ahk_id " geminiHwnd)
            Gemini_WaitForPromptFieldDiscoverable(uia, 8000)
            focusPhase := ""
            promptField := 0
            deadline := A_TickCount + 5000
            while (A_TickCount < deadline) {
                if (promptField := Gemini_FocusPromptWithChime(uia, "", &focusPhase))
                    break
                Sleep 100
            }
            if (promptField) {
            }
        } catch {
        }

        StandardLoadingBar_Hide(0)
    } catch Error as err {
        ; Hide banner5 on error
        StandardLoadingBar_Hide(0)
    }
}

; =============================================================================
; Open Gemini / resolved companion
; Hotkey: Win+Alt+Shift+I
;   1× tap  = open/focus the existing companion (unchanged)
;   2× tap  = new tab in that companion window, fresh page, new chat
; =============================================================================
global g_GeminiOpen_DoubleTapArmed := false
global g_GeminiOpen_LastPressTick := 0
global g_GeminiOpen_DoubleTapTimer := 0

class GeminiOpen_DoubleTapTimerObj {
    static OnSingleTapTimeout() {
        global g_GeminiOpen_DoubleTapArmed, g_GeminiOpen_DoubleTapTimer
        if (!g_GeminiOpen_DoubleTapArmed)
            return
        g_GeminiOpen_DoubleTapArmed := false
        g_GeminiOpen_DoubleTapTimer := 0
        GeminiOpen_OpenOrFocus()
    }
}

GeminiOpen_DisarmDoubleTap() {
    global g_GeminiOpen_DoubleTapArmed, g_GeminiOpen_DoubleTapTimer, g_GeminiOpen_LastPressTick
    g_GeminiOpen_DoubleTapArmed := false
    g_GeminiOpen_LastPressTick := 0
    if (g_GeminiOpen_DoubleTapTimer) {
        SetTimer(g_GeminiOpen_DoubleTapTimer, 0)
        g_GeminiOpen_DoubleTapTimer := 0
    }
}

GeminiOpen_WaitChordUp(timeoutMs := 500) {
    start := A_TickCount
    while (A_TickCount - start < timeoutMs) {
        if (!GetKeyState("LWin", "P") && !GetKeyState("RWin", "P") && !GetKeyState("LAlt", "P")
        && !GetKeyState("RAlt", "P") && !GetKeyState("LShift", "P") && !GetKeyState("RShift", "P")
        && !GetKeyState("i", "P"))
            return true
        Sleep 15
    }
    return false
}

GeminiOpen_LaunchUrl() {
    companion := ResolveGlobalAICompanion()
    if (companion = "enterprise") {
        try return GeminiEnterprise_GetLaunchUrl()
        return ""
    }
    if (companion = "copilot") {
        try return CopilotWeb_GetLaunchUrl()
        return ""
    }
    return "https://gemini.google.com/"
}

GeminiOpen_UrlHost(url) {
    if RegExMatch(url, "i)^https?://([^/]+)", &m)
        return m[1]
    return url
}

GeminiOpen_CompanionHwnd() {
    companion := ResolveGlobalAICompanion()
    if (companion = "enterprise") {
        try return GetGeminiEnterpriseWindowHwnd()
        return 0
    }
    if (companion = "copilot") {
        try return GetCopilotWebWindowHwnd()
        return 0
    }
    return GetGeminiWindowHwnd()
}

GeminiOpen_FindOmnibox(hwnd) {
    try {
        root := UIA.ElementFromHandle(hwnd)
        if (!root)
            return 0
        for criteria in [{ Type: "Edit", AcceleratorKey: "Ctrl+L" }, { Type: "Edit", AutomationId: "view_1012" }, {
            Type: "Edit", Name: "Address and search bar" }] {
            try {
                el := root.FindFirst(criteria)
                if (el)
                    return el
            } catch {
            }
        }
    } catch {
    }
    return 0
}

GeminiOpen_OmniboxText(el) {
    if (!IsObject(el))
        return ""
    try return String(el.Value)
    return ""
}

GeminiOpen_TabCount(hwnd) {
    try {
        root := UIA.ElementFromHandle(hwnd)
        if (!root)
            return 0
        tabs := root.FindAll({ Type: "TabItem" })
        return tabs.Length
    } catch {
        return 0
    }
}

GeminiOpen_OmniboxIsFresh(text, previous) {
    if (text = "" || InStr(text, "chrome://new", false))
        return true
    return previous != "" && text != previous
}

; Ctrl+T in the companion window, then confirm a new tab took the address bar.
GeminiOpen_OpenBlankTab(hwnd) {
    beforeCount := GeminiOpen_TabCount(hwnd)
    beforeText := GeminiOpen_OmniboxText(GeminiOpen_FindOmnibox(hwnd))
    ChordSend_WithoutRestoringModifiers("^t")
    start := A_TickCount
    while (A_TickCount - start < 1500) {
        count := GeminiOpen_TabCount(hwnd)
        text := GeminiOpen_OmniboxText(GeminiOpen_FindOmnibox(hwnd))
        countGrew := beforeCount > 0 && count > beforeCount
        fresh := GeminiOpen_OmniboxIsFresh(text, beforeText)
        if (countGrew && fresh)
            return true
        if (beforeCount = 0 && fresh && text != beforeText)
            return true
        Sleep 40
    }
    return false
}

GeminiOpen_CommitUrl(hwnd, url, host) {
    omnibox := GeminiOpen_FindOmnibox(hwnd)
    if (!omnibox)
        return false
    try omnibox.SetFocus()
    catch {
    }
    Sleep 40
    stuck := false
    try {
        omnibox.ValuePattern.SetValue(url)
        stuck := InStr(GeminiOpen_OmniboxText(omnibox), host, false)
    } catch {
        stuck := false
    }
    if (!stuck) {
        try {
            omnibox.Value := url
            stuck := InStr(GeminiOpen_OmniboxText(omnibox), host, false)
        } catch {
            stuck := false
        }
    }
    if (!stuck) {
        clipSaved := ClipboardAll()
        try {
            A_Clipboard := url
            if (ClipWait(1)) {
                try omnibox.SetFocus()
                ChordSend_WithoutRestoringModifiers("^a^v")
                Sleep 60
                stuck := InStr(GeminiOpen_OmniboxText(GeminiOpen_FindOmnibox(hwnd)), host, false)
            }
        } catch {
            stuck := false
        } finally {
            try A_Clipboard := clipSaved
            catch {
            }
        }
    }
    if (!stuck)
        return false
    try omnibox.SetFocus()
    ControlSend("{Enter}", , "ahk_id " hwnd)
    return true
}

GeminiOpen_FindOnscreenDoc(hwnd, host) {
    try {
        root := UIA.ElementFromHandle(hwnd)
        if (!root)
            return 0
        docs := root.FindAll({ Type: "Document" })
        for doc in docs {
            val := ""
            try val := doc.Value
            if (!InStr(val, host, false))
                continue
            off := true
            try off := doc.GetPropertyValue(UIA.Property.IsOffscreen)
            if (!off)
                return doc
        }
    } catch {
    }
    return 0
}

GeminiOpen_FocusDocPrompt(doc) {
    for name in GEMINI_PROMPT_FIELD_NAMES {
        try {
            el := doc.FindFirst({ Name: name, Type: 50004 })
            if (el)
                return el
        } catch {
        }
    }
    try {
        el := doc.FindFirst({ Type: "Edit", ClassName: "ql-editor", matchmode: "Substring" })
        if (el)
            return el
    } catch {
    }
    return 0
}

GeminiOpen_FocusFreshPage(hwnd, companion, host, playChime := true) {
    deadline := A_TickCount + 8000
    while (A_TickCount < deadline) {
        if (!WinExist("ahk_id " hwnd))
            return false
        doc := GeminiOpen_FindOnscreenDoc(hwnd, host)
        if (IsObject(doc)) {
            try {
                if (companion = "enterprise") {
                    if (GeminiEnterprise_FocusComposer(doc, playChime)) {
                        GeminiEnterprise_CacheHwnd(hwnd)
                        GeminiEnterprise_RefreshHotkeyContext(hwnd, true)
                        return true
                    }
                } else if (companion = "copilot") {
                    if (CopilotWeb_FocusComposer(doc, playChime)) {
                        CopilotWeb_CacheHwnd(hwnd)
                        CopilotWeb_RefreshHotkeyContext(hwnd, true)
                        return true
                    }
                } else {
                    el := GeminiOpen_FocusDocPrompt(doc)
                    if (el) {
                        try el.SetFocus()
                        catch {
                        }
                        Sleep 40
                        focusedOk := false
                        try {
                            if (el.HasKeyboardFocus)
                                focusedOk := true
                        } catch {
                        }
                        if (!focusedOk) {
                            try el.Click()
                            Sleep 40
                        }
                        startFocus := A_TickCount
                        while (A_TickCount - startFocus < 350) {
                            try {
                                if (el.HasKeyboardFocus) {
                                    focusedOk := true
                                    break
                                }
                            } catch {
                                focusedOk := true
                                break
                            }
                            Sleep 25
                        }
                        if (focusedOk) {
                            if (playChime)
                                Utils_PlayGeminiFocusedChime()
                            return true
                        }
                    }
                }
            }
            catch {
            }
        }
        Sleep 200
    }
    return false
}

; New tab in the main companion window, loaded with a fresh companion page and a new chat.
GeminiOpen_NewCompanionTab() {
    companion := ResolveGlobalAICompanion()
    res := AiCompanionCDP_Ensure(companion, true, true)
    if (!res.ok)
        AiCompanionCDP_NotifyError("Open new tab", res.error, companion)
    else {
        Utils_PlayGeminiFocusedChime()
        try ShowCenteredOverlay_Utils("🌐 Opened new " . (companion = "enterprise" ? "Enterprise" : "Gemini") . " tab (CDP)", 1200, BANNER_ACCENT_SUCCESS)
        catch {
        }
    }
    return res.ok
}

GeminiOpen_OpenOrFocus() {
    companion := ResolveGlobalAICompanion()
    if (companion = "copilot")
        return CopilotWeb_OpenOrFocus()
    compLabel := (companion = "enterprise" ? "Enterprise" : "Gemini")
    try StandardLoadingBar_Show("🌐 Opening " . compLabel . " (CDP)…", BANNER_ACCENT_INTERMEDIATE)
    catch {
    }
    res := AiCompanionCDP_Ensure(companion, true, false)
    try StandardLoadingBar_Hide(0) 
    catch {
    }
    if (!res.ok)
        AiCompanionCDP_NotifyError("Open companion", res.error, companion)
    else {
        Utils_PlayGeminiFocusedChime()
        try ShowCenteredOverlay_Utils("🌐 Focused " . compLabel . " (CDP)", 1000, BANNER_ACCENT_SUCCESS)
        catch {
        }
    }
    return res.ok
}

#!+i:: {
    global g_GeminiOpen_DoubleTapArmed, g_GeminiOpen_LastPressTick, g_GeminiOpen_DoubleTapTimer

    ; Programmatic / synthetic trigger (SendInput or Send): open/focus companion directly
    if !GetKeyState("i", "P") {
        GeminiOpen_OpenOrFocus()
        return
    }

    thresholdMs := 400
    try thresholdMs := AI_QD_DOUBLE_TAP_MS
    catch {
        thresholdMs := 400
    }

    pressTime := A_TickCount
    elapsed := (g_GeminiOpen_LastPressTick > 0) ? (pressTime - g_GeminiOpen_LastPressTick) : 9999
    isSecondTap := g_GeminiOpen_DoubleTapArmed && elapsed >= 0 && elapsed < thresholdMs

    KeyWait "i"

    if (isSecondTap) {
        GeminiOpen_DisarmDoubleTap()
        GeminiOpen_NewCompanionTab()
        return
    }

    if (g_GeminiOpen_DoubleTapTimer) {
        SetTimer(g_GeminiOpen_DoubleTapTimer, 0)
        g_GeminiOpen_DoubleTapTimer := 0
    }
    g_GeminiOpen_LastPressTick := A_TickCount
    g_GeminiOpen_DoubleTapArmed := true
    g_GeminiOpen_DoubleTapTimer := ObjBindMethod(GeminiOpen_DoubleTapTimerObj, "OnSingleTapTimeout")
    SetTimer(g_GeminiOpen_DoubleTapTimer, -thresholdMs)
}

; Ready chime for flows that do not use Gemini_FocusPromptWithChime (e.g. legacy call sites).
Gemini_PlayReadyChime(minIntervalMs := 400) {
    return Utils_PlayGeminiFocusedChime(minIntervalMs)
}
