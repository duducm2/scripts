; =============================================================================
; Utils module: utility_shortcuts.ahk
; Utility shortcuts #!+U, #!+W, #!+L and ^!# secondary triggers
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

; Win+Alt+Shift+W — Utility Shortcuts → Macros
; Same UI as #!+U then [M]; toggles closed if Macros is already open.
#!+w::
{
    ShowHotstringSelector("Macros")
}

; Extract the first http(s) URL from plain text; normalize bare www. hosts.
OpenClipboardLinkInChrome_ExtractHttpUrl(text) {
    t := Trim(text)
    if (t = "")
        return ""
    if StudyLink_IsValidHttpUrl(t)
        return t
    if RegExMatch(t, "i)(https?://[^\s`"<>]+)", &m) {
        url := RegExReplace(m[1], "[)\].,;:!?]+$")
        if StudyLink_IsValidHttpUrl(url)
            return url
    }
    if RegExMatch(t, "i)\b((?:https?://|www\.)[^\s`"<>]+)", &m2) {
        url := RegExReplace(m2[1], "[)\].,;:!?]+$")
        if (SubStr(StrLower(url), 1, 4) = "www.")
            url := "https://" url
        if StudyLink_IsValidHttpUrl(url)
            return url
    }
    return ""
}

OpenClipboardLinkInChrome_UiaElementUrl(el) {
    if !IsObject(el)
        return ""
    fields := []
    for prop in ["Value", "LegacyIAccessibleValue", "HelpText", "Name"] {
        try fields.Push(el.%prop%)
        catch {
        }
    }
    for field in fields {
        url := OpenClipboardLinkInChrome_ExtractHttpUrl(field)
        if (url != "")
            return url
    }
    return ""
}

; Prefer UIA at the mouse point — copy/double-click often returns anchor text, not href.
OpenClipboardLinkInChrome_UiaUrlFromPoint(mx := "", my := "") {
    if (mx = "" || my = "")
        MouseGetPos(&mx, &my)
    el := 0
    try el := UIA.SmallestElementFromPoint(mx, my)
    catch {
        try el := UIA.ElementFromPoint(mx, my)
        catch
            return ""
    }
    loop 12 {
        if !IsObject(el)
            break
        url := OpenClipboardLinkInChrome_UiaElementUrl(el)
        if (url != "")
            return url
        try el := el.Parent
        catch
            break
    }
    return ""
}

; Resolve URL and/or plain text from hover, selection, or clipboard fallback.
OpenClipboardLinkInChrome_ResolveTarget(&url := "", &text := "") {
    url := ""
    text := ""
    savedClipAll := ClipboardAll()
    savedClipText := Trim(A_Clipboard)
    try {
        url := OpenClipboardLinkInChrome_UiaUrlFromPoint()
        if (url != "")
            return

        MouseGetPos(, , &hwndUnder)
        if (hwndUnder) {
            try WinActivate("ahk_id " hwndUnder)
            Sleep 50
        }

        A_Clipboard := ""
        if (TryCopySelectionToClipboard_QuickLookAware()) {
            text := Trim(A_Clipboard)
            url := OpenClipboardLinkInChrome_ExtractHttpUrl(text)
            if (url != "" || text != "")
                return
        }

        A_Clipboard := ""
        Send "{LWin Up}{RWin Up}{LAlt Up}{RAlt Up}{LShift Up}{RShift Up}"
        Sleep 40
        Click 2
        Sleep 80
        Send "^c"
        if ClipWait(0.7) {
            text := Trim(A_Clipboard)
            url := OpenClipboardLinkInChrome_ExtractHttpUrl(text)
            if (url != "" || text != "")
                return
        }

        text := savedClipText
        url := OpenClipboardLinkInChrome_ExtractHttpUrl(text)
    } finally {
        try A_Clipboard := savedClipAll
    }
}

; Macros [L] — hovered hyperlink, selected text, or clipboard; open in Chrome or Google search.
OpenClipboardLinkInChrome() {
    url := ""
    text := ""
    OpenClipboardLinkInChrome_ResolveTarget(&url, &text)
    if (url = "" && text = "") {
        ShowCenteredOverlay_Utils("❌ No text to open or search.", 2000, BANNER_ACCENT_ERROR)
        return
    }
    if (url != "") {
        banner := "✅ Opening link in Chrome…"
    } else {
        url := "https://www.google.com/search?q=" . StudyLink_UrlEncode(text)
        banner := "✅ Searching on Google…"
    }
    if !StudyLink_OpenUrlInChrome(url, true) {
        ShowCenteredOverlay_Utils("❌ Could not open Chrome.", 2000, BANNER_ACCENT_ERROR)
        return
    }
    ShowCenteredOverlay_Utils(banner, 1200, BANNER_ACCENT_SUCCESS)
}

RegisterMacro(OpenClipboardLinkInChrome, "🔗 Open hovered text in Chrome / Google search", "l")

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

; Ctrl+Alt+Win+2..8 / J — dedicated chords (not listed in #!+U Macros)
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
; Win+Alt+Shift+O (same tiering as #!+8): 1× cut, 2× open, hold 700ms+ copy path
#!+o:: DesktopCutNewest_OnHotkey()
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
