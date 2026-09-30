; =============================================================================
; Utils module: utility_shortcuts.ahk
; Utility shortcuts #!+U, #!+W, #!+L and ^!# secondary triggers
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; Hotkey Handler: Windows + Alt + Shift + U (*#!+U)
; * so the menu still opens while Control is stuck down.
; =============================================================================
*#!+U::
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
; * so the menu still opens while Control is stuck down.
*#!+w::
{
    ShowHotstringSelector("Macros")
}

; Wait out the opener chord (briefly) so its key-up is not lost when the GUI takes focus.
UtilityShortcuts_WaitForOpenerChord() {
    th := ""
    try th := A_ThisHotkey
    catch {
        th := ""
    }
    if (th = "")
        return
    tw := "T0.4"
    if InStr(th, "#") {
        try KeyWait "LWin", tw
        try KeyWait "RWin", tw
    }
    if InStr(th, "!") {
        try KeyWait "LAlt", tw
        try KeyWait "RAlt", tw
    }
    if InStr(th, "+") {
        try KeyWait "LShift", tw
        try KeyWait "RShift", tw
    }
    if InStr(th, "^") {
        try KeyWait "LControl", tw
        try KeyWait "RControl", tw
    }
    hk := RegExReplace(th, "^[$*~]+")
    hk := RegExReplace(hk, "[#^!+<>*]+", "")
    if (StrLen(hk) = 1) {
        try KeyWait hk, tw
    }
}

; Force Control/Alt/Shift/Win up. {Blind} so AutoHotkey does not press them again.
; keybd_event covers the case where the hook already dropped the physical key-up.
UtilityShortcuts_ReleaseStuckModifiers() {
    ChordSend_WithoutRestoringModifiers()
    for vk in [0x11, 0xA2, 0xA3, 0x12, 0xA4, 0xA5, 0x10, 0xA0, 0xA1, 0x5B, 0x5C]
        DllCall("keybd_event", "UChar", vk, "UChar", 0, "UInt", 2, "UPtr", 0)
}

MacroReleaseStuckControl(*) {
    UtilityShortcuts_ReleaseStuckModifiers()
    try ShowCenteredOverlay_Utils("Control released", 1200, BANNER_ACCENT_SUCCESS)
}

RegisterMacro(MacroReleaseStuckControl, "🔓 Release stuck Control", "u")

; Macros [W] — force-kill heavy apps and their child processes to free RAM.
; Stop resident servers first so warmup does not relaunch Finance / Memory Palace.
MacroWipeHeavyApps(*) {
    ShowCenteredOverlay_Utils("🧹 Closing heavy apps...", 1500, BANNER_ACCENT_INTERMEDIATE)
    try StopResidentWebServers()
    catch {
    }
    for name in [
        "chrome.exe", "Cursor.exe", "Code.exe",
        "ms-teams.exe", "Teams.exe", "MSTeams.exe",
        "OUTLOOK.EXE", "olk.exe",
        "msedge.exe", "ONENOTE.EXE", "WhatsApp.exe", "Spotify.exe"
    ]
        WipeKillProcessTree(name)
    try WhatsAppJump_InvalidateHwndCache()
    catch {
    }
    ShowCenteredOverlay_Utils("✅ Heavy apps closed", 1500, BANNER_ACCENT_SUCCESS)
}

; taskkill /T so Electron/Chrome children (renderers, language servers) die with the parent.
; ProcessClose leaves those holding RAM. Skip images that are not running.
WipeKillProcessTree(imageName) {
    if !ProcessExist(imageName)
        return
    try RunWait(A_ComSpec . " /c taskkill /F /T /IM " . imageName . " >nul 2>&1", , "Hide")
    catch {
    }
}

RegisterMacro(MacroWipeHeavyApps, "🧹 Wipe heavy apps (free RAM)", "w")

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

; Macros [L] — selected text (Ctrl+C like #!+8) or fallback clipboard; open in Chrome / Google search.
OpenClipboardLinkInChrome() {
    savedClip := ""
    try savedClip := A_Clipboard
    catch {
        savedClip := ""
    }
    ; Prefer selection in the window that opened Utility Shortcuts (#!+8 pattern).
    UtilitySelector_RestorePreviousHwnd()
    Sleep 150
    A_Clipboard := ""
    copied := TryCopySelectionToClipboard_QuickLookAware()
    if (copied) {
        text := Trim(A_Clipboard)
    } else {
        try A_Clipboard := savedClip
        catch {
        }
        text := Trim(savedClip)
    }
    if (text = "") {
        ShowCenteredOverlay_Utils("❌ No selection and clipboard is empty.", 2000, BANNER_ACCENT_ERROR)
        return
    }
    url := OpenClipboardLinkInChrome_ExtractHttpUrl(text)
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

RegisterMacro(OpenClipboardLinkInChrome, "🔗 Open selected/clipboard text/link in Chrome / Google search", "l")

; Study links — one-shot set (clipboard URL) / open (new Chrome). Keys: 1/2/3 set, v/a/f open.
MacroStudyLink_SetVideo(*) {
    StudyLink_SetFromClipboard(STUDYLINK_KEY_YOUTUBE, "video link")
}
MacroStudyLink_SetArticle(*) {
    StudyLink_SetFromClipboard(STUDYLINK_KEY_ARTICLE, "article link")
}
MacroStudyLink_SetFavorite(*) {
    StudyLink_SetFromClipboard(STUDYLINK_KEY_FAVORITE, "favorite link")
}
MacroStudyLink_OpenVideo(*) {
    StudyLink_Open(STUDYLINK_KEY_YOUTUBE)
}
MacroStudyLink_OpenArticle(*) {
    StudyLink_Open(STUDYLINK_KEY_ARTICLE)
}
MacroStudyLink_OpenFavorite(*) {
    StudyLink_Open(STUDYLINK_KEY_FAVORITE)
}

RegisterMacro(MacroStudyLink_SetVideo, "📹 Set video from clipboard", "1")
RegisterMacro(MacroStudyLink_SetArticle, "📖 Set article from clipboard", "2")
RegisterMacro(MacroStudyLink_SetFavorite, "❤️ Set favorite from clipboard", "3")
RegisterMacro(MacroStudyLink_OpenVideo, "📹 Open video", "v")
RegisterMacro(MacroStudyLink_OpenArticle, "📖 Open article", "a")
RegisterMacro(MacroStudyLink_OpenFavorite, "❤️ Open favorite", "f")

; Win+Alt+Shift+L — paste OS clipboard (^v) to a picked visible window (same as D2C menu [W]).
; After pick: [Y] paste+Enter, [N] paste only, [Esc] abort, 3s timeout = paste only.
; After paste (+ optional learn-field prompt), restores focus to the window that was active before the picker.
; If a main text field is saved for that exe+title/url (assets/data/paste_field_mappings.ini),
; focus it via UIA before paste; if unknown, after paste ask [Y]/[N] to persist the focused field.
; In the picker: slot key = paste; [Q] close mode (slot key closes that window; [Q] leaves it; [ESC] closes the picker);
; [R] then slot = ignore that process (exe) for AutoSlot;
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

; Map ResolveGlobalAICompanion() → AI_COMPANION_* ids for AiCompanionModels_*.
GlobalAICompanionModelsId() {
    switch ResolveGlobalAICompanion() {
        case "enterprise":
            return AI_COMPANION_ENTERPRISE
        case "copilot":
            return AI_COMPANION_COPILOT
        default:
            return AI_COMPANION_GEMINI
    }
}

; Resolve companion hwnd + label (same pattern as ToggleAICompanionChromeTab).
GlobalAICompanionHwndAndLabel(&hwnd, &label) {
    hwnd := 0
    label := "Gemini"
    switch ResolveGlobalAICompanion() {
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
}

; Global Fast / Deep — same as in-app Shift+Q / Shift+M via AiCompanionModels_SelectRole.
; Captures the focused window first and reactivates it after the model switch
; (Macros menu restores the pre-selector window before calling us; #!q/#!m use the live foreground).
GlobalAICompanionSelectRole(role) {
    prevHwnd := 0
    try prevHwnd := WinGetID("A")
    catch {
        prevHwnd := 0
    }
    hwnd := 0
    label := "Gemini"
    GlobalAICompanionHwndAndLabel(&hwnd, &label)
    if (!hwnd) {
        ShowCenteredOverlay_Utils("❌ " . label . " is not open.", 1800, BANNER_ACCENT_ERROR)
        return false
    }
    ok := AiCompanionModels_SelectRole(GlobalAICompanionModelsId(), role)
    Handy_RestorePrevWindow(prevHwnd, hwnd)
    return ok
}

; Macros list Q/M — live Shift+L Fast/Deep name for the resolved companion.
AICompanionRoleMacroTitle(role) {
    id := GlobalAICompanionModelsId()
    name := (role = "fast") ? AiCompanionModels_GetFast(id) : AiCompanionModels_GetDeep(id)
    if (Trim(name) = "")
        name := "(not set — Shift+L)"
    label := AiCompanionModels_DisplayName(id)
    if (role = "fast")
        return "⚡ Quick — " . label . ": " . name
    return "🔄 Deep — " . label . ": " . name
}

MacroAICompanionQuickModel(*) {
    return GlobalAICompanionSelectRole("fast")
}

MacroAICompanionDeepModel(*) {
    GlobalAICompanionSelectRole("deep")
}

; Win+Alt+Q / Win+Alt+M — OS-wide Quick / Deep model on the resolved AI companion
#!q:: MacroAICompanionQuickModel()
#!m:: MacroAICompanionDeepModel()

RegisterMacro(MacroAICompanionQuickModel, "⚡ AI companion Quick / Fast model", "q")
RegisterMacro(MacroAICompanionDeepModel, "🔄 AI companion Deep model", "m")

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
; Win+Alt+Shift+O: 1× cut, 2× open, 3× paste→Desktop (clipboard / Explorer selection / Cursor·Code active file), hold 700ms+ copy path
#!+o:: DesktopCutNewest_OnHotkey()
; Ctrl+Alt+Win+9 / +B - Handy Nemotron Portuguese / Parakeet Unified English (slots 2 and 1)
^!#9:: ExecuteHandyAiModelSelection(HANDY_AI_SLOT_PORTUGUESE)
^!#b:: ExecuteHandyAiModelSelection(HANDY_AI_SLOT_ENGLISH)

; $ so the forwarded chord does not re-enter this hotkey and swallow Control's key-up.
$#^!m::
{
    Sleep 50
    SendInput "#^!m"
    UtilityShortcuts_ReleaseStuckModifiers()
    SendInput "{Blind}" "{Left}"
}
