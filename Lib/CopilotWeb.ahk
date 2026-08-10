; M365 Copilot web automation (Chrome). Included from Gemini.ahk after Utils.ahk.
; Requires: env.ahk, Utils.ahk, UIA_Browser.

COPILOT_WEB_ACTIVATE_WAIT_MS := 2000
COPILOT_WEB_UIA_SETTLE_MS := 120
COPILOT_WEB_SCROLL_SETTLE_MS := 350
COPILOT_WEB_TAB_SETTLE_MS := 150
COPILOT_WEB_ASYNC_POLL_MS := 500
COPILOT_WEB_ASYNC_MAX_RETRIES := 60
COPILOT_WEB_COPY_MAX_RETRIES := 3
COPILOT_WEB_COPY_RETRY_SLEEP_MS := 400
COPILOT_WEB_STREAM_GONE_VERIFY_MS := 200
COPILOT_WEB_STREAM_GONE_LOOPS := 4
COPILOT_WEB_POST_COPY_SYNC_TIMEOUT_MS := 2000
COPILOT_WEB_CLIPBOARD_POLL_MS := 10
COPILOT_WEB_FIRST_LAUNCH_WAIT_MS := 2500

COPILOT_COPY_RESPONSE_NAMES := ["Copy Response", "Copy", "Copiar"]

global g_CopilotWebCachedHwnd := 0
global g_CopilotWebHotkeyActive := false
global g_CopilotWebCachedTitle := ""
global g_CopilotWeb_ForegroundHookHandle := 0
global g_CopilotWeb_ForegroundHookCallback := 0
COPILOT_READ_ALOUD_NAMES := ["Read aloud", "Read Aloud", "Ler em voz alta"]
COPILOT_MORE_OPTIONS_NAMES := ["More options", "Show more options", "Mais opções"]
COPILOT_TTS_PAUSE_NAMES := ["Pause", "Pausar"]
COPILOT_TTS_RESUME_NAMES := ["Resume", "Retomar"]
COPILOT_VOICE_START_NAMES := ["Start a new voice chat", "Iniciar um novo chat de voz"]
COPILOT_VOICE_END_NAMES := ["End voice chat", "Encerrar chat de voz"]

UIA_Copilot_ControlType_Button := 50000
UIA_Copilot_ControlType_MenuItem := 50011
UIA_Copilot_ControlType_RadioButton := 50013
UIA_Copilot_ControlType_Text := 50020

COPILOT_NAV_EXPAND_NAMES := ["Expand navigation", "Expandir navegação"]
COPILOT_NAV_COLLAPSE_NAMES := ["Collapse navigation", "Recolher navegação"]
COPILOT_NEW_CHAT_NAMES := ["New chat", "Novo chat"]
COPILOT_NAV_SEARCH_NAMES := ["Search", "Pesquisar", "Buscar"]
COPILOT_MODEL_SELECTOR_CRITERIA := [{ AutomationId: "gptModeSwitcher", ControlType: "Button" }, { Name: "Model Selector",
    ControlType: "Button" }]
; Longer / more specific phrases first — used for scoring when multiple menu rows match.
COPILOT_DEEP_REASONING_NEEDLES := [
    "think deeper",
    "mais profundo",
    "pensar mais",
    "deeper",
    "profundo",
    "think",
    "pensar"
]
COPILOT_MODEL_MENU_WAIT_MS := 2000
COPILOT_MODEL_MENU_POLL_MS := 80
UIA_Copilot_ControlType_ListItem := 50007
COPILOT_SOURCES_BUTTON_CRITERIA := [{ Name: "Add and manage sources", ControlType: "Button" }, { Name: "Adicionar e gerenciar fontes",
    ControlType: "Button" }, { Name: "Add and manage sources", matchmode: "Substring" }]
COPILOT_SOURCES_MENU_MARKERS := [{ Name: "Add capabilities", ControlType: "MenuItem" }, { Name: "Adicionar capacidades",
    ControlType: "MenuItem" }, { Name: "Upload images and files", ControlType: "MenuItem" }, { Name: "Add work content",
        ControlType: "MenuItem" }, { Name: "Adicionar conteúdo de trabalho", ControlType: "MenuItem" }, { Name: "Add capabilities",
            matchmode: "Substring" }, { Name: "Attach cloud files", matchmode: "Substring" }
]
COPILOT_ADD_CAPABILITIES_NAMES := ["Add capabilities", "Adicionar capacidades"]
COPILOT_CAPABILITY_IMAGE_NAMES := ["Generate an image", "Criar imagem", "Criar uma imagem"]
COPILOT_CAPABILITY_IMAGE_AID := "capability-id-imageGeneration"
COPILOT_CAPABILITY_IMAGE_ENGAGE_NAMES := ["Generate an image", "Criar imagem", "Criar uma imagem", "Designer"]
COPILOT_CAPABILITY_RESEARCH_NAMES := ["Research a topic", "Pesquisar um tópico", "Researcher", "Pesquisador"]
; Sources popup order (screenshot): work content, upload, cloud, Add capabilities → Research, Analyze, Generate an image
COPILOT_CAP_MENU_DOWN_TO_ADD_CAP := 3
COPILOT_CAP_SUBMENU_DOWN_IMAGE := 2
COPILOT_CAP_SUBMENU_DOWN_RESEARCH := 0
COPILOT_COMPOSER_EXPAND_NAMES := [
    "Expand message copilot input box",
    "Expand input to Fullscreen"
]
COPILOT_COMPOSER_COLLAPSE_NAMES := [
    "Collapse message copilot input box",
    "Collapse input from Fullscreen"
]

; Optional overrides in env.ahk (included before this file). Defaults silence #Warn Unset on load.
if (!IsSet(COPILOT_WEB_URL))
    global COPILOT_WEB_URL := "https://m365.cloud.microsoft/chat"
if (!IsSet(COPILOT_WEB_URL_NEEDLE))
    global COPILOT_WEB_URL_NEEDLE := "m365.cloud.microsoft/chat"
if (!IsSet(COPILOT_WEB_TITLE_NEEDLE))
    global COPILOT_WEB_TITLE_NEEDLE := ""

; Provider routing: personal = Gemini; work = Enterprise if open (or default), else Copilot.
; GetGeminiEnterpriseWindowHwnd is defined in GeminiEnterprise.ahk (included after this file).

ResolveGlobalAICompanion() {
    global IS_WORK_ENVIRONMENT
    if (!IS_WORK_ENVIRONMENT)
        return "gemini"
    ; Prefer Enterprise whenever it is open; default launch target at work is also Enterprise.
    try {
        if (GetGeminiEnterpriseWindowHwnd())
            return "enterprise"
    } catch {
    }
    try {
        if (GetCopilotWebWindowHwnd())
            return "copilot"
    } catch {
    }
    return "enterprise"
}

UseCopilotWebForGlobalAI() {
    return ResolveGlobalAICompanion() = "copilot"
}

GetGlobalAIProviderLabel() {
    switch ResolveGlobalAICompanion() {
        case "enterprise":
            return "Gemini Enterprise"
        case "copilot":
            return "Copilot"
        default:
            return "Gemini"
    }
}

CopilotWeb_GetLaunchUrl() {
    global COPILOT_WEB_URL
    return (COPILOT_WEB_URL != "") ? COPILOT_WEB_URL : "https://m365.cloud.microsoft/chat"
}

CopilotWeb_Notify(message, durationMs := 800, fontSize := 22) {
    StandardLoadingBar_Show(message, BANNER_ACCENT_INTERMEDIATE, { passive: true, fontSize: fontSize })
    StandardLoadingBar_Hide(durationMs)
}

CopilotWeb_UrlIsChat(url) {
    global COPILOT_WEB_URL_NEEDLE
    if (!url)
        return false
    needle := (COPILOT_WEB_URL_NEEDLE != "") ? COPILOT_WEB_URL_NEEDLE : "m365.cloud.microsoft/chat"
    return InStr(url, needle)
}

CopilotWeb_IsChromeHwnd(hwnd) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    try {
        return StrLower(WinGetProcessName("ahk_id " hwnd)) = "chrome.exe"
    } catch {
        return false
    }
}

CopilotWeb_TitleMatchesCopilot(title) {
    global COPILOT_WEB_TITLE_NEEDLE
    if (!title)
        return false
    if (COPILOT_WEB_TITLE_NEEDLE != "" && InStr(title, COPILOT_WEB_TITLE_NEEDLE, false))
        return true
    if (InStr(title, "Chat | M365 Copilot", false))
        return true
    return false
}

; Background-safe UIA root (no UIA_Browser init — avoids WinActivate side effects).
CopilotWeb_ReadRootFromHwnd(hwnd) {
    try
        return UIA.ElementFromHandle(hwnd)
    catch
        return 0
}

CopilotWeb_TryUrlFromAddressBar(hwnd) {
    try {
        root := CopilotWeb_ReadRootFromHwnd(hwnd)
        if (!root)
            return false
        edit := CopilotWeb_FindFirstInUia(root, [{ Name: "Address and search bar", ControlType: "Edit" }, { AutomationId: "view_1012",
            ControlType: "Edit" }])
        if (!edit)
            return false
        url := edit.Value
        if (!url)
            return false
        if (!RegExMatch(url, "^https?://"))
            url := "https://" url
        return CopilotWeb_UrlIsChat(url)
    } catch {
        return false
    }
}

CopilotWeb_TryUrlFromDocument(hwnd) {
    try {
        root := CopilotWeb_ReadRootFromHwnd(hwnd)
        if (!root)
            return false
        doc := root.FindFirst({ ControlType: "Document" })
        if (!doc)
            return false
        return CopilotWeb_UrlIsChat(doc.Value)
    } catch {
        return false
    }
}

CopilotWeb_TryUiaFingerprint(hwnd) {
    try {
        uia := CopilotWeb_ReadRootFromHwnd(hwnd)
        if (!uia)
            return false
        if (CopilotWeb_FindComposer(uia))
            return true
        for criteria in [{ AutomationId: "m365-copilot-app-layout-main" }, { AutomationId: "copilot-message-rbp-title" }, { AutomationId: "m365-chat-editor-target-element" }, { AutomationId: "gptModeSwitcher",
            ControlType: "Button" }, { Name: "Message Copilot" }, { Name: "Copilot said:" }, { Name: "Expand navigation",
                ControlType: "Button" }, { Name: "New chat", ControlType: "Link" }, { Name: "New chat", ControlType: "Hyperlink" }, { Name: "New chat",
                    ControlType: "MenuItem" }
        ] {
            try {
                if (uia.FindFirst(criteria))
                    return true
            } catch {
            }
        }
    } catch {
    }
    return false
}

; mode: "fast" = title + address bar only; "full" = also document URL + UIA fingerprint
CopilotWeb_IsCopilotHwnd(hwnd, mode := "full") {
    if (!CopilotWeb_IsChromeHwnd(hwnd))
        return false
    try {
        if (CopilotWeb_TitleMatchesCopilot(WinGetTitle("ahk_id " hwnd)))
            return true
    } catch {
    }
    if (CopilotWeb_TryUrlFromAddressBar(hwnd))
        return true
    if (mode = "fast")
        return false
    if (CopilotWeb_TryUrlFromDocument(hwnd))
        return true
    return CopilotWeb_TryUiaFingerprint(hwnd)
}

CopilotWeb_IsCopilotWindow(hwnd) {
    return CopilotWeb_IsCopilotHwnd(hwnd, "full")
}

CopilotWeb_InvalidateCache() {
    global g_CopilotWebCachedHwnd, g_CopilotWebHotkeyActive, g_CopilotWebCachedTitle
    g_CopilotWebCachedHwnd := 0
    g_CopilotWebHotkeyActive := false
    g_CopilotWebCachedTitle := ""
}

CopilotWeb_CacheHwnd(hwnd) {
    global g_CopilotWebCachedHwnd
    if (hwnd && CopilotWeb_IsChromeHwnd(hwnd))
        g_CopilotWebCachedHwnd := hwnd
    else
        g_CopilotWebCachedHwnd := 0
}

GetCopilotWebWindowHwnd() {
    global g_CopilotWebCachedHwnd
    if (g_CopilotWebCachedHwnd && WinExist("ahk_id " g_CopilotWebCachedHwnd)) {
        if (CopilotWeb_IsCopilotHwnd(g_CopilotWebCachedHwnd, "fast"))
            return g_CopilotWebCachedHwnd
    }
    CopilotWeb_InvalidateCache()
    hwnd := WinExist("A")
    if (CopilotWeb_IsCopilotHwnd(hwnd, "full")) {
        CopilotWeb_CacheHwnd(hwnd)
        return hwnd
    }
    try {
        for h in WinGetList("ahk_exe chrome.exe") {
            if (CopilotWeb_IsCopilotHwnd(h, "full")) {
                CopilotWeb_CacheHwnd(h)
                return h
            }
        }
    } catch {
    }
    return 0
}

CopilotWeb_FindFirstInUia(uia, criteriaList) {
    if (!IsObject(uia))
        return 0
    for criteria in criteriaList {
        try {
            el := uia.FindFirst(criteria)
            if (el)
                return el
        } catch {
        }
    }
    return 0
}

CopilotWeb_FindComposer(uia) {
    return CopilotWeb_FindFirstInUia(uia, [{ AutomationId: "m365-chat-editor-target-element", ControlType: "Edit" }, { AutomationId: "m365-chat-editor-target-element" }, { Name: "Message Copilot",
        ControlType: "Edit" }])
}

CopilotWeb_FindStopGenerating(uia) {
    if (!IsObject(uia))
        return 0
    el := CopilotWeb_FindFirstInUia(uia, [{ Name: "Stop generating", ControlType: "Button" }, { Name: "Stop generating",
        Type: UIA_Copilot_ControlType_Button }])
    if (el)
        return el
    try {
        return uia.FindFirst({ Name: "Stop generating", matchmode: "Substring", ControlType: "Button" })
    } catch {
    }
    return 0
}

CopilotWeb_FindSendButton(uia) {
    return CopilotWeb_FindFirstInUia(uia, [{ Name: "Send ", matchmode: "Substring", ControlType: "Button" }, { ClassName: "fai-SendButton",
        matchmode: "Substring", ControlType: "Button" }])
}

CopilotWeb_TrySubmit(uia) {
    if (!IsObject(uia))
        return false
    if (CopilotWeb_FindStopGenerating(uia))
        return false
    sendBtn := CopilotWeb_FindSendButton(uia)
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
    SendInput "{Enter}"
    return true
}

CopilotWeb_PlayFocusedChime(minIntervalMs := 400) {
    static lastChimeTick := 0
    if (!IsSoundEnabled())
        return false
    now := A_TickCount
    if (lastChimeTick && (now - lastChimeTick) < minIntervalMs)
        return false
    lastChimeTick := now
    ScriptSoundPlay(A_ScriptDir . "\assets\sounds\gemini-focused.wav")
    return true
}

CopilotWeb_FocusComposer(uia, playChime := true) {
    el := CopilotWeb_FindComposer(uia)
    if (!el)
        return 0
    try {
        if (el.HasKeyboardFocus) {
            if (playChime)
                CopilotWeb_PlayFocusedChime()
            return el
        }
    } catch {
    }
    try el.ScrollIntoView()
    try el.SetFocus()
    Sleep 40
    try el.Click()
    Sleep 40
    try {
        if (el.HasKeyboardFocus && playChime)
            CopilotWeb_PlayFocusedChime()
    } catch {
    }
    return el
}

CopilotWeb_FocusComposerForHwnd(copilotHwnd, playChime := false) {
    if (!copilotHwnd || !WinActive("ahk_id " copilotHwnd))
        return false
    try {
        uia := UIA_Browser("ahk_id " copilotHwnd)
        return !!CopilotWeb_FocusComposer(uia, playChime)
    } catch {
    }
    return false
}

; After nav / chrome UIA clicks: type a letter into the prompt then erase it (Shift shortcuts).
CopilotWeb_ReturnToComposer() {
    Sleep 40
    Send "{Blind}d{Backspace}"
}

CopilotWeb_WaitForComposerDiscoverable(uia, timeoutMs := 500) {
    start := A_TickCount
    while (A_TickCount - start < timeoutMs) {
        if (CopilotWeb_FindComposer(uia))
            return true
        Sleep 25
    }
    return !!CopilotWeb_FindComposer(uia)
}

CopilotWeb_ActivateWindow(hwnd, timeoutMs := COPILOT_WEB_ACTIVATE_WAIT_MS) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    try {
        WinActivate("ahk_id " hwnd)
    } catch {
        return false
    }
    return WinWaitActive("ahk_id " hwnd, , timeoutMs // 1000)
}

CopilotWeb_OpenOrFocus() {
    SetTitleMatchMode(2)
    hwnd := GetCopilotWebWindowHwnd()
    if (hwnd) {
        alreadyActive := false
        try alreadyActive := WinActive("ahk_id " hwnd)
        if (!alreadyActive) {
            if !CopilotWeb_ActivateWindow(hwnd)
                return false
        }
        root := CopilotWeb_ReadRootFromHwnd(hwnd)
        if (!root) {
            ShowCenteredOverlay_Utils("❌ Error: Could not attach to Copilot window.", 2000, BANNER_ACCENT_ERROR)
            return false
        }
        if (!alreadyActive && WinActive("ahk_id " hwnd))
            CopilotWeb_WaitForComposerDiscoverable(root)
        if (WinActive("ahk_id " hwnd))
            CopilotWeb_FocusComposer(root, true)
        CopilotWeb_CacheHwnd(hwnd)
        CopilotWeb_RefreshHotkeyContext(hwnd, true)
        return true
    }
    CopilotWeb_InvalidateCache()
    try {
        StandardLoadingBar_Show("📤 Opening Copilot...", BANNER_ACCENT_INTERMEDIATE)
        Run 'chrome.exe --new-window "' CopilotWeb_GetLaunchUrl() '"'
        if !WinWaitActive("ahk_exe chrome.exe", , 8) {
            StandardLoadingBar_Hide(0)
            return false
        }
        Sleep COPILOT_WEB_FIRST_LAUNCH_WAIT_MS
        hwnd := GetCopilotWebWindowHwnd()
        if (!hwnd)
            hwnd := WinExist("A")
        if (!hwnd) {
            StandardLoadingBar_Hide(0)
            return false
        }
        root := CopilotWeb_ReadRootFromHwnd(hwnd)
        if (!root) {
            StandardLoadingBar_Hide(0)
            return false
        }
        CopilotWeb_WaitForComposerDiscoverable(root, 4000)
        if (WinActive("ahk_id " hwnd))
            CopilotWeb_FocusComposer(root, true)
        CopilotWeb_CacheHwnd(hwnd)
        CopilotWeb_RefreshHotkeyContext(hwnd, true)
        StandardLoadingBar_Hide(0)
        return true
    } catch {
        StandardLoadingBar_Hide(0)
        return false
    }
}

CopilotWeb_ComposerGetText(copilotHwnd := 0) {
    try {
        if (copilotHwnd) {
            uia := CopilotWeb_ReadRootFromHwnd(copilotHwnd)
            if (!uia)
                return ""
        } else {
            uia := UIA_Browser()
        }
        pf := CopilotWeb_FindComposer(uia)
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

; Strip human-reminder block after the last --- in the composer (keep --- + blank lines).
CopilotWeb_StripComposerHumanReminders() {
    hwnd := WinExist("A")
    text := CopilotWeb_ComposerGetText(hwnd)
    if (text = "")
        return false
    newText := StripPromptHumanReminders(text)
    if (newText = "")
        return false
    root := CopilotWeb_ReadRootFromHwnd(hwnd)
    if (IsObject(root))
        CopilotWeb_FocusComposer(root, false)
    return ReplaceFocusedEditWithText(newText)
}

; Navigate to Copilot, focus composer, paste snippet (optional text or Clip Angel first snippet).
CopilotWeb_NavigateFocusAndPaste(optionalPromptText := "", autoSubmit := false) {
    SetTitleMatchMode(2)
    copilotHwnd := GetCopilotWebWindowHwnd()
    if (!copilotHwnd) {
        Run 'chrome.exe --new-window "' CopilotWeb_GetLaunchUrl() '"'
        if !WinWaitActive("ahk_exe chrome.exe", , 5)
            return 0
        Sleep COPILOT_WEB_FIRST_LAUNCH_WAIT_MS
        copilotHwnd := GetCopilotWebWindowHwnd()
        if (!copilotHwnd)
            copilotHwnd := WinExist("A")
    }
    if (!copilotHwnd)
        return 0
    WinActivate("ahk_id " copilotHwnd)
    WinWaitActive("ahk_id " copilotHwnd, , 2)
    Sleep COPILOT_WEB_TAB_SETTLE_MS
    try {
        uia := UIA_Browser("ahk_id " copilotHwnd)
        Sleep 80
        CopilotWeb_FocusComposer(uia, false)
    } catch {
        return 0
    }
    WinActivate("ahk_id " copilotHwnd)
    WinWaitActive("ahk_id " copilotHwnd, , 2)
    Sleep 150
    if (optionalPromptText != "")
        InsertText(optionalPromptText)
    else
        ClipAngel_SendTopListItem(copilotHwnd)
    Sleep 250
    CopilotWeb_PlayFocusedChime()
    if (autoSubmit) {
        Sleep 1000
        try {
            uia := UIA_Browser("ahk_id " copilotHwnd)
            CopilotWeb_TrySubmit(uia)
        } catch {
            Send "{Enter}"
        }
    }
    if (copilotHwnd)
        CopilotWeb_CacheHwnd(copilotHwnd)
    return copilotHwnd
}

CopilotWeb_IsCopyResponseButton(name) {
    if (!name)
        return false
    if (InStr(name, "Copy prompt", false) || InStr(name, "Copiar prompt", false))
        return false
    for n in COPILOT_COPY_RESPONSE_NAMES {
        if (name = n)
            return true
    }
    if (InStr(name, "Copy Response", false) = 1)
        return true
    return false
}

CopilotWeb_GetCopyButtonsArray(uia) {
    out := []
    if (!IsObject(uia))
        return out
    try {
        allButtons := uia.FindAll({ Type: "Button" })
        for button in allButtons {
            if (CopilotWeb_IsCopyResponseButton(button.Name))
                out.Push(button)
        }
    } catch {
    }
    return out
}

CopilotWeb_GetLastCopyButton(uia) {
    arr := CopilotWeb_GetCopyButtonsArray(uia)
    if (arr.Length = 0)
        return 0
    lastEl := 0
    lastTop := ""
    for btn in arr {
        try {
            br := btn.BoundingRectangle
        } catch {
            continue
        }
        if (!IsObject(br))
            continue
        if ((br.r - br.l) <= 0 || (br.b - br.t) <= 0)
            continue
        if (lastEl = 0 || br.t >= lastTop) {
            lastEl := btn
            lastTop := br.t
        }
    }
    return lastEl ? lastEl : arr[arr.Length]
}

CopilotWeb_GetLowestButtonByNames(uia, names) {
    if (!IsObject(uia))
        return 0
    candidates := []
    try {
        allButtons := uia.FindAll({ Type: "Button" })
        for button in allButtons {
            for n in names {
                if (button.Name = n || InStr(button.Name, n, false) = 1) {
                    candidates.Push(button)
                    break
                }
            }
        }
    } catch {
    }
    if (candidates.Length = 0)
        return 0
    lastEl := 0
    lastTop := ""
    for btn in candidates {
        try {
            br := btn.BoundingRectangle
        } catch {
            continue
        }
        if (!IsObject(br))
            continue
        if (lastEl = 0 || br.t >= lastTop) {
            lastEl := btn
            lastTop := br.t
        }
    }
    return lastEl
}

CopilotWeb_FindPauseResumeButton(uia, which) {
    names := (which = "Resume") ? COPILOT_TTS_RESUME_NAMES : COPILOT_TTS_PAUSE_NAMES
    return CopilotWeb_GetLowestButtonByNames(uia, names)
}

CopilotWeb_CopyLastMessageToClipboard(options := "", copilotHwnd := 0) {
    restoreWindow := (options = "" || !options.HasProp("restoreWindow")) ? true : options.restoreWindow
    playChimeAndNotify := (options = "" || !options.HasProp("playChimeAndNotify")) ? true : options.playChimeAndNotify
    alreadyActive := (options != "" && options.HasProp("alreadyActive")) ? options.alreadyActive : false
    try {
        SetTitleMatchMode(2)
        if (!copilotHwnd)
            copilotHwnd := GetCopilotWebWindowHwnd()
        if (!copilotHwnd)
            return false
        if (!alreadyActive) {
            if !CopilotWeb_ActivateWindow(copilotHwnd)
                return false
            Sleep COPILOT_WEB_TAB_SETTLE_MS
        }
        uia := alreadyActive ? UIA_Browser() : UIA_Browser("ahk_id " copilotHwnd)
        Sleep COPILOT_WEB_UIA_SETTLE_MS
        Send "^{End}"
        Sleep COPILOT_WEB_SCROLL_SETTLE_MS
        copyBtn := CopilotWeb_GetLastCopyButton(uia)
        if (!copyBtn)
            return false
        A_Clipboard := ""
        copyBtn.Click()
        if !ClipWait(2)
            return false
        if (playChimeAndNotify) {
            try ScriptSoundPlay(A_ScriptDir . "\assets\sounds\copy.wav")
            CopilotWeb_Notify("Copied!", 800, 24)
        }
        if (restoreWindow)
            Send "!{Tab}"
        else
            CopilotWeb_FocusComposerForHwnd(copilotHwnd, false)
        return true
    } catch {
        return false
    }
}

CopilotWeb_CopyLastMessageWithRetry(options := "", copilotHwnd := 0, maxRetries := COPILOT_WEB_COPY_MAX_RETRIES) {
    baseDelay := COPILOT_WEB_COPY_RETRY_SLEEP_MS
    loop maxRetries {
        if (CopilotWeb_CopyLastMessageToClipboard(options, copilotHwnd))
            return true
        if (A_Index < maxRetries)
            Sleep baseDelay * (1 << (A_Index - 1))
    }
    return false
}

CopilotBackgroundSetTimer(task, callback, periodMs := COPILOT_WEB_ASYNC_POLL_MS) {
    CopilotBackgroundStopTimer(task)
    task.TimerCallback := callback
    SetTimer(task.TimerCallback, periodMs)
}

CopilotBackgroundStopTimer(task) {
    cb := ""
    try cb := task.TimerCallback
    catch
        cb := ""
    if (cb)
        SetTimer(cb, 0)
    try task.TimerCallback := ""
}

CopilotWeb_VerifyStreamingStopped(copilotHwnd) {
    loop COPILOT_WEB_STREAM_GONE_LOOPS {
        Sleep COPILOT_WEB_STREAM_GONE_VERIFY_MS
        root := CopilotWeb_ReadRootFromHwnd(copilotHwnd)
        if (!root)
            return true
        try {
            if (!CopilotWeb_FindStopGenerating(root))
                continue
            return false
        } catch {
            return true
        }
    }
    return true
}

CopilotWeb_MonitorStreamingTransition(task, onCompleteCallback) {
    task.RetryCount++
    if (task.RetryCount > task.MaxRetries) {
        CopilotBackgroundStopTimer(task)
        return "timeout"
    }
    if (!task.CopilotHwnd || !WinExist("ahk_id " task.CopilotHwnd)) {
        CopilotBackgroundStopTimer(task)
        return "unavailable"
    }
    uia := CopilotWeb_ReadRootFromHwnd(task.CopilotHwnd)
    if (!uia)
        return "unavailable"
    if (CopilotWeb_FindStopGenerating(uia)) {
        task.ButtonEverFound := true
        return "streaming"
    }
    if (!task.ButtonEverFound)
        return "waiting"
    if (!CopilotWeb_VerifyStreamingStopped(task.CopilotHwnd))
        return "streaming"
    CopilotBackgroundStopTimer(task)
    onCompleteCallback.Call()
    return "completed"
}

CopilotWeb_TriggerReadAloud(copyFirst := true, options := "") {
    return (CopilotAsyncReadAloud(copyFirst, options)).Start()
}

; --- Async read aloud (#!+o, D2C R) ---
class CopilotAsyncReadAloud {
    __New(copyFirst := true, options := "") {
        this.CopyFirst := copyFirst
        this.OriginalHwnd := (options != "" && options.HasProp("originalHwnd")) ? options.originalHwnd : 0
        this.CopilotHwnd := (options != "" && options.HasProp("copilotHwnd")) ? options.copilotHwnd : 0
        this.AlreadyActive := (options != "" && options.HasProp("alreadyActive")) ? options.alreadyActive : false
        this.Uia := 0
        this.CopyRetryCount := 0
    }

    Start() {
        if (!this.OriginalHwnd)
            try this.OriginalHwnd := WinExist("A")
        SetTitleMatchMode(2)
        if (!this.CopilotHwnd)
            this.CopilotHwnd := GetCopilotWebWindowHwnd()
        if (!this.CopilotHwnd) {
            CopilotWeb_Notify("Read aloud failed – Copilot is not open", 1800, 22)
            return false
        }
        SetTimer(this.Run.Bind(this), -1)
        return true
    }

    Run(*) {
        try {
            if (!WinActive("ahk_id " this.CopilotHwnd)) {
                StandardLoadingBar_Show("⏳ Switching to Copilot…", BANNER_ACCENT_INTERMEDIATE)
                if !CopilotWeb_ActivateWindow(this.CopilotHwnd) {
                    StandardLoadingBar_Hide(0)
                    this.Fail()
                    return
                }
            }
            try this.Uia := UIA_Browser("ahk_id " this.CopilotHwnd)
            catch {
                StandardLoadingBar_Hide(0)
                this.Fail()
                return
            }
            Sleep COPILOT_WEB_UIA_SETTLE_MS
            StandardLoadingBar_Hide(0)

            pauseBtn := CopilotWeb_FindPauseResumeButton(this.Uia, "Pause")
            if (pauseBtn) {
                try pauseBtn.Click()
                CopilotWeb_Notify("Paused", 800, 24)
                this.RestoreFocus()
                return
            }
            resumeBtn := CopilotWeb_FindPauseResumeButton(this.Uia, "Resume")
            if (resumeBtn) {
                try resumeBtn.Click()
                CopilotWeb_Notify("Resumed", 800, 24)
                this.RestoreFocus()
                return
            }

            Send "^{End}"
            Sleep COPILOT_WEB_SCROLL_SETTLE_MS
            if (this.CopyFirst) {
                copyOpt := { restoreWindow: false, playChimeAndNotify: false, alreadyActive: true }
                if (!CopilotWeb_CopyLastMessageToClipboard(copyOpt, this.CopilotHwnd))
                    this.CopyRetryCount := COPILOT_WEB_COPY_MAX_RETRIES
            }

            readBtn := CopilotWeb_GetLowestButtonByNames(this.Uia, COPILOT_READ_ALOUD_NAMES)
            if (!readBtn) {
                moreBtn := CopilotWeb_GetLowestButtonByNames(this.Uia, COPILOT_MORE_OPTIONS_NAMES)
                if (moreBtn) {
                    try moreBtn.Click()
                    Sleep 200
                    try this.Uia := UIA_Browser("ahk_id " this.CopilotHwnd)
                    for n in COPILOT_READ_ALOUD_NAMES {
                        try {
                            item := this.Uia.FindFirst({ Name: n, Type: UIA_Copilot_ControlType_MenuItem })
                            if (item) {
                                item.Click()
                                readBtn := item
                                break
                            }
                        } catch {
                        }
                    }
                }
            } else {
                try readBtn.Click()
            }

            if (!readBtn) {
                ShowCenteredOverlay_Utils("Read aloud not available for Copilot web", 2500, BANNER_ACCENT_ERROR)
                this.RestoreFocus()
                return
            }

            CopilotWeb_Notify(this.CopyFirst ? "Copied & Reading aloud" : "Reading aloud", 800, 24)
            this.RestoreFocus()
        } catch {
            this.Fail()
        }
    }

    RestoreFocus() {
        orig := this.OriginalHwnd
        if (orig && orig != this.CopilotHwnd && WinExist("ahk_id " orig)) {
            try {
                WinActivate("ahk_id " orig)
                WinWaitActive("ahk_id " orig, , 0.5)
            } catch {
            }
        } else if (this.CopilotHwnd && WinActive("ahk_id " this.CopilotHwnd))
            CopilotWeb_FocusComposerForHwnd(this.CopilotHwnd, false)
    }

    Fail() {
        StandardLoadingBar_Hide(0)
        CopilotWeb_Notify("Read aloud failed – Copilot UI not ready", 2000, 22)
        this.RestoreFocus()
    }
}

; --- TTS from selection (GeminiAsyncTTS / CopilotAsyncTTS; no global hotkey — Win+Alt+Shift+7 freed) ---
class CopilotAsyncTTS {
    static TTSPrompt :=
        "Repeat the following text exactly as it is. Do not add any introduction, explanation, or markdown formatting. Just output the text itself:`n`n"
    static PostStreamingDelayMs := 600

    __New() {
        this.OriginalHwnd := 0
        this.CopilotHwnd := 0
        this.RetryCount := 0
        this.MaxRetries := COPILOT_WEB_ASYNC_MAX_RETRIES
        this.ButtonEverFound := false
    }

    Start() {
        this.OriginalHwnd := WinExist("A")
        if !this.OriginalHwnd
            return
        StandardLoadingBar_Show("⏳ Loading…", BANNER_ACCENT_INTERMEDIATE)
        A_Clipboard := ""
        Send "^c"
        if !ClipWait(2) {
            StandardLoadingBar_Hide(0)
            return
        }
        SetTitleMatchMode(2)
        this.CopilotHwnd := GetCopilotWebWindowHwnd()
        if !this.CopilotHwnd {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ Copilot is not open", 2000, BANNER_ACCENT_ERROR)
            return
        }
        if !CopilotWeb_ActivateWindow(this.CopilotHwnd) {
            StandardLoadingBar_Hide(0)
            return
        }
        try uia := UIA_Browser("ahk_id " this.CopilotHwnd)
        catch {
            StandardLoadingBar_Hide(0)
            return
        }
        Sleep 300
        if (!CopilotWeb_FocusComposer(uia, false)) {
            StandardLoadingBar_Hide(0)
            return
        }
        A_Clipboard := CopilotAsyncTTS.TTSPrompt . A_Clipboard
        Sleep 100
        Send("^a")
        Sleep 500
        Send("^v")
        Sleep 500
        Send("{Enter}")
        Sleep 300
        origHwnd := this.OriginalHwnd
        try {
            if WinExist("ahk_id " origHwnd) {
                WinActivate("ahk_id " origHwnd)
                WinWaitActive("ahk_id " origHwnd, , 1)
            }
        } catch {
        }
        this.RetryCount := 0
        CopilotBackgroundSetTimer(this, this.CheckCompletion.Bind(this), COPILOT_WEB_ASYNC_POLL_MS)
    }

    CheckCompletion() {
        state := CopilotWeb_MonitorStreamingTransition(this, this.OnStreamingCompleted.Bind(this))
        if (state = "timeout")
            StandardLoadingBar_Hide(0)
    }

    OnStreamingCompleted() {
        StandardLoadingBar_Hide(0)
        try ScriptSoundPlay(A_ScriptDir . "\assets\sounds\gemini-completion.wav")
        catch {
        }
        Sleep(CopilotAsyncTTS.PostStreamingDelayMs)
        CopilotWeb_TriggerReadAloud(false, { originalHwnd: this.OriginalHwnd, copilotHwnd: this.CopilotHwnd })
    }
}

; --- Pronunciation lookup (#!+8) ---
class CopilotAsyncLookup {
    __New(lang, selectedText := "", preCopiedText := "") {
        this.Lang := lang
        this.PreCopiedText := preCopiedText
        this.OriginalHwnd := 0
        this.CopilotHwnd := 0
        this.RetryCount := 0
        this.MaxRetries := COPILOT_WEB_ASYNC_MAX_RETRIES
        this.ButtonEverFound := false
        if (selectedText != "")
            this.PreCopiedText := selectedText
    }

    Start() {
        this.OriginalHwnd := WinExist("A")
        if !this.OriginalHwnd
            return
        StandardLoadingBar_Show("⏳ Loading…", BANNER_ACCENT_INTERMEDIATE)
        clipOk := false
        if (this.PreCopiedText != "") {
            A_Clipboard := this.PreCopiedText
            clipOk := true
        } else {
            A_Clipboard := ""
            clipOk := TryCopySelectionToClipboard_QuickLookAware()
        }
        if !clipOk {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ Copy failed (no clipboard text).", 2400, BANNER_ACCENT_ERROR)
            return
        }
        SetTitleMatchMode(2)
        this.CopilotHwnd := GetCopilotWebWindowHwnd()
        if !this.CopilotHwnd {
            StandardLoadingBar_Hide(0)
            return
        }
        if !CopilotWeb_ActivateWindow(this.CopilotHwnd) {
            StandardLoadingBar_Hide(0)
            return
        }
        try uia := UIA_Browser("ahk_id " this.CopilotHwnd)
        catch {
            StandardLoadingBar_Hide(0)
            return
        }
        Sleep 300
        if (!CopilotWeb_FocusComposer(uia, false)) {
            StandardLoadingBar_Hide(0)
            return
        }
        promptName := this.Lang ? "pronunciation-lookup-" . this.Lang : "pronunciation-lookup"
        lookupLangTitle := Map("pt", "Português brasileiro", "en", "English", "de", "Deutsch")
        lead := ""
        if (this.Lang != "" && lookupLangTitle.Has(this.Lang))
            lead :=
                "Answer preamble (mandatory): The very first line of your reply must be exactly this language label (so the reader sees which lookup mode was used):`n"
                . lookupLangTitle[this.Lang]
                .
                "`nThe second line must be blank. After that, follow every instruction below—including the seven sections—with no section titles or markdown headings for those sections (this preamble is the only allowed title line).`n`n"
        searchString := lead . RTrim(GetPromptText(promptName), "`r`n")
        A_Clipboard := searchString . "`n`nContent: " . A_Clipboard
        Sleep 100
        Send("^a")
        Sleep 500
        Send("^v")
        Sleep 500
        Send("{Enter}")
        Sleep 300
        origHwnd := this.OriginalHwnd
        try {
            if WinExist("ahk_id " origHwnd) {
                WinActivate("ahk_id " origHwnd)
                WinWaitActive("ahk_id " origHwnd, , 1)
            }
        } catch {
        }
        this.RetryCount := 0
        CopilotBackgroundSetTimer(this, this.CheckCompletion.Bind(this), COPILOT_WEB_ASYNC_POLL_MS)
    }

    CheckCompletion() {
        state := CopilotWeb_MonitorStreamingTransition(this, this.OnStreamingCompleted.Bind(this))
        if (state = "timeout")
            StandardLoadingBar_Hide(0)
    }

    OnStreamingCompleted() {
        try ScriptSoundPlay(A_ScriptDir . "\assets\sounds\gemini-completion.wav")
        catch {
        }
        this.RetrieveResponse()
    }

    RetrieveResponse() {
        try WinActivate("ahk_id " this.CopilotHwnd)
        catch {
            StandardLoadingBar_Hide(0)
            return
        }
        if !WinWaitActive("ahk_exe chrome.exe", , 2) {
            StandardLoadingBar_Hide(0)
            return
        }
        copyOpt := { restoreWindow: false, playChimeAndNotify: false, alreadyActive: true }
        seqBefore := Clipboard_GetSequenceNumber()
        if !CopilotWeb_CopyLastMessageWithRetry(copyOpt, this.CopilotHwnd) {
            StandardLoadingBar_Hide(0)
            return
        }
        syncElapsed := 0
        while (syncElapsed < COPILOT_WEB_POST_COPY_SYNC_TIMEOUT_MS) {
            if (Clipboard_GetSequenceNumber() != seqBefore)
                break
            Sleep COPILOT_WEB_CLIPBOARD_POLL_MS
            syncElapsed += COPILOT_WEB_CLIPBOARD_POLL_MS
        }
        bannerText := A_Clipboard
        origHwnd := this.OriginalHwnd
        if (origHwnd && WinExist("ahk_id " origHwnd)) {
            try {
                WinActivate("ahk_id " origHwnd)
                if (!WinActive("ahk_id " origHwnd))
                    WinWaitActive("ahk_id " origHwnd, , 0.5)
            } catch {
            }
        }
        StandardLoadingBar_Hide(0)
        if (!bannerText || StrLen(Trim(bannerText)) = 0)
            return
        state := "ℹ " . bannerText
        closeNoOp(*) {
        }
        closeKeys := Map("Enter", closeNoOp, "Escape", closeNoOp, "E", closeNoOp)
        StandardLoadingBar_ShowWithKeys(state, closeKeys, 0, 0, "",
            BANNER_ACCENT_INTERMEDIATE, 600, 17, "", false,
            "[Enter] [E] [Esc] Close", true)
    }
}

; --- Shift keys.ahk helpers (same letter shortcuts as Gemini web) ---

; Cache-first #HotIf context: full UIA/URL detect on foreground change only (efficiency-canon §3–§4).
CopilotWeb_RefreshHotkeyContext(hwnd, useFull := false) {
    global g_CopilotWebHotkeyActive, g_CopilotWebCachedHwnd, g_CopilotWebCachedTitle
    if (!hwnd || !CopilotWeb_IsChromeHwnd(hwnd)) {
        g_CopilotWebHotkeyActive := false
        g_CopilotWebCachedTitle := ""
        return false
    }
    mode := useFull ? "full" : "fast"
    active := CopilotWeb_IsCopilotHwnd(hwnd, mode)
    g_CopilotWebHotkeyActive := active
    if (active) {
        g_CopilotWebCachedHwnd := hwnd
        CopilotWeb_CacheHwnd(hwnd)
        try {
            g_CopilotWebCachedTitle := WinGetTitle("ahk_id " hwnd)
        } catch {
            g_CopilotWebCachedTitle := ""
        }
    } else {
        if (g_CopilotWebCachedHwnd = hwnd)
            g_CopilotWebCachedHwnd := 0
        g_CopilotWebCachedTitle := ""
    }
    return active
}

CopilotWeb_OnForegroundChanged(hwnd) {
    if (!hwnd || !WinExist("ahk_id " hwnd)) {
        global g_CopilotWebHotkeyActive, g_CopilotWebCachedTitle
        g_CopilotWebHotkeyActive := false
        g_CopilotWebCachedTitle := ""
        return
    }
    if (!CopilotWeb_IsChromeHwnd(hwnd)) {
        global g_CopilotWebHotkeyActive, g_CopilotWebCachedTitle
        g_CopilotWebHotkeyActive := false
        g_CopilotWebCachedTitle := ""
        return
    }
    CopilotWeb_RefreshHotkeyContext(hwnd, true)
}

CopilotWeb_ForegroundHookProc(hHook, event, hwnd, idObject, idChild, idEventThread, dwmsEventTime) {
    SetTimer(CopilotWeb_OnForegroundChanged.Bind(hwnd), -0)
}

CopilotWeb_EnsureForegroundHook() {
    global g_CopilotWeb_ForegroundHookHandle, g_CopilotWeb_ForegroundHookCallback
    if (g_CopilotWeb_ForegroundHookHandle)
        return
    cb := CallbackCreate(CopilotWeb_ForegroundHookProc, "F", 7)
    h := DllCall("user32\SetWinEventHook", "UInt", 0x0003, "UInt", 0x0003, "Ptr", 0, "Ptr", cb, "UInt", 0, "UInt", 0,
        "UInt", 0, "Ptr")
    if (h) {
        g_CopilotWeb_ForegroundHookHandle := h
        g_CopilotWeb_ForegroundHookCallback := cb
    }
}

IsCopilotWebChromeActiveForHotkey() {
    CopilotWeb_EnsureForegroundHook()
    hwnd := WinExist("A")
    if (!hwnd || !CopilotWeb_IsChromeHwnd(hwnd))
        return false
    global g_CopilotWebHotkeyActive, g_CopilotWebCachedHwnd, g_CopilotWebCachedTitle
    if (g_CopilotWebHotkeyActive && hwnd = g_CopilotWebCachedHwnd) {
        try {
            title := WinGetTitle("ahk_id " hwnd)
        } catch {
            title := ""
        }
        if (title != g_CopilotWebCachedTitle)
            return CopilotWeb_RefreshHotkeyContext(hwnd, false)
        return true
    }
    return CopilotWeb_RefreshHotkeyContext(hwnd, true)
}

CopilotWeb_GetActiveUia() {
    ; Prefer ElementFromHandle for in-page Shift actions (cheaper than UIA_Browser init).
    return CopilotWeb_ReadRootFromHwnd(WinExist("A"))
}

; Bind UIA to a specific Chrome hwnd (prefer over bare UIA_Browser() in wait loops).
CopilotWeb_GetBoundUia(hwnd := 0) {
    if (!hwnd)
        hwnd := WinExist("A")
    if (!hwnd)
        return 0
    root := CopilotWeb_ReadRootFromHwnd(hwnd)
    if (root)
        return root
    try {
        return UIA_Browser("ahk_id " hwnd)
    } catch {
        return 0
    }
}

CopilotWeb_RefreshBoundUia(hwnd) {
    if (!hwnd)
        hwnd := WinExist("A")
    return CopilotWeb_GetBoundUia(hwnd)
}

; Loading Indication while mouse/UIA automation runs — always Hide in finally.
CopilotWeb_RunWithBusyBanner(message, fn, hwnd := 0) {
    if (!hwnd)
        hwnd := WinExist("A")
    StandardLoadingBar_Show(message, BANNER_ACCENT_INTERMEDIATE, { centerOnHwnd: hwnd, fontSize: 17 })
    try {
        return fn.Call()
    } finally {
        StandardLoadingBar_Hide(0)
    }
}

; After busy banner Hide — avoid clearing a same-frame error overlay.
CopilotWeb_ShowErrorAfterBanner(text, durationMs := 2200) {
    SetTimer(() => ShowCenteredOverlay_Utils(text, durationMs, BANNER_ACCENT_ERROR), -60)
}

CopilotWeb_ClickUiaElement(el) {
    if (!IsObject(el))
        return false
    try {
        if (el.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)) {
            el.InvokePattern.Invoke()
            return true
        }
    } catch {
    }
    try {
        el.Click()
        return true
    } catch {
    }
    return false
}

; Force a physical screen click (Fluent flyout items often ignore Invoke).
CopilotWeb_ClickUiaElementMouse(el) {
    if (!IsObject(el))
        return false
    saveCoord := A_CoordModeMouse
    CoordMode("Mouse", "Screen")
    try {
        try {
            pt := el.GetClickablePoint()
            if (IsObject(pt) && (pt.x || pt.y)) {
                Click(pt.x " " pt.y)
                CoordMode("Mouse", saveCoord)
                return true
            }
        } catch {
        }
        try {
            loc := el.Location
            if (IsObject(loc) && loc.w > 1 && loc.h > 1) {
                Click((loc.x + loc.w // 2) " " (loc.y + loc.h // 2))
                CoordMode("Mouse", saveCoord)
                return true
            }
        } catch {
        }
        try {
            if (el.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable)) {
                el.SelectionItemPattern.Select()
                CoordMode("Mouse", saveCoord)
                return true
            }
        } catch {
        }
        try {
            el.Click("left")
            CoordMode("Mouse", saveCoord)
            return true
        } catch {
        }
    } finally {
        try CoordMode("Mouse", saveCoord)
        catch {
        }
    }
    return false
}

CopilotWeb_HoverUiaElement(el) {
    if (!IsObject(el))
        return false
    try {
        loc := el.Location
        if (!IsObject(loc) || loc.w < 1 || loc.h < 1)
            return false
        CoordMode("Mouse", "Screen")
        MouseMove(loc.x + loc.w // 2, loc.y + loc.h // 2, 0)
        return true
    } catch {
    }
    return false
}

; Open a submenu parent when hover alone did not reveal children.
; Do not call this after hover has already opened the flyout — a second click can dismiss it.
CopilotWeb_ForceOpenSubmenuMenuItem(el) {
    if (!IsObject(el))
        return false
    try {
        if (el.GetPropertyValue(UIA.Property.IsExpandCollapsePatternAvailable)) {
            pat := el.ExpandCollapsePattern
            state := pat.ExpandCollapseState
            if (state = UIA.ExpandCollapseState.Expanded)
                return true
            if (state = UIA.ExpandCollapseState.Collapsed) {
                pat.Expand()
                return true
            }
        }
    } catch {
    }
    if (CopilotWeb_ClickUiaElementMouse(el))
        return true
    try {
        el.SetFocus()
        Sleep 40
        Send "{Right}"
        return true
    } catch {
    }
    return false
}

CopilotWeb_FindButtonByNames(uia, names) {
    if (!IsObject(uia) || !IsObject(names))
        return 0
    criteria := []
    for n in names
        criteria.Push({ Name: n, ControlType: "Button" })
    return CopilotWeb_FindFirstInUia(uia, criteria)
}

CopilotWeb_FindComposerExpandToggleButton(uia) {
    if (!IsObject(uia))
        return 0
    collapse := CopilotWeb_FindButtonByNames(uia, COPILOT_COMPOSER_COLLAPSE_NAMES)
    if (collapse)
        return collapse
    expand := CopilotWeb_FindButtonByNames(uia, COPILOT_COMPOSER_EXPAND_NAMES)
    if (expand)
        return expand
    try {
        for btn in uia.FindAll({ ControlType: "Button" }) {
            cn := ""
            name := ""
            try cn := btn.ClassName
            try name := btn.Name
            if (InStr(cn, "fai-ChatInput__expandButton") || InStr(cn, "fui-ExpandButton")
            || InStr(cn, "ExpandableChatInput__expandButton"))
                return btn
            if (InStr(name, "copilot input box", false) && (InStr(name, "Expand", false) || InStr(name, "Collapse",
                false)))
                return btn
        }
    } catch {
    }
    return 0
}

CopilotWeb_ToggleComposerFullscreen(uia := 0) {
    if (!uia)
        uia := CopilotWeb_GetActiveUia()
    if (!IsObject(uia))
        return false
    btn := CopilotWeb_FindComposerExpandToggleButton(uia)
    if (!btn)
        return false
    return CopilotWeb_ClickUiaElement(btn)
}

CopilotWeb_ToggleNavDrawer(uia := 0) {
    if (!uia)
        uia := CopilotWeb_GetActiveUia()
    if (!IsObject(uia))
        return false
    collapse := CopilotWeb_FindButtonByNames(uia, COPILOT_NAV_COLLAPSE_NAMES)
    if (collapse)
        return CopilotWeb_ClickUiaElement(collapse)
    expand := CopilotWeb_FindButtonByNames(uia, COPILOT_NAV_EXPAND_NAMES)
    if (expand)
        return CopilotWeb_ClickUiaElement(expand)
    return false
}

CopilotWeb_EnsureNavDrawerOpen(uia := 0) {
    if (!uia)
        uia := CopilotWeb_GetActiveUia()
    if (!IsObject(uia))
        return false
    if (!CopilotWeb_FindButtonByNames(uia, COPILOT_NAV_EXPAND_NAMES))
        return true
    return CopilotWeb_ToggleNavDrawer(uia)
}

CopilotWeb_ClickNewChat(uia := 0) {
    if (!uia)
        uia := CopilotWeb_GetActiveUia()
    if (!IsObject(uia))
        return false
    ; Live UI: Link Type 50005 — works collapsed or expanded. Do not EnsureNavDrawerOpen.
    criteria := []
    for n in COPILOT_NEW_CHAT_NAMES {
        criteria.Push({ Name: n, Type: 50005 })
        criteria.Push({ Name: n, ControlType: "Button" })
        criteria.Push({ Name: n, ControlType: "MenuItem" })
    }
    el := CopilotWeb_FindFirstInUia(uia, criteria)
    return el && CopilotWeb_ClickUiaElement(el)
}

CopilotWeb_FindNavSearchLink(uia) {
    if (!IsObject(uia))
        return 0
    criteria := []
    for n in COPILOT_NAV_SEARCH_NAMES {
        criteria.Push({ Name: n, Type: 50005 })
        criteria.Push({ Name: n, ControlType: "Button" })
        criteria.Push({ Name: n, ControlType: "MenuItem" })
    }
    return CopilotWeb_FindFirstInUia(uia, criteria)
}

CopilotWeb_ClickNavSearch(uia := 0) {
    hwnd := WinExist("A")
    if (!uia)
        uia := CopilotWeb_GetActiveUia()
    if (!IsObject(uia))
        return false
    el := CopilotWeb_FindNavSearchLink(uia)
    if (el)
        return CopilotWeb_ClickUiaElement(el)
    ; Link not visible (drawer fully closed) — expand, then bounded wait.
    if (!CopilotWeb_EnsureNavDrawerOpen(uia))
        return false
    deadline := A_TickCount + 400
    while (A_TickCount < deadline) {
        uia := CopilotWeb_ReadRootFromHwnd(hwnd)
        if (!IsObject(uia)) {
            Sleep 40
            continue
        }
        el := CopilotWeb_FindNavSearchLink(uia)
        if (el)
            return CopilotWeb_ClickUiaElement(el)
        Sleep 40
    }
    return false
}

CopilotWeb_DeepReasoningNameScore(name) {
    global COPILOT_DEEP_REASONING_NEEDLES
    if (name = "")
        return 0
    ; Progress / status labels are not selectable models.
    if RegExMatch(name, "i)\bthinking\b") && !RegExMatch(name, "i)\bthink\s+deeper\b")
        return 0
    best := 0
    for needle in COPILOT_DEEP_REASONING_NEEDLES {
        pat := "i)\b" . RegExReplace(needle, "\s+", "\\s+") . "\b"
        if RegExMatch(name, pat) {
            score := StrLen(needle)
            if (score > best)
                best := score
        }
    }
    return best
}

CopilotWeb_IsDeepReasoningModelName(name) {
    return CopilotWeb_DeepReasoningNameScore(name) > 0
}

; Cheap menu-open probe for wait loops — FindFirst only, no FindAll.
CopilotWeb_DeepReasoningMenuReady(uia) {
    if (!IsObject(uia))
        return false
    for needle in ["think deeper", "mais profundo", "pensar mais", "deeper"] {
        for typeSpec in [UIA_Copilot_ControlType_RadioButton, "RadioButton", UIA_Copilot_ControlType_MenuItem,
            "MenuItem"] {
            try {
                el := uia.FindFirst({ Name: needle, Type: typeSpec, matchmode: "Substring" })
                if (el)
                    return true
            } catch {
            }
        }
    }
    return false
}

CopilotWeb_FindModelSelectorButton(uia) {
    return CopilotWeb_FindFirstInUia(uia, COPILOT_MODEL_SELECTOR_CRITERIA)
}

; Visible mode label lives on a Text child under gptModeSwitcher (button Name stays "Model Selector").
CopilotWeb_GetModelSelectorLabel(btn) {
    if (!IsObject(btn))
        return ""
    try {
        for child in btn.FindAll({ Type: UIA_Copilot_ControlType_Text }) {
            try {
                t := Trim(child.Name)
                if (t != "")
                    return t
            } catch {
            }
        }
    } catch {
    }
    try {
        return Trim(btn.Name)
    } catch {
    }
    return ""
}

CopilotWeb_FindDeepReasoningMenuItem(uia) {
    if (!IsObject(uia))
        return 0
    ; Targeted FindFirst first (cheap); FindAll scoring only as fallback.
    for needle in COPILOT_DEEP_REASONING_NEEDLES {
        if (needle = "")
            continue
        for typeSpec in [UIA_Copilot_ControlType_RadioButton, "RadioButton", UIA_Copilot_ControlType_MenuItem,
            "MenuItem"] {
            try {
                el := uia.FindFirst({ Name: needle, Type: typeSpec, matchmode: "Substring" })
            } catch {
                el := 0
            }
            if (!el)
                continue
            try {
                aid := ""
                try aid := el.AutomationId
                if (aid = "gptModeSwitcher")
                    continue
                name := ""
                try name := el.Name
                if (CopilotWeb_DeepReasoningNameScore(name) <= 0)
                    continue
                try {
                    if (!el.GetPropertyValue(UIA.Property.IsEnabled))
                        continue
                } catch {
                }
                return el
            } catch {
            }
        }
    }
    bestEl := 0
    bestScore := 0
    for typeSpec in [UIA_Copilot_ControlType_RadioButton, "RadioButton", UIA_Copilot_ControlType_MenuItem, "MenuItem",
        UIA_Copilot_ControlType_ListItem, "ListItem", UIA_Copilot_ControlType_Button, "Button"] {
        try {
            items := uia.FindAll({ Type: typeSpec })
        } catch {
            continue
        }
        if !IsObject(items)
            continue
        for item in items {
            try {
                aid := ""
                try aid := item.AutomationId
                if (aid = "gptModeSwitcher")
                    continue
                name := ""
                try name := item.Name
                score := CopilotWeb_DeepReasoningNameScore(name)
                if (score <= 0)
                    continue
                try {
                    if (!item.GetPropertyValue(UIA.Property.IsEnabled))
                        continue
                } catch {
                }
                if (score > bestScore) {
                    bestScore := score
                    bestEl := item
                }
            } catch {
            }
        }
        if (bestEl)
            return bestEl
    }
    return bestEl
}

CopilotWeb_WaitForModelMenu(uia := 0, timeoutMs := 0, hwnd := 0) {
    global COPILOT_MODEL_MENU_WAIT_MS, COPILOT_MODEL_MENU_POLL_MS
    if (timeoutMs <= 0)
        timeoutMs := COPILOT_MODEL_MENU_WAIT_MS
    if (!hwnd)
        hwnd := WinExist("A")
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        cur := CopilotWeb_ReadRootFromHwnd(hwnd)
        if (!IsObject(cur))
            cur := uia
        if (IsObject(cur) && CopilotWeb_DeepReasoningMenuReady(cur))
            return true
        Sleep COPILOT_MODEL_MENU_POLL_MS
    }
    cur := CopilotWeb_ReadRootFromHwnd(hwnd)
    if (!IsObject(cur))
        cur := uia
    return IsObject(cur) && CopilotWeb_DeepReasoningMenuReady(cur)
}

; After model click — wait until label matches Deep (INI / needles) or model menu is gone.
CopilotWeb_WaitForModelSelectionSettled(hwnd := 0, timeoutMs := 400) {
    if (!hwnd)
        hwnd := WinExist("A")
    deepName := "Think deeper"
    try {
        fromIni := AiCompanionModels_GetDeep(AI_COMPANION_COPILOT)
        if (fromIni != "")
            deepName := fromIni
    } catch {
    }
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        uia := CopilotWeb_ReadRootFromHwnd(hwnd)
        if (!IsObject(uia)) {
            Sleep 80
            continue
        }
        btn := CopilotWeb_FindModelSelectorButton(uia)
        if (btn && CopilotWeb_ModelLabelMatches(CopilotWeb_GetModelSelectorLabel(btn), deepName, "deep"))
            return true
        if (!CopilotWeb_DeepReasoningMenuReady(uia) && CopilotWeb_FindSourcesButton(uia))
            return true
        Sleep 80
    }
    return true
}

; True when the model-selector label matches a configured name (Deep uses needle fallback).
CopilotWeb_ModelLabelMatches(label, modelName, role := "") {
    label := Trim(label)
    modelName := Trim(modelName)
    if (label = "" || modelName = "")
        return false
    if (InStr(label, modelName, false) || InStr(modelName, label, false))
        return true
    role := StrLower(Trim(role))
    if (role = "deep" || CopilotWeb_IsDeepReasoningModelName(modelName))
        return CopilotWeb_IsDeepReasoningModelName(label)
    return false
}

CopilotWeb_FindMenuItemByExactOrSubstringName(uia, modelName) {
    if (!IsObject(uia) || modelName = "")
        return 0
    for typeSpec in [UIA_Copilot_ControlType_RadioButton, "RadioButton", UIA_Copilot_ControlType_MenuItem,
        "MenuItem", UIA_Copilot_ControlType_ListItem, "ListItem", UIA_Copilot_ControlType_Button, "Button"] {
        try {
            el := uia.FindFirst({ Name: modelName, Type: typeSpec })
            if (el)
                return el
        } catch {
        }
        try {
            el := uia.FindFirst({ Name: modelName, Type: typeSpec, matchmode: "Substring" })
            if (el)
                return el
        } catch {
        }
    }
    return CopilotWeb_FindCapabilityByNames(uia, [modelName])
}

CopilotWeb_WaitForNamedModelMenuItem(hwnd, modelName, timeoutMs := 0) {
    global COPILOT_MODEL_MENU_WAIT_MS, COPILOT_MODEL_MENU_POLL_MS
    if (timeoutMs <= 0)
        timeoutMs := COPILOT_MODEL_MENU_WAIT_MS
    if (!hwnd)
        hwnd := WinExist("A")
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        uia := CopilotWeb_ReadRootFromHwnd(hwnd)
        if (IsObject(uia)) {
            item := CopilotWeb_FindMenuItemByExactOrSubstringName(uia, modelName)
            if (item)
                return item
            ; Deep nickname: needle scoring while menu is open.
            if (CopilotWeb_IsDeepReasoningModelName(modelName) || CopilotWeb_DeepReasoningMenuReady(uia)) {
                deepItem := CopilotWeb_FindDeepReasoningMenuItem(uia)
                if (deepItem && (CopilotWeb_IsDeepReasoningModelName(modelName) || modelName = ""))
                    return deepItem
            }
        }
        Sleep COPILOT_MODEL_MENU_POLL_MS
    }
    uia := CopilotWeb_ReadRootFromHwnd(hwnd)
    if !IsObject(uia)
        return 0
    item := CopilotWeb_FindMenuItemByExactOrSubstringName(uia, modelName)
    if (item)
        return item
    if (CopilotWeb_IsDeepReasoningModelName(modelName))
        return CopilotWeb_FindDeepReasoningMenuItem(uia)
    return 0
}

; Select any Copilot web model by UIA-visible name (opens model menu if needed).
CopilotWeb_SelectModelByName(modelName, uia := 0) {
    modelName := Trim(modelName)
    if (modelName = "")
        return false
    hwnd := WinExist("A")
    if (!uia)
        uia := CopilotWeb_GetBoundUia(hwnd)
    if (!IsObject(uia))
        return false
    btn := CopilotWeb_FindModelSelectorButton(uia)
    if (!btn)
        return false
    label := CopilotWeb_GetModelSelectorLabel(btn)
    if (CopilotWeb_ModelLabelMatches(label, modelName))
        return true
    if (!CopilotWeb_ClickUiaElement(btn))
        return false
    item := CopilotWeb_WaitForNamedModelMenuItem(hwnd, modelName)
    if (!item) {
        ; Fallback: deep needle path when waiting for Deep / Think deeper.
        if (CopilotWeb_WaitForModelMenu(uia, 0, hwnd) && CopilotWeb_IsDeepReasoningModelName(modelName)) {
            uia := CopilotWeb_RefreshBoundUia(hwnd)
            item := CopilotWeb_FindDeepReasoningMenuItem(uia)
        }
    }
    if (item && CopilotWeb_ClickUiaElement(item))
        return true
    Send "{Escape}"
    CopilotWeb_ShowErrorAfterBanner("❌ Model not found: " . modelName)
    return false
}

; Shift+A: Think deeper then Generate an image, paste bosch-brand-image, strip human reminders.
CopilotWeb_ShiftArt() {
    CopilotWeb_OpenModelSelector()
    CopilotWeb_WaitForModelSelectionSettled()
    if !CopilotWeb_ClickAddCapability(COPILOT_CAPABILITY_IMAGE_NAMES)
        return false
    promptText := GetPromptText("bosch-brand-image")
    if (InStr(promptText, "[PROMPT FILE MISSING:"))
        return false
    stripped := StripPromptHumanReminders(promptText)
    if (stripped != "")
        promptText := stripped
    root := CopilotWeb_ReadRootFromHwnd(WinExist("A"))
    if (IsObject(root))
        CopilotWeb_FocusComposer(root, false)
    return ReplaceFocusedEditWithText(promptText)
}

; Select configured Deep model (INI), with Think-deeper needle fallback.
CopilotWeb_OpenModelSelector(uia := 0) {
    deepName := "Think deeper"
    try {
        fromIni := AiCompanionModels_GetDeep(AI_COMPANION_COPILOT)
        if (fromIni != "")
            deepName := fromIni
    } catch {
    }
    return CopilotWeb_SelectModelByName(deepName, uia)
}

CopilotWeb_FindSourcesButton(uia) {
    return CopilotWeb_FindFirstInUia(uia, COPILOT_SOURCES_BUTTON_CRITERIA)
}

CopilotWeb_IsSourcesMenuOpen(uia) {
    if (!IsObject(uia))
        return false
    return !!CopilotWeb_FindFirstInUia(uia, COPILOT_SOURCES_MENU_MARKERS)
}

CopilotWeb_OpenSourcesMenu(uia := 0) {
    if (!uia)
        uia := CopilotWeb_GetBoundUia()
    if (!IsObject(uia))
        return false
    if (CopilotWeb_IsSourcesMenuOpen(uia))
        return true
    btn := CopilotWeb_FindSourcesButton(uia)
    if (!btn)
        return false
    if (!CopilotWeb_ClickUiaElement(btn))
        return false
    return true
}

CopilotWeb_EnsureSourcesMenuOpen(&uia) {
    hwnd := WinExist("A")
    if (!IsObject(uia))
        uia := CopilotWeb_GetBoundUia(hwnd)
    if (!IsObject(uia))
        return false
    if (CopilotWeb_IsSourcesMenuOpen(uia))
        return true
    if (!CopilotWeb_OpenSourcesMenu(uia))
        return false
    deadline := A_TickCount + 3000
    while (A_TickCount < deadline) {
        uia := CopilotWeb_RefreshBoundUia(hwnd)
        if (!IsObject(uia))
            return false
        if (CopilotWeb_IsSourcesMenuOpen(uia))
            return true
        Sleep 80
    }
    return CopilotWeb_IsSourcesMenuOpen(uia)
}

CopilotWeb_ClickSourcesCapability(automationId, nameSubstrings, uia := 0) {
    if (!uia)
        uia := CopilotWeb_GetBoundUia()
    if (!IsObject(uia))
        return false
    if (!CopilotWeb_EnsureSourcesMenuOpen(&uia))
        return false
    criteria := []
    if (automationId != "") {
        criteria.Push({ AutomationId: automationId, ControlType: "MenuItem" })
        criteria.Push({ AutomationId: automationId, ControlType: "Button" })
        criteria.Push({ AutomationId: automationId })
    }
    if (IsObject(nameSubstrings)) {
        for n in nameSubstrings {
            criteria.Push({ Name: n, ControlType: "MenuItem", matchmode: "Substring" })
            criteria.Push({ Name: n, ControlType: "Button", matchmode: "Substring" })
            criteria.Push({ Name: n, Type: UIA_Copilot_ControlType_ListItem, matchmode: "Substring" })
        }
    }
    el := CopilotWeb_FindFirstInUia(uia, criteria)
    if (!el)
        return false
    if (CopilotWeb_ClickUiaElementMouse(el))
        return true
    return CopilotWeb_ClickUiaElement(el)
}

; Broad finder for capability / submenu rows (MenuItem, Button, ListItem).
CopilotWeb_FindCapabilityByNames(uia, names) {
    if (!IsObject(uia) || !IsObject(names))
        return 0
    for needle in names {
        if (needle = "")
            continue
        for typeSpec in [UIA_Copilot_ControlType_MenuItem, "MenuItem", UIA_Copilot_ControlType_Button, "Button",
            UIA_Copilot_ControlType_ListItem, "ListItem"] {
            try {
                el := uia.FindFirst({ Name: needle, Type: typeSpec, matchmode: "Substring" })
                if (el)
                    return el
            } catch {
            }
            try {
                el := uia.FindFirst({ Name: needle, ControlType: typeSpec, matchmode: "Substring" })
                if (el)
                    return el
            } catch {
            }
        }
        ; Name-only substring (emoji-prefixed Fluent rows).
        try {
            el := uia.FindFirst({ Name: needle, matchmode: "Substring" })
            if (el)
                return el
        } catch {
        }
    }
    return 0
}

CopilotWeb_FindMenuItemByNameNeedles(uia, nameNeedles) {
    return CopilotWeb_FindCapabilityByNames(uia, nameNeedles)
}

CopilotWeb_WaitSourcesMenuClosed(hwnd, timeoutMs := 1200) {
    if (!hwnd)
        hwnd := WinExist("A")
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        root := CopilotWeb_ReadRootFromHwnd(hwnd)
        if (IsObject(root) && !CopilotWeb_IsSourcesMenuOpen(root))
            return true
        Sleep 40
    }
    return false
}

; True when an engaged capability marker is visible and sources menu is closed.
CopilotWeb_CapabilityEngaged(hwnd, engageNames, automationIds := 0, timeoutMs := 1200) {
    if (!hwnd)
        hwnd := WinExist("A")
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        root := CopilotWeb_ReadRootFromHwnd(hwnd)
        if (!IsObject(root)) {
            Sleep 40
            continue
        }
        if (CopilotWeb_IsSourcesMenuOpen(root)) {
            Sleep 40
            continue
        }
        if (IsObject(automationIds)) {
            for aid in automationIds {
                if (aid = "")
                    continue
                try {
                    el := root.FindFirst({ AutomationId: aid })
                    if (el)
                        return true
                } catch {
                }
            }
        }
        if (CopilotWeb_FindCapabilityByNames(root, engageNames))
            return true
        Sleep 40
    }
    return false
}

CopilotWeb_ImageCapabilityEngaged(hwnd := 0, timeoutMs := 1200) {
    return CopilotWeb_CapabilityEngaged(hwnd, COPILOT_CAPABILITY_IMAGE_ENGAGE_NAMES, [COPILOT_CAPABILITY_IMAGE_AID],
    timeoutMs)
}

CopilotWeb_CapabilitySubmenuDowns(nameNeedles) {
    if (!IsObject(nameNeedles))
        return COPILOT_CAP_SUBMENU_DOWN_IMAGE
    for n in nameNeedles {
        if InStr(n, "Research", false) || InStr(n, "Pesquis", false) || InStr(n, "topic", false) || InStr(n, "tópico",
            false)
            return COPILOT_CAP_SUBMENU_DOWN_RESEARCH
        if InStr(n, "image", false) || InStr(n, "imagem", false)
            return COPILOT_CAP_SUBMENU_DOWN_IMAGE
    }
    return COPILOT_CAP_SUBMENU_DOWN_IMAGE
}

; Keyboard path: Fluent nested menus (hover/UIA Invoke often fail to activate the row).
CopilotWeb_ClickAddCapabilityViaKeyboard(nameNeedles, hwnd := 0) {
    if (!hwnd)
        hwnd := WinExist("A")
    root := CopilotWeb_ReadRootFromHwnd(hwnd)
    if (!IsObject(root))
        root := CopilotWeb_GetBoundUia(hwnd)
    if (!CopilotWeb_EnsureSourcesMenuOpen(&root))
        return false
    Sleep 120
    loop COPILOT_CAP_MENU_DOWN_TO_ADD_CAP {
        Send "{Down}"
        Sleep 35
    }
    Sleep 60
    Send "{Right}"
    Sleep 150
    downs := CopilotWeb_CapabilitySubmenuDowns(nameNeedles)
    loop downs {
        Send "{Down}"
        Sleep 35
    }
    Sleep 50
    Send "{Enter}"
    return true
}

CopilotWeb_OpenAddCapabilitiesSubmenu(addCap, nameNeedles, hwnd) {
    if (!IsObject(addCap))
        return 0
    CopilotWeb_HoverUiaElement(addCap)
    Sleep 180
    item := CopilotWeb_WaitForCapabilityByNames(nameNeedles, 600, hwnd)
    if (item)
        return item
    expanded := false
    try {
        if (addCap.GetPropertyValue(UIA.Property.IsExpandCollapsePatternAvailable)) {
            pat := addCap.ExpandCollapsePattern
            state := pat.ExpandCollapseState
            if (state = UIA.ExpandCollapseState.Expanded)
                expanded := true
            else if (state = UIA.ExpandCollapseState.Collapsed) {
                pat.Expand()
                expanded := true
            }
        }
    } catch {
    }
    if (!expanded) {
        try {
            addCap.SetFocus()
            Sleep 40
            Send "{Right}"
        } catch {
        }
    }
    item := CopilotWeb_WaitForCapabilityByNames(nameNeedles, 900, hwnd)
    if (item)
        return item
    if (!CopilotWeb_ClickUiaElementMouse(addCap))
        return 0
    Sleep 120
    return CopilotWeb_WaitForCapabilityByNames(nameNeedles, 2000, hwnd)
}

CopilotWeb_WaitForCapabilityByNames(nameNeedles, timeoutMs := 2000, hwnd := 0) {
    if (!hwnd)
        hwnd := WinExist("A")
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        root := CopilotWeb_ReadRootFromHwnd(hwnd)
        if (!IsObject(root)) {
            try root := UIA_Browser("ahk_id " hwnd)
            catch
                root := 0
        }
        if (IsObject(root)) {
            item := CopilotWeb_FindCapabilityByNames(root, nameNeedles)
            if (item)
                return item
        }
        Sleep 40
    }
    return 0
}

CopilotWeb_WaitForMenuItemByNameNeedles(nameNeedles, timeoutMs := 2000, hwnd := 0) {
    return CopilotWeb_WaitForCapabilityByNames(nameNeedles, timeoutMs, hwnd)
}

; Resolve AID / engage markers for built-in capability presets.
CopilotWeb_CapabilityPresetMeta(nameNeedles) {
    meta := { AutomationId: "", EngageNames: nameNeedles }
    if (!IsObject(nameNeedles))
        return meta
    for n in nameNeedles {
        if InStr(n, "image", false) || InStr(n, "imagem", false) {
            meta.AutomationId := COPILOT_CAPABILITY_IMAGE_AID
            meta.EngageNames := COPILOT_CAPABILITY_IMAGE_ENGAGE_NAMES
            return meta
        }
    }
    return meta
}

; Sources -> keyboard Add capabilities path, then UIA hover/click, then AID. QC until engaged.
CopilotWeb_ClickAddCapability(nameNeedles, uia := 0) {
    hwnd := WinExist("A")
    meta := CopilotWeb_CapabilityPresetMeta(nameNeedles)
    engageNames := meta.EngageNames
    automationId := meta.AutomationId
    aids := automationId != "" ? [automationId] : 0

    loop 2 {
        attempt := A_Index
        if (attempt > 1) {
            Send "{Escape}"
            Sleep 100
        }

        ; Strategy K: keyboard Down × N → Right → Down × M → Enter (matches live Fluent menu).
        if (CopilotWeb_ClickAddCapabilityViaKeyboard(nameNeedles, hwnd)) {
            if (CopilotWeb_WaitSourcesMenuClosed(hwnd, 1200) || CopilotWeb_CapabilityEngaged(hwnd, engageNames, aids))
                return true
        }

        ; Strategy B: hover Add capabilities then physical click target row.
        root := CopilotWeb_ReadRootFromHwnd(hwnd)
        if (!IsObject(root)) {
            try root := UIA_Browser("ahk_id " hwnd)
            catch
                root := 0
        }
        if (!IsObject(root))
            continue
        if (!CopilotWeb_EnsureSourcesMenuOpen(&root))
            continue
        Sleep 80
        addCap := CopilotWeb_FindCapabilityByNames(root, COPILOT_ADD_CAPABILITIES_NAMES)
        if (!addCap) {
            root := CopilotWeb_ReadRootFromHwnd(hwnd)
            if (IsObject(root))
                addCap := CopilotWeb_FindCapabilityByNames(root, COPILOT_ADD_CAPABILITIES_NAMES)
        }
        if (!addCap)
            continue
        item := CopilotWeb_OpenAddCapabilitiesSubmenu(addCap, nameNeedles, hwnd)
        if (!item)
            continue
        ; Keep Add capabilities hovered open — move and physical-click Generate an image.
        if (!(CopilotWeb_ClickUiaElementMouse(item) || CopilotWeb_ClickUiaElement(item)))
            continue
        if (CopilotWeb_WaitSourcesMenuClosed(hwnd, 1200) || CopilotWeb_CapabilityEngaged(hwnd, engageNames, aids))
            return true

        ; Strategy A: legacy AutomationId row under sources (if present).
        if (automationId != "") {
            root := CopilotWeb_ReadRootFromHwnd(hwnd)
            if (CopilotWeb_ClickSourcesCapability(automationId, nameNeedles, root)) {
                if (CopilotWeb_WaitSourcesMenuClosed(hwnd, 1200) || CopilotWeb_CapabilityEngaged(hwnd, engageNames,
                    aids))
                    return true
            }
        }
    }
    return false
}

CopilotWeb_ScrollFeedToBottom(hwnd := 0) {
    if (!hwnd)
        hwnd := WinExist("A")
    if (!hwnd)
        return false
    try {
        uia := ChromeChat_ScrollFeedToBottomFast(hwnd)
        if (!IsObject(uia))
            return false
        pf := CopilotWeb_FindComposer(uia)
        if (pf) {
            try pf.ScrollIntoView()
            catch {
            }
        }
        return true
    } catch {
        return false
    }
}

CopilotWeb_SendPromptFromFile(promptFilePath := "") {
    if (promptFilePath = "")
        promptFilePath := A_ScriptDir "\assets\data\Gemini_Prompt.txt"
    uia := CopilotWeb_GetActiveUia()
    if (!IsObject(uia))
        return false
    if (!CopilotWeb_FocusComposer(uia, false))
        return false
    if (!FileExist(promptFilePath))
        return false
    oldClipboard := A_Clipboard
    try {
        promptText := FileRead(promptFilePath, "UTF-8")
        if (!promptText)
            return false
        A_Clipboard := promptText
        if !ClipWait(1, 1)
            return false
        Send "^a"
        Sleep 50
        Send "^v"
        Sleep 100
        A_Clipboard := oldClipboard
        Sleep 400
        CopilotWeb_TrySubmit(uia)
        return true
    } catch {
        try A_Clipboard := oldClipboard
        catch {
        }
        return false
    }
}

CopilotWeb_ShiftCopyLastMessage() {
    return CopilotWeb_CopyLastMessageToClipboard({ restoreWindow: false, playChimeAndNotify: true, alreadyActive: true })
}

CopilotWeb_ShiftReadAloud() {
    return CopilotWeb_TriggerReadAloud(false, { alreadyActive: true })
}

CopilotWeb_IsVoiceChatActive(uia := 0) {
    if (!uia)
        uia := CopilotWeb_GetActiveUia()
    if (!IsObject(uia))
        return false
    return !!CopilotWeb_FindButtonByNames(uia, COPILOT_VOICE_END_NAMES)
}

CopilotWeb_ToggleVoiceChat(uia := 0) {
    if (!uia)
        uia := CopilotWeb_GetActiveUia()
    if (!IsObject(uia))
        return false
    endBtn := CopilotWeb_FindButtonByNames(uia, COPILOT_VOICE_END_NAMES)
    if (endBtn) {
        if (!CopilotWeb_ClickUiaElement(endBtn))
            return false
        CopilotWeb_Notify("Voice chat ended", 800, 24)
        CopilotWeb_ReturnToComposer()
        return true
    }
    startBtn := CopilotWeb_FindButtonByNames(uia, COPILOT_VOICE_START_NAMES)
    if (!startBtn)
        return false
    try startBtn.ScrollIntoView()
    catch {
    }
    if (!CopilotWeb_ClickUiaElement(startBtn))
        return false
    CopilotWeb_Notify("Voice chat started", 800, 24)
    return true
}

CopilotWeb_PlayCompletionChime() {
    try {
        static lastTick := 0
        if (A_TickCount - lastTick < 1500)
            return
        lastTick := A_TickCount
        ScriptSoundPlay(A_ScriptDir . "\assets\sounds\gemini-completion.wav")
    } catch {
    }
}

CopilotWeb_WaitForGenerationComplete(timeout := 300000) {
    hwnd := WinExist("A")
    if (!hwnd || !CopilotWeb_IsCopilotHwnd(hwnd, "fast"))
        return
    uia := CopilotWeb_ReadRootFromHwnd(hwnd)
    if (!uia)
        return
    deadline := (timeout > 0) ? (A_TickCount + timeout) : 0
    found := false
    while (timeout <= 0 || A_TickCount < deadline) {
        if (CopilotWeb_FindStopGenerating(uia)) {
            found := true
            break
        }
        Sleep 250
        uia := CopilotWeb_ReadRootFromHwnd(hwnd)
        if (!uia)
            return
    }
    if (!found)
        return
    while (timeout <= 0 || A_TickCount < deadline) {
        while (CopilotWeb_FindStopGenerating(uia)) {
            Sleep 250
            uia := CopilotWeb_ReadRootFromHwnd(hwnd)
            if (!uia)
                return
        }
        if (CopilotWeb_VerifyStreamingStopped(hwnd)) {
            CopilotWeb_PlayCompletionChime()
            return
        }
    }
}

CopilotHotkey_ShowPronunciationLanguagePicker(selectedText) {
    onSelect(lang) {
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        if (lang != "")
            (CopilotAsyncLookup(lang, selectedText)).Start()
    }
    onTimeout() {
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        StandardLoadingBar_Show("⏳ Detecting language…", BANNER_ACCENT_INTERMEDIATE, { textWidth: 450, fontSize: 17 })
        lang := DetectLang_AhkFallback(selectedText)
        if !(lang = "pt" || lang = "en" || lang = "de")
            lang := "en"
        (CopilotAsyncLookup(lang, selectedText)).Start()
    }
    keyCallbacks := Map(
        "1", (*) => onSelect("pt"),
        "2", (*) => onSelect("en"),
        "3", (*) => onSelect("de"),
        "*Escape", (*) => onSelect("")
    )
    StandardLoadingBar_ShowWithKeys("❓ Auto-detect in 2s — press to override", keyCallbacks, 2000, 0, onTimeout,
        BANNER_ACCENT_INTERMEDIATE,
        450, 17, "", false, "[1] Portuguese  [2] English  [3] German  [Esc] Cancel", false, true)
}
