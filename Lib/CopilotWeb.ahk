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

UIA_Copilot_ControlType_Button := 50000
UIA_Copilot_ControlType_MenuItem := 50011

COPILOT_NAV_EXPAND_NAMES := ["Expand navigation", "Expandir navegação"]
COPILOT_NAV_COLLAPSE_NAMES := ["Collapse navigation", "Recolher navegação"]
COPILOT_NEW_CHAT_NAMES := ["New chat", "Novo chat"]
COPILOT_NAV_SEARCH_NAMES := ["Search", "Pesquisar", "Buscar"]
COPILOT_MODEL_SELECTOR_CRITERIA := [{ AutomationId: "gptModeSwitcher", ControlType: "Button" }, { Name: "Model Selector",
    ControlType: "Button" }]
COPILOT_SOURCES_BUTTON_CRITERIA := [{ Name: "Add and manage sources", ControlType: "Button" }, { Name: "Adicionar e gerenciar fontes",
    ControlType: "Button" }]
COPILOT_SOURCES_MENU_MARKERS := [{ AutomationId: "capability-id-researcher", ControlType: "MenuItem" }, { Name: "Upload images and files",
    ControlType: "MenuItem" }, { Name: "Add work content", ControlType: "MenuItem" }
]
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

; Personal rig: Gemini. Work rig: M365 Copilot web (IS_WORK_ENVIRONMENT from env.ahk).
UseCopilotWebForGlobalAI() {
    global IS_WORK_ENVIRONMENT
    return IS_WORK_ENVIRONMENT
}

GetGlobalAIProviderLabel() {
    return UseCopilotWebForGlobalAI() ? "Copilot" : "Gemini"
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

CopilotWeb_TryUrlFromAddressBar(hwnd) {
    try {
        url := UIA_Browser("ahk_id " hwnd).GetCurrentURL(true)
        return CopilotWeb_UrlIsChat(url)
    } catch {
        return false
    }
}

CopilotWeb_TryUrlFromDocument(hwnd) {
    try {
        uia := UIA_Browser("ahk_id " hwnd)
        if (!uia.IsBrowserVisible()) {
            try WinActivate("ahk_id " hwnd)
            Sleep 80
        }
        return CopilotWeb_UrlIsChat(uia.GetCurrentURL(false))
    } catch {
        return false
    }
}

CopilotWeb_TryUiaFingerprint(hwnd) {
    try {
        uia := UIA_Browser("ahk_id " hwnd)
        if (CopilotWeb_FindComposer(uia))
            return true
        for criteria in [{ AutomationId: "m365-copilot-app-layout-main" }, { AutomationId: "copilot-message-rbp-title" }, { AutomationId: "m365-chat-editor-target-element" }, { AutomationId: "gptModeSwitcher",
            ControlType: "Button" }, { Name: "Message Copilot" }, { Name: "Copilot said:" }, { Name: "Expand navigation",
                ControlType: "Button" }, { Name: "New chat", ControlType: "MenuItem" }
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
    ScriptSoundPlay(A_ScriptDir . "\sounds\gemini-focused.wav")
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
        try uia := UIA_Browser("ahk_id " hwnd)
        catch {
            ShowCenteredOverlay_Utils("❌ Error: Could not attach to Copilot window.", 2000, BANNER_ACCENT_ERROR)
            return false
        }
        if (!alreadyActive)
            CopilotWeb_WaitForComposerDiscoverable(uia)
        CopilotWeb_FocusComposer(uia, true)
        CopilotWeb_CacheHwnd(hwnd)
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
        try uia := UIA_Browser("ahk_id " hwnd)
        catch {
            StandardLoadingBar_Hide(0)
            return false
        }
        CopilotWeb_WaitForComposerDiscoverable(uia, 4000)
        CopilotWeb_FocusComposer(uia, true)
        CopilotWeb_CacheHwnd(hwnd)
        StandardLoadingBar_Hide(0)
        return true
    } catch {
        StandardLoadingBar_Hide(0)
        return false
    }
}

CopilotWeb_ComposerGetText(copilotHwnd := 0) {
    try {
        uia := copilotHwnd ? UIA_Browser("ahk_id " copilotHwnd) : UIA_Browser()
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
    else {
        Send "!v"
        Sleep 50
        Send "^!b"
    }
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
            try ScriptSoundPlay(A_ScriptDir . "\sounds\copy.wav")
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
        try {
            uia := UIA_Browser("ahk_id " copilotHwnd)
            if (!CopilotWeb_FindStopGenerating(uia))
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
    try {
        uia := UIA_Browser("ahk_id " task.CopilotHwnd)
    } catch {
        return "unavailable"
    }
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

; --- TTS from selection (#!+7) ---
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
        try ScriptSoundPlay(A_ScriptDir . "\sounds\gemini-completion.wav")
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
        try ScriptSoundPlay(A_ScriptDir . "\sounds\gemini-completion.wav")
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
    hwnd := WinExist("A")
    if (!hwnd)
        return 0
    try {
        return UIA_Browser("ahk_id " hwnd)
    } catch {
        return 0
    }
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
    criteria := []
    for n in COPILOT_NEW_CHAT_NAMES {
        criteria.Push({ Name: n, ControlType: "Button" })
        criteria.Push({ Name: n, ControlType: "MenuItem" })
    }
    el := CopilotWeb_FindFirstInUia(uia, criteria)
    return el && CopilotWeb_ClickUiaElement(el)
}

CopilotWeb_ClickNavSearch(uia := 0) {
    if (!uia)
        uia := CopilotWeb_GetActiveUia()
    if (!IsObject(uia))
        return false
    CopilotWeb_EnsureNavDrawerOpen(uia)
    Sleep 150
    try
        uia := UIA_Browser()
    catch
        return false
    criteria := []
    for n in COPILOT_NAV_SEARCH_NAMES
        criteria.Push({ Name: n, ControlType: "MenuItem", ClassName: "fai-CopilotNavItem" })
    for n in COPILOT_NAV_SEARCH_NAMES
        criteria.Push({ Name: n, ControlType: "MenuItem" })
    el := CopilotWeb_FindFirstInUia(uia, criteria)
    return el && CopilotWeb_ClickUiaElement(el)
}

CopilotWeb_OpenModelSelector(uia := 0) {
    if (!uia)
        uia := CopilotWeb_GetActiveUia()
    if (!IsObject(uia))
        return false
    el := CopilotWeb_FindFirstInUia(uia, COPILOT_MODEL_SELECTOR_CRITERIA)
    return el && CopilotWeb_ClickUiaElement(el)
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
        uia := CopilotWeb_GetActiveUia()
    if (!IsObject(uia))
        return false
    if (CopilotWeb_IsSourcesMenuOpen(uia))
        return true
    btn := CopilotWeb_FindSourcesButton(uia)
    if (!btn)
        return false
    if (!CopilotWeb_ClickUiaElement(btn))
        return false
    Sleep 250
    return true
}

CopilotWeb_EnsureSourcesMenuOpen(&uia) {
    if (!IsObject(uia))
        uia := CopilotWeb_GetActiveUia()
    if (!IsObject(uia))
        return false
    if (CopilotWeb_IsSourcesMenuOpen(uia))
        return true
    if (!CopilotWeb_OpenSourcesMenu(uia))
        return false
    deadline := A_TickCount + 3000
    while (A_TickCount < deadline) {
        try
            uia := UIA_Browser()
        catch
            return false
        if (CopilotWeb_IsSourcesMenuOpen(uia))
            return true
        Sleep 80
    }
    return CopilotWeb_IsSourcesMenuOpen(uia)
}

CopilotWeb_ClickSourcesCapability(automationId, nameSubstrings, uia := 0) {
    if (!uia)
        uia := CopilotWeb_GetActiveUia()
    if (!IsObject(uia))
        return false
    if (!CopilotWeb_EnsureSourcesMenuOpen(&uia))
        return false
    criteria := [{ AutomationId: automationId, ControlType: "MenuItem" }]
    if (IsObject(nameSubstrings)) {
        for n in nameSubstrings
            criteria.Push({ Name: n, ControlType: "MenuItem" })
    }
    el := CopilotWeb_FindFirstInUia(uia, criteria)
    return el && CopilotWeb_ClickUiaElement(el)
}

CopilotWeb_ScrollFeedToBottom(hwnd := 0) {
    if (!hwnd)
        hwnd := WinExist("A")
    if (!hwnd)
        return false
    try {
        uia := UIA_Browser("ahk_id " hwnd)
        try {
            msgs := uia.FindAll({ ClassName: "fai-CopilotMessage", matchmode: "Substring" })
            if (msgs && msgs.Length > 0) {
                try msgs[msgs.Length].ScrollIntoView()
                catch {
                }
            }
        } catch {
        }
        rw := 0
        try rw := ControlGetHwnd("Chrome_RenderWidgetHostHWND1", "ahk_id " hwnd)
        catch
            rw := 0
        if (rw) {
            try ControlSend "{Blind}^{End}", , "ahk_id " rw
            catch {
                Send "^{End}"
            }
        } else {
            Send "^{End}"
        }
        Sleep COPILOT_WEB_SCROLL_SETTLE_MS
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
        promptFilePath := A_ScriptDir "\data\Gemini_Prompt.txt"
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

CopilotWeb_PlayCompletionChime() {
    try {
        static lastTick := 0
        if (A_TickCount - lastTick < 1500)
            return
        lastTick := A_TickCount
        ScriptSoundPlay(A_ScriptDir . "\sounds\gemini-completion.wav")
    } catch {
    }
}

CopilotWeb_WaitForGenerationComplete(timeout := 300000) {
    hwnd := WinExist("A")
    if (!hwnd || !CopilotWeb_IsCopilotHwnd(hwnd, "fast"))
        return
    try
        uia := UIA_Browser("ahk_id " hwnd)
    catch
        return
    deadline := (timeout > 0) ? (A_TickCount + timeout) : 0
    found := false
    while (timeout <= 0 || A_TickCount < deadline) {
        if (CopilotWeb_FindStopGenerating(uia)) {
            found := true
            break
        }
        Sleep 250
        try
            uia := UIA_Browser("ahk_id " hwnd)
        catch
            return
    }
    if (!found)
        return
    while (timeout <= 0 || A_TickCount < deadline) {
        while (CopilotWeb_FindStopGenerating(uia)) {
            Sleep 250
            try
                uia := UIA_Browser("ahk_id " hwnd)
            catch
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

CopilotWeb_EnsureForegroundHook()