; Gemini Enterprise (Vertex AI Search / AskBosch) Chrome automation.
; Included from Utils.ahk after CopilotWeb.ahk. Requires: env.ahk, Utils.ahk, UIA.

GEMINI_ENTERPRISE_URL_NEEDLE := "vertexaisearch.cloud.google.com"
GEMINI_ENTERPRISE_TITLE_NEEDLE := "Gemini Enterprise"
GEMINI_ENTERPRISE_LAUNCH_URL :=
    "https://vertexaisearch.cloud.google.com/u/1/eu/home/cid/bcb383f1-26d8-41fd-9a55-623f7e93de92?pli=1"
GEMINI_ENTERPRISE_MENU_WAIT_MS := 2000
GEMINI_ENTERPRISE_MENU_POLL_MS := 80
GEMINI_ENTERPRISE_UIA_SETTLE_MS := 120
GEMINI_ENTERPRISE_SCROLL_SETTLE_MS := 350
GEMINI_ENTERPRISE_FIRST_LAUNCH_WAIT_MS := 2500
GEMINI_ENTERPRISE_ACTIVATE_WAIT_MS := 2000
GEMINI_ENTERPRISE_DEEP_MODEL := "3.1 Pro"
GEMINI_ENTERPRISE_ASYNC_POLL_MS := 500
GEMINI_ENTERPRISE_ASYNC_MAX_RETRIES := 60
GEMINI_ENTERPRISE_COPY_MAX_RETRIES := 3
GEMINI_ENTERPRISE_COPY_RETRY_SLEEP_MS := 400
GEMINI_ENTERPRISE_STREAM_GONE_VERIFY_MS := 200
GEMINI_ENTERPRISE_STREAM_GONE_LOOPS := 4
GEMINI_ENTERPRISE_POST_COPY_SYNC_TIMEOUT_MS := 2000
GEMINI_ENTERPRISE_CLIPBOARD_POLL_MS := 10
; Copy response button names (EN/PT). Excludes "Copy prompt" / "Copiar prompt".
GEMINI_ENTERPRISE_COPY_RESPONSE_NAMES := ["Copy response", "Copy Response", "Copy message", "Copy", "Copiar"]

global g_GeminiEnterpriseCachedHwnd := 0
global g_GeminiEnterpriseHotkeyActive := false
global g_GeminiEnterpriseCachedTitle := ""
global g_GeminiEnterprise_ForegroundHookHandle := 0
global g_GeminiEnterprise_ForegroundHookCallback := 0

; Consumer Gemini title: contains "gemini" but is not Enterprise (avoids false HotIf).
IsConsumerGeminiChromeTitle(title) {
    if (!title)
        return false
    if (!InStr(title, "gemini", false))
        return false
    return !GeminiEnterprise_TitleMatches(title)
}

GeminiEnterprise_TitleMatches(title) {
    if (!title)
        return false
    return InStr(title, GEMINI_ENTERPRISE_TITLE_NEEDLE, false)
}

GeminiEnterprise_UrlMatches(url) {
    if (!url)
        return false
    return InStr(url, GEMINI_ENTERPRISE_URL_NEEDLE, false)
}

GeminiEnterprise_IsChromeHwnd(hwnd) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    try {
        return StrLower(WinGetProcessName("ahk_id " hwnd)) = "chrome.exe"
    } catch {
        return false
    }
}

GeminiEnterprise_ReadRootFromHwnd(hwnd) {
    try
        return UIA.ElementFromHandle(hwnd)
    catch
        return 0
}

GeminiEnterprise_FindFirstInUia(uia, criteriaList) {
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

GeminiEnterprise_ClickUiaElement(el) {
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

GeminiEnterprise_TryUrlFromAddressBar(hwnd) {
    try {
        root := GeminiEnterprise_ReadRootFromHwnd(hwnd)
        if (!root)
            return false
        edit := GeminiEnterprise_FindFirstInUia(root, [{ Name: "Address and search bar", ControlType: "Edit" }, { AutomationId: "view_1012",
            ControlType: "Edit" }])
        if (!edit)
            return false
        url := edit.Value
        if (!url)
            return false
        if (!RegExMatch(url, "^https?://"))
            url := "https://" url
        return GeminiEnterprise_UrlMatches(url)
    } catch {
        return false
    }
}

GeminiEnterprise_TryUrlFromDocument(hwnd) {
    try {
        root := GeminiEnterprise_ReadRootFromHwnd(hwnd)
        if (!root)
            return false
        doc := root.FindFirst({ ControlType: "Document" })
        if (!doc)
            return false
        return GeminiEnterprise_UrlMatches(doc.Value)
    } catch {
        return false
    }
}

GeminiEnterprise_TryUiaFingerprint(hwnd) {
    try {
        uia := GeminiEnterprise_ReadRootFromHwnd(hwnd)
        if (!uia)
            return false
        for criteria in [{ AutomationId: "main-panel" }, { AutomationId: "agent-search-prosemirror-editor" }, { AutomationId: "tool-selector-menu-anchor" }, { AutomationId: "model-selector-menu-anchor" }, { AutomationId: "agent-gallery-button" }, { Name: "Select tools",
            ControlType: "Button" }, { Name: "Choose model", ControlType: "Button" }, { Name: "Gemini Enterprise",
                ControlType: "Document" }
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

; mode: "fast" = title + address bar; "full" = + document URL + UIA fingerprint
GeminiEnterprise_IsEnterpriseHwnd(hwnd, mode := "full") {
    if (!GeminiEnterprise_IsChromeHwnd(hwnd))
        return false
    try {
        if (GeminiEnterprise_TitleMatches(WinGetTitle("ahk_id " hwnd)))
            return true
    } catch {
    }
    if (GeminiEnterprise_TryUrlFromAddressBar(hwnd))
        return true
    if (mode = "fast")
        return false
    if (GeminiEnterprise_TryUrlFromDocument(hwnd))
        return true
    return GeminiEnterprise_TryUiaFingerprint(hwnd)
}

GeminiEnterprise_InvalidateCache() {
    global g_GeminiEnterpriseCachedHwnd, g_GeminiEnterpriseHotkeyActive, g_GeminiEnterpriseCachedTitle
    g_GeminiEnterpriseCachedHwnd := 0
    g_GeminiEnterpriseHotkeyActive := false
    g_GeminiEnterpriseCachedTitle := ""
}

GeminiEnterprise_CacheHwnd(hwnd) {
    global g_GeminiEnterpriseCachedHwnd
    if (hwnd && GeminiEnterprise_IsChromeHwnd(hwnd))
        g_GeminiEnterpriseCachedHwnd := hwnd
    else
        g_GeminiEnterpriseCachedHwnd := 0
}

GeminiEnterprise_RefreshHotkeyContext(hwnd, useFull := false) {
    global g_GeminiEnterpriseHotkeyActive, g_GeminiEnterpriseCachedHwnd, g_GeminiEnterpriseCachedTitle
    if (!hwnd || !GeminiEnterprise_IsChromeHwnd(hwnd)) {
        g_GeminiEnterpriseHotkeyActive := false
        g_GeminiEnterpriseCachedTitle := ""
        return false
    }
    mode := useFull ? "full" : "fast"
    active := GeminiEnterprise_IsEnterpriseHwnd(hwnd, mode)
    g_GeminiEnterpriseHotkeyActive := active
    if (active) {
        g_GeminiEnterpriseCachedHwnd := hwnd
        GeminiEnterprise_CacheHwnd(hwnd)
        try {
            g_GeminiEnterpriseCachedTitle := WinGetTitle("ahk_id " hwnd)
        } catch {
            g_GeminiEnterpriseCachedTitle := ""
        }
    } else {
        if (g_GeminiEnterpriseCachedHwnd = hwnd)
            g_GeminiEnterpriseCachedHwnd := 0
        g_GeminiEnterpriseCachedTitle := ""
    }
    return active
}

GeminiEnterprise_OnForegroundChanged(hwnd) {
    if (!hwnd || !WinExist("ahk_id " hwnd)) {
        global g_GeminiEnterpriseHotkeyActive, g_GeminiEnterpriseCachedTitle
        g_GeminiEnterpriseHotkeyActive := false
        g_GeminiEnterpriseCachedTitle := ""
        return
    }
    if (!GeminiEnterprise_IsChromeHwnd(hwnd)) {
        global g_GeminiEnterpriseHotkeyActive, g_GeminiEnterpriseCachedTitle
        g_GeminiEnterpriseHotkeyActive := false
        g_GeminiEnterpriseCachedTitle := ""
        return
    }
    GeminiEnterprise_RefreshHotkeyContext(hwnd, true)
}

GeminiEnterprise_ForegroundHookProc(hHook, event, hwnd, idObject, idChild, idEventThread, dwmsEventTime) {
    SetTimer(GeminiEnterprise_OnForegroundChanged.Bind(hwnd), -0)
}

GeminiEnterprise_EnsureForegroundHook() {
    global g_GeminiEnterprise_ForegroundHookHandle, g_GeminiEnterprise_ForegroundHookCallback
    if (g_GeminiEnterprise_ForegroundHookHandle)
        return
    cb := CallbackCreate(GeminiEnterprise_ForegroundHookProc, "F", 7)
    h := DllCall("user32\SetWinEventHook", "UInt", 0x0003, "UInt", 0x0003, "Ptr", 0, "Ptr", cb, "UInt", 0, "UInt", 0,
        "UInt", 0, "Ptr")
    if (h) {
        g_GeminiEnterprise_ForegroundHookHandle := h
        g_GeminiEnterprise_ForegroundHookCallback := cb
    }
}

IsGeminiEnterpriseChromeActiveForHotkey() {
    GeminiEnterprise_EnsureForegroundHook()
    hwnd := WinExist("A")
    if (!hwnd || !GeminiEnterprise_IsChromeHwnd(hwnd))
        return false
    global g_GeminiEnterpriseHotkeyActive, g_GeminiEnterpriseCachedHwnd, g_GeminiEnterpriseCachedTitle
    if (g_GeminiEnterpriseHotkeyActive && hwnd = g_GeminiEnterpriseCachedHwnd) {
        try {
            title := WinGetTitle("ahk_id " hwnd)
        } catch {
            title := ""
        }
        if (title != g_GeminiEnterpriseCachedTitle)
            return GeminiEnterprise_RefreshHotkeyContext(hwnd, false)
        return true
    }
    return GeminiEnterprise_RefreshHotkeyContext(hwnd, true)
}

GeminiEnterprise_GetActiveUia() {
    return GeminiEnterprise_ReadRootFromHwnd(WinExist("A"))
}

GeminiEnterprise_RunWithBusyBanner(message, fn, hwnd := 0) {
    if (!hwnd)
        hwnd := WinExist("A")
    StandardLoadingBar_Show(message, BANNER_ACCENT_INTERMEDIATE, { centerOnHwnd: hwnd, fontSize: 17 })
    try {
        return fn.Call()
    } finally {
        StandardLoadingBar_Hide(0)
    }
}

GeminiEnterprise_ReturnToComposer() {
    Sleep 40
    Send "{Blind}d{Backspace}"
}

GeminiEnterprise_PlayFocusedChime(minIntervalMs := 400) {
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

GetGeminiEnterpriseWindowHwnd() {
    global g_GeminiEnterpriseCachedHwnd
    if (g_GeminiEnterpriseCachedHwnd && WinExist("ahk_id " g_GeminiEnterpriseCachedHwnd)) {
        if (GeminiEnterprise_IsEnterpriseHwnd(g_GeminiEnterpriseCachedHwnd, "fast"))
            return g_GeminiEnterpriseCachedHwnd
    }
    GeminiEnterprise_InvalidateCache()
    hwnd := WinExist("A")
    if (GeminiEnterprise_IsEnterpriseHwnd(hwnd, "full")) {
        GeminiEnterprise_CacheHwnd(hwnd)
        return hwnd
    }
    try {
        for h in WinGetList("ahk_exe chrome.exe") {
            if (GeminiEnterprise_IsEnterpriseHwnd(h, "full")) {
                GeminiEnterprise_CacheHwnd(h)
                return h
            }
        }
    } catch {
    }
    return 0
}

GeminiEnterprise_GetLaunchUrl() {
    return GEMINI_ENTERPRISE_LAUNCH_URL
}

GeminiEnterprise_ActivateWindow(hwnd, timeoutMs := GEMINI_ENTERPRISE_ACTIVATE_WAIT_MS) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    try {
        WinActivate("ahk_id " hwnd)
    } catch {
        return false
    }
    return WinWaitActive("ahk_id " hwnd, , timeoutMs // 1000)
}

GeminiEnterprise_WaitForComposerDiscoverable(uia, timeoutMs := 500) {
    start := A_TickCount
    while (A_TickCount - start < timeoutMs) {
        if (GeminiEnterprise_FindComposer(uia))
            return true
        Sleep 25
    }
    return !!GeminiEnterprise_FindComposer(uia)
}

; Open/focus Enterprise and put caret in the prompt (global companion #!+I path).
GeminiEnterprise_OpenOrFocus() {
    SetTitleMatchMode(2)
    hwnd := GetGeminiEnterpriseWindowHwnd()
    if (hwnd) {
        alreadyActive := false
        try alreadyActive := WinActive("ahk_id " hwnd)
        if (!alreadyActive) {
            if !GeminiEnterprise_ActivateWindow(hwnd)
                return false
        }
        root := GeminiEnterprise_ReadRootFromHwnd(hwnd)
        if (!root) {
            ShowCenteredOverlay_Utils("❌ Error: Could not attach to Gemini Enterprise window.", 2000,
                BANNER_ACCENT_ERROR)
            return false
        }
        if (!alreadyActive && WinActive("ahk_id " hwnd))
            GeminiEnterprise_WaitForComposerDiscoverable(root)
        if (WinActive("ahk_id " hwnd))
            GeminiEnterprise_FocusComposer(root, true)
        GeminiEnterprise_CacheHwnd(hwnd)
        GeminiEnterprise_RefreshHotkeyContext(hwnd, true)
        return true
    }
    GeminiEnterprise_InvalidateCache()
    try {
        StandardLoadingBar_Show("📤 Opening Gemini Enterprise (2 tabs)...", BANNER_ACCENT_INTERMEDIATE)
        url := GeminiEnterprise_GetLaunchUrl()
        Run 'chrome.exe --new-window "' url '" "' url '"'
        if !WinWaitActive("ahk_exe chrome.exe", , 8) {
            StandardLoadingBar_Hide(0)
            return false
        }
        Sleep GEMINI_ENTERPRISE_FIRST_LAUNCH_WAIT_MS
        hwnd := GetGeminiEnterpriseWindowHwnd()
        if (!hwnd)
            hwnd := WinExist("A")
        if (!hwnd) {
            StandardLoadingBar_Hide(0)
            return false
        }
        root := GeminiEnterprise_ReadRootFromHwnd(hwnd)
        if (!root) {
            StandardLoadingBar_Hide(0)
            return false
        }
        GeminiEnterprise_WaitForComposerDiscoverable(root, 4000)
        if (WinActive("ahk_id " hwnd))
            GeminiEnterprise_FocusComposer(root, true)
        GeminiEnterprise_CacheHwnd(hwnd)
        GeminiEnterprise_RefreshHotkeyContext(hwnd, true)
        StandardLoadingBar_Hide(0)
        return true
    } catch {
        StandardLoadingBar_Hide(0)
        return false
    }
}

; For companion hotkeys that cannot be transposed yet: same as open/focus prompt.
GeminiEnterprise_FocusPromptOnly() {
    return GeminiEnterprise_OpenOrFocus()
}

; Navigate to Enterprise, focus composer, paste snippet (optional text or Clip Angel top item).
GeminiEnterprise_NavigateFocusAndPaste(optionalPromptText := "", autoSubmit := false) {
    SetTitleMatchMode(2)
    hwnd := GetGeminiEnterpriseWindowHwnd()
    if (!hwnd) {
        Run 'chrome.exe --new-window "' GeminiEnterprise_GetLaunchUrl() '"'
        if !WinWaitActive("ahk_exe chrome.exe", , 5)
            return 0
        Sleep GEMINI_ENTERPRISE_FIRST_LAUNCH_WAIT_MS
        hwnd := GetGeminiEnterpriseWindowHwnd()
        if (!hwnd)
            hwnd := WinExist("A")
    }
    if (!hwnd)
        return 0
    WinActivate("ahk_id " hwnd)
    WinWaitActive("ahk_id " hwnd, , 2)
    Sleep GEMINI_ENTERPRISE_UIA_SETTLE_MS
    try {
        uia := UIA_Browser("ahk_id " hwnd)
        Sleep 80
        GeminiEnterprise_FocusComposer(uia, false)
    } catch {
        return 0
    }
    WinActivate("ahk_id " hwnd)
    WinWaitActive("ahk_id " hwnd, , 2)
    Sleep 150
    if (optionalPromptText != "")
        InsertText(optionalPromptText)
    else
        ClipAngel_SendTopListItem(hwnd)
    Sleep 250
    GeminiEnterprise_PlayFocusedChime()
    if (autoSubmit) {
        Sleep 1000
        try {
            uia := UIA_Browser("ahk_id " hwnd)
            GeminiEnterprise_TrySubmit(uia)
        } catch {
            Send "{Enter}"
        }
    }
    if (hwnd)
        GeminiEnterprise_CacheHwnd(hwnd)
    return hwnd
}

; --- Finders -----------------------------------------------------------------

GeminiEnterprise_FindComposer(uia) {
    return GeminiEnterprise_FindFirstInUia(uia, [{ AutomationId: "agent-search-prosemirror-editor" }, { Name: "Search",
        AutomationId: "agent-search-prosemirror-editor" }, { ClassName: "ProseMirror", matchmode: "Substring" }
    ])
}

GeminiEnterprise_FindMenuButton(uia) {
    return GeminiEnterprise_FindFirstInUia(uia, [{ Name: "Menu", ControlType: "Button" }, { Name: "Menu", Type: 50000 }])
}

GeminiEnterprise_FindNewChatButton(uia) {
    return GeminiEnterprise_FindFirstInUia(uia, [{ Name: "New chat", ControlType: "Button" }, { Name: "New chat",
        ClassName: "chat-button", matchmode: "Substring" }, { Name: "Novo chat", ControlType: "Button" }
    ])
}

GeminiEnterprise_FindSearchButton(uia) {
    return GeminiEnterprise_FindFirstInUia(uia, [{ Name: "Search", ControlType: "Button", ClassName: "search-button" }, { Name: "Search",
        ControlType: "Button" }, { Name: "Pesquisar", ControlType: "Button" }, { Name: "Buscar", ControlType: "Button" }
    ])
}

GeminiEnterprise_FindSelectToolsButton(uia) {
    btn := GeminiEnterprise_FindFirstInUia(uia, [{ Name: "Select tools", ControlType: "Button" }, { Name: "Select tools",
        Type: 50000 }])
    if (btn)
        return btn
    anchor := GeminiEnterprise_FindFirstInUia(uia, [{ AutomationId: "tool-selector-menu-anchor" }])
    if (anchor) {
        try {
            nested := anchor.FindFirst({ Name: "Select tools", ControlType: "Button" })
            if (nested)
                return nested
        } catch {
        }
    }
    return 0
}

GeminiEnterprise_FindChooseModelButton(uia) {
    btn := GeminiEnterprise_FindFirstInUia(uia, [{ Name: "Choose model", ControlType: "Button" }, { Name: "Choose model",
        Type: 50000 }])
    if (btn)
        return btn
    anchor := GeminiEnterprise_FindFirstInUia(uia, [{ AutomationId: "model-selector-menu-anchor" }])
    if (anchor) {
        try {
            nested := anchor.FindFirst({ Name: "Choose model", ControlType: "Button" })
            if (nested)
                return nested
        } catch {
        }
    }
    return 0
}

GeminiEnterprise_FindSubmitButton(uia) {
    return GeminiEnterprise_FindFirstInUia(uia, [{ Name: "Submit", ControlType: "Button" }, { ClassName: "send-button",
        matchmode: "Substring" }])
}

GeminiEnterprise_FindStopButton(uia) {
    if (!IsObject(uia))
        return 0
    el := GeminiEnterprise_FindFirstInUia(uia, [{ Name: "Stop generating", ControlType: "Button" }, { Name: "Stop response",
        ControlType: "Button" }, { Name: "Stop", ControlType: "Button" }
    ])
    if (el)
        return el
    try {
        return uia.FindFirst({ Name: "Stop", matchmode: "Substring", ControlType: "Button" })
    } catch {
    }
    return 0
}

GeminiEnterprise_FindMenuItemByNames(uia, names) {
    if (!IsObject(uia) || !IsObject(names))
        return 0
    criteria := []
    for n in names {
        criteria.Push({ Name: n, ControlType: "MenuItem" })
        criteria.Push({ Name: n, Type: 50011 })
        criteria.Push({ Name: n, matchmode: "Substring", ControlType: "MenuItem" })
    }
    return GeminiEnterprise_FindFirstInUia(uia, criteria)
}

GeminiEnterprise_WaitForMenuItemByNames(names, timeoutMs := 2000, hwnd := 0) {
    if (!hwnd)
        hwnd := WinExist("A")
    start := A_TickCount
    while (A_TickCount - start < timeoutMs) {
        uia := GeminiEnterprise_ReadRootFromHwnd(hwnd)
        if (IsObject(uia)) {
            el := GeminiEnterprise_FindMenuItemByNames(uia, names)
            if (el)
                return el
        }
        Sleep GEMINI_ENTERPRISE_MENU_POLL_MS
    }
    return 0
}

; --- Actions -----------------------------------------------------------------

GeminiEnterprise_FocusComposer(uia := 0, playChime := true) {
    if (!uia)
        uia := GeminiEnterprise_GetActiveUia()
    el := GeminiEnterprise_FindComposer(uia)
    if (!el)
        return 0
    try {
        if (el.HasKeyboardFocus) {
            if (playChime)
                GeminiEnterprise_PlayFocusedChime()
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
            GeminiEnterprise_PlayFocusedChime()
    } catch {
        if (playChime)
            GeminiEnterprise_PlayFocusedChime()
    }
    return el
}

GeminiEnterprise_ToggleNavDrawer(uia := 0) {
    if (!uia)
        uia := GeminiEnterprise_GetActiveUia()
    if (!IsObject(uia))
        return false
    btn := GeminiEnterprise_FindMenuButton(uia)
    return btn && GeminiEnterprise_ClickUiaElement(btn)
}

GeminiEnterprise_ClickNewChat(uia := 0) {
    if (!uia)
        uia := GeminiEnterprise_GetActiveUia()
    if (!IsObject(uia))
        return false
    el := GeminiEnterprise_FindNewChatButton(uia)
    return el && GeminiEnterprise_ClickUiaElement(el)
}

GeminiEnterprise_ClickNavSearch(uia := 0) {
    if (!uia)
        uia := GeminiEnterprise_GetActiveUia()
    if (!IsObject(uia))
        return false
    el := GeminiEnterprise_FindSearchButton(uia)
    if (el)
        return GeminiEnterprise_ClickUiaElement(el)
    ; Drawer may be collapsed — open Menu then retry.
    if (!GeminiEnterprise_ToggleNavDrawer(uia))
        return false
    Sleep GEMINI_ENTERPRISE_UIA_SETTLE_MS
    uia := GeminiEnterprise_GetActiveUia()
    el := GeminiEnterprise_FindSearchButton(uia)
    return el && GeminiEnterprise_ClickUiaElement(el)
}

GeminiEnterprise_OpenToolsMenu(uia := 0) {
    if (!uia)
        uia := GeminiEnterprise_GetActiveUia()
    if (!IsObject(uia))
        return false
    ; Already open?
    if (GeminiEnterprise_FindMenuItemByNames(uia, ["Search company data", "Create images", "Deep Research"]))
        return true
    btn := GeminiEnterprise_FindSelectToolsButton(uia)
    if (!btn)
        return false
    if (!GeminiEnterprise_ClickUiaElement(btn))
        return false
    Sleep GEMINI_ENTERPRISE_UIA_SETTLE_MS
    return !!GeminiEnterprise_WaitForMenuItemByNames(["Search company data", "Create images", "Deep Research",
        "Search the web"])
}

GeminiEnterprise_ClickToolMenuItem(nameNeedles, uia := 0) {
    hwnd := WinExist("A")
    if (!uia)
        uia := GeminiEnterprise_GetActiveUia()
    if (!IsObject(uia))
        return false
    item := GeminiEnterprise_FindMenuItemByNames(uia, nameNeedles)
    if (!item) {
        if (!GeminiEnterprise_OpenToolsMenu(uia))
            return false
        item := GeminiEnterprise_WaitForMenuItemByNames(nameNeedles, GEMINI_ENTERPRISE_MENU_WAIT_MS, hwnd)
    }
    if (!item)
        return false
    return GeminiEnterprise_ClickUiaElement(item)
}

GeminiEnterprise_ClickCreateImages(uia := 0) {
    return GeminiEnterprise_ClickToolMenuItem(["Create images", "Create image", "Criar imagens"], uia)
}

GeminiEnterprise_ClickDeepResearch(uia := 0) {
    return GeminiEnterprise_ClickToolMenuItem(["Deep Research", "Deep research", "Pesquisa aprofundada"], uia)
}

GeminiEnterprise_OpenModelSelector(uia := 0) {
    if (!uia)
        uia := GeminiEnterprise_GetActiveUia()
    if (!IsObject(uia))
        return false
    btn := GeminiEnterprise_FindChooseModelButton(uia)
    if (!btn)
        return false
    if (!GeminiEnterprise_ClickUiaElement(btn))
        return false
    Sleep GEMINI_ENTERPRISE_UIA_SETTLE_MS
    return true
}

GeminiEnterprise_GetModelSelectorLabel(btn) {
    if (!IsObject(btn))
        return ""
    label := ""
    try label := btn.Name
    if (label != "" && !InStr(label, "Choose model", false))
        return label
    try {
        for child in btn.FindAll({ ControlType: "Text" }) {
            n := ""
            try n := child.Name
            if (n != "" && !InStr(n, "Choose model", false))
                return n
        }
    } catch {
    }
    ; Parent group may hold the visible model name (e.g. "Auto", "3.1 Pro").
    try {
        parent := btn.Parent
        if (IsObject(parent)) {
            for child in parent.FindAll({ ControlType: "Text" }) {
                n := ""
                try n := child.Name
                if (n != "" && !InStr(n, "Choose model", false))
                    return n
            }
        }
    } catch {
    }
    return label
}

GeminiEnterprise_IsModelSelected(modelName, uia := 0) {
    if (!modelName)
        return false
    if (!uia)
        uia := GeminiEnterprise_GetActiveUia()
    btn := GeminiEnterprise_FindChooseModelButton(uia)
    if (!btn)
        return false
    label := GeminiEnterprise_GetModelSelectorLabel(btn)
    return (label != "" && InStr(label, modelName, false))
}

GeminiEnterprise_SelectModelByName(modelName, hwnd := 0) {
    if (!modelName)
        return false
    if (!hwnd)
        hwnd := WinExist("A")
    uia := GeminiEnterprise_ReadRootFromHwnd(hwnd)
    if (!IsObject(uia))
        return false
    if (GeminiEnterprise_IsModelSelected(modelName, uia))
        return true
    if (!GeminiEnterprise_OpenModelSelector(uia))
        return false
    item := GeminiEnterprise_WaitForMenuItemByNames([modelName], GEMINI_ENTERPRISE_MENU_WAIT_MS, hwnd)
    if (!item) {
        ; Some builds expose models as buttons / list items rather than MenuItem.
        uia := GeminiEnterprise_ReadRootFromHwnd(hwnd)
        item := GeminiEnterprise_FindFirstInUia(uia, [{ Name: modelName, ControlType: "Button" }, { Name: modelName,
            ControlType: "ListItem" }, { Name: modelName, matchmode: "Substring", ControlType: "MenuItem" }, { Name: modelName,
                matchmode: "Substring", ControlType: "Button" }
        ])
    }
    if (!item)
        return false
    return GeminiEnterprise_ClickUiaElement(item)
}

GeminiEnterprise_SelectDeepReasoningModel(hwnd := 0) {
    modelName := GEMINI_ENTERPRISE_DEEP_MODEL
    try {
        fromIni := AiCompanionModels_GetDeep(AI_COMPANION_ENTERPRISE)
        if (fromIni != "")
            modelName := fromIni
    } catch {
    }
    return GeminiEnterprise_SelectModelByName(modelName, hwnd)
}

GeminiEnterprise_TrySubmit(uia := 0) {
    if (!uia)
        uia := GeminiEnterprise_GetActiveUia()
    if (!IsObject(uia))
        return false
    if (GeminiEnterprise_FindStopButton(uia))
        return false
    sendBtn := GeminiEnterprise_FindSubmitButton(uia)
    if (sendBtn) {
        if (GeminiEnterprise_ClickUiaElement(sendBtn))
            return true
    }
    SendInput "{Enter}"
    return true
}

GeminiEnterprise_WaitForGenerationComplete(timeoutMs := 300000) {
    hwnd := WinExist("A")
    if (!hwnd || !GeminiEnterprise_IsEnterpriseHwnd(hwnd, "fast"))
        return
    start := A_TickCount
    sawStop := false
    while (A_TickCount - start < timeoutMs) {
        if (!WinExist("ahk_id " hwnd))
            return
        uia := GeminiEnterprise_ReadRootFromHwnd(hwnd)
        stopBtn := IsObject(uia) ? GeminiEnterprise_FindStopButton(uia) : 0
        if (stopBtn) {
            sawStop := true
        } else if (sawStop) {
            if (IsSoundEnabled())
                ScriptSoundPlay(A_ScriptDir . "\assets\sounds\gemini-completion.wav")
            return
        } else {
            ; No stop control ever appeared — do not block forever on landing submit.
            if (A_TickCount - start > 2500)
                return
        }
        Sleep 500
    }
}

; --- Copy last response + async streaming monitor (#!+8 pronunciation) ----------

GeminiEnterprise_IsCopyResponseButton(name) {
    if (!name)
        return false
    if (InStr(name, "Copy prompt", false) || InStr(name, "Copiar prompt", false))
        return false
    for n in GEMINI_ENTERPRISE_COPY_RESPONSE_NAMES {
        if (name = n)
            return true
    }
    if (InStr(name, "Copy response", false) = 1 || InStr(name, "Copy Response", false) = 1)
        return true
    if (InStr(name, "Copy message", false) = 1)
        return true
    return false
}

GeminiEnterprise_GetCopyButtonsArray(uia) {
    out := []
    if (!IsObject(uia))
        return out
    scope := uia
    usedPanel := false
    try {
        panel := GeminiEnterprise_FindFirstInUia(uia, [{ AutomationId: "main-panel" }])
        if (IsObject(panel)) {
            scope := panel
            usedPanel := true
        }
    } catch {
    }
    try {
        allButtons := scope.FindAll({ Type: "Button" })
        for button in allButtons {
            if (GeminiEnterprise_IsCopyResponseButton(button.Name))
                out.Push(button)
        }
    } catch {
    }
    if (out.Length = 0 && usedPanel) {
        try {
            allButtons := uia.FindAll({ Type: "Button" })
            for button in allButtons {
                if (GeminiEnterprise_IsCopyResponseButton(button.Name))
                    out.Push(button)
            }
        } catch {
        }
    }
    return out
}

GeminiEnterprise_GetLastCopyButton(uia) {
    arr := GeminiEnterprise_GetCopyButtonsArray(uia)
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

GeminiEnterprise_CopyLastMessageToClipboard(options := "", enterpriseHwnd := 0) {
    restoreWindow := (options = "" || !options.HasProp("restoreWindow")) ? true : options.restoreWindow
    playChimeAndNotify := (options = "" || !options.HasProp("playChimeAndNotify")) ? true : options.playChimeAndNotify
    alreadyActive := (options != "" && options.HasProp("alreadyActive")) ? options.alreadyActive : false
    try {
        SetTitleMatchMode(2)
        if (!enterpriseHwnd)
            enterpriseHwnd := GetGeminiEnterpriseWindowHwnd()
        if (!enterpriseHwnd)
            return false
        if (!alreadyActive) {
            if !GeminiEnterprise_ActivateWindow(enterpriseHwnd)
                return false
            Sleep GEMINI_ENTERPRISE_UIA_SETTLE_MS
        }
        GeminiEnterprise_ScrollFeedToBottom(enterpriseHwnd)
        uia := alreadyActive ? UIA_Browser() : UIA_Browser("ahk_id " enterpriseHwnd)
        Sleep GEMINI_ENTERPRISE_UIA_SETTLE_MS
        copyBtn := GeminiEnterprise_GetLastCopyButton(uia)
        if (!copyBtn)
            return false
        A_Clipboard := ""
        if (!GeminiEnterprise_ClickUiaElement(copyBtn)) {
            try copyBtn.Click()
            catch {
                return false
            }
        }
        if !ClipWait(2)
            return false
        if (playChimeAndNotify) {
            try ScriptSoundPlay(A_ScriptDir . "\assets\sounds\copy.wav")
            ShowCenteredOverlay_Utils("Copied!", 800, BANNER_ACCENT_SUCCESS)
        }
        if (restoreWindow)
            Send "!{Tab}"
        else {
            root := GeminiEnterprise_ReadRootFromHwnd(enterpriseHwnd)
            if (IsObject(root))
                GeminiEnterprise_FocusComposer(root, false)
        }
        return true
    } catch {
        return false
    }
}

GeminiEnterprise_CopyLastMessageWithRetry(options := "", enterpriseHwnd := 0, maxRetries :=
    GEMINI_ENTERPRISE_COPY_MAX_RETRIES) {
    baseDelay := GEMINI_ENTERPRISE_COPY_RETRY_SLEEP_MS
    loop maxRetries {
        if (GeminiEnterprise_CopyLastMessageToClipboard(options, enterpriseHwnd))
            return true
        if (A_Index < maxRetries)
            Sleep baseDelay * (1 << (A_Index - 1))
    }
    return false
}

GeminiEnterpriseBackgroundSetTimer(task, callback, periodMs := GEMINI_ENTERPRISE_ASYNC_POLL_MS) {
    GeminiEnterpriseBackgroundStopTimer(task)
    task.TimerCallback := callback
    SetTimer(task.TimerCallback, periodMs)
}

GeminiEnterpriseBackgroundStopTimer(task) {
    cb := ""
    try cb := task.TimerCallback
    catch
        cb := ""
    if (cb)
        SetTimer(cb, 0)
    try task.TimerCallback := ""
}

GeminiEnterprise_VerifyStreamingStopped(enterpriseHwnd) {
    loop GEMINI_ENTERPRISE_STREAM_GONE_LOOPS {
        Sleep GEMINI_ENTERPRISE_STREAM_GONE_VERIFY_MS
        root := GeminiEnterprise_ReadRootFromHwnd(enterpriseHwnd)
        if (!root)
            return true
        try {
            if (!GeminiEnterprise_FindStopButton(root))
                continue
            return false
        } catch {
            return true
        }
    }
    return true
}

GeminiEnterprise_MonitorStreamingTransition(task, onCompleteCallback) {
    task.RetryCount++
    if (task.RetryCount > task.MaxRetries) {
        GeminiEnterpriseBackgroundStopTimer(task)
        return "timeout"
    }
    if (!task.EnterpriseHwnd || !WinExist("ahk_id " task.EnterpriseHwnd)) {
        GeminiEnterpriseBackgroundStopTimer(task)
        return "unavailable"
    }
    uia := GeminiEnterprise_ReadRootFromHwnd(task.EnterpriseHwnd)
    if (!uia)
        return "unavailable"
    if (GeminiEnterprise_FindStopButton(uia)) {
        task.ButtonEverFound := true
        return "streaming"
    }
    if (!task.ButtonEverFound)
        return "waiting"
    if (!GeminiEnterprise_VerifyStreamingStopped(task.EnterpriseHwnd))
        return "streaming"
    GeminiEnterpriseBackgroundStopTimer(task)
    onCompleteCallback.Call()
    return "completed"
}

; ProseMirror often lacks Value/TextPattern — clipboard read after focus is primary.
GeminiEnterprise_ComposerGetTextViaClipboard(hwnd := 0) {
    if (!hwnd)
        hwnd := WinExist("A")
    if (!hwnd)
        return ""
    root := GeminiEnterprise_ReadRootFromHwnd(hwnd)
    if (!IsObject(root) || !GeminiEnterprise_FocusComposer(root, false))
        return ""
    Sleep 60
    saved := ClipboardAll()
    try {
        A_Clipboard := ""
        Send "^a"
        Sleep 40
        Send "^c"
        if !ClipWait(1, 1)
            return ""
        text := A_Clipboard
        if (Type(text) != "String")
            text := ""
        text := Trim(text)
        if (text = "" || InStr(text, "Ask anything", false))
            return ""
        return text
    } finally {
        Sleep 40
        try A_Clipboard := saved
        catch {
        }
    }
}

GeminiEnterprise_ComposerGetTextViaUia(hwnd := 0) {
    try {
        if (!hwnd)
            hwnd := WinExist("A")
        uia := GeminiEnterprise_ReadRootFromHwnd(hwnd)
        if (!uia)
            return ""
        pf := GeminiEnterprise_FindComposer(uia)
        if (!pf)
            return ""
        try {
            text := Trim(pf.Value)
            if (text != "" && !InStr(text, "Ask anything", false))
                return text
        } catch {
        }
        try {
            text := Trim(pf.TextPattern.DocumentRange.GetText(-1))
            if (text != "" && !InStr(text, "Ask anything", false))
                return text
        } catch {
        }
    } catch {
    }
    return ""
}

GeminiEnterprise_ComposerGetText(hwnd := 0) {
    if (!hwnd)
        hwnd := WinExist("A")
    text := GeminiEnterprise_ComposerGetTextViaClipboard(hwnd)
    if (text != "")
        return text
    return GeminiEnterprise_ComposerGetTextViaUia(hwnd)
}

; reason: "" on success; "empty" if unreadable; "nodivider" if no --- block.
GeminiEnterprise_StripComposerHumanReminders(&reason := "") {
    reason := ""
    hwnd := WinExist("A")
    text := GeminiEnterprise_ComposerGetText(hwnd)
    if (text = "") {
        reason := "empty"
        return false
    }
    newText := StripPromptHumanReminders(text)
    if (newText = "") {
        reason := "nodivider"
        return false
    }
    root := GeminiEnterprise_ReadRootFromHwnd(hwnd)
    if (IsObject(root))
        GeminiEnterprise_FocusComposer(root, false)
    if !ReplaceFocusedEditWithText(newText) {
        reason := "empty"
        return false
    }
    return true
}

GeminiEnterprise_ScrollFeedToBottom(hwnd := 0) {
    if (!hwnd)
        hwnd := WinExist("A")
    if (!hwnd)
        return false
    try {
        preUia := 0
        try preUia := UIA_Browser("ahk_id " hwnd)
        pf := IsObject(preUia) ? GeminiEnterprise_FindComposer(preUia) : 0
        snapshot := ChromeChat_ComposerSnapshot(pf)

        mainPanel := 0
        omnibar := 0
        copyBtn := 0
        panelOk := false
        pfOk := false
        omniOk := false
        copyOk := false
        if (IsObject(preUia)) {
            try {
                root := preUia.DocumentElement
                if (!IsObject(root))
                    root := UIA.ElementFromHandle(hwnd)
                mainPanel := ChromeChat_FindFirstInUia(root, [{ AutomationId: "main-panel" }, { ClassName: "enable-full-page-scroll-view",
                    matchmode: "Substring" }])
                panelOk := ChromeChat_ScrollElementToBottom(mainPanel)
                copyBtn := GeminiEnterprise_GetLastCopyButton(preUia)
                if (copyBtn) {
                    try {
                        copyBtn.ScrollIntoView()
                        copyOk := true
                    } catch {
                    }
                }
                try omnibar := root.FindFirst({ ClassName: "omnibar", matchmode: "Substring" })
            } catch {
            }
        }
        if (pf) {
            try {
                pf.ScrollIntoView()
                pfOk := true
            } catch {
            }
        }
        if (IsObject(omnibar)) {
            try {
                omnibar.ScrollIntoView()
                omniOk := true
            } catch {
            }
        }
        wheelTarget := mainPanel ? mainPanel : (omnibar ? omnibar : (pf ? pf : 0))
        if (wheelTarget)
            ChromeChat_ScrollViaMouseWheelAtElement(wheelTarget, hwnd, 80)
        else
            ChromeChat_ScrollFeedToBottomFallback(hwnd)
        ; #region agent log
        ChromeChat_DebugLog("H-D", "GeminiEnterprise:scroll", "uia scroll", '{"hasPf":' . (IsObject(pf) ? 1 : 0) .
        ',"pfOk":' . (pfOk ? 1 : 0) . ',"hasMainPanel":' . (IsObject(mainPanel) ? 1 : 0) . ',"panelOk":' . (
            panelOk ? 1 : 0) . ',"hasCopyBtn":' . (IsObject(copyBtn) ? 1 : 0) . ',"copyOk":' . (copyOk ? 1 : 0) .
        ',"omniOk":' . (omniOk ? 1 : 0) . '}')
        ; #endregion

        if (!IsObject(pf) && IsObject(preUia))
            pf := GeminiEnterprise_FindComposer(preUia)
        if (pf) {
            try pf.ScrollIntoView()
            catch {
            }
        }
        ChromeChat_ComposerRestore(pf, snapshot)
        return true
    } catch {
        return false
    }
}

; Shift+A: 3.1 Pro + Create images + bosch-brand-image (strip reminders).
GeminiEnterprise_ShiftArt() {
    hwnd := WinExist("A")
    if (!GeminiEnterprise_SelectDeepReasoningModel(hwnd))
        return false
    Sleep GEMINI_ENTERPRISE_UIA_SETTLE_MS
    if !GeminiEnterprise_ClickCreateImages()
        return false
    promptText := GetPromptText("bosch-brand-image")
    if (InStr(promptText, "[PROMPT FILE MISSING:"))
        return false
    stripped := StripPromptHumanReminders(promptText)
    if (stripped != "")
        promptText := stripped
    root := GeminiEnterprise_ReadRootFromHwnd(hwnd)
    if (IsObject(root))
        GeminiEnterprise_FocusComposer(root, false)
    return ReplaceFocusedEditWithText(promptText)
}

; --- Async pronunciation lookup (Win+Alt+Shift+8) -------------------------------

class GeminiEnterpriseAsyncLookup {
    __New(lang, selectedText := "", preCopiedText := "") {
        this.Lang := lang
        this.PreCopiedText := preCopiedText
        this.OriginalHwnd := 0
        this.EnterpriseHwnd := 0
        this.RetryCount := 0
        this.MaxRetries := GEMINI_ENTERPRISE_ASYNC_MAX_RETRIES
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
        this.EnterpriseHwnd := GetGeminiEnterpriseWindowHwnd()
        if !this.EnterpriseHwnd {
            GeminiEnterprise_OpenOrFocus()
            this.EnterpriseHwnd := GetGeminiEnterpriseWindowHwnd()
        }
        if !this.EnterpriseHwnd {
            StandardLoadingBar_Hide(0)
            return
        }
        if !GeminiEnterprise_ActivateWindow(this.EnterpriseHwnd) {
            StandardLoadingBar_Hide(0)
            return
        }
        try uia := UIA_Browser("ahk_id " this.EnterpriseHwnd)
        catch {
            StandardLoadingBar_Hide(0)
            return
        }
        Sleep 300
        if (!GeminiEnterprise_FocusComposer(uia, false)) {
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
        GeminiEnterpriseBackgroundSetTimer(this, this.CheckCompletion.Bind(this), GEMINI_ENTERPRISE_ASYNC_POLL_MS)
    }

    CheckCompletion() {
        state := GeminiEnterprise_MonitorStreamingTransition(this, this.OnStreamingCompleted.Bind(this))
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
        try WinActivate("ahk_id " this.EnterpriseHwnd)
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
        if !GeminiEnterprise_CopyLastMessageWithRetry(copyOpt, this.EnterpriseHwnd) {
            StandardLoadingBar_Hide(0)
            return
        }
        syncElapsed := 0
        while (syncElapsed < GEMINI_ENTERPRISE_POST_COPY_SYNC_TIMEOUT_MS) {
            if (Clipboard_GetSequenceNumber() != seqBefore)
                break
            Sleep GEMINI_ENTERPRISE_CLIPBOARD_POLL_MS
            syncElapsed += GEMINI_ENTERPRISE_CLIPBOARD_POLL_MS
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

GeminiEnterpriseHotkey_ShowPronunciationLanguagePicker(selectedText) {
    onSelect(lang) {
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        if (lang != "")
            (GeminiEnterpriseAsyncLookup(lang, selectedText)).Start()
    }
    onTimeout() {
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        StandardLoadingBar_Show("⏳ Detecting language…", BANNER_ACCENT_INTERMEDIATE, { textWidth: 450, fontSize: 17 })
        lang := DetectLang_AhkFallback(selectedText)
        if !(lang = "pt" || lang = "en" || lang = "de")
            lang := "en"
        (GeminiEnterpriseAsyncLookup(lang, selectedText)).Start()
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
