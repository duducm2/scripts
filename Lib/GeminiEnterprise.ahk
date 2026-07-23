; Gemini Enterprise (Vertex AI Search / AskBosch) Chrome automation.
; Included from Utils.ahk after CopilotWeb.ahk. Requires: env.ahk, Utils.ahk, UIA.

GEMINI_ENTERPRISE_URL_NEEDLE := "vertexaisearch.cloud.google.com"
GEMINI_ENTERPRISE_TITLE_NEEDLE := "Gemini Enterprise"
GEMINI_ENTERPRISE_MENU_WAIT_MS := 2000
GEMINI_ENTERPRISE_MENU_POLL_MS := 80
GEMINI_ENTERPRISE_UIA_SETTLE_MS := 120

global g_GeminiEnterpriseCachedHwnd := 0
global g_GeminiEnterpriseHotkeyActive := false
global g_GeminiEnterpriseCachedTitle := ""
global g_GeminiEnterprise_ForegroundHookHandle := 0
global g_GeminiEnterprise_ForegroundHookCallback := 0

global g_GeminiEnterpriseModelSelectorGui := false
global g_GeminiEnterpriseModelSelectorActive := false
global g_GeminiEnterpriseModelHotkeyHandlers := []
global g_GeminiEnterpriseModelCharSequence := ["1", "2", "3", "4"]
global g_GeminiEnterpriseModels := [{ name: "Auto", description: "Gemini chooses the best fit" }, { name: "3.1 Pro",
    description: "State-of-the-art reasoning" }, { name: "3.5 Flash", description: "Frontier intelligence built for speed" }, { name: "2.5 Pro",
        description: "Solves complex problems" }
]

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

GeminiEnterprise_SelectModelByName(modelName, hwnd := 0) {
    if (!modelName)
        return false
    if (!hwnd)
        hwnd := WinExist("A")
    uia := GeminiEnterprise_ReadRootFromHwnd(hwnd)
    if (!IsObject(uia))
        return false
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

GeminiEnterprise_ComposerGetText(hwnd := 0) {
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
        ; Fallback: concatenate text children under ProseMirror (skip placeholder).
        try {
            out := ""
            for child in pf.FindAll({ ControlType: "Text" }) {
                n := ""
                try n := child.Name
                if (n = "" || InStr(n, "Ask anything", false) || n = "`n")
                    continue
                out .= n
            }
            if (out != "")
                return Trim(out)
        } catch {
        }
    } catch {
    }
    return ""
}

GeminiEnterprise_StripComposerHumanReminders() {
    hwnd := WinExist("A")
    text := GeminiEnterprise_ComposerGetText(hwnd)
    if (text = "")
        return false
    newText := StripPromptHumanReminders(text)
    if (newText = "")
        return false
    root := GeminiEnterprise_ReadRootFromHwnd(hwnd)
    if (IsObject(root))
        GeminiEnterprise_FocusComposer(root, false)
    return ReplaceFocusedEditWithText(newText)
}

; --- Model selector wizard (Shift+M) ----------------------------------------

CleanupGeminiEnterpriseModelSelector() {
    global g_GeminiEnterpriseModelSelectorActive, g_GeminiEnterpriseModelSelectorGui
    global g_GeminiEnterpriseModelHotkeyHandlers
    g_GeminiEnterpriseModelSelectorActive := false
    for handler in g_GeminiEnterpriseModelHotkeyHandlers {
        try Hotkey(handler.char, handler.handler, "Off")
        catch {
        }
    }
    g_GeminiEnterpriseModelHotkeyHandlers := []
    if (IsObject(g_GeminiEnterpriseModelSelectorGui)) {
        try g_GeminiEnterpriseModelSelectorGui.Destroy()
        catch {
        }
        g_GeminiEnterpriseModelSelectorGui := false
    }
}

CreateGeminiEnterpriseModelHandler(char) {
    return (*) => HandleGeminiEnterpriseModelSelection(char)
}

HandleGeminiEnterpriseModelSelection(char) {
    global g_GeminiEnterpriseModelSelectorActive, g_GeminiEnterpriseModels
    global g_GeminiEnterpriseModelCharSequence
    if (!g_GeminiEnterpriseModelSelectorActive)
        return
    modelIndex := -1
    for idx, ch in g_GeminiEnterpriseModelCharSequence {
        if (ch = char) {
            modelIndex := idx
            break
        }
    }
    if (modelIndex < 1 || modelIndex > g_GeminiEnterpriseModels.Length)
        return
    modelInfo := g_GeminiEnterpriseModels[modelIndex]
    if (!IsObject(modelInfo))
        return
    modelName := ""
    try modelName := modelInfo.name
    if (modelName = "")
        return
    try CleanupGeminiEnterpriseModelSelector()
    catch {
    }
    hwnd := WinExist("A")
    ok := GeminiEnterprise_RunWithBusyBanner("🔄 Switching model to " modelName "…", (*) =>
        GeminiEnterprise_SelectModelByName(modelName, hwnd), hwnd)
    if (ok) {
        try GeminiEnterprise_FocusComposer(GeminiEnterprise_GetActiveUia(), false)
        GeminiEnterprise_ReturnToComposer()
    } else
        ShowCenteredOverlay_Utils("❌ Could not select model: " . modelName, 2800, BANNER_ACCENT_ERROR)
}

ShowGeminiEnterpriseModelSelector() {
    global g_GeminiEnterpriseModelSelectorGui, g_GeminiEnterpriseModelSelectorActive
    global g_GeminiEnterpriseModelHotkeyHandlers, g_GeminiEnterpriseModels
    global g_GeminiEnterpriseModelCharSequence
    if (!IsObject(g_GeminiEnterpriseModels) || g_GeminiEnterpriseModels.Length = 0)
        return
    if (g_GeminiEnterpriseModelSelectorActive && IsObject(g_GeminiEnterpriseModelSelectorGui)) {
        try g_GeminiEnterpriseModelSelectorGui.Show()
        return
    }
    CleanupGeminiEnterpriseModelSelector()

    lines := "Gemini Enterprise — Choose model`n`n"
    charIndex := 0
    for model in g_GeminiEnterpriseModels {
        if (charIndex < g_GeminiEnterpriseModelCharSequence.Length) {
            char := g_GeminiEnterpriseModelCharSequence[charIndex + 1]
            desc := ""
            try desc := model.description
            lines .= "[" char "] " model.name
            if (desc != "")
                lines .= " — " desc
            lines .= "`n"
            charIndex++
        }
    }
    lines .= "`n[Esc] Close"

    g_GeminiEnterpriseModelSelectorGui := Gui("+AlwaysOnTop +ToolWindow +E0x08000000", "Gemini Enterprise Model")
    g_GeminiEnterpriseModelSelectorGui.SetFont("s14", "Segoe UI")
    g_GeminiEnterpriseModelSelectorGui.MarginX := 10
    g_GeminiEnterpriseModelSelectorGui.MarginY := 5
    g_GeminiEnterpriseModelSelectorGui.AddEdit("w420 h180 ReadOnly VScroll", lines)
    closeBtn := g_GeminiEnterpriseModelSelectorGui.AddButton("w100 Default Center", "Close")
    closeBtn.OnEvent("Click", (*) => CleanupGeminiEnterpriseModelSelector())
    g_GeminiEnterpriseModelSelectorGui.OnEvent("Close", (*) => CleanupGeminiEnterpriseModelSelector())
    g_GeminiEnterpriseModelSelectorGui.OnEvent("Escape", (*) => CleanupGeminiEnterpriseModelSelector())

    try {
        MouseGetPos(&mx, &my)
        g_GeminiEnterpriseModelSelectorGui.Show("NA w440 h230 x" (mx + 12) " y" (my + 12))
    } catch {
        g_GeminiEnterpriseModelSelectorGui.Show("NA w440 h230")
    }

    g_GeminiEnterpriseModelSelectorActive := true
    g_GeminiEnterpriseModelHotkeyHandlers := []
    charIndex := 0
    for model in g_GeminiEnterpriseModels {
        if (charIndex < g_GeminiEnterpriseModelCharSequence.Length) {
            char := g_GeminiEnterpriseModelCharSequence[charIndex + 1]
            handler := CreateGeminiEnterpriseModelHandler(char)
            try Hotkey(char, handler, "On")
            g_GeminiEnterpriseModelHotkeyHandlers.Push({ char: char, handler: handler })
            charIndex++
        }
    }
}
