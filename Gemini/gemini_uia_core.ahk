; =============================================================================
; Gemini module: gemini_uia_core.ahk
; GeminiState, copy-button UIA, tab banner, model picker
; Extracted verbatim from Gemini.ahk; loaded via #include into the
; Gemini.ahk process, which remains the entry point / source of truth.
; =============================================================================

; --- GeminiState singleton: cache by hwnd for O(1) validation on subsequent use in same lifecycle ---
class GeminiState {
    static _hwnd := 0
    static _lastCopyButton := ""

    static Invalidate() {
        GeminiState._hwnd := 0
        GeminiState._lastCopyButton := ""
    }

    ; Get last copy button; use cached element if same hwnd and element still valid (O(1) check), else discover and cache.
    static GetLastCopyButtonCached(uia, geminiHwnd) {
        if (geminiHwnd && GeminiState._hwnd = geminiHwnd && GeminiState._lastCopyButton != "") {
            try {
                _ := GeminiState._lastCopyButton.Name
                return GeminiState._lastCopyButton
            } catch
                GeminiState.Invalidate()
        }
        btn := GetLastGeminiCopyButton(uia)
        if (btn && geminiHwnd) {
            GeminiState._hwnd := geminiHwnd
            GeminiState._lastCopyButton := btn
        }
        return btn
    }
}

; Callable by name from Utils.ahk (avoids #Warn UseUnsetLocal for GeminiState).
GeminiStateInvalidate() {
    GeminiState.Invalidate()
}

; --- Helper Functions --------------------------------------------------------
; FindGeminiPromptField and GEMINI_PROMPT_FIELD_NAMES are defined in Utils.ahk (included above).

; True if button name is the "Copy [last response]" button (EN or PT), not "Copy prompt" / "Copy code".
IsGeminiCopyResponseButton(name) {
    if (!name || InStr(name, "prompt") || InStr(name, "code"))
        return false
    for n in GEMINI_COPY_RESPONSE_NAMES {
        if (name = n || InStr(name, n, false))
            return true
    }
    return false
}

; True if button name is the code-fence "Copy code" button (EN or PT), not "Copy prompt" / bare "Copy".
IsGeminiCopyCodeButton(name) {
    if (!name || InStr(name, "prompt"))
        return false
    for n in GEMINI_COPY_CODE_NAMES {
        if (name = n || InStr(name, n, false) = 1)
            return true
    }
    return false
}

Gemini_ClassHasToken(className, token) {
    if (!className || !token)
        return false
    normalized := StrReplace(StrReplace(StrReplace(className, "`n", " "), "`r", " "), "`t", " ")
    for part in StrSplit(normalized, " ", " `t") {
        if (part = token)
            return true
    }
    return false
}

; Active gemini.google.com document. A second Chrome tab in the same window must not
; contribute Copy / Copy code buttons (its bounds can sit on top of the visible footer Copy).
Gemini_ActivePageRoot(uia) {
    if (!IsObject(uia))
        return uia
    docs := []
    try docs := uia.FindAll({ Type: "Document" })
    catch
        docs := []
    geminiDocs := []
    for doc in docs {
        name := ""
        val := ""
        try name := doc.Name
        try val := doc.Value
        if (InStr(val, "gemini.google.com", false) || InStr(name, "Gemini", false))
            geminiDocs.Push(doc)
    }
    if (geminiDocs.Length = 0)
        return uia
    if (geminiDocs.Length = 1)
        return geminiDocs[1]
    title := ""
    try {
        hwnd := 0
        try hwnd := uia.BrowserId
        if (hwnd)
            title := WinGetTitle("ahk_id " hwnd)
    } catch {
        title := ""
    }
    if (!title) {
        try title := WinGetTitle("A")
        catch
            title := ""
    }
    best := 0
    bestLen := 0
    for doc in geminiDocs {
        name := ""
        try name := doc.Name
        if (name && title && InStr(title, name, false) = 1 && StrLen(name) > bestLen) {
            best := doc
            bestLen := StrLen(name)
        }
    }
    if (best)
        return best
    for doc in geminiDocs {
        off := true
        try off := doc.GetPropertyValue(UIA.Property.IsOffscreen)
        if (!off)
            return doc
    }
    return geminiDocs[geminiDocs.Length]
}

; Footer "Copy" lives on response-container, outside presented-response-container.
; The fence control lives in group.copy-button inside the presented body (or legacy code-block).
Gemini_IsResponseFooterControl(el) {
    if (!IsObject(el))
        return false
    inBody := false
    node := el
    loop 8 {
        if (!IsObject(node))
            return false
        cls := ""
        try cls := node.ClassName
        if (Gemini_ClassHasToken(cls, "presented-response-container") || Gemini_ClassHasToken(cls, "code-block"))
            inBody := true
        if (Gemini_ClassHasToken(cls, "response-container"))
            return !inBody
        parent := 0
        try parent := node.Parent
        catch
            return false
        node := parent
    }
    return false
}

Gemini_IsFenceCopyButton(el) {
    if (!IsObject(el))
        return false
    name := ""
    try name := el.Name
    catch
        return false
    if (!IsGeminiCopyCodeButton(name) || IsGeminiCopyResponseButton(name))
        return false
    if (Gemini_IsResponseFooterControl(el))
        return false
    hasHost := false
    inBody := false
    node := el
    loop 8 {
        if (!IsObject(node))
            break
        cls := ""
        try cls := node.ClassName
        if (Gemini_ClassHasToken(cls, "copy-button"))
            hasHost := true
        if (Gemini_ClassHasToken(cls, "presented-response-container") || Gemini_ClassHasToken(cls, "code-block"))
            inBody := true
        if (Gemini_ClassHasToken(cls, "response-container") && inBody)
            break
        parent := 0
        try parent := node.Parent
        catch
            break
        node := parent
    }
    return hasHost && inBody
}

Gemini_SameElement(a, b) {
    if (!IsObject(a) || !IsObject(b))
        return false
    try {
        return UIA.RuntimeIdToString(a.GetRuntimeId()) = UIA.RuntimeIdToString(b.GetRuntimeId())
    } catch {
        return false
    }
}

Gemini_HasAncestor(el, ancestor) {
    if (!IsObject(el) || !IsObject(ancestor))
        return false
    node := el
    loop 12 {
        if (!IsObject(node))
            return false
        if (Gemini_SameElement(node, ancestor))
            return true
        parent := 0
        try parent := node.Parent
        catch
            return false
        node := parent
    }
    return false
}

; Last response-container on the page (not presented-response-container).
Gemini_LastResponseContainer(root) {
    if (!IsObject(root))
        return 0
    containers := []
    try {
        groups := root.FindAll({ ClassName: "response-container", matchmode: "Substring" })
        for g in groups {
            cls := ""
            try cls := g.ClassName
            if (Gemini_ClassHasToken(cls, "response-container"))
                containers.Push(g)
        }
    } catch {
    }
    if (containers.Length = 0)
        return 0
    return containers[containers.Length]
}

Gemini_ScrollElementIntoView(el) {
    node := el
    loop 4 {
        if (!IsObject(node))
            return
        try {
            if (node.GetPropertyValue(UIA.Property.IsScrollItemPatternAvailable)) {
                node.ScrollItemPattern.ScrollIntoView()
                return
            }
        } catch {
        }
        parent := 0
        try parent := node.Parent
        catch
            return
        node := parent
    }
}

; Code copy must be the fence control (group.copy-button inside the presented response).
; The whole-answer "Copy" sits lower, on the response footer. Invoke only: a coordinate
; click lands on that footer when the fence button's bounds overlap it.
Gemini_InvokeAcceptedCodeCopyButton(el) {
    if (!IsObject(el))
        return false
    name := ""
    try name := el.Name
    catch
        return false
    if (!Gemini_IsFenceCopyButton(el)) {
        if (!IsGeminiCopyCodeButton(name) || IsGeminiCopyResponseButton(name) || Gemini_IsResponseFooterControl(el))
            return false
    }
    Gemini_ScrollElementIntoView(el)
    try {
        if (el.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)) {
            el.InvokePattern.Invoke()
            return true
        }
    } catch {
    }
    return false
}

; Return count of Gemini "Copy response" buttons. Uses centralized GetGeminiCopyButtonsArray. Caller must ensure tab is active and scrolled to bottom.
GetGeminiCopyButtonCount(uia) {
    return GetGeminiCopyButtonsArray(uia).Length
}

; True if at least one "Show more options" / "More options" exists (last response can then offer read aloud). Same labels as GeminiTriggerReadAloud.
GeminiHasMoreOptionsForResponse(uia) {
    allMoreOptionsButtons := []
    try {
        allMoreOptionsButtons := uia.FindAll({ Name: "Show more options" })
        try {
            moreOpt := uia.FindAll({ Name: "More options" })
            for btn in moreOpt
                allMoreOptionsButtons.Push(btn)
        } catch {
        }
    } catch {
    }
    if (allMoreOptionsButtons.Length = 0) {
        try {
            allMenuItems := uia.FindAll({ Type: 50011 })
            for menuItem in allMenuItems {
                name := menuItem.Name
                if (name = "Show more options" || name = "More options" || InStr(name, "Show more options", false) = 1 ||
                InStr(name, "More options", false) = 1)
                    allMoreOptionsButtons.Push(menuItem)
            }
        } catch {
        }
    }
    return allMoreOptionsButtons.Length > 0
}

; Find consumer Gemini browser window (excludes Gemini Enterprise)
GetGeminiWindowHwnd() {
    try {
        for hwnd in WinGetList("ahk_exe chrome.exe") {
            try {
                if IsConsumerGeminiChromeTitle(WinGetTitle("ahk_id " hwnd))
                    return hwnd
            } catch {
                ; Silently skip invalid windows
            }
        }
    } catch {
        ; Silently handle WinGetList errors
    }
    return 0
}

; --- Phase 4: Event-driven new Chrome window detection (SetWinEventHook) ---------
; Wait for a new top-level Chrome window that is not in existingHwnds. Returns hwnd or 0.
; Hook is an accelerator only: WinGetList is polled every tick so a window created
; before the hook was installed (or missed by EVENT_OBJECT_CREATE) is found immediately.
WaitForNewChromeWindow(existingHwnds, timeoutMs) {
    hHook := 0
    try {
        if (GEMINI_USE_WIN_EVENT_HOOK && timeoutMs > 0) {
            global g_GeminiCreatedHwnds := []
            cb := CallbackCreate(GeminiWinEventProc, "F Fast", 7)
            hHook := DllCall("user32\SetWinEventHook", "UInt", GEMINI_EVENT_OBJECT_CREATE, "UInt",
                GEMINI_EVENT_OBJECT_CREATE, "Ptr", 0, "Ptr", cb, "UInt", 0, "UInt", 0, "UInt", 0, "Ptr")
        }
        deadline := A_TickCount + Max(timeoutMs, 0)
        loop {
            hwnd := FindNewChromeWindowHwnd(existingHwnds)
            if (hwnd)
                return hwnd
            if (hHook) {
                for createdHwnd in g_GeminiCreatedHwnds {
                    if (IsNewChromeWindowHwnd(createdHwnd, existingHwnds) && WinExist("ahk_id " createdHwnd))
                        return createdHwnd
                }
                g_GeminiCreatedHwnds := []
            }
            if (A_TickCount >= deadline)
                return 0
            Sleep 80
        }
    } finally {
        if (hHook)
            DllCall("user32\UnhookWinEvent", "Ptr", hHook)
    }
}

FindNewChromeWindowHwnd(existingHwnds) {
    try {
        for hwnd in WinGetList("ahk_exe chrome.exe") {
            if IsNewChromeWindowHwnd(hwnd, existingHwnds)
                return hwnd
        }
    } catch {
    }
    return 0
}

IsNewChromeWindowHwnd(hwnd, existingHwnds) {
    if (!hwnd)
        return false
    for existing in existingHwnds {
        if (existing = hwnd)
            return false
    }
    try {
        return (WinGetProcessName("ahk_id " hwnd) = "chrome.exe")
    } catch {
        return false
    }
}

GeminiWinEventProc(hWinEventHook, event, hwnd, idObject, idChild, idEventThread, dwmsEventTime) {
    global g_GeminiCreatedHwnds
    if (event = GEMINI_EVENT_OBJECT_CREATE && idObject = GEMINI_OBJID_WINDOW && hwnd)
        g_GeminiCreatedHwnds.Push(hwnd)
}

; =============================================================================
; Show tab indicator banner (1 = blue, 2 = yellow): square, center of active-window monitor.
; Delegates to Utils for identical behavior as #!+U tab-switching in Utils.ahk. Auto-hides after 700 ms.
; =============================================================================
ShowGeminiTabBanner(tabNumber, geminiHwnd := 0) {
    ShowSingleCharTabBanner_Utils(tabNumber)
}
