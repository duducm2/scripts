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
