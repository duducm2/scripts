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

; True if button name is the "Copy [last response]" button (EN or PT), not "Copy prompt".
IsGeminiCopyResponseButton(name) {
    if (!name || InStr(name, "prompt"))
        return false
    for n in GEMINI_COPY_RESPONSE_NAMES {
        if (name = n || InStr(name, n, false))
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
; Wait for a new Chrome window that is not in existingHwnds. Returns hwnd or 0. Timeout in ms.
; Uses EVENT_OBJECT_CREATE hook when GEMINI_USE_WIN_EVENT_HOOK is true; else falls back to polling.
WaitForNewChromeWindow(existingHwnds, timeoutMs) {
    if (GEMINI_USE_WIN_EVENT_HOOK && timeoutMs > 0) {
        global g_GeminiCreatedHwnds := []
        cb := CallbackCreate(GeminiWinEventProc, "F Fast", 7)
        hHook := DllCall("user32\SetWinEventHook", "UInt", GEMINI_EVENT_OBJECT_CREATE, "UInt",
            GEMINI_EVENT_OBJECT_CREATE, "Ptr", 0, "Ptr", cb, "UInt", 0, "UInt", 0, "UInt", 0, "Ptr")
        if (hHook) {
            start := A_TickCount
            while (A_TickCount - start < timeoutMs) {
                Sleep 80
                for hwnd in g_GeminiCreatedHwnds {
                    try {
                        if (WinGetProcessName("ahk_id " hwnd) = "chrome.exe") {
                            isNew := true
                            for ex in existingHwnds {
                                if (ex = hwnd) {
                                    isNew := false
                                    break
                                }
                            }
                            if (isNew) {
                                DllCall("user32\UnhookWinEvent", "Ptr", hHook)
                                return hwnd
                            }
                        }
                    } catch {
                    }
                }
                g_GeminiCreatedHwnds := []
            }
            DllCall("user32\UnhookWinEvent", "Ptr", hHook)
        }
    }
    ; Fallback: polling loop (legacy or when hook failed)
    deadline := A_TickCount + timeoutMs
    loop {
        if (A_TickCount >= deadline)
            return 0
        for hwnd in WinGetList("ahk_exe chrome.exe") {
            isNew := true
            for existing in existingHwnds {
                if (existing = hwnd) {
                    isNew := false
                    break
                }
            }
            if (isNew)
                return hwnd
        }
        Sleep GEMINI_FIRST_LAUNCH_POLL_MS
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

GeminiWaitForModelPickerReady(geminiHwnd := 0, timeoutMs := GEMINI_FIRST_LAUNCH_MODEL_READY_TIMEOUT_MS) {
    deadline := A_TickCount + (timeoutMs > 0 ? timeoutMs : GEMINI_FIRST_LAUNCH_MODEL_READY_TIMEOUT_MS)
    if (!geminiHwnd)
        geminiHwnd := FindGeminiChromeHwnd()
    while (A_TickCount < deadline) {
        try {
            uia := geminiHwnd ? UIA_Browser("ahk_id " geminiHwnd) : UIA_Browser()
            if (FindGeminiModePickerButton(uia))
                return true
        } catch {
        }
        Sleep GEMINI_FIRST_LAUNCH_MODEL_READY_POLL_MS
    }
    return false
}

GeminiSetModelForActiveTabWhenReady(modelName, geminiHwnd := 0) {
    if (!GeminiWaitForModelPickerReady(geminiHwnd))
        return false
    if (EnsureGeminiModelViaMenu(modelName))
        return true
    if (!GeminiWaitForModelPickerReady(geminiHwnd, GEMINI_FIRST_LAUNCH_MODEL_READY_POLL_MS * 4))
        return false
    return EnsureGeminiModelViaMenu(modelName)
}

GeminiConfigureFirstLaunchTabModels(geminiHwnd) {
    if (!geminiHwnd || !WinExist("ahk_id " geminiHwnd))
        return false

    if (!WinActive("ahk_id " geminiHwnd)) {
        try WinActivate("ahk_id " geminiHwnd)
        if (!WinWaitActive("ahk_id " geminiHwnd, , GEMINI_ACTIVATE_WAIT_MS / 1000))
            return false
    }

    StandardLoadingBar_Update("⚙️ Setting tab 1 model: 3.1 Pro...", BANNER_ACCENT_INTERMEDIATE)

    ; Ensure tab 1 is active before assigning Pro.
    try {
        uia := UIA_Browser("ahk_id " geminiHwnd)
        tabInfo := GetChromeActiveTabIndex(uia)
        if (tabInfo && tabInfo.index != 1) {
            Send "^1"
            Sleep GEMINI_FIRST_LAUNCH_TAB_SWITCH_SETTLE_MS
        }
    } catch {
    }

    proSet := GeminiSetModelForActiveTabWhenReady("3.1 Pro", geminiHwnd)

    ; Move to tab 2, set Fast, then return to tab 1 as requested.
    Send "^{Tab}"
    Sleep GEMINI_FIRST_LAUNCH_TAB_SWITCH_SETTLE_MS

    StandardLoadingBar_Update("⚙️ Setting tab 2 model: 3.1 Flash-Lite...", BANNER_ACCENT_INTERMEDIATE)
    fastSet := GeminiSetModelForActiveTabWhenReady("3.1 Flash-Lite", geminiHwnd)

    Send "^+{Tab}"
    Sleep GEMINI_FIRST_LAUNCH_TAB_SWITCH_SETTLE_MS

    if (proSet && fastSet) {
        StandardLoadingBar_Update("✅ Tab 1: 3.1 Pro, Tab 2: 3.1 Flash-Lite", BANNER_ACCENT_INTERMEDIATE)
        Sleep 200
        return true
    }
    return false
}
