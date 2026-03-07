#Requires AutoHotkey v2.0
#SingleInstance Force

; --- Includes ----------------------------------------------------------------
#include UIA-v2\Lib\UIA.ahk
#include UIA-v2\Lib\UIA_Browser.ahk
#include %A_ScriptDir%\env.ahk
#include %A_ScriptDir%\Utils.ahk

; --- Config ---------------------------------------------------------------
; Path to the file containing the initial prompt Gemini should receive.
PROMPT_FILE := A_ScriptDir "\data\Gemini_Prompt.txt"

; Copy response button names (EN/PT). Excludes "Copy prompt" / "Copiar prompt" which are different controls.
GEMINI_COPY_RESPONSE_NAMES := ["Copy", "Copiar"]

; #region agent log
; Append one NDJSON line to debug-a6ae60.log for copy-index debugging.
_GeminiCopyDebugLog(location, message, data := "") {
    logPath := A_ScriptDir "\debug-a6ae60.log"
    try {
        q := Chr(34)
        dataStr := ""
        if (data && Type(data) = "Map") {
            for k, v in data {
                vStr := String(v)
                esc := StrReplace(StrReplace(StrReplace(vStr, "\", "\\"), q, "\" q), "`n", "\n")
                dataStr .= (dataStr ? "," : "") q k q ":" q esc q
            }
        }
        line := "{" q "sessionId" q ":" q "a6ae60" q "," q "location" q ":" q location q "," q "message" q ":" q message q "," q "data" q ":{" dataStr "}," q "timestamp" q ":" A_TickCount "}"
        FileAppend line "`n", logPath
    } catch
        return
}
; #endregion

; --- Refactor: threshold constants (no magic numbers) ---------------------------------
; Timeouts (ms)
GEMINI_ACTIVATE_WAIT_MS := 2000
GEMINI_SCROLL_SETTLE_MS := 350
GEMINI_UIA_SETTLE_MS := 120
GEMINI_TAB_SWITCH_MS := 150
GEMINI_MENU_OPEN_MS := 200
GEMINI_READ_ALOUD_SETTLE_MS := 1500
GEMINI_PROMPT_FOCUS_MS := 300
GEMINI_FIRST_LAUNCH_POLL_MS := 300
GEMINI_FIRST_LAUNCH_MAX_LOOPS := 35
GEMINI_TITLE_READY_MS := 6000
GEMINI_TITLE_POLL_MS := 250
GEMINI_WAIT_BUTTON_POLL_MS := 250
GEMINI_WAIT_BUTTON_TIMEOUT_MS := 15000
GEMINI_ASYNC_POLL_MS := 500
GEMINI_COPY_RETRY_SLEEP_MS := 400
GEMINI_STREAM_GONE_VERIFY_MS := 200
GEMINI_STREAM_GONE_LOOPS := 4
; Retries
GEMINI_ASYNC_LOOKUP_MAX_RETRIES := 60   ; 60 * 500ms = 30s
GEMINI_DELAYED_SUBMIT_MAX_RETRIES := 300  ; 300 * 500ms = 150s
GEMINI_ASYNC_TTS_MAX_RETRIES := 60
GEMINI_COPY_MAX_RETRIES := 3
; Pronunciation result banner: long timeout so user can read (ms)
GEMINI_PRONUNCIATION_BANNER_TIMEOUT_MS := 50000
; Post-copy sync: max wait for clipboard change (sequence-number detection). Fallback if detection fails (ms).
GEMINI_POST_COPY_SYNC_TIMEOUT_MS := 2000
; Poll interval when waiting for clipboard sequence number to change (ms). Low value = minimal latency.
GEMINI_CLIPBOARD_POLL_MS := 10
; Performance instrumentation (set to true to log latencies to script dir)
GEMINI_PERF_LOG_ENABLED := false
GEMINI_PERF_LOG_PATH := A_ScriptDir "\.cursor\gemini_perf.log"

; Feature flags for refactor phases (set false to fall back to legacy behavior)
GEMINI_USE_WIN_EVENT_HOOK := true
GEMINI_USE_PYTHON_IPC := false

; Phase 7: Python daemon IPC (local socket). Daemon must be started separately; AHK connects on demand.
GEMINI_PYTHON_DAEMON_PORT := 29512
GEMINI_PYTHON_IPC_TIMEOUT_MS := 5000

; --- Phase 4: WinEvent hook constants (user32) ---------------------------------
; EVENT_OBJECT_CREATE = 0x8000, OBJID_WINDOW = 0
GEMINI_EVENT_OBJECT_CREATE := 0x8000
GEMINI_OBJID_WINDOW := 0

; --- Clipboard sequence number (user32) – O(1) change detection ------------------
; Returns Windows clipboard sequence number (increments on any clipboard change). Used to wait for clipboard update with minimal latency instead of fixed sleep.
GetClipboardSequenceNumber() {
    try
        return DllCall("user32\GetClipboardSequenceNumber", "UInt")
    catch
        return 0
}

; --- Performance instrumentation (Phase 1) -------------------------------------
; Log latency for a flow; no-op if GEMINI_PERF_LOG_ENABLED is false. Non-blocking.
GeminiPerfLog(flowName, startTick) {
    if (!GEMINI_PERF_LOG_ENABLED)
        return
    elapsed := A_TickCount - startTick
    try FileAppend(FormatTime(, "yyyy-MM-dd HH:mm:ss") "`t" flowName "`t" elapsed "`n", GEMINI_PERF_LOG_PATH)
    catch
        return
}

; Non-blocking error log for empty catch blocks (Phase 6). Does not interrupt workflow.
GeminiLogError(context, err := "") {
    try FileAppend(FormatTime(, "yyyy-MM-dd HH:mm:ss") " [Gemini] " context (err ? " " err : "") "`n", A_ScriptDir "\.cursor\gemini_err.log"
    )
    catch
        return
}

; --- Phase 7: Python IPC client (socket bridge) ----------------------------------
; Send a request to the persistent Python daemon; returns response map or empty on failure/timeout.
; Protocol: 4-byte big-endian length + UTF-8 JSON. No RunWait; daemon runs separately.
; When GEMINI_USE_PYTHON_IPC is false or daemon is unavailable, returns "" (AHK continues with local path).
GeminiIpcSend(requestObj) {
    if (!GEMINI_USE_PYTHON_IPC)
        return ""
    ; Placeholder: full implementation uses Winsock DllCall (connect to 127.0.0.1:GEMINI_PYTHON_DAEMON_PORT,
    ; send framed JSON, recv with timeout). When implemented, return parsed response object.
    return ""
}

; --- Phase 2: UIA control type constants (strict integer; no string coercion) ----
; UIA ControlType: Button=50000, MenuItem=50011. Use these instead of magic numbers.
UIA_ControlType_Button := 50000
UIA_ControlType_MenuItem := 50011

; --- Phase 2: Centralized UIA discovery (single tree walk per element type) ------
; Returns array of Copy response buttons in document order; empty array on error. Caller must ensure tab active and scrolled to bottom.
GetGeminiCopyButtonsArray(uia) {
    prevBatch := -1
    try prevBatch := A_BatchLines
    catch
        prevBatch := -1
    try {
        out := []
        A_BatchLines := -1
        allButtons := uia.FindAll({ Type: "Button" })
        nameMatchCount := 0
        classMatchCount := 0
        sampleNames := ""
        sampleClasses := ""
        for button in allButtons {
            isNameMatch := IsGeminiCopyResponseButton(button.Name)
            if (isNameMatch)
                nameMatchCount++
            if (isNameMatch && (InStr(button.ClassName, "icon-button") || InStr(button.ClassName, "mdc-button"))) {
                classMatchCount++
                out.Push(button)
            }
            if (out.Length <= 2 && isNameMatch) {
                sampleNames .= (sampleNames ? " | " : "") (button.Name ? button.Name : "(empty)")
                sampleClasses .= (sampleClasses ? " | " : "") (button.ClassName ? SubStr(button.ClassName, 1, 80) :
                    "(empty)")
            }
        }
        if (out.Length = 0) {
            allButtons := uia.FindAll({ Type: "Button" })
            for button in allButtons {
                if (IsGeminiCopyResponseButton(button.Name))
                    out.Push(button)
            }
        }
        ; #region agent log
        if (out.Length > 0) {
            lastBtn := out[out.Length]
            _GeminiCopyDebugLog("GetGeminiCopyButtonsArray", "exit", Map("totalCopyButtons", String(out.Length),
            "lastName", lastBtn.Name ? lastBtn.Name : "(empty)", "lastClass", lastBtn.ClassName ? SubStr(lastBtn.ClassName,
                1, 100) : "(empty)"))
            if (out.Length >= 2)
                _GeminiCopyDebugLog("GetGeminiCopyButtonsArray", "penultimate", Map("idx", String(out.Length - 1),
                "name", out[out.Length - 1].Name ? out[out.Length - 1].Name : "(empty)"))
        } else
            _GeminiCopyDebugLog("GetGeminiCopyButtonsArray", "exit", Map("totalCopyButtons", "0"))
        ; #endregion
    } catch as err {
        try A_BatchLines := prevBatch
        return []
    }
    try A_BatchLines := prevBatch
    return out
}

; Returns the last Copy button (last response) or 0. Uses centralized discovery.
GetLastGeminiCopyButton(uia) {
    arr := GetGeminiCopyButtonsArray(uia)
    ; #region agent log
    _GeminiCopyDebugLog("GetLastGeminiCopyButton", "return", Map("arrLen", String(arr.Length), "usingIdx", arr.Length >
    0 ? String(arr.Length) : "0"))
    ; #endregion
    return (arr.Length > 0) ? arr[arr.Length] : 0
}

; Copy last Gemini message with retry using exponential backoff (Phase 6). Returns true if clipboard changed.
; options/geminiHwnd same as CopyLastGeminiMessageToClipboard. maxRetries includes first attempt.
CopyLastGeminiMessageWithRetry(options := "", geminiHwnd := 0, maxRetries := GEMINI_COPY_MAX_RETRIES) {
    baseDelay := GEMINI_COPY_RETRY_SLEEP_MS
    contentBefore := A_Clipboard
    loop maxRetries {
        if (CopyLastGeminiMessageToClipboard(options, geminiHwnd) && A_Clipboard != "" && A_Clipboard != contentBefore)
            return true
        if (A_Index < maxRetries)
            Sleep baseDelay * (1 << (A_Index - 1))
    }
    return false
}

; Find Pause or Resume button (TTS). which = "Pause" or "Resume". Returns element or 0. Uses UIA_ControlType_Button.
FindGeminiPauseResumeButton(uia, which) {
    try {
        btn := uia.FindFirst({ Name: which, Type: UIA_ControlType_Button })
        if (btn)
            return btn
        btn := uia.FindFirst({ Type: "Button", Name: which })
        if (btn)
            return btn
        allButtons := uia.FindAll({ Type: UIA_ControlType_Button })
        for button in allButtons {
            if (button.Name = which || InStr(button.Name, which, false) = 1) {
                if (InStr(button.ClassName, "tts-button") || InStr(button.ClassName, "mdc-icon-button"))
                    return button
            }
        }
    } catch {
        return 0
    }
    return 0
}

; --- Phase 3: Scoped discovery (conversation panel / main pane to avoid full-doc traversal) ---
; Get root for scoped search: main pane when available, else full tree. Returns element to call FindAll on.
GetGeminiSearchRoot(uia) {
    try {
        root := uia.GetCurrentMainPaneElement()
        if (root)
            return root
    } catch {
    }
    return uia
}

; Find "Show more options" / "More options" buttons scoped to main pane when possible. Returns array.
GetGeminiMoreOptionsButtonsScoped(uia) {
    root := GetGeminiSearchRoot(uia)
    out := []
    try {
        els := root.FindAll({ Name: "Show more options" })
        for e in els
            out.Push(e)
        try {
            moreOpt := root.FindAll({ Name: "More options" })
            for btn in moreOpt
                out.Push(btn)
        } catch {
        }
    } catch {
    }
    if (out.Length = 0) {
        try {
            allMenuItems := root.FindAll({ Type: UIA_ControlType_MenuItem })
            for menuItem in allMenuItems {
                name := menuItem.Name
                if (name = "Show more options" || name = "More options" || InStr(name, "Show more options", false) = 1 ||
                InStr(name, "More options", false) = 1)
                    out.Push(menuItem)
            }
        } catch {
        }
    }
    return out
}

; Find "Text to speech" menu item. Returns element or 0. Uses UIA_ControlType_MenuItem.
FindGeminiTextToSpeechMenuItem(uia) {
    try {
        mi := uia.FindFirst({ Name: "Text to speech", Type: UIA_ControlType_MenuItem })
        if (mi)
            return mi
        mi := uia.FindFirst({ Type: "MenuItem", Name: "Text to speech" })
        if (mi)
            return mi
        allMenuItems := uia.FindAll({ Type: UIA_ControlType_MenuItem })
        for menuItem in allMenuItems {
            if (menuItem.Name = "Text to speech" || InStr(menuItem.Name, "Text to speech", false) = 1) {
                if (InStr(menuItem.ClassName, "mat-mdc-menu-item"))
                    return menuItem
            }
        }
        for menuItem in allMenuItems {
            if (menuItem.Name = "Text to speech" || InStr(menuItem.Name, "Text to speech", false) = 1)
                return menuItem
        }
    } catch {
        return 0
    }
    return 0
}

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
                ; #region agent log
                _GeminiCopyDebugLog("GetLastCopyButtonCached", "cached", Map("hwnd", String(geminiHwnd)))
                ; #endregion
                return GeminiState._lastCopyButton
            } catch
                GeminiState.Invalidate()
        }
        ; #region agent log
        _GeminiCopyDebugLog("GetLastCopyButtonCached", "fresh", Map("hwnd", String(geminiHwnd)))
        ; #endregion
        btn := GetLastGeminiCopyButton(uia)
        if (btn && geminiHwnd) {
            GeminiState._hwnd := geminiHwnd
            GeminiState._lastCopyButton := btn
        }
        return btn
    }
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

; Find Gemini browser window (case-insensitive contains match for "gemini")
GetGeminiWindowHwnd() {
    try {
        for hwnd in WinGetList("ahk_exe chrome.exe") {
            try {
                if InStr(WinGetTitle("ahk_id " hwnd), "gemini", false)
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
; Get work area (left, top, right, bottom) of the monitor that contains the given window
; =============================================================================
GetWorkAreaForWindow(hwnd) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return ""
    try {
        WinGetPos(&winX, &winY, &winW, &winH, "ahk_id " hwnd)
        centerX := winX + winW / 2
        centerY := winY + winH / 2
        n := MonitorGetCount()
        loop n {
            MonitorGet(A_Index, &L, &T, &R, &B)
            if (centerX >= L && centerX < R && centerY >= T && centerY < B) {
                MonitorGetWorkArea(A_Index, &wLeft, &wTop, &wRight, &wBottom)
                return { left: wLeft, top: wTop, right: wRight, bottom: wBottom }
            }
        }
    } catch {
    }
    return ""
}

; =============================================================================
; Get 1-based active tab index and tab count in Chrome via UIA (tab bar TabItem elements).
; Returns {index: n, count: c} on success; 0 if detection fails. Used when #!+i is triggered
; on an existing Gemini window (banner shown only when count >= 2).
;
; Implementation proposals for Chrome tab identification (for testing/verification):
;   A) UIA (current): Use UIA_Browser.GetTabs() and GetTab("") with SelectionItemIsSelected,
;      then match selected tab by RuntimeId in the tab list to get 1-based index. Reliable for
;      Chrome/Edge when the tab bar is exposed to UIA.
;   B) Window title: WinGetTitle() often includes the active tab's page title; does not give
;      tab index, but could be used to show "current tab title" in a text banner instead of a number.
;   C) Chrome DevTools Protocol (CDP): Requires Chrome started with --remote-debugging-port;
;      can query tabs via HTTP/json. More setup, not used here.
; =============================================================================
GetChromeActiveTabIndex(uia) {
    try {
        uia.GetCurrentMainPaneElement()
        tabs := uia.GetTabs()
        if (!tabs.Length)
            return 0
        current := uia.GetTab("")
        if (!current)
            return 0
        rid := current.RuntimeId
        for i, tab in tabs {
            try {
                if (tab.RuntimeId = rid)
                    return { index: i, count: tabs.Length }
            } catch {
                continue
            }
        }
    } catch {
    }
    return 0
}

; =============================================================================
; Show tab indicator banner (1 = blue, 2 = yellow): square, center of active-window monitor.
; Delegates to Utils for identical behavior as #!+U tab-switching in Utils.ahk. Auto-hides after 700 ms.
; =============================================================================
ShowGeminiTabBanner(tabNumber, geminiHwnd := 0) {
    ShowSingleCharTabBanner_Utils(tabNumber)
}

; =============================================================================
; Helper function to show a notification using the standard loading indicator (passive, auto-hide).
; =============================================================================
ShowNotification(message, durationMs := 500, bgColor := "FFFF00", fontColor := "000000", fontSize := 17) {
    StandardLoadingBar_Show(message, BANNER_ACCENT_INTERMEDIATE, { passive: true, fontSize: fontSize })
    StandardLoadingBar_Hide(durationMs)
}

; =============================================================================
; Copy completed chime (single beep, debounced)
; =============================================================================
PlayCopyCompletedChime() {
    try {
        static lastTick := 0
        if (A_TickCount - lastTick < 1500)
            return
        lastTick := A_TickCount

        if (IsSoundEnabled()) {
            SoundPlay(A_ScriptDir . "\sounds\copy.wav")
        }
    } catch {
        ; Silently ignore errors
    }
}

; =============================================================================
; Small Loading Indicator Helpers (delegate to standard loading bar in Utils)
; =============================================================================
ShowSmallLoadingIndicator(state := "⏳ Loading…", bgColor := BANNER_ACCENT_INTERMEDIATE, centerOnHwnd := 0, textWidth :=
    500, fontSize :=
    17) {
    global g_StandardLoadingBarGui
    if (g_StandardLoadingBarGui)
        StandardLoadingBar_Update(state)
    else
        StandardLoadingBar_Show(state, bgColor, { passive: true, centerOnHwnd: centerOnHwnd, textWidth: textWidth,
            fontSize: fontSize })
}

HideSmallLoadingIndicator() {
    StandardLoadingBar_Hide(0)
}

WaitForButtonAndShowSmallLoading(buttonNames, stateText := "⏳ Loading…", timeout := GEMINI_WAIT_BUTTON_TIMEOUT_MS) {
    try cUIA := UIA_Browser()
    catch {
        ; Silently ignore UIA browser errors
        return
    }
    start := A_TickCount
    deadline := (timeout > 0) ? (start + timeout) : 0
    btn := ""
    indicatorShown := false
    buttonEverFound := false
    buttonDisappeared := false
    while (timeout <= 0 || A_TickCount < deadline) {
        btn := ""
        for n in buttonNames {
            try {
                btn := cUIA.FindElement({ Name: n, Type: "Button" })
            } catch {
                btn := ""
            }
            if btn
                break
        }
        if btn {
            buttonEverFound := true
            if (!indicatorShown) {
                StandardLoadingBar_Show(stateText, BANNER_ACCENT_INTERMEDIATE)
                indicatorShown := true
            }
            while btn && (timeout <= 0 || A_TickCount < deadline) {
                Sleep GEMINI_WAIT_BUTTON_POLL_MS
                btn := ""
                for n in buttonNames {
                    try {
                        btn := cUIA.FindElement({ Name: n, Type: "Button" })
                    } catch {
                        btn := ""
                    }
                    if btn
                        break
                }
            }
            if !btn
                buttonDisappeared := true
            break
        }
        Sleep GEMINI_WAIT_BUTTON_POLL_MS
    }
    ; Play completion sound only for actual AI responses when we saw the button and it disappeared
    try {
        if (buttonEverFound && buttonDisappeared && InStr(StrLower(stateText), "transcrib") = 0)
            PlayCopyCompletedChime()
    } catch {
        ; Silently ignore errors
    }
    StandardLoadingBar_Hide(0)
}

; =============================================================================
; Helper function to center mouse on the active window
; =============================================================================
CenterMouse() {
    Sleep 200
    Send("#!+q")
}

; --- Hotkeys ----------------------------------------------------------------

; Reusable: activate Gemini, handle Pause/Resume, then optionally copy last message and trigger "Text to speech".
; copyFirst: true = copy last response then read aloud (#!+o); false = only read aloud (#!+7).
; useTrashTab: when true, explicitly target the second Gemini tab (trash tab) instead of the main tab.
GeminiTriggerReadAloud(copyFirst := true, useTrashTab := false) {
    t0 := A_TickCount
    ; Step 1: Activate Gemini window globally
    SetTitleMatchMode(2)
    if hwnd := GetGeminiWindowHwnd() {
        try {
            WinActivate("ahk_id " hwnd)
        } catch {
            ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
            return
        }
    }
    if !WinWaitActive("ahk_exe chrome.exe", , GEMINI_ACTIVATE_WAIT_MS // 1000)
        return
    Sleep GEMINI_TAB_SWITCH_MS

    ; When requested (#!+o trash tab), explicitly switch to the second Gemini tab.
    ; Chrome convention: Ctrl+2 selects the second tab in the window.
    if (useTrashTab) {
        Send("^2")
        Sleep GEMINI_TAB_SWITCH_MS
        ShowGeminiTabBanner(2, hwnd)
    }

    ; Step 2: Check if "Pause" button exists (if reading is active, pause it)
    uia := UIA_Browser()
    Sleep GEMINI_UIA_SETTLE_MS

    pauseButton := FindGeminiPauseResumeButton(uia, "Pause")
    if (pauseButton) {
        pauseButton.Click()
        ShowNotification("Paused", 800, "FFFF00", "000000", 24)
        Send "!{Tab}"
        return
    }

    resumeButton := FindGeminiPauseResumeButton(uia, "Resume")
    if (resumeButton) {
        resumeButton.Click()
        ShowNotification("Resumed", 800, "FFFF00", "000000", 24)
        Send "!{Tab}"
        return
    }

    ; Step 3: If copyFirst, find and click the last Copy button; else just scroll so last response is in view
    Send "^{End}"
    Sleep GEMINI_SCROLL_SETTLE_MS

    if (copyFirst) {
        lastCopyButton := GeminiState.GetLastCopyButtonCached(uia, hwnd)
        if (lastCopyButton) {
            lastCopyButton.Click()
            PlayCopyCompletedChime()
        }
    }

    ; Step 4: Find the final "More options" / "Show more options" in the Gemini response tree (bottom-up).
    ; We target only the most recent Gemini response to avoid reading older messages. See gemini-tree.txt for tree structure.
    centerHwnd := WinExist("A")
    StandardLoadingBar_Show(copyFirst ? "🔍 Finding read aloud button and copying..." :
        "🔍 Finding read aloud button...", BANNER_ACCENT_INTERMEDIATE, { passive: true, centerOnHwnd: centerHwnd })
    Sleep GEMINI_WAIT_BUTTON_POLL_MS

    allMoreOptionsButtons := GetGeminiMoreOptionsButtonsScoped(uia)

    if (allMoreOptionsButtons.Length = 0) {
        StandardLoadingBar_Hide(0)
        return
    }

    ; Bottom-up: select the last instance in the response tree = most recent response only.
    ; 1) Prefer button with the largest bottom Y (true bottom of page = final response).
    ; 2) Fallback: last element in FindAll order (document/tree order = last in tree).
    lastMoreOptionsButton := 0
    highestBottomY := -1
    for moreOptionsButton in allMoreOptionsButtons {
        try {
            btnPos := moreOptionsButton.Location
            bottomY := btnPos.y + btnPos.h
            if (bottomY > highestBottomY) {
                highestBottomY := bottomY
                lastMoreOptionsButton := moreOptionsButton
            }
        } catch {
            continue
        }
    }
    if (!lastMoreOptionsButton && allMoreOptionsButtons.Length > 0)
        lastMoreOptionsButton := allMoreOptionsButtons[allMoreOptionsButtons.Length]
    if (!lastMoreOptionsButton) {
        StandardLoadingBar_Hide(0)
        return
    }

    try {
        lastMoreOptionsButton.Click()
        Sleep GEMINI_MENU_OPEN_MS

        textToSpeechMenuItem := FindGeminiTextToSpeechMenuItem(uia)
        if (textToSpeechMenuItem) {
            textToSpeechMenuItem.Click()
            Sleep GEMINI_MENU_OPEN_MS
        } else {
            Send "{Down}"
            Sleep GEMINI_MENU_OPEN_MS
            Send "{Enter}"
        }
    } catch {
    }

    StandardLoadingBar_Hide(0)

    Sleep GEMINI_READ_ALOUD_SETTLE_MS

    isReading := (FindGeminiPauseResumeButton(uia, "Pause") != 0)

    if (!isReading) {
        ShowNotification("Retrying read aloud...", 800, "FFFF00", "000000", 24)
        try {
            lastMoreOptionsButton.Click()
            Sleep GEMINI_MENU_OPEN_MS

            textToSpeechMenuItem := FindGeminiTextToSpeechMenuItem(uia)
            if (textToSpeechMenuItem) {
                textToSpeechMenuItem.Click()
                Sleep GEMINI_MENU_OPEN_MS
            } else {
                Send "{Down}"
                Sleep GEMINI_MENU_OPEN_MS
                Send "{Enter}"
            }
        } catch {
        }
    }

    GeminiPerfLog("read_aloud", t0)
    ShowNotification(copyFirst ? "Copied & Reading aloud" : "Reading aloud", 800, "FFFF00", "000000", 24)
    Send "!{Tab}"
}

; Win+Alt+Shift+O : Read aloud the last message in Gemini (or Pause/Resume if already reading)
#!+o:: {
    try {
        ; Standard behavior: operate on the currently active Gemini tab.
        GeminiTriggerReadAloud()
    } catch Error as e {
        ;
    }
}

; Copy last Gemini message to clipboard. Used by #!+p and by async pronunciation flow.
; options.restoreWindow (default true): send !{Tab} after copy. options.playChimeAndNotify (default true): play chime and show "Copied!".
; options.alreadyActive (default false): when true, skip activation; assume Gemini is already the active window (use UIA_Browser() with no arg).
; geminiHwnd: optional; if 0, uses GetGeminiWindowHwnd(). Returns true if copy succeeded, false otherwise.
CopyLastGeminiMessageToClipboard(options := "", geminiHwnd := 0) {
    t0 := A_TickCount
    restoreWindow := (options = "" || !options.HasProp("restoreWindow")) ? true : options.restoreWindow
    playChimeAndNotify := (options = "" || !options.HasProp("playChimeAndNotify")) ? true : options.playChimeAndNotify
    alreadyActive := (options != "" && options.HasProp("alreadyActive")) ? options.alreadyActive : false
    try {
        SetTitleMatchMode(2)
        if !geminiHwnd
            geminiHwnd := GetGeminiWindowHwnd()
        if !geminiHwnd {
            GeminiPerfLog("copy", t0)
            return false
        }
        if (!alreadyActive) {
            try {
                WinActivate("ahk_id " geminiHwnd)
            } catch {
                ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
                return false
            }
            if !WinWaitActive("ahk_exe chrome.exe", , GEMINI_ACTIVATE_WAIT_MS // 1000)
                return false
            Sleep GEMINI_TAB_SWITCH_MS
        }

        ; Scroll to bottom *before* UIA so the last response is in the tree and we go down the chat.
        Send "^{End}"
        Sleep GEMINI_SCROLL_SETTLE_MS

        uia := alreadyActive ? UIA_Browser() : UIA_Browser("ahk_id " geminiHwnd)
        Sleep GEMINI_UIA_SETTLE_MS

        lastCopyButton := GeminiState.GetLastCopyButtonCached(uia, geminiHwnd)

        if (!lastCopyButton) {
            GeminiPerfLog("copy", t0)
            return false
        }
        lastCopyButton.Click()
        if !ClipWait(2) {
            GeminiPerfLog("copy", t0)
            return false
        }
        if (playChimeAndNotify) {
            PlayCopyCompletedChime()
            ShowNotification("Copied!", 800, "FFFF00", "000000", 24)
        }
        if (restoreWindow)
            Send "!{Tab}"
        GeminiPerfLog("copy", t0)
        return true
    } catch {
        GeminiPerfLog("copy", t0)
        return false
    }
}

; Win+Alt+Shift+P : Click the last Copy button in Gemini (activates Gemini, scrolls to bottom with Ctrl+End, then copies last response)
; Works in EN ("Copy") and PT ("Copiar") UI. Uses tree order: last Copy button in the UI tree = last response.
#!+p:: {
    try {
        t0 := A_TickCount
        if (!CopyLastGeminiMessageToClipboard())
            ShowNotification("Copy failed – ensure Gemini is open and has a response", 2500, "FF6666", "FFFFFF", 22)
        GeminiPerfLog("hotkey_copy", t0)
    } catch as err {
        ShowNotification("Copy error: " (err.Message ? err.Message : "unknown"), 2500, "FF6666", "FFFFFF", 22)
    }
}

; Custom message so WindowManagement.ahk can trigger copy without Send (Send does not trigger hotkeys in another script).
WM_COPY_LAST_GEMINI := 0x8001
; Start background completion monitor for Ctrl+Alt+Win+L (wParam = originalHwnd, lParam = geminiHwnd). Sent from Utils.ahk.
WM_START_DELAYED_SUBMIT_MONITOR := 0x8002
; Path for bridge to verify that Copy Last Response (same as #!+p) actually succeeded
GEMINI_COPY_RESULT_PATH := A_ScriptDir "\.cursor\gemini_copy_result.txt"

OnMessage(WM_COPY_LAST_GEMINI, copyFromBridge)
OnMessage(WM_START_DELAYED_SUBMIT_MONITOR, handleStartDelayedSubmitMonitor)
handleStartDelayedSubmitMonitor(wParam, lParam, msg, hwnd) {
    GeminiDelayedSubmitMonitorStart(wParam, lParam)
}
copyFromBridge(*) {
    ; Guarantee layer: write result so bridge can confirm we copied Gemini's last response (same path as #!+p).
    try
        FileDelete(GEMINI_COPY_RESULT_PATH)
    try
        FileAppend("0", GEMINI_COPY_RESULT_PATH)
    r := CopyLastGeminiMessageToClipboard({ restoreWindow: false, playChimeAndNotify: false })
    try
        FileDelete(GEMINI_COPY_RESULT_PATH)
    try
        FileAppend(r ? "1" : "0", GEMINI_COPY_RESULT_PATH)
}

; =============================================================================
; TTS from selection – Win+Alt+Shift+7: copy selection, send "repeat exactly" to Gemini, then trigger read aloud
; =============================================================================
#!+7:: {
    (GeminiAsyncTTS()).Start()
}

; =============================================================================
; Get Pronunciation
; Hotkey: Win+Alt+Shift+8 — async: submit to Gemini, restore focus, show result in banner when ready
; =============================================================================
#!+8:: {
    (GeminiAsyncLookup()).Start()
}

; =============================================================================
; Initialize Gemini window on first-time opening
; =============================================================================
SendPromptToActiveGeminiTab(promptText) {
    try {
        if (StrLen(promptText) = 0)
            return false

        uia := UIA_Browser()
        Sleep 300

        promptField := FindGeminiPromptField(uia)
        if (!promptField)
            return false

        promptField.SetFocus()
        Sleep 100
        if (!promptField.HasKeyboardFocus) {
            try
                promptField.Click()
            Sleep 100
        }

        oldClip := A_Clipboard
        A_Clipboard := ""
        A_Clipboard := promptText
        ClipWait 1
        Send("^v")
        Sleep 100
        Send("{Enter}")
        Sleep 100
        A_Clipboard := oldClip
        return true
    } catch {
        return false
    }
}

InitializeGeminiFirstTime() {
    try {
        ; Show banner to inform user
        StandardLoadingBar_Show("📤 Opening Gemini (2 tabs)...", BANNER_ACCENT_INTERMEDIATE)

        ; Remember existing Chrome windows so we can find the one we're about to create
        existingChromeHwnds := []
        try {
            for hwnd in WinGetList("ahk_exe chrome.exe")
                existingChromeHwnds.Push(hwnd)
        } catch {
        }

        ; Run Chrome with new window and two Gemini tabs
        Run "chrome.exe --new-window https://gemini.google.com/ https://gemini.google.com/"
        Sleep 700   ; Give the system time to start Chrome before waiting for it

        ; Find the newly created Chrome window: event-driven hook or polling fallback
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
        if !WinWaitActive("ahk_id " geminiHwnd, , 4) {
            StandardLoadingBar_Hide(0)
            return
        }
        ; Wait for the first tab to load so the title contains "Gemini" (timeout-bounded condition wait)
        SetTitleMatchMode(2)
        start := A_TickCount
        while (A_TickCount - start < GEMINI_TITLE_READY_MS) {
            try {
                if InStr(WinGetTitle("ahk_id " geminiHwnd), "Gemini", false)
                    break
            } catch {
            }
            Sleep GEMINI_TITLE_POLL_MS
        }
        Sleep 550   ; Give window and tabs time to fully settle

        ; Read initial prompt from external file
        promptText := ""
        try promptText := FileRead(PROMPT_FILE, "UTF-8")
        if (StrLen(promptText) = 0)
            promptText := "hey, what's up?"

        ; Update banner status
        StandardLoadingBar_Show("📤 Sending prompt to Gemini tabs...", BANNER_ACCENT_INTERMEDIATE)

        geminiHwnd := GetGeminiWindowHwnd()
        ; Ensure first tab is active and send prompt (no tab banner during initial launch)
        Send("^1")
        Sleep 280
        SendPromptToActiveGeminiTab(promptText)

        ; Switch to second tab and send the same prompt
        Send("^2")
        Sleep 280
        SendPromptToActiveGeminiTab(promptText)

        ; Hide banner on success
        StandardLoadingBar_Hide(0)
    } catch Error as err {
        ; Hide banner on error
        StandardLoadingBar_Hide(0)
    }
}

; =============================================================================
; Open Gemini
; Hotkey: Win+Alt+Shift+I
; =============================================================================
#!+i:: {
    t0 := A_TickCount
    SetTitleMatchMode(2)
    if hwnd := GetGeminiWindowHwnd() {
        try {
            WinActivate("ahk_id " hwnd)
        } catch {
            ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        if WinWaitActive("ahk_id " hwnd, , GEMINI_ACTIVATE_WAIT_MS // 1000) {
            Sleep GEMINI_UIA_SETTLE_MS   ; Let the window and Chrome content settle before UIA attaches
            ; Bind UIA to this window so we never attach to a different Chrome window
            uia := UIA_Browser("ahk_id " hwnd)
            Sleep GEMINI_UIA_SETTLE_MS   ; UIA settle time (align with CopyLastGeminiMessageToClipboard)

            ; Show current active tab only when this window already has two Gemini tabs (not during initial launch).
            ; Brief extra delay so Chrome tab bar is ready for UIA; retry once if first attempt fails (timing).
            Sleep 80
            tabInfo := GetChromeActiveTabIndex(uia)
            if (!tabInfo) {
                Sleep 150
                tabInfo := GetChromeActiveTabIndex(uia)
            }
            ; Show tab-position banner: use UIA index when available, otherwise assume position 1 so banner always appears
            tabPosition := (tabInfo && tabInfo.count >= 2 && tabInfo.index) ? tabInfo.index : 1
            ShowSingleCharTabBanner_Utils(tabPosition)

            ; Find the anchor element: "Open upload file menu" button
            ; Combined search: Try exact match first, then case-insensitive (most efficient)
            anchorButton := 0
            try {
                anchorButton := uia.FindFirst({ Type: UIA_ControlType_Button, Name: "Open upload file menu",
                    ControlType: "Button" })
                if (!anchorButton) {
                    anchorButton := uia.FindFirst({ Type: UIA_ControlType_Button, Name: "Open upload file menu", cs: false })
                }
            } catch {
            }

            ; Fallback: Only use expensive FindAll if first two strategies failed
            if (!anchorButton) {
                try {
                    allButtons := uia.FindAll({ Type: UIA_ControlType_Button })
                    for button in allButtons {
                        try {
                            if (InStr(button.Name, "Open upload file menu", false)) {
                                anchorButton := button
                                break
                            }
                        } catch {
                            continue
                        }
                    }
                } catch {
                }
            }

            if (anchorButton) {
                ; Focus the anchor button (do NOT click) and navigate back
                try {
                    anchorButton.SetFocus()
                    Sleep 25   ; minimal wait for focus
                    SendInput "+{Tab}"  ; Use SendInput for faster keystroke
                    Sleep 15   ; minimal delay for navigation
                } catch {
                    ; Anchor strategy failed; will use direct prompt field below
                }
            }
            ; Ensure the prompt field actually has keyboard focus (same as SendPromptToActiveGeminiTab)
            promptField := FindGeminiPromptField(uia)
            if (promptField) {
                try {
                    promptField.SetFocus()
                    Sleep 100
                    if (!promptField.HasKeyboardFocus) {
                        try promptField.Click()
                        Sleep 100
                    }
                } catch {
                }
                if (IsSoundEnabled())
                    SoundPlay(A_ScriptDir . "\sounds\gemini-focused.wav")
            }
            GeminiPerfLog("activation", t0)
        }
    } else {
        InitializeGeminiFirstTime()
    }
}

; =============================================================================
; GeminiAsyncLookup – async pronunciation lookup (Win+Alt+Shift+8)
; User keeps focus; timer polls for completion; result shown in banner.
; =============================================================================
class GeminiAsyncLookup {
    static PronunciationPrompt :=
        "Below, you will find a word or phrase. I'd like you to answer in five sections: the 1st section you will repeat the word twice. For each time you repeat, use a point to finish the phrase. The 2nd section should have the definition of the word (You should also say each part of speech does the different definitions belong to). The 3rd section should have the pronunciation of this word using the International Phonetic Alphabet characters (for American English).The 4th section should have the same word applied in a real sentence (put that in quotations, so I can identify that). In the 5th, Write down the translation of the word into Portuguese. Please, do not title any section. Thanks!"

    __New() {
        this.OriginalHwnd := 0
        this.GeminiHwnd := 0
        this.TimerCallback := ""
        this.RetryCount := 0
        this.MaxRetries := 60   ; 60 * 500ms = 30s timeout
        this.ButtonEverFound := false
    }

    Start() {
        this.OriginalHwnd := WinExist("A")
        if !this.OriginalHwnd
            return
        ; Show loading banner immediately, centered on the monitor where this window is (with warning)
        StandardLoadingBar_Show("⏳ Loading…", BANNER_ACCENT_INTERMEDIATE)

        A_Clipboard := ""
        Send "^c"
        if !ClipWait(2)
            return
        SetTitleMatchMode(2)
        this.GeminiHwnd := GetGeminiWindowHwnd()
        if !this.GeminiHwnd {
            StandardLoadingBar_Hide(0)
            return
        }
        try {
            WinActivate("ahk_id " this.GeminiHwnd)
        } catch {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        if !WinWaitActive("ahk_exe chrome.exe", , 2) {
            StandardLoadingBar_Hide(0)
            return
        }
        ; For pronunciation lookup (#!+8), always use the trash tab (second Gemini tab).
        ; Chrome convention: Ctrl+2 selects the second tab in the window.
        Send("^2")
        Sleep 150
        ShowGeminiTabBanner(2, this.GeminiHwnd)
        uia := UIA_Browser()
        Sleep 300
        promptField := FindGeminiPromptField(uia)
        if (!promptField) {
            StandardLoadingBar_Hide(0)
            return
        }
        promptField.SetFocus()
        Sleep 100
        if (!promptField.HasKeyboardFocus) {
            try promptField.Click()
            Sleep 100
        }
        ; Switch to Fast model before the prompt (enough for this task)
        Send("@fast ")
        Sleep 200
        searchString := GeminiAsyncLookup.PronunciationPrompt
        A_Clipboard := searchString . "`n`nContent: " . A_Clipboard
        Sleep 100
        Send("^a")
        Sleep 500
        Send("^v")
        Sleep 500
        Send("{Enter}")
        Sleep 300
        ; Go back to the window where you triggered the hotkey so you can keep working (activate only; do not WinRestore or we lose maximized state)
        origHwnd := this.OriginalHwnd
        try {
            if WinExist("ahk_id " origHwnd) {
                WinActivate("ahk_id " origHwnd)
                WinWaitActive("ahk_id " origHwnd, , 1)
            }
        } catch {
            if (WinExist("ahk_id " origHwnd))
                WinActivate("ahk_id " origHwnd)
        }
        this.RetryCount := 0
        this.TimerCallback := this.CheckCompletion.Bind(this)
        SetTimer(this.TimerCallback, 500)
    }

    CheckCompletion() {
        this.RetryCount++
        if (this.RetryCount > this.MaxRetries) {
            SetTimer(this.TimerCallback, 0)
            StandardLoadingBar_Hide(0)
            return
        }
        ; Poll in background using raw UIA (no UIA_Browser) so the library never activates Gemini
        btn := ""
        buttonNames := ["Stop streaming", "Interromper transmissão", "Stop response"]
        try {
            root := UIA.ElementFromHandle(this.GeminiHwnd)
            for n in buttonNames {
                try {
                    btn := root.FindElement({ Name: n, Type: "Button" })
                } catch {
                    btn := ""
                }
                if btn
                    break
            }
        } catch {
            return
        }

        if btn {
            this.ButtonEverFound := true
            return   ; Still streaming
        }

        ; If button was found and now is gone, verify it's truly finished (timeout-bounded)
        if (this.ButtonEverFound) {
            isTrulyGone := true
            loop GEMINI_STREAM_GONE_LOOPS {
                Sleep GEMINI_STREAM_GONE_VERIFY_MS
                try {
                    for n in buttonNames {
                        if root.ElementExist({ Name: n, Type: "Button" }) {
                            isTrulyGone := false
                            break
                        }
                    }
                } catch
                    isTrulyGone := true
                if !isTrulyGone
                    break
            }

            if isTrulyGone {
                SetTimer(this.TimerCallback, 0)
                ; Use the same sound as Shift keys.ahk for consistency
                try {
                    if (IsSoundEnabled()) {
                        SoundPlay(A_ScriptDir . "\sounds\gemini-completion.wav")
                    }
                } catch {
                    PlayCopyCompletedChime()
                }
                this.RetrieveResponse()
            }
        }
    }

    RetrieveResponse() {
        ; Activate Gemini once, then copy with retry (exponential backoff) without switching back until done
        try {
            WinActivate("ahk_id " this.GeminiHwnd)
        } catch {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        if !WinWaitActive("ahk_exe chrome.exe", , 2) {
            StandardLoadingBar_Hide(0)
            return
        }
        copyOpt := { restoreWindow: false, playChimeAndNotify: false, alreadyActive: true }
        seqBefore := GetClipboardSequenceNumber()
        if !CopyLastGeminiMessageWithRetry(copyOpt, this.GeminiHwnd) {
            StandardLoadingBar_Hide(0)
            return
        }
        ; Wait for clipboard change via sequence number (O(1) per check); proceed as soon as it changes instead of fixed delay.
        syncElapsed := 0
        while (syncElapsed < GEMINI_POST_COPY_SYNC_TIMEOUT_MS) {
            if (GetClipboardSequenceNumber() != seqBefore)
                break
            Sleep GEMINI_CLIPBOARD_POLL_MS
            syncElapsed += GEMINI_CLIPBOARD_POLL_MS
        }
        ; Use clipboard content only after change detected (or timeout) so the banner shows current content.
        bannerText := A_Clipboard
        WinActivate("ahk_id " this.OriginalHwnd)
        StandardLoadingBar_Hide(0)
        this.ShowResultBanner(bannerText)
    }

    ShowResultBanner(text) {
        if (!text || StrLen(Trim(text)) = 0)
            return
        state := "ℹ " . text
        closeNoOp(*) {
        }
        closeKeys := Map("Enter", closeNoOp, "Escape", closeNoOp, "E", closeNoOp)
        StandardLoadingBar_ShowWithKeys(state, closeKeys, GEMINI_PRONUNCIATION_BANNER_TIMEOUT_MS, this.OriginalHwnd, "",
            BANNER_ACCENT_INTERMEDIATE, 600,
            17, "", false,
            "[Enter] [E] Close")
    }
}

; =============================================================================
; GeminiDelayedSubmitMonitor – background completion monitor for Ctrl+Alt+Win+L
; Reuses #!+8 completion detection; on completion shows "Copy? [N] [R]" with 4s timeout (N = no copy, R = copy + read aloud).
; =============================================================================
class GeminiDelayedSubmitMonitor {
    __New() {
        this.OriginalHwnd := 0
        this.GeminiHwnd := 0
        this.TimerCallback := ""
        this.RetryCount := 0
        this.MaxRetries := 300   ; 300 * 500ms = 150s timeout (covers Gemini Pro deep thinking)
        this.ButtonEverFound := false
        this.CopyBannerGui := ""
        this.CopyTimeoutTimer := ""
    }

    Start(originalHwnd, geminiHwnd) {
        if (!originalHwnd || !geminiHwnd)
            return
        this.OriginalHwnd := originalHwnd
        this.GeminiHwnd := geminiHwnd
        this.RetryCount := 0
        this.ButtonEverFound := false
        this.TimerCallback := this.CheckCompletion.Bind(this)
        SetTimer(this.TimerCallback, 500)
    }

    CheckCompletion() {
        this.RetryCount++
        if (this.RetryCount > this.MaxRetries) {
            SetTimer(this.TimerCallback, 0)
            return
        }
        btn := ""
        buttonNames := ["Stop streaming", "Interromper transmissão", "Stop response"]
        try {
            root := UIA.ElementFromHandle(this.GeminiHwnd)
            for n in buttonNames {
                try {
                    btn := root.FindElement({ Name: n, Type: "Button" })
                } catch {
                    btn := ""
                }
                if btn
                    break
            }
        } catch {
            return
        }

        if btn {
            this.ButtonEverFound := true
            return
        }

        if (this.ButtonEverFound) {
            isTrulyGone := true
            loop GEMINI_STREAM_GONE_LOOPS {
                Sleep GEMINI_STREAM_GONE_VERIFY_MS
                try {
                    for n in buttonNames {
                        if root.ElementExist({ Name: n, Type: "Button" }) {
                            isTrulyGone := false
                            break
                        }
                    }
                } catch
                    isTrulyGone := true
                if !isTrulyGone
                    break
            }

            if isTrulyGone {
                SetTimer(this.TimerCallback, 0)
                try {
                    if (IsSoundEnabled())
                        SoundPlay(A_ScriptDir . "\sounds\gemini-completion.wav")
                } catch {
                    PlayCopyCompletedChime()
                }
                this.ShowCopyDecisionBanner()
            }
        }
    }

    ShowCopyDecisionBanner() {
        this.CopyBannerGui := ""
        this.CopyTimeoutTimer := ""
        copyKeyCallbacks := Map("N", this.CancelCopy.Bind(this), "Y", this.DoCopyOnly.Bind(this), "R", this.CopyAndReadAloud
        .Bind(this), "E", this.CancelCopy.Bind(this))
        StandardLoadingBar_ShowWithKeys("❓ Copy response?", copyKeyCallbacks, 4000, this.OriginalHwnd, this.DoCopyOnTimeout
            .Bind(this), BANNER_ACCENT_INTERMEDIATE, 380, 17, "", false, "[Y] Copy  [N] No  [R] Copy+Read  [E] Close")
    }

    ; Y key: copy latest response only (same as timeout default), then close banner and restore focus.
    DoCopyOnly(*) {
        this.DoCopyOnTimeout()
    }

    ; Shared cleanup: clear Gemini-side state only. Key/timer unregister and overlay hide are handled by Utils.
    CleanupCopyBanner() {
        this.CopyBannerGui := ""
        this.CopyTimeoutTimer := ""
    }

    CancelCopy(*) {
        this.CleanupCopyBanner()
    }

    DoCopyOnTimeout(*) {
        this.CleanupCopyBanner()
        GeminiState.Invalidate()
        try {
            WinActivate("ahk_id " this.GeminiHwnd)
        } catch {
            if (WinExist("ahk_id " this.OriginalHwnd))
                WinActivate("ahk_id " this.OriginalHwnd)
            return
        }
        if !WinWaitActive("ahk_exe chrome.exe", , 2) {
            if (WinExist("ahk_id " this.OriginalHwnd))
                WinActivate("ahk_id " this.OriginalHwnd)
            return
        }
        copyOpt := { restoreWindow: false, playChimeAndNotify: false, alreadyActive: true }
        if CopyLastGeminiMessageWithRetry(copyOpt, this.GeminiHwnd)
            PlayCopyCompletedChime()
        if (WinExist("ahk_id " this.OriginalHwnd))
            WinActivate("ahk_id " this.OriginalHwnd)
    }

    ; R key: copy last message and read it aloud, then restore focus (same tab as delayed submit).
    CopyAndReadAloud(*) {
        this.CleanupCopyBanner()
        GeminiState.Invalidate()
        try {
            WinActivate("ahk_id " this.GeminiHwnd)
        } catch {
            if (WinExist("ahk_id " this.OriginalHwnd))
                WinActivate("ahk_id " this.OriginalHwnd)
            return
        }
        if !WinWaitActive("ahk_exe chrome.exe", , 2) {
            if (WinExist("ahk_id " this.OriginalHwnd))
                WinActivate("ahk_id " this.OriginalHwnd)
            return
        }
        copyOpt := { restoreWindow: false, playChimeAndNotify: false, alreadyActive: true }
        if CopyLastGeminiMessageWithRetry(copyOpt, this.GeminiHwnd)
            PlayCopyCompletedChime()
        GeminiTriggerReadAloud(false, false)   ; read aloud only (already copied)
        if (WinExist("ahk_id " this.OriginalHwnd))
            WinActivate("ahk_id " this.OriginalHwnd)
    }
}

; Callable from Utils.ahk after successful auto-send (Ctrl+Alt+Win+L).
GeminiDelayedSubmitMonitorStart(originalHwnd, geminiHwnd) {
    (GeminiDelayedSubmitMonitor()).Start(originalHwnd, geminiHwnd)
}

; =============================================================================
; GeminiAsyncTTS – copy selection, send "repeat exactly" to Gemini, then trigger read aloud (Win+Alt+Shift+7)
; =============================================================================
class GeminiAsyncTTS {
    static TTSPrompt :=
        "Repeat the following text exactly as it is. Do not add any introduction, explanation, or markdown formatting. Just output the text itself:`n`n"
    static PostStreamingDelayMs := 600

    __New() {
        this.OriginalHwnd := 0
        this.GeminiHwnd := 0
        this.TimerCallback := ""
        this.RetryCount := 0
        this.MaxRetries := 60   ; 60 * 500ms = 30s timeout
        this.ButtonEverFound := false
        this.CopyCountAtSubmit := 0
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
        this.GeminiHwnd := GetGeminiWindowHwnd()
        if !this.GeminiHwnd {
            StandardLoadingBar_Hide(0)
            return
        }
        try {
            WinActivate("ahk_id " this.GeminiHwnd)
        } catch {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        if !WinWaitActive("ahk_exe chrome.exe", , 2) {
            StandardLoadingBar_Hide(0)
            return
        }
        ; For TTS from selection (#!+7), always use the trash tab (second Gemini tab) when sending the prompt.
        ; Chrome convention: Ctrl+2 selects the second tab in the window.
        Send("^2")
        Sleep 150
        ShowGeminiTabBanner(2, this.GeminiHwnd)
        uia := UIA_Browser()
        Sleep 300
        promptField := FindGeminiPromptField(uia)
        if (!promptField) {
            StandardLoadingBar_Hide(0)
            return
        }
        promptField.SetFocus()
        Sleep 100
        if (!promptField.HasKeyboardFocus) {
            try promptField.Click()
            Sleep 100
        }
        ; Paste prompt + selected text and submit
        A_Clipboard := GeminiAsyncTTS.TTSPrompt . A_Clipboard
        Sleep 100
        Send("^a")
        Sleep 500
        Send("^v")
        Sleep 500
        ; Record Copy button count before submit (used when multiple-message validation is needed).
        Send("^{End}")
        Sleep 350
        this.CopyCountAtSubmit := GetGeminiCopyButtonCount(uia)
        Send("{Enter}")
        Sleep 300
        ; Return focus to original window (activate only; do not WinRestore or we lose maximized state)
        origHwnd := this.OriginalHwnd
        try {
            if WinExist("ahk_id " origHwnd) {
                WinActivate("ahk_id " origHwnd)
                WinWaitActive("ahk_id " origHwnd, , 1)
            }
        } catch {
            if (WinExist("ahk_id " origHwnd))
                WinActivate("ahk_id " origHwnd)
        }
        this.RetryCount := 0
        this.TimerCallback := this.CheckCompletion.Bind(this)
        SetTimer(this.TimerCallback, 500)
    }

    CheckCompletion() {
        this.RetryCount++
        if (this.RetryCount > this.MaxRetries) {
            SetTimer(this.TimerCallback, 0)
            StandardLoadingBar_Hide(0)
            return
        }
        btn := ""
        buttonNames := ["Stop streaming", "Interromper transmissão", "Stop response"]
        try {
            root := UIA.ElementFromHandle(this.GeminiHwnd)
            for n in buttonNames {
                try {
                    btn := root.FindElement({ Name: n, Type: "Button" })
                } catch {
                    btn := ""
                }
                if btn
                    break
            }
        } catch {
            return
        }

        if btn {
            this.ButtonEverFound := true
            return
        }

        ; Layer 1: streaming stopped
        if (this.ButtonEverFound) {
            isTrulyGone := true
            loop GEMINI_STREAM_GONE_LOOPS {
                Sleep GEMINI_STREAM_GONE_VERIFY_MS
                try {
                    for n in buttonNames {
                        if root.ElementExist({ Name: n, Type: "Button" }) {
                            isTrulyGone := false
                            break
                        }
                    }
                } catch
                    isTrulyGone := true
                if !isTrulyGone
                    break
            }

            if isTrulyGone {
                SetTimer(this.TimerCallback, 0)
                StandardLoadingBar_Hide(0)
                ; Completion detection matches GeminiAsyncLookup (#!+8): Layer 1 only (Stop button gone). No extra Layer 2 so we don't miss completion.
                try {
                    if (IsSoundEnabled())
                        SoundPlay(A_ScriptDir . "\sounds\gemini-completion.wav")
                } catch {
                    PlayCopyCompletedChime()
                }
                ; Allow DOM to finish rendering, then activate trash tab and trigger read aloud (same pattern as #!+8 RetrieveResponse).
                Sleep(GeminiAsyncTTS.PostStreamingDelayMs)
                try {
                    WinActivate("ahk_id " this.GeminiHwnd)
                } catch {
                    return
                }
                if !WinWaitActive("ahk_exe chrome.exe", , 2)
                    return
                Send("^2")
                Sleep GEMINI_TAB_SWITCH_MS
                ShowGeminiTabBanner(2, this.GeminiHwnd)
                ; After TTS from selection (#!+7), read aloud from the trash tab (second Gemini tab).
                GeminiTriggerReadAloud(false, true)   ; read aloud only, no copy (text was just sent via #!+7)
            }
        }
    }
}
