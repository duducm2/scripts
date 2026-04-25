#Requires AutoHotkey v2.0
#SingleInstance Force

; --- Includes ----------------------------------------------------------------
#include UIA-v2\Lib\UIA.ahk
#include UIA-v2\Lib\UIA_Browser.ahk
#include %A_ScriptDir%\env.ahk
#include %A_ScriptDir%\Utils.ahk
#include %A_ScriptDir%\aux\WMIPC.ahk
#include %A_ScriptDir%\aux\GeminiIPC.ahk

; --- Config ---------------------------------------------------------------
; Copy response button names (EN/PT). Excludes "Copy prompt" / "Copiar prompt" which are different controls.
GEMINI_COPY_RESPONSE_NAMES := ["Copy", "Copiar"]

; --- Refactor: threshold constants (no magic numbers) ---------------------------------
; Timeouts (ms)
GEMINI_ACTIVATE_WAIT_MS := 2000
GEMINI_SCROLL_SETTLE_MS := 350
GEMINI_UIA_SETTLE_MS := 120
GEMINI_TAB_SWITCH_MS := 150
GEMINI_MENU_OPEN_MS := 200
GEMINI_LISTEN_MENU_WAIT_MS := 1500   ; Bounded wait for Listen menu item after opening More options.
GEMINI_LISTEN_MENU_MAX_ATTEMPTS := 2 ; Max attempts to open menu and click Listen.
GEMINI_READ_ALOUD_SETTLE_MS := 1500
GEMINI_PROMPT_FOCUS_MS := 300
GEMINI_FIRST_LAUNCH_POLL_MS := 300
GEMINI_FIRST_LAUNCH_MAX_LOOPS := 35
GEMINI_TITLE_READY_MS := 6000
GEMINI_TITLE_POLL_MS := 250
GEMINI_WAIT_BUTTON_POLL_MS := 250
GEMINI_WAIT_BUTTON_TIMEOUT_MS := 15000
GEMINI_ASYNC_POLL_MS := 500
GEMINI_READ_ALOUD_START_POLL_MS := 150
GEMINI_COPY_RETRY_SLEEP_MS := 400
GEMINI_STREAM_GONE_VERIFY_MS := 200
GEMINI_STREAM_GONE_LOOPS := 4
; Retries
GEMINI_ASYNC_LOOKUP_MAX_RETRIES := 60   ; 60 * 500ms = 30s
GEMINI_DELAYED_SUBMIT_MAX_RETRIES := 300  ; 300 * 500ms = 150s
GEMINI_ASYNC_TTS_MAX_RETRIES := 60
GEMINI_READ_ALOUD_START_MAX_RETRIES := 10
GEMINI_COPY_MAX_RETRIES := 3
; Minimum clipboard length for Gemini-to-Cursor transfer (same as bridge validation)
GEMINI_TRANSFER_MIN_CLIPBOARD_LENGTH := 10
; Post-copy sync: max wait for clipboard change (sequence-number detection). Fallback if detection fails (ms).
GEMINI_POST_COPY_SYNC_TIMEOUT_MS := 2000
; Poll interval when waiting for clipboard sequence number to change (ms). Low value = minimal latency.
GEMINI_CLIPBOARD_POLL_MS := 10
; After copy, before Clip Angel favorite — lets newest clip appear as row 0 (ms).
GEMINI_POST_COPY_FAVORITE_DELAY_MS := 150
; Performance instrumentation (set to true to log latencies to script dir)
GEMINI_PERF_LOG_ENABLED := false
GEMINI_PERF_LOG_PATH := A_ScriptDir "\.cursor\gemini_perf.log"

; Feature flags for refactor phases (set false to fall back to legacy behavior)
GEMINI_USE_WIN_EVENT_HOOK := true
GEMINI_USE_PYTHON_IPC := false

; Phase 7: Python daemon IPC (Named Pipe). AHK starts the daemon on demand when enabled.
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
    } catch as err {
        try A_BatchLines := prevBatch
        return []
    }
    try A_BatchLines := prevBatch
    return out
}

; Returns the last Copy button (last response) or 0. Uses centralized discovery.
; Robust against Gemini UI changes that might add non-response Copy buttons after the chat history.
GetLastGeminiCopyButton(uia) {
    arr := GetGeminiCopyButtonsArray(uia)
    if (arr.Length = 0)
        return 0
    ; Prefer the visually lowest button (largest BoundingRectangle.t), which should correspond
    ; to the last assistant response in the chat, and ignore offscreen/zero-size elements.
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

; Copy last Gemini message with retry using exponential backoff (Phase 6). Returns true if clipboard changed.
; options/geminiHwnd same as CopyLastGeminiMessageToClipboard. maxRetries includes first attempt.
CopyLastGeminiMessageWithRetry(options := "", geminiHwnd := 0, maxRetries := GEMINI_COPY_MAX_RETRIES) {
    baseDelay := GEMINI_COPY_RETRY_SLEEP_MS
    loop maxRetries {
        if (CopyLastGeminiMessageToClipboard(options, geminiHwnd))
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

; --- Listen menu item (Type 50011): strict Name+Type, last instance = newest response ---
; Returns the last valid "Listen" MenuItem in the UI tree (belongs to newest message context menu).
; When multiple responses exist, only one context menu is open; the last matching MenuItem in tree
; order (by largest bottom Y, then last in FindAll order) is the one from that menu. No list-position
; targeting. Search scoped to GetGeminiSearchRoot first; fallback to full uia if popup is outside pane.
; Excludes stale/zero-size bounds. Returns element or 0.
GetLastGeminiListenMenuItem(uia) {
    listenItems := []
    root := GetGeminiSearchRoot(uia)
    try
        listenItems := root.FindAll({ Name: "Listen", Type: UIA_ControlType_MenuItem })
    catch
        listenItems := []
    if (listenItems.Length = 0) {
        try
            listenItems := uia.FindAll({ Name: "Listen", Type: UIA_ControlType_MenuItem })
        catch
            listenItems := []
    }
    if (listenItems.Length = 0)
        return 0
    ; Prefer candidate with largest bottom Y (last in layout); exclude zero-size/offscreen.
    lastEl := 0
    bestBottom := -1
    for item in listenItems {
        try {
            br := item.BoundingRectangle
        } catch {
            continue
        }
        if (!IsObject(br))
            continue
        if ((br.r - br.l) <= 0 || (br.b - br.t) <= 0)
            continue
        if (br.b > bestBottom) {
            bestBottom := br.b
            lastEl := item
        }
    }
    if (lastEl)
        return lastEl
    return listenItems[listenItems.Length]
}

; Poll for "Listen" MenuItem to appear after opening More options. Timeout-bounded; returns element or 0.
WaitForListenMenuItem(uia, timeoutMs := GEMINI_LISTEN_MENU_WAIT_MS) {
    deadline := A_TickCount + (timeoutMs > 0 ? timeoutMs : GEMINI_LISTEN_MENU_WAIT_MS)
    while (A_TickCount < deadline) {
        el := GetLastGeminiListenMenuItem(uia)
        if (el)
            return el
        Sleep GEMINI_WAIT_BUTTON_POLL_MS
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
; Shared background helpers
; =============================================================================
GeminiBackgroundSetTimer(task, callback, periodMs := GEMINI_ASYNC_POLL_MS) {
    GeminiBackgroundStopTimer(task)
    task.TimerCallback := callback
    SetTimer(task.TimerCallback, periodMs)
}

GeminiBackgroundStopTimer(task) {
    cb := ""
    try cb := task.TimerCallback
    catch
        cb := ""
    if (cb)
        SetTimer(cb, 0)
    try task.TimerCallback := ""
}

GeminiCanUseWMAutomationContext() {
    return WM_USE_DAEMON && WM_USE_PIPE_IPC && WM_USE_EVENT_HOOK_CACHE
}

GeminiBeginAutomationSwitch(reason := "", durationMs := 0) {
    if (!GeminiCanUseWMAutomationContext())
        return Map()
    try
        return WMIPC_BeginAutomationSwitch(reason, durationMs)
    catch
        return Map()
}

GeminiEndAutomationSwitch(reason := "") {
    if (!GeminiCanUseWMAutomationContext())
        return Map()
    try
        return WMIPC_EndAutomationSwitch(reason)
    catch
        return Map()
}

GeminiResolveOriginalHwnd(fallbackHwnd := 0) {
    originalHwnd := fallbackHwnd ? fallbackHwnd : WinExist("A")
    if (!GeminiCanUseWMAutomationContext())
        return originalHwnd
    try {
        ctx := WMIPC_GetAutomationContext()
        if (ctx.Has("foregroundHwnd") && Integer(ctx["foregroundHwnd"]) != 0)
            originalHwnd := Integer(ctx["foregroundHwnd"])
        title := ctx.Has("foregroundTitle") ? String(ctx["foregroundTitle"]) : ""
        if (InStr(title, "gemini", false) && ctx.Has("lastNonGeminiHwnd") && Integer(ctx["lastNonGeminiHwnd"]) != 0)
            return Integer(ctx["lastNonGeminiHwnd"])
    } catch {
    }
    return originalHwnd
}

GeminiActivateWindow(hwnd, waitMs := GEMINI_ACTIVATE_WAIT_MS) {
    if (!hwnd)
        return false
    GeminiBeginAutomationSwitch("gemini_activate_window", waitMs + GEMINI_TAB_SWITCH_MS + 1000)
    try {
        WinActivate("ahk_id " hwnd)
    } catch {
        return false
    }
    if (WinActive("ahk_id " hwnd))
        return true
    return !!WinWaitActive("ahk_id " hwnd, , waitMs / 1000)
}

GeminiRestoreWindow(hwnd, waitMs := 1000) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    GeminiBeginAutomationSwitch("gemini_restore_window", waitMs + 1000)
    return GeminiActivateWindow(hwnd, waitMs)
}

GeminiGetStreamingButtonNames() {
    static buttonNames := ["Stop streaming", "Interromper transmissão", "Stop response"]
    return buttonNames
}

GeminiReadRootFromHwnd(hwnd) {
    try
        return UIA.ElementFromHandle(hwnd)
    catch
        return 0
}

GeminiFindStreamingStopButton(root) {
    if (!IsObject(root))
        return 0
    for n in GeminiGetStreamingButtonNames() {
        try {
            btn := root.FindElement({ Name: n, Type: "Button" })
        } catch {
            btn := ""
        }
        if (btn)
            return btn
        try {
            btn := root.FindElement({ Name: n, Type: UIA_ControlType_Button })
        } catch {
            btn := ""
        }
        if (btn)
            return btn
    }
    return 0
}

GeminiVerifyStreamingStopped(geminiHwnd) {
    loop GEMINI_STREAM_GONE_LOOPS {
        Sleep GEMINI_STREAM_GONE_VERIFY_MS
        root := GeminiReadRootFromHwnd(geminiHwnd)
        if (!root)
            return true
        if (GeminiFindStreamingStopButton(root))
            return false
    }
    return true
}

GeminiMonitorStreamingTransition(task, onCompleteCallback) {
    task.RetryCount++
    if (task.RetryCount > task.MaxRetries) {
        GeminiBackgroundStopTimer(task)
        return "timeout"
    }
    root := GeminiReadRootFromHwnd(task.GeminiHwnd)
    if (!root)
        return "unavailable"
    if (GeminiFindStreamingStopButton(root)) {
        task.ButtonEverFound := true
        return "streaming"
    }
    if (!task.ButtonEverFound)
        return "waiting"
    if (!GeminiVerifyStreamingStopped(task.GeminiHwnd))
        return "streaming"
    GeminiBackgroundStopTimer(task)
    onCompleteCallback.Call()
    return "completed"
}

GetLastGeminiMoreOptionsButton(uia) {
    allMoreOptionsButtons := GetGeminiMoreOptionsButtonsScoped(uia)
    if (allMoreOptionsButtons.Length = 0)
        return 0
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
    return lastMoreOptionsButton
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

        ScriptSoundPlay(A_ScriptDir . "\sounds\copy.wav")
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

; --- Hotkeys ----------------------------------------------------------------

; Reusable launcher: activate Gemini asynchronously, handle Pause/Resume, then optionally
; copy the last message and trigger read aloud without holding the hotkey thread open.
GeminiTriggerReadAloud(copyFirst := true, useTrashTab := false, options := "") {
    return (GeminiAsyncReadAloud(copyFirst, useTrashTab, options)).Start()
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
; options.restoreWindow (default true): send !{Tab} after copy. When false, caret is moved to Ask Gemini via FocusGeminiAskFieldForHwnd (stays on Gemini tab).
; options.playChimeAndNotify (default true): play chime and show "Copied!".
; options.alreadyActive (default false): when true, skip activation; assume Gemini is already the active window (use UIA_Browser() with no arg).
; geminiHwnd: optional; if 0, uses GetGeminiWindowHwnd(). Returns true if copy succeeded, false otherwise.
CopyLastGeminiMessageToClipboard(options := "", geminiHwnd := 0) {
    t0 := A_TickCount
    restoreWindow := (options = "" || !options.HasProp("restoreWindow")) ? true : options.restoreWindow
    playChimeAndNotify := (options = "" || !options.HasProp("playChimeAndNotify")) ? true : options.playChimeAndNotify
    alreadyActive := (options != "" && options.HasProp("alreadyActive")) ? options.alreadyActive : false
    try {
        GeminiState.Invalidate()
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

        uia := alreadyActive ? UIA_Browser() : UIA_Browser("ahk_id " geminiHwnd)
        Sleep GEMINI_UIA_SETTLE_MS

        ; Scroll to bottom so the newest response controls are discoverable.
        Send "^{End}"
        Sleep GEMINI_SCROLL_SETTLE_MS

        lastCopyButton := GeminiState.GetLastCopyButtonCached(uia, geminiHwnd)

        if (!lastCopyButton) {
            GeminiPerfLog("copy", t0)
            return false
        }
        A_Clipboard := ""
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
        else
            FocusGeminiAskFieldForHwnd(geminiHwnd, false)
        GeminiPerfLog("copy", t0)
        return true
    } catch {
        GeminiPerfLog("copy", t0)
        return false
    }
}

; Win+Alt+Shift+P : Click the last Copy button in Gemini (activates Gemini, scrolls to bottom with Ctrl+End, then copies last response)
; Works in EN ("Copy") and PT ("Copiar") UI. Uses tree order: last Copy button in the UI tree = last response.
; Stays on Gemini (no Alt+Tab); leaves caret in Ask field + ready chime after copy.
#!+p:: {
    try {
        t0 := A_TickCount
        if (!CopyLastGeminiMessageToClipboard({ restoreWindow: false, playChimeAndNotify: true }))
            ShowNotification("Copy failed – ensure Gemini is open and has a response", 2500, "FF6666", "FFFFFF", 22)
        else if (hwnd := GetGeminiWindowHwnd())
            FocusGeminiAskFieldForHwnd(hwnd, true)
        GeminiPerfLog("hotkey_copy", t0)
    } catch as err {
        ShowNotification("Copy error: " (err.Message ? err.Message : "unknown"), 2500, "FF6666", "FFFFFF", 22)
    }
}

; Custom message so WindowManagement.ahk can trigger copy without Send (Send does not trigger hotkeys in another script).
WM_COPY_LAST_GEMINI := 0x8001
; Start background completion monitor for Ctrl+Alt+Win+L (wParam = originalHwnd, lParam = geminiHwnd). Sent from Utils.ahk.
WM_START_DELAYED_SUBMIT_MONITOR := 0x8002
; Stop any running delayed-submit monitor (e.g. when user chose S or N at 6s dictation confirm). Sent from Utils.ahk.
WM_STOP_DELAYED_SUBMIT_MONITOR := 0x8003
; Trigger read aloud from another script (e.g. D2C "Copy response?" R). Send does not trigger hotkeys in another script.
WM_TRIGGER_READ_ALOUD := 0x8004
; Path for bridge to verify that Copy Last Response (same as #!+p) actually succeeded
GEMINI_COPY_RESULT_PATH := A_ScriptDir "\.cursor\gemini_copy_result.txt"

OnMessage(WM_COPY_LAST_GEMINI, copyFromBridge)
OnMessage(WM_START_DELAYED_SUBMIT_MONITOR, handleStartDelayedSubmitMonitor)
OnMessage(WM_STOP_DELAYED_SUBMIT_MONITOR, handleStopDelayedSubmitMonitor)
OnMessage(WM_TRIGGER_READ_ALOUD, handleTriggerReadAloud)
handleStartDelayedSubmitMonitor(wParam, lParam, msg, hwnd) {
    GeminiDelayedSubmitMonitorStart(wParam, lParam)
}
handleStopDelayedSubmitMonitor(*) {
    GeminiDelayedSubmitMonitorStop()
}
handleTriggerReadAloud(wParam, lParam, msg, hwnd) {
    ; wParam 1: D2C already ran WM_COPY_LAST_GEMINI; skip internal Copy click, open Listen only.
    GeminiTriggerReadAloud(wParam = 0)
}
copyFromBridge(wParam, lParam, msg, hwnd) {
    geminiHwnd := Integer(lParam)
    ; Guarantee layer: write result so bridge can confirm we copied Gemini's last response (same path as #!+p).
    try
        FileDelete(GEMINI_COPY_RESULT_PATH)
    try
        FileAppend("0", GEMINI_COPY_RESULT_PATH)
    opts := { restoreWindow: false, playChimeAndNotify: false }
    if (geminiHwnd && WinActive("ahk_id " geminiHwnd))
        opts.alreadyActive := true
    r := CopyLastGeminiMessageToClipboard(opts, geminiHwnd)
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
    global g_StandardLoadingBarIsKeysOverlay
    if (g_StandardLoadingBarIsKeysOverlay) {
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        return
    }

    onSelect(lang) {
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        if (lang != "")
            (GeminiAsyncLookup(lang)).Start()
    }

    keyCallbacks := Map(
        "1", (*) => onSelect("pt"),
        "2", (*) => onSelect("en"),
        "3", (*) => onSelect("de"),
        "*Escape", (*) => onSelect("")
    )

    StandardLoadingBar_ShowWithKeys("❓ Select Translation Language", keyCallbacks, 0, 0, "", BANNER_ACCENT_INTERMEDIATE,
        450, 17, "", false, "[1] Portuguese  [2] English  [3] German  [Esc] Cancel")
}

; =============================================================================
; Initialize Gemini window on first-time opening
; =============================================================================
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
            ; Ensure the prompt field actually has keyboard focus (ready-to-type chime plays inside helper)
            Gemini_FocusPromptSameAsOpenHotkey(uia)
            GeminiPerfLog("activation", t0)
        }
    } else {
        InitializeGeminiFirstTime()
    }
}

; Direct focus to prompt text field (refactored for maximum efficiency)
Gemini_PlayReadyChime(minIntervalMs := 400) {
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

Gemini_FocusPromptSameAsOpenHotkey(uia) {
    if (!IsObject(uia)) {
        try {
            uia := UIA_Browser()
        } catch {
            return false
        }
    }

    localSettleMs := 120
    try {
        Sleep localSettleMs
        promptField := FindGeminiPromptField(uia)
        if (promptField) {
            ; Condition B: already focused/active (play immediately)
            try {
                if (promptField.HasKeyboardFocus) {
                    Gemini_PlayReadyChime()
                    return promptField
                }
            } catch {
            }

            ; Condition A: focus it now, then chime once caret is active
            try promptField.SetFocus()
            Sleep 100
            if (!promptField.HasKeyboardFocus) {
                try promptField.Click()
                Sleep 80
            }
            if (promptField.HasKeyboardFocus) {
                Gemini_PlayReadyChime()
                return promptField
            }
            return false
        }
    } catch {
        ; ignore and fall through
    }
    return false
}

; =============================================================================
; GeminiAsyncReadAloud – async read aloud / pause / resume (Win+Alt+Shift+O)
; =============================================================================
class GeminiAsyncReadAloud {
    __New(copyFirst := true, useTrashTab := false, options := "") {
        this.CopyFirst := copyFirst
        this.UseTrashTab := useTrashTab
        this.OriginalHwnd := (options != "" && options.HasProp("originalHwnd")) ? options.originalHwnd : 0
        this.GeminiHwnd := (options != "" && options.HasProp("geminiHwnd")) ? options.geminiHwnd : 0
        this.AlreadyActive := (options != "" && options.HasProp("alreadyActive")) ? options.alreadyActive : false
        this.TimerCallback := ""
        this.StartCallback := ""
        this.StepCallback := ""
        this.StartRetryAttempted := false
        this.StartRetryCount := 0
        this.StartTick := 0
        this.SkipCopy := false
        this.CopyRetryCount := 0
        this.MenuRetryCount := 0
        this.MenuPollCount := 0
        this.QueueTaskId := ""
        this.QueuePollCount := 0
        this.Uia := 0
    }

    Start() {
        if (!this.OriginalHwnd)
            this.OriginalHwnd := GeminiResolveOriginalHwnd()
        if (!this.StartTick)
            this.StartTick := A_TickCount
        SetTitleMatchMode(2)
        if (!this.GeminiHwnd)
            this.GeminiHwnd := GetGeminiWindowHwnd()
        if (!this.GeminiHwnd) {
            ShowNotification("Read aloud failed – Gemini is not open", 1800, "FF6666", "FFFFFF", 22)
            return false
        }
        if (this.AlreadyActive && WinActive("ahk_id " this.GeminiHwnd))
            return this.TryStartReadAloud(false)
        queuedTask := GeminiQueueBackgroundTask("ReadAloud", Map("geminiHwnd", this.GeminiHwnd, "originalHwnd",
            this.OriginalHwnd, "copyFirst", this.CopyFirst ? 1 : 0, "useTrashTab", this.UseTrashTab ? 1 : 0))
        if (queuedTask is Map && queuedTask.Has("taskId")) {
            this.QueueTaskId := String(queuedTask["taskId"])
            this.QueuePollCount := 0
            this.StartCallback := this.WaitForQueuedTask.Bind(this)
            SetTimer(this.StartCallback, -1)
            return true
        }
        this.StartCallback := this.Launch.Bind(this)
        SetTimer(this.StartCallback, -1)
        return true
    }

    ScheduleStep(callback, delayMs := 1) {
        this.ClearStep()
        this.StepCallback := callback
        SetTimer(this.StepCallback, -delayMs)
    }

    ClearStep() {
        cb := ""
        try cb := this.StepCallback
        catch
            cb := ""
        if (cb)
            SetTimer(cb, 0)
        this.StepCallback := ""
    }

    WaitForQueuedTask(*) {
        this.StartCallback := ""
        if (!this.QueueTaskId) {
            this.Launch()
            return
        }
        resp := GeminiIPC_GetTaskStatus(this.QueueTaskId)
        if (GeminiIPC_ResponseOk(resp)) {
            result := GeminiIPC_ResponseResultMap(resp)
            status := result.Has("status") ? String(result["status"]) : ""
            if (status = "ready") {
                this.QueueTaskId := ""
                this.Launch()
                return
            }
        } else {
            this.QueueTaskId := ""
            this.Launch()
            return
        }
        this.QueuePollCount++
        if (this.QueuePollCount >= 40) {
            this.QueueTaskId := ""
            this.Launch()
            return
        }
        this.StartCallback := this.WaitForQueuedTask.Bind(this)
        SetTimer(this.StartCallback, -50)
    }

    Launch(*) {
        this.StartCallback := ""
        this.TryStartReadAloud(false)
    }

    RetryLaunch(*) {
        this.StartCallback := ""
        this.TryStartReadAloud(true)
    }

    TryStartReadAloud(skipCopy := false) {
        try {
            this.SkipCopy := skipCopy
            this.CopyRetryCount := 0
            this.MenuRetryCount := 0
            this.MenuPollCount := 0
            this.Uia := 0
            GeminiState.Invalidate()
            if (!this.AlreadyActive || !WinActive("ahk_id " this.GeminiHwnd)) {
                PlayPreMovementWarning("Gemini")
                if !GeminiActivateWindow(this.GeminiHwnd, GEMINI_ACTIVATE_WAIT_MS) {
                    ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
                    return false
                }
                this.ScheduleStep(this.AfterActivation.Bind(this), GEMINI_TAB_SWITCH_MS)
                return true
            }
            this.ScheduleStep(this.AfterActivation.Bind(this), 1)
            return true
        } catch {
            this.Fail()
            return false
        }
    }

    AfterActivation(*) {
        this.ClearStep()
        if (this.UseTrashTab) {
            Send("^2")
            ShowGeminiTabBanner(2, this.GeminiHwnd)
            this.ScheduleStep(this.BuildUIA.Bind(this), GEMINI_TAB_SWITCH_MS)
            return
        }
        this.ScheduleStep(this.BuildUIA.Bind(this), 1)
    }

    BuildUIA(*) {
        this.ClearStep()
        try
            this.Uia := UIA_Browser("ahk_id " this.GeminiHwnd)
        catch {
            this.Fail()
            return false
        }
        this.ScheduleStep(this.InspectControls.Bind(this), GEMINI_UIA_SETTLE_MS)
        return true
    }

    InspectControls(*) {
        this.ClearStep()
        pauseButton := FindGeminiPauseResumeButton(this.Uia, "Pause")
        if (pauseButton) {
            try pauseButton.Click()
            ShowNotification("Paused", 800, "FFFF00", "000000", 24)
            this.RestoreOriginalFocus()
            return true
        }

        resumeButton := FindGeminiPauseResumeButton(this.Uia, "Resume")
        if (resumeButton) {
            try resumeButton.Click()
            ShowNotification("Resumed", 800, "FFFF00", "000000", 24)
            this.RestoreOriginalFocus()
            return true
        }

        Send "^{End}"
        this.ScheduleStep(this.AfterScroll.Bind(this), GEMINI_SCROLL_SETTLE_MS)
        return true
    }

    AfterScroll(*) {
        this.ClearStep()
        if (this.CopyFirst && !this.SkipCopy) {
            this.ScheduleStep(this.RunCopyAttempt.Bind(this), 1)
            return
        }
        this.BeginListenMenuPhase()
    }

    RunCopyAttempt(*) {
        this.ClearStep()
        copyOpt := { restoreWindow: false, playChimeAndNotify: false, alreadyActive: true }
        if (CopyLastGeminiMessageToClipboard(copyOpt, this.GeminiHwnd)) {
            PlayCopyCompletedChime()
            this.BeginListenMenuPhase()
            return true
        }
        this.CopyRetryCount++
        if (this.CopyRetryCount < GEMINI_COPY_MAX_RETRIES) {
            backoffMs := GEMINI_COPY_RETRY_SLEEP_MS * (1 << (this.CopyRetryCount - 1))
            this.ScheduleStep(this.RunCopyAttempt.Bind(this), backoffMs)
            return false
        }
        this.BeginListenMenuPhase()
        return false
    }

    BeginListenMenuPhase() {
        StandardLoadingBar_Show(this.CopyFirst && !this.SkipCopy ? "🔍 Finding read aloud button and copying..." :
            "🔍 Finding read aloud button...", BANNER_ACCENT_INTERMEDIATE, { passive: true, centerOnHwnd: 0 })
        this.ScheduleStep(this.OpenListenMenu.Bind(this), GEMINI_WAIT_BUTTON_POLL_MS)
    }

    OpenListenMenu(*) {
        this.ClearStep()
        lastMoreOptionsButton := GetLastGeminiMoreOptionsButton(this.Uia)
        if (!lastMoreOptionsButton) {
            StandardLoadingBar_Hide(0)
            this.Fail()
            return false
        }
        try lastMoreOptionsButton.Click()
        catch {
            this.HandleListenMenuRetry()
            return false
        }
        this.MenuPollCount := 0
        this.ScheduleStep(this.WaitForListenMenuReady.Bind(this), GEMINI_MENU_OPEN_MS)
        return true
    }

    WaitForListenMenuReady(*) {
        this.ClearStep()
        listenMenuItem := GetLastGeminiListenMenuItem(this.Uia)
        if (listenMenuItem) {
            try {
                listenMenuItem.Click()
                StandardLoadingBar_Hide(0)
                this.RestoreOriginalFocus()
                this.StartRetryCount := 0
                GeminiBackgroundSetTimer(this, this.CheckStarted.Bind(this), GEMINI_READ_ALOUD_START_POLL_MS)
                return true
            } catch {
                this.HandleListenMenuRetry()
                return false
            }
        }

        this.MenuPollCount++
        maxPolls := Ceil(GEMINI_LISTEN_MENU_WAIT_MS / GEMINI_WAIT_BUTTON_POLL_MS)
        if (this.MenuPollCount < maxPolls) {
            this.ScheduleStep(this.WaitForListenMenuReady.Bind(this), GEMINI_WAIT_BUTTON_POLL_MS)
            return false
        }
        this.HandleListenMenuRetry()
        return false
    }

    HandleListenMenuRetry() {
        this.ClearStep()
        this.MenuRetryCount++
        if (this.MenuRetryCount < GEMINI_LISTEN_MENU_MAX_ATTEMPTS) {
            SendEscape()
            this.ScheduleStep(this.OpenListenMenu.Bind(this), GEMINI_MENU_OPEN_MS)
            return
        }
        try StandardLoadingBar_Hide(0)
        this.Fail()
    }

    CheckStarted() {
        this.StartRetryCount++
        root := GeminiReadRootFromHwnd(this.GeminiHwnd)
        if (root) {
            try {
                if (FindGeminiPauseResumeButton(root, "Pause")) {
                    GeminiBackgroundStopTimer(this)
                    GeminiPerfLog("read_aloud", this.StartTick)
                    GeminiEndAutomationSwitch("gemini_read_aloud_started")
                    ShowNotification(this.CopyFirst ? "Copied & Reading aloud" : "Reading aloud", 800, "FFFF00",
                        "000000", 24)
                    if (this.GeminiHwnd && WinActive("ahk_id " this.GeminiHwnd))
                        FocusGeminiAskFieldForHwnd(this.GeminiHwnd, false)
                    return
                }
            } catch {
            }
        }

        if (this.StartRetryCount < GEMINI_READ_ALOUD_START_MAX_RETRIES)
            return

        GeminiBackgroundStopTimer(this)
        if (!this.StartRetryAttempted) {
            this.StartRetryAttempted := true
            ShowNotification("Retrying read aloud...", 800, "FFFF00", "000000", 24)
            this.StartCallback := this.RetryLaunch.Bind(this)
            SetTimer(this.StartCallback, -1)
            return
        }
        this.Fail()
    }

    RestoreOriginalFocus() {
        this.ClearStep()
        this.AlreadyActive := false
        if (this.OriginalHwnd)
            GeminiRestoreWindow(this.OriginalHwnd)
        if (this.GeminiHwnd && WinActive("ahk_id " this.GeminiHwnd))
            FocusGeminiAskFieldForHwnd(this.GeminiHwnd, false)
    }

    Fail() {
        this.ClearStep()
        try StandardLoadingBar_Hide(0)
        ShowNotification("Read aloud failed – Gemini UI not ready", 2000, "FF6666", "FFFFFF", 22)
        this.RestoreOriginalFocus()
    }
}

; =============================================================================
; GeminiAsyncLookup – async pronunciation lookup (Win+Alt+Shift+8)
; User keeps focus; timer polls for completion; result shown in banner.
; =============================================================================
class GeminiAsyncLookup {
    __New(lang := "") {
        this.Lang := lang
        this.OriginalHwnd := 0
        this.GeminiHwnd := 0
        this.TimerCallback := ""
        this.RetryCount := 0
        this.MaxRetries := 60   ; 60 * 500ms = 30s timeout
        this.ButtonEverFound := false
    }

    Start() {
        ; #region agent log (debug-d75ab1)
        __d75ab1_log := (hypothesisId, message, data := 0) => DebugLog_d75ab1(hypothesisId, "pre-fix",
            "Gemini.ahk:1472", message, data)
        __d75ab1_log("A", "Start() entry", Map(
            "activeHwnd", WinExist("A"),
            "activeTitle", SafeWinGetTitle_d75ab1("A"),
            "activeProcess", SafeWinGetProcessName_d75ab1("A")
        ))
        ; #endregion agent log (debug-d75ab1)

        this.OriginalHwnd := WinExist("A")
        if !this.OriginalHwnd {
            ; #region agent log (debug-d75ab1)
            __d75ab1_log("E", "Early return: no OriginalHwnd", 0)
            ; #endregion agent log (debug-d75ab1)
            return
        }
        ; Show loading banner immediately, centered on the monitor where this window is (with warning)
        StandardLoadingBar_Show("⏳ Loading…", BANNER_ACCENT_INTERMEDIATE)

        A_Clipboard := ""
        clipOk := TryCopySelectionToClipboard_d75ab1(__d75ab1_log)
        ; #region agent log (debug-d75ab1)
        __d75ab1_log("A", "ClipWait finished", Map(
            "clipOk", clipOk ? 1 : 0,
            "clipTextLen", StrLen(A_Clipboard),
            "clipFormats", ClipboardFormatsSummary_d75ab1()
        ))
        ; #endregion agent log (debug-d75ab1)
        if !clipOk {
            ; #region agent log (debug-d75ab1)
            __d75ab1_log("A", "Early return: ClipWait(2) failed (likely QuickLook copy)", 0)
            ; #endregion agent log (debug-d75ab1)
            try StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ Copy failed (no clipboard text). QuickLook selection may not support Ctrl+C.",
                2400,
                BANNER_ACCENT_ERROR)
            ; #region agent log (debug-d75ab1)
            __d75ab1_log("A", "Handled ClipWait failure: hid loading bar and showed error banner", 0)
            ; #endregion agent log (debug-d75ab1)
            return
        }
        SetTitleMatchMode(2)
        this.GeminiHwnd := GetGeminiWindowHwnd()
        if !this.GeminiHwnd {
            ; #region agent log (debug-d75ab1)
            __d75ab1_log("C", "Early return: GetGeminiWindowHwnd() returned 0", 0)
            ; #endregion agent log (debug-d75ab1)
            StandardLoadingBar_Hide(0)
            return
        }
        try {
            WinActivate("ahk_id " this.GeminiHwnd)
        } catch {
            ; #region agent log (debug-d75ab1)
            __d75ab1_log("C", "Early return: WinActivate Gemini threw", Map("geminiHwnd", this.GeminiHwnd))
            ; #endregion agent log (debug-d75ab1)
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        if !WinWaitActive("ahk_exe chrome.exe", , 2) {
            ; #region agent log (debug-d75ab1)
            __d75ab1_log("C", "Early return: WinWaitActive(chrome.exe,2s) failed after activation", Map("geminiHwnd",
                this.GeminiHwnd))
            ; #endregion agent log (debug-d75ab1)
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
        promptField := Gemini_FocusPromptSameAsOpenHotkey(uia)
        if (!promptField) {
            ; #region agent log (debug-d75ab1)
            __d75ab1_log("D", "Early return: Gemini_FocusPromptSameAsOpenHotkey() returned falsey", 0)
            ; #endregion agent log (debug-d75ab1)
            StandardLoadingBar_Hide(0)
            return
        }
        ; Switch to Fast model via mode picker menu (Gemini 3 MenuItem list), not @ text
        if (!EnsureGeminiModelViaMenu("Fast"))
            ShowCenteredOverlay_Utils("❌ Could not switch Gemini to Fast model.", 2200, BANNER_ACCENT_ERROR)
        try {
            promptField.SetFocus()
            Sleep 80
        } catch {
        }
        promptName := this.Lang ? "pronunciation-lookup-" . this.Lang : "pronunciation-lookup"
        searchString := RTrim(GetPromptText(promptName), "`r`n")
        A_Clipboard := searchString . "`n`nContent: " . A_Clipboard
        ; #region agent log (debug-d75ab1)
        __d75ab1_log("B", "Prepared prompt+content for paste", Map(
            "lang", this.Lang,
            "promptName", promptName,
            "promptLen", StrLen(searchString),
            "contentLen", StrLen(A_Clipboard) - (StrLen(searchString) + StrLen("`n`nContent: "))
        ))
        ; #endregion agent log (debug-d75ab1)
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
        GeminiBackgroundSetTimer(this, this.CheckCompletion.Bind(this), GEMINI_ASYNC_POLL_MS)
    }

    CheckCompletion() {
        state := GeminiMonitorStreamingTransition(this, this.OnStreamingCompleted.Bind(this))
        if (state = "timeout") {
            StandardLoadingBar_Hide(0)
            return
        }
    }

    OnStreamingCompleted() {
        ; Use the same sound as Shift keys.ahk for consistency
        try {
            ScriptSoundPlay(A_ScriptDir . "\sounds\gemini-completion.wav")
        } catch {
            PlayCopyCompletedChime()
        }
        this.RetrieveResponse()
    }

    RetrieveResponse() {
        ; Activate Gemini once, then copy with retry (exponential backoff) without switching back until done.
        ; Invalidate last-Copy-button cache so we discover the newly completed response (avoid penultimate message).
        GeminiState.Invalidate()
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
        if (this.OriginalHwnd = this.GeminiHwnd)
            FocusGeminiAskFieldForHwnd(this.GeminiHwnd, false)
        ; Original may have been closed while we waited on Gemini — never WinActivate a stale HWND.
        origHwnd := this.OriginalHwnd
        if (origHwnd && WinExist("ahk_id " origHwnd)) {
            try {
                WinActivate("ahk_id " origHwnd)
                if (!WinActive("ahk_id " origHwnd))
                    WinWaitActive("ahk_id " origHwnd, , 0.5)
            } catch {
                if (WinExist("ahk_id " origHwnd))
                    try WinActivate("ahk_id " origHwnd)
            }
        }
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
        ; timeoutMs 0 = no auto-dismiss; user closes with Enter, E, or Escape (Utils.ahk StandardLoadingBar_ShowWithKeys).
        StandardLoadingBar_ShowWithKeys(state, closeKeys, 0, 0, "",
            BANNER_ACCENT_INTERMEDIATE, 600,
            17, "", false,
            "[Enter] [E] [Esc] Close",
            true)
    }
}

; #region agent log (debug-d75ab1)
DebugLog_d75ab1(hypothesisId, runId, location, message, data := 0) {
    static __logPath := A_ScriptDir "\debug-d75ab1.log"
    payload := Map(
        "sessionId", "d75ab1",
        "runId", runId,
        "hypothesisId", hypothesisId,
        "location", location,
        "message", message,
        "timestamp", A_NowUTC
    )
    if (IsObject(data))
        payload["data"] := data
    else if (data != 0)
        payload["data"] := data
    try FileAppend(JsonStringify_d75ab1(payload) "`n", __logPath, "UTF-8")
}

TryCopySelectionToClipboard_d75ab1(loggerFn) {
    ; Attempt 1: Ctrl+C
    proc := ""
    try {
        proc := WinGetProcessName("A")
    } catch {
        proc := ""
    }
    ; #region agent log (debug-d75ab1)
    loggerFn("A", "Copy attempt 1: Ctrl+C", Map("activeProcess", proc))
    ; #endregion agent log (debug-d75ab1)
    A_Clipboard := ""
    Send "^c"
    if ClipWait(0.7)
        return true

    ; Attempt 2: Ctrl+Insert (common alternative)
    ; #region agent log (debug-d75ab1)
    loggerFn("A", "Copy attempt 2: Ctrl+Insert", 0)
    ; #endregion agent log (debug-d75ab1)
    A_Clipboard := ""
    Send "^{Insert}"
    if ClipWait(0.7)
        return true

    ; Attempt 3: QuickLook context menu copy
    if (proc = "QuickLook.exe") {
        ; #region agent log (debug-d75ab1)
        loggerFn("A", "Copy attempt 3: QuickLook context menu (AppsKey then C)", 0)
        ; #endregion agent log (debug-d75ab1)
        A_Clipboard := ""
        Send "{AppsKey}"
        Sleep 60
        Send "c"
        if ClipWait(0.9)
            return true

        ; Attempt 4: UIA selection extraction (no clipboard)
        ; #region agent log (debug-d75ab1)
        loggerFn("F", "Copy attempt 4: UIA selected-text extraction (QuickLook)", 0)
        ; #endregion agent log (debug-d75ab1)
        try {
            txt := TryGetSelectedTextViaUIA_d75ab1(loggerFn)
            if (txt != "" && StrLen(Trim(txt)) > 0) {
                A_Clipboard := txt
                ; #region agent log (debug-d75ab1)
                loggerFn("F", "UIA extracted text; placing into clipboard variable", Map("len", StrLen(txt)))
                ; #endregion agent log (debug-d75ab1)
                return true
            }
        } catch {
        }
    }

    return false
}

TryGetSelectedTextViaUIA_d75ab1(loggerFn) {
    hwnd := WinExist("A")
    ; Focused element is often the text host; try it first.
    focused := 0
    try {
        focused := UIA.GetFocusedElement()
    } catch {
        focused := 0
    }

    if (focused) {
        ; #region agent log (debug-d75ab1)
        loggerFn("F", "UIA focused element snapshot", Map(
            "name", focused.Name,
            "type", focused.LocalizedControlType,
            "hasFocus", focused.HasKeyboardFocus,
            "isText", focused.IsTextPatternAvailable ? 1 : 0
        ))
        ; #endregion agent log (debug-d75ab1)
        if (focused.IsTextPatternAvailable) {
            try {
                ranges := focused.TextPattern.GetSelection()
                if (IsObject(ranges) && ranges.Length >= 1) {
                    txt := ranges[1].GetText(512)
                    return Trim(txt)
                }
            } catch {
            }
        }
    }

    ; Fallback: find a Document element under the window and read its selection/document range.
    try {
        root := UIA.ElementFromHandle(hwnd)
        doc := root.FindFirst({ Type: UIA.ControlType.Document })
        if (doc) {
            ; #region agent log (debug-d75ab1)
            loggerFn("F", "UIA document element found", Map(
                "name", doc.Name,
                "isText", doc.IsTextPatternAvailable ? 1 : 0
            ))
            ; #endregion agent log (debug-d75ab1)
            if (doc.IsTextPatternAvailable) {
                try {
                    ranges := doc.TextPattern.GetSelection()
                    if (IsObject(ranges) && ranges.Length >= 1) {
                        txt := ranges[1].GetText(512)
                        if (Trim(txt) != "")
                            return Trim(txt)
                    }
                } catch {
                }
                try {
                    txt := doc.TextPattern.DocumentRange.GetText(512)
                    return Trim(txt)
                } catch {
                }
            }
        }
    } catch {
    }
    return ""
}

SafeWinGetTitle_d75ab1(win) {
    try {
        return WinGetTitle(win)
    } catch {
        return ""
    }
}

SafeWinGetProcessName_d75ab1(win) {
    try {
        return WinGetProcessName(win)
    } catch {
        return ""
    }
}

ClipboardFormatsSummary_d75ab1() {
    out := ""
    try {
        if !DllCall("OpenClipboard", "ptr", 0, "int")
            return ""
        fmt := 0
        loop 32 {
            fmt := DllCall("EnumClipboardFormats", "uint", fmt, "uint")
            if (!fmt)
                break
            out .= (out = "" ? "" : ",") fmt
        }
    } catch {
        out := out ? out : ""
    }
    try DllCall("CloseClipboard")
    return out
}

JsonStringify_d75ab1(val) {
    q := Chr(34)
    if (!IsObject(val)) {
        if (val is Number)
            return val + 0
        if (val = true)
            return "true"
        if (val = false)
            return "false"
        if (val = "")
            return q q
        return q JsonEscape_d75ab1(val) q
    }
    if (val is Map) {
        s := "{"
        first := true
        for k, v in val {
            if (!first)
                s .= ","
            first := false
            s .= q JsonEscape_d75ab1(k) q ":" JsonStringify_d75ab1(v)
        }
        return s . "}"
    }
    ; Fallback: stringify unknown objects as their Type name
    try {
        return q JsonEscape_d75ab1(Type(val)) q
    } catch {
        return q q
    }
}

JsonEscape_d75ab1(s) {
    s := StrReplace(s, "\", "\\")
    s := StrReplace(s, '"', '\"')
    s := StrReplace(s, "`r", "\r")
    s := StrReplace(s, "`n", "\n")
    s := StrReplace(s, "`t", "\t")
    return s
}
; #endregion agent log (debug-d75ab1)

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
        this.HasCopiedForThisResponse := false
        this.CopyBannerShownForThisResponse := false
    }

    Start(originalHwnd, geminiHwnd) {
        if (!originalHwnd || !geminiHwnd)
            return
        this.OriginalHwnd := originalHwnd
        this.GeminiHwnd := geminiHwnd
        this.RetryCount := 0
        this.ButtonEverFound := false
        this.HasCopiedForThisResponse := false
        this.CopyBannerShownForThisResponse := false
        GeminiBackgroundSetTimer(this, this.CheckCompletion.Bind(this), GEMINI_ASYNC_POLL_MS)
    }

    ; Stop polling; used when user chose S or N at 6s so "Copy response?" is never shown for this flow.
    Stop() {
        GeminiBackgroundStopTimer(this)
    }

    CheckCompletion() {
        state := GeminiMonitorStreamingTransition(this, this.OnStreamingCompleted.Bind(this))
        if (state = "timeout")
            return
    }

    OnStreamingCompleted() {
        try {
            ScriptSoundPlay(A_ScriptDir . "\sounds\gemini-completion.wav")
        } catch {
            PlayCopyCompletedChime()
        }
        this.ShowCopyDecisionBanner()
    }

    ShowCopyDecisionBanner() {
        if (this.CopyBannerShownForThisResponse)
            return
        this.CopyBannerShownForThisResponse := true
        this.CopyBannerGui := ""
        this.CopyTimeoutTimer := ""
        copyKeyCallbacks := Map("N", this.CancelCopy.Bind(this), "Y", this.DoCopyOnly.Bind(this), "R", this.CopyAndReadAloud
        .Bind(this), "C", this.CopyAndTransferToCursor.Bind(this), "F", this.CopyAndFavorite.Bind(this))
        StandardLoadingBar_ShowWithKeys("❓ Copy response?", copyKeyCallbacks, 5000, 0, this.DoCopyOnTimeout
            .Bind(this), BANNER_ACCENT_INTERMEDIATE, 520, 17, "", false,
            "[Y] Copy  [N] No  [R] Copy+Read  [C] Transfer  [F] Copy+Favorite",
            true)
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

    ; N key: close overlay and stop (no copy/read/transfer). Explicitly close Utils overlay so cancel always takes effect.
    CancelCopy(*) {
        try StandardLoadingBar_CloseKeysOverlay()
        catch {
        }
        try StandardLoadingBar_Hide(0)
        catch {
        }
        this.CleanupCopyBanner()
    }

    DoCopyCore(readAloud := false, skipRestoreFocus := false) {
        if (this.HasCopiedForThisResponse)
            return
        this.HasCopiedForThisResponse := true
        GeminiState.Invalidate()
        if !WinExist("ahk_id " this.GeminiHwnd) {
            if (WinExist("ahk_id " this.OriginalHwnd))
                WinActivate("ahk_id " this.OriginalHwnd)
            return
        }
        ; Hands off cue before activating Gemini to copy the last response (manual Y/R/C and timeout).
        PlayPreMovementWarning("Gemini")
        ; If Gemini is not active when the monitor fires, activate it now.
        if !WinActive("ahk_id " this.GeminiHwnd) {
            try {
                WinActivate("ahk_id " this.GeminiHwnd)
            } catch {
                if (WinExist("ahk_id " this.OriginalHwnd))
                    WinActivate("ahk_id " this.OriginalHwnd)
                return
            }
            if !WinWaitActive("ahk_exe chrome.exe", , 0.5) {
                if (WinExist("ahk_id " this.OriginalHwnd))
                    WinActivate("ahk_id " this.OriginalHwnd)
                return
            }
        }
        copyOpt := { restoreWindow: false, playChimeAndNotify: false, alreadyActive: true }
        if CopyLastGeminiMessageWithRetry(copyOpt, this.GeminiHwnd) {
            PlayCopyCompletedChime()
        }
        if (readAloud)
            GeminiTriggerReadAloud(false, false, { originalHwnd: this.OriginalHwnd, geminiHwnd: this.GeminiHwnd,
                alreadyActive: true })
        if (WinActive("ahk_id " this.GeminiHwnd))
            FocusGeminiAskFieldForHwnd(this.GeminiHwnd, false)
        ; Gemini/Clipboard → Original: return transitions are immediate (no warning).
        if (!skipRestoreFocus && WinExist("ahk_id " this.OriginalHwnd) && !WinActive("ahk_id " this.OriginalHwnd)) {
            WinActivate("ahk_id " this.OriginalHwnd)
            ; Fast-path: avoid WinWaitActive if we are already active.
            if (!WinActive("ahk_id " this.OriginalHwnd))
                WinWaitActive("ahk_id " this.OriginalHwnd, , 0.5)
        }
    }

    DoCopyOnTimeout(*) {
        this.CleanupCopyBanner()
        this.DoCopyCore(false)
    }

    ; R key: copy last message and read it aloud, then restore focus (same tab as delayed submit).
    CopyAndReadAloud(*) {
        this.CleanupCopyBanner()
        this.DoCopyCore(true)
    }

    ; F key: copy last response, then mark the new top clip as favorite in Clip Angel (same as MarkLastClipAsFavorite).
    CopyAndFavorite(*) {
        this.CleanupCopyBanner()
        this.DoCopyCore(false)
        clipRaw := A_Clipboard
        clip := Trim(clipRaw)
        if (clip = "" || StrLen(clip) < GEMINI_TRANSFER_MIN_CLIPBOARD_LENGTH) {
            ShowCenteredOverlay_Utils("❌ Copy failed or empty – try again", 2000, BANNER_ACCENT_ERROR)
            if (WinExist("ahk_id " this.OriginalHwnd))
                WinActivate("ahk_id " this.OriginalHwnd)
            return
        }
        Sleep(GEMINI_POST_COPY_FAVORITE_DELAY_MS)
        MarkLastClipAsFavorite()
    }

    ; C key: copy response, then show Cursor window selector (1–9), activate selected window, focus AI field, paste and send.
    CopyAndTransferToCursor(*) {
        this.CleanupCopyBanner()
        ; Skip restoring focus so clipboard is not overwritten by the previously focused window before we read it.
        this.DoCopyCore(false, true)
        clipRaw := A_Clipboard
        clip := Trim(clipRaw)
        if (clip = "" || StrLen(clip) < GEMINI_TRANSFER_MIN_CLIPBOARD_LENGTH) {
            ShowCenteredOverlay_Utils("❌ Copy failed or empty – try again", 2000, BANNER_ACCENT_ERROR)
            if (WinExist("ahk_id " this.OriginalHwnd))
                WinActivate("ahk_id " this.OriginalHwnd)
            return
        }

        ; Restore the pre-handoff anchored window so the user sees the selector/paste in the exact context.
        if (this.OriginalHwnd && WinExist("ahk_id " this.OriginalHwnd)) {
            try {
                WinActivate("ahk_id " this.OriginalHwnd)
                ; Fast-path: avoid WinWaitActive if we are already active.
                if (!WinActive("ahk_id " this.OriginalHwnd))
                    WinWaitActive("ahk_id " this.OriginalHwnd, , 0.5)
            } catch {
            }
        }
        try A_Clipboard := clipRaw

        hwnd := CursorTransfer_ShowWindowSelector(0)
        if (!hwnd) {
            if (WinExist("ahk_id " this.OriginalHwnd))
                WinActivate("ahk_id " this.OriginalHwnd)
            try A_Clipboard := clipRaw
            return
        }
        ; Gemini → Cursor: no pre-movement warning (source is not Original).
        try A_Clipboard := clipRaw
        CursorTransfer_ActivateFocusPaste(hwnd, this.OriginalHwnd)
    }
}

; Current monitor instance so we can stop it when user chooses S or N at 6s (no copy/transfer follow-up).
global g_GeminiDelayedSubmitMonitor := ""

; Callable from Utils.ahk after successful auto-send (Ctrl+Alt+Win+L).
GeminiDelayedSubmitMonitorStart(originalHwnd, geminiHwnd) {
    global g_GeminiDelayedSubmitMonitor
    g_GeminiDelayedSubmitMonitor := GeminiDelayedSubmitMonitor()
    g_GeminiDelayedSubmitMonitor.Start(originalHwnd, geminiHwnd)
}

; Stop any running monitor so "Copy response?" does not show (e.g. after S or N at 6s dictation confirm).
GeminiDelayedSubmitMonitorStop() {
    global g_GeminiDelayedSubmitMonitor
    if (g_GeminiDelayedSubmitMonitor)
        try g_GeminiDelayedSubmitMonitor.Stop()
    g_GeminiDelayedSubmitMonitor := ""
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
        promptField := Gemini_FocusPromptSameAsOpenHotkey(uia)
        if (!promptField) {
            StandardLoadingBar_Hide(0)
            return
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
        GeminiBackgroundSetTimer(this, this.CheckCompletion.Bind(this), GEMINI_ASYNC_POLL_MS)
    }

    CheckCompletion() {
        state := GeminiMonitorStreamingTransition(this, this.OnStreamingCompleted.Bind(this))
        if (state = "timeout") {
            StandardLoadingBar_Hide(0)
            return
        }
    }

    OnStreamingCompleted() {
        StandardLoadingBar_Hide(0)
        ; Completion detection matches GeminiAsyncLookup (#!+8): Layer 1 only (Stop button gone). No extra Layer 2 so we don't miss completion.
        try {
            ScriptSoundPlay(A_ScriptDir . "\sounds\gemini-completion.wav")
        } catch {
            PlayCopyCompletedChime()
        }
        ; Allow DOM to finish rendering, then hand off read aloud without keeping focus on Gemini.
        Sleep(GeminiAsyncTTS.PostStreamingDelayMs)
        GeminiTriggerReadAloud(false, true, { originalHwnd: this.OriginalHwnd, geminiHwnd: this.GeminiHwnd })
    }
}

; Optional Python sidecar contract. AHK remains the default orchestration layer; if enabled later,
; the helper should accept a lightweight task envelope such as:
; { kind, geminiHwnd, originalHwnd, copyFirst, useTrashTab, requestedAt }
GeminiQueueBackgroundTask(taskKind, payload := "") {
    if (!GEMINI_USE_PYTHON_IPC)
        return false
    if (!GeminiIPC_EnsureReady(400))
        return false
    reqPayload := payload ? payload : Map()
    if (!(reqPayload is Map))
        reqPayload := Map()
    reqPayload["requestedAt"] := A_TickCount
    resp := GeminiIPC_QueueTask(taskKind, reqPayload)
    if (!GeminiIPC_ResponseOk(resp))
        return false
    return GeminiIPC_ResponseResultMap(resp)
}
