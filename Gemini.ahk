#Requires AutoHotkey v2.0
#SingleInstance Force

; --- Includes ----------------------------------------------------------------
#include UIA-v2\Lib\UIA.ahk
#include UIA-v2\Lib\UIA_Browser.ahk
#include %A_ScriptDir%\env.ahk
#include %A_ScriptDir%\Utils.ahk
try Hotkey("#!+Y", "Off")
try Hotkey("#!+X", "Off")
#include %A_ScriptDir%\aux\WMIPC.ahk

#include %A_ScriptDir%\aux\GeminiIPC.ahk

; -----------------------------------------------------------------------------
; MODULE MAP - Gemini.ahk stays the runnable entry point and #includes each
; module below. Early preamble includes (UIA, env, Utils, WMIPC, GeminiIPC) stay
; here. See Gemini/MODULARIZATION_PROGRESS.md for the full module list.
; -----------------------------------------------------------------------------

; --- Config ---------------------------------------------------------------
; Copy response button names (EN/PT). Excludes "Copy prompt" / "Copiar prompt" which are different controls.
GEMINI_COPY_RESPONSE_NAMES := ["Copy", "Copiar"]

; TTS Pause/Resume button names (EN/PT). Used by FindGeminiPauseResumeButton during read-aloud verification.
GEMINI_TTS_PAUSE_NAMES := ["Pause", "Pausar"]
GEMINI_TTS_RESUME_NAMES := ["Resume", "Retomar"]

; --- Refactor: threshold constants (no magic numbers) ---------------------------------
; Timeouts (ms)
GEMINI_ACTIVATE_WAIT_MS := 2000
GEMINI_SCROLL_SETTLE_MS := 350
GEMINI_UIA_SETTLE_MS := 120
GEMINI_PROMPT_FOCUS_POLL_MS := 25
GEMINI_PROMPT_FOCUS_TIMEOUT_MS := 300
GEMINI_OPEN_FAST_SETTLE_MS := 0
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
GEMINI_FIRST_LAUNCH_MODEL_READY_TIMEOUT_MS := 12000
GEMINI_FIRST_LAUNCH_MODEL_READY_POLL_MS := 250
GEMINI_FIRST_LAUNCH_TAB_SWITCH_SETTLE_MS := 350
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
; Minimum ~3 effective polls after immediate CheckStarted; extra margin for cold TTS while Gemini stays foreground.
GEMINI_READ_ALOUD_START_MAX_RETRIES := 8
; Total Listen-menu attempts (Gemini often needs >1) before full RetryLaunch: 1 initial + (MAX-1) replays.
GEMINI_READ_ALOUD_LISTEN_PHASE_MAX := 3
; Dictation "Copy + Read" path: response just finished streaming, so TTS engine is cold.
; Use the pre-optimization budget (10 * 150ms = 1500ms) so the Pause button has time to render
; without triggering RetryLaunch (which would emit a second "Hands off!" cue).
GEMINI_DICTATION_READ_ALOUD_MAX_RETRIES := 10
GEMINI_COPY_MAX_RETRIES := 3
; Minimum clipboard length for Gemini-to-Cursor transfer (same as bridge validation)
GEMINI_TRANSFER_MIN_CLIPBOARD_LENGTH := 10
; Post-copy sync: max wait for clipboard change (sequence-number detection). Fallback if detection fails (ms).
GEMINI_POST_COPY_SYNC_TIMEOUT_MS := 2000
; Poll interval when waiting for clipboard sequence number to change (ms). Low value = minimal latency.
GEMINI_CLIPBOARD_POLL_MS := 10
; Performance instrumentation (set to true to log latencies to script dir)
; Logs focus phase (fast_already_focused, direct_focus, anchor_fallback) and tab_banner_deferred.
GEMINI_PERF_LOG_ENABLED := false
GEMINI_PERF_LOG_PATH := A_ScriptDir "\.cursor\gemini_perf.log"

; Feature flags for refactor phases (set false to fall back to legacy behavior)
GEMINI_USE_WIN_EVENT_HOOK := true
GEMINI_USE_PYTHON_IPC := true

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

; Find Pause or Resume button (TTS). which = "Pause" or "Resume". Returns element or 0.
; Resolves localized names (EN/PT) via GEMINI_TTS_PAUSE_NAMES / GEMINI_TTS_RESUME_NAMES and runs
; cached FindFirstBuildCache per name. No FindAll fallback (canon §3 / §4): one cheap COM call per name.
FindGeminiPauseResumeButton(uia, which) {
    static cacheRequest := ""
    if (!cacheRequest)
        cacheRequest := UIA.CreateCacheRequest(["Name", "ClassName"], , 5)
    names := (which = "Resume") ? GEMINI_TTS_RESUME_NAMES : GEMINI_TTS_PAUSE_NAMES
    for n in names {
        try {
            btn := uia.FindFirstBuildCache(cacheRequest, { Name: n, Type: UIA_ControlType_Button })
            if (btn)
                return btn
        } catch {
            continue
        }
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
    global WM_USE_DAEMON := false, WM_USE_PIPE_IPC := false, WM_USE_EVENT_HOOK_CACHE := false
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
        if UseCopilotWebForGlobalAI()
            return CopilotWeb_TriggerReadAloud()
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
        if UseCopilotWebForGlobalAI() {
            if (!CopilotWeb_CopyLastMessageToClipboard({ restoreWindow: false, playChimeAndNotify: true }))
                ShowNotification("Copy failed – ensure Copilot is open and has a response", 2500, "FF6666", "FFFFFF",
                    22)
            else if (hwnd := GetCopilotWebWindowHwnd())
                CopilotWeb_FocusComposerForHwnd(hwnd, true)
            return
        }
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
; wParam: 1 = caller already copied (skip Copy in Gemini). lParam: anchored original hwnd for focus restore (0 = resolve like #!+o).
WM_TRIGGER_READ_ALOUD := 0x8004
; Work environment: M365 Copilot web (Chrome) copy / read-aloud IPC from Utils D2C_FlowManager.
WM_COPY_LAST_COPILOT := 0x8005
WM_TRIGGER_COPILOT_READ_ALOUD := 0x8006
; Path for bridge to verify that Copy Last Response (same as #!+p) actually succeeded
GEMINI_COPY_RESULT_PATH := A_ScriptDir "\.cursor\gemini_copy_result.txt"

OnMessage(WM_COPY_LAST_GEMINI, copyFromBridge)
OnMessage(WM_COPY_LAST_COPILOT, copyCopilotFromBridge)
OnMessage(WM_START_DELAYED_SUBMIT_MONITOR, handleStartDelayedSubmitMonitor)
OnMessage(WM_STOP_DELAYED_SUBMIT_MONITOR, handleStopDelayedSubmitMonitor)
OnMessage(WM_TRIGGER_READ_ALOUD, handleTriggerReadAloud)
OnMessage(WM_TRIGGER_COPILOT_READ_ALOUD, handleTriggerCopilotReadAloud)
handleStartDelayedSubmitMonitor(wParam, lParam, msg, hwnd) {
    GeminiDelayedSubmitMonitorStart(wParam, lParam)
}
handleStopDelayedSubmitMonitor(*) {
    GeminiDelayedSubmitMonitorStop()
}
handleTriggerReadAloud(wParam, lParam, msg, hwnd) {
    ; wParam 1: D2C already ran WM_COPY_LAST_GEMINI; skip internal Copy click, open Listen only.
    ; lParam: anchored original hwnd from dictation D2C (0 = resolve like local #!+o).
    wp := Integer(wParam)
    lp := Integer(lParam)
    copyFirst := !(wp = 1)
    gemHwnd := GetGeminiWindowHwnd()
    if (lp && WinExist("ahk_id " lp))
        return GeminiTriggerReadAloud(copyFirst, false, { originalHwnd: lp, geminiHwnd: gemHwnd ? gemHwnd : 0,
            alreadyActive: true, verifyMaxRetries: GEMINI_DICTATION_READ_ALOUD_MAX_RETRIES })
    return GeminiTriggerReadAloud(copyFirst)
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

copyCopilotFromBridge(wParam, lParam, msg, hwnd) {
    copilotHwnd := Integer(lParam)
    try
        FileDelete(GEMINI_COPY_RESULT_PATH)
    try
        FileAppend("0", GEMINI_COPY_RESULT_PATH)
    opts := { restoreWindow: false, playChimeAndNotify: false }
    if (copilotHwnd && WinActive("ahk_id " copilotHwnd))
        opts.alreadyActive := true
    r := CopilotWeb_CopyLastMessageToClipboard(opts, copilotHwnd)
    try
        FileDelete(GEMINI_COPY_RESULT_PATH)
    try
        FileAppend(r ? "1" : "0", GEMINI_COPY_RESULT_PATH)
}

handleTriggerCopilotReadAloud(wParam, lParam, msg, hwnd) {
    wp := Integer(wParam)
    lp := Integer(lParam)
    copyFirst := !(wp = 1)
    copHwnd := GetCopilotWebWindowHwnd()
    if (lp && WinExist("ahk_id " lp))
        return CopilotWeb_TriggerReadAloud(copyFirst, { originalHwnd: lp, copilotHwnd: copHwnd ? copHwnd : 0,
            alreadyActive: true })
    return CopilotWeb_TriggerReadAloud(copyFirst)
}

; =============================================================================
; TTS from selection – Win+Alt+Shift+7: copy selection, send "repeat exactly" to Gemini, then trigger read aloud
; =============================================================================
#!+7:: {
    if UseCopilotWebForGlobalAI()
        return (CopilotAsyncTTS()).Start()
    (GeminiAsyncTTS()).Start()
}

; [Gemini module] Language picker and #!+8 pronunciation hotkey -> Gemini\hotkey_pronunciation.ahk
#include %A_ScriptDir%\Gemini\hotkey_pronunciation.ahk

; [Gemini module] InitializeGeminiFirstTime and #!+I open/focus hotkey -> Gemini\gemini_open.ahk
#include %A_ScriptDir%\Gemini\gemini_open.ahk

; Gemini_FocusPromptSameAsOpenHotkey lives in Utils.ahk (shared with Shift keys Fast Copy).

; [Gemini module] GeminiAsyncReadAloud async read-aloud class -> Gemini\gemini_async_readaloud.ahk
#include %A_ScriptDir%\Gemini\gemini_async_readaloud.ahk

; [Gemini module] GeminiAsyncLookup pronunciation async class -> Gemini\gemini_async_lookup.ahk
#include %A_ScriptDir%\Gemini\gemini_async_lookup.ahk

; [Gemini module] GeminiDelayedSubmitMonitor and start/stop helpers -> Gemini\gemini_delayed_submit.ahk
#include %A_ScriptDir%\Gemini\gemini_delayed_submit.ahk

; [Gemini module] GeminiAsyncTTS class and GeminiQueueBackgroundTask -> Gemini\gemini_async_tts.ahk
#include %A_ScriptDir%\Gemini\gemini_async_tts.ahk
