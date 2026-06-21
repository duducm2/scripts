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

; [Gemini module] GeminiState, copy-button UIA, tab banner, model picker -> Gemini\gemini_uia_core.ahk
#include %A_ScriptDir%\Gemini\gemini_uia_core.ahk

; [Gemini module] ShowNotification, background timers, copy chime -> Gemini\background_helpers.ahk
#include %A_ScriptDir%\Gemini\background_helpers.ahk

; [Gemini module] Small loading indicator and WaitForButton helpers -> Gemini\loading_wait.ahk
#include %A_ScriptDir%\Gemini\loading_wait.ahk

; [Gemini module] #!+O/#!+P/#!+7 and CopyLastGeminiMessageToClipboard -> Gemini\hotkey_read_copy.ahk
#include %A_ScriptDir%\Gemini\hotkey_read_copy.ahk

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
