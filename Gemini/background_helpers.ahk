; =============================================================================
; Gemini module: background_helpers.ahk
; ShowNotification, background timers, copy chime
; Extracted verbatim from Gemini.ahk; loaded via #include into the
; Gemini.ahk process, which remains the entry point / source of truth.
; =============================================================================

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

        ScriptSoundPlay(A_ScriptDir . "\assets\sounds\copy.wav")
    } catch {
        ; Silently ignore errors
    }
}
