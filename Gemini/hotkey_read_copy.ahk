; =============================================================================
; Gemini module: hotkey_read_copy.ahk
; #!+P (1× message / 2× code): destination banner first, then companion copy + action after choice.
; Also: CopyLastGeminiMessageToClipboard / CopyLastGeminiCodeSnippetToClipboard, read-aloud IPC.
; (Win+Alt+Shift+O lives in Utils: DesktopCutNewest_OnHotkey — cut / open / paste clipboard / copy path)
; Extracted verbatim from Gemini.ahk; loaded via #include into the
; Gemini.ahk process, which remains the entry point / source of truth.
; =============================================================================

; --- Hotkeys ----------------------------------------------------------------

; Reusable launcher: activate Gemini asynchronously, handle Pause/Resume, then optionally
; copy the last message and trigger read aloud (D2C R / IPC / TTS / delayed submit).
GeminiTriggerReadAloud(copyFirst := true, useTrashTab := false, options := "") {
    return (GeminiAsyncReadAloud(copyFirst, useTrashTab, options)).Start()
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

; Copy last Gemini code-fence snippet to clipboard (#!+p double-tap). Same options as CopyLastGeminiMessageToClipboard.
CopyLastGeminiCodeSnippetToClipboard(options := "", geminiHwnd := 0) {
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
            GeminiPerfLog("copy_code", t0)
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

        Send "^{End}"
        Sleep GEMINI_SCROLL_SETTLE_MS

        lastCodeButton := GetLastGeminiCopyCodeButton(uia)
        if (!lastCodeButton) {
            GeminiPerfLog("copy_code", t0)
            return false
        }
        A_Clipboard := ""
        lastCodeButton.Click()
        if !ClipWait(2) {
            GeminiPerfLog("copy_code", t0)
            return false
        }
        if (playChimeAndNotify) {
            PlayCopyCompletedChime()
            ShowNotification("Code copied!", 800, "FFFF00", "000000", 24)
        }
        if (restoreWindow)
            Send "!{Tab}"
        else
            FocusGeminiAskFieldForHwnd(geminiHwnd, false)
        GeminiPerfLog("copy_code", t0)
        return true
    } catch {
        GeminiPerfLog("copy_code", t0)
        return false
    }
}

; #!+p orchestrator: show destination banner first; copy starts only after user picks Y/F/C/R/W/O.
HotkeyCopy_RunIntentFlow(isCode := false) {
    originHwnd := 0
    try originHwnd := WinGetID("A")
    companion := ""
    try companion := ResolveGlobalAICompanion()
    catch {
        companion := ""
    }
    HotkeyCopy_ShowIntentBanner(isCode, originHwnd, companion)
}

; Companion-aware copy last message (#!+p single-tap worker). No banner; reports via OnCopyWorkerDone.
HotkeyCopy_RunCopyLastMessage(gen := 0) {
    ok := false
    err := "Copy failed – ensure Gemini is open and has a response"
    try {
        t0 := A_TickCount
        companion := ResolveGlobalAICompanion()
        opts := { restoreWindow: false, playChimeAndNotify: false }
        if (companion = "enterprise") {
            err := "Copy failed – ensure Gemini Enterprise is open and has a response"
            ok := GeminiEnterprise_CopyLastMessageToClipboard(opts)
            if (ok) {
                if (hwnd := GetGeminiEnterpriseWindowHwnd()) {
                    root := GeminiEnterprise_ReadRootFromHwnd(hwnd)
                    if (IsObject(root))
                        GeminiEnterprise_FocusComposer(root, true)
                }
            }
        } else if (companion = "copilot") {
            err := "Copy failed – ensure Copilot is open and has a response"
            ok := CopilotWeb_CopyLastMessageToClipboard(opts)
            if (ok) {
                if (hwnd := GetCopilotWebWindowHwnd())
                    CopilotWeb_FocusComposerForHwnd(hwnd, true)
            }
        } else {
            ok := CopyLastGeminiMessageToClipboard(opts)
            if (ok) {
                if (hwnd := GetGeminiWindowHwnd())
                    FocusGeminiAskFieldForHwnd(hwnd, true)
            }
        }
        GeminiPerfLog("hotkey_copy", t0)
    } catch as e {
        ok := false
        err := "Copy error: " (e.Message ? e.Message : "unknown")
    }
    HotkeyCopy_OnCopyWorkerDone(ok, ok ? "" : err, gen)
}

; Companion-aware copy last code snippet (#!+p double-tap worker). No banner; reports via OnCopyWorkerDone.
HotkeyCopy_RunCopyLastCode(gen := 0) {
    ok := false
    err := "No code snippet found – ensure Gemini has a code block"
    try {
        t0 := A_TickCount
        companion := ResolveGlobalAICompanion()
        opts := { restoreWindow: false, playChimeAndNotify: false }
        if (companion = "enterprise") {
            err := "No code snippet found – ensure Gemini Enterprise has a code block"
            ok := GeminiEnterprise_CopyLastCodeSnippetToClipboard(opts)
            if (ok) {
                if (hwnd := GetGeminiEnterpriseWindowHwnd()) {
                    root := GeminiEnterprise_ReadRootFromHwnd(hwnd)
                    if (IsObject(root))
                        GeminiEnterprise_FocusComposer(root, true)
                }
            }
        } else if (companion = "copilot") {
            err := "No code snippet found – ensure Copilot has a code block"
            ok := CopilotWeb_CopyLastCodeSnippetToClipboard(opts)
            if (ok) {
                if (hwnd := GetCopilotWebWindowHwnd())
                    CopilotWeb_FocusComposerForHwnd(hwnd, true)
            }
        } else {
            ok := CopyLastGeminiCodeSnippetToClipboard(opts)
            if (ok) {
                if (hwnd := GetGeminiWindowHwnd())
                    FocusGeminiAskFieldForHwnd(hwnd, true)
            }
        }
        GeminiPerfLog("hotkey_copy_code", t0)
    } catch as e {
        ok := false
        err := "Copy code error: " (e.Message ? e.Message : "unknown")
    }
    HotkeyCopy_OnCopyWorkerDone(ok, ok ? "" : err, gen)
}

; Win+Alt+Shift+P tap-dance (400 ms = AI_QD_DOUBLE_TAP_MS):
;   1× = destination banner, then copy last message/response after choice
;   2× = destination banner, then copy most recent code snippet after choice
; Copy starts only after Y/F/C/R/W/O; N/Esc/timeout dismisses without copying.
global g_HotkeyCopy_DoubleTapArmed := false
global g_HotkeyCopy_LastPressTick := 0
global g_HotkeyCopy_DoubleTapTimer := 0

class HotkeyCopy_DoubleTapTimerObj {
    static OnSingleTapTimeout() {
        global g_HotkeyCopy_DoubleTapArmed, g_HotkeyCopy_DoubleTapTimer
        if (!g_HotkeyCopy_DoubleTapArmed)
            return
        g_HotkeyCopy_DoubleTapArmed := false
        g_HotkeyCopy_DoubleTapTimer := 0
        SetTimer((*) => HotkeyCopy_RunIntentFlow(false), -1)
    }
}

#!+p:: {
    global g_HotkeyCopy_DoubleTapArmed, g_HotkeyCopy_LastPressTick, g_HotkeyCopy_DoubleTapTimer

    thresholdMs := 400
    try thresholdMs := AI_QD_DOUBLE_TAP_MS
    catch {
        thresholdMs := 400
    }

    now := A_TickCount
    elapsed := (g_HotkeyCopy_LastPressTick > 0) ? (now - g_HotkeyCopy_LastPressTick) : 9999

    if (g_HotkeyCopy_DoubleTapArmed && elapsed >= 0 && elapsed < thresholdMs) {
        g_HotkeyCopy_DoubleTapArmed := false
        g_HotkeyCopy_LastPressTick := 0
        if (g_HotkeyCopy_DoubleTapTimer) {
            SetTimer(g_HotkeyCopy_DoubleTapTimer, 0)
            g_HotkeyCopy_DoubleTapTimer := 0
        }
        SetTimer((*) => HotkeyCopy_RunIntentFlow(true), -1)
        return
    }

    g_HotkeyCopy_LastPressTick := now
    g_HotkeyCopy_DoubleTapArmed := true
    g_HotkeyCopy_DoubleTapTimer := ObjBindMethod(HotkeyCopy_DoubleTapTimerObj, "OnSingleTapTimeout")
    SetTimer(g_HotkeyCopy_DoubleTapTimer, -thresholdMs)
}

; Custom message so WindowManagement.ahk can trigger copy without Send (Send does not trigger hotkeys in another script).
WM_COPY_LAST_GEMINI := 0x8001
; Start background completion monitor for Ctrl+Alt+Win+L (wParam = originalHwnd, lParam = geminiHwnd). Sent from Utils.ahk.
WM_START_DELAYED_SUBMIT_MONITOR := 0x8002
; Stop any running delayed-submit monitor (e.g. when user chose S or N at 6s dictation confirm). Sent from Utils.ahk.
WM_STOP_DELAYED_SUBMIT_MONITOR := 0x8003
; Trigger read aloud from another script (e.g. D2C "Copy response?" R). Send does not trigger hotkeys in another script.
; wParam: 1 = caller already copied (skip Copy in Gemini). lParam: anchored original hwnd for focus restore (0 = resolve default).
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
    ; lParam: anchored original hwnd from dictation D2C (0 = resolve default).
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
