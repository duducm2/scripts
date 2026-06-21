; =============================================================================
; Gemini module: gemini_open.ahk
; InitializeGeminiFirstTime and #!+I open/focus hotkey
; Extracted verbatim from Gemini.ahk; loaded via #include into the
; Gemini.ahk process, which remains the entry point / source of truth.
; =============================================================================

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

        ; First-launch bootstrap: tab 1 -> Pro, tab 2 -> Fast, then return to tab 1.
        StandardLoadingBar_Update("⚙️ Configuring Gemini tabs...", BANNER_ACCENT_INTERMEDIATE)
        GeminiConfigureFirstLaunchTabModels(geminiHwnd)

        StandardLoadingBar_Hide(0)
    } catch Error as err {
        ; Hide banner5 on error
        StandardLoadingBar_Hide(0)
    }
}

; =============================================================================
; Open Gemini
; Hotkey: Win+Alt+Shift+I
; =============================================================================
#!+i:: {
    if UseCopilotWebForGlobalAI()
        return CopilotWeb_OpenOrFocus()
    t0 := A_TickCount
    SetTitleMatchMode(2)
    if hwnd := GetGeminiWindowHwnd() {
        alreadyActive := false
        try
            alreadyActive := WinActive("ahk_id " hwnd)
        catch {
        }
        if (!alreadyActive) {
            try {
                WinActivate("ahk_id " hwnd)
            } catch {
                ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
                return
            }
            if !WinWaitActive("ahk_id " hwnd, , GEMINI_ACTIVATE_WAIT_MS // 1000)
                return
        }

        if (alreadyActive && GEMINI_OPEN_FAST_SETTLE_MS > 0)
            Sleep GEMINI_OPEN_FAST_SETTLE_MS

        try
            uia := UIA_Browser("ahk_id " hwnd)
        catch {
            ShowCenteredOverlay_Utils("❌ Error: Could not attach to Gemini window.", 2000, BANNER_ACCENT_ERROR)
            return
        }

        if (!alreadyActive)
            Gemini_WaitForPromptFieldDiscoverable(uia)

        focusPhase := ""
        promptField := Gemini_FocusPromptWithChime(uia, "", &focusPhase)
        if (promptField) {
            tBanner := A_TickCount
            SetTimer(() => (Gemini_ShowDeferredTabBanner(uia), GeminiPerfLog("tab_banner_deferred", tBanner)), -1)
        }
        GeminiPerfLog(focusPhase != "" ? focusPhase : "focus_failed", t0)
        GeminiPerfLog("activation", t0)
    } else {
        InitializeGeminiFirstTime()
    }
}

; Ready chime for flows that do not use Gemini_FocusPromptWithChime (e.g. legacy call sites).
Gemini_PlayReadyChime(minIntervalMs := 400) {
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
