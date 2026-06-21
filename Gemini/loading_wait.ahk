; =============================================================================
; Gemini module: loading_wait.ahk
; Small loading indicator and WaitForButton helpers
; Extracted verbatim from Gemini.ahk; loaded via #include into the
; Gemini.ahk process, which remains the entry point / source of truth.
; =============================================================================

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
