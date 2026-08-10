; =============================================================================
; Utils module: dictation_legacy.ahk
; Deprecated dictation Gemini confirm banner
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; Dictation: "Send dictation?" confirmation banner (5s, Y to confirm; uses standard loading indicator)
; =============================================================================
; DEPRECATED: Use D2C_FlowManager
DEPRECATED_DictationGeminiConfirm_Show() {
    ; No-op; use DictationGeminiConfirm_ShowAndWait() which uses StandardLoadingBar_ShowWithKeys.
}

DEPRECATED_DictationGeminiConfirm_Hide(*) {
    StandardLoadingBar_CloseKeysOverlay()
}

; submitToGemini=false (N or timeout): terminal. submitToGemini=true: delayed-submit (paste+Enter). pasteOnly=true: paste to Gemini only, no Enter.
DEPRECATED_DictationGeminiConfirm_CleanupAndMaybeSubmit(submitToGemini, pasteOnly := false) {
    global g_DictationGeminiConfirmBannerVisible
    ; Only clear banner-visible when proceeding (Y/S/timeout); leave true on N so a stray ShowAndWait does not re-show and register a second 5s timer (logs showed second timeout firing after N).
    if (submitToGemini || pasteOnly)
        g_DictationGeminiConfirmBannerVisible := false
    ; Unregister 5s banner keys (same * prefix as StandardLoadingBar_RegisterKeyHandler uses)
    try Hotkey("*y", "Off")
    try Hotkey("*Y", "Off")
    try Hotkey("*s", "Off")
    try Hotkey("*S", "Off")
    try Hotkey("*n", "Off")
    try Hotkey("*N", "Off")
    SetTimer(DEPRECATED_DictationGeminiConfirm_OnTimeout, 0)
    DEPRECATED_DictationGeminiConfirm_Hide()
    ; S or N at 5s: no submit flow, so stop any running "Copy response?" monitor so it never shows.
    if (!submitToGemini)
        GeminiDelayedSubmitMonitorStopFromUtils()
    if (pasteOnly) {
        Sleep 350
        DEPRECATED_GeminiDictationPasteOnlyFlow()
    } else if (submitToGemini) {
        Sleep 350
        DEPRECATED_GeminiDelayedSubmitFlow()
    }
}

DEPRECATED_DictationGeminiConfirm_OnY(*) {
    DEPRECATED_DictationGeminiConfirm_CleanupAndMaybeSubmit(true)
}

; S = paste to Gemini only (no Enter, no 4s banner).
DEPRECATED_DictationGeminiConfirm_OnS(*) {
    DEPRECATED_DictationGeminiConfirm_CleanupAndMaybeSubmit(false, true)
}

; Default action on 5s timeout: proceed as Yes (DelayedSubmitFlow), same as user pressing Y.
DEPRECATED_DictationGeminiConfirm_OnTimeout(*) {
    DEPRECATED_DictationGeminiConfirm_CleanupAndMaybeSubmit(true)
}

; N = terminate flow: no paste, no Enter, no 4s, no copy; only cleanup and cancel overlay.
DEPRECATED_DictationGeminiConfirm_OnCancel(*) {
    DEPRECATED_DictationGeminiConfirm_CleanupAndMaybeSubmit(false)  ; submitToGemini=false, pasteOnly=false => no flow runs
    ShowCenteredOverlay_Utils("⚠ Gemini submission cancelled", 1500, BANNER_ACCENT_INTERMEDIATE)
}

DEPRECATED_DictationGeminiConfirm_ShowAndWait() {
    global g_DictationGeminiConfirmBannerVisible
    ; Only one banner: atomic check-and-set so only one invocation can pass (prevents duplicate from multiple PlayDictationCompletionChime runs).
    Critical "On"
    if (g_DictationGeminiConfirmBannerVisible) {
        Critical "Off"
        return
    }
    g_DictationGeminiConfirmBannerVisible := true
    Critical "Off"
    ; Only the official loading bar (standard loading indicator) may show this content. Hide any other bar/overlay first.
    StandardLoadingBar_CloseKeysOverlay()
    StandardLoadingBar_Hide(0)
    HideDictationIndicator()
    Sleep 50
    keyCallbacks := Map("Y", DEPRECATED_DictationGeminiConfirm_OnY, "S", DEPRECATED_DictationGeminiConfirm_OnS, "N",
        DEPRECATED_DictationGeminiConfirm_OnCancel)
    ; Official loading bar only; no blue; single banner (no border); fixed bottom strip for input.
    StandardLoadingBar_ShowWithKeys("❓ Send to " . GetGlobalAIProviderLabel() . "? (5s)", keyCallbacks,
    D2C_SUBMIT_MENU_TIMEOUT_MS,
    0,
    DEPRECATED_DictationGeminiConfirm_OnTimeout, BANNER_ACCENT_INTERMEDIATE, 380, 17, "", true,
    "[Y] Send  [S] Paste only  [N] Cancel",
    true,
    true)
}
