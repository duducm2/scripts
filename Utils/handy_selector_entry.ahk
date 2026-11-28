; =============================================================================
; Utils module: handy_selector_entry.ahk
; SelectAiModelInHandy entry and ShowCenteredOverlay_Utils
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; SelectAiModelInHandy() - Opens or closes the selector GUI (Utility Shortcuts [C])
; =============================================================================
; Select AI model in Handy via interactive GUI selector.
; Targets the correct Handy instance by environment: work = Documents\Handy\handy.exe, home = any.
; Blocks opening while a background model switch is already running.
SelectAiModelInHandy(restoreHwnd := 0) {
    global g_AiModelSelectorActive, g_HandyModelSwitchBusy
    if (g_HandyModelSwitchBusy)
        return
    if (g_AiModelSelectorActive)
        AiModelSelector_Close()
    else
        ShowAiModelSelector(restoreHwnd)
}

; =============================================================================
; Helper: Show centered overlay banner (uses standard loading indicator; non-blocking).
; =============================================================================
ShowCenteredOverlay_Utils(text, duration := 1500, bgColor := BANNER_ACCENT_INTERMEDIATE) {
    if (duration < 1)
        duration := 1
    ; centerOnHwnd 0 = foreground monitor (GetActiveMonitorWorkArea_StandardBar); same intent as prior WinGetID("A") path.
    StandardLoadingBar_Show(text, bgColor, { passive: true, noActivate: true, centerOnHwnd: 0, textWidth: 500, fontSize: 17,
        passiveBgColor: bgColor, maxDurationMs: duration + 2000 })
    StandardLoadingBar_Hide(duration)
}
