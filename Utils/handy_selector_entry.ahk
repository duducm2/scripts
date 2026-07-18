; =============================================================================
; Utils module: handy_selector_entry.ahk
; SelectAiModelInHandy entry and pre-movement warning
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; SelectAiModelInHandy() - Opens or closes the selector GUI (hotkey entry point)
; =============================================================================
; Select AI model in Handy via interactive GUI selector.
; Win+Alt+Shift+C toggles: open when closed, close when open.
; Targets the correct Handy instance by environment: work = Documents\Handy\handy.exe, home = any.
SelectAiModelInHandy() {
    if (!HandyAi_IsOwnerProcess())
        return
    global g_AiModelSelectorActive
    if (g_AiModelSelectorActive)
        AiModelSelector_Close()
    else
        ShowAiModelSelector()
}

; =============================================================================
; Helper: Show centered overlay banner (uses standard loading indicator; non-blocking).
; =============================================================================
ShowCenteredOverlay_Utils(text, duration := 1500, bgColor := BANNER_ACCENT_INTERMEDIATE) {
    ; centerOnHwnd 0 = foreground monitor (GetActiveMonitorWorkArea_StandardBar); same intent as prior WinGetID("A") path.
    StandardLoadingBar_Show(text, bgColor, { passive: true, centerOnHwnd: 0, textWidth: 500, fontSize: 17,
        passiveBgColor: bgColor })
    if (duration < 1)
        duration := 1
    StandardLoadingBar_Hide(duration)
    ; Hard max (5s default) so a missed/raced hide cannot leave the banner stuck.
    ; Keys overlays use ShowWithKeys and do not call this path.
    StandardLoadingBar_ArmForceHide()
}

; =============================================================================
; Helper: Pre-movement warning (sound + 2s delay) before automated window changes.
; =============================================================================
PlayPreMovementWarning(targetName) {
    ScriptSoundPlay(A_ScriptDir . "\assets\sounds\pre-movement.wav")
    ShowCenteredOverlay_Utils("✋ Hands off! Moving to " . targetName . "...", 2000, BANNER_ACCENT_INTERMEDIATE)
    Sleep 2000
}
