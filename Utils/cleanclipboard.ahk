; =============================================================================
; Utils module: cleanclipboard.ahk
; Clean clipboard countdown macro
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; Clean the Clipboard macro function
; Shows a non-modal 4s countdown; auto-continues unless user cancels with N; Y proceeds immediately.
CleanClipboard() {
    CleanClipboard_ShowCountdown()
}

; Strip markdown escape backslashes chat UIs add on copy (e.g. meeting\_id → meeting_id).
UnescapeMarkdownFromText(text) {
    if (text = "")
        return { text: "", count: 0 }
    if RegExMatch(text, '^\s*```[^\R]*\R(.*?)\R```\s*$', &fence)
        text := fence[1]
    count := 0
    for ch in ["*", "_", "{", "}", "[", "]", "(", ")", "#", "+", ".", "!", "-", Chr(96), Chr(34)] {
        replaced := 0
        text := StrReplace(text, '\' . ch, ch, , &replaced)
        count += replaced
    }
    replaced := 0
    text := StrReplace(text, '\\', '\', , &replaced)
    count += replaced
    return { text: text, count: count }
}

UnescapeMarkdownClipboard() {
    clip := A_Clipboard
    if (Trim(clip) = "") {
        ShowCenteredOverlay_Utils("Clipboard empty", 1500, BANNER_ACCENT_INTERMEDIATE)
        return
    }
    result := UnescapeMarkdownFromText(clip)
    if (result.text = clip) {
        ShowCenteredOverlay_Utils("No markdown escapes found", 1500, BANNER_ACCENT_INTERMEDIATE)
        return
    }
    saved := ClipboardAll()
    try {
        A_Clipboard := result.text
        ClipWait(0.3)
    } finally {
        if (A_Clipboard != result.text)
            A_Clipboard := saved
    }
    msg := (result.count = 1) ? "Removed 1 markdown escape" : "Removed " result.count " markdown escapes"
    ShowCenteredOverlay_Utils("📋 " msg, 1500, BANNER_ACCENT_SUCCESS)
    try ScriptSoundPlay(A_ScriptDir . "\sounds\copy.wav")
}

CleanClipboard_ShowCountdown() {
    global g_CleanClipboardInProgress

    if (g_CleanClipboardInProgress) {
        ShowCenteredOverlay_Utils("Clipboard cleanup already running", 1500, BANNER_ACCENT_INTERMEDIATE)
        return
    }

    CleanClipboard_BeginSession()

    ; Ensure any previous keys overlay is closed before showing a new one
    try StandardLoadingBar_CloseKeysOverlay()
    catch {
    }
    StandardLoadingBar_Hide(0)
    Sleep 50

    state := "❓ Clean the clipboard? (removes stored clips, 4s)`nPress [Y] to clean now, or [N] within 4s to cancel."
    keyCallbacks := Map("N", CleanClipboard_OnCancel, "Y", CleanClipboard_OnYConfirm, "*Escape",
        CleanClipboard_OnCancel)

    ; Center on active monitor (centerOnHwnd := 0), use red accent for destructive action.
    StandardLoadingBar_ShowWithKeys(
        state,
        keyCallbacks,
        4000,
        0,
        CleanClipboard_OnTimeout.Bind(g_CleanClipboardSessionId),
        BANNER_ACCENT_ERROR,
        0,
        17,
        "",
        false,
        "[Y] Clean now  [N] Cancel (auto-continue in 4s)",
        true,
        true,
        true)
}

CleanClipboard_OnCancel(*) {
    global g_CleanClipboardCanceled
    g_CleanClipboardCanceled := true
    try StandardLoadingBar_CloseKeysOverlay()
    catch {
    }
    StandardLoadingBar_Hide(0)
    ShowCenteredOverlay_Utils("Clipboard cleanup cancelled", 1500, BANNER_ACCENT_INTERMEDIATE)
}

CleanClipboard_Proceed(sessionId := 0) {
    global g_CleanClipboardCanceled, g_CleanClipboardInProgress, g_CleanClipboardSessionId

    if (g_CleanClipboardCanceled)
        return
    if (g_CleanClipboardInProgress)
        return
    if (sessionId && sessionId != g_CleanClipboardSessionId)
        return

    g_CleanClipboardInProgress := true
    try {
        if (g_CleanClipboardCanceled)
            return
        try StandardLoadingBar_CloseKeysOverlay()
        catch {
        }
        StandardLoadingBar_Hide(0)
        CleanClipboard_SetAbortHotkeys(true)
        CleanClipboardInternal(sessionId)
    } finally {
        CleanClipboard_SetAbortHotkeys(false)
        CleanClipboard_EndSession()
    }
}

CleanClipboard_OnYConfirm(*) {
    global g_CleanClipboardCanceled, g_CleanClipboardSessionId, g_CleanClipboardProceedClaimed
    if (g_CleanClipboardCanceled)
        return
    ; Hotkey + poll can both fire Y before Proceed sets inProgress
    if (g_CleanClipboardProceedClaimed)
        return
    g_CleanClipboardProceedClaimed := true
    PlayCleaningDesktopSound()
    CleanClipboard_Proceed(g_CleanClipboardSessionId)
}

; Auto-continue when countdown ends (no chime; only Y plays the sound)
CleanClipboard_OnTimeout(sessionId, *) {
    global g_CleanClipboardCanceled, g_CleanClipboardProceedClaimed
    if (g_CleanClipboardCanceled)
        return
    if (g_CleanClipboardProceedClaimed)
        return
    g_CleanClipboardProceedClaimed := true
    CleanClipboard_Proceed(sessionId)
}
