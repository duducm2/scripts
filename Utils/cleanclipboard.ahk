; =============================================================================
; Utils module: cleanclipboard.ahk
; Clean clipboard countdown macro
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; Clean the Clipboard macro function
; Runs immediately in the background without blocking or stealing focus.
CleanClipboard() {
    global g_CleanClipboardInProgress, g_CleanClipboardSessionId

    if (g_CleanClipboardInProgress) {
        ShowCenteredOverlay_Utils("Clipboard cleanup already running", 1500, BANNER_ACCENT_INTERMEDIATE)
        return
    }

    CleanClipboard_BeginSession()
    try StandardLoadingBar_CloseKeysOverlay()
    catch {
    }
    StandardLoadingBar_Hide(0)
    SetTimer((*) => CleanClipboard_Proceed(g_CleanClipboardSessionId), -1)
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
    try ScriptSoundPlay(A_ScriptDir . "\assets\sounds\copy.wav")
}

CleanClipboard_ShowCountdown() {
    CleanClipboard()
}

CleanClipboard_OnCancel(*) {
    global g_CleanClipboardCanceled
    g_CleanClipboardCanceled := true
    try CleanClipboard_CancelActiveRun()
    catch {
    }
    try StandardLoadingBar_CloseKeysOverlay()
    catch {
    }
    StandardLoadingBar_Hide(0)
    ShowCenteredOverlay_Utils("Clipboard cleanup cancelled", 1500, BANNER_ACCENT_INTERMEDIATE)
}

CleanClipboard_IsOsClipboardEmpty() {
    try return (DllCall("CountClipboardFormats", "int") = 0)
    catch
        return false
}

CleanClipboard_TryClearOsClipboardOnce() {
    if (CleanClipboard_IsOsClipboardEmpty())
        return true
    try {
        if (DllCall("OpenClipboard", "Ptr", 0)) {
            emptied := false
            try emptied := DllCall("EmptyClipboard")
            finally DllCall("CloseClipboard")
            if (emptied && CleanClipboard_IsOsClipboardEmpty())
                return true
        }
    } catch {
    }
    try {
        A_Clipboard := ""
        if (CleanClipboard_IsOsClipboardEmpty() || A_Clipboard = "")
            return true
    } catch {
    }
    return false
}

CleanClipboard_ClearOsClipboard(retries := 6) {
    loop retries {
        if (CleanClipboard_TryClearOsClipboardOnce())
            return true
        if (A_Index < retries)
            Sleep 30
    }
    return false
}

CleanClipboard_StartClearOsClipboardAsync(retries := 8, onDone := 0, shouldAbortFn := 0) {
    hasDoneCb := (IsObject(onDone) && HasMethod(onDone, "Call"))
    if (CleanClipboard_TryClearOsClipboardOnce()) {
        if (hasDoneCb)
            onDone(true)
        return 0
    }
    hasAbortFn := (IsObject(shouldAbortFn) && HasMethod(shouldAbortFn, "Call"))
    state := { attempt: 1, max: retries, timerFn: 0, cancelled: false }
    step(*) {
        if (state.cancelled || (hasAbortFn && shouldAbortFn())) {
            state.cancelled := true
            SetTimer(state.timerFn, 0)
            if (hasDoneCb)
                onDone(false)
            return
        }
        state.attempt += 1
        if (CleanClipboard_TryClearOsClipboardOnce()) {
            SetTimer(state.timerFn, 0)
            if (hasDoneCb)
                onDone(true)
            return
        }
        if (state.attempt >= state.max) {
            SetTimer(state.timerFn, 0)
            if (hasDoneCb)
                onDone(false)
        }
    }
    state.timerFn := step
    SetTimer(step, 25)
    return state
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
    startedAsync := false
    try {
        if (g_CleanClipboardCanceled)
            return
        try StandardLoadingBar_CloseKeysOverlay()
        catch {
        }
        CleanClipboard_SetAbortHotkeys(true)
        startedAsync := CleanClipboardInternal(sessionId)
    } catch Error as e {
        try CleanClipboard_LogError(0, "Proceed", e.Message)
        catch {
        }
        try StandardLoadingBar_Hide(0)
        catch {
        }
        ShowCenteredOverlay_Utils("❌ Clipboard cleanup error: " e.Message, 3500, BANNER_ACCENT_ERROR)
    } finally {
        if (!startedAsync) {
            CleanClipboard_SetAbortHotkeys(false)
            CleanClipboard_EndSession()
        }
    }
}

; Auto-continue when countdown ends. Success chime plays inside CleanClipboardInternal.
CleanClipboard_OnTimeout(sessionId, *) {
    global g_CleanClipboardCanceled, g_CleanClipboardProceedClaimed
    if (g_CleanClipboardCanceled)
        return
    if (g_CleanClipboardProceedClaimed)
        return
    g_CleanClipboardProceedClaimed := true
    ; Show continuous loading banner immediately before KeysOverlay closes so there is no visual gap.
    try StandardLoadingBar_Show("⏳ Cleaning clipboard & Clip Angel...", BANNER_ACCENT_INTERMEDIATE, {
        passive: false,
        manualProgress: true,
        noActivate: true,
        fontSize: 17
    })
    catch {
    }
    SetTimer((*) => CleanClipboard_Proceed(sessionId), -1)
}
