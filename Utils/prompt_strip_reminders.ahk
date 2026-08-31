; =============================================================================
; Utils module: prompt_strip_reminders.ahk
; Strip human-reminder blocks after the last --- divider in AI prompt text.
; Loaded via #include into Utils.ahk.
; =============================================================================

global g_PromptPasteWaitChoice := ""
global g_PromptPasteWaitActive := false

; Keep text through the last standalone "---" line, then two blank lines for comments.
; Returns "" if no --- divider is found (caller shows error).
StripPromptHumanReminders(text) {
    if (text = "")
        return ""
    t := StrReplace(StrReplace(text, "`r`n", "`n"), "`r", "`n")
    lines := StrSplit(t, "`n")
    last := 0
    for line in lines {
        if (Trim(line) = "---")
            last := A_Index
    }
    if (!last)
        return ""
    out := ""
    loop last {
        if (A_Index > 1)
            out .= "`n"
        out .= lines[A_Index]
    }
    return out . "`n`n"
}

; Text after the last standalone "---" line (human reminders). "" if no divider.
ExtractPromptHumanReminders(text) {
    if (text = "")
        return ""
    t := StrReplace(StrReplace(text, "`r`n", "`n"), "`r", "`n")
    lines := StrSplit(t, "`n")
    last := 0
    for line in lines {
        if (Trim(line) = "---")
            last := A_Index
    }
    if (!last || last >= lines.Length)
        return ""
    out := ""
    loop (lines.Length - last) {
        if (A_Index > 1)
            out .= "`n"
        out .= lines[last + A_Index]
    }
    return out
}

; Replace the focused edit (composer/prompt) with newText via select-all + paste.
ReplaceFocusedEditWithText(newText) {
    if (newText = "")
        return false
    saved := ClipboardAll()
    try {
        A_Clipboard := newText
        if !ClipWait(1, 1)
            return false
        Send "^a"
        Sleep 40
        Send "^v"
        Sleep 80
        return true
    } finally {
        Sleep 100
        try A_Clipboard := saved
        catch {
        }
    }
}

; After a full prompt is already in the composer, strip reminders from fullText in memory
; and replace the composer (native Ctrl+Z restores the full paste). Returns false if no ---.
ReplaceComposerWithStrippedReminders(fullText, settleMs := 250) {
    stripped := StripPromptHumanReminders(fullText)
    if (stripped = "")
        return false
    if (settleMs > 0)
        Sleep settleMs
    return ReplaceFocusedEditWithText(stripped)
}

PromptPaste_ResolveText(fullText, includeReminders) {
    if (includeReminders)
        return fullText
    stripped := StripPromptHumanReminders(fullText)
    return (stripped != "") ? stripped : fullText
}

PromptPaste_FinishWait(choice) {
    global g_PromptPasteWaitChoice, g_PromptPasteWaitActive
    if (!g_PromptPasteWaitActive)
        return
    g_PromptPasteWaitChoice := choice
    g_PromptPasteWaitActive := false
}

PromptPaste_OnIncludeReminders(*) {
    PromptPaste_FinishWait("reminders")
}

PromptPaste_OnPasteOnly(*) {
    PromptPaste_FinishWait("strip")
}

PromptPaste_OnSendNow(*) {
    PromptPaste_FinishWait("send")
}

PromptPaste_OnTimeout(*) {
    PromptPaste_FinishWait("strip")
}

; Show banner immediately; block until Y / H / N / Esc / timeout. Returns "reminders", "strip", or "send".
PromptPaste_ShowOptionsAndWait() {
    global g_PromptPasteWaitChoice, g_PromptPasteWaitActive
    g_PromptPasteWaitChoice := ""
    g_PromptPasteWaitActive := true
    StandardLoadingBar_CloseKeysOverlay()
    StandardLoadingBar_Hide(0)
    keyCallbacks := Map(
        "Y", PromptPaste_OnSendNow,
        "H", PromptPaste_OnIncludeReminders,
        "N", PromptPaste_OnPasteOnly,
        "Escape", PromptPaste_OnPasteOnly
    )
    StandardLoadingBar_ShowWithKeys(
        "❓ Paste prompt? (3s)",
        keyCallbacks,
        3000,
        0,
        PromptPaste_OnTimeout,
        BANNER_ACCENT_INTERMEDIATE,
        420,
        17,
        "",
        true,
        "[Y] Send now  [H] Human notes  [N]/Esc Paste",
        true,
        true,
        true
    )
    while (g_PromptPasteWaitActive)
        Sleep 30
    return g_PromptPasteWaitChoice
}

PromptPaste_ApplyChoice(choice, fullText, onAfterPaste := "", restoreFocus := "", submitOpts := "") {
    if (fullText = "" || choice = "")
        return
    includeReminders := (choice = "reminders")
    doSend := (choice = "send")
    if (restoreFocus != "") {
        try restoreFocus()
        catch {
        }
        Sleep 80
    }
    textToPaste := PromptPaste_ResolveText(fullText, includeReminders)
    InsertText(textToPaste)
    if (onAfterPaste != "") {
        try onAfterPaste()
        catch {
        }
    }
    if (doSend) {
        hwnd := 0
        companionId := ""
        attachCount := 0
        if (IsObject(submitOpts)) {
            if (submitOpts.HasProp("hwnd"))
                hwnd := submitOpts.hwnd
            if (submitOpts.HasProp("companionId"))
                companionId := submitOpts.companionId
            if (submitOpts.HasProp("attachCount"))
                attachCount := submitOpts.attachCount
        }
        PromptPaste_SubmitWhenReady(hwnd, companionId, attachCount)
    }
}
