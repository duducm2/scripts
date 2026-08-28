; =============================================================================
; Utils module: prompt_strip_reminders.ahk
; Strip human-reminder blocks after the last --- divider in AI prompt text.
; Loaded via #include into Utils.ahk.
; =============================================================================

global g_PendingPromptPasteFullText := ""
global g_PendingPromptPasteOnAfter := ""
global g_PendingPromptPasteRestoreFocus := ""

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

PromptPaste_ClearPending(*) {
    global g_PendingPromptPasteFullText, g_PendingPromptPasteOnAfter, g_PendingPromptPasteRestoreFocus
    g_PendingPromptPasteFullText := ""
    g_PendingPromptPasteOnAfter := ""
    g_PendingPromptPasteRestoreFocus := ""
}

PromptPaste_ExecuteChoice(choice) {
    global g_PendingPromptPasteFullText, g_PendingPromptPasteOnAfter, g_PendingPromptPasteRestoreFocus
    fullText := g_PendingPromptPasteFullText
    onAfter := g_PendingPromptPasteOnAfter
    restoreFocus := g_PendingPromptPasteRestoreFocus
    PromptPaste_ClearPending()
    if (fullText = "")
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
    if (onAfter != "") {
        try onAfter()
        catch {
        }
    }
    if (doSend)
        Send "{Enter}"
}

PromptPaste_OnIncludeReminders(*) {
    PromptPaste_ExecuteChoice("reminders")
}

PromptPaste_OnPasteOnly(*) {
    PromptPaste_ExecuteChoice("strip")
}

PromptPaste_OnSendNow(*) {
    PromptPaste_ExecuteChoice("send")
}

PromptPaste_OnTimeout(*) {
    PromptPaste_OnPasteOnly()
}

; Pre-paste banner: Y = full text, Esc/timeout = strip, S = strip + Enter after paste (and onAfter).
PromptPaste_ShowOptionsThenPaste(fullText, onAfterPaste := "", restoreFocus := "") {
    global g_PendingPromptPasteFullText, g_PendingPromptPasteOnAfter, g_PendingPromptPasteRestoreFocus
    if (fullText = "")
        return
    g_PendingPromptPasteFullText := fullText
    g_PendingPromptPasteOnAfter := onAfterPaste
    g_PendingPromptPasteRestoreFocus := restoreFocus
    StandardLoadingBar_CloseKeysOverlay()
    StandardLoadingBar_Hide(0)
    keyCallbacks := Map(
        "Y", PromptPaste_OnIncludeReminders,
        "S", PromptPaste_OnSendNow,
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
        "[Y] Include reminders  [S] Send now  [Esc] Paste",
        true,
        true,
        true
    )
}
