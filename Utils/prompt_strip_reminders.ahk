; =============================================================================
; Utils module: prompt_strip_reminders.ahk
; Strip human-reminder blocks after the last --- divider in AI prompt text.
; Loaded via #include into Utils.ahk.
; =============================================================================

global g_PendingPromptHumanReminders := ""

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

PromptHumanReminders_ClearPending(*) {
    global g_PendingPromptHumanReminders
    g_PendingPromptHumanReminders := ""
}

PromptHumanReminders_OnYes(*) {
    global g_PendingPromptHumanReminders, g_lastExpansion
    reminders := g_PendingPromptHumanReminders
    g_PendingPromptHumanReminders := ""
    if (reminders = "")
        return
    g_lastExpansion := 0
    InsertText(reminders)
}

PromptHumanReminders_OnTimeout(*) {
    PromptHumanReminders_ClearPending()
}

; Strip in memory, paste once, then optional 3s Y banner to append reminders.
PasteStrippedPromptOfferReminders(fullText) {
    global g_PendingPromptHumanReminders
    stripped := StripPromptHumanReminders(fullText)
    reminders := ExtractPromptHumanReminders(fullText)
    if (stripped = "") {
        InsertText(fullText)
        return
    }
    InsertText(stripped)
    if (reminders = "")
        return
    g_PendingPromptHumanReminders := reminders
    StandardLoadingBar_CloseKeysOverlay()
    StandardLoadingBar_Hide(0)
    keyCallbacks := Map("Y", PromptHumanReminders_OnYes, "Escape", PromptHumanReminders_ClearPending)
    StandardLoadingBar_ShowWithKeys(
        "❓ Paste human reminders? (3s)",
        keyCallbacks,
        3000,
        0,
        PromptHumanReminders_OnTimeout,
        BANNER_ACCENT_INTERMEDIATE,
        420,
        17,
        "",
        true,
        "[Y] Yes  [Esc] Skip",
        true,
        true,
        true
    )
}
