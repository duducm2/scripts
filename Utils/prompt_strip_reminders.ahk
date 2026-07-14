; =============================================================================
; Utils module: prompt_strip_reminders.ahk
; Strip human-reminder blocks after the last --- divider in AI prompt text.
; Loaded via #include into Utils.ahk.
; =============================================================================

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
