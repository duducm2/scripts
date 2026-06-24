; =============================================================================
; Utils module: hotstrings_core.ahk
; Prompt text helpers and safe clipboard paste (InsertText / InsertFiles).
; Loaded via #include into the Utils.ahk orchestrator / shared library entry point.
; =============================================================================

global g_lastExpansion := 0

; Safe paste insertion to avoid app shortcuts and re-triggers
InsertText(text) {
    global g_lastExpansion
    if (A_TickCount - g_lastExpansion) < 250
        return
    g_lastExpansion := A_TickCount

    saved := ClipboardAll()
    try {
        A_Clipboard := text
        ClipWait(0.3)
        Sleep 50
        Send "^v"
    } finally {
        Sleep 150
        A_Clipboard := saved
    }
}

; Paste file(s) via CF_HDROP (file attachment), not path text. Returns true on success.
InsertFiles(paths) {
    global g_lastExpansion
    if (A_TickCount - g_lastExpansion) < 250
        return false
    if (!paths || paths.Length = 0)
        return false
    for path in paths {
        if !Clipboard_PathIsExistingFile(path)
            return false
    }

    g_lastExpansion := A_TickCount
    saved := ClipboardAll()
    ok := false
    try {
        if !Clipboard_SetFiles(paths)
            return false
        if !Clipboard_WaitForFileDrop(800)
            return false
        Sleep 50
        Send "^v"
        ok := true
        Sleep 400
    } finally {
        try A_Clipboard := saved
    }
    return ok
}

InsertFiles_IsAiChatForeground() {
    try {
        title := WinGetTitle("A")
        if (title = "")
            return false
        lower := StrLower(title)
        return InStr(lower, "gemini") || InStr(lower, "copilot")
    } catch {
        return false
    }
}

GetPromptDir() {
    return A_ScriptDir "\assets\prompt"
}

GetPromptText(key) {
    try {
        return FileRead(GetPromptDir() "\" key ".txt")
    } catch {
        return "[PROMPT FILE MISSING: " key "]"
    }
}
