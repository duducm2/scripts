; =============================================================================
; Utils module: hotstrings_core.ahk
; Hotstrings core (InsertText, prompt file helpers)
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; Hotstring Core Functions
; =============================================================================

; ----------------------
; Safer hotstrings core
; ----------------------
global g_hotstrings := []
global g_lastExpansion := 0

; Cross-process IPC for Hotstring Selector (symmetric with WindowManagement project selector IPC)
global g_HS_SelectorOpenFile := A_ScriptDir "\.cursor\hs_selector_open"
global g_HS_SelectorCloseRequestFile := A_ScriptDir "\.cursor\hs_selector_close_request"
global g_HS_SelectorCloseCheckTimer := ""

Utils_CheckHotstringSelectorCloseRequest() {
    global g_HotstringSelectorActive, g_HS_SelectorCloseRequestFile
    if (!g_HotstringSelectorActive)
        return
    if (FileExist(g_HS_SelectorCloseRequestFile)) {
        try FileDelete(g_HS_SelectorCloseRequestFile)
        catch {
        }
        CleanupHotstringSelector()
    }
}

RegisterHotstring(trigger, expansion, category := "", title := "", char := "") {
    global g_hotstrings
    g_hotstrings.Push({ trigger: trigger, expansion: expansion, category: category, title: title, char: char })
}

; Safe paste insertion to avoid app shortcuts and re-triggers
InsertText(text) {
    global g_lastExpansion
    ; Debounce to prevent rapid duplicate expansions (e.g., double Space)
    if (A_TickCount - g_lastExpansion) < 250
        return
    g_lastExpansion := A_TickCount

    saved := ClipboardAll()
    try {
        A_Clipboard := text
        ClipWait(0.3)
        Sleep 50  ; Give time for clipboard to fully update
        Send "^v"
    } finally {
        Sleep 150  ; Wait longer for paste to complete before restoring clipboard
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
        ; Brief settle so browser upload handlers receive the paste before clipboard restore.
        ; Multi-file packs need a longer settle before the attach helper's upload-idle wait.
        Sleep (paths.Length > 1) ? 800 : 400
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
        return InStr(lower, "gemini") || InStr(lower, "copilot") || InStr(lower, "askbosch") || InStr(lower,
            "vertex")
    } catch {
        return false
    }
}

GetPromptDir() {
    return A_ScriptDir "\assets\prompt"
}

; Prompt .txt files are UTF-8; default FileRead uses ANSI on Windows without BOM.
ReadUtf8File(path) {
    return FileRead(path, "UTF-8")
}

WriteUtf8File(path, content) {
    try FileDelete(path)
    catch {
    }
    FileAppend(content, path, "UTF-8")
}

GetPromptText(key) {
    try {
        return ReadUtf8File(GetPromptDir() "\" key ".txt")
    } catch {
        return "[PROMPT FILE MISSING: " key "]"
    }
}

; Technique prompts: scripts mirror first, then live notes repo if present.
GetTechniquePromptFilePath(fileName) {
    mirror := A_ScriptDir "\mnemonics\technique\prompts\" fileName
    if FileExist(mirror)
        return mirror
    repo := GetNotesRepoPath()
    dir := (repo != "") ? repo "\studies\technique\prompts" : ""
    if (dir != "" && FileExist(dir "\" fileName))
        return dir "\" fileName
    legacy := A_ScriptDir "\assets\prompt\technique\" fileName
    if FileExist(legacy)
        return legacy
    return mirror
}

; Prompts / Hotstrings / Projects for #!+U now live in INI (prompt_data.ahk, hotstring_data.ahk, projects.ini).
