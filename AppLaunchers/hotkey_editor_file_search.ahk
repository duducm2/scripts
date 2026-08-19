; =============================================================================
; AppLaunchers module: hotkey_editor_file_search.ahk
; Win+Alt+Shift+K — search and open a file across VS Code and Cursor windows
; Loaded via #include into AppLaunchers.ahk
; =============================================================================

global g_EditorFileSearchActive := false
global EDITOR_FILE_SEARCH_ACTIVATE_TIMEOUT_SEC := 2
global EDITOR_FILE_SEARCH_QUICKINPUT_WAIT_MS := 800
global EDITOR_FILE_SEARCH_RESULT_POLL_MS := 600
global EDITOR_FILE_SEARCH_DISMISS_VERIFY_MS := 800
global EDITOR_FILE_SEARCH_POLL_STEP_MS := 40
global EDITOR_FILE_SEARCH_USE_HIT_CACHE := true
global EDITOR_FILE_SEARCH_SKIP_IF_ALREADY_OPEN := true
global EDITOR_FILE_SEARCH_USE_TITLE_VERIFY := true
global EDITOR_FILE_SEARCH_TITLE_VERIFY_MS := 1500
global EDITOR_FILE_SEARCH_TITLE_STABLE_POLLS := 2

; #region agent log
EditorFileSearch_DebugJsonEscape(s) {
    s := StrReplace(String(s), "\", "\\")
    s := StrReplace(s, '"', '\"')
    s := StrReplace(s, "`n", "\n")
    s := StrReplace(s, "`r", "")
    return s
}

EditorFileSearch_DebugJoin(arr, delim := ",") {
    s := ""
    for i, v in arr
        s .= (i = 1 ? "" : delim) v
    return s
}

EditorFileSearch_DebugLog(hypothesisId, location, message, data := "") {
    try {
        logPath := A_ScriptDir "\debug-c69194.log"
        dataJson := "{}"
        if IsObject(data) && data.Count > 0 {
            pairs := []
            for k, v in data {
                key := EditorFileSearch_DebugJsonEscape(k)
                if (v is Integer || v is Float)
                    pairs.Push('"' key '":' v)
                else if (v = true || v = false)
                    pairs.Push('"' key '":' (v ? "true" : "false"))
                else
                    pairs.Push('"' key '":"' EditorFileSearch_DebugJsonEscape(v) '"')
            }
            dataJson := "{" EditorFileSearch_DebugJoin(pairs) "}"
        }
        line := '{"sessionId":"c69194","hypothesisId":"' EditorFileSearch_DebugJsonEscape(hypothesisId)
        . '","location":"' EditorFileSearch_DebugJsonEscape(location)
        . '","message":"' EditorFileSearch_DebugJsonEscape(message)
        . '","data":' dataJson ',"timestamp":' A_TickCount '}' "`n"
        FileAppend line, logPath, "UTF-8"
    } catch as err {
        try FileAppend '{"sessionId":"c69194","message":"log_fail","data":"'
            . EditorFileSearch_DebugJsonEscape(err.Message) '"}' "`n", A_ScriptDir "\debug-c69194.log", "UTF-8"
        catch {
        }
    }
}
; #endregion

EditorFileSearch_NormalizeBasename(raw) {
    if (raw = "")
        return ""
    s := Trim(raw)
    s := RegExReplace(s, "^[●\s]+")
    if InStr(s, ",")
        s := Trim(SubStr(s, 1, InStr(s, ",") - 1))
    if (InStr(s, "\") || InStr(s, "/")) {
        SplitPath s, &name
        if (name != "")
            s := name
    }
    return s
}

EditorFileSearch_IsPlausibleBasename(raw) {
    s := EditorFileSearch_NormalizeBasename(raw)
    if (s = "" || StrLen(s) > 180)
        return false
    lower := StrLower(s)
    if InStr(lower, "not accessible") || InStr(lower, "screen reader") || InStr(lower, "agentswindow")
        return false
    if InStr(s, "`n") || InStr(s, "`r")
        return false
    return true
}

EditorFileSearch_BasenameFromTitle(hwnd) {
    if !(hwnd is Integer) || hwnd <= 0
        return ""
    try {
        title := WinGetTitle("ahk_id " hwnd)
        if (title = "")
            return ""
        parts := StrSplit(title, " - ", , 2)
        if (parts.Length >= 1 && parts[1] != "") {
            candidate := EditorFileSearch_NormalizeBasename(Trim(parts[1]))
            if EditorFileSearch_IsPlausibleBasename(candidate)
                return candidate
        }
    } catch {
    }
    return ""
}

EditorFileSearch_NormalizeQuery(query) {
    q := StrLower(Trim(query))
    if (q = "")
        return ""
    if (InStr(q, "\") || InStr(q, "/")) {
        SplitPath q, &nameOnly
        if (nameOnly != "")
            q := StrLower(nameOnly)
    }
    return q
}

EditorFileSearch_StripExtension(name) {
    if (name = "")
        return ""
    return RegExReplace(name, "\.[^.\\\/]+$")
}

EditorFileSearch_QueryTokens(query) {
    q := EditorFileSearch_NormalizeQuery(query)
    tokens := []
    if (q = "")
        return tokens
    for part in StrSplit(q, A_Space) {
        part := Trim(part)
        if (part != "")
            tokens.Push(part)
    }
    return tokens
}

EditorFileSearch_BasenameWords(b) {
    words := []
    if (b = "")
        return words
    for part in StrSplit(b, "_-. " Chr(9)) {
        part := Trim(part)
        if (part != "")
            words.Push(part)
    }
    return words
}

EditorFileSearch_WordPrefixMatchesToken(word, token) {
    if (word = token)
        return true
    if (StrLen(token) > StrLen(word))
        return false
    return SubStr(word, 1, StrLen(token)) = token
}

EditorFileSearch_TokenMatchesBasenameWords(token, b, bStem) {
    if (token = "" || b = "")
        return false
    for word in EditorFileSearch_BasenameWords(b) {
        if EditorFileSearch_WordPrefixMatchesToken(word, token)
            return true
    }
    if (bStem != "" && bStem != b) {
        for word in EditorFileSearch_BasenameWords(bStem) {
            if EditorFileSearch_WordPrefixMatchesToken(word, token)
                return true
        }
    }
    return false
}

EditorFileSearch_QueryMatchesBasename(query, basename) {
    b := StrLower(EditorFileSearch_NormalizeBasename(basename))
    if (b = "")
        return false
    tokens := EditorFileSearch_QueryTokens(query)
    if (tokens.Length = 0)
        return false
    if (tokens.Length = 1) {
        q := tokens[1]
        if (b = q)
            return true
        qStem := EditorFileSearch_StripExtension(q)
        bStem := EditorFileSearch_StripExtension(b)
        if (qStem != "" && bStem != "" && (qStem = bStem || InStr(bStem, qStem) || InStr(qStem, bStem)))
            return true
        return InStr(b, q) || InStr(q, b)
    }
    bStem := EditorFileSearch_StripExtension(b)
    for token in tokens {
        if !EditorFileSearch_TokenMatchesBasenameWords(token, b, bStem)
            return false
    }
    return true
}

EditorFileSearch_HwndAlreadyHasFile(hwnd, query) {
    basename := EditorFileSearch_BasenameFromTitle(hwnd)
    if (basename = "")
        return false
    return EditorFileSearch_QueryMatchesBasename(query, basename)
}

EditorFileSearch_FindHwndAlreadyOpen(query, hwnds) {
    for hwnd in hwnds {
        if (EditorFileSearch_HwndAlreadyHasFile(hwnd, query))
            return hwnd
    }
    return 0
}

EditorFileSearch_WaitForTitleBasenameMatch(hwnd, query, timeoutMs) {
    global EDITOR_FILE_SEARCH_POLL_STEP_MS, EDITOR_FILE_SEARCH_TITLE_STABLE_POLLS
    deadline := A_TickCount + timeoutMs
    stable := 0
    need := EDITOR_FILE_SEARCH_TITLE_STABLE_POLLS
    while (A_TickCount < deadline) {
        if (EditorFileSearch_HwndAlreadyHasFile(hwnd, query)) {
            stable++
            if (stable >= need)
                return true
        } else {
            stable := 0
        }
        Sleep EDITOR_FILE_SEARCH_POLL_STEP_MS
    }
    return false
}

EditorFileSearch_IniPath() {
    return A_ScriptDir "\assets\data\editor_file_search.ini"
}

EditorFileSearch_NormalizeQueryKey(query) {
    key := StrLower(Trim(query))
    return StrReplace(key, "=", "`=")
}

EditorFileSearch_LoadHitWindowTitle(query) {
    global EDITOR_FILE_SEARCH_USE_HIT_CACHE
    if !EDITOR_FILE_SEARCH_USE_HIT_CACHE
        return ""
    key := EditorFileSearch_NormalizeQueryKey(query)
    if (key = "")
        return ""
    try {
        return Trim(IniRead(EditorFileSearch_IniPath(), "Hits", key, ""))
    } catch {
        return ""
    }
}

EditorFileSearch_SaveHit(hwnd, query) {
    global EDITOR_FILE_SEARCH_USE_HIT_CACHE
    if !EDITOR_FILE_SEARCH_USE_HIT_CACHE
        return
    if !(hwnd is Integer) || hwnd <= 0
        return
    if !EditorFileSearch_IsEditorProcessHwnd(hwnd)
        return
    key := EditorFileSearch_NormalizeQueryKey(query)
    if (key = "")
        return
    try {
        title := WinGetTitle("ahk_id " hwnd)
        if (title = "")
            return
        iniPath := EditorFileSearch_IniPath()
        SplitPath(iniPath, , &dir)
        if (dir != "" && !DirExist(dir))
            DirCreate(dir)
        IniWrite(title, iniPath, "Hits", key)
    } catch {
    }
}

EditorFileSearch_FindHwndBySavedTitle(savedTitle) {
    if (savedTitle = "")
        return 0
    savedLower := StrLower(savedTitle)
    for exe in ["ahk_exe Cursor.exe", "ahk_exe Code.exe"] {
        try {
            for hwnd in WinGetList(exe) {
                if !EditorFileSearch_IsEligibleEditorHwnd(hwnd)
                    continue
                try {
                    if (StrLower(WinGetTitle("ahk_id " hwnd)) = savedLower)
                        return hwnd
                } catch {
                }
            }
        } catch {
        }
    }
    return 0
}

EditorFileSearch_Start() {
    global g_EditorFileSearchActive
    if (g_EditorFileSearchActive)
        return

    query := Trim(InputBox("Enter target file name:", "Open File in Editor").Value)
    if (query = "")
        return

    hwnds := EditorFileSearch_CollectEditorHwnds(query)
    if (hwnds.Length = 0) {
        ShowCenteredOverlay_Utils("File not found in any editor window", 2000, BANNER_ACCENT_ERROR)
        return
    }

    global EDITOR_FILE_SEARCH_SKIP_IF_ALREADY_OPEN
    if (EDITOR_FILE_SEARCH_SKIP_IF_ALREADY_OPEN) {
        alreadyHwnd := EditorFileSearch_FindHwndAlreadyOpen(query, hwnds)
        ; #region agent log
        EditorFileSearch_DebugLog("A", "Start:alreadyOpenScan", "global already-open scan", Map(
            "alreadyHwnd", alreadyHwnd, "query", query,
            "basename", alreadyHwnd ? EditorFileSearch_BasenameFromTitle(alreadyHwnd) : ""))
        ; #endregion
        if (alreadyHwnd) {
            try WinActivate("ahk_id " alreadyHwnd)
            EditorFileSearch_SaveHit(alreadyHwnd, query)
            ; #region agent log
            EditorFileSearch_DebugLog("A", "Start:alreadyOpenSkip", "return before Quick Open", Map("hwnd", alreadyHwnd
            ))
            ; #endregion
            return
        }
    }

    g_EditorFileSearchActive := true
    try {
        ; #region agent log
        EditorFileSearch_DebugLog("E", "Start:tryLoop", "entering hwnd loop", Map("hwndCount", hwnds.Length, "query",
            query))
        ; #endregion
        for hwnd in hwnds {
            if (EditorFileSearch_TryOpenInInstance(hwnd, query)) {
                EditorFileSearch_SaveHit(hwnd, query)
                return
            }
        }
        ShowCenteredOverlay_Utils("File not found in any editor window", 2000, BANNER_ACCENT_ERROR)
    } finally {
        g_EditorFileSearchActive := false
    }
}

EditorFileSearch_CollectEditorHwnds(query := "") {
    global EDITOR_FILE_SEARCH_USE_HIT_CACHE
    result := []
    seen := Map()

    if (EDITOR_FILE_SEARCH_USE_HIT_CACHE && query != "") {
        savedTitle := EditorFileSearch_LoadHitWindowTitle(query)
        if (savedTitle != "") {
            hitHwnd := EditorFileSearch_FindHwndBySavedTitle(savedTitle)
            if (hitHwnd) {
                result.Push(hitHwnd)
                seen[hitHwnd] := true
            }
        }
    }

    try {
        fg := WinGetID("A")
        if (fg && EditorFileSearch_IsEligibleEditorHwnd(fg) && !seen.Has(fg)) {
            result.Push(fg)
            seen[fg] := true
        }
    } catch {
    }

    for exe in ["ahk_exe Cursor.exe", "ahk_exe Code.exe"] {
        try {
            for hwnd in WinGetList(exe) {
                if (seen.Has(hwnd))
                    continue
                if (EditorFileSearch_IsEligibleEditorHwnd(hwnd)) {
                    result.Push(hwnd)
                    seen[hwnd] := true
                }
            }
        } catch {
        }
    }
    return result
}

EditorFileSearch_IsEditorProcessHwnd(hwnd) {
    try {
        proc := WinGetProcessName("ahk_id " hwnd)
        return (proc = "Cursor.exe" || proc = "Code.exe")
    } catch {
        return false
    }
}

EditorFileSearch_IsEligibleEditorHwnd(hwnd) {
    if !(hwnd is Integer) || hwnd <= 0
        return false
    if !EditorFileSearch_IsEditorProcessHwnd(hwnd)
        return false
    if !DllCall("IsWindowVisible", "ptr", hwnd)
        return false
    try {
        if (WinGetMinMax("ahk_id " hwnd) = -1)
            return false
    } catch {
        return false
    }
    try {
        title := WinGetTitle("ahk_id " hwnd)
        if (InStr(StrLower(title), "preview"))
            return false
    } catch {
        return false
    }
    return true
}

EditorFileSearch_TryOpenInInstance(hwnd, query) {
    global EDITOR_FILE_SEARCH_ACTIVATE_TIMEOUT_SEC, EDITOR_FILE_SEARCH_QUICKINPUT_WAIT_MS,
        EDITOR_FILE_SEARCH_RESULT_POLL_MS, EDITOR_FILE_SEARCH_DISMISS_VERIFY_MS,
        EDITOR_FILE_SEARCH_POLL_STEP_MS, EDITOR_FILE_SEARCH_USE_TITLE_VERIFY,
        EDITOR_FILE_SEARCH_TITLE_VERIFY_MS, EDITOR_FILE_SEARCH_SKIP_IF_ALREADY_OPEN

    ; #region agent log
    EditorFileSearch_DebugLog("E", "TryOpen:enter", "TryOpenInInstance", Map("hwnd", hwnd, "query", query))
    ; #endregion

    if !EditorFileSearch_IsEditorProcessHwnd(hwnd) {
        ; #region agent log
        EditorFileSearch_DebugLog("E", "TryOpen:exit", "not editor process", Map("hwnd", hwnd, "ok", false))
        ; #endregion
        return false
    }

    try {
        WinActivate("ahk_id " hwnd)
    } catch {
        ; #region agent log
        EditorFileSearch_DebugLog("E", "TryOpen:exit", "WinActivate failed", Map("hwnd", hwnd, "ok", false))
        ; #endregion
        return false
    }
    if (!WinWaitActive("ahk_id " hwnd, , EDITOR_FILE_SEARCH_ACTIVATE_TIMEOUT_SEC)) {
        ; #region agent log
        EditorFileSearch_DebugLog("E", "TryOpen:exit", "WinWaitActive timeout", Map("hwnd", hwnd, "ok", false))
        ; #endregion
        return false
    }

    if (EDITOR_FILE_SEARCH_SKIP_IF_ALREADY_OPEN && EditorFileSearch_HwndAlreadyHasFile(hwnd, query)) {
        ; #region agent log
        EditorFileSearch_DebugLog("A", "TryOpen:exit", "per-instance already-open skip", Map(
            "hwnd", hwnd, "basename", EditorFileSearch_BasenameFromTitle(hwnd), "ok", true))
        ; #endregion
        return true
    }

    EditorFileSearch_ReleaseHotkeyModifiers()
    Send "{Escape}"
    Sleep 50

    Send "^p"
    waitOk := EditorFileSearch_WaitForQuickInput(hwnd, EDITOR_FILE_SEARCH_QUICKINPUT_WAIT_MS)
    uiaMiss := !waitOk
    if (uiaMiss && WinActive("ahk_id " hwnd))
        waitOk := true
    ; #region agent log
    fgBefore := 0
    try fgBefore := WinGetID("A")
    EditorFileSearch_DebugLog("B", "TryOpen:afterWait", "WaitForQuickInput", Map(
        "hwnd", hwnd, "waitOk", waitOk, "uiaMiss", uiaMiss, "fgHwnd", fgBefore,
        "widget", EditorFileSearch_QuickInputVisible(hwnd)))
    ; #endregion
    if (!waitOk) {
        Send "{Escape}"
        ; #region agent log
        EditorFileSearch_DebugLog("B", "TryOpen:exit", "quick input wait failed", Map("hwnd", hwnd, "ok", false))
        ; #endregion
        return false
    }

    ; #region agent log
    fgSend := 0
    try fgSend := WinGetID("A")
    EditorFileSearch_DebugLog("C", "TryOpen:beforeSendText", "about to type query", Map(
        "hwnd", hwnd, "fgHwnd", fgSend, "fgMatches", fgSend = hwnd,
        "queryLen", StrLen(query), "widget", EditorFileSearch_QuickInputVisible(hwnd)))
    ; #endregion
    typeMethod := EditorFileSearch_TypeQueryIntoQuickOpen(hwnd, query)
    ; #region agent log
    fgAfter := 0
    try fgAfter := WinGetID("A")
    EditorFileSearch_DebugLog("C", "TryOpen:afterSendText", "type query completed", Map(
        "hwnd", hwnd, "fgHwnd", fgAfter, "typeMethod", typeMethod,
        "widget", EditorFileSearch_QuickInputVisible(hwnd)))
    ; #endregion

    deadline := A_TickCount + EDITOR_FILE_SEARCH_RESULT_POLL_MS
    while (A_TickCount < deadline) {
        if (EditorFileSearch_QuickPickHasNoResults(hwnd)) {
            Send "{Escape}"
            ; #region agent log
            EditorFileSearch_DebugLog("D", "TryOpen:exit", "no results in poll", Map("hwnd", hwnd, "ok", false))
            ; #endregion
            return false
        }
        if (EditorFileSearch_HwndAlreadyHasFile(hwnd, query)) {
            Send "{Escape}"
            ; #region agent log
            EditorFileSearch_DebugLog("D", "TryOpen:exit", "title matched in poll before Enter", Map(
                "hwnd", hwnd, "basename", EditorFileSearch_BasenameFromTitle(hwnd), "ok", true))
            ; #endregion
            return true
        }
        Sleep EDITOR_FILE_SEARCH_POLL_STEP_MS
    }

    Send "{Enter}"
    if (EDITOR_FILE_SEARCH_USE_TITLE_VERIFY) {
        if (EditorFileSearch_WaitForTitleBasenameMatch(hwnd, query, EDITOR_FILE_SEARCH_TITLE_VERIFY_MS)) {
            ; #region agent log
            EditorFileSearch_DebugLog("D", "TryOpen:exit", "title verify success", Map("hwnd", hwnd, "ok", true))
            ; #endregion
            return true
        }
    }
    if (EditorFileSearch_WaitForQuickInputDismissed(hwnd, EDITOR_FILE_SEARCH_DISMISS_VERIFY_MS)) {
        ; #region agent log
        EditorFileSearch_DebugLog("D", "TryOpen:exit", "dismiss verify success", Map("hwnd", hwnd, "ok", true))
        ; #endregion
        return true
    }

    Send "{Escape}"
    ; #region agent log
    EditorFileSearch_DebugLog("D", "TryOpen:exit", "exhausted verify", Map("hwnd", hwnd, "ok", false))
    ; #endregion
    return false
}

EditorFileSearch_ReleaseHotkeyModifiers() {
    Send "{LWin up}{RWin up}{Alt up}{Shift up}{Ctrl up}"
    Sleep 30
}

EditorFileSearch_GetRoot(hwnd) {
    try {
        return UIA.ElementFromHandle(hwnd)
    } catch {
        return 0
    }
}

EditorFileSearch_QuickInputWidget(hwnd) {
    root := EditorFileSearch_GetRoot(hwnd)
    if (!root)
        return 0
    try {
        return root.FindFirst({ Type: 50004, ClassName: "quick-input-widget", cs: false })
    } catch {
        return 0
    }
}

EditorFileSearch_QuickInputVisible(hwnd) {
    if (EditorFileSearch_QuickInputWidget(hwnd))
        return true
    root := EditorFileSearch_GetRoot(hwnd)
    if (root) {
        try {
            if (root.FindFirst({ AutomationId: "quickInput.list.filter", cs: false }))
                return true
        } catch {
        }
    }
    if WinActive("ahk_id " hwnd) {
        try {
            fe := UIA.GetFocusedElement()
            if (fe && fe.ControlType = UIA.Type.Edit)
                return true
        } catch {
        }
    }
    return false
}

EditorFileSearch_TypeQueryIntoQuickOpen(hwnd, query) {
    if !WinActive("ahk_id " hwnd) {
        try WinActivate("ahk_id " hwnd)
        if !WinWaitActive("ahk_id " hwnd, , 1)
            return "no_foreground"
    }
    try {
        ControlSendText query, , "ahk_id " hwnd
        return "control_sendtext"
    } catch {
    }
    try {
        SendText query
        return "sendtext"
    } catch {
        return "type_failed"
    }
}

EditorFileSearch_WaitForQuickInput(hwnd, timeoutMs) {
    global EDITOR_FILE_SEARCH_POLL_STEP_MS
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        if (EditorFileSearch_QuickInputVisible(hwnd)) {
            ; #region agent log
            EditorFileSearch_DebugLog("B", "WaitForQuickInput:ready", "quick input detected", Map("hwnd", hwnd))
            ; #endregion
            return true
        }
        Sleep EDITOR_FILE_SEARCH_POLL_STEP_MS
    }
    ; #region agent log
    EditorFileSearch_DebugLog("B", "WaitForQuickInput:timeout", "no quick input before deadline", Map("hwnd", hwnd))
    ; #endregion
    return false
}

EditorFileSearch_WaitForQuickInputDismissed(hwnd, timeoutMs) {
    global EDITOR_FILE_SEARCH_POLL_STEP_MS
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        if (!EditorFileSearch_QuickInputVisible(hwnd))
            return true
        Sleep EDITOR_FILE_SEARCH_POLL_STEP_MS
    }
    return false
}

EditorFileSearch_QuickPickHasNoResults(hwnd) {
    root := EditorFileSearch_GetRoot(hwnd)
    if (!root)
        return false
    for name in ["No matching results", "Nenhum resultado", "Sin resultados"] {
        try {
            if (root.FindFirst({ Name: name, matchmode: "Substring", cs: false }))
                return true
        } catch {
        }
    }
    return false
}

; =============================================================================
; Hotkey: Win+Alt+Shift+K — open file in VS Code / Cursor via Quick Open
; =============================================================================
#!+k:: EditorFileSearch_Start()