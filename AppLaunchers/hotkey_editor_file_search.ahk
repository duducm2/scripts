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

EditorFileSearch_QueryMatchesBasename(query, basename) {
    q := EditorFileSearch_NormalizeQuery(query)
    b := StrLower(EditorFileSearch_NormalizeBasename(basename))
    if (q = "" || b = "")
        return false
    if (b = q)
        return true
    qStem := EditorFileSearch_StripExtension(q)
    bStem := EditorFileSearch_StripExtension(b)
    if (qStem != "" && bStem != "" && (qStem = bStem || InStr(bStem, qStem) || InStr(qStem, bStem)))
        return true
    return InStr(b, q) || InStr(q, b)
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
        if (alreadyHwnd) {
            try WinActivate("ahk_id " alreadyHwnd)
            EditorFileSearch_SaveHit(alreadyHwnd, query)
            return
        }
    }

    g_EditorFileSearchActive := true
    try {
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

    if !EditorFileSearch_IsEditorProcessHwnd(hwnd)
        return false

    try {
        WinActivate("ahk_id " hwnd)
    } catch {
        return false
    }
    if (!WinWaitActive("ahk_id " hwnd, , EDITOR_FILE_SEARCH_ACTIVATE_TIMEOUT_SEC))
        return false

    if (EDITOR_FILE_SEARCH_SKIP_IF_ALREADY_OPEN && EditorFileSearch_HwndAlreadyHasFile(hwnd, query))
        return true

    Send "{Escape}"
    Sleep 50

    Send "^p"
    if (!EditorFileSearch_WaitForQuickInput(hwnd, EDITOR_FILE_SEARCH_QUICKINPUT_WAIT_MS)) {
        Send "{Escape}"
        return false
    }

    SendText query

    deadline := A_TickCount + EDITOR_FILE_SEARCH_RESULT_POLL_MS
    while (A_TickCount < deadline) {
        if (EditorFileSearch_QuickPickHasNoResults(hwnd)) {
            Send "{Escape}"
            return false
        }
        if (EditorFileSearch_HwndAlreadyHasFile(hwnd, query)) {
            Send "{Escape}"
            return true
        }
        Sleep EDITOR_FILE_SEARCH_POLL_STEP_MS
    }

    Send "{Enter}"
    if (EDITOR_FILE_SEARCH_USE_TITLE_VERIFY) {
        if (EditorFileSearch_WaitForTitleBasenameMatch(hwnd, query, EDITOR_FILE_SEARCH_TITLE_VERIFY_MS))
            return true
    }
    if (EditorFileSearch_WaitForQuickInputDismissed(hwnd, EDITOR_FILE_SEARCH_DISMISS_VERIFY_MS))
        return true

    Send "{Escape}"
    return false
}

EditorFileSearch_GetRoot(hwnd) {
    try {
        return UIA.ElementFromHandle(hwnd)
    } catch {
        return 0
    }
}

EditorFileSearch_QuickInputVisible(hwnd) {
    root := EditorFileSearch_GetRoot(hwnd)
    if (!root)
        return false
    try {
        if (root.FindFirst({ Type: 50004, ClassName: "quick-input-widget", cs: false }))
            return true
    } catch {
    }
    try {
        fe := UIA.GetFocusedElement()
        if (fe && fe.ControlType = UIA.Type.Edit)
            return true
    } catch {
    }
    return false
}

EditorFileSearch_WaitForQuickInput(hwnd, timeoutMs) {
    global EDITOR_FILE_SEARCH_POLL_STEP_MS
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        if (EditorFileSearch_QuickInputVisible(hwnd))
            return true
        Sleep EDITOR_FILE_SEARCH_POLL_STEP_MS
    }
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