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

EditorFileSearch_Start() {
    global g_EditorFileSearchActive
    if (g_EditorFileSearchActive)
        return

    query := Trim(InputBox("Enter target file name:", "Open File in Editor").Value)
    if (query = "")
        return

    hwnds := EditorFileSearch_CollectEditorHwnds()
    if (hwnds.Length = 0) {
        ShowCenteredOverlay_Utils("File not found in any editor window", 2000, BANNER_ACCENT_ERROR)
        return
    }

    g_EditorFileSearchActive := true
    try {
        for hwnd in hwnds {
            if (EditorFileSearch_TryOpenInInstance(hwnd, query))
                return
        }
        ShowCenteredOverlay_Utils("File not found in any editor window", 2000, BANNER_ACCENT_ERROR)
    } finally {
        g_EditorFileSearchActive := false
    }
}

EditorFileSearch_CollectEditorHwnds() {
    result := []
    seen := Map()

    try {
        fg := WinGetID("A")
        if (fg && EditorFileSearch_IsEligibleEditorHwnd(fg)) {
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

EditorFileSearch_IsTargetEditorForeground(hwnd) {
    if !(hwnd is Integer) || hwnd <= 0
        return false
    if !WinActive("ahk_id " hwnd)
        return false
    return EditorFileSearch_IsEditorProcessHwnd(hwnd)
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
        EDITOR_FILE_SEARCH_POLL_STEP_MS

    if !EditorFileSearch_IsEditorProcessHwnd(hwnd)
        return false

    try {
        WinActivate("ahk_id " hwnd)
    } catch {
        return false
    }
    if (!WinWaitActive("ahk_id " hwnd, , EDITOR_FILE_SEARCH_ACTIVATE_TIMEOUT_SEC))
        return false
    if !EditorFileSearch_IsTargetEditorForeground(hwnd)
        return false

    Send "{Escape}"
    Sleep 50

    if !EditorFileSearch_IsTargetEditorForeground(hwnd)
        return false
    Send "^p"
    if (!EditorFileSearch_WaitForQuickInput(hwnd, EDITOR_FILE_SEARCH_QUICKINPUT_WAIT_MS)) {
        if (EditorFileSearch_IsTargetEditorForeground(hwnd))
            Send "{Escape}"
        return false
    }

    if !EditorFileSearch_IsTargetEditorForeground(hwnd)
        return false
    SendText query

    deadline := A_TickCount + EDITOR_FILE_SEARCH_RESULT_POLL_MS
    while (A_TickCount < deadline) {
        if !EditorFileSearch_IsTargetEditorForeground(hwnd)
            return false
        if (EditorFileSearch_QuickPickHasNoResults(hwnd)) {
            Send "{Escape}"
            return false
        }
        Sleep EDITOR_FILE_SEARCH_POLL_STEP_MS
    }

    if !EditorFileSearch_IsTargetEditorForeground(hwnd)
        return false
    Send "{Enter}"
    if (EditorFileSearch_WaitForQuickInputDismissed(hwnd, EDITOR_FILE_SEARCH_DISMISS_VERIFY_MS))
        return true

    if (EditorFileSearch_IsTargetEditorForeground(hwnd))
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
    if !EditorFileSearch_IsTargetEditorForeground(hwnd)
        return false
    root := EditorFileSearch_GetRoot(hwnd)
    if (!root)
        return false
    try {
        if (root.FindFirst({ Type: 50004, ClassName: "quick-input-widget", cs: false }))
            return true
    } catch {
    }
    try {
        if (root.FindFirst({ Type: 50004, AutomationId: "quickInput.list.filter", cs: false }))
            return true
    } catch {
    }
    return false
}

EditorFileSearch_WaitForQuickInput(hwnd, timeoutMs) {
    global EDITOR_FILE_SEARCH_POLL_STEP_MS
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        if !EditorFileSearch_IsTargetEditorForeground(hwnd)
            return false
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
        if !EditorFileSearch_IsTargetEditorForeground(hwnd)
            return false
        if (!EditorFileSearch_QuickInputVisible(hwnd))
            return true
        Sleep EDITOR_FILE_SEARCH_POLL_STEP_MS
    }
    return false
}

EditorFileSearch_QuickPickHasNoResults(hwnd) {
    if !EditorFileSearch_IsTargetEditorForeground(hwnd)
        return false
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