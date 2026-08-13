; =============================================================================
; Utils module: gemini_cursor_transfer.ahk
; Gemini-to-Cursor transfer numeric window selector
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; Gemini-to-Cursor transfer: numeric Cursor window selector (1-9) and activate/focus/paste
; Used when user presses [C] Transfer in Gemini copy-decision banner.
; =============================================================================
global g_CursorTransferSelectorGui := false
global g_CursorTransferSelectorActive := false
global g_CursorTransferSelectorResult := ""   ; "" = waiting, 0 = cancel, integer = selected hwnd
global g_CursorTransferWindowList := []      ; up to 9 { hwnd, title }
global g_CursorTransferHotkeyHandlers := []
global g_CursorTransferPidCmdCache := Map()
global g_CursorTransferLastHandledIndex := 0 ; Diagnostic: track last selected index

; === Environment-aware transfer target resolution ===

; Return transfer target app executable based on IS_WORK_ENVIRONMENT: "Cursor.exe" or "Code.exe"
CursorTransfer_GetTargetAppExecutable() {
    if (IsSet(IS_WORK_ENVIRONMENT) && IS_WORK_ENVIRONMENT)
        return "Code.exe"
    return "Cursor.exe"
}

; Return display name for transfer target app: "Cursor" or "VS Code"
CursorTransfer_GetTargetAppName() {
    if (IsSet(IS_WORK_ENVIRONMENT) && IS_WORK_ENVIRONMENT)
        return "VS Code"
    return "Cursor"
}

CursorTransfer_SelectorClose() {
    global g_CursorTransferSelectorActive, g_CursorTransferSelectorGui, g_CursorTransferHotkeyHandlers
    if (!g_CursorTransferSelectorActive)
        return
    g_CursorTransferSelectorActive := false
    for h in g_CursorTransferHotkeyHandlers {
        if (IsObject(h) && h.HasProp("key") && h.HasProp("callback")) {
            try Hotkey(h.key, h.callback, "Off")
        } else {
            try Hotkey(h, "Off")
        }
    }
    g_CursorTransferHotkeyHandlers := []
    if (IsObject(g_CursorTransferSelectorGui) && g_CursorTransferSelectorGui.Hwnd) {
        try g_CursorTransferSelectorGui.Destroy()
    }
    g_CursorTransferSelectorGui := false
}

CursorTransfer_SelectorHandleKey(index, *) {
    global g_CursorTransferSelectorResult, g_CursorTransferWindowList, g_CursorTransferLastHandledIndex

    ; Validate index bounds
    if (!IsInteger(index) || index < 1 || index > g_CursorTransferWindowList.Length) {
        ; Invalid index: keep selector open, log for diagnostics
        g_CursorTransferLastHandledIndex := index
        return
    }

    ; Retrieve target hwnd
    targetItem := g_CursorTransferWindowList[index]
    if (!targetItem || !targetItem.HasProp("hwnd")) {
        g_CursorTransferLastHandledIndex := index
        return
    }

    targetHwnd := targetItem.hwnd

    ; Verify hwnd still exists before accepting selection
    if (!WinExist("ahk_id " targetHwnd)) {
        g_CursorTransferLastHandledIndex := index
        return
    }

    ; Valid selection: set result and close
    g_CursorTransferSelectorResult := targetHwnd
    g_CursorTransferLastHandledIndex := index
    CursorTransfer_SelectorClose()
}

CursorTransfer_SelectorEscape(*) {
    global g_CursorTransferSelectorResult
    g_CursorTransferSelectorResult := 0
    CursorTransfer_SelectorClose()
}

; Return project order index from g_Projects for a window title; 0 = no match.
CursorTransfer_GetProjectOrderForTitle(winTitle) {
    idx := CursorTransfer_GetMatchingProjectIndexForTitle(winTitle)
    return idx
}

; Build canonical project-index -> character mapping (same logic as standard project selector).
CursorTransfer_BuildProjectIndexToChar() {
    global g_Projects
    ProjectData_Load()
    projectIndexToChar := Map()
    taken := Map()
    loop g_Projects.Length {
        projectIndex := A_Index
        project := g_Projects[projectIndex]
        if (project.name = "" && project.path = "" && project.workPath = "")
            continue
        if (!project.HasProp("char") || project.char = "")
            continue
        ch := project.char
        if (ProjectData_IsValidChar(ch) && !taken.Has(ch)) {
            projectIndexToChar[projectIndex] := ch
            taken[ch] := true
        }
    }
    return projectIndexToChar
}

; Return matching project index from g_Projects for a window title; 0 = no match.
; Uses longest matching path segment so "user-scripts" wins over "scripts" when both match.
CursorTransfer_GetMatchingProjectIndexForTitle(winTitle) {
    global g_Projects
    ProjectData_Load()
    if (!winTitle || !IsObject(g_Projects))
        return 0
    try {
        isWork := (IsSet(IS_WORK_ENVIRONMENT) && IS_WORK_ENVIRONMENT)
        winLow := StrLower(winTitle)
        bestIdx := 0
        bestScore := 0
        loop g_Projects.Length {
            project := g_Projects[A_Index]
            if (project.name = "" && project.path = "" && project.workPath = "")
                continue
            projectPath := isWork ? project.workPath : project.path
            if (isWork && (projectPath = ""))
                projectPath := project.path
            if (projectPath = "")
                continue
            matchSegments := ExtractProjectMatchSegments(projectPath)
            for segment in matchSegments {
                if (segment = "")
                    continue
                if (InStr(winLow, StrLower(segment))) {
                    len := StrLen(segment)
                    if (len > bestScore) {
                        bestScore := len
                        bestIdx := A_Index
                    }
                }
            }
        }
        return bestIdx
    } catch {
    }
    return 0
}

; Return project path according to current environment for a given project object.
CursorTransfer_GetEffectiveProjectPath(project) {
    isWork := (IsSet(IS_WORK_ENVIRONMENT) && IS_WORK_ENVIRONMENT)
    projectPath := isWork ? project.workPath : project.path
    if (isWork && (projectPath = ""))
        projectPath := project.path
    return projectPath
}

; True when the OS title has no workspace/file segment (e.g. bare "Cursor" welcome screen).
CursorTransfer_IsUninformativeCursorTitle(winTitle) {
    t := Trim(winTitle)
    if (t = "")
        return true
    tl := StrLower(t)
    if (tl = "cursor")
        return true
    ; Strip trailing " - Cursor"; if nothing remains, title had no file/workspace name.
    rest := RegExReplace(tl, "\s*[---]\s*cursor\s*$", "")
    if (Trim(rest) = "")
        return true
    return false
}

; Longest project path wins when multiple g_Projects paths appear in the same command line.
CursorTransfer_GetMatchingProjectIndexByCmdLine(cmdLine) {
    global g_Projects
    ProjectData_Load()
    if (!cmdLine || !IsObject(g_Projects))
        return 0
    cmdLow := StrLower(cmdLine)
    bestIdx := 0
    bestLen := 0
    try {
        loop g_Projects.Length {
            project := g_Projects[A_Index]
            if (project.name = "" && project.path = "" && project.workPath = "")
                continue
            projectPath := CursorTransfer_GetEffectiveProjectPath(project)
            if (projectPath = "")
                continue
            projLow := StrLower(RTrim(projectPath, "\"))
            if (projLow != "" && InStr(cmdLow, projLow)) {
                len := StrLen(projLow)
                if (len > bestLen) {
                    bestLen := len
                    bestIdx := A_Index
                }
            }
        }
    } catch {
    }
    return bestIdx
}

; Get process command line by PID (cached). Returns "" on failure.
CursorTransfer_GetProcessCommandLine(pid) {
    global g_CursorTransferPidCmdCache
    if (!pid)
        return ""
    if (g_CursorTransferPidCmdCache.Has(pid))
        return g_CursorTransferPidCmdCache[pid]
    cmd := ""
    try {
        locator := ComObject("WbemScripting.SWbemLocator")
        svc := locator.ConnectServer(".", "root\cimv2")
        for proc in svc.ExecQuery("SELECT CommandLine FROM Win32_Process WHERE ProcessId = " pid) {
            try cmd := proc.CommandLine
            break
        }
    } catch {
        cmd := ""
    }
    g_CursorTransferPidCmdCache[pid] := cmd ? cmd : ""
    return g_CursorTransferPidCmdCache[pid]
}

; Title match when the title lists workspace/file; cmd-line when title is bare "Cursor" etc.
; (Same PID often shares one command line - use longest path in cmd for disambiguation.)
CursorTransfer_GetMatchingProjectIndex(hwnd, winTitle := "") {
    global g_Projects
    ProjectData_Load()
    if (!IsObject(g_Projects))
        return 0
    pid := 0
    try pid := WinGetPID("ahk_id " hwnd)
    cmdLine := CursorTransfer_GetProcessCommandLine(pid)
    idxCmd := CursorTransfer_GetMatchingProjectIndexByCmdLine(cmdLine)
    idxTitle := (winTitle != "") ? CursorTransfer_GetMatchingProjectIndexForTitle(winTitle) : 0
    if (CursorTransfer_IsUninformativeCursorTitle(winTitle)) {
        if (idxCmd > 0)
            return idxCmd
        return idxTitle
    }
    if (idxTitle > 0)
        return idxTitle
    return idxCmd
}

; Stable insertion sort for small arrays by project order, then by title.
CursorTransfer_SortWindowsByProjectOrder(&arr) {
    n := arr.Length
    if (n <= 1)
        return
    loop n - 1 {
        i := A_Index + 1
        key := arr[i]
        j := i - 1
        while (j >= 1) {
            left := arr[j]
            shouldShift := false
            ; Coerce to integer so comparison never sees a string (projectOrder can be string from g_Projects).
            try
                leftOrd := Integer(left.projectOrder)
            catch
                leftOrd := 0
            try
                keyOrd := Integer(key.projectOrder)
            catch
                keyOrd := 0
            if (leftOrd > keyOrd) {
                shouldShift := true
            } else if (leftOrd = keyOrd) {
                leftName := ""
                keyName := ""
                try
                    leftName := String(left.displayName)
                catch
                    leftName := ""
                try
                    keyName := String(key.displayName)
                catch
                    keyName := ""
                if (StrCompare(leftName, keyName) > 0)
                    shouldShift := true
            }
            if (!shouldShift)
                break
            arr[j + 1] := left
            j--
        }
        arr[j + 1] := key
    }
}

; Return project name from g_Projects if window title matches a project path; otherwise "".
CursorTransfer_GetProjectNameForTitle(winTitle) {
    global g_Projects
    idx := CursorTransfer_GetMatchingProjectIndexForTitle(winTitle)
    if (!idx || !IsObject(g_Projects))
        return ""
    try {
        project := g_Projects[idx]
        if (project.name != "")
            return project.name
    } catch {
    }
    return ""
}

CursorTransfer_StripStaticScriptTokenForDisplay(projectName) {
    ; Removes redundant static token(s) like "Script"/"Scripts" from the project label.
    ; If stripping would erase the whole name (e.g. folder is literally "Scripts"), keep the original.
    if (!projectName)
        return ""
    orig := Trim(projectName)
    cleaned := projectName
    ; Match whole-word "Script" or "Scripts" (case-insensitive).
    cleaned := RegExReplace(cleaned, "(?i)\bscript(s)?\b", "")
    cleaned := RegExReplace(cleaned, "\s{2,}", " ")
    cleaned := Trim(cleaned)
    if (StrLen(cleaned) < 2)
        return orig
    return cleaned
}

; Collapse "file.ext (Label) (file.ext)" in Cursor titles to a single filename + label.
CursorTransfer_StripDuplicateFilenameInParens(title) {
    if (!title)
        return ""
    return RegExReplace(title, "(\S+\.\w+)\s+(\([^)]+\))\s+\(\1\)", "$1 $2")
}

; If the window title starts with the project name (already shown in brackets), drop that prefix.
CursorTransfer_StripLeadingProjectFromTitle(title, cleanProjName, projName) {
    if (!title)
        return ""
    ; Longer label first so a shorter prefix cannot steal a match from a longer project name.
    names := []
    if (StrLen(cleanProjName) > StrLen(projName)) {
        if (cleanProjName != "")
            names.Push(cleanProjName)
        if (projName != "")
            names.Push(projName)
    } else {
        if (projName != "")
            names.Push(projName)
        if (cleanProjName != "" && cleanProjName != projName)
            names.Push(cleanProjName)
    }
    for name in names {
        if (StrLen(name) < 2)
            continue
        tl := StrLower(title)
        nl := StrLower(name)
        if (SubStr(tl, 1, StrLen(nl)) != nl)
            continue
        rest := SubStr(title, StrLen(name) + 1)
        rest := Trim(rest)
        if (rest = "")
            return ""
        if (SubStr(rest, 1, 1) = "-" || SubStr(rest, 1, 1) = "|" || SubStr(rest, 1, 1) = ":" || SubStr(rest, 1, 1) =
        "-")
            rest := Trim(SubStr(rest, 2))
        rest := Trim(LTrim(rest, "- "))
        return (rest != "") ? rest : title
    }
    return title
}

; Remove repetitive app suffix (e.g., " - Cursor", " - Visual Studio Code") from list labels; all windows are already from target app.
CursorTransfer_StripTrailingCursorAppSuffix(s) {
    if (!s)
        return ""
    t := Trim(s)
    targetApp := CursorTransfer_GetTargetAppName()
    if (StrLower(t) = StrLower(targetApp))
        return ""
    ; Strip " - Cursor", " - VS Code", " - Visual Studio Code", and similar variants
    result := Trim(RegExReplace(t, "i)\s*[---]\s*(?:Cursor|VS Code|Visual Studio Code)\s*$", ""))
    return result
}

Clipboard_GetSequenceNumber() {
    ; WinAPI: https://learn.microsoft.com/windows/win32/api/winuser/nf-winuser-getclipboardsequencenumber
    try {
        return DllCall("GetClipboardSequenceNumber", "uint")
    } catch {
        return 0
    }
}

Clipboard_WaitForSequenceChange(seqBefore, totalTimeoutMs := 2000, fastPhaseMs := 850) {
    ; Tight, bounded wait: returns true as soon as the clipboard sequence changes.
    ; Uses a fast polling phase first, then a slower phase for the remainder.
    start := A_TickCount
    deadline := start + totalTimeoutMs
    fastDeadline := start + fastPhaseMs
    while (A_TickCount < deadline) {
        seqNow := Clipboard_GetSequenceNumber()
        if (seqNow && seqNow != seqBefore)
            return true
        Sleep((A_TickCount < fastDeadline) ? 20 : 50)
    }
    return false
}

GetGeminiScriptMsgTargetHwnd() {
    ; Cache-first resolver for Gemini.ahk AutoHotkey script window.
    static cached := 0
    if (cached && WinExist("ahk_id " cached)) {
        try {
            if (InStr(WinGetTitle("ahk_id " cached), "Gemini.ahk"))
                return cached
        } catch {
        }
    }

    prevMatch := A_TitleMatchMode
    DetectHiddenWindows(true)
    SetTitleMatchMode(2)
    found := 0
    try {
        for hwnd in WinGetList("ahk_exe AutoHotkey64.exe") {
            try {
                if (InStr(WinGetTitle("ahk_id " hwnd), "Gemini.ahk")) {
                    found := hwnd
                    break
                }
            } catch {
                continue
            }
        }
        if (!found) {
            for hwnd in WinGetList("ahk_exe AutoHotkey32.exe") {
                try {
                    if (InStr(WinGetTitle("ahk_id " hwnd), "Gemini.ahk")) {
                        found := hwnd
                        break
                    }
                } catch {
                    continue
                }
            }
        }
    } finally {
        SetTitleMatchMode(prevMatch)
        DetectHiddenWindows(false)
    }

    if (found)
        cached := found
    return found
}

; Returns selected Cursor window HWND or 0 on cancel/timeout/no windows. Blocking with timeout.
; centerOnHwnd: optional window to center the modal on (uses that window's monitor); 0 = foreground monitor.
CursorTransfer_ShowWindowSelector(centerOnHwnd := 0) {
    global g_CursorTransferSelectorGui, g_CursorTransferSelectorActive, g_CursorTransferSelectorResult
    global g_CursorTransferWindowList, g_CursorTransferHotkeyHandlers
    global g_Projects, g_ProjectCharSequence
    CursorTransfer_SelectorClose()
    list := []

    ; Resolve target app executable based on environment
    targetAppExe := CursorTransfer_GetTargetAppExecutable()
    appDisplayName := CursorTransfer_GetTargetAppName()

    ; Only enumerate visible windows (DetectHiddenWindows off = default).
    ; Hidden Cursor/VS Code renderer/worker processes can have non-empty titles, pass the
    ; title filter, and end up at position 1 in the sorted list. Their hwnd cannot be found
    ; by WinExist in the hotkey handler thread (which also runs with DetectHiddenWindows off),
    ; causing the first-item selection to silently fail while all other items work fine.
    try {
        for hwnd in WinGetList("ahk_exe " targetAppExe) {
            try {
                title := WinGetTitle("ahk_id " hwnd)
                if (title = "" || InStr(StrLower(title), "preview"))
                    continue
                if (title = "MSCTFIME UI" || title = "Default IME")
                    continue
                list.Push({ hwnd: hwnd, title: title })
                if (list.Length >= 9)
                    break
            } catch {
                continue
            }
        }
    } catch {
    }
    if (list.Length = 0) {
        msg := "❌ No " appDisplayName " windows found"
        ShowCenteredOverlay_Utils(msg, 2000, BANNER_ACCENT_ERROR)
        return 0
    }
    ; Enrich list using path-first project identification, then sort by canonical project order.
    enriched := []
    for w in list {
        winTitle := w.title ? w.title : ""
        projectIndex := CursorTransfer_GetMatchingProjectIndex(w.hwnd, winTitle)
        projectOrder := Integer(projectIndex > 0 ? projectIndex : 10000 + enriched.Length)
        projName := ""
        if (projectIndex > 0) {
            try {
                project := g_Projects[projectIndex]
                projName := project.name
            }
        }
        displayName := ""
        shortAfterParens := ""
        shortAfterLeadStrip := ""
        cleanProjName := ""
        if (projName != "") {
            shortTitle := winTitle ? winTitle : ""
            cleanProjName := CursorTransfer_StripStaticScriptTokenForDisplay(projName)
            if (shortTitle != "") {
                shortTitle := CursorTransfer_StripDuplicateFilenameInParens(shortTitle)
                shortAfterParens := shortTitle
                shortTitle := CursorTransfer_StripLeadingProjectFromTitle(shortTitle, cleanProjName, projName)
                shortAfterLeadStrip := shortTitle
            }
            ; Label = window title only (no g_Projects name prefix). If stripping the duplicate workspace
            ; segment would leave only "Cursor", keep the fuller line (e.g. "scripts - Cursor").
            if (shortAfterParens != "") {
                cand := shortAfterLeadStrip
                if (cand = "" || CursorTransfer_IsUninformativeCursorTitle(cand))
                    displayName := shortAfterParens
                else
                    displayName := cand
                if (CursorTransfer_IsUninformativeCursorTitle(displayName))
                    displayName := displayName . " · #" . w.hwnd
            } else {
                displayName := winTitle ? winTitle : (appDisplayName . " Window " . w.hwnd)
            }
        } else {
            displayName := winTitle ? winTitle : (appDisplayName . " Window " . w.hwnd)
        }
        displayName := CursorTransfer_StripTrailingCursorAppSuffix(displayName)
        if (displayName = "")
            displayName := "#" . w.hwnd
        enriched.Push({
            hwnd: w.hwnd,
            title: winTitle,
            displayName: displayName,
            projectOrder: projectOrder,
            projectIndex: projectIndex,
            hotkeyChar: ""
        })
    }
    if (enriched.Length = 0) {
        ShowCenteredOverlay_Utils("❌ No mapped Cursor projects found", 2000, BANNER_ACCENT_ERROR)
        return 0
    }
    try {
        CursorTransfer_SortWindowsByProjectOrder(&enriched)
        ; Most important first: put active Cursor window at position 1 if it's in the list.
        try {
            activeHwnd := WinGetID("A")
            if (activeHwnd) {
                loop enriched.Length {
                    if (enriched[A_Index].hwnd = activeHwnd) {
                        if (A_Index > 1) {
                            swap := enriched[1]
                            enriched[1] := enriched[A_Index]
                            enriched[A_Index] := swap
                        }
                        break
                    }
                }
            }
        } catch {
        }
        if (enriched.Length > 9) {
            trimmed := []
            loop 9
                trimmed.Push(enriched[A_Index])
            list := trimmed
        } else {
            list := enriched
        }
        ; Number keys 1-9 (like #!+C / SelectAiModelInHandy modal).
        loop list.Length {
            list[A_Index].hotkeyChar := String(A_Index)
        }
        g_CursorTransferWindowList := list
        g_CursorTransferSelectorResult := ""
        g_CursorTransferSelectorActive := true
        g_CursorTransferSelectorGui := Gui("+AlwaysOnTop -Caption +ToolWindow")
        g_CursorTransferSelectorGui.BackColor := "1E1E2E"
        g_CursorTransferSelectorGui.MarginX := 20
        g_CursorTransferSelectorGui.MarginY := 15
        g_CursorTransferSelectorGui.OnEvent("Escape", CursorTransfer_SelectorEscape)
        g_CursorTransferSelectorGui.SetFont("s14 cCDD6F4 Bold", "Segoe UI")
        transferSelGuiW := 720
        g_CursorTransferSelectorGui.Add("Text", "w" . transferSelGuiW . " Center", "📋 Transfer to " . appDisplayName)
        g_CursorTransferSelectorGui.Add("Text", "w" . transferSelGuiW . " h1 Background45475A")
        g_CursorTransferSelectorGui.SetFont("s12 cCDD6F4", "Segoe UI")
        for w in list {
            g_CursorTransferSelectorGui.Add("Text", "w" . transferSelGuiW, "[" . w.hotkeyChar . "] " . w.displayName)
        }
        g_CursorTransferSelectorGui.Add("Text", "w" . transferSelGuiW . " h1 Background45475A y+10")
        g_CursorTransferSelectorGui.SetFont("s9 c6C7086", "Segoe UI")
        g_CursorTransferSelectorGui.Add("Text", "w" . transferSelGuiW . " Center", "Press 1-9 | N or Esc to cancel")
        g_CursorTransferSelectorGui.Show("AutoSize Hide")
        g_CursorTransferSelectorGui.GetPos(&gx, &gy, &gw, &gh)
        ; Center on centerOnHwnd's monitor when provided; otherwise foreground window's monitor (dictation / transfer flow).
        monitorIndex := 1
        if (centerOnHwnd && WinExist("ahk_id " centerOnHwnd)) {
            rect := Buffer(16, 0)
            if (DllCall("GetWindowRect", "ptr", centerOnHwnd, "ptr", rect)) {
                cx := NumGet(rect, 0, "int") + (NumGet(rect, 8, "int") - NumGet(rect, 0, "int")) // 2
                cy := NumGet(rect, 4, "int") + (NumGet(rect, 12, "int") - NumGet(rect, 4, "int")) // 2
                monitorCount := MonitorGetCount()
                loop monitorCount {
                    idx := A_Index
                    MonitorGet(idx, &ml, &mt, &mr, &mb)
                    if (cx >= ml && cx <= mr && cy >= mt && cy <= mb) {
                        monitorIndex := idx
                        break
                    }
                }
            }
        } else {
            monitorIndex := GetMonitorIndexForForeground_StandardBar()
        }
        MonitorGetWorkArea(monitorIndex, &ml, &mt, &mr, &mb)
        mw := mr - ml
        mh := mb - mt
        cx := ml + (mw - gw) // 2
        cy := mt + (mh - gh) // 2
        g_CursorTransferSelectorGui.Show("x" . cx . " y" . cy . " NA")
    } catch as err {
        ShowCenteredOverlay_Utils("❌ Selector error", 2000, BANNER_ACCENT_ERROR)
        return 0
    }
    ; Wildcard prefix so hotkeys fire even when modifier held (e.g. C still down when modal opens).
    g_CursorTransferHotkeyHandlers := []
    loop list.Length {
        i := A_Index
        keyChar := list[i].hotkeyChar
        if (keyChar = "")
            continue
        try {
            fn := CursorTransfer_SelectorHandleKey.Bind(i)
            hotkeyKey := "*" . keyChar
            Hotkey(hotkeyKey, fn, "On")
            g_CursorTransferHotkeyHandlers.Push({ key: hotkeyKey, callback: fn })
        } catch {
        }
    }
    try {
        Hotkey("*Escape", CursorTransfer_SelectorEscape, "On")
        g_CursorTransferHotkeyHandlers.Push({ key: "*Escape", callback: CursorTransfer_SelectorEscape })
        Hotkey("*N", CursorTransfer_SelectorEscape, "On")
        g_CursorTransferHotkeyHandlers.Push({ key: "*N", callback: CursorTransfer_SelectorEscape })
        Hotkey("*n", CursorTransfer_SelectorEscape, "On")
        g_CursorTransferHotkeyHandlers.Push({ key: "*n", callback: CursorTransfer_SelectorEscape })
    } catch {
    }
    start := A_TickCount
    timeoutMs := 30000
    lastCursorTransferMonitorIdx := monitorIndex
    keyWasDownByIndex := []
    loop list.Length
        keyWasDownByIndex.Push(false)
    cancelWasDown := false
    try {
        while (g_CursorTransferSelectorResult = "") {
            if ((A_TickCount - start) >= timeoutMs)
                break
            Sleep 50

            ; Fallback key polling (edge-triggered): survives global Hotkey("1") conflicts from other scripts.
            loop list.Length {
                if (g_CursorTransferSelectorResult != "")
                    break
                keyChar := list[A_Index].hotkeyChar
                if (keyChar = "")
                    continue
                isDown := false
                try isDown := GetKeyState(keyChar, "P") || GetKeyState("Numpad" . keyChar, "P")
                if (isDown && !keyWasDownByIndex[A_Index])
                    CursorTransfer_SelectorHandleKey(A_Index)
                keyWasDownByIndex[A_Index] := isDown
            }

            isCancelDown := false
            try isCancelDown := GetKeyState("Escape", "P") || GetKeyState("n", "P") || GetKeyState("N", "P")
            if (isCancelDown && !cancelWasDown && g_CursorTransferSelectorResult = "")
                CursorTransfer_SelectorEscape()
            cancelWasDown := isCancelDown

            if (IsObject(g_CursorTransferSelectorGui)) {
                curIdx := GetMonitorIndexForForeground_StandardBar()
                if (curIdx != lastCursorTransferMonitorIdx) {
                    lastCursorTransferMonitorIdx := curIdx
                    MonitorGetWorkArea(curIdx, &ml, &mt, &mr, &mb)
                    mw := mr - ml
                    mh := mb - mt
                    try {
                        g_CursorTransferSelectorGui.GetPos(&gxOld, &gyOld, &gw, &gh)
                        cx := ml + (mw - gw) // 2
                        cy := mt + (mh - gh) // 2
                        g_CursorTransferSelectorGui.Show("x" . cx . " y" . cy . " NA")
                    } catch {
                    }
                }
            }
        }
    } catch as loopErr {
        ; ignore loop exceptions
    }
    durationMs := A_TickCount - start
    result := (g_CursorTransferSelectorResult = "") ? 0 : Integer(g_CursorTransferSelectorResult)
    CursorTransfer_SelectorClose()
    return result
}
