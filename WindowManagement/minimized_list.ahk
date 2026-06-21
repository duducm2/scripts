; =============================================================================
; WindowManagement module: minimized_list.ahk
; Minimized/hidden background window list GUI
; Extracted verbatim from WindowManagement.ahk; loaded via #include into the
; WindowManagement.ahk process, which remains the entry point / source of truth.
; =============================================================================

WM_BackgroundIsEligibleForExcludePicker(hwnd, foreHwnd) {
    if (!hwnd || hwnd = foreHwnd)
        return false
    try {
        exStyle := DllCall("GetWindowLongPtr", "ptr", hwnd, "int", -20, "ptr")
        if (exStyle & 0x00000080)
            return false
        class := WinGetClass(hwnd)
        if (WM_IsDesktopOrTaskbarClass(class))
            return false
        if (WinGetTitle(hwnd) = "")
            return false
        if (WM_IsExcludedIndicatorWindow(hwnd))
            return false
    } catch {
        return false
    }
    return true
}

WM_BackgroundAddRowForExcludePicker(&rows, &seen, hwnd, foreHwnd) {
    if (seen.Has(hwnd) || !WM_BackgroundIsEligibleForExcludePicker(hwnd, foreHwnd))
        return
    try {
        rows.Push({
            hwnd: hwnd,
            title: WinGetTitle(hwnd),
            exe: WinGetProcessName("ahk_id " hwnd)
        })
        seen[hwnd] := true
    } catch {
    }
}

WM_CollectMinimizedWindowsForExcludePicker() {
    global g_WM_MinimizedListCollectForeHwnd
    WM_BackgroundTitleExcludes_Ensure()
    collectForeHwnd := g_WM_MinimizedListCollectForeHwnd
    if (!collectForeHwnd) {
        try collectForeHwnd := WinGetID("A")
    }
    return WM_BackgroundFilterRowsByTitleExcludes(WM_CollectBackgroundWindows(collectForeHwnd))
}

WM_BackgroundIsEligibleWindow(hwnd, foreHwnd) {
    if (!hwnd || hwnd = foreHwnd)
        return false
    try {
        exStyle := DllCall("GetWindowLongPtr", "ptr", hwnd, "int", -20, "ptr")
        if (exStyle & 0x00000080)
            return false
        class := WinGetClass(hwnd)
        if (WM_IsDesktopOrTaskbarClass(class))
            return false
        title := WinGetTitle(hwnd)
        if (title = "")
            return false
        exe := ""
        try exe := WinGetProcessName("ahk_id " hwnd)
        if (WM_BackgroundTitleIsExcluded(title, exe))
            return false
        if (WM_IsExcludedIndicatorWindow(hwnd))
            return false
    } catch {
        return false
    }
    return true
}

WM_BackgroundAddRow(&rows, &seen, hwnd, foreHwnd) {
    if (seen.Has(hwnd))
        return false
    title := ""
    exe := ""
    try title := WinGetTitle(hwnd)
    catch
        return false
    if (title = "")
        return false
    exe := ""
    try exe := WinGetProcessName("ahk_id " hwnd)
    catch
        exe := ""
    if (WM_BackgroundTitleIsExcluded(title, exe))
        return false
    rows.Push({ hwnd: hwnd, title: title, exe: exe })
    seen[hwnd] := true
    return true
}

; Hidden windows: not visible on any monitor (z-order covered and/or taskbar-minimized); excludes foreground hwnd.
WM_CollectBackgroundWindows(foreHwndOverride := 0) {
    global g_WM_BackgroundTitleExcludes
    WM_BackgroundTitleExcludes_Init()
    foreHwnd := foreHwndOverride
    foreClass := ""
    foreTitle := ""
    foreExe := ""
    try {
        if (!foreHwnd)
            foreHwnd := WinGetID("A")
        foreClass := WinGetClass(foreHwnd)
        foreTitle := WM_TruncateTitleForList(WinGetTitle(foreHwnd), 40)
        foreExe := WinGetProcessName(foreHwnd)
    } catch {
    }
    rows := []
    seen := Map()
    hiddenHwnds := WM_BackgroundEnumerateHiddenHwnds()
    rejectStats := Map()
    added := 0
    for hwnd in hiddenHwnds {
        try {
            reason := WM_BackgroundExplainReject(hwnd, foreHwnd)
            if (reason != "") {
                rejectStats[reason] := rejectStats.Has(reason) ? rejectStats[reason] + 1 : 1
                continue
            }
            if (WM_BackgroundAddRow(&rows, &seen, hwnd, foreHwnd))
                added++
            else
                rejectStats["row_build_failed"] := rejectStats.Has("row_build_failed") ? rejectStats["row_build_failed"
                    ] + 1 : 1
        } catch {
            rejectStats["addrow_exception"] := rejectStats.Has("addrow_exception") ? rejectStats["addrow_exception"] +
                1 : 1
        }
    }
    WM_SortBackgroundRows(&rows)
    rowsBeforeTitleFilter := rows.Length
    rows := WM_BackgroundFilterRowsByTitleExcludes(rows)
    if (rowsBeforeTitleFilter > rows.Length)
        rejectStats["title_excluded_postfilter"] := rowsBeforeTitleFilter - rows.Length
    global g_WM_LastEnumerateStats, g_WM_LastBackgroundCollectStats
    g_WM_LastBackgroundCollectStats := Map(
        "visibleOnMonitors", g_WM_LastEnumerateStats.Get("visibleOnMonitors", 0),
        "hiddenCandidates", g_WM_LastEnumerateStats.Get("hiddenCandidates", 0),
        "deduped", hiddenHwnds.Length,
        "rows", rows.Length,
        "added", added,
        "foreExe", foreExe,
        "foreClass", foreClass,
        "rejects", rejectStats.Clone())
    return rows
}

WM_CheckMinimizedListCloseRequest() {
    global g_WM_MinimizedListActive, g_WM_MinimizedListCloseRequestFile
    if (!g_WM_MinimizedListActive)
        return
    if (FileExist(g_WM_MinimizedListCloseRequestFile)) {
        try FileDelete(g_WM_MinimizedListCloseRequestFile)
        catch {
        }
        WM_MinimizedList_Cancel()
    }
}

WM_MinimizedList_ModifiersDown() {
    try {
        return GetKeyState("LWin", "P") || GetKeyState("RWin", "P") || GetKeyState("Ctrl", "P") || GetKeyState("Alt",
            "P") ||
        GetKeyState("Shift", "P")
    } catch {
        return false
    }
}

WM_MinimizedList_ShouldCaptureKey() {
    global g_WM_MinimizedListActive, g_WM_MinimizedListExcludePickerActive, g_WM_MinimizedListRefreshing
    if (!g_WM_MinimizedListActive || g_WM_MinimizedListRefreshing || g_WM_MinimizedListExcludePickerActive)
        return false
    return !WM_MinimizedList_ModifiersDown()
}

WM_MinimizedList_ShouldCapturePickerKey() {
    global g_WM_MinimizedListActive, g_WM_MinimizedListExcludePickerActive, g_WM_MinimizedListRefreshing
    if (!g_WM_MinimizedListActive || g_WM_MinimizedListRefreshing || !g_WM_MinimizedListExcludePickerActive)
        return false
    return !WM_MinimizedList_ModifiersDown()
}

; AHK v2: do not use c >= "0" on slot letters a–z — throws "Expected a Number but got a String".
WM_IsDigitSlotChar(c) {
    if (StrLen(c) != 1)
        return false
    o := Ord(c)
    return o >= Ord("0") && o <= Ord("9")
}

; [A] and [C] are command keys — never assign hidden-window slots to a/c (same physical key for exclude/close).
WM_MinimizedList_IsReservedSlotChar(char) {
    c := StrLower(char)
    return c = "a" || c = "c"
}

WM_MinimizedList_WindowSlotChars() {
    global g_WM_MinimizedCharSequence
    chars := []
    for ch in g_WM_MinimizedCharSequence {
        if (!WM_MinimizedList_IsReservedSlotChar(ch))
            chars.Push(ch)
    }
    return chars
}

; Hotkey + keys poll both fire on one press — consume duplicate within debounce window.
WM_MinimizedList_TryConsumeCharAction(char) {
    global g_WM_MinimizedListCharActionLock, g_WM_MinimizedListLastCharActionTick, g_WM_MinimizedListLastCharActionKey
    if (g_WM_MinimizedListCharActionLock)
        return false
    key := StrLower(char)
    if (g_WM_MinimizedListLastCharActionKey = key && A_TickCount - g_WM_MinimizedListLastCharActionTick < 400)
        return false
    g_WM_MinimizedListCharActionLock := true
    g_WM_MinimizedListLastCharActionKey := key
    g_WM_MinimizedListLastCharActionTick := A_TickCount
    return true
}

WM_MinimizedList_ReleaseCharActionLock() {
    global g_WM_MinimizedListCharActionLock
    g_WM_MinimizedListCharActionLock := false
}

WM_MinimizedList_KeyDown(keyName) {
    try {
        if (WM_IsDigitSlotChar(keyName))
            return GetKeyState(keyName, "P") || GetKeyState("Numpad" . keyName, "P")
        return GetKeyState(keyName, "P")
    } catch {
        return false
    }
}

WM_MinimizedList_StopKeysPoll() {
    global g_WM_MinimizedKeysPollTimer, g_WM_MinimizedKeysPollPrev, g_WM_MinimizedKeysPollCallbacks
    try SetTimer(g_WM_MinimizedKeysPollTimer, 0)
    catch {
    }
    g_WM_MinimizedKeysPollTimer := ""
    g_WM_MinimizedKeysPollPrev := Map()
    g_WM_MinimizedKeysPollCallbacks := Map()
}

; Edge-triggered poll: survives global Hotkey("1") conflicts (see StandardLoadingBar_KeysSelectionPoll in Utils.ahk).
WM_MinimizedList_KeysPoll() {
    global g_WM_MinimizedListActive, g_WM_MinimizedListExcludePickerActive, g_WM_MinimizedListRefreshing,
        g_WM_MinimizedKeysPollCallbacks, g_WM_MinimizedKeysPollPrev
    if (!g_WM_MinimizedListActive) {
        WM_MinimizedList_StopKeysPoll()
        return
    }
    if (g_WM_MinimizedListRefreshing)
        return
    if (g_WM_MinimizedListExcludePickerActive) {
        if (!WM_MinimizedList_ShouldCapturePickerKey())
            return
    } else if (!WM_MinimizedList_ShouldCaptureKey())
        return
    for keyName, cb in g_WM_MinimizedKeysPollCallbacks {
        if (!cb)
            continue
        isDown := WM_MinimizedList_KeyDown(keyName)
        wasDown := g_WM_MinimizedKeysPollPrev.Has(keyName) ? g_WM_MinimizedKeysPollPrev[keyName] : false
        g_WM_MinimizedKeysPollPrev[keyName] := isDown
        if (isDown && !wasDown) {
            try cb.Call()
            catch {
            }
        }
    }
}

WM_MinimizedList_StartKeysPoll(windows, forExcludePicker := false) {
    global g_WM_MinimizedKeysPollCallbacks, g_WM_MinimizedKeysPollPrev, g_WM_MinimizedKeysPollTimer
    WM_MinimizedList_StopKeysPoll()
    g_WM_MinimizedKeysPollCallbacks := Map()
    g_WM_MinimizedKeysPollPrev := Map()
    for w in windows {
        slotChar := w.char
        g_WM_MinimizedKeysPollCallbacks[slotChar] := forExcludePicker ?
            HandleMinimizedListExcludePickerByChar.Bind(slotChar) : HandleMinimizedListByChar.Bind(slotChar)
        g_WM_MinimizedKeysPollPrev[slotChar] := WM_MinimizedList_KeyDown(slotChar)
    }
    if (g_WM_MinimizedKeysPollCallbacks.Count > 0)
        g_WM_MinimizedKeysPollTimer := SetTimer(WM_MinimizedList_KeysPoll, 50)
}

WM_MinimizedList_RegisterHotkey(hk, handler) {
    global g_WM_MinimizedHotkeyHandlers
    try {
        #InputLevel 10
        Hotkey(hk, handler, "On")
        #InputLevel 0
        g_WM_MinimizedHotkeyHandlers.Push({ hk: hk, handler: handler })
    } catch {
    }
}

WM_MinimizedList_UnbindHotkeys() {
    global g_WM_MinimizedHotkeyHandlers
    WM_MinimizedList_StopKeysPoll()
    try HotIf()
    catch {
    }
    for entry in g_WM_MinimizedHotkeyHandlers {
        try Hotkey(entry.hk, "Off")
        catch {
        }
    }
    g_WM_MinimizedHotkeyHandlers := []
    try HotIf()
    catch {
    }
}

WM_MinimizedList_BindHotkeys(windows) {
    WM_MinimizedList_UnbindHotkeys()
    if (windows.Length = 0)
        return
    try HotIf (*) => WM_MinimizedList_ShouldCaptureKey()
    catch {
    }
    for w in windows {
        slotChar := w.char
        WM_MinimizedList_RegisterHotkey("$*" . slotChar, HandleMinimizedListByChar.Bind(slotChar))
        if (WM_IsDigitSlotChar(slotChar))
            WM_MinimizedList_RegisterHotkey("$*Numpad" . slotChar, HandleMinimizedListByChar.Bind(slotChar))
    }
    WM_MinimizedList_RegisterHotkey("$*A", HandleMinimizedListAddExcludeTrigger)
    WM_MinimizedList_RegisterHotkey("$*C", HandleMinimizedListCloseModeArm)
    try HotIf()
    catch {
    }
    WM_MinimizedList_StartKeysPoll(windows, false)
}

WM_MinimizedList_BindPickerHotkeys(pickerWindows) {
    WM_MinimizedList_UnbindHotkeys()
    if (pickerWindows.Length = 0)
        return
    try HotIf (*) => WM_MinimizedList_ShouldCapturePickerKey()
    catch {
    }
    for w in pickerWindows {
        slotChar := w.char
        WM_MinimizedList_RegisterHotkey("$*" . slotChar, HandleMinimizedListExcludePickerByChar.Bind(slotChar))
        if (WM_IsDigitSlotChar(slotChar))
            WM_MinimizedList_RegisterHotkey("$*Numpad" . slotChar, HandleMinimizedListExcludePickerByChar.Bind(slotChar
            ))
    }
    try HotIf()
    catch {
    }
    WM_MinimizedList_StartKeysPoll(pickerWindows, true)
}

WM_MinimizedList_EscapePoll() {
    global g_WM_MinimizedListActive, g_WM_MinimizedListEscPollPrev
    if (!g_WM_MinimizedListActive) {
        try SetTimer(WM_MinimizedList_EscapePoll, 0)
        catch {
        }
        return
    }
    if (WM_MinimizedList_ModifiersDown())
        return
    escDown := GetKeyState("Escape", "P") || (DllCall("user32\GetAsyncKeyState", "int", 0x1B) & 0x8000)
    if (escDown) {
        if (!g_WM_MinimizedListEscPollPrev) {
            g_WM_MinimizedListEscPollPrev := true
            WM_MinimizedList_Cancel()
        }
    } else
        g_WM_MinimizedListEscPollPrev := false
}

HandleMinimizedListEscape(*) {
    global g_WM_MinimizedListActive
    if (g_WM_MinimizedListActive)
        WM_MinimizedList_Cancel()
}

WM_MinimizedList_BindEscape() {
    global g_WM_MinimizedListEscPollPrev, g_OnEscapePressed, g_WM_MinimizedListOpenFile,
        g_WM_MinimizedListCloseCheckTimer
    try SetTimer(WM_MinimizedList_EscapePoll, 0)
    catch {
    }
    g_OnEscapePressed := HandleMinimizedListEscape
    Utils_EnsureGlobalEscapeHotkey()
    try HotIf()
    catch {
    }
    try {
        #InputLevel 10
        Hotkey("$*Escape", WM_MinimizedList_Cancel, "On")
        #InputLevel 0
    } catch {
    }
    g_WM_MinimizedListEscPollPrev := false
    SetTimer(WM_MinimizedList_EscapePoll, 50)
    try {
        DirCreate(A_ScriptDir "\.cursor")
        try FileDelete(g_WM_MinimizedListOpenFile)
        FileAppend "", g_WM_MinimizedListOpenFile
    } catch {
    }
    try SetTimer(WM_CheckMinimizedListCloseRequest, 0)
    catch {
    }
    g_WM_MinimizedListCloseCheckTimer := SetTimer(WM_CheckMinimizedListCloseRequest, 120)
}

WM_MinimizedList_UnbindEscape() {
    global g_WM_MinimizedListEscPollPrev, g_OnEscapePressed, g_WM_MinimizedListOpenFile,
        g_WM_MinimizedListCloseRequestFile,
        g_WM_MinimizedListCloseCheckTimer
    try SetTimer(WM_MinimizedList_EscapePoll, 0)
    catch {
    }
    try SetTimer(WM_CheckMinimizedListCloseRequest, 0)
    catch {
    }
    g_WM_MinimizedListCloseCheckTimer := ""
    g_WM_MinimizedListEscPollPrev := false
    try {
        #InputLevel 10
        Hotkey("$*Escape", WM_MinimizedList_Cancel, "Off")
        #InputLevel 0
    } catch {
    }
    try FileDelete(g_WM_MinimizedListOpenFile)
    catch {
    }
    try FileDelete(g_WM_MinimizedListCloseRequestFile)
    catch {
    }
    if (g_OnEscapePressed = HandleMinimizedListEscape)
        g_OnEscapePressed := ""
    Utils_EnsureGlobalEscapeHotkey()
}

WM_CenterGuiOnActiveMonitor(gui) {
    WM_MinimizedList_RepositionToActiveMonitor(0, gui)
}

WM_MinimizedList_StopActiveMonitorTracking() {
    global g_WM_MinimizedListTrackTimer
    try SetTimer(WM_MinimizedList_TrackActiveMonitorTick, 0)
    catch {
    }
    g_WM_MinimizedListTrackTimer := ""
}

WM_MinimizedList_GuiHasWindow(gui := unset) {
    if (!IsSet(gui))
        gui := g_WM_MinimizedListGui
    if (!IsObject(gui))
        return false
    try return !!gui.Hwnd
    catch
        return false
}

WM_MinimizedList_ActivateGui(gui) {
    if (!WM_MinimizedList_GuiHasWindow(gui))
        return
    try WinActivate("ahk_id " gui.Hwnd)
    catch {
    }
    try DllCall("SetForegroundWindow", "ptr", gui.Hwnd)
    catch {
    }
}

WM_MinimizedList_ShowAt(gui, cx, cy) {
    try gui.Show("x" . cx . " y" . cy)
    catch {
        return
    }
    WM_MinimizedList_ActivateGui(gui)
}

; Reposition modal to center of foreground monitor (parity with StandardLoadingBar trackActiveMonitor / StudyTopicSelector).
WM_MinimizedList_RepositionToActiveMonitor(forMonitorIdx := 0, gui := unset) {
    global g_WM_MinimizedListGui
    if (!IsSet(gui))
        gui := g_WM_MinimizedListGui
    if (!WM_MinimizedList_GuiHasWindow(gui))
        return
    idx := forMonitorIdx
    if (idx < 1 || idx > MonitorGetCount())
        idx := GetMonitorIndexForForeground_StandardBar()
    MonitorGetWorkArea(idx, &ml, &mt, &mr, &mb)
    monitorWidth := mr - ml
    monitorHeight := mb - mt
    try {
        gui.Show("AutoSize Hide")
        gui.GetPos(, , &gw, &gh)
    } catch {
        return
    }
    marginX := 20
    marginY := 20
    cx := ml + (monitorWidth - gw) // 2
    cy := mt + (monitorHeight - gh) // 2
    if (cx < ml + marginX)
        cx := ml + marginX
    if (cy < mt + marginY)
        cy := mt + marginY
    if (cx + gw > mr - marginX)
        cx := mr - gw - marginX
    if (cy + gh > mb - marginY)
        cy := mb - gh - marginY
    WM_MinimizedList_ShowAt(gui, cx, cy)
}

WM_MinimizedList_TrackActiveMonitorTick() {
    global g_WM_MinimizedListActive, g_WM_MinimizedListGui, g_WM_MinimizedListLastForegroundMonitorIdx
    if (!g_WM_MinimizedListActive || !WM_MinimizedList_GuiHasWindow()) {
        WM_MinimizedList_StopActiveMonitorTracking()
        return
    }
    newIdx := GetMonitorIndexForForeground_StandardBar()
    if (newIdx != g_WM_MinimizedListLastForegroundMonitorIdx) {
        g_WM_MinimizedListLastForegroundMonitorIdx := newIdx
        WM_MinimizedList_RepositionToActiveMonitor(newIdx)
    }
}

WM_MinimizedList_StartActiveMonitorTracking() {
    global g_WM_MinimizedListTrackTimer, g_WM_MinimizedListLastForegroundMonitorIdx
    WM_MinimizedList_StopActiveMonitorTracking()
    g_WM_MinimizedListLastForegroundMonitorIdx := GetMonitorIndexForForeground_StandardBar()
    g_WM_MinimizedListTrackTimer := SetTimer(WM_MinimizedList_TrackActiveMonitorTick, 115)
}

WM_MinimizedList_KeyLabel(char) {
    return (char = "0") ? "10" : char
}

WM_MinimizedList_AssignKeys(rows) {
    global g_WM_MinimizedKeyMap
    slotChars := WM_MinimizedList_WindowSlotChars()
    g_WM_MinimizedKeyMap := Map()
    windows := []
    limit := Min(rows.Length, slotChars.Length)
    loop limit {
        ch := slotChars[A_Index]
        row := rows[A_Index]
        g_WM_MinimizedKeyMap[ch] := row.hwnd
        windows.Push({ hwnd: row.hwnd, title: row.title, char: ch, label: WM_MinimizedList_KeyLabel(ch) })
    }
    return windows
}

WM_MinimizedList_WaitForHwndClosed(hwnd, timeoutMs := 2000) {
    if (!hwnd)
        return true
    deadline := A_TickCount + timeoutMs
    loop {
        if !WinExist("ahk_id " hwnd)
            break
        if (A_TickCount >= deadline)
            return false
        Sleep 50
    }
    Sleep 50
    return !WinExist("ahk_id " hwnd)
}

WM_MinimizedList_FilterExcludedHwnd(rows, excludeHwnd) {
    if (!excludeHwnd)
        return rows
    filtered := []
    for row in rows {
        if (row.hwnd != excludeHwnd)
            filtered.Push(row)
    }
    return filtered
}

WM_MinimizedList_CloseHwnd(hwnd) {
    if (!hwnd)
        return true
    try {
        WinClose "ahk_id " hwnd
        if !WinWaitClose("ahk_id " hwnd, , 0.2) {
            try WinShow "ahk_id " hwnd
            WinActivate "ahk_id " hwnd
            WinWaitActive "ahk_id " hwnd, , 1.2
            WinClose "ahk_id " hwnd
            if !WinWaitClose("ahk_id " hwnd, , 0.35) {
                PostMessage 0x0010, 0, 0, , "ahk_id " hwnd
                WinWaitClose "ahk_id " hwnd, , 1.5
            }
        }
    } catch {
    }
    return WM_MinimizedList_WaitForHwndClosed(hwnd)
}

WM_MinimizedList_OpenHwnd(hwnd) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    try {
        if (WM_WindowIsTaskbarMinimized(hwnd))
            WinRestore "ahk_id " hwnd
        WinShow "ahk_id " hwnd
        WinActivate "ahk_id " hwnd
        WinWaitActive "ahk_id " hwnd, , 1.5
    } catch {
        return false
    }
    Sleep 50
    return true
}

HandleMinimizedListCloseModeArm(*) {
    global g_WM_MinimizedListActive, g_WM_MinimizedListRefreshing, g_WM_MinimizedListExcludePickerActive,
        g_WM_MinimizedListCloseModeArmed, g_WM_MinimizedListLastCloseArmTick
    if (!g_WM_MinimizedListActive || g_WM_MinimizedListRefreshing || g_WM_MinimizedListExcludePickerActive)
        return
    if (A_TickCount - g_WM_MinimizedListLastCloseArmTick < 250)
        return
    g_WM_MinimizedListLastCloseArmTick := A_TickCount
    g_WM_MinimizedListCloseModeArmed := true
    WM_MinimizedList_RepaintMainList()
}

HandleMinimizedListByChar(char, *) {
    global g_WM_MinimizedListActive, g_WM_MinimizedKeyMap, g_WM_MinimizedListRefreshing,
        g_WM_MinimizedListExcludePickerActive,
        g_WM_MinimizedListCloseModeArmed
    if (!g_WM_MinimizedListActive || g_WM_MinimizedListRefreshing || g_WM_MinimizedListExcludePickerActive)
        return
    if (WM_MinimizedList_IsReservedSlotChar(char))
        return
    if (!WM_MinimizedList_TryConsumeCharAction(char))
        return
    try {
        hwnd := g_WM_MinimizedKeyMap.Get(char, "")
        if (hwnd = "")
            hwnd := g_WM_MinimizedKeyMap.Get(StrLower(char), "")
        if (!hwnd)
            return
        if (g_WM_MinimizedListCloseModeArmed) {
            g_WM_MinimizedListCloseModeArmed := false
            WM_MinimizedList_CloseHwnd(hwnd)
            WM_MinimizedList_Refresh(hwnd)
            return
        }
        WM_MinimizedList_UnbindHotkeys()
        WM_MinimizedList_OpenHwnd(hwnd)
        WM_MinimizedList_Cleanup()
    } finally {
        WM_MinimizedList_ReleaseCharActionLock()
    }
}

WM_MinimizedList_BuildDisplayText(rows, windows) {
    global g_WM_MinimizedListCloseModeArmed
    slotCount := WM_MinimizedList_WindowSlotChars().Length
    displayText := "=== HIDDEN WINDOWS (NOT VISIBLE ON MONITORS) ===`n`n"
    for w in windows
        displayText .= "[" . w.label . "] " . w.title . "`n"
    if (rows.Length > slotCount)
        displayText .= "`n(" . (rows.Length - slotCount) . " more — close some and reopen)`n"
    if (g_WM_MinimizedListCloseModeArmed)
        displayText .= "`n>>> Press a slot key to CLOSE <<<`n"
    else
        displayText .= "`n>>> [C] then slot key = CLOSE  |  slot key alone = OPEN <<<`n"
    displayText .= "`n[C] Arm close  [A] Add to exclude list  [ESC] Cancel"
    return displayText
}

WM_MinimizedList_RepaintMainList() {
    global g_WM_MinimizedListRows
    if (g_WM_MinimizedListRows.Length = 0)
        return
    windows := WM_MinimizedList_AssignKeys(g_WM_MinimizedListRows)
    displayText := WM_MinimizedList_BuildDisplayText(g_WM_MinimizedListRows, windows)
    WM_MinimizedList_RebuildListGui(displayText)
    WM_MinimizedList_BindHotkeys(windows)
}

WM_MinimizedList_BuildExcludePickerDisplayText(rows, pickerWindows) {
    displayText := "=== ADD TO EXCLUDE LIST ===`n`n"
    for w in pickerWindows
        displayText .= "[" . w.label . "] " . w.title . "`n"
    if (rows.Length > pickerWindows.Length)
        displayText .= "`n(" . (rows.Length - pickerWindows.Length) . " more — not shown)`n"
    displayText .= "`n[ESC] Cancel — back to list"
    return displayText
}

WM_MinimizedList_RebuildListGui(displayText) {
    global g_WM_MinimizedListGui, g_WM_MinimizedListActive
    WM_MinimizedList_StopActiveMonitorTracking()
    if (WM_MinimizedList_GuiHasWindow()) {
        try g_WM_MinimizedListGui.Destroy()
        catch {
        }
        g_WM_MinimizedListGui := false
    }
    fontSize := 11
    baseWidth := 480
    g_WM_MinimizedListGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Owner")
    g_WM_MinimizedListGui.BackColor := "1E1E2E"
    g_WM_MinimizedListGui.MarginX := 15
    g_WM_MinimizedListGui.MarginY := 10
    g_WM_MinimizedListGui.SetFont("s" . fontSize . " cCDD6F4", "Segoe UI")
    g_WM_MinimizedListGui.Add("Text", "w" . (baseWidth - 30), displayText)
    g_WM_MinimizedListGui.OnEvent("Escape", WM_MinimizedList_Cancel)
    WM_MinimizedList_RepositionToActiveMonitor(0, g_WM_MinimizedListGui)
    WM_MinimizedList_ActivateGui(g_WM_MinimizedListGui)
    if (g_WM_MinimizedListActive)
        WM_MinimizedList_StartActiveMonitorTracking()
}

HandleMinimizedListExcludePickerByChar(char, *) {
    global g_WM_MinimizedListExcludePickerActive, g_WM_MinimizedListExcludePickerMap, g_WM_MinimizedListRefreshing
    if (!g_WM_MinimizedListExcludePickerActive || g_WM_MinimizedListRefreshing)
        return
    row := g_WM_MinimizedListExcludePickerMap.Get(char, "")
    if (!IsObject(row) || !row.HasProp("title"))
        return
    if (!WM_BackgroundTitleExcludes_PersistAppend(row.title, row.exe))
        return
    g_WM_MinimizedListExcludePickerActive := false
    g_WM_MinimizedListExcludePickerMap := Map()
    g_WM_MinimizedListExcludePickerRows := []
    WM_MinimizedList_Refresh(0)
}

HandleMinimizedListAddExcludeTrigger(*) {
    global g_WM_MinimizedListActive, g_WM_MinimizedListRefreshing, g_WM_MinimizedListExcludePickerActive
    if (!g_WM_MinimizedListActive || g_WM_MinimizedListRefreshing || g_WM_MinimizedListExcludePickerActive)
        return
    WM_MinimizedList_ShowExcludePicker()
}

WM_MinimizedList_ShowExcludePicker() {
    global g_WM_MinimizedListExcludePickerActive, g_WM_MinimizedListExcludePickerRows,
        g_WM_MinimizedListExcludePickerMap,
        g_WM_MinimizedListExcludePickerDigitSequence, g_WM_MinimizedListCloseModeArmed,
        g_WM_MinimizedListCollectForeHwnd
    g_WM_MinimizedListCloseModeArmed := false
    WM_MinimizedList_UnbindHotkeys()
    WM_BackgroundTitleExcludes_Init()
    allRows := WM_CollectBackgroundWindows(g_WM_MinimizedListCollectForeHwnd)
    rows := WM_BackgroundFilterRowsByTitleExcludes(allRows)
    if (rows.Length = 0) {
        ShowCenteredOverlay_Utils("ℹ️ No hidden windows to add to exclude list", 2500, BANNER_ACCENT_INFO)
        WM_MinimizedList_Refresh(0)
        return
    }
    g_WM_MinimizedListExcludePickerActive := true
    g_WM_MinimizedListExcludePickerRows := rows
    g_WM_MinimizedListExcludePickerMap := Map()
    pickerWindows := []
    limit := Min(rows.Length, g_WM_MinimizedListExcludePickerDigitSequence.Length)
    loop limit {
        ch := g_WM_MinimizedListExcludePickerDigitSequence[A_Index]
        row := rows[A_Index]
        g_WM_MinimizedListExcludePickerMap[ch] := row
        pickerWindows.Push({ char: ch, title: row.title, label: WM_MinimizedList_KeyLabel(ch) })
    }
    displayText := WM_MinimizedList_BuildExcludePickerDisplayText(rows, pickerWindows)
    WM_MinimizedList_RebuildListGui(displayText)
    WM_MinimizedList_BindPickerHotkeys(pickerWindows)
}

WM_MinimizedList_CancelExcludePicker() {
    global g_WM_MinimizedListExcludePickerActive, g_WM_MinimizedListExcludePickerRows,
        g_WM_MinimizedListExcludePickerMap,
        g_WM_MinimizedListCloseModeArmed
    g_WM_MinimizedListCloseModeArmed := false
    g_WM_MinimizedListExcludePickerActive := false
    g_WM_MinimizedListExcludePickerRows := []
    g_WM_MinimizedListExcludePickerMap := Map()
    WM_MinimizedList_Refresh(0)
}

WM_MinimizedList_Cleanup() {
    global g_WM_MinimizedListGui, g_WM_MinimizedListActive, g_WM_MinimizedListRows, g_WM_MinimizedKeyMap,
        g_WM_MinimizedListRefreshing, g_WM_MinimizedListExcludePickerActive,
        g_WM_MinimizedListExcludePickerRows, g_WM_MinimizedListExcludePickerMap, g_WM_MinimizedListCloseModeArmed
    if (!g_WM_MinimizedListActive && !WM_MinimizedList_GuiHasWindow())
        return
    g_WM_MinimizedListActive := false
    g_WM_MinimizedListRefreshing := false
    g_WM_MinimizedListCloseModeArmed := false
    g_WM_MinimizedListExcludePickerActive := false
    g_WM_MinimizedListExcludePickerRows := []
    g_WM_MinimizedListExcludePickerMap := Map()
    WM_MinimizedList_StopActiveMonitorTracking()
    g_WM_MinimizedListRows := []
    g_WM_MinimizedKeyMap := Map()
    WM_MinimizedList_UnbindHotkeys()
    WM_MinimizedList_UnbindEscape()
    try Utils_EnsureGlobalEscapeHotkey()
    if (WM_MinimizedList_GuiHasWindow()) {
        try g_WM_MinimizedListGui.Destroy()
    }
    g_WM_MinimizedListGui := false
}

WM_MinimizedList_Cancel(*) {
    global g_WM_MinimizedListExcludePickerActive
    if (g_WM_MinimizedListExcludePickerActive) {
        WM_MinimizedList_CancelExcludePicker()
        return
    }
    WM_MinimizedList_Cleanup()
}

WM_MinimizedList_Refresh(closedHwnd := 0) {
    global g_WM_MinimizedListActive, g_WM_MinimizedListRefreshing, g_WM_MinimizedListExcludePickerActive,
        g_WM_MinimizedListExcludePickerRows, g_WM_MinimizedListExcludePickerMap, g_WM_MinimizedListCloseModeArmed,
        g_WM_MinimizedListCollectForeHwnd
    if (!g_WM_MinimizedListActive || g_WM_MinimizedListRefreshing)
        return
    g_WM_MinimizedListRefreshing := true
    WM_BackgroundTitleExcludes_Init()
    g_WM_MinimizedListCloseModeArmed := false
    g_WM_MinimizedListExcludePickerActive := false
    g_WM_MinimizedListExcludePickerRows := []
    g_WM_MinimizedListExcludePickerMap := Map()
    try {
        if (closedHwnd)
            WM_MinimizedList_WaitForHwndClosed(closedHwnd)
        WM_MinimizedList_UnbindHotkeys()
        collectForeHwnd := g_WM_MinimizedListCollectForeHwnd
        rows := WM_CollectBackgroundWindows(collectForeHwnd)
        rows := WM_MinimizedList_FilterExcludedHwnd(rows, closedHwnd)
        g_WM_MinimizedListRows := rows
        if (rows.Length = 0) {
            WM_MinimizedList_Cleanup()
            WM_PresentNoBackgroundWindowsEmpty(, false)
            return
        }
        WM_ShowMinimizedBackgroundList(rows, true)
    } finally {
        g_WM_MinimizedListRefreshing := false
    }
}

WM_ShowMinimizedBackgroundList(rows := unset, refresh := false) {
    global g_WM_MinimizedListGui, g_WM_MinimizedListActive, g_WM_MinimizedListRows, g_WM_MinimizedCharSequence
    if (g_WM_MinimizedListActive && !refresh)
        return
    if (WM_DebugBackgroundEnabled())
        WM_DebugBackgroundWindowScan()
    if (!IsSet(rows)) {
        global g_WM_MinimizedListCollectForeHwnd
        rows := WM_CollectBackgroundWindows(g_WM_MinimizedListCollectForeHwnd)
    }
    if (rows.Length = 0) {
        if (refresh)
            WM_MinimizedList_Cleanup()
        WM_PresentNoBackgroundWindowsEmpty(, false)
        return
    }
    g_WM_MinimizedListRows := rows
    if (refresh && IsObject(g_WM_MinimizedListGui)) {
        try g_WM_MinimizedListGui.Destroy()
        g_WM_MinimizedListGui := false
    }
    windows := WM_MinimizedList_AssignKeys(rows)
    displayText := WM_MinimizedList_BuildDisplayText(rows, windows)
    fontSize := 11
    baseWidth := 480
    lineHeight := fontSize + 6
    lineCount := StrSplit(displayText, "`n").Length
    textControlHeight := lineCount * lineHeight + 10
    g_WM_MinimizedListGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Owner")
    g_WM_MinimizedListGui.BackColor := "1E1E2E"
    g_WM_MinimizedListGui.MarginX := 15
    g_WM_MinimizedListGui.MarginY := 10
    g_WM_MinimizedListGui.SetFont("s" . fontSize . " cCDD6F4", "Segoe UI")
    g_WM_MinimizedListGui.Add("Text", "w" . (baseWidth - 30), displayText)
    g_WM_MinimizedListGui.OnEvent("Escape", WM_MinimizedList_Cancel)
    if (!refresh)
        g_WM_MinimizedListActive := true
    WM_MinimizedList_BindEscape()
    WM_MinimizedList_BindHotkeys(windows)
    WM_MinimizedList_RepositionToActiveMonitor(0, g_WM_MinimizedListGui)
    WM_MinimizedList_ActivateGui(g_WM_MinimizedListGui)
    WM_MinimizedList_StartActiveMonitorTracking()
}
