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

WM_MinimizedList_DwmRegisterThumbnail(destHwnd, sourceHwnd) {
    if (!destHwnd || !sourceHwnd || destHwnd = sourceHwnd)
        return 0
    thumbId := 0
    try {
        hr := DllCall("dwmapi\DwmRegisterThumbnail", "Ptr", destHwnd, "Ptr", sourceHwnd, "Ptr*", &thumbId, "HRESULT")
        if (hr < 0 || !thumbId)
            return 0
        return thumbId
    } catch {
        return 0
    }
}

WM_MinimizedList_DwmUpdateThumbnailRect(thumbId, x, y, w, h, sourceClientOnly := false) {
    global WM_DWM_TNP_RECTDESTINATION, WM_DWM_TNP_VISIBLE, WM_DWM_TNP_SOURCECLIENTAREAONLY
    if (!thumbId)
        return false
    props := Buffer(48, 0)
    flags := WM_DWM_TNP_RECTDESTINATION | WM_DWM_TNP_VISIBLE | WM_DWM_TNP_SOURCECLIENTAREAONLY
    NumPut("UInt", flags, props, 0)
    NumPut("Int", x, props, 4)
    NumPut("Int", y, props, 8)
    NumPut("Int", x + w, props, 12)
    NumPut("Int", y + h, props, 16)
    NumPut("Int", 1, props, 40)
    NumPut("Int", sourceClientOnly ? 1 : 0, props, 44)
    try {
        return DllCall("dwmapi\DwmUpdateThumbnailProperties", "Ptr", thumbId, "Ptr", props, "HRESULT") >= 0
    } catch {
        return false
    }
}

WM_MinimizedList_DwmUnregisterThumbnail(thumbId) {
    if (!thumbId)
        return
    try DllCall("dwmapi\DwmUnregisterThumbnail", "Ptr", thumbId, "HRESULT")
    catch {
    }
}

WM_MinimizedList_UnregisterAllThumbnails() {
    global g_WM_MinimizedListThumbnails
    for entry in g_WM_MinimizedListThumbnails {
        if (IsObject(entry) && entry.HasProp("thumbId"))
            WM_MinimizedList_DwmUnregisterThumbnail(entry.thumbId)
    }
    g_WM_MinimizedListThumbnails := []
}

WM_MinimizedList_PanelDestRect(gui, panel, &outX, &outY, &outW, &outH) {
    if (!WM_MinimizedList_GuiHasWindow(gui) || !IsObject(panel))
        return false
    try panelHwnd := panel.Hwnd
    catch
        return false
    if (!panelHwnd)
        return false
    guiHwnd := gui.Hwnd
    panelRect := Buffer(16, 0)
    if !DllCall("GetWindowRect", "Ptr", panelHwnd, "Ptr", panelRect)
        return false
    ptTL := Buffer(8, 0)
    ptBR := Buffer(8, 0)
    NumPut("Int", NumGet(panelRect, 0, "Int"), ptTL, 0)
    NumPut("Int", NumGet(panelRect, 4, "Int"), ptTL, 4)
    NumPut("Int", NumGet(panelRect, 8, "Int"), ptBR, 0)
    NumPut("Int", NumGet(panelRect, 12, "Int"), ptBR, 4)
    if !DllCall("MapWindowPoints", "Ptr", 0, "Ptr", guiHwnd, "Ptr", ptTL, "UInt", 1)
        return false
    if !DllCall("MapWindowPoints", "Ptr", 0, "Ptr", guiHwnd, "Ptr", ptBR, "UInt", 1)
        return false
    outX := NumGet(ptTL, 0, "Int")
    outY := NumGet(ptTL, 4, "Int")
    outW := NumGet(ptBR, 0, "Int") - outX
    outH := NumGet(ptBR, 4, "Int") - outY
    return outW > 2 && outH > 2
}

WM_MinimizedList_AttachThumbnails(gui, panelRows) {
    global g_WM_MinimizedListThumbnails
    if (!WM_MinimizedList_GuiHasWindow(gui))
        return
    destHwnd := gui.Hwnd
    for row in panelRows {
        if (!IsObject(row) || !row.HasProp("panel") || !row.HasProp("sourceHwnd"))
            continue
        sourceHwnd := row.sourceHwnd
        if (!sourceHwnd || sourceHwnd = destHwnd || !WinExist("ahk_id " sourceHwnd))
            continue
        if (!WM_MinimizedList_PanelDestRect(gui, row.panel, &px, &py, &pw, &ph))
            continue
        thumbId := WM_MinimizedList_DwmRegisterThumbnail(destHwnd, sourceHwnd)
        if (!thumbId)
            continue
        if (!WM_MinimizedList_DwmUpdateThumbnailRect(thumbId, px, py, pw, ph, false)) {
            WM_MinimizedList_DwmUnregisterThumbnail(thumbId)
            continue
        }
        g_WM_MinimizedListThumbnails.Push({ thumbId: thumbId, sourceHwnd: sourceHwnd })
    }
}

WM_MinimizedList_GridSlotChar(col, row) {
    if (row = 1)
        return ["a", "s", "d", "f"][col]
    return ["z", "x", "c", "v"][col]
}

WM_MinimizedList_GetWindowRestoreRect(hwnd, &left, &top, &right, &bottom) {
    left := top := right := bottom := 0
    if (!hwnd)
        return false
    wp := Buffer(44, 0)
    NumPut("UInt", 44, wp, 0)
    try {
        if !DllCall("GetWindowPlacement", "ptr", hwnd, "ptr", wp)
            return false
        left := NumGet(wp, 28, "int")
        top := NumGet(wp, 32, "int")
        right := NumGet(wp, 36, "int")
        bottom := NumGet(wp, 40, "int")
        if (right <= left || bottom <= top) {
            rect := Buffer(16, 0)
            if DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect) {
                left := NumGet(rect, 0, "int")
                top := NumGet(rect, 4, "int")
                right := NumGet(rect, 8, "int")
                bottom := NumGet(rect, 12, "int")
            }
        }
        return right > left && bottom > top
    } catch {
        return false
    }
}

WM_MinimizedList_GetWindowMonitorIndex(hwnd) {
    if (!hwnd)
        return 0
    try {
        hMon := DllCall("MonitorFromWindow", "ptr", hwnd, "uint", 2, "ptr")
        loop MonitorGetCount() {
            MonitorGet A_Index, &ml, &mt, &mr, &mb
            cx := (ml + mr) // 2
            cy := (mt + mb) // 2
            point64 := (cy & 0xFFFFFFFF) << 32 | (cx & 0xFFFFFFFF)
            if (Integer(DllCall("MonitorFromPoint", "int64", point64, "uint", 2, "ptr")) = Integer(hMon))
                return A_Index
        }
    } catch {
    }
    return 0
}

; Returns ordered monitor column 1..4, or 0 if beyond grid columns.
WM_MinimizedList_GetWindowOrderedMonitorCol(hwnd) {
    monIdx := WM_MinimizedList_GetWindowMonitorIndex(hwnd)
    if (!monIdx)
        return 0
    loop Min(WM_MINIMIZED_LIST_COL_COUNT, MonitorGetCount()) {
        if (GetMonitorIndexByOrder(A_Index) = monIdx)
            return A_Index
    }
    return 0
}

WM_MinimizedList_GetSnapSplitAxis(monIdx) {
    MonitorGetWorkArea(monIdx, &wl, &wt, &wr, &wb)
    return (wr - wl >= wb - wt) ? "h" : "v"
}

WM_MinimizedList_ClassifyWindowPaneRow(monIdx, left, top, right, bottom) {
    MonitorGetWorkArea(monIdx, &wl, &wt, &wr, &wb)
    if (WM_MinimizedList_GetSnapSplitAxis(monIdx) = "h") {
        center := wl + (wr - wl) // 2
        return ((left + right) // 2 < center) ? 1 : 2
    }
    center := wt + (wb - wt) // 2
    return ((top + bottom) // 2 < center) ? 1 : 2
}

WM_MinimizedList_PaneSizeOnAxis(monIdx, left, top, right, bottom) {
    return (WM_MinimizedList_GetSnapSplitAxis(monIdx) = "h") ? (right - left) : (bottom - top)
}

WM_MinimizedList_ComputeLayout() {
    global WM_MINIMIZED_LIST_COL_COUNT, WM_MINIMIZED_LIST_COL_GAP, WM_MINIMIZED_LIST_THUMB_MAX_W,
        WM_MINIMIZED_LIST_TITLE_H, WM_MINIMIZED_LIST_ROW_PAD
    colCount := WM_MINIMIZED_LIST_COL_COUNT
    colGap := WM_MINIMIZED_LIST_COL_GAP
    monIdx := GetMonitorIndexForForeground_StandardBar()
    MonitorGetWorkArea(monIdx, &ml, &mt, &mr, &mb)
    workW := mr - ml
    contentW := Max(640, Floor(workW * 0.92))
    cellW := (contentW - (colCount - 1) * colGap) // colCount
    thumbW := Min(WM_MINIMIZED_LIST_THUMB_MAX_W, cellW - 2)
    thumbH := Round(thumbW * 9 / 16)
    keyH := 28
    keyGap := 6
    titleGap := 12
    titleH := WM_MINIMIZED_LIST_TITLE_H
    rowPad := WM_MINIMIZED_LIST_ROW_PAD
    cellH := keyH + keyGap + thumbH + 2 + titleGap + titleH + rowPad
    return {
        colCount: colCount,
        colGap: colGap,
        cellW: cellW,
        cellH: cellH,
        thumbW: thumbW,
        thumbH: thumbH,
        keyH: keyH,
        keyGap: keyGap,
        titleGap: titleGap,
        titleH: titleH,
        contentW: colCount * cellW + (colCount - 1) * colGap
    }
}

WM_MinimizedList_BuildHiddenWindowGrid(rows) {
    global WM_MINIMIZED_LIST_COL_COUNT, WM_MINIMIZED_LIST_ROW_COUNT, g_WM_MinimizedCharSequence
    displayCols := WM_MINIMIZED_LIST_COL_COUNT
    rowCount := WM_MINIMIZED_LIST_ROW_COUNT
    grid := []
    loop displayCols {
        colCells := []
        loop rowCount
            colCells.Push("")
        grid.Push(colCells)
    }
    overflowWins := []
    byCol := Map()
    loop displayCols
        byCol[A_Index] := []
    for row in rows {
        col := WM_MinimizedList_GetWindowOrderedMonitorCol(row.hwnd)
        monIdx := WM_MinimizedList_GetWindowMonitorIndex(row.hwnd)
        left := top := right := bottom := 0
        if (!WM_MinimizedList_GetWindowRestoreRect(row.hwnd, &left, &top, &right, &bottom)) {
            overflowWins.Push({ hwnd: row.hwnd, title: row.title })
            continue
        }
        if (col < 1 || col > displayCols) {
            overflowWins.Push({ hwnd: row.hwnd, title: row.title })
            continue
        }
        paneSize := monIdx ? WM_MinimizedList_PaneSizeOnAxis(monIdx, left, top, right, bottom) : (right - left)
        byCol[col].Push({
            hwnd: row.hwnd,
            title: row.title,
            mon: monIdx,
            left: left,
            top: top,
            right: right,
            bottom: bottom,
            paneSize: paneSize
        })
    }
    loop displayCols {
        col := A_Index
        monWins := byCol[col]
        if (monWins.Length = 0)
            continue
        mon := GetMonitorIndexByOrder(col)
        if (!mon)
            continue
        for win in monWins {
            idx := A_Index
            if (idx > rowCount) {
                overflowWins.Push({ hwnd: win.hwnd, title: win.title })
                continue
            }
            grid[col][idx] := { hwnd: win.hwnd, title: win.title, paneSize: win.paneSize, mon: mon }
        }
    }
    gridSlots := []
    loop displayCols {
        col := A_Index
        loop rowCount {
            row := A_Index
            cell := grid[col][row]
            if (!IsObject(cell) || !cell.HasProp("hwnd"))
                continue
            ch := WM_MinimizedList_GridSlotChar(col, row)
            slot := {
                hwnd: cell.hwnd,
                title: cell.title,
                char: ch,
                label: StrUpper(ch),
                col: col,
                row: row,
                isOverflow: false
            }
            gridSlots.Push(slot)
            grid[col][row] := slot
        }
    }
    overflowSlots := []
    overflowKeys := []
    for ch in g_WM_MinimizedCharSequence {
        if (!WM_MinimizedList_IsReservedSlotChar(ch))
            overflowKeys.Push(ch)
    }
    loop overflowWins.Length {
        if (A_Index > overflowKeys.Length)
            break
        win := overflowWins[A_Index]
        ch := overflowKeys[A_Index]
        overflowSlots.Push({
            hwnd: win.hwnd,
            title: win.title,
            char: ch,
            label: WM_MinimizedList_KeyLabel(ch),
            isOverflow: true
        })
    }
    unkeyedOverflowCount := overflowWins.Length - overflowSlots.Length
    slots := []
    for slot in gridSlots
        slots.Push(slot)
    for slot in overflowSlots
        slots.Push(slot)
    return { grid: grid, slots: slots, overflowSlots: overflowSlots, unkeyedOverflowCount: unkeyedOverflowCount }
}

WM_MinimizedList_AssignKeysFromSlots(slots) {
    global g_WM_MinimizedKeyMap
    g_WM_MinimizedKeyMap := Map()
    windows := []
    for slot in slots {
        g_WM_MinimizedKeyMap[slot.char] := slot.hwnd
        windows.Push(slot)
    }
    return windows
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

; [E] and [Q] are command keys — never assign hidden-window overflow slots to e/q.
WM_MinimizedList_IsReservedSlotChar(char) {
    c := StrLower(char)
    return c = "e" || c = "q"
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
    WM_MinimizedList_RegisterHotkey("$*E", HandleMinimizedListAddExcludeTrigger)
    WM_MinimizedList_RegisterHotkey("$*Q", HandleMinimizedListCloseModeArm)
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
        WM_MinimizedList_Cleanup()
        Sleep 30
        WM_MinimizedList_OpenHwnd(hwnd)
    } finally {
        WM_MinimizedList_ReleaseCharActionLock()
    }
}

WM_MinimizedList_ShowGridModal(gridData) {
    global g_WM_MinimizedListGui, g_WM_MinimizedListActive, g_WM_MinimizedListCloseModeArmed
    global WM_MINIMIZED_LIST_ROW_COUNT, WM_MINIMIZED_LIST_KEY_FONT, WM_MINIMIZED_LIST_TITLE_FONT,
        WM_MINIMIZED_LIST_HEADER_FONT, WM_MINIMIZED_LIST_MON_LABEL_FONT
    WM_MinimizedList_StopActiveMonitorTracking()
    WM_MinimizedList_UnregisterAllThumbnails()
    if (WM_MinimizedList_GuiHasWindow()) {
        try g_WM_MinimizedListGui.Destroy()
        catch {
        }
        g_WM_MinimizedListGui := false
    }
    slots := gridData.slots
    grid := gridData.grid
    overflowSlots := gridData.overflowSlots
    unkeyedOverflowCount := gridData.unkeyedOverflowCount
    windows := WM_MinimizedList_AssignKeysFromSlots(slots)
    layout := WM_MinimizedList_ComputeLayout()
    marginX := 15
    marginY := 10
    colCount := layout.colCount
    colGap := layout.colGap
    cellW := layout.cellW
    cellH := layout.cellH
    thumbW := layout.thumbW
    thumbH := layout.thumbH
    keyH := layout.keyH
    keyGap := layout.keyGap
    titleGap := layout.titleGap
    titleH := layout.titleH
    contentW := layout.contentW
    rowCount := WM_MINIMIZED_LIST_ROW_COUNT
    monLabelH := 18
    monLabelGap := 4

    g_WM_MinimizedListGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Owner")
    g_WM_MinimizedListGui.Opt("-DPIScale")
    g_WM_MinimizedListGui.BackColor := "1E1E2E"
    g_WM_MinimizedListGui.MarginX := marginX
    g_WM_MinimizedListGui.MarginY := marginY
    g_WM_MinimizedListGui.SetFont("s" . WM_MINIMIZED_LIST_HEADER_FONT . " cCDD6F4 Bold", "Segoe UI")
    g_WM_MinimizedListGui.Add("Text", "w" . contentW . " Center", "=== HIDDEN WINDOWS ===")

    gridStartY := marginY + 34
    monRowY := gridStartY
    g_WM_MinimizedListGui.SetFont("s" . WM_MINIMIZED_LIST_MON_LABEL_FONT . " c6C7086", "Segoe UI")
    loop colCount {
        col := A_Index - 1
        cellX := marginX + col * (cellW + colGap)
        g_WM_MinimizedListGui.Add("Text", "x" . cellX . " y" . monRowY . " w" . cellW . " h" . monLabelH . " Center",
            "M" . A_Index)
    }
    gridStartY := monRowY + monLabelH + monLabelGap
    thumbPanels := []
    loop colCount {
        col := A_Index
        loop rowCount {
            row := A_Index
            rowIdx := row - 1
            cellX := marginX + (col - 1) * (cellW + colGap)
            cellY := gridStartY + rowIdx * cellH
            thumbX := cellX + (cellW - thumbW) // 2
            thumbY := cellY + keyH + keyGap
            cell := grid[col][row]
            occupied := IsObject(cell) && cell.HasProp("hwnd") && cell.HasProp("char")

            if (occupied) {
                g_WM_MinimizedListGui.SetFont("s" . WM_MINIMIZED_LIST_KEY_FONT . " cCDD6F4 Bold", "Segoe UI")
                g_WM_MinimizedListGui.Add("Text", "x" . cellX . " y" . cellY . " w" . cellW . " h" . keyH . " Center",
                    "[" . cell.label . "]")
                g_WM_MinimizedListGui.Add("Text",
                    "x" . (thumbX - 1) . " y" . (thumbY - 1) . " w" . (thumbW + 2) . " h" . (thumbH + 2) .
                    " Background6C7086",
                    "")
                panel := g_WM_MinimizedListGui.Add("Text",
                    "x" . thumbX . " y" . thumbY . " w" . thumbW . " h" . thumbH . " Background1E1E2E", "")
                g_WM_MinimizedListGui.SetFont("s" . WM_MINIMIZED_LIST_TITLE_FONT . " cCDD6F4", "Segoe UI")
                g_WM_MinimizedListGui.Add("Text",
                    "x" . cellX . " y" . (thumbY + thumbH + titleGap) . " w" . cellW . " h" . titleH . " Center Wrap",
                    WM_TruncateTitleForList(cell.title, 38))
                thumbPanels.Push({ sourceHwnd: cell.hwnd, panel: panel })
            } else {
                g_WM_MinimizedListGui.Add("Text", "x" . cellX . " y" . cellY . " w" . cellW . " h" . keyH . " Center",
                    "")
                g_WM_MinimizedListGui.Add("Text",
                    "x" . (thumbX - 1) . " y" . (thumbY - 1) . " w" . (thumbW + 2) . " h" . (thumbH + 2) .
                    " Background45475A",
                    "")
                g_WM_MinimizedListGui.Add("Text",
                    "x" . cellX . " y" . (thumbY + thumbH + titleGap) . " w" . cellW . " h" . titleH . " Center", "")
            }
        }
    }

    footerY := gridStartY + rowCount * cellH + 8

    if (overflowSlots.Length > 0) {
        footerY += 8
        g_WM_MinimizedListGui.SetFont("s11 c6C7086", "Segoe UI")
        g_WM_MinimizedListGui.Add("Text", "x" . marginX . " y" . footerY . " w" . contentW . " Left",
            "Additional windows:")
        footerY += 20
        ovKeyW := 36
        ovThumbW := Max(120, Min(160, cellW // 2))
        ovThumbH := Round(ovThumbW * 9 / 16)
        ovRowH := ovThumbH + 8
        ovTitleW := contentW - ovKeyW - ovThumbW - 16
        for slot in overflowSlots {
            rowY := footerY
            g_WM_MinimizedListGui.SetFont("s" . WM_MINIMIZED_LIST_KEY_FONT . " cCDD6F4 Bold", "Segoe UI")
            g_WM_MinimizedListGui.Add("Text", "x" . marginX . " y" . rowY . " w" . ovKeyW . " h" . ovThumbH .
                " Center", "[" . slot.label . "]")
            ovThumbX := marginX + ovKeyW + 4
            g_WM_MinimizedListGui.Add("Text",
                "x" . (ovThumbX - 1) . " y" . (rowY - 1) . " w" . (ovThumbW + 2) . " h" . (ovThumbH + 2) .
                " Background6C7086",
                "")
            ovPanel := g_WM_MinimizedListGui.Add("Text",
                "x" . ovThumbX . " y" . rowY . " w" . ovThumbW . " h" . ovThumbH . " Background1E1E2E", "")
            ovTitleX := ovThumbX + ovThumbW + 8
            g_WM_MinimizedListGui.SetFont("s" . WM_MINIMIZED_LIST_TITLE_FONT . " cCDD6F4", "Segoe UI")
            g_WM_MinimizedListGui.Add("Text",
                "x" . ovTitleX . " y" . rowY . " w" . ovTitleW . " h" . ovThumbH . " Left Wrap",
                WM_TruncateTitleForList(slot.title, 52))
            thumbPanels.Push({ sourceHwnd: slot.hwnd, panel: ovPanel })
            footerY += ovRowH
        }
    }

    if (unkeyedOverflowCount > 0) {
        g_WM_MinimizedListGui.SetFont("s12 c6C7086", "Segoe UI")
        g_WM_MinimizedListGui.Add("Text", "x" . marginX . " y" . footerY . " w" . contentW . " Center",
            "(" . unkeyedOverflowCount . " more — close some and reopen)")
        footerY += 24
    }

    g_WM_MinimizedListGui.SetFont("s12 cCDD6F4", "Segoe UI")
    if (g_WM_MinimizedListCloseModeArmed)
        g_WM_MinimizedListGui.Add("Text", "x" . marginX . " y" . footerY . " w" . contentW . " Center",
            ">>> Press a slot key to CLOSE <<<")
    else
        g_WM_MinimizedListGui.Add("Text", "x" . marginX . " y" . footerY . " w" . contentW . " Center",
            ">>> [Q] then slot key = CLOSE  |  slot key alone = OPEN <<<")
    footerY += 24
    g_WM_MinimizedListGui.Add("Text", "x" . marginX . " y" . footerY . " w" . contentW . " Center",
        "[Q] Arm close  [E] Add to exclude list  [ESC] Cancel")
    g_WM_MinimizedListGui.OnEvent("Escape", WM_MinimizedList_Cancel)
    WM_MinimizedList_RepositionToActiveMonitor(0, g_WM_MinimizedListGui)
    WM_MinimizedList_ActivateGui(g_WM_MinimizedListGui)
    if (g_WM_MinimizedListActive)
        WM_MinimizedList_StartActiveMonitorTracking()
    Sleep 30
    WM_MinimizedList_AttachThumbnails(g_WM_MinimizedListGui, thumbPanels)
    return windows
}

WM_MinimizedList_RepaintMainList() {
    global g_WM_MinimizedListRows
    if (g_WM_MinimizedListRows.Length = 0)
        return
    gridData := WM_MinimizedList_BuildHiddenWindowGrid(g_WM_MinimizedListRows)
    windows := WM_MinimizedList_ShowGridModal(gridData)
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
    WM_MinimizedList_UnregisterAllThumbnails()
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
    global g_WM_MinimizedListGui, g_WM_MinimizedListActive, g_WM_MinimizedListRows
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
    gridData := WM_MinimizedList_BuildHiddenWindowGrid(rows)
    if (gridData.slots.Length = 0) {
        if (refresh)
            WM_MinimizedList_Cleanup()
        WM_PresentNoBackgroundWindowsEmpty(, false)
        return
    }
    if (!refresh)
        g_WM_MinimizedListActive := true
    WM_MinimizedList_BindEscape()
    windows := WM_MinimizedList_ShowGridModal(gridData)
    WM_MinimizedList_BindHotkeys(windows)
}
