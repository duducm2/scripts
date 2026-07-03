; =============================================================================
; Utils module: dictation_visible_paste.ahk
; Post-dictation visible-window picker: select a window and paste clipboard (Ctrl+V).
; =============================================================================

global g_DictationVisiblePasteGui := false
global g_DictationVisiblePasteActive := false
global g_DictationVisiblePasteResult := ""   ; "" = waiting, 0 = cancel, integer = hwnd
global g_DictationVisiblePasteKeyMap := Map()
global g_DictationVisiblePasteHotkeyHandlers := []
global g_DictationVisiblePasteKeysPollTimer := ""
global g_DictationVisiblePasteKeysPollPrev := Map()
global g_DictationVisiblePasteKeysPollCallbacks := Map()
global g_DictationVisiblePasteCharActionLock := false
global g_DictationVisiblePasteLastCharActionKey := ""
global g_DictationVisiblePasteLastCharActionTick := 0
global g_DictationVisiblePasteTrackTimer := ""
global g_DictationVisiblePasteLastForegroundMonitorIdx := 0
global g_DictationVisiblePasteCharSequence := ["1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "a", "b", "c", "d",
    "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z"]
global g_DictationVisiblePasteThumbnails := []  ; [{ thumbId, sourceHwnd }]

global DICTATION_VISIBLE_PASTE_COL_COUNT := 4
global DICTATION_VISIBLE_PASTE_ROW_COUNT := 2
global DICTATION_VISIBLE_PASTE_MAX_KEYS := 8
global DICTATION_VISIBLE_PASTE_COL_GAP := 18
global DICTATION_VISIBLE_PASTE_THUMB_MAX_W := 320
global DICTATION_VISIBLE_PASTE_TITLE_H := 48
global DICTATION_VISIBLE_PASTE_ROW_PAD := 22
global DICTATION_VISIBLE_PASTE_KEY_FONT := 16
global DICTATION_VISIBLE_PASTE_TITLE_FONT := 13
global DICTATION_VISIBLE_PASTE_HEADER_FONT := 14
global DICTATION_VISIBLE_PASTE_MON_LABEL_FONT := 11

; DWM thumbnail property flags
global DWM_TNP_RECTDESTINATION := 0x1
global DWM_TNP_VISIBLE := 0x8
global DWM_TNP_SOURCECLIENTAREAONLY := 0x10

Dictation_DwmRegisterThumbnail(destHwnd, sourceHwnd) {
    if (!destHwnd || !sourceHwnd || destHwnd = sourceHwnd)
        return 0
    thumbId := 0
    hr := 0
    try {
        hr := DllCall("dwmapi\DwmRegisterThumbnail", "Ptr", destHwnd, "Ptr", sourceHwnd, "Ptr*", &thumbId, "HRESULT")
    } catch {
        return 0
    }
    if (hr < 0 || !thumbId)
        return 0
    return thumbId
}

Dictation_DwmUpdateThumbnailRect(thumbId, x, y, w, h, sourceClientOnly := false) {
    if (!thumbId)
        return false
    props := Buffer(48, 0)
    flags := DWM_TNP_RECTDESTINATION | DWM_TNP_VISIBLE | DWM_TNP_SOURCECLIENTAREAONLY
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

Dictation_DwmUnregisterThumbnail(thumbId) {
    if (!thumbId)
        return
    try DllCall("dwmapi\DwmUnregisterThumbnail", "Ptr", thumbId, "HRESULT")
    catch {
    }
}

Dictation_VisiblePasteUnregisterAllThumbnails() {
    global g_DictationVisiblePasteThumbnails
    for entry in g_DictationVisiblePasteThumbnails {
        if (IsObject(entry) && entry.HasProp("thumbId"))
            Dictation_DwmUnregisterThumbnail(entry.thumbId)
    }
    g_DictationVisiblePasteThumbnails := []
}

Dictation_VisiblePasteTruncateTitle(title, maxLen := 38) {
    if (StrLen(title) <= maxLen)
        return title
    return SubStr(title, 1, maxLen - 1) . "…"
}

; Map panel control screen rect to destination GUI client coordinates (DWM rcDestination).
Dictation_VisiblePastePanelDestRect(gui, panel, &outX, &outY, &outW, &outH) {
    if (!Dictation_VisiblePasteGuiHasWindow(gui) || !IsObject(panel))
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

Dictation_VisiblePasteAttachThumbnails(gui, panelRows) {
    global g_DictationVisiblePasteThumbnails
    if (!Dictation_VisiblePasteGuiHasWindow(gui))
        return
    destHwnd := gui.Hwnd
    for row in panelRows {
        if (!IsObject(row) || !row.HasProp("panel") || !row.HasProp("sourceHwnd"))
            continue
        sourceHwnd := row.sourceHwnd
        if (!sourceHwnd || sourceHwnd = destHwnd || !WinExist("ahk_id " sourceHwnd))
            continue
        if (Dictation_IsExcludedPasteTarget(sourceHwnd))
            continue
        if (!Dictation_VisiblePastePanelDestRect(gui, row.panel, &px, &py, &pw, &ph))
            continue
        thumbId := Dictation_DwmRegisterThumbnail(destHwnd, sourceHwnd)
        if (!thumbId)
            continue
        if (!Dictation_DwmUpdateThumbnailRect(thumbId, px, py, pw, ph, false)) {
            Dictation_DwmUnregisterThumbnail(thumbId)
            continue
        }
        g_DictationVisiblePasteThumbnails.Push({ thumbId: thumbId, sourceHwnd: sourceHwnd })
    }
}

Dictation_IsExcludedPasteTarget(hwnd) {
    if (!hwnd)
        return true
    try {
        exe := StrLower(WinGetProcessName("ahk_id " hwnd))
    } catch {
        return true
    }
    if (exe = "handy.exe" || exe = "clipangel.exe")
        return true
    try {
        title := WinGetTitle(hwnd)
    } catch {
        return true
    }
    if (InStr(StrLower(title), "windowmanagement.ahk"))
        return true
    return false
}

Dictation_IsDigitSlotChar(c) {
    if (StrLen(c) != 1)
        return false
    o := Ord(c)
    return o >= Ord("0") && o <= Ord("9")
}

Dictation_VisiblePasteKeyLabel(char) {
    return (char = "0") ? "10" : char
}

; Left-to-right monitor order (parity with WindowManagement GetMonitorIndexByOrder).
Dictation_GetMonitorIndexByOrder(order) {
    count := MonitorGetCount()
    if (order < 1 || order > count)
        return 0
    monitors := []
    loop count {
        i := A_Index
        MonitorGet i, &l, &t, &r, &b
        cx := (l + r) // 2
        cy := (t + b) // 2
        monitors.Push({ idx: i, cx: cx, cy: cy })
    }
    n := monitors.Length
    loop n - 1 {
        loop n - A_Index {
            j := A_Index
            a := monitors[j]
            b := monitors[j + 1]
            if (a.cx > b.cx || (a.cx = b.cx && a.cy > b.cy)) {
                monitors[j] := b
                monitors[j + 1] := a
            }
        }
    }
    return monitors[order].idx
}

Dictation_GetSnapSplitAxis(monIdx) {
    MonitorGetWorkArea(monIdx, &wl, &wt, &wr, &wb)
    return (wr - wl >= wb - wt) ? "h" : "v"
}

; Returns grid row 1 = start pane (left/top), 2 = end pane (right/bottom).
Dictation_ClassifyWindowPaneRow(monIdx, left, top, right, bottom) {
    MonitorGetWorkArea(monIdx, &wl, &wt, &wr, &wb)
    if (Dictation_GetSnapSplitAxis(monIdx) = "h") {
        center := wl + (wr - wl) // 2
        return ((left + right) // 2 < center) ? 1 : 2
    }
    center := wt + (wb - wt) // 2
    return ((top + bottom) // 2 < center) ? 1 : 2
}

Dictation_PaneSizeOnAxis(monIdx, left, top, right, bottom) {
    return (Dictation_GetSnapSplitAxis(monIdx) = "h") ? (right - left) : (bottom - top)
}

Dictation_VisiblePasteResolveLayoutMonitor(centerOnHwnd := 0) {
    if (centerOnHwnd && WinExist("ahk_id " centerOnHwnd)) {
        rect := Buffer(16, 0)
        if (DllCall("GetWindowRect", "ptr", centerOnHwnd, "ptr", rect)) {
            cx := NumGet(rect, 0, "int") + (NumGet(rect, 8, "int") - NumGet(rect, 0, "int")) // 2
            cy := NumGet(rect, 4, "int") + (NumGet(rect, 12, "int") - NumGet(rect, 4, "int")) // 2
            loop MonitorGetCount() {
                MonitorGet(A_Index, &ml, &mt, &mr, &mb)
                if (cx >= ml && cx <= mr && cy >= mt && cy <= mb)
                    return A_Index
            }
        }
    }
    return GetMonitorIndexForForeground_StandardBar()
}

Dictation_VisiblePasteComputeLayout(centerOnHwnd := 0) {
    global DICTATION_VISIBLE_PASTE_COL_COUNT, DICTATION_VISIBLE_PASTE_COL_GAP, DICTATION_VISIBLE_PASTE_THUMB_MAX_W,
        DICTATION_VISIBLE_PASTE_TITLE_H, DICTATION_VISIBLE_PASTE_ROW_PAD
    colCount := DICTATION_VISIBLE_PASTE_COL_COUNT
    colGap := DICTATION_VISIBLE_PASTE_COL_GAP
    monIdx := Dictation_VisiblePasteResolveLayoutMonitor(centerOnHwnd)
    MonitorGetWorkArea(monIdx, &ml, &mt, &mr, &mb)
    workW := mr - ml
    contentW := Max(640, Floor(workW * 0.92))
    cellW := (contentW - (colCount - 1) * colGap) // colCount
    thumbW := Min(DICTATION_VISIBLE_PASTE_THUMB_MAX_W, cellW - 2)
    thumbH := Round(thumbW * 9 / 16)
    keyH := 28
    keyGap := 6
    titleGap := 12
    titleH := DICTATION_VISIBLE_PASTE_TITLE_H
    rowPad := DICTATION_VISIBLE_PASTE_ROW_PAD
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

; Fixed 4x2 grid: columns = monitors L→R, rows = start/end pane. Keys 1-8 column-major.
Dictation_BuildMonitorGrid() {
    global DICTATION_VISIBLE_PASTE_COL_COUNT, DICTATION_VISIBLE_PASTE_ROW_COUNT, DICTATION_VISIBLE_PASTE_MAX_KEYS
    displayCols := DICTATION_VISIBLE_PASTE_COL_COUNT
    rowCount := DICTATION_VISIBLE_PASTE_ROW_COUNT
    grid := []
    loop displayCols {
        colCells := []
        loop rowCount
            colCells.Push("")
        grid.Push(colCells)
    }
    overflowCount := 0
    physicalCols := Min(displayCols, MonitorGetCount())
    loop physicalCols {
        col := A_Index
        mon := Dictation_GetMonitorIndexByOrder(col)
        if (!mon)
            continue
        for win in Dictation_GetVisibleWindowsOnMonitor(mon) {
            paneRow := Dictation_ClassifyWindowPaneRow(mon, win.left, win.top, win.right, win.bottom)
            paneSize := Dictation_PaneSizeOnAxis(mon, win.left, win.top, win.right, win.bottom)
            title := ""
            try title := WinGetTitle(win.hwnd)
            candidate := { hwnd: win.hwnd, title: title, paneSize: paneSize, mon: mon }
            existing := grid[col][paneRow]
            if (!IsObject(existing) || !existing.HasProp("hwnd")) {
                grid[col][paneRow] := candidate
            } else if (paneSize > existing.paneSize) {
                overflowCount++
                grid[col][paneRow] := candidate
            } else {
                overflowCount++
            }
        }
    }
    slots := []
    keyIdx := 0
    loop displayCols {
        col := A_Index
        loop rowCount {
            row := A_Index
            cell := grid[col][row]
            if (!IsObject(cell) || !cell.HasProp("hwnd"))
                continue
            keyIdx++
            if (keyIdx > DICTATION_VISIBLE_PASTE_MAX_KEYS) {
                overflowCount++
                continue
            }
            ch := String(keyIdx)
            slot := {
                hwnd: cell.hwnd,
                title: cell.title,
                char: ch,
                label: Dictation_VisiblePasteKeyLabel(ch),
                col: col,
                row: row
            }
            slots.Push(slot)
            grid[col][row] := slot
        }
    }
    return { grid: grid, slots: slots, overflowCount: overflowCount, physicalCols: physicalCols }
}

; Legacy visible-window enumeration for one monitor (no daemon IPC).
Dictation_GetVisibleWindowsOnMonitor(mon) {
    MonitorGet mon, &ml, &mt, &mr, &mb
    cx := (ml + mr) // 2
    cy := (mt + mb) // 2
    point64 := (cy & 0xFFFFFFFF) << 32 | (cx & 0xFFFFFFFF)
    hTarget := DllCall("MonitorFromPoint", "int64", point64, "uint", 2, "ptr")

    hwnds := WinGetList()
    GWL_EXSTYLE := -20
    WS_EX_TOOLWINDOW := 0x00000080
    TOL := 40
    visible := []

    for hwnd in hwnds {
        zIdx := hwnds.Length - A_Index
        try {
            if (WinGetMinMax(hwnd) = -1)
                continue
            if !DllCall("IsWindowVisible", "ptr", hwnd)
                continue
            exStyle := DllCall("GetWindowLongPtr", "ptr", hwnd, "int", GWL_EXSTYLE, "ptr")
            if (exStyle & WS_EX_TOOLWINDOW)
                continue
            hMon := DllCall("MonitorFromWindow", "ptr", hwnd, "uint", 2, "ptr")
            if (Integer(hMon) != Integer(hTarget))
                continue
            class := WinGetClass(hwnd)
            if (class = "Progman" || class = "WorkerW")
                continue
            title := WinGetTitle(hwnd)
            if (title = "")
                continue
            if (Dictation_IsExcludedPasteTarget(hwnd))
                continue

            rect := Buffer(16, 0)
            if !DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect)
                continue
            left := NumGet(rect, 0, "int")
            top := NumGet(rect, 4, "int")
            right := NumGet(rect, 8, "int")
            bottom := NumGet(rect, 12, "int")

            centerX := (left + right) // 2
            centerY := (top + bottom) // 2
            covered := false
            for win in visible {
                if (centerX >= win.left && centerX <= win.right && centerY >= win.top && centerY <= win.bottom) {
                    covered := true
                    break
                }
            }
            if (covered)
                continue

            visible.Push({ hwnd: hwnd, left: left, top: top, right: right, bottom: bottom, z: zIdx })
        } catch {
            continue
        }
    }

    n := visible.Length
    if (n > 1) {
        loop n - 1 {
            i := A_Index
            loop n - i {
                j := A_Index
                rowDiff := visible[j].top - visible[j + 1].top
                if (rowDiff > TOL || (Abs(rowDiff) <= TOL && visible[j].left > visible[j + 1].left)) {
                    tmp := visible[j]
                    visible[j] := visible[j + 1]
                    visible[j + 1] := tmp
                }
            }
        }
    }
    return visible
}

Dictation_VisiblePasteModifiersDown() {
    try {
        return GetKeyState("LWin", "P") || GetKeyState("RWin", "P") || GetKeyState("Ctrl", "P") || GetKeyState("Alt",
            "P") || GetKeyState("Shift", "P")
    } catch {
        return false
    }
}

Dictation_VisiblePasteShouldCaptureKey() {
    global g_DictationVisiblePasteActive
    if (!g_DictationVisiblePasteActive)
        return false
    return !Dictation_VisiblePasteModifiersDown()
}

Dictation_VisiblePasteTryConsumeCharAction(char) {
    global g_DictationVisiblePasteCharActionLock, g_DictationVisiblePasteLastCharActionKey,
        g_DictationVisiblePasteLastCharActionTick
    if (g_DictationVisiblePasteCharActionLock)
        return false
    key := StrLower(char)
    if (g_DictationVisiblePasteLastCharActionKey = key && A_TickCount - g_DictationVisiblePasteLastCharActionTick < 400
    )
        return false
    g_DictationVisiblePasteCharActionLock := true
    g_DictationVisiblePasteLastCharActionKey := key
    g_DictationVisiblePasteLastCharActionTick := A_TickCount
    return true
}

Dictation_VisiblePasteReleaseCharActionLock() {
    global g_DictationVisiblePasteCharActionLock
    g_DictationVisiblePasteCharActionLock := false
}

Dictation_VisiblePasteKeyDown(keyName) {
    try {
        if (Dictation_IsDigitSlotChar(keyName))
            return GetKeyState(keyName, "P") || GetKeyState("Numpad" . keyName, "P")
        return GetKeyState(keyName, "P")
    } catch {
        return false
    }
}

Dictation_VisiblePasteStopKeysPoll() {
    global g_DictationVisiblePasteKeysPollTimer, g_DictationVisiblePasteKeysPollPrev,
        g_DictationVisiblePasteKeysPollCallbacks
    try SetTimer(g_DictationVisiblePasteKeysPollTimer, 0)
    catch {
    }
    g_DictationVisiblePasteKeysPollTimer := ""
    g_DictationVisiblePasteKeysPollPrev := Map()
    g_DictationVisiblePasteKeysPollCallbacks := Map()
}

Dictation_VisiblePasteKeysPoll() {
    global g_DictationVisiblePasteActive, g_DictationVisiblePasteKeysPollCallbacks, g_DictationVisiblePasteKeysPollPrev
    if (!g_DictationVisiblePasteActive) {
        Dictation_VisiblePasteStopKeysPoll()
        return
    }
    if (!Dictation_VisiblePasteShouldCaptureKey())
        return
    for keyName, cb in g_DictationVisiblePasteKeysPollCallbacks {
        if (!cb)
            continue
        isDown := Dictation_VisiblePasteKeyDown(keyName)
        wasDown := g_DictationVisiblePasteKeysPollPrev.Has(keyName) ? g_DictationVisiblePasteKeysPollPrev[keyName] :
            false
        g_DictationVisiblePasteKeysPollPrev[keyName] := isDown
        if (isDown && !wasDown) {
            try cb.Call()
            catch {
            }
        }
    }
}

Dictation_VisiblePasteStartKeysPoll(windows) {
    global g_DictationVisiblePasteKeysPollCallbacks, g_DictationVisiblePasteKeysPollPrev,
        g_DictationVisiblePasteKeysPollTimer
    Dictation_VisiblePasteStopKeysPoll()
    g_DictationVisiblePasteKeysPollCallbacks := Map()
    g_DictationVisiblePasteKeysPollPrev := Map()
    for w in windows {
        slotChar := w.char
        g_DictationVisiblePasteKeysPollCallbacks[slotChar] := Dictation_VisiblePasteHandleChar.Bind(slotChar)
        g_DictationVisiblePasteKeysPollPrev[slotChar] := Dictation_VisiblePasteKeyDown(slotChar)
    }
    if (g_DictationVisiblePasteKeysPollCallbacks.Count > 0)
        g_DictationVisiblePasteKeysPollTimer := SetTimer(Dictation_VisiblePasteKeysPoll, 50)
}

Dictation_VisiblePasteRegisterHotkey(hk, handler) {
    global g_DictationVisiblePasteHotkeyHandlers
    try {
        #InputLevel 10
        Hotkey(hk, handler, "On")
        #InputLevel 0
        g_DictationVisiblePasteHotkeyHandlers.Push({ hk: hk, handler: handler })
    } catch {
    }
}

Dictation_VisiblePasteUnbindHotkeys() {
    global g_DictationVisiblePasteHotkeyHandlers
    Dictation_VisiblePasteStopKeysPoll()
    try HotIf()
    catch {
    }
    for entry in g_DictationVisiblePasteHotkeyHandlers {
        try Hotkey(entry.hk, "Off")
        catch {
        }
    }
    g_DictationVisiblePasteHotkeyHandlers := []
    try HotIf()
    catch {
    }
}

Dictation_VisiblePasteBindHotkeys(windows) {
    Dictation_VisiblePasteUnbindHotkeys()
    if (windows.Length = 0)
        return
    try HotIf (*) => Dictation_VisiblePasteShouldCaptureKey()
    catch {
    }
    for w in windows {
        slotChar := w.char
        Dictation_VisiblePasteRegisterHotkey("$*" . slotChar, Dictation_VisiblePasteHandleChar.Bind(slotChar))
        if (Dictation_IsDigitSlotChar(slotChar))
            Dictation_VisiblePasteRegisterHotkey("$*Numpad" . slotChar, Dictation_VisiblePasteHandleChar.Bind(slotChar))
    }
    try HotIf()
    catch {
    }
    Dictation_VisiblePasteStartKeysPoll(windows)
}

Dictation_VisiblePasteGuiHasWindow(gui := unset) {
    global g_DictationVisiblePasteGui
    if (!IsSet(gui))
        gui := g_DictationVisiblePasteGui
    if (!IsObject(gui))
        return false
    try return !!gui.Hwnd
    catch
        return false
}

Dictation_VisiblePasteActivateGui(gui) {
    if (!Dictation_VisiblePasteGuiHasWindow(gui))
        return
    try WinActivate("ahk_id " gui.Hwnd)
    catch {
    }
    try DllCall("SetForegroundWindow", "ptr", gui.Hwnd)
    catch {
    }
}

Dictation_VisiblePasteRepositionToActiveMonitor(forMonitorIdx := 0, gui := unset) {
    global g_DictationVisiblePasteGui
    if (!IsSet(gui))
        gui := g_DictationVisiblePasteGui
    if (!Dictation_VisiblePasteGuiHasWindow(gui))
        return
    monitorIndex := forMonitorIdx ? forMonitorIdx : GetMonitorIndexForForeground_StandardBar()
    MonitorGetWorkArea(monitorIndex, &ml, &mt, &mr, &mb)
    mw := mr - ml
    mh := mb - mt
    try gui.Show("AutoSize Hide")
    catch {
        return
    }
    try gui.GetPos(&gx, &gy, &gw, &gh)
    catch {
        return
    }
    cx := ml + (mw - gw) // 2
    cy := mt + (mh - gh) // 2
    try gui.Show("x" . cx . " y" . cy . " NA")
    catch {
    }
    Dictation_VisiblePasteActivateGui(gui)
}

Dictation_VisiblePasteStopMonitorTracking() {
    global g_DictationVisiblePasteTrackTimer
    try SetTimer(Dictation_VisiblePasteTrackActiveMonitorTick, 0)
    catch {
    }
    g_DictationVisiblePasteTrackTimer := ""
}

Dictation_VisiblePasteTrackActiveMonitorTick(*) {
    global g_DictationVisiblePasteActive, g_DictationVisiblePasteLastForegroundMonitorIdx
    if (!g_DictationVisiblePasteActive) {
        Dictation_VisiblePasteStopMonitorTracking()
        return
    }
    newIdx := GetMonitorIndexForForeground_StandardBar()
    if (newIdx != g_DictationVisiblePasteLastForegroundMonitorIdx) {
        g_DictationVisiblePasteLastForegroundMonitorIdx := newIdx
        Dictation_VisiblePasteRepositionToActiveMonitor(newIdx)
    }
}

Dictation_VisiblePasteStartMonitorTracking() {
    global g_DictationVisiblePasteTrackTimer, g_DictationVisiblePasteLastForegroundMonitorIdx
    Dictation_VisiblePasteStopMonitorTracking()
    g_DictationVisiblePasteLastForegroundMonitorIdx := GetMonitorIndexForForeground_StandardBar()
    g_DictationVisiblePasteTrackTimer := SetTimer(Dictation_VisiblePasteTrackActiveMonitorTick, 115)
}

Dictation_VisiblePasteAssignKeys(slots) {
    global g_DictationVisiblePasteKeyMap
    g_DictationVisiblePasteKeyMap := Map()
    windows := []
    for slot in slots {
        g_DictationVisiblePasteKeyMap[slot.char] := slot.hwnd
        windows.Push(slot)
    }
    return windows
}

Dictation_VisiblePasteShowModal(gridData, centerOnHwnd := 0) {
    global g_DictationVisiblePasteGui, g_DictationVisiblePasteActive, g_DictationVisiblePasteResult
    global DICTATION_VISIBLE_PASTE_COL_COUNT, DICTATION_VISIBLE_PASTE_ROW_COUNT,
        DICTATION_VISIBLE_PASTE_KEY_FONT, DICTATION_VISIBLE_PASTE_TITLE_FONT, DICTATION_VISIBLE_PASTE_HEADER_FONT,
        DICTATION_VISIBLE_PASTE_MON_LABEL_FONT
    Dictation_VisiblePasteUnregisterAllThumbnails()
    slots := gridData.slots
    grid := gridData.grid
    overflowCount := gridData.overflowCount
    windows := Dictation_VisiblePasteAssignKeys(slots)
    layout := Dictation_VisiblePasteComputeLayout(centerOnHwnd)
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
    rowCount := DICTATION_VISIBLE_PASTE_ROW_COUNT
    monLabelH := 18
    monLabelGap := 4

    g_DictationVisiblePasteGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Owner")
    g_DictationVisiblePasteGui.Opt("-DPIScale")
    g_DictationVisiblePasteGui.BackColor := "1E1E2E"
    g_DictationVisiblePasteGui.MarginX := marginX
    g_DictationVisiblePasteGui.MarginY := marginY
    g_DictationVisiblePasteGui.SetFont("s" . DICTATION_VISIBLE_PASTE_HEADER_FONT . " cCDD6F4 Bold", "Segoe UI")
    g_DictationVisiblePasteGui.Add("Text", "w" . contentW . " Center", "=== VISIBLE WINDOWS ===")

    gridStartY := marginY + 34
    monRowY := gridStartY
    g_DictationVisiblePasteGui.SetFont("s" . DICTATION_VISIBLE_PASTE_MON_LABEL_FONT . " c6C7086", "Segoe UI")
    loop colCount {
        col := A_Index - 1
        cellX := marginX + col * (cellW + colGap)
        g_DictationVisiblePasteGui.Add("Text", "x" . cellX . " y" . monRowY . " w" . cellW . " h" . monLabelH .
            " Center",
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
                g_DictationVisiblePasteGui.SetFont("s" . DICTATION_VISIBLE_PASTE_KEY_FONT . " cCDD6F4 Bold", "Segoe UI"
                )
                g_DictationVisiblePasteGui.Add("Text", "x" . cellX . " y" . cellY . " w" . cellW . " h" . keyH .
                    " Center",
                    "[" . cell.label . "]")
                g_DictationVisiblePasteGui.Add("Text",
                    "x" . (thumbX - 1) . " y" . (thumbY - 1) . " w" . (thumbW + 2) . " h" . (thumbH + 2) .
                    " Background6C7086",
                    "")
                panel := g_DictationVisiblePasteGui.Add("Text",
                    "x" . thumbX . " y" . thumbY . " w" . thumbW . " h" . thumbH . " Background1E1E2E", "")
                g_DictationVisiblePasteGui.SetFont("s" . DICTATION_VISIBLE_PASTE_TITLE_FONT . " cCDD6F4", "Segoe UI")
                g_DictationVisiblePasteGui.Add("Text",
                    "x" . cellX . " y" . (thumbY + thumbH + titleGap) . " w" . cellW . " h" . titleH . " Center Wrap",
                    Dictation_VisiblePasteTruncateTitle(cell.title))
                thumbPanels.Push({ sourceHwnd: cell.hwnd, panel: panel })
            } else {
                g_DictationVisiblePasteGui.Add("Text", "x" . cellX . " y" . cellY . " w" . cellW . " h" . keyH .
                    " Center", "")
                g_DictationVisiblePasteGui.Add("Text",
                    "x" . (thumbX - 1) . " y" . (thumbY - 1) . " w" . (thumbW + 2) . " h" . (thumbH + 2) .
                    " Background45475A",
                    "")
                g_DictationVisiblePasteGui.Add("Text",
                    "x" . cellX . " y" . (thumbY + thumbH + titleGap) . " w" . cellW . " h" . titleH . " Center", "")
            }
        }
    }

    footerY := gridStartY + rowCount * cellH + 8

    if (overflowCount > 0) {
        g_DictationVisiblePasteGui.SetFont("s12 c6C7086", "Segoe UI")
        g_DictationVisiblePasteGui.Add("Text", "x" . marginX . " y" . footerY . " w" . contentW . " Center",
            "(" . overflowCount . " more — close some and reopen)")
        footerY += 24
    }

    g_DictationVisiblePasteGui.SetFont("s12 cCDD6F4", "Segoe UI")
    g_DictationVisiblePasteGui.Add("Text", "x" . marginX . " y" . footerY . " w" . contentW . " Center",
        ">>> slot key = GO + PASTE <<<")
    g_DictationVisiblePasteGui.Add("Text", "x" . marginX . " y" . (footerY + 24) . " w" . contentW . " Center",
    "[ESC] Cancel")
    g_DictationVisiblePasteGui.OnEvent("Escape", Dictation_VisiblePasteCancel)
    g_DictationVisiblePasteActive := true
    g_DictationVisiblePasteResult := ""
    Dictation_VisiblePasteBindHotkeys(windows)
    try {
        #InputLevel 10
        Hotkey("$*Escape", Dictation_VisiblePasteCancel, "On")
        #InputLevel 0
    } catch {
    }
    if (centerOnHwnd && WinExist("ahk_id " centerOnHwnd))
        Dictation_VisiblePasteRepositionToActiveMonitor(Dictation_VisiblePasteResolveLayoutMonitor(centerOnHwnd))
    else
        Dictation_VisiblePasteRepositionToActiveMonitor(0)
    Sleep 30
    Dictation_VisiblePasteAttachThumbnails(g_DictationVisiblePasteGui, thumbPanels)
    Dictation_VisiblePasteStartMonitorTracking()
}

Dictation_VisiblePasteClose() {
    global g_DictationVisiblePasteActive, g_DictationVisiblePasteGui, g_DictationVisiblePasteHotkeyHandlers
    if (!g_DictationVisiblePasteActive && !Dictation_VisiblePasteGuiHasWindow())
        return
    g_DictationVisiblePasteActive := false
    Dictation_VisiblePasteStopMonitorTracking()
    Dictation_VisiblePasteUnregisterAllThumbnails()
    g_DictationVisiblePasteKeyMap := Map()
    Dictation_VisiblePasteUnbindHotkeys()
    try {
        #InputLevel 10
        Hotkey("$*Escape", Dictation_VisiblePasteCancel, "Off")
        #InputLevel 0
    } catch {
    }
    if (Dictation_VisiblePasteGuiHasWindow()) {
        try g_DictationVisiblePasteGui.Destroy()
    }
    g_DictationVisiblePasteGui := false
}

Dictation_VisiblePasteCancel(*) {
    global g_DictationVisiblePasteResult
    g_DictationVisiblePasteResult := 0
    Dictation_VisiblePasteClose()
}

Dictation_VisiblePasteHandleChar(char, *) {
    global g_DictationVisiblePasteActive, g_DictationVisiblePasteKeyMap, g_DictationVisiblePasteResult
    if (!g_DictationVisiblePasteActive)
        return
    if (!Dictation_VisiblePasteTryConsumeCharAction(char))
        return
    try {
        hwnd := g_DictationVisiblePasteKeyMap.Get(char, "")
        if (hwnd = "")
            hwnd := g_DictationVisiblePasteKeyMap.Get(StrLower(char), "")
        if (!hwnd || !WinExist("ahk_id " hwnd))
            return
        g_DictationVisiblePasteResult := hwnd
        Dictation_VisiblePasteClose()
    } finally {
        Dictation_VisiblePasteReleaseCharActionLock()
    }
}

; Returns selected window HWND or 0 on cancel/timeout/no windows. Blocking with timeout.
Dictation_ShowVisiblePasteSelector(centerOnHwnd := 0) {
    global g_DictationVisiblePasteResult, g_DictationVisiblePasteLastForegroundMonitorIdx
    clipBackup := ""
    try clipBackup := A_Clipboard
    Dictation_VisiblePasteClose()
    gridData := Dictation_BuildMonitorGrid()
    if (gridData.slots.Length = 0) {
        ShowCenteredOverlay_Utils("ℹ️ No visible windows found", 2500, BANNER_ACCENT_INFO)
        try A_Clipboard := clipBackup
        return 0
    }
    Dictation_VisiblePasteShowModal(gridData, centerOnHwnd)
    start := A_TickCount
    timeoutMs := 30000
    lastMonitorIdx := g_DictationVisiblePasteLastForegroundMonitorIdx
    escWasDown := false
    try {
        while (g_DictationVisiblePasteResult = "") {
            if ((A_TickCount - start) >= timeoutMs)
                break
            Sleep 50
            isEscDown := false
            try isEscDown := GetKeyState("Escape", "P")
            if (isEscDown && !escWasDown && g_DictationVisiblePasteResult = "")
                Dictation_VisiblePasteCancel()
            escWasDown := isEscDown
            if (Dictation_VisiblePasteGuiHasWindow()) {
                curIdx := GetMonitorIndexForForeground_StandardBar()
                if (curIdx != lastMonitorIdx) {
                    lastMonitorIdx := curIdx
                    Dictation_VisiblePasteRepositionToActiveMonitor(curIdx)
                }
            }
        }
    } catch {
    }
    result := (g_DictationVisiblePasteResult = "") ? 0 : Integer(g_DictationVisiblePasteResult)
    Dictation_VisiblePasteClose()
    try A_Clipboard := clipBackup
    return result
}
