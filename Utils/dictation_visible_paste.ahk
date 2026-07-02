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

Dictation_SortVisiblePasteRows(&rows) {
    n := rows.Length
    if (n < 2)
        return
    loop n - 1 {
        loop n - A_Index {
            j := A_Index
            a := rows[j]
            b := rows[j + 1]
            if (StrCompare(a.title, b.title, true) > 0) {
                tmp := rows[j]
                rows[j] := rows[j + 1]
                rows[j + 1] := tmp
            }
        }
    }
}

Dictation_CollectVisibleWindowRows(promoteHwnd := 0) {
    rows := []
    seen := Map()
    loop MonitorGetCount() {
        for win in Dictation_GetVisibleWindowsOnMonitor(A_Index) {
            if (seen.Has(win.hwnd))
                continue
            seen[win.hwnd] := true
            try {
                rows.Push({ hwnd: win.hwnd, title: WinGetTitle(win.hwnd) })
            } catch {
            }
        }
    }
    Dictation_SortVisiblePasteRows(&rows)
    if (promoteHwnd && rows.Length > 1) {
        loop rows.Length {
            if (rows[A_Index].hwnd = promoteHwnd) {
                if (A_Index > 1) {
                    swap := rows[1]
                    rows[1] := rows[A_Index]
                    rows[A_Index] := swap
                }
                break
            }
        }
    }
    return rows
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

Dictation_VisiblePasteAssignKeys(rows) {
    global g_DictationVisiblePasteKeyMap, g_DictationVisiblePasteCharSequence
    g_DictationVisiblePasteKeyMap := Map()
    windows := []
    limit := Min(rows.Length, g_DictationVisiblePasteCharSequence.Length)
    loop limit {
        ch := g_DictationVisiblePasteCharSequence[A_Index]
        row := rows[A_Index]
        g_DictationVisiblePasteKeyMap[ch] := row.hwnd
        windows.Push({ hwnd: row.hwnd, title: row.title, char: ch, label: Dictation_VisiblePasteKeyLabel(ch) })
    }
    return windows
}

Dictation_VisiblePasteBuildDisplayText(rows, windows) {
    global g_DictationVisiblePasteCharSequence
    slotCount := g_DictationVisiblePasteCharSequence.Length
    displayText := "=== VISIBLE WINDOWS ===`n`n"
    for w in windows
        displayText .= "[" . w.label . "] " . w.title . "`n"
    if (rows.Length > slotCount)
        displayText .= "`n(" . (rows.Length - slotCount) . " more — close some and reopen)`n"
    displayText .= "`n>>> slot key = GO + PASTE <<<`n"
    displayText .= "`n[ESC] Cancel"
    return displayText
}

Dictation_VisiblePasteClose() {
    global g_DictationVisiblePasteActive, g_DictationVisiblePasteGui, g_DictationVisiblePasteHotkeyHandlers
    if (!g_DictationVisiblePasteActive && !Dictation_VisiblePasteGuiHasWindow())
        return
    g_DictationVisiblePasteActive := false
    Dictation_VisiblePasteStopMonitorTracking()
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

Dictation_VisiblePasteShowModal(rows, centerOnHwnd := 0) {
    global g_DictationVisiblePasteGui, g_DictationVisiblePasteActive, g_DictationVisiblePasteResult
    windows := Dictation_VisiblePasteAssignKeys(rows)
    displayText := Dictation_VisiblePasteBuildDisplayText(rows, windows)
    fontSize := 11
    baseWidth := 480
    g_DictationVisiblePasteGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Owner")
    g_DictationVisiblePasteGui.BackColor := "1E1E2E"
    g_DictationVisiblePasteGui.MarginX := 15
    g_DictationVisiblePasteGui.MarginY := 10
    g_DictationVisiblePasteGui.SetFont("s" . fontSize . " cCDD6F4", "Segoe UI")
    g_DictationVisiblePasteGui.Add("Text", "w" . (baseWidth - 30), displayText)
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
    if (centerOnHwnd && WinExist("ahk_id " centerOnHwnd)) {
        monitorIndex := 1
        rect := Buffer(16, 0)
        if (DllCall("GetWindowRect", "ptr", centerOnHwnd, "ptr", rect)) {
            cx := NumGet(rect, 0, "int") + (NumGet(rect, 8, "int") - NumGet(rect, 0, "int")) // 2
            cy := NumGet(rect, 4, "int") + (NumGet(rect, 12, "int") - NumGet(rect, 4, "int")) // 2
            loop MonitorGetCount() {
                MonitorGet(A_Index, &ml, &mt, &mr, &mb)
                if (cx >= ml && cx <= mr && cy >= mt && cy <= mb) {
                    monitorIndex := A_Index
                    break
                }
            }
        }
        Dictation_VisiblePasteRepositionToActiveMonitor(monitorIndex)
    } else {
        Dictation_VisiblePasteRepositionToActiveMonitor(0)
    }
    Dictation_VisiblePasteStartMonitorTracking()
}

; Returns selected window HWND or 0 on cancel/timeout/no windows. Blocking with timeout.
Dictation_ShowVisiblePasteSelector(centerOnHwnd := 0) {
    global g_DictationVisiblePasteResult, g_DictationVisiblePasteLastForegroundMonitorIdx
    clipBackup := ""
    try clipBackup := A_Clipboard
    Dictation_VisiblePasteClose()
    rows := Dictation_CollectVisibleWindowRows(centerOnHwnd)
    if (rows.Length = 0) {
        ShowCenteredOverlay_Utils("ℹ️ No visible windows found", 2500, BANNER_ACCENT_INFO)
        try A_Clipboard := clipBackup
        return 0
    }
    Dictation_VisiblePasteShowModal(rows, centerOnHwnd)
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
