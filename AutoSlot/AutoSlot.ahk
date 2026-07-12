; =============================================================================
; AutoSlot — optional auto-position for newly opened windows (multi-monitor).
;
; Detection/placement policy is self-contained. 50/50 placement reuses the
; proven gapless snap APIs from WindowManagement\tile_snap.ahk (same as ^!#x).
; After a successful snap, a 2s ShowWithKeys modal offers [M] undo.
; Delete: remove the #include in WindowManagement.ahk + this folder (see README).
;
; Placement (MonitorGetCount() > 1 only):
;   1) Origin monitor has exactly 1 other window → restore if maximized, 50/50
;   2) Else first empty ordinal monitor → maximize onto it
;   3) Else first half-full ordinal → 50/50
;   4) Else maximize new window in place
; =============================================================================

global g_AutoSlotHook := 0
global g_AutoSlotHookCb := 0
global g_AutoSlotShellMsg := 0
global g_AutoSlotPending := Map()
global g_AutoSlotRecent := Map()
global g_AutoSlotGui := 0
global g_AutoSlotUndo := 0

AutoSlot_EVENT_OBJECT_SHOW := 0x8002
AutoSlot_OBJID_WINDOW := 0
AutoSlot_DEBOUNCE_MS := 250
AutoSlot_RECENT_MS := 4000
AutoSlot_MAX_ORDINAL := 4
AutoSlot_HSHELL_WINDOWCREATED := 1
AutoSlot_UNDO_MODAL_MS := 2000

; --- Init --------------------------------------------------------------------

AutoSlot_Init() {
    global g_AutoSlotHook, g_AutoSlotHookCb, g_AutoSlotShellMsg, g_AutoSlotGui
    if (g_AutoSlotHook || g_AutoSlotShellMsg)
        return

    if (!g_AutoSlotGui) {
        g_AutoSlotGui := Gui("+ToolWindow -Caption +E0x08000000")
        g_AutoSlotGui.Show("x0 y0 w0 h0 Hide")
    }
    sinkHwnd := g_AutoSlotGui.Hwnd
    if (DllCall("RegisterShellHookWindow", "ptr", sinkHwnd)) {
        g_AutoSlotShellMsg := DllCall("RegisterWindowMessage", "str", "SHELLHOOK", "uint")
        if (g_AutoSlotShellMsg)
            OnMessage(g_AutoSlotShellMsg, AutoSlot_OnShellHook)
    }

    g_AutoSlotHookCb := CallbackCreate(AutoSlot_OnWinEvent, "F", 7)
    g_AutoSlotHook := DllCall("user32\SetWinEventHook",
        "UInt", AutoSlot_EVENT_OBJECT_SHOW,
        "UInt", AutoSlot_EVENT_OBJECT_SHOW,
        "Ptr", 0,
        "Ptr", g_AutoSlotHookCb,
        "UInt", 0,
        "UInt", 0,
        "UInt", 0,
        "Ptr")

    OnMessage(0x007E, AutoSlot_OnDisplayChange)  ; WM_DISPLAYCHANGE
}

AutoSlot_OnDisplayChange(*) {
    return 0
}

AutoSlot_OnShellHook(wParam, lParam, *) {
    if (wParam = AutoSlot_HSHELL_WINDOWCREATED && lParam)
        AutoSlot_Schedule(Integer(lParam))
    return 0
}

AutoSlot_OnWinEvent(hWinEventHook, event, hwnd, idObject, idChild, idEventThread, dwmsEventTime) {
    if (event != AutoSlot_EVENT_OBJECT_SHOW || idObject != AutoSlot_OBJID_WINDOW || !hwnd)
        return
    AutoSlot_Schedule(Integer(hwnd))
}

AutoSlot_Schedule(hwnd) {
    global g_AutoSlotPending, g_AutoSlotRecent
    if (!hwnd || MonitorGetCount() <= 1)
        return
    if (g_AutoSlotRecent.Has(hwnd) && A_TickCount - g_AutoSlotRecent[hwnd] < AutoSlot_RECENT_MS)
        return
    if (g_AutoSlotPending.Has(hwnd))
        return
    g_AutoSlotPending[hwnd] := true
    SetTimer(() => AutoSlot_ProcessPending(hwnd), -AutoSlot_DEBOUNCE_MS)
}

AutoSlot_ProcessPending(hwnd) {
    global g_AutoSlotPending, g_AutoSlotRecent
    g_AutoSlotPending.Delete(hwnd)
    if (!hwnd || !DllCall("IsWindow", "ptr", hwnd))
        return
    if (MonitorGetCount() <= 1)
        return
    if (g_AutoSlotRecent.Has(hwnd) && A_TickCount - g_AutoSlotRecent[hwnd] < AutoSlot_RECENT_MS)
        return
    if (!AutoSlot_IsEligibleNewWindow(hwnd))
        return
    g_AutoSlotRecent[hwnd] := A_TickCount
    AutoSlot_PruneRecent()
    AutoSlot_Place(hwnd)
}

AutoSlot_PruneRecent() {
    global g_AutoSlotRecent
    now := A_TickCount
    toDelete := []
    for h, tick in g_AutoSlotRecent {
        if (now - tick > AutoSlot_RECENT_MS * 2)
            toDelete.Push(h)
    }
    for h in toDelete
        g_AutoSlotRecent.Delete(h)
}

; --- Eligibility / excludes (local; not WM_Background*) ----------------------

AutoSlot_IsExcludedExeOrTitle(hwnd) {
    try {
        exe := StrLower(WinGetProcessName("ahk_id " hwnd))
    } catch {
        exe := ""
    }
    if (exe = "handy.exe" || exe = "clipangel.exe")
        return true
    try {
        title := WinGetTitle(hwnd)
    } catch {
        return false
    }
    t := StrLower(title)
    if (InStr(t, "windowmanagement.ahk"))
        return true
    if (InStr(t, "autohotkey") && InStr(StrLower(WinGetClass(hwnd)), "ahk"))
        return true
    return false
}

AutoSlot_IsDesktopOrTaskbarClass(cls) {
    return cls = "Progman" || cls = "WorkerW" || cls = "Shell_TrayWnd" || cls = "Shell_SecondaryTrayWnd"
}

AutoSlot_IsEligibleNewWindow(hwnd) {
    if (!hwnd)
        return false
    try {
        if (!DllCall("IsWindowVisible", "ptr", hwnd))
            return false
        if (DllCall("GetParent", "ptr", hwnd))
            return false
        if (WinGetMinMax("ahk_id " hwnd) = -1)
            return false
        exStyle := DllCall("GetWindowLongPtr", "ptr", hwnd, "int", -20, "ptr")
        if (exStyle & 0x00000080)
            return false
        class := WinGetClass(hwnd)
        if (AutoSlot_IsDesktopOrTaskbarClass(class))
            return false
        title := WinGetTitle(hwnd)
        if (title = "")
            return false
        if (AutoSlot_IsExcludedExeOrTitle(hwnd))
            return false
    } catch {
        return false
    }
    return true
}

AutoSlot_IsOccupancyCandidate(hwnd, excludeHwnd := 0) {
    if (!hwnd || (excludeHwnd && hwnd = excludeHwnd))
        return false
    try {
        if (WinGetMinMax("ahk_id " hwnd) = -1)
            return false
        if (!DllCall("IsWindowVisible", "ptr", hwnd))
            return false
        if (DllCall("GetParent", "ptr", hwnd))
            return false
        exStyle := DllCall("GetWindowLongPtr", "ptr", hwnd, "int", -20, "ptr")
        if (exStyle & 0x00000080)
            return false
        class := WinGetClass(hwnd)
        if (AutoSlot_IsDesktopOrTaskbarClass(class))
            return false
        if (WinGetTitle(hwnd) = "")
            return false
        if (AutoSlot_IsExcludedExeOrTitle(hwnd))
            return false
    } catch {
        return false
    }
    return true
}

; --- Monitors / occupancy ----------------------------------------------------

AutoSlot_GetMonitorIndexByOrder(order) {
    count := MonitorGetCount()
    if (order < 1 || order > count)
        return 0
    monitors := []
    loop count {
        i := A_Index
        MonitorGet i, &l, &t, &r, &b
        monitors.Push({ idx: i, cx: (l + r) // 2, cy: (t + b) // 2 })
    }
    n := monitors.Length
    loop n - 1 {
        i := A_Index
        loop n - i {
            j := A_Index
            a := monitors[j]
            b := monitors[j + 1]
            if (a.cx > b.cx || (a.cx == b.cx && a.cy > b.cy)) {
                monitors[j] := b
                monitors[j + 1] := a
            }
        }
    }
    return monitors[order].idx
}

AutoSlot_OrderForMonitorIndex(monIdx) {
    ordinalCount := Min(MonitorGetCount(), AutoSlot_MAX_ORDINAL)
    loop ordinalCount {
        if (AutoSlot_GetMonitorIndexByOrder(A_Index) = monIdx)
            return A_Index
    }
    return 0
}

AutoSlot_GetHwndMonitorIndex(hwnd) {
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

AutoSlot_OccupancyOnMonitor(monIdx, excludeHwnd := 0) {
    rows := []
    if (monIdx < 1 || monIdx > MonitorGetCount())
        return rows
    MonitorGet monIdx, &ml, &mt, &mr, &mb
    cx := (ml + mr) // 2
    cy := (mt + mb) // 2
    point64 := (cy & 0xFFFFFFFF) << 32 | (cx & 0xFFFFFFFF)
    hTarget := DllCall("MonitorFromPoint", "int64", point64, "uint", 2, "ptr")

    for hwnd in WinGetList() {
        if (!AutoSlot_IsOccupancyCandidate(hwnd, excludeHwnd))
            continue
        try {
            hMon := DllCall("MonitorFromWindow", "ptr", hwnd, "uint", 2, "ptr")
            if (Integer(hMon) != Integer(hTarget))
                continue
            rect := Buffer(16, 0)
            if !DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect)
                continue
            rows.Push({
                hwnd: hwnd,
                left: NumGet(rect, 0, "int"),
                top: NumGet(rect, 4, "int"),
                right: NumGet(rect, 8, "int"),
                bottom: NumGet(rect, 12, "int")
            })
        } catch {
            continue
        }
    }
    return rows
}

; --- Geometry / move ---------------------------------------------------------

AutoSlot_PrepareHwnd(hwnd) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    if (AutoSlot_IsExcludedExeOrTitle(hwnd))
        return false
    try {
        state := WinGetMinMax("ahk_id " hwnd)
        if (state = 1 || state = -1) {
            WinRestore "ahk_id " hwnd
            Sleep 80
        }
        WinShow "ahk_id " hwnd
    } catch {
        return false
    }
    return true
}

AutoSlot_MoveHwndToRect(hwnd, left, top, right, bottom) {
    w := right - left
    h := bottom - top
    if (!hwnd || w < 1 || h < 1)
        return false
    ok := 0
    try ok := WinMove(hwnd, left, top, w, h)
    catch
        ok := 0
    if !ok {
        try ok := DllCall("MoveWindow", "ptr", hwnd, "int", left, "int", top, "int", w, "int", h, "int", true)
    }
    return !!ok
}

AutoSlot_MaximizeHwnd(hwnd) {
    if !hwnd
        return
    try {
        WinMaximize "ahk_id " hwnd
    } catch {
        try PostMessage 0x0112, 0xF030, , , "ahk_id " hwnd
    }
}

AutoSlot_ActivateHwnd(hwnd) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    try {
        WinActivate("ahk_id " hwnd)
        if (!WinActive("ahk_id " hwnd))
            WinWaitActive("ahk_id " hwnd, , 0.3)
        return !!WinActive("ahk_id " hwnd)
    } catch {
        return false
    }
}

AutoSlot_MaximizeOnMonitor(hwnd, monIdx) {
    if (!hwnd || monIdx < 1 || monIdx > MonitorGetCount())
        return false
    if (!AutoSlot_PrepareHwnd(hwnd))
        return false
    MonitorGet monIdx, &left, &top, &right, &bottom
    AutoSlot_MoveHwndToRect(hwnd, left, top, right, bottom)
    AutoSlot_MaximizeHwnd(hwnd)
    AutoSlot_ActivateHwnd(hwnd)
    return true
}

AutoSlot_MaximizeInPlace(hwnd) {
    if (!AutoSlot_PrepareHwnd(hwnd))
        return false
    AutoSlot_MaximizeHwnd(hwnd)
    AutoSlot_ActivateHwnd(hwnd)
    return true
}

; 50/50 via tile_snap gapless APIs (same engine as ^!#x / WM_SnapHalfPairActiveWindow).
; Returns pane name the new window was placed into, or "".
; On success, stashes undo context in g_AutoSlotUndo (partner pre-snap state).
AutoSlot_CapturePartnerState(partnerHwnd) {
    if (!partnerHwnd || !WinExist("ahk_id " partnerHwnd))
        return 0
    minMax := 0
    try minMax := WinGetMinMax("ahk_id " partnerHwnd)
    catch
        minMax := 0
    rect := Buffer(16, 0)
    if !DllCall("GetWindowRect", "ptr", partnerHwnd, "ptr", rect)
        return 0
    return {
        hwnd: partnerHwnd,
        left: NumGet(rect, 0, "int"),
        top: NumGet(rect, 4, "int"),
        right: NumGet(rect, 8, "int"),
        bottom: NumGet(rect, 12, "int"),
        minMax: minMax,
        wasMaximized: (minMax = 1)
    }
}

AutoSlot_SnapPair(newHwnd, partnerHwnd, monIdx) {
    global g_AutoSlotUndo
    if (!newHwnd || !partnerHwnd || monIdx < 1)
        return ""

    partnerState := AutoSlot_CapturePartnerState(partnerHwnd)
    partnerWasMax := IsObject(partnerState) ? partnerState.wasMaximized : false
    if (!partnerWasMax) {
        try partnerWasMax := (WinGetMinMax("ahk_id " partnerHwnd) = 1)
        catch
            partnerWasMax := false
    }

    if (!WM_PrepareHwndForTile(newHwnd) || !WM_PrepareHwndForTile(partnerHwnd))
        return ""

    axis := WM_GetSnapSplitAxis(monIdx)
    partnerPane := "start"
    if (!partnerWasMax && WM_GetWindowRectHwnd(partnerHwnd, &pl, &pt, &pr, &pb)) {
        WM_GetSnapInnerWorkArea(monIdx, &wl, &wt, &wr, &wb)
        pane := ""
        paneSize := 0
        if (WM_ClassifySnapPane(axis, wl, wt, wr, wb, pl, pt, pr, pb, &pane, &paneSize) && pane != "")
            partnerPane := pane
    }
    newPane := (partnerPane = "start") ? "end" : "start"

    ok := WM_SnapPairGaplessRects(monIdx, axis, newHwnd, newPane, partnerHwnd, partnerPane)
    if (!ok)
        return ""
    if (!WM_WaitValidateSnapBipartitionStrict(monIdx, newHwnd, partnerHwnd))
        return ""

    if (IsObject(partnerState)) {
        g_AutoSlotUndo := {
            newHwnd: newHwnd,
            partner: partnerState,
            snapMonIdx: monIdx,
            newPane: newPane
        }
    } else {
        g_AutoSlotUndo := 0
    }
    AutoSlot_ActivateHwnd(newHwnd)
    return newPane
}

; --- Toast / undo modal ------------------------------------------------------

AutoSlot_Toast(msg) {
    if (msg = "")
        return
    try ShowCenteredOverlay_Utils(msg, 1800, BANNER_ACCENT_INFO)
    catch {
    }
}

AutoSlot_ClearUndo() {
    global g_AutoSlotUndo
    g_AutoSlotUndo := 0
}

AutoSlot_CloseUndoModal() {
    try StandardLoadingBar_CloseKeysOverlay()
    catch {
    }
    try StandardLoadingBar_Hide(0)
    catch {
    }
}

AutoSlot_RestorePartner(partnerState) {
    if (!IsObject(partnerState) || !partnerState.hwnd)
        return false
    hwnd := partnerState.hwnd
    if (!WinExist("ahk_id " hwnd))
        return false
    try {
        state := WinGetMinMax("ahk_id " hwnd)
        if (state = 1 || state = -1) {
            WinRestore "ahk_id " hwnd
            Sleep 80
        }
    } catch {
    }
    if (!AutoSlot_MoveHwndToRect(hwnd, partnerState.left, partnerState.top, partnerState.right, partnerState.bottom))
        return false
    if (partnerState.wasMaximized) {
        Sleep 50
        AutoSlot_MaximizeHwnd(hwnd)
    }
    return true
}

AutoSlot_TargetMonitorForUndoMax() {
    monIdx := AutoSlot_GetMonitorIndexByOrder(2)
    if (monIdx >= 1)
        return monIdx
    try return MonitorGetPrimary()
    catch
        return 1
}

AutoSlot_OnUndoM(*) {
    global g_AutoSlotUndo
    undo := g_AutoSlotUndo
    AutoSlot_CloseUndoModal()
    if (!IsObject(undo) || !undo.newHwnd) {
        AutoSlot_ClearUndo()
        return
    }
    newHwnd := undo.newHwnd
    partnerState := undo.partner
    AutoSlot_ClearUndo()

    monIdx := AutoSlot_TargetMonitorForUndoMax()
    if (WinExist("ahk_id " newHwnd))
        AutoSlot_MaximizeOnMonitor(newHwnd, monIdx)
    if (IsObject(partnerState))
        AutoSlot_RestorePartner(partnerState)
    ; Partner restore can steal focus — always finish on the new window.
    AutoSlot_ActivateHwnd(newHwnd)
    AutoSlot_Toast("ℹ️ Auto-slot undone → M2")
}

AutoSlot_OnUndoTimeout(*) {
    global g_AutoSlotUndo
    newHwnd := IsObject(g_AutoSlotUndo) ? g_AutoSlotUndo.newHwnd : 0
    AutoSlot_CloseUndoModal()
    AutoSlot_ClearUndo()
    AutoSlot_ActivateHwnd(newHwnd)
}

AutoSlot_ShowUndoModal(paneLabel := "") {
    global g_AutoSlotUndo
    if (!IsObject(g_AutoSlotUndo) || !g_AutoSlotUndo.newHwnd)
        return
    newHwnd := g_AutoSlotUndo.newHwnd
    try StandardLoadingBar_CloseKeysOverlay()
    catch {
    }
    try StandardLoadingBar_Hide(0)
    catch {
    }
    AutoSlot_ActivateHwnd(newHwnd)
    suffix := paneLabel != "" ? " (" paneLabel ")" : ""
    keyCallbacks := Map("M", AutoSlot_OnUndoM)
    try {
        StandardLoadingBar_ShowWithKeys(
            "❓ Undo auto-slot? (2s)" suffix,
            keyCallbacks,
            AutoSlot_UNDO_MODAL_MS,
            0,
            AutoSlot_OnUndoTimeout,
            BANNER_ACCENT_INTERMEDIATE,
            480,
            17,
            "",
            true,
            "[M] Max on M2 + restore partner",
            true,
            true,
            true
        )
    } catch {
        AutoSlot_Toast("ℹ️ Auto-slotted" suffix)
        AutoSlot_ClearUndo()
    }
    ; After modal dismiss (timeout/Esc/M), ensure the new window is foreground.
    AutoSlot_ActivateHwnd(newHwnd)
}

; --- Placement ---------------------------------------------------------------

AutoSlot_Place(hwnd) {
    global g_AutoSlotUndo
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return
    ordinalCount := Min(MonitorGetCount(), AutoSlot_MAX_ORDINAL)
    if (ordinalCount < 2)
        return

    g_AutoSlotUndo := 0
    msg := ""

    ; 1) Origin-first: exactly one other window on the monitor where the new hwnd appeared.
    originMon := AutoSlot_GetHwndMonitorIndex(hwnd)
    if (originMon >= 1) {
        originOthers := AutoSlot_OccupancyOnMonitor(originMon, hwnd)
        if (originOthers.Length = 1) {
            pane := AutoSlot_SnapPair(hwnd, originOthers[1].hwnd, originMon)
            order := AutoSlot_OrderForMonitorIndex(originMon)
            if (pane != "") {
                label := order > 0 ? order : originMon
                AutoSlot_ShowUndoModal("M" label " " pane)
                return
            }
            if (AutoSlot_MaximizeInPlace(hwnd)) {
                AutoSlot_Toast("ℹ️ Auto-slot snap failed — maximized")
                return
            }
        }
    }

    emptyOrder := 0
    emptyMon := 0
    halfOrder := 0
    halfMon := 0
    halfPartner := 0

    loop ordinalCount {
        order := A_Index
        monIdx := AutoSlot_GetMonitorIndexByOrder(order)
        if (!monIdx)
            continue
        others := AutoSlot_OccupancyOnMonitor(monIdx, hwnd)
        if (others.Length = 0) {
            if (!emptyMon) {
                emptyOrder := order
                emptyMon := monIdx
            }
        } else if (others.Length = 1) {
            if (!halfMon) {
                halfOrder := order
                halfMon := monIdx
                halfPartner := others[1].hwnd
            }
        }
    }

    if (emptyMon) {
        if (AutoSlot_MaximizeOnMonitor(hwnd, emptyMon))
            msg := "ℹ️ Auto-slotted → M" emptyOrder " (maximized)"
    } else if (halfMon && halfPartner) {
        pane := AutoSlot_SnapPair(hwnd, halfPartner, halfMon)
        if (pane != "") {
            AutoSlot_ShowUndoModal("M" halfOrder " " pane)
            return
        }
        if (AutoSlot_MaximizeInPlace(hwnd))
            msg := "ℹ️ Auto-slot snap failed — maximized"
    } else {
        if (AutoSlot_MaximizeInPlace(hwnd))
            msg := "ℹ️ Grid full — maximized"
    }

    AutoSlot_Toast(msg)
}

; Auto-execute when #included.
AutoSlot_Init()