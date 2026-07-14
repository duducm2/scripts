; =============================================================================
; AutoSlot — optional auto-position for newly opened windows (multi-monitor).
;
; Detection/placement policy is self-contained. 50/50 placement reuses the
; proven gapless snap APIs from WindowManagement\tile_snap.ahk (same as ^!#x).
; After a successful snap, a 2s ShowWithKeys modal offers [M] undo.
; On window close, fills the freed monitor slot from background windows.
; Delete: remove the #include in WindowManagement.ahk + this folder (see README).
;
; Placement (MonitorGetCount() > 1 only):
;   1) First empty ordinal monitor → maximize onto it
;   2) Else first half-full ordinal → 50/50
;   3) Else maximize new window in place
;
; Fill-on-close (same multi-monitor gate):
;   Shell destroy primary (WinEvent deduped) → debounce/cooldown → occupancy 0/1
;   already-filled skip → else WM_CollectBackgroundWindows → maximize or gapless
;   50/50 without strict validate; no candidate → companion heal.
;
; Enable toggle: Win+Alt+Shift+W [5]; persisted in assets\data\wm_autoslot.ini.
; =============================================================================

global g_AutoSlotHook := 0
global g_AutoSlotHookCb := 0
global g_AutoSlotShellMsg := 0
global g_AutoSlotPending := Map()
global g_AutoSlotRecent := Map()
global g_AutoSlotGui := 0
global g_AutoSlotUndo := 0
global g_AutoSlotFillPending := Map()
global g_AutoSlotHealPending := Map()
global g_AutoSlotFillRetry := Map()
global g_AutoSlotHwndMon := Map()
global g_AutoSlotFillCooldown := Map()
global g_AutoSlotLastDestroyHwnd := 0
global g_AutoSlotLastDestroyTick := 0
global g_AutoSlotEnabled := true

AutoSlot_EVENT_OBJECT_DESTROY := 0x8001
AutoSlot_EVENT_OBJECT_SHOW := 0x8002
AutoSlot_OBJID_WINDOW := 0
AutoSlot_DEBOUNCE_MS := 250
AutoSlot_RECENT_MS := 4000
AutoSlot_FILL_COOLDOWN_MS := 1500
AutoSlot_FILL_RETRY_MS := 400
AutoSlot_DESTROY_DEDUP_MS := 250
AutoSlot_MAX_ORDINAL := 4
AutoSlot_HSHELL_WINDOWCREATED := 1
AutoSlot_HSHELL_WINDOWDESTROYED := 2
AutoSlot_UNDO_MODAL_MS := 2000

; --- Enable / persist --------------------------------------------------------

AutoSlot_IniPath() {
    return A_ScriptDir "\assets\data\wm_autoslot.ini"
}

AutoSlot_IsEnabled() {
    global g_AutoSlotEnabled
    return !!g_AutoSlotEnabled
}

AutoSlot_LoadEnabled() {
    global g_AutoSlotEnabled
    path := AutoSlot_IniPath()
    raw := "1"
    try raw := IniRead(path, "AutoSlot", "Enabled", "1")
    catch
        raw := "1"
    g_AutoSlotEnabled := (raw = "1" || raw = "true" || raw = "True")
}

AutoSlot_SetEnabled(on) {
    global g_AutoSlotEnabled
    g_AutoSlotEnabled := !!on
    path := AutoSlot_IniPath()
    try {
        dir := A_ScriptDir "\assets\data"
        if !DirExist(dir)
            DirCreate(dir)
        IniWrite(g_AutoSlotEnabled ? "1" : "0", path, "AutoSlot", "Enabled")
    } catch {
    }
    return g_AutoSlotEnabled
}

; --- Init --------------------------------------------------------------------

AutoSlot_Init() {
    global g_AutoSlotHook, g_AutoSlotHookCb, g_AutoSlotShellMsg, g_AutoSlotGui
    if (g_AutoSlotHook || g_AutoSlotShellMsg)
        return

    AutoSlot_LoadEnabled()

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

    ; One hook covers DESTROY..SHOW; OnWinEvent branches by event.
    g_AutoSlotHookCb := CallbackCreate(AutoSlot_OnWinEvent, "F", 7)
    g_AutoSlotHook := DllCall("user32\SetWinEventHook",
        "UInt", AutoSlot_EVENT_OBJECT_DESTROY,
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
    if (!lParam)
        return 0
    hwnd := Integer(lParam)
    if (wParam = AutoSlot_HSHELL_WINDOWCREATED)
        AutoSlot_Schedule(hwnd)
    else if (wParam = AutoSlot_HSHELL_WINDOWDESTROYED)
        AutoSlot_OnDestroy(hwnd, true)  ; shell is primary fill trigger
    return 0
}

AutoSlot_OnWinEvent(hWinEventHook, event, hwnd, idObject, idChild, idEventThread, dwmsEventTime) {
    if (idObject != AutoSlot_OBJID_WINDOW || !hwnd)
        return
    hwnd := Integer(hwnd)
    if (event = AutoSlot_EVENT_OBJECT_SHOW)
        AutoSlot_Schedule(hwnd)
    else if (event = AutoSlot_EVENT_OBJECT_DESTROY)
        AutoSlot_OnDestroy(hwnd, false)  ; secondary; skip if shell just handled same hwnd
}

AutoSlot_RememberHwndMon(hwnd) {
    global g_AutoSlotHwndMon
    if (!hwnd)
        return
    monIdx := AutoSlot_GetHwndMonitorIndex(hwnd)
    if (monIdx >= 1)
        g_AutoSlotHwndMon[hwnd] := monIdx
}

AutoSlot_PruneHwndMon() {
    global g_AutoSlotHwndMon
    toDelete := []
    for h, _ in g_AutoSlotHwndMon {
        if (!DllCall("IsWindow", "ptr", h))
            toDelete.Push(h)
    }
    for h in toDelete
        g_AutoSlotHwndMon.Delete(h)
}

; Cheap gate before MonitorFromWindow — avoid DESTROY storm work for chrome/tool windows.
AutoSlot_DestroyLooksInteresting(hwnd) {
    global g_AutoSlotHwndMon
    if (!hwnd)
        return false
    if (g_AutoSlotHwndMon.Has(hwnd))
        return true
    if (!DllCall("IsWindow", "ptr", hwnd))
        return false
    try {
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

; fromShell: true for HSHELL_WINDOWDESTROYED (primary). WinEvent skips if same hwnd just armed.
AutoSlot_OnDestroy(hwnd, fromShell := false) {
    global g_AutoSlotHwndMon, g_AutoSlotLastDestroyHwnd, g_AutoSlotLastDestroyTick
    if (!hwnd || MonitorGetCount() <= 1)
        return

    cached := g_AutoSlotHwndMon.Has(hwnd) ? g_AutoSlotHwndMon[hwnd] : 0
    alive := !!DllCall("IsWindow", "ptr", hwnd)

    if (!fromShell) {
        if (hwnd = g_AutoSlotLastDestroyHwnd && A_TickCount - g_AutoSlotLastDestroyTick < AutoSlot_DESTROY_DEDUP_MS)
            return
        if (!cached && !AutoSlot_DestroyLooksInteresting(hwnd))
            return
    } else {
        ; Shell primary: skip chrome noise when alive; dead HWNDs need cache to resolve mon.
        if (!cached) {
            if (!alive)
                return
            if (!AutoSlot_DestroyLooksInteresting(hwnd))
                return
        }
    }

    monIdx := 0
    if (alive)
        monIdx := AutoSlot_GetHwndMonitorIndex(hwnd)
    if (monIdx < 1 && cached)
        monIdx := cached
    if (g_AutoSlotHwndMon.Has(hwnd))
        g_AutoSlotHwndMon.Delete(hwnd)
    if (monIdx < 1)
        return

    g_AutoSlotLastDestroyHwnd := hwnd
    g_AutoSlotLastDestroyTick := A_TickCount
    AutoSlot_ScheduleFill(monIdx)
}

; Place claimed this monitor — suppress deferred background fill (companion heal still allowed).
AutoSlot_ClaimMonitor(monIdx) {
    global g_AutoSlotFillPending, g_AutoSlotFillCooldown
    if (monIdx < 1)
        return
    g_AutoSlotFillCooldown[monIdx] := A_TickCount
    if (g_AutoSlotFillPending.Has(monIdx))
        g_AutoSlotFillPending.Delete(monIdx)
}

AutoSlot_FillCooldownActive(monIdx) {
    global g_AutoSlotFillCooldown
    return g_AutoSlotFillCooldown.Has(monIdx)
    && A_TickCount - g_AutoSlotFillCooldown[monIdx] < AutoSlot_FILL_COOLDOWN_MS
}

; During Place claim cooldown: still arm companion heal (not background promote).
AutoSlot_ScheduleHealOnly(monIdx) {
    global g_AutoSlotHealPending
    if (!AutoSlot_IsEnabled() || monIdx < 1 || MonitorGetCount() <= 1)
        return
    if (g_AutoSlotHealPending.Has(monIdx))
        return
    g_AutoSlotHealPending[monIdx] := true
    SetTimer(() => AutoSlot_ProcessHealOnly(monIdx), -AutoSlot_DEBOUNCE_MS)
}

AutoSlot_ProcessHealOnly(monIdx) {
    global g_AutoSlotHealPending, g_AutoSlotFillCooldown
    if (g_AutoSlotHealPending.Has(monIdx))
        g_AutoSlotHealPending.Delete(monIdx)
    if (!AutoSlot_IsEnabled())
        return
    if (AutoSlot_HealLoneCompanion(monIdx))
        g_AutoSlotFillCooldown[monIdx] := A_TickCount
}

AutoSlot_ScheduleFill(monIdx) {
    global g_AutoSlotFillPending
    if (!AutoSlot_IsEnabled())
        return
    if (monIdx < 1 || MonitorGetCount() <= 1)
        return
    ; Claim cooldown: block background promote, still allow companion heal.
    if (AutoSlot_FillCooldownActive(monIdx)) {
        AutoSlot_ScheduleHealOnly(monIdx)
        return
    }
    if (g_AutoSlotFillPending.Has(monIdx))
        return
    g_AutoSlotFillPending[monIdx] := true
    SetTimer(() => AutoSlot_ProcessFillPending(monIdx), -AutoSlot_DEBOUNCE_MS)
}

AutoSlot_ProcessFillPending(monIdx) {
    global g_AutoSlotFillPending, g_AutoSlotFillCooldown
    if (g_AutoSlotFillPending.Has(monIdx))
        g_AutoSlotFillPending.Delete(monIdx)
    if (!AutoSlot_IsEnabled())
        return
    ; Place may have claimed this monitor after the timer was armed — heal only.
    if (AutoSlot_FillCooldownActive(monIdx)) {
        if (AutoSlot_HealLoneCompanion(monIdx))
            g_AutoSlotFillCooldown[monIdx] := A_TickCount
        return
    }
    result := AutoSlot_FillMonitorFromBackground(monIdx)
    if (result = "stale") {
        AutoSlot_ScheduleFillRetry(monIdx)
        return
    }
    if (result = "ok")
        g_AutoSlotFillCooldown[monIdx] := A_TickCount
}

AutoSlot_ScheduleFillRetry(monIdx) {
    global g_AutoSlotFillRetry
    if (!AutoSlot_IsEnabled() || monIdx < 1)
        return
    if (g_AutoSlotFillRetry.Has(monIdx))
        return
    g_AutoSlotFillRetry[monIdx] := true
    SetTimer(() => AutoSlot_ProcessFillRetry(monIdx), -AutoSlot_FILL_RETRY_MS)
}

AutoSlot_ProcessFillRetry(monIdx) {
    global g_AutoSlotFillRetry, g_AutoSlotFillCooldown
    if (g_AutoSlotFillRetry.Has(monIdx))
        g_AutoSlotFillRetry.Delete(monIdx)
    if (!AutoSlot_IsEnabled())
        return
    if (AutoSlot_FillCooldownActive(monIdx)) {
        if (AutoSlot_HealLoneCompanion(monIdx))
            g_AutoSlotFillCooldown[monIdx] := A_TickCount
        return
    }
    result := AutoSlot_FillMonitorFromBackground(monIdx)
    ; One retry only — no further stale reschedule.
    if (result = "ok")
        g_AutoSlotFillCooldown[monIdx] := A_TickCount
    else if (result = "stale")
        AutoSlot_HealLoneCompanion(monIdx)  ; no-op if still 2+; heals if dead HWND cleared to 1
}

AutoSlot_CompanionAlreadyFilled(hwnd, monIdx) {
    if (!hwnd || monIdx < 1)
        return false
    try {
        if (WinGetMinMax("ahk_id " hwnd) = 1)
            return true
    } catch {
    }
    try {
        MonitorGetWorkArea monIdx, &wl, &wt, &wr, &wb
        rect := Buffer(16, 0)
        if !DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect)
            return false
        l := NumGet(rect, 0, "int"), t := NumGet(rect, 4, "int")
        r := NumGet(rect, 8, "int"), b := NumGet(rect, 12, "int")
        ; Work-area fill from WM_MaximizeHwndBackground (not OS-maximized).
        return (Abs(l - wl) <= 8 && Abs(t - wt) <= 8 && Abs(r - wr) <= 8 && Abs(b - wb) <= 8)
    } catch {
        return false
    }
}

AutoSlot_Schedule(hwnd) {
    global g_AutoSlotPending, g_AutoSlotRecent
    if (!AutoSlot_IsEnabled())
        return
    if (!hwnd || MonitorGetCount() <= 1)
        return
    if (g_AutoSlotRecent.Has(hwnd) && A_TickCount - g_AutoSlotRecent[hwnd] < AutoSlot_RECENT_MS)
        return
    if (g_AutoSlotPending.Has(hwnd))
        return
    AutoSlot_RememberHwndMon(hwnd)
    g_AutoSlotPending[hwnd] := true
    SetTimer(() => AutoSlot_ProcessPending(hwnd), -AutoSlot_DEBOUNCE_MS)
}

AutoSlot_ProcessPending(hwnd) {
    global g_AutoSlotPending, g_AutoSlotRecent
    g_AutoSlotPending.Delete(hwnd)
    if (!AutoSlot_IsEnabled())
        return
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
    AutoSlot_RememberHwndMon(hwnd)
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
    AutoSlot_PruneHwndMon()
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

AutoSlot_SnapPair(newHwnd, partnerHwnd, monIdx, acceptUnvalidated := false) {
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

    ; Single prepare (parity with ^!#x) — avoid double Sleep from AutoSlot_Prepare + WM_Prepare.
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
    ; Close-fill skips strict validate (gapless already applied); new-window place keeps it.
    if (!acceptUnvalidated) {
        if (!WM_WaitValidateSnapBipartitionStrict(monIdx, newHwnd, partnerHwnd))
            return ""
    }

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

; --- Fill-on-close (background → freed slot) ---------------------------------

AutoSlot_PickBackgroundCandidate(monIdx, occupancyRows) {
    occupied := Map()
    for row in occupancyRows {
        if (row.hwnd)
            occupied[row.hwnd] := true
    }
    rows := []
    try rows := WM_CollectBackgroundWindows()
    catch
        return 0
    for row in rows {
        hwnd := 0
        try hwnd := Integer(row.hwnd)
        catch
            continue
        if (!hwnd || !DllCall("IsWindow", "ptr", hwnd))
            continue
        if (occupied.Has(hwnd))
            continue
        if (AutoSlot_IsExcludedExeOrTitle(hwnd))
            continue
        return hwnd
    }
    return 0
}

AutoSlot_OccupancyHasRecent(occupancyRows) {
    global g_AutoSlotRecent
    for row in occupancyRows {
        h := row.hwnd
        if (h && g_AutoSlotRecent.Has(h) && A_TickCount - g_AutoSlotRecent[h] < AutoSlot_RECENT_MS)
            return true
    }
    return false
}

; Maximize the sole visible window on monIdx when it is not already filled.
AutoSlot_HealLoneCompanion(monIdx) {
    if (monIdx < 1 || monIdx > MonitorGetCount())
        return false
    others := AutoSlot_OccupancyOnMonitor(monIdx)
    if (others.Length != 1)
        return false
    companion := others[1].hwnd
    if (AutoSlot_CompanionAlreadyFilled(companion, monIdx))
        return true
    healed := false
    try healed := !!WM_MaximizeHwndBackground(companion)
    catch
        healed := false
    if (!healed) {
        try {
            AutoSlot_MaximizeHwnd(companion)
            healed := true
        } catch {
        }
    }
    if (healed) {
        order := AutoSlot_OrderForMonitorIndex(monIdx)
        label := order > 0 ? order : monIdx
        AutoSlot_Toast("ℹ️ Companion maximized → M" label)
    }
    return healed
}

; Returns "ok" | "noop" | "stale" (stale = occupancy still 2+, caller may retry).
AutoSlot_FillMonitorFromBackground(monIdx) {
    global g_AutoSlotRecent, g_AutoSlotUndo
    if (monIdx < 1 || monIdx > MonitorGetCount() || MonitorGetCount() <= 1)
        return "noop"
    others := AutoSlot_OccupancyOnMonitor(monIdx)
    if (others.Length >= 2)
        return "stale"

    order := AutoSlot_OrderForMonitorIndex(monIdx)
    label := order > 0 ? order : monIdx

    ; Leftover half of a pair — maximize companion (even if that hwnd is recent-placed).
    if (others.Length = 1) {
        if (AutoSlot_CompanionAlreadyFilled(others[1].hwnd, monIdx))
            return "ok"
        ; Do not promote background over a window AutoSlot_Place just assigned.
        if (AutoSlot_OccupancyHasRecent(others))
            return AutoSlot_HealLoneCompanion(monIdx) ? "ok" : "noop"
        cand := AutoSlot_PickBackgroundCandidate(monIdx, others)
        if (!cand)
            return AutoSlot_HealLoneCompanion(monIdx) ? "ok" : "noop"
        g_AutoSlotRecent[cand] := A_TickCount
        AutoSlot_PruneRecent()
        g_AutoSlotUndo := 0
        pane := AutoSlot_SnapPair(cand, others[1].hwnd, monIdx, true)
        AutoSlot_ClearUndo()
        if (pane != "") {
            AutoSlot_RememberHwndMon(cand)
            AutoSlot_Toast("ℹ️ Slot filled → M" label " " pane)
            return "ok"
        }
        if (AutoSlot_MaximizeOnMonitor(cand, monIdx)) {
            AutoSlot_RememberHwndMon(cand)
            AutoSlot_Toast("ℹ️ Slot filled → M" label " (maximized)")
            return "ok"
        }
        return AutoSlot_HealLoneCompanion(monIdx) ? "ok" : "noop"
    }

    ; Empty monitor — promote background (skip if somehow contested by recent; empty has none).
    cand := AutoSlot_PickBackgroundCandidate(monIdx, others)
    if (!cand)
        return "noop"
    g_AutoSlotRecent[cand] := A_TickCount
    AutoSlot_PruneRecent()
    if (AutoSlot_MaximizeOnMonitor(cand, monIdx)) {
        AutoSlot_RememberHwndMon(cand)
        AutoSlot_Toast("ℹ️ Slot filled → M" label " (maximized)")
        return "ok"
    }
    return "noop"
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
    AutoSlot_RememberHwndMon(hwnd)

    ; Empty-monitor-first: never 50/50 on an occupied screen when any ordinal is empty.
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
        if (AutoSlot_MaximizeOnMonitor(hwnd, emptyMon)) {
            AutoSlot_ClaimMonitor(emptyMon)
            msg := "ℹ️ Auto-slotted → M" emptyOrder " (maximized)"
        }
    } else if (halfMon && halfPartner) {
        pane := AutoSlot_SnapPair(hwnd, halfPartner, halfMon)
        if (pane != "") {
            AutoSlot_ClaimMonitor(halfMon)
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