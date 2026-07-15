; =============================================================================
; AutoSlot — optional auto-position for newly opened windows (multi-monitor).
;
; Detection/placement policy is self-contained. 50/50 placement reuses the
; proven gapless snap APIs from WindowManagement\tile_snap.ahk (same as ^!#x).
; After a successful snap, a 2s ShowWithKeys modal offers [M] undo.
; On window close, fills the freed monitor slot from background windows.
; Delete: remove the #include in WindowManagement.ahk + this folder (see README).
;
; Placement (MonitorGetCount() > 1 only); 2 slots per ordinal monitor (max 8):
;   1) First empty ordinal → maximize onto it
;   2) Else origin has exactly 1 other (max or half) → 50/50 on origin
;   3) Else first half-full ordinal anywhere → 50/50 (maximized still has a free half-slot)
;   4) Else maximize new window in place
;
; Fill-on-close (same multi-monitor gate):
;   Shell destroy primary (WinEvent deduped) → debounce/cooldown → occupancy 0/1
;   empty → two backgrounds 50/50 or one maximize; half → SnapPair bg else heal.
;   Place freeze / claim cooldown → heal-only (no background import).
;
; Rearrange-on-move (same fill/heal rules, no visible reshuffle):
;   MOVESIZEEND / suite leave / MaximizeOnMonitor → debounced RearrangeUnderfilled.
;
; Foreground monitor swap (suite MoveWinToMonitor when AutoSlot ON):
;   Full↔pair / half↔full / full↔full — exchange whole-monitor layouts; quiet rearrange.
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
global g_AutoSlotPlaceFreezeUntil := 0
global g_AutoSlotSnapPairs := Map()
global g_AutoSlotMaxCompanionSuppress := Map()
global g_AutoSlotPairMaxPending := Map()
global g_AutoSlotLocHook := 0
global g_AutoSlotLocHookCb := 0
global g_AutoSlotMoveHook := 0
global g_AutoSlotMoveHookCb := 0
global g_AutoSlotRearrangePending := false
global g_AutoSlotRearrangeExclude := 0
global g_AutoSlotSwapQuietUntil := 0

AutoSlot_EVENT_OBJECT_DESTROY := 0x8001
AutoSlot_EVENT_OBJECT_SHOW := 0x8002
AutoSlot_EVENT_OBJECT_LOCATIONCHANGE := 0x800B
AutoSlot_EVENT_SYSTEM_MOVESIZEEND := 0x000B
AutoSlot_OBJID_WINDOW := 0
AutoSlot_DEBOUNCE_MS := 250
AutoSlot_RECENT_MS := 4000
AutoSlot_FILL_COOLDOWN_MS := 1500
AutoSlot_FILL_RETRY_MS := 400
AutoSlot_DESTROY_DEDUP_MS := 250
AutoSlot_PAIR_MAX_DEBOUNCE_MS := 120
AutoSlot_PAIR_SUPPRESS_MS := 500
AutoSlot_REARRANGE_MS := 350
AutoSlot_SWAP_QUIET_MS := 900
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
    global g_AutoSlotLocHook, g_AutoSlotLocHookCb, g_AutoSlotMoveHook, g_AutoSlotMoveHookCb
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

    ; Narrow hook: paired maximize when one half of a 50/50 becomes maximized.
    g_AutoSlotLocHookCb := CallbackCreate(AutoSlot_OnLocationChange, "F", 7)
    g_AutoSlotLocHook := DllCall("user32\SetWinEventHook",
        "UInt", AutoSlot_EVENT_OBJECT_LOCATIONCHANGE,
        "UInt", AutoSlot_EVENT_OBJECT_LOCATIONCHANGE,
        "Ptr", 0,
        "Ptr", g_AutoSlotLocHookCb,
        "UInt", 0,
        "UInt", 0,
        "UInt", 0,
        "Ptr")

    ; Manual drag/resize end → rearrange underfilled monitors.
    g_AutoSlotMoveHookCb := CallbackCreate(AutoSlot_OnMoveSizeEnd, "F", 7)
    g_AutoSlotMoveHook := DllCall("user32\SetWinEventHook",
        "UInt", AutoSlot_EVENT_SYSTEM_MOVESIZEEND,
        "UInt", AutoSlot_EVENT_SYSTEM_MOVESIZEEND,
        "Ptr", 0,
        "Ptr", g_AutoSlotMoveHookCb,
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
    if (!hwnd)
        return

    AutoSlot_UnregisterSnapPair(hwnd)

    if (MonitorGetCount() <= 1)
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

; --- 50/50 pair registry (maximize either side → maximize companion) ----------

AutoSlot_RegisterSnapPair(a, b) {
    global g_AutoSlotSnapPairs
    if (!a || !b || a = b)
        return
    g_AutoSlotSnapPairs[a] := b
    g_AutoSlotSnapPairs[b] := a
}

AutoSlot_UnregisterSnapPair(hwnd) {
    global g_AutoSlotSnapPairs
    if (!hwnd || !g_AutoSlotSnapPairs.Has(hwnd))
        return
    partner := g_AutoSlotSnapPairs[hwnd]
    g_AutoSlotSnapPairs.Delete(hwnd)
    if (partner && g_AutoSlotSnapPairs.Has(partner) && g_AutoSlotSnapPairs[partner] = hwnd)
        g_AutoSlotSnapPairs.Delete(partner)
}

AutoSlot_PairSuppressActive(hwnd) {
    global g_AutoSlotMaxCompanionSuppress
    if (!hwnd || !g_AutoSlotMaxCompanionSuppress.Has(hwnd))
        return false
    return A_TickCount - g_AutoSlotMaxCompanionSuppress[hwnd] < AutoSlot_PAIR_SUPPRESS_MS
}

AutoSlot_PairSuppressMark(hwnd) {
    global g_AutoSlotMaxCompanionSuppress
    if (hwnd)
        g_AutoSlotMaxCompanionSuppress[hwnd] := A_TickCount
}

AutoSlot_OnLocationChange(hWinEventHook, event, hwnd, idObject, idChild, idEventThread, dwmsEventTime) {
    if (idObject != AutoSlot_OBJID_WINDOW || !hwnd)
        return
    if (!AutoSlot_IsEnabled())
        return
    hwnd := Integer(hwnd)
    global g_AutoSlotSnapPairs, g_AutoSlotPairMaxPending
    if (!g_AutoSlotSnapPairs.Has(hwnd))
        return
    if (AutoSlot_PairSuppressActive(hwnd))
        return
    try {
        if (WinGetMinMax("ahk_id " hwnd) != 1)
            return
    } catch {
        return
    }
    if (g_AutoSlotPairMaxPending.Has(hwnd))
        return
    g_AutoSlotPairMaxPending[hwnd] := true
    SetTimer(() => AutoSlot_ProcessPairedMaximizePending(hwnd), -AutoSlot_PAIR_MAX_DEBOUNCE_MS)
}

AutoSlot_ProcessPairedMaximizePending(hwnd) {
    global g_AutoSlotPairMaxPending
    if (g_AutoSlotPairMaxPending.Has(hwnd))
        g_AutoSlotPairMaxPending.Delete(hwnd)
    AutoSlot_OnPairedMaximize(hwnd)
}

; When hwnd is (or became) maximized, maximize its registered 50/50 companion.
AutoSlot_OnPairedMaximize(hwnd) {
    global g_AutoSlotSnapPairs
    if (!AutoSlot_IsEnabled() || !hwnd)
        return false
    if (AutoSlot_PairSuppressActive(hwnd))
        return false
    if (!g_AutoSlotSnapPairs.Has(hwnd))
        return false
    partner := g_AutoSlotSnapPairs[hwnd]
    if (!partner || !DllCall("IsWindow", "ptr", partner)) {
        AutoSlot_UnregisterSnapPair(hwnd)
        return false
    }
    try {
        if (WinGetMinMax("ahk_id " hwnd) != 1)
            return false
    } catch {
        return false
    }
    monIdx := AutoSlot_GetHwndMonitorIndex(partner)
    if (monIdx < 1)
        monIdx := AutoSlot_GetHwndMonitorIndex(hwnd)
    if (monIdx >= 1 && AutoSlot_CompanionAlreadyFilled(partner, monIdx))
        return false

    AutoSlot_PairSuppressMark(hwnd)
    AutoSlot_PairSuppressMark(partner)
    healed := false
    try healed := !!WM_MaximizeHwndBackground(partner)
    catch
        healed := false
    if (!healed) {
        try {
            AutoSlot_MaximizeHwnd(partner)
            healed := true
        } catch {
        }
    }
    return healed
}

; --- Rearrange underfilled slots after a move --------------------------------

AutoSlot_BeginSwapQuiet(ms := 0) {
    global g_AutoSlotSwapQuietUntil
    if (ms < 1)
        ms := AutoSlot_SWAP_QUIET_MS
    g_AutoSlotSwapQuietUntil := A_TickCount + ms
}

AutoSlot_SwapQuietActive() {
    global g_AutoSlotSwapQuietUntil
    return g_AutoSlotSwapQuietUntil > 0 && A_TickCount < g_AutoSlotSwapQuietUntil
}

AutoSlot_ScheduleRearrange(excludeHwnd := 0) {
    global g_AutoSlotRearrangePending, g_AutoSlotRearrangeExclude
    if (!AutoSlot_IsEnabled() || MonitorGetCount() <= 1)
        return
    if (AutoSlot_SwapQuietActive())
        return
    g_AutoSlotRearrangeExclude := excludeHwnd
    g_AutoSlotRearrangePending := true
    SetTimer(AutoSlot_ProcessRearrange, -AutoSlot_REARRANGE_MS)
}

AutoSlot_ProcessRearrange(*) {
    global g_AutoSlotRearrangePending, g_AutoSlotRearrangeExclude
    if (!g_AutoSlotRearrangePending)
        return
    g_AutoSlotRearrangePending := false
    excludeHwnd := g_AutoSlotRearrangeExclude
    g_AutoSlotRearrangeExclude := 0
    AutoSlot_RearrangeUnderfilled(excludeHwnd)
}

; Heal/fill every underfilled ordinal monitor (same policy as fill-on-close).
AutoSlot_RearrangeUnderfilled(excludeHwnd := 0) {
    if (!AutoSlot_IsEnabled() || MonitorGetCount() <= 1)
        return
    ordinalCount := Min(MonitorGetCount(), AutoSlot_MAX_ORDINAL)
    loop ordinalCount {
        monIdx := AutoSlot_GetMonitorIndexByOrder(A_Index)
        if (!monIdx)
            continue
        others := AutoSlot_OccupancyOnMonitor(monIdx, excludeHwnd)
        if (others.Length = 0) {
            if (AutoSlot_FillCooldownActive(monIdx))
                continue
            AutoSlot_FillMonitorFromBackground(monIdx)
        } else if (others.Length = 1) {
            ; Fill handles SnapPair import / heal / place-freeze / claim cooldown.
            AutoSlot_FillMonitorFromBackground(monIdx)
        }
    }
}

AutoSlot_OnMoveSizeEnd(hWinEventHook, event, hwnd, idObject, idChild, idEventThread, dwmsEventTime) {
    global g_AutoSlotHwndMon
    if (idObject != AutoSlot_OBJID_WINDOW || !hwnd)
        return
    if (!AutoSlot_IsEnabled())
        return
    hwnd := Integer(hwnd)
    if (AutoSlot_SwapQuietActive())
        return
    if (AutoSlot_PairSuppressActive(hwnd))
        return
    if (AutoSlot_PlaceFreezeActive())
        return
    if (!AutoSlot_IsOccupancyCandidate(hwnd))
        return
    oldMon := g_AutoSlotHwndMon.Has(hwnd) ? g_AutoSlotHwndMon[hwnd] : 0
    newMon := AutoSlot_GetHwndMonitorIndex(hwnd)
    AutoSlot_RememberHwndMon(hwnd)
    changedMon := (oldMon >= 1 && newMon >= 1 && oldMon != newMon)
    becameMax := false
    try becameMax := (WinGetMinMax("ahk_id " hwnd) = 1)
    catch
        becameMax := false
    if (changedMon || becameMax)
        AutoSlot_ScheduleRearrange(hwnd)
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
    if (g_AutoSlotPending.Has(hwnd))
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

AutoSlot_IsAhkHostProcess(exe) {
    exe := StrLower(exe)
    return exe = "autohotkey.exe" || exe = "autohotkey64.exe" || exe = "autohotkey32.exe"
        || exe = "autohotkeyux.exe"
}

AutoSlot_IsSameScriptPid(hwnd) {
    if (!hwnd)
        return false
    pid := 0
    try DllCall("GetWindowThreadProcessId", "ptr", hwnd, "uint*", &pid)
    catch
        return false
    if (!pid)
        return false
    return Integer(pid) = Integer(DllCall("GetCurrentProcessId", "uint"))
}

; Handy/ClipAngel overlays, WindowManagement identity, and suite AHK GUIs/prompts.
AutoSlot_IsExcludedExeOrTitle(hwnd) {
    if (!hwnd)
        return false
    try {
        exe := StrLower(WinGetProcessName("ahk_id " hwnd))
    } catch {
        exe := ""
    }
    if (exe = "handy.exe" || exe = "clipangel.exe")
        return true
    if (AutoSlot_IsAhkHostProcess(exe))
        return true
    if (AutoSlot_IsSameScriptPid(hwnd))
        return true
    try {
        class := StrLower(WinGetClass(hwnd))
    } catch {
        class := ""
    }
    if (class = "autohotkeygui")
        return true
    try {
        title := WinGetTitle(hwnd)
    } catch {
        return false
    }
    t := StrLower(title)
    if (InStr(t, "windowmanagement.ahk"))
        return true
    if (InStr(t, "autohotkey") && InStr(class, "ahk"))
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

AutoSlot_MaximizeOnMonitor(hwnd, monIdx, scheduleRearrange := true) {
    if (!hwnd || monIdx < 1 || monIdx > MonitorGetCount())
        return false
    if (!AutoSlot_PrepareHwnd(hwnd))
        return false
    MonitorGet monIdx, &left, &top, &right, &bottom
    AutoSlot_MoveHwndToRect(hwnd, left, top, right, bottom)
    AutoSlot_MaximizeHwnd(hwnd)
    AutoSlot_ActivateHwnd(hwnd)
    if (scheduleRearrange) {
        try AutoSlot_ScheduleRearrange(hwnd)
        catch {
        }
    }
    return true
}

; --- Foreground monitor swap (suite move-to-monitor) -------------------------

; Presentation pane for a window on monIdx ("start"|"end"|"").
AutoSlot_GetHwndPaneOnMonitor(hwnd, monIdx) {
    if (!hwnd || monIdx < 1)
        return ""
    axis := WM_GetSnapSplitAxis(monIdx)
    WM_GetSnapInnerWorkArea(monIdx, &wl, &wt, &wr, &wb)
    if (!WM_GetWindowRectHwnd(hwnd, &l, &t, &r, &b))
        return ""
    pane := ""
    paneSize := 0
    if (WM_ClassifySnapPane(axis, wl, wt, wr, wb, l, t, r, b, &pane, &paneSize) && pane != "")
        return pane
    return ""
}

; Dest/source FG layout: empty | full | single | pair | other.
AutoSlot_ClassifyFgLayout(monIdx, excludeHwnd := 0) {
    rows := AutoSlot_OccupancyOnMonitor(monIdx, excludeHwnd)
    if (rows.Length = 0)
        return { kind: "empty" }
    if (rows.Length = 1) {
        h := rows[1].hwnd
        if (AutoSlot_CompanionAlreadyFilled(h, monIdx))
            return { kind: "full", hwnd: h }
        return { kind: "single", hwnd: h }
    }
    if (rows.Length = 2) {
        a := rows[1].hwnd
        b := rows[2].hwnd
        aPane := AutoSlot_GetHwndPaneOnMonitor(a, monIdx)
        bPane := AutoSlot_GetHwndPaneOnMonitor(b, monIdx)
        if (aPane = "" || bPane = "" || aPane = bPane) {
            if (rows[1].left <= rows[2].left) {
                aPane := "start"
                bPane := "end"
            } else {
                aPane := "end"
                bPane := "start"
            }
        }
        return { kind: "pair", a: a, b: b, aPane: aPane, bPane: bPane }
    }
    return { kind: "other" }
}

; Mover role on source: full | half | halfAlone | other (+ companion when half).
AutoSlot_ClassifyMoverRole(hwnd, sourceMon) {
    if (!hwnd || sourceMon < 1)
        return { role: "other" }
    others := AutoSlot_OccupancyOnMonitor(sourceMon, hwnd)
    filled := AutoSlot_CompanionAlreadyFilled(hwnd, sourceMon)
    if (others.Length = 0) {
        if (filled)
            return { role: "full" }
        return { role: "halfAlone" }
    }
    if (others.Length = 1)
        return { role: "half", companion: others[1].hwnd }
    return { role: "other" }
}

AutoSlot_ApplyPairOnMonitor(a, b, aPane, bPane, monIdx) {
    if (!a || !b || a = b || monIdx < 1)
        return false
    if (aPane = "" || bPane = "" || aPane = bPane) {
        aPane := "start"
        bPane := "end"
    }
    AutoSlot_UnregisterSnapPair(a)
    AutoSlot_UnregisterSnapPair(b)
    if (!WM_PrepareHwndForTile(a) || !WM_PrepareHwndForTile(b))
        return false
    axis := WM_GetSnapSplitAxis(monIdx)
    if (!WM_SnapPairGaplessRects(monIdx, axis, a, aPane, b, bPane))
        return false
    AutoSlot_RegisterSnapPair(a, b)
    AutoSlot_RememberHwndMon(a)
    AutoSlot_RememberHwndMon(b)
    return true
}

; Exchange whole-monitor FG layouts when MoveWinToMonitor would collide.
; Returns true when swap was applied (caller should skip normal move).
AutoSlot_TryForegroundSwap(hwnd, sourceMon, destMon) {
    if (!AutoSlot_IsEnabled() || !hwnd || sourceMon < 1 || destMon < 1 || sourceMon = destMon)
        return false
    if (!DllCall("IsWindow", "ptr", hwnd) || !AutoSlot_IsOccupancyCandidate(hwnd))
        return false

    mover := AutoSlot_ClassifyMoverRole(hwnd, sourceMon)
    dest := AutoSlot_ClassifyFgLayout(destMon)
    role := mover.role
    if (role = "other" || dest.kind = "empty" || dest.kind = "other" || dest.kind = "single")
        return false

    ; Half with a left-behind companion cannot host an incoming pair on source.
    if (role = "half" && dest.kind = "pair")
        return false

    wantSwap := false
    if ((role = "full" || role = "halfAlone") && dest.kind = "pair")
        wantSwap := true
    else if ((role = "full" || role = "halfAlone" || role = "half") && dest.kind = "full")
        wantSwap := true
    else if (role = "full" && dest.kind = "full")
        wantSwap := true
    if (!wantSwap)
        return false

    AutoSlot_BeginSwapQuiet()
    AutoSlot_PairSuppressMark(hwnd)
    companion := 0
    if (role = "half") {
        try companion := Integer(mover.companion)
        catch
            companion := 0
    }
    if (companion)
        AutoSlot_PairSuppressMark(companion)
    if (dest.kind = "pair") {
        AutoSlot_PairSuppressMark(dest.a)
        AutoSlot_PairSuppressMark(dest.b)
    } else if (dest.kind = "full")
        AutoSlot_PairSuppressMark(dest.hwnd)

    AutoSlot_UnregisterSnapPair(hwnd)

    if (dest.kind = "pair") {
        ; Mover takes whole dest; pair keeps relative panes on source.
        if (!AutoSlot_MaximizeOnMonitor(hwnd, destMon, false)) {
            global g_AutoSlotSwapQuietUntil
            g_AutoSlotSwapQuietUntil := 0
            return false
        }
        AutoSlot_RememberHwndMon(hwnd)
        if (!AutoSlot_ApplyPairOnMonitor(dest.a, dest.b, dest.aPane, dest.bPane, sourceMon)) {
            ; Mover already owns dest — do not roll back; avoid fill fighting.
        }
    } else if (dest.kind = "full") {
        other := dest.hwnd
        if (other = hwnd) {
            global g_AutoSlotSwapQuietUntil
            g_AutoSlotSwapQuietUntil := 0
            return false
        }
        AutoSlot_UnregisterSnapPair(other)
        if (!AutoSlot_MaximizeOnMonitor(hwnd, destMon, false)) {
            global g_AutoSlotSwapQuietUntil
            g_AutoSlotSwapQuietUntil := 0
            return false
        }
        AutoSlot_RememberHwndMon(hwnd)
        placedOther := false
        if (companion && companion != other) {
            AutoSlot_UnregisterSnapPair(companion)
            placedOther := AutoSlot_ApplyPairOnMonitor(other, companion, "start", "end", sourceMon)
        }
        if (!placedOther)
            placedOther := AutoSlot_MaximizeOnMonitor(other, sourceMon, false)
        if (placedOther)
            AutoSlot_RememberHwndMon(other)
    } else {
        global g_AutoSlotSwapQuietUntil
        g_AutoSlotSwapQuietUntil := 0
        return false
    }

    AutoSlot_ClaimMonitor(destMon)
    AutoSlot_ClaimMonitor(sourceMon)
    AutoSlot_ActivateHwnd(hwnd)
    try WM_MaybeCenterMouse(hwnd, "autoslot_fg_swap", true)
    catch {
    }
    srcLabel := AutoSlot_OrderForMonitorIndex(sourceMon)
    dstLabel := AutoSlot_OrderForMonitorIndex(destMon)
    if (srcLabel < 1)
        srcLabel := sourceMon
    if (dstLabel < 1)
        dstLabel := destMon
    AutoSlot_Toast("ℹ️ Swapped M" srcLabel " ↔ M" dstLabel)
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

    ; Partner keeps its pane (shrink in place); new always takes the opposite.
    ; Maximized / unclassified → "start" once — do not flip sides after placement.
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
    AutoSlot_RegisterSnapPair(newHwnd, partnerHwnd)
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

; True when hwnd still belongs to a living 50/50 pair (must not be relocated by fill).
AutoSlot_BackgroundCandHasLivingSnapPartner(hwnd) {
    global g_AutoSlotSnapPairs
    if (!hwnd || !g_AutoSlotSnapPairs.Has(hwnd))
        return false
    partner := g_AutoSlotSnapPairs[hwnd]
    return !!(partner && DllCall("IsWindow", "ptr", partner))
}

; True when hwnd's home monitor has another window in F11 fullscreen (hwnd is covered there).
AutoSlot_BackgroundCandCoveredByF11(hwnd) {
    global g_AutoSlotHwndMon
    if (!hwnd)
        return false
    monIdx := 0
    if (g_AutoSlotHwndMon.Has(hwnd))
        monIdx := g_AutoSlotHwndMon[hwnd]
    if (monIdx < 1)
        monIdx := AutoSlot_GetHwndMonitorIndex(hwnd)
    if (monIdx < 1)
        return false
    try {
        for other in WinGetList() {
            if (!other || other = hwnd)
                continue
            if (AutoSlot_GetHwndMonitorIndex(other) != monIdx)
                continue
            try {
                if (WM_WindowIsF11Fullscreen(other))
                    return true
            } catch {
            }
        }
    } catch {
    }
    return false
}

; excludeExtra: optional hwnd or Map/Array of hwnds already chosen (for dual-slot empty fill).
AutoSlot_PickBackgroundCandidate(monIdx, occupancyRows, excludeExtra := 0) {
    occupied := Map()
    for row in occupancyRows {
        if (row.hwnd)
            occupied[row.hwnd] := true
    }
    if (IsObject(excludeExtra)) {
        try {
            for h, _ in excludeExtra
                if (h)
                    occupied[Integer(h)] := true
        } catch {
            for h in excludeExtra
                if (h)
                    occupied[Integer(h)] := true
        }
    } else if (excludeExtra) {
        occupied[Integer(excludeExtra)] := true
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
        ; Do not steal F11-covered / still-paired 50/50 companions into another slot.
        if (AutoSlot_BackgroundCandHasLivingSnapPartner(hwnd))
            continue
        if (AutoSlot_BackgroundCandCoveredByF11(hwnd))
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
AutoSlot_PlaceFreezeActive() {
    global g_AutoSlotPlaceFreezeUntil
    return g_AutoSlotPlaceFreezeUntil > 0 && A_TickCount < g_AutoSlotPlaceFreezeUntil
}

AutoSlot_BeginPlaceFreeze() {
    global g_AutoSlotPlaceFreezeUntil
    g_AutoSlotPlaceFreezeUntil := A_TickCount + AutoSlot_RECENT_MS
}

; Promote background into free capacity: empty → up to two 50/50 (or one max);
; half → SnapPair one background with residual (else heal). No undo modal.
AutoSlot_FillMonitorFromBackground(monIdx) {
    global g_AutoSlotRecent, g_AutoSlotUndo
    if (monIdx < 1 || monIdx > MonitorGetCount() || MonitorGetCount() <= 1)
        return "noop"
    others := AutoSlot_OccupancyOnMonitor(monIdx)
    if (others.Length >= 2)
        return "stale"

    order := AutoSlot_OrderForMonitorIndex(monIdx)
    label := order > 0 ? order : monIdx

    ; Place freeze / claim cooldown: companion heal only (no background import).
    blockImport := AutoSlot_PlaceFreezeActive() || AutoSlot_FillCooldownActive(monIdx)

    if (others.Length = 1) {
        residual := others[1].hwnd
        if (AutoSlot_CompanionAlreadyFilled(residual, monIdx))
            return "ok"
        if (blockImport)
            return AutoSlot_HealLoneCompanion(monIdx) ? "ok" : "noop"
        cand := AutoSlot_PickBackgroundCandidate(monIdx, others)
        if (!cand)
            return AutoSlot_HealLoneCompanion(monIdx) ? "ok" : "noop"
        g_AutoSlotRecent[cand] := A_TickCount
        AutoSlot_PruneRecent()
        pane := AutoSlot_SnapPair(cand, residual, monIdx, true)
        g_AutoSlotUndo := 0
        if (pane = "")
            return AutoSlot_HealLoneCompanion(monIdx) ? "ok" : "noop"
        AutoSlot_RememberHwndMon(cand)
        AutoSlot_RememberHwndMon(residual)
        AutoSlot_Toast("ℹ️ Slot filled → M" label " (50/50)")
        return "ok"
    }

    ; Empty monitor — fill both slots when two candidates exist.
    if (blockImport)
        return "noop"
    cand1 := AutoSlot_PickBackgroundCandidate(monIdx, others)
    if (!cand1)
        return "noop"
    cand2 := AutoSlot_PickBackgroundCandidate(monIdx, others, cand1)
    if (cand2) {
        g_AutoSlotRecent[cand1] := A_TickCount
        g_AutoSlotRecent[cand2] := A_TickCount
        AutoSlot_PruneRecent()
        pane := AutoSlot_SnapPair(cand1, cand2, monIdx, true)
        g_AutoSlotUndo := 0
        if (pane != "") {
            AutoSlot_RememberHwndMon(cand1)
            AutoSlot_RememberHwndMon(cand2)
            AutoSlot_Toast("ℹ️ Slot filled → M" label " (50/50)")
            return "ok"
        }
        ; Snap failed — fall through to single maximize of cand1.
    }
    g_AutoSlotRecent[cand1] := A_TickCount
    AutoSlot_PruneRecent()
    if (!AutoSlot_MaximizeOnMonitor(cand1, monIdx))
        return "noop"
    AutoSlot_RememberHwndMon(cand1)
    ; Still one free half-slot after maximize — fill it now (same pass; claim cooldown
    ; would otherwise block rearrange from importing the second background).
    if (!cand2)
        cand2 := AutoSlot_PickBackgroundCandidate(monIdx, [{ hwnd: cand1 }], cand1)
    if (cand2) {
        g_AutoSlotRecent[cand2] := A_TickCount
        AutoSlot_PruneRecent()
        pane := AutoSlot_SnapPair(cand2, cand1, monIdx, true)
        g_AutoSlotUndo := 0
        if (pane != "") {
            AutoSlot_RememberHwndMon(cand2)
            AutoSlot_Toast("ℹ️ Slot filled → M" label " (50/50)")
            return "ok"
        }
    }
    AutoSlot_Toast("ℹ️ Slot filled → M" label " (maximized)")
    return "ok"
}

; --- Placement ---------------------------------------------------------------

; First ordinal monitor with exactly one occupant (excl. hwnd). One maximized
; window counts as half-full — the other half-slot is still available.
AutoSlot_FindHalfFullMonitor(excludeHwnd := 0) {
    ordinalCount := Min(MonitorGetCount(), AutoSlot_MAX_ORDINAL)
    loop ordinalCount {
        order := A_Index
        monIdx := AutoSlot_GetMonitorIndexByOrder(order)
        if (!monIdx)
            continue
        others := AutoSlot_OccupancyOnMonitor(monIdx, excludeHwnd)
        if (others.Length = 1)
            return { order: order, monIdx: monIdx, partner: others[1].hwnd }
    }
    return 0
}

AutoSlot_TrySnapOnMonitor(hwnd, monIdx, orderLabel := 0) {
    others := AutoSlot_OccupancyOnMonitor(monIdx, hwnd)
    if (others.Length != 1)
        return false
    partner := others[1].hwnd
    pane := AutoSlot_SnapPair(hwnd, partner, monIdx)
    if (pane = "")
        return false
    AutoSlot_ClaimMonitor(monIdx)
    if (orderLabel < 1)
        orderLabel := AutoSlot_OrderForMonitorIndex(monIdx)
    label := orderLabel > 0 ? orderLabel : monIdx
    AutoSlot_ShowUndoModal("M" label " " pane)
    return true
}

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
    AutoSlot_BeginPlaceFreeze()

    ; Empty-monitor-first: never 50/50 on an occupied screen when any ordinal is empty.
    emptyOrder := 0
    emptyMon := 0

    loop ordinalCount {
        order := A_Index
        monIdx := AutoSlot_GetMonitorIndexByOrder(order)
        if (!monIdx)
            continue
        others := AutoSlot_OccupancyOnMonitor(monIdx, hwnd)
        if (others.Length = 0) {
            emptyOrder := order
            emptyMon := monIdx
            break
        }
    }

    if (emptyMon) {
        if (AutoSlot_MaximizeOnMonitor(hwnd, emptyMon)) {
            AutoSlot_ClaimMonitor(emptyMon)
            msg := "ℹ️ Auto-slotted → M" emptyOrder " (maximized)"
        }
        AutoSlot_Toast(msg)
        return
    }

    ; Prefer partitioning the origin slot when it has exactly one resident (incl. maximized).
    originMon := AutoSlot_GetHwndMonitorIndex(hwnd)
    if (originMon >= 1 && AutoSlot_TrySnapOnMonitor(hwnd, originMon))
        return

    ; Half-slots remain open on any monitor with a single window (maximized = 1 of 2 slots).
    half := AutoSlot_FindHalfFullMonitor(hwnd)
    if (IsObject(half) && AutoSlot_TrySnapOnMonitor(hwnd, half.monIdx, half.order))
        return

    if (AutoSlot_MaximizeInPlace(hwnd))
        msg := "ℹ️ Grid full — maximized"
    AutoSlot_Toast(msg)
}

; Auto-execute when #included.
AutoSlot_Init()