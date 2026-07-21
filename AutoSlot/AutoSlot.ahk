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
; Rearrange-on-move / minimize (same fill/heal rules, no visible reshuffle):
;   MOVESIZEEND / MINIMIZEEND / suite leave / MaximizeOnMonitor → debounced RearrangeUnderfilled.
;
; Foreground monitor swap (suite MoveWinToMonitor when AutoSlot ON):
;   Full/halfAlone/half↔pair; full/halfAlone/half↔full|single.
;   After swap: 2s ShowWithKeys [F] replaces (minimizes + fill-skip displaced).
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
global g_AutoSlotMinHook := 0
global g_AutoSlotMinHookCb := 0
global g_AutoSlotRearrangePending := false
global g_AutoSlotRearrangeExclude := 0
global g_AutoSlotSwapQuietUntil := 0
global g_AutoSlotSwapDisplaced := []
global g_AutoSlotSwapMoverHwnd := 0
global g_AutoSlotReplaceSkip := Map()
global g_AutoSlotJustRestored := Map()
global g_AutoSlotLastToastMsg := ""
global g_AutoSlotLastToastTick := 0
global g_AutoSlotWasF11 := Map()          ; snap-pair hwnd → true while (or after) F11 fullscreen
global g_AutoSlotF11RestorePending := Map()

AutoSlot_EVENT_OBJECT_DESTROY := 0x8001
AutoSlot_EVENT_OBJECT_SHOW := 0x8002
AutoSlot_EVENT_OBJECT_LOCATIONCHANGE := 0x800B
AutoSlot_EVENT_SYSTEM_MOVESIZEEND := 0x000B
AutoSlot_EVENT_SYSTEM_MINIMIZESTART := 0x0016
AutoSlot_EVENT_SYSTEM_MINIMIZEEND := 0x0017
AutoSlot_OBJID_WINDOW := 0
AutoSlot_DEBOUNCE_MS := 250
AutoSlot_RECENT_MS := 4000
AutoSlot_FILL_COOLDOWN_MS := 1500
AutoSlot_FILL_RETRY_MS := 400
AutoSlot_DESTROY_DEDUP_MS := 250
AutoSlot_PAIR_MAX_DEBOUNCE_MS := 120
AutoSlot_PAIR_SUPPRESS_MS := 500
AutoSlot_REARRANGE_MS := 350
AutoSlot_SWAP_QUIET_MS := 2500
AutoSlot_SWAP_MODAL_MS := 2000
AutoSlot_SWAP_PAIR_SUPPRESS_MS := 3500
AutoSlot_REPLACE_SKIP_TTL_MS := 600000
AutoSlot_RESTORE_GUARD_MS := 4000   ; MINIMIZEEND → suppress SHOW-triggered auto-place for this long
AutoSlot_TOAST_DEBOUNCE_MS := 4000  ; identical toast text — avoid hide-timer reset spam
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
    global g_AutoSlotMinHook, g_AutoSlotMinHookCb
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

    ; Minimize start → clear snap pair, heal leftover companion, rearrange/fill.
    g_AutoSlotMinHookCb := CallbackCreate(AutoSlot_OnMinimize, "F", 7)
    g_AutoSlotMinHook := DllCall("user32\SetWinEventHook",
        "UInt", AutoSlot_EVENT_SYSTEM_MINIMIZESTART,
        "UInt", AutoSlot_EVENT_SYSTEM_MINIMIZEEND,
        "Ptr", 0,
        "Ptr", g_AutoSlotMinHookCb,
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
    global g_AutoSlotHwndMon, g_AutoSlotLastDestroyHwnd, g_AutoSlotLastDestroyTick, g_AutoSlotSnapPairs
    if (!hwnd)
        return

    ; Capture living snap partner BEFORE unregister — dead HWND often has no mon cache,
    ; and fill must still heal the leftover half.
    partner := 0
    partnerMon := 0
    if (g_AutoSlotSnapPairs.Has(hwnd)) {
        partner := g_AutoSlotSnapPairs[hwnd]
        if (partner && partner != hwnd && DllCall("IsWindow", "ptr", partner)) {
            partnerMon := AutoSlot_GetHwndMonitorIndex(partner)
            if (partnerMon < 1 && g_AutoSlotHwndMon.Has(partner))
                partnerMon := g_AutoSlotHwndMon[partner]
        } else
            partner := 0
    }

    AutoSlot_UnregisterSnapPair(hwnd)

    if (MonitorGetCount() <= 1)
        return

    cached := g_AutoSlotHwndMon.Has(hwnd) ? g_AutoSlotHwndMon[hwnd] : 0
    alive := !!DllCall("IsWindow", "ptr", hwnd)

    if (!fromShell) {
        if (hwnd = g_AutoSlotLastDestroyHwnd && A_TickCount - g_AutoSlotLastDestroyTick < AutoSlot_DESTROY_DEDUP_MS)
            return
        if (!cached && !partner && !AutoSlot_DestroyLooksInteresting(hwnd))
            return
    } else {
        ; Shell primary: skip chrome noise when alive; dead HWNDs need cache or snap partner.
        if (!cached && !partner) {
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
    if (monIdx < 1 && partnerMon >= 1)
        monIdx := partnerMon
    if (g_AutoSlotHwndMon.Has(hwnd))
        g_AutoSlotHwndMon.Delete(hwnd)
    if (monIdx < 1)
        return

    g_AutoSlotLastDestroyHwnd := hwnd
    g_AutoSlotLastDestroyTick := A_TickCount
    ; Closing one half of a registered pair → maximize that partner (bypass occupancy races).
    if (partner) {
        p := partner
        m := monIdx
        SetTimer(() => AutoSlot_HealKnownCompanion(p, m), -AutoSlot_DEBOUNCE_MS)
        ; Second pass after destroy HWND clears from occupancy lists.
        SetTimer(() => AutoSlot_HealKnownCompanion(p, m), -AutoSlot_FILL_RETRY_MS)
    }
    AutoSlot_ScheduleHealOnly(monIdx)
    AutoSlot_ScheduleFill(monIdx)
    ; Late heal for unregistered 50/50 (no snap pair) once zombie HWNDs drop out.
    lateMon := monIdx
    SetTimer(() => AutoSlot_HealLoneCompanion(lateMon), -(AutoSlot_FILL_RETRY_MS + 200))
}

; Maximize a known leftover companion after its snap-pair partner closed.
AutoSlot_HealKnownCompanion(companion, monIdx := 0) {
    if (!companion || !DllCall("IsWindow", "ptr", companion))
        return false
    if (AutoSlot_IsClipAngelHwnd(companion))
        return false
    if (monIdx < 1)
        monIdx := AutoSlot_GetHwndMonitorIndex(companion)
    if (monIdx < 1)
        return false
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
    if (!healed) {
        try {
            WinMaximize("ahk_id " companion)
            healed := true
        } catch {
        }
    }
    if (healed) {
        AutoSlot_RememberHwndMon(companion)
        ; Suppress MoveSizeEnd/paired-max rearrange so heal does not re-toast in a loop.
        AutoSlot_PairSuppressMark(companion, AutoSlot_RECENT_MS)
        AutoSlot_ClaimMonitor(monIdx)
        order := AutoSlot_OrderForMonitorIndex(monIdx)
        label := order > 0 ? order : monIdx
        AutoSlot_Toast("ℹ️ Companion maximized → M" label)
    }
    return healed
}

; --- 50/50 pair registry (maximize either side → maximize companion) ----------

AutoSlot_RegisterSnapPair(a, b) {
    global g_AutoSlotSnapPairs
    if (!a || !b || a = b)
        return
    ; Clear stale reverse links (e.g. C→A) before re-pairing A↔B.
    if (g_AutoSlotSnapPairs.Has(a)) {
        old := g_AutoSlotSnapPairs[a]
        if (old && old != b && g_AutoSlotSnapPairs.Has(old) && g_AutoSlotSnapPairs[old] = a)
            g_AutoSlotSnapPairs.Delete(old)
    }
    if (g_AutoSlotSnapPairs.Has(b)) {
        old := g_AutoSlotSnapPairs[b]
        if (old && old != a && g_AutoSlotSnapPairs.Has(old) && g_AutoSlotSnapPairs[old] = b)
            g_AutoSlotSnapPairs.Delete(old)
    }
    g_AutoSlotSnapPairs[a] := b
    g_AutoSlotSnapPairs[b] := a
    AutoSlot_RememberHwndMon(a)
    AutoSlot_RememberHwndMon(b)
}

AutoSlot_UnregisterSnapPair(hwnd) {
    global g_AutoSlotSnapPairs, g_AutoSlotWasF11, g_AutoSlotF11RestorePending
    if (!hwnd || !g_AutoSlotSnapPairs.Has(hwnd))
        return
    partner := g_AutoSlotSnapPairs[hwnd]
    g_AutoSlotSnapPairs.Delete(hwnd)
    if (partner && g_AutoSlotSnapPairs.Has(partner) && g_AutoSlotSnapPairs[partner] = hwnd)
        g_AutoSlotSnapPairs.Delete(partner)
    if (g_AutoSlotWasF11.Has(hwnd))
        g_AutoSlotWasF11.Delete(hwnd)
    if (g_AutoSlotF11RestorePending.Has(hwnd))
        g_AutoSlotF11RestorePending.Delete(hwnd)
    if (partner) {
        if (g_AutoSlotWasF11.Has(partner))
            g_AutoSlotWasF11.Delete(partner)
        if (g_AutoSlotF11RestorePending.Has(partner))
            g_AutoSlotF11RestorePending.Delete(partner)
    }
}

AutoSlot_PairSuppressActive(hwnd) {
    global g_AutoSlotMaxCompanionSuppress
    if (!hwnd || !g_AutoSlotMaxCompanionSuppress.Has(hwnd))
        return false
    return A_TickCount < g_AutoSlotMaxCompanionSuppress[hwnd]
}

AutoSlot_PairSuppressMark(hwnd, ms := 0) {
    global g_AutoSlotMaxCompanionSuppress
    if (!hwnd)
        return
    if (ms < 1)
        ms := AutoSlot_PAIR_SUPPRESS_MS
    g_AutoSlotMaxCompanionSuppress[hwnd] := A_TickCount + ms
}

AutoSlot_OnLocationChange(hWinEventHook, event, hwnd, idObject, idChild, idEventThread, dwmsEventTime) {
    if (idObject != AutoSlot_OBJID_WINDOW || !hwnd)
        return
    if (!AutoSlot_IsEnabled())
        return
    hwnd := Integer(hwnd)
    global g_AutoSlotSnapPairs, g_AutoSlotPairMaxPending, g_AutoSlotWasF11, g_AutoSlotF11RestorePending
    if (!g_AutoSlotSnapPairs.Has(hwnd))
        return
    if (AutoSlot_PairSuppressActive(hwnd))
        return

    ; Track F11 on snap-pair members: enter → remember; exit → restore 50/50 (not 2-slot + bg companion).
    isF11 := false
    try isF11 := !!WM_WindowIsF11Fullscreen(hwnd)
    catch
        isF11 := false
    if (isF11) {
        g_AutoSlotWasF11[hwnd] := true
        return
    }
    if (g_AutoSlotWasF11.Has(hwnd)) {
        if (g_AutoSlotF11RestorePending.Has(hwnd))
            return
        g_AutoSlotF11RestorePending[hwnd] := true
        SetTimer(() => AutoSlot_ProcessF11ExitRestore(hwnd), -AutoSlot_PAIR_MAX_DEBOUNCE_MS)
        return
    }

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

AutoSlot_ProcessF11ExitRestore(hwnd) {
    global g_AutoSlotF11RestorePending, g_AutoSlotWasF11
    if (g_AutoSlotF11RestorePending.Has(hwnd))
        g_AutoSlotF11RestorePending.Delete(hwnd)
    if (g_AutoSlotWasF11.Has(hwnd))
        g_AutoSlotWasF11.Delete(hwnd)
    if (AutoSlot_SwapQuietActive())
        return
    AutoSlot_RestoreSnapPairAfterF11(hwnd)
}

; After F11 on one half of a 50/50 pair: snap both back instead of leaving one maximized
; (2 slots) with the companion stuck in the background.
AutoSlot_RestoreSnapPairAfterF11(hwnd) {
    global g_AutoSlotSnapPairs
    if (!AutoSlot_IsEnabled() || !hwnd)
        return false
    if (!g_AutoSlotSnapPairs.Has(hwnd))
        return false
    partner := g_AutoSlotSnapPairs[hwnd]
    if (!partner || !DllCall("IsWindow", "ptr", partner)) {
        AutoSlot_UnregisterSnapPair(hwnd)
        return false
    }
    ; Still F11 (transient event) — keep the mark and wait.
    try {
        if (WM_WindowIsF11Fullscreen(hwnd)) {
            global g_AutoSlotWasF11
            g_AutoSlotWasF11[hwnd] := true
            return false
        }
    } catch {
    }
    monIdx := AutoSlot_GetHwndMonitorIndex(hwnd)
    if (monIdx < 1)
        monIdx := AutoSlot_GetHwndMonitorIndex(partner)
    if (monIdx < 1)
        return false

    AutoSlot_PairSuppressMark(hwnd, AutoSlot_RECENT_MS)
    AutoSlot_PairSuppressMark(partner, AutoSlot_RECENT_MS)
    AutoSlot_ClearPairMaxPending(hwnd)
    AutoSlot_ClearPairMaxPending(partner)
    pane := AutoSlot_SnapPair(hwnd, partner, monIdx, true)
    if (pane = "")
        return false
    AutoSlot_RememberHwndMon(hwnd)
    AutoSlot_RememberHwndMon(partner)
    AutoSlot_ClaimMonitor(monIdx)
    order := AutoSlot_OrderForMonitorIndex(monIdx)
    label := order > 0 ? order : monIdx
    AutoSlot_Toast("ℹ️ Restored 50/50 → M" label)
    return true
}

AutoSlot_ProcessPairedMaximizePending(hwnd) {
    global g_AutoSlotPairMaxPending
    if (g_AutoSlotPairMaxPending.Has(hwnd))
        g_AutoSlotPairMaxPending.Delete(hwnd)
    if (AutoSlot_SwapQuietActive())
        return
    AutoSlot_OnPairedMaximize(hwnd)
}

; When hwnd is (or became) maximized, maximize its registered 50/50 companion.
AutoSlot_OnPairedMaximize(hwnd) {
    global g_AutoSlotSnapPairs, g_AutoSlotWasF11, g_AutoSlotF11RestorePending
    if (!AutoSlot_IsEnabled() || !hwnd)
        return false
    if (AutoSlot_SwapQuietActive())
        return false
    if (AutoSlot_PairSuppressActive(hwnd))
        return false
    ; F11 enter/exit owns this hwnd — do not maximize the companion on top of fullscreen.
    if (g_AutoSlotWasF11.Has(hwnd) || g_AutoSlotF11RestorePending.Has(hwnd))
        return false
    try {
        if (WM_WindowIsF11Fullscreen(hwnd)) {
            g_AutoSlotWasF11[hwnd] := true
            return false
        }
    } catch {
    }
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

AutoSlot_ClearPairMaxPending(hwnd) {
    global g_AutoSlotPairMaxPending
    if (hwnd && g_AutoSlotPairMaxPending.Has(hwnd))
        g_AutoSlotPairMaxPending.Delete(hwnd)
}

; --- Rearrange underfilled slots after a move --------------------------------

AutoSlot_BeginSwapQuiet(ms := 0) {
    global g_AutoSlotSwapQuietUntil
    if (ms < 1)
        ms := AutoSlot_SWAP_QUIET_MS
    g_AutoSlotSwapQuietUntil := A_TickCount + ms
    ; Catch-up rearrange after quiet — requests during quiet are otherwise lost.
    SetTimer(AutoSlot_PostQuietRearrange, -(ms + AutoSlot_REARRANGE_MS))
}

AutoSlot_SwapQuietActive() {
    global g_AutoSlotSwapQuietUntil
    return g_AutoSlotSwapQuietUntil > 0 && A_TickCount < g_AutoSlotSwapQuietUntil
}

; Fired after SwapQuiet expires; heals/fills underfilled ordinals missed during quiet.
AutoSlot_PostQuietRearrange(*) {
    if (!AutoSlot_IsEnabled() || MonitorGetCount() <= 1)
        return
    if (AutoSlot_SwapQuietActive())
        return
    AutoSlot_RearrangeUnderfilled()
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
    ; Quiet started after this timer was armed — drop; PostQuietRearrange will catch up.
    if (AutoSlot_SwapQuietActive()) {
        g_AutoSlotRearrangePending := false
        g_AutoSlotRearrangeExclude := 0
        return
    }
    g_AutoSlotRearrangePending := false
    excludeHwnd := g_AutoSlotRearrangeExclude
    g_AutoSlotRearrangeExclude := 0
    AutoSlot_RearrangeUnderfilled(excludeHwnd)
}

; Heal/fill every underfilled ordinal monitor (same policy as fill-on-close).
; Counts non-filled occupants only — maximized-behind must not block heal/fill.
; Lone maximized = half-full (free half-slot for background SnapPair).
AutoSlot_RearrangeUnderfilled(excludeHwnd := 0) {
    if (!AutoSlot_IsEnabled() || MonitorGetCount() <= 1)
        return
    ordinalCount := Min(MonitorGetCount(), AutoSlot_MAX_ORDINAL)
    loop ordinalCount {
        monIdx := AutoSlot_GetMonitorIndexByOrder(A_Index)
        if (!monIdx)
            continue
        part := AutoSlot_PartitionOccupancy(monIdx, excludeHwnd)
        if (part.nonFilled.Length >= 2)
            continue
        if (part.nonFilled.Length = 0 && part.filled.Length >= 2)
            continue  ; no free slot
        if (part.nonFilled.Length = 0 && part.filled.Length = 0) {
            if (AutoSlot_FillCooldownActive(monIdx))
                continue
            AutoSlot_FillMonitorFromBackground(monIdx)
        } else if (part.nonFilled.Length = 1 || (part.nonFilled.Length = 0 && part.filled.Length = 1)) {
            ; Half (unfilled) or lone maximized — Fill SnapPairs / heals.
            AutoSlot_FillMonitorFromBackground(monIdx)
        }
    }
}

; Ctrl+Alt+Win+Y when AutoSlot ON: fill underfilled ordinal slots from background only
; (never relocates windows already occupying slots). Returns { ok, message }.
; Prefer empty monitors before splitting a lone maximized (half-slot).
AutoSlot_RunTileBackground() {
    if (!AutoSlot_IsEnabled())
        return { ok: false, message: "AutoSlot is OFF" }
    if (MonitorGetCount() <= 1)
        return { ok: false, message: "Need more than one monitor for slot fill" }
    filled := 0
    healed := 0
    skippedFull := 0
    ordinalCount := Min(MonitorGetCount(), AutoSlot_MAX_ORDINAL)
    emptyMons := []
    halfMons := []
    loop ordinalCount {
        monIdx := AutoSlot_GetMonitorIndexByOrder(A_Index)
        if (!monIdx)
            continue
        part := AutoSlot_PartitionOccupancy(monIdx)
        if (part.nonFilled.Length >= 2 || (part.nonFilled.Length = 0 && part.filled.Length >= 2)) {
            skippedFull++
            continue
        }
        if (part.nonFilled.Length = 0 && part.filled.Length = 0)
            emptyMons.Push(monIdx)
        else
            halfMons.Push(monIdx)  ; nonFilled=1 or lone maximized
    }
    for monIdx in emptyMons {
        beforeNon := 0
        beforeFilled := 0
        result := AutoSlot_FillMonitorFromBackground(monIdx, true)
        if (result != "ok")
            continue
        filled++
    }
    for monIdx in halfMons {
        part := AutoSlot_PartitionOccupancy(monIdx)
        beforeNon := part.nonFilled.Length
        beforeFilled := part.filled.Length
        result := AutoSlot_FillMonitorFromBackground(monIdx, true)
        if (result != "ok")
            continue
        after := AutoSlot_PartitionOccupancy(monIdx)
        if (beforeNon = 1 && after.nonFilled.Length = 0 && after.filled.Length > beforeFilled)
            healed++
        else
            filled++
    }
    total := filled + healed
    if (total = 0) {
        if (skippedFull >= ordinalCount && emptyMons.Length = 0 && halfMons.Length = 0)
            return { ok: false, message: "All ordinal slots full — nothing to fill" }
        return { ok: false, noBackground: true, message: "No background windows to fill free slots" }
    }
    msg := "Filled " total " slot group(s) from background"
    if (healed > 0)
        msg .= " (" healed " companion heal)"
    msg .= " — slotted windows left in place"
    return { ok: true, message: msg, filled: filled, healed: healed }
}

AutoSlot_OnMoveSizeEnd(hWinEventHook, event, hwnd, idObject, idChild, idEventThread, dwmsEventTime) {
    global g_AutoSlotHwndMon, g_AutoSlotSnapPairs, g_AutoSlotWasF11, g_AutoSlotF11RestorePending
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
    ; Clip Angel maximize/move must not trigger fill — excludeHwnd=ClipAngel made the
    ; monitor look empty and SnapPair/Maximize imported backgrounds on top of it.
    if (AutoSlot_IsClipAngelHwnd(hwnd))
        return

    ; Backup F11-exit detection when LOCATIONCHANGE did not schedule restore.
    if (g_AutoSlotSnapPairs.Has(hwnd)) {
        isF11 := false
        try isF11 := !!WM_WindowIsF11Fullscreen(hwnd)
        catch
            isF11 := false
        if (isF11) {
            g_AutoSlotWasF11[hwnd] := true
            return
        }
        if (g_AutoSlotWasF11.Has(hwnd) && !g_AutoSlotF11RestorePending.Has(hwnd)) {
            g_AutoSlotF11RestorePending[hwnd] := true
            SetTimer(() => AutoSlot_ProcessF11ExitRestore(hwnd), -AutoSlot_PAIR_MAX_DEBOUNCE_MS)
            return
        }
    }

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

; On minimize start: clear pair registry, heal leftover companion, then rearrange/fill.
AutoSlot_OnMinimize(hWinEventHook, event, hwnd, idObject, idChild, idEventThread, dwmsEventTime) {
    global g_AutoSlotHwndMon, g_AutoSlotSnapPairs
    if (idObject != AutoSlot_OBJID_WINDOW || !hwnd)
        return
    if (!AutoSlot_IsEnabled())
        return
    hwnd := Integer(hwnd)
    ; MINIMIZEEND fires when a window is restored from the taskbar, not when it minimizes.
    ; Arm a short guard so the subsequent EVENT_OBJECT_SHOW does not auto-place the window.
    ; Also clear any stale replace-skip so the window is available as a Y/fill candidate
    ; next time it is minimized again.
    if (event = AutoSlot_EVENT_SYSTEM_MINIMIZEEND) {
        global g_AutoSlotJustRestored, g_AutoSlotReplaceSkip
        g_AutoSlotJustRestored[hwnd] := A_TickCount
        if (g_AutoSlotReplaceSkip.Has(hwnd))
            g_AutoSlotReplaceSkip.Delete(hwnd)
        return
    }
    cached := g_AutoSlotHwndMon.Has(hwnd) ? g_AutoSlotHwndMon[hwnd] : 0
    partner := g_AutoSlotSnapPairs.Has(hwnd) ? g_AutoSlotSnapPairs[hwnd] : 0
    if (!cached && !AutoSlot_IsMinimizeRearrangeCandidate(hwnd))
        return
    monIdx := cached >= 1 ? cached : 0
    partnerHwnd := partner
    ; Always break the 50/50 link when one side minimizes — place freeze / swap quiet
    ; must not leave a stale pair that blocks later Y/fill (skipPartner).
    AutoSlot_UnregisterSnapPair(hwnd)
    AutoSlot_ForgetHwndMon(hwnd)
    ; User-initiated minimize clears the [F]-swap replace-skip so the window becomes
    ; available again as a background candidate for Ctrl+Alt+Win+Y.
    global g_AutoSlotReplaceSkip
    if (g_AutoSlotReplaceSkip.Has(hwnd))
        g_AutoSlotReplaceSkip.Delete(hwnd)
    if (partnerHwnd && monIdx >= 1) {
        p := partnerHwnd
        m := monIdx
        SetTimer(() => AutoSlot_HealKnownCompanion(p, m), -AutoSlot_DEBOUNCE_MS)
    }
    if (AutoSlot_SwapQuietActive())
        return
    ; Rearrange may import background; skip only while swap-quiet. Place freeze still
    ; allows heal above; Y uses forceImport and does not depend on this path.
    AutoSlot_ScheduleRearrange(hwnd)
}

; Like IsOccupancyCandidate but allows iconic (minMax = -1).
AutoSlot_IsMinimizeRearrangeCandidate(hwnd) {
    if (!hwnd || !DllCall("IsWindow", "ptr", hwnd))
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
        if (AutoSlot_IsOccupancySkipExeOrTitle(hwnd))
            return false
    } catch {
        return false
    }
    return true
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
        MonitorGetWorkArea monIdx, &wl, &wt, &wr, &wb
        rect := Buffer(16, 0)
        if !DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect)
            return false
        l := NumGet(rect, 0, "int"), t := NumGet(rect, 4, "int")
        r := NumGet(rect, 8, "int"), b := NumGet(rect, 12, "int")
        workW := wr - wl
        workH := wb - wt
        if (workW < 1 || workH < 1)
            return false
        w := r - l
        h := b - t
        ; OS-maximized: trust flag unless the window is clearly a half-pane (stale max bit).
        try {
            if (WinGetMinMax("ahk_id " hwnd) = 1) {
                if (w >= Round(workW * 0.60) && h >= Round(workH * 0.60))
                    return true
            }
        } catch {
        }
        ; Exact work-area match (±8).
        if (Abs(l - wl) <= 8 && Abs(t - wt) <= 8 && Abs(r - wr) <= 8 && Abs(b - wb) <= 8)
            return true
        ; Overlap with work area ≥90% both axes (handles chrome/shadow outside work rect).
        ovW := Min(r, wr) - Max(l, wl)
        ovH := Min(b, wb) - Max(t, wt)
        if (ovW >= Round(workW * 0.90) && ovH >= Round(workH * 0.90))
            return true
        ; Near-full outer size aligned near top-left.
        return (w >= Round(workW * 0.90) && h >= Round(workH * 0.90)
        && Abs(l - wl) <= 24 && Abs(t - wt) <= 24)
    } catch {
        return false
    }
}

AutoSlot_Schedule(hwnd) {
    global g_AutoSlotPending, g_AutoSlotRecent, g_AutoSlotHwndMon, g_AutoSlotJustRestored
    if (!AutoSlot_IsEnabled())
        return
    if (!hwnd || MonitorGetCount() <= 1)
        return
    ; Already tracked — ignore SHOW spam from activation, not a new window.
    if (g_AutoSlotHwndMon.Has(hwnd))
        return
    ; Recently restored from taskbar — the SHOW event is from the restore animation,
    ; not a new window opening. Skip auto-placement for the guard duration.
    if (g_AutoSlotJustRestored.Has(hwnd)) {
        if (A_TickCount - g_AutoSlotJustRestored[hwnd] < AutoSlot_RESTORE_GUARD_MS)
            return
        g_AutoSlotJustRestored.Delete(hwnd)
    }
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

; Handy overlays, WindowManagement identity, suite AHK GUIs/prompts, ClipAngel,
; and Win+Shift+S screen clip / Snipping Tool (must not be auto-slotted or resized).
; ClipAngel is fully excluded — never auto-slotted / 50/50'd and never counted as occupancy.
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
    ; Win+Shift+S region capture overlay + Snipping Tool / Screen Sketch editors.
    if (exe = "screenclippinghost.exe" || exe = "snippingtool.exe" || exe = "screensketch.exe")
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
    ; Microsoft Teams screen-sharing control bar (floating toolbar) — never auto-slot / resize.
    if (InStr(t, "sharing control bar |"))
        return true
    if (InStr(t, "windowmanagement.ahk"))
        return true
    if (InStr(t, "autohotkey") && InStr(class, "ahk"))
        return true
    return false
}

; Occupancy skip: Clip Angel is fully ignored — never counts as a slot occupant, so
; rearrange/fill arranges real windows as if it is not there.
AutoSlot_IsOccupancySkipExeOrTitle(hwnd) {
    return AutoSlot_IsExcludedExeOrTitle(hwnd)
}

AutoSlot_IsClipAngelHwnd(hwnd) {
    if (!hwnd)
        return false
    try return StrLower(WinGetProcessName("ahk_id " hwnd)) = "clipangel.exe"
    catch
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
    if (!hwnd)
        return false
    if (excludeHwnd && hwnd = excludeHwnd)
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
        if (AutoSlot_IsOccupancySkipExeOrTitle(hwnd))
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
        ; Prefer window center (same as WM_GetHwndMonitorIndex) so end halves near
        ; monitor edges are not attributed to the neighbor.
        if (WM_GetWindowRectHwnd(hwnd, &l, &t, &r, &b)) {
            wcx := (l + r) // 2
            wcy := (t + b) // 2
            wPoint := (wcy & 0xFFFFFFFF) << 32 | (wcx & 0xFFFFFFFF)
            hMon := DllCall("MonitorFromPoint", "int64", wPoint, "uint", 2, "ptr")
        } else
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
            rect := Buffer(16, 0)
            if !DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect)
                continue
            left := NumGet(rect, 0, "int")
            top := NumGet(rect, 4, "int")
            right := NumGet(rect, 8, "int")
            bottom := NumGet(rect, 12, "int")
            ; Center point — more reliable than MonitorFromWindow for end (right/bottom) halves
            ; that sit near a monitor edge and can be attributed to the neighbor.
            wcx := (left + right) // 2
            wcy := (top + bottom) // 2
            wPoint := (wcy & 0xFFFFFFFF) << 32 | (wcx & 0xFFFFFFFF)
            hMon := DllCall("MonitorFromPoint", "int64", wPoint, "uint", 2, "ptr")
            if (Integer(hMon) != Integer(hTarget))
                continue
            rows.Push({
                hwnd: hwnd,
                left: left,
                top: top,
                right: right,
                bottom: bottom
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
    ; Never rearrange because Clip Angel moved/maximized (see OnMoveSizeEnd).
    if (scheduleRearrange && !AutoSlot_IsClipAngelHwnd(hwnd)) {
        try AutoSlot_ScheduleRearrange(hwnd)
        catch {
        }
    }
    return true
}

; --- Foreground monitor swap (suite move-to-monitor) -------------------------

; Presentation pane for a window on monIdx ("start"|"end"|"").
; Prefer DWM/visible frame so end (right/bottom) halves classify like snap validation.
AutoSlot_GetHwndPaneOnMonitor(hwnd, monIdx) {
    if (!hwnd || monIdx < 1)
        return ""
    axis := WM_GetSnapSplitAxis(monIdx)
    WM_GetSnapInnerWorkArea(monIdx, &wl, &wt, &wr, &wb)
    l := 0, t := 0, r := 0, b := 0
    got := false
    try got := !!WM_GetSnapVisibleFrame(hwnd, monIdx, &l, &t, &r, &b)
    catch
        got := false
    if (!got && !WM_GetWindowRectHwnd(hwnd, &l, &t, &r, &b))
        return ""
    pane := ""
    paneSize := 0
    if (WM_ClassifySnapPane(axis, wl, wt, wr, wb, l, t, r, b, &pane, &paneSize) && pane != "")
        return pane
    ; Size/edge classify failed — still assign by center along the split axis.
    if (axis = "h") {
        center := wl + (wr - wl) // 2
        return ((l + r) // 2 < center) ? "start" : "end"
    }
    center := wt + (wb - wt) // 2
    return ((t + b) // 2 < center) ? "start" : "end"
}

; Among non-filled others on sourceMon, pick opposite-pane companion (or sole non-filled).
; If panes are ambiguous, pick the geometrically opposite window along the split axis.
AutoSlot_PickOppositeNonFilledCompanion(hwnd, sourceMon) {
    if (!hwnd || sourceMon < 1)
        return 0
    moverPane := AutoSlot_GetHwndPaneOnMonitor(hwnd, sourceMon)
    others := AutoSlot_OccupancyOnMonitor(sourceMon, hwnd)
    nonFilled := []
    for row in others {
        h := row.hwnd
        if (!h || AutoSlot_CompanionAlreadyFilled(h, sourceMon))
            continue
        nonFilled.Push(row)
    }
    if (nonFilled.Length = 0)
        return 0
    if (nonFilled.Length = 1)
        return nonFilled[1].hwnd
    if (moverPane != "") {
        wantPane := (moverPane = "start") ? "end" : "start"
        for row in nonFilled {
            if (AutoSlot_GetHwndPaneOnMonitor(row.hwnd, sourceMon) = wantPane)
                return row.hwnd
        }
    }
    ; Geometric fallback: farthest along split axis from mover center.
    axis := WM_GetSnapSplitAxis(sourceMon)
    if (!WM_GetWindowRectHwnd(hwnd, &ml, &mt, &mr, &mb))
        return nonFilled[1].hwnd
    mMid := (axis = "h") ? ((ml + mr) // 2) : ((mt + mb) // 2)
    bestH := 0
    bestDist := -1
    for row in nonFilled {
        oMid := (axis = "h") ? ((row.left + row.right) // 2) : ((row.top + row.bottom) // 2)
        dist := Abs(oMid - mMid)
        if (dist > bestDist) {
            bestDist := dist
            bestH := row.hwnd
        }
    }
    return bestH
}

; Dest/source FG layout: empty | full | single | pair | other.
; When 2+ occupants, pick two for a swappable pair (prefer opposite panes).
AutoSlot_PickPairFromOccupancyRows(rows) {
    if (!IsObject(rows) || rows.Length < 2)
        return 0
    monIdx := 0
    try monIdx := AutoSlot_GetHwndMonitorIndex(rows[1].hwnd)
    catch
        monIdx := 0
    bestA := 0
    bestB := 0
    bestAPane := ""
    bestBPane := ""
    if (monIdx >= 1) {
        found := false
        loop rows.Length - 1 {
            i := A_Index
            a := rows[i].hwnd
            aPane := AutoSlot_GetHwndPaneOnMonitor(a, monIdx)
            if (aPane = "")
                continue
            j := i + 1
            while (j <= rows.Length) {
                b := rows[j].hwnd
                bPane := AutoSlot_GetHwndPaneOnMonitor(b, monIdx)
                if (bPane != "" && bPane != aPane) {
                    bestA := a
                    bestB := b
                    bestAPane := aPane
                    bestBPane := bPane
                    found := true
                    break
                }
                j += 1
            }
            if (found)
                break
        }
    }
    if (!bestA) {
        bestA := rows[1].hwnd
        bestB := rows[2].hwnd
        if (monIdx >= 1) {
            bestAPane := AutoSlot_GetHwndPaneOnMonitor(bestA, monIdx)
            bestBPane := AutoSlot_GetHwndPaneOnMonitor(bestB, monIdx)
        }
        if (bestAPane = "" || bestBPane = "" || bestAPane = bestBPane) {
            ; Axis-aware: portrait (v) uses top; landscape (h) uses left.
            axis := (monIdx >= 1) ? WM_GetSnapSplitAxis(monIdx) : "h"
            aFirst := true
            if (axis = "v")
                aFirst := (rows[1].top <= rows[2].top)
            else
                aFirst := (rows[1].left <= rows[2].left)
            if (aFirst) {
                bestAPane := "start"
                bestBPane := "end"
            } else {
                bestAPane := "end"
                bestBPane := "start"
            }
        }
    }
    return { kind: "pair", a: bestA, b: bestB, aPane: bestAPane, bPane: bestBPane }
}

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
    ; 2+ visible candidates → swappable pair (stray 3rd must not kill swap).
    picked := AutoSlot_PickPairFromOccupancyRows(rows)
    if (IsObject(picked))
        return picked
    return { kind: "other" }
}

; Mover role on source: full | half | halfAlone | other (+ companion when half).
; Maximized / work-area-filled always counts as full even if covered extras share the monitor.
; Half detection ignores maximized-behind: strict snap, then opposite-pane / geometric non-filled.
AutoSlot_ClassifyMoverRole(hwnd, sourceMon) {
    if (!hwnd || sourceMon < 1)
        return { role: "other" }
    if (AutoSlot_CompanionAlreadyFilled(hwnd, sourceMon))
        return { role: "full" }
    companion := 0
    try companion := Integer(WM_FindStrictSnapCompanion(hwnd, sourceMon))
    catch
        companion := 0
    if (companion)
        return { role: "half", companion: companion }
    companion := AutoSlot_PickOppositeNonFilledCompanion(hwnd, sourceMon)
    if (companion)
        return { role: "half", companion: companion }
    others := AutoSlot_OccupancyOnMonitor(sourceMon, hwnd)
    nonFilled := []
    for row in others {
        h := row.hwnd
        if (!h || AutoSlot_CompanionAlreadyFilled(h, sourceMon))
            continue
        nonFilled.Push(h)
    }
    if (nonFilled.Length = 0)
        return { role: "halfAlone" }
    ; Any remaining non-filled other is a half companion (prefer first).
    return { role: "half", companion: nonFilled[1] }
}

AutoSlot_ApplyPairOnMonitor(a, b, aPane, bPane, monIdx, registerPair := true) {
    if (!a || !b || a = b || monIdx < 1)
        return false
    if (aPane = "" || bPane = "" || aPane = bPane) {
        aPane := "start"
        bPane := "end"
    }
    AutoSlot_UnregisterSnapPair(a)
    AutoSlot_UnregisterSnapPair(b)
    prepared := true
    if (!WM_PrepareHwndForTile(a) || !WM_PrepareHwndForTile(b))
        prepared := false
    axis := WM_GetSnapSplitAxis(monIdx)
    ok := false
    if (prepared)
        ok := !!WM_SnapPairGaplessRects(monIdx, axis, a, aPane, b, bPane)
    if (!ok && prepared && (aPane != "start" || bPane != "end")) {
        aPane := "start"
        bPane := "end"
        ok := !!WM_SnapPairGaplessRects(monIdx, axis, a, aPane, b, bPane)
    }
    ; Gapless can flake (DPI/chrome) — crude work-area halves still complete the swap.
    if (!ok)
        ok := AutoSlot_ApplyPairCrude(a, b, aPane, bPane, monIdx)
    if (!ok)
        return false
    ; If either half got OS-maximized during place, restore and re-gapless once.
    needRetry := false
    try needRetry := (WinGetMinMax("ahk_id " a) = 1) || (WinGetMinMax("ahk_id " b) = 1)
    catch
        needRetry := false
    if (needRetry) {
        try {
            if (WinGetMinMax("ahk_id " a) = 1) {
                WinRestore "ahk_id " a
                Sleep 60
            }
        } catch {
        }
        try {
            if (WinGetMinMax("ahk_id " b) = 1) {
                WinRestore "ahk_id " b
                Sleep 60
            }
        } catch {
        }
        if (WM_PrepareHwndForTile(a) && WM_PrepareHwndForTile(b)) {
            if (!WM_SnapPairGaplessRects(monIdx, axis, a, aPane, b, bPane))
                AutoSlot_ApplyPairCrude(a, b, aPane, bPane, monIdx)
        }
    }
    if (registerPair)
        AutoSlot_RegisterSnapPair(a, b)
    AutoSlot_RememberHwndMon(a)
    AutoSlot_RememberHwndMon(b)
    return true
}

; Best-effort 50/50 without gapless inset math (swap fallback).
AutoSlot_ApplyPairCrude(a, b, aPane, bPane, monIdx) {
    if (!a || !b || a = b || monIdx < 1)
        return false
    if (aPane = "" || bPane = "" || aPane = bPane) {
        aPane := "start"
        bPane := "end"
    }
    MonitorGetWorkArea monIdx, &wl, &wt, &wr, &wb
    axis := WM_GetSnapSplitAxis(monIdx)
    g := 4
    if (axis = "h") {
        half := (wr - wl - g) // 2
        startR := [wl, wt, wl + half, wb]
        endR := [wl + half + g, wt, wr, wb]
    } else {
        half := (wb - wt - g) // 2
        startR := [wl, wt, wr, wt + half]
        endR := [wl, wt + half + g, wr, wb]
    }
    aR := (aPane = "start") ? startR : endR
    bR := (bPane = "start") ? startR : endR
    try {
        if (WinGetMinMax("ahk_id " a) != 0) {
            WinRestore "ahk_id " a
            Sleep 40
        }
    } catch {
    }
    try {
        if (WinGetMinMax("ahk_id " b) != 0) {
            WinRestore "ahk_id " b
            Sleep 40
        }
    } catch {
    }
    okB := false
    okA := false
    try okB := !!WinMove(b, bR[1], bR[2], bR[3] - bR[1], bR[4] - bR[2])
    catch
        okB := false
    if (!okB)
        okB := !!DllCall("MoveWindow", "ptr", b, "int", bR[1], "int", bR[2], "int", bR[3] - bR[1], "int", bR[4] - bR[2],
            "int", true)
    try okA := !!WinMove(a, aR[1], aR[2], aR[3] - aR[1], aR[4] - aR[2])
    catch
        okA := false
    if (!okA)
        okA := !!DllCall("MoveWindow", "ptr", a, "int", aR[1], "int", aR[2], "int", aR[3] - aR[1], "int", aR[4] - aR[2],
            "int", true)
    return okA && okB
}

; After swap place: register pair and keep paired-max suppressed through the banner window.
AutoSlot_FinishSwappedPair(a, b) {
    if (!a || !b || a = b)
        return
    AutoSlot_RegisterSnapPair(a, b)
    AutoSlot_PairSuppressMark(a, AutoSlot_SWAP_PAIR_SUPPRESS_MS)
    AutoSlot_PairSuppressMark(b, AutoSlot_SWAP_PAIR_SUPPRESS_MS)
    AutoSlot_ClearPairMaxPending(a)
    AutoSlot_ClearPairMaxPending(b)
}

AutoSlot_ForgetHwndMon(hwnd) {
    global g_AutoSlotHwndMon
    if (hwnd && g_AutoSlotHwndMon.Has(hwnd))
        g_AutoSlotHwndMon.Delete(hwnd)
}

; True while [F]-replaced hwnd must not be fill/rearrange-promoted (clears if restored).
AutoSlot_ReplaceSkipActive(hwnd) {
    global g_AutoSlotReplaceSkip
    if (!hwnd || !g_AutoSlotReplaceSkip.Has(hwnd))
        return false
    if (!DllCall("IsWindow", "ptr", hwnd)) {
        g_AutoSlotReplaceSkip.Delete(hwnd)
        return false
    }
    if (A_TickCount - g_AutoSlotReplaceSkip[hwnd] > AutoSlot_REPLACE_SKIP_TTL_MS) {
        g_AutoSlotReplaceSkip.Delete(hwnd)
        return false
    }
    minMax := 0
    try minMax := WinGetMinMax("ahk_id " hwnd)
    catch
        minMax := 0
    ; User restored the window — allow promote again.
    if (minMax != -1) {
        try {
            if (!WM_WindowIsTaskbarMinimized(hwnd)) {
                g_AutoSlotReplaceSkip.Delete(hwnd)
                return false
            }
        } catch {
            g_AutoSlotReplaceSkip.Delete(hwnd)
            return false
        }
    }
    return true
}

AutoSlot_MarkReplaceSkip(hwnd) {
    global g_AutoSlotReplaceSkip
    if (hwnd)
        g_AutoSlotReplaceSkip[hwnd] := A_TickCount
}

AutoSlot_ClearSwapDisplaced() {
    global g_AutoSlotSwapDisplaced, g_AutoSlotSwapMoverHwnd
    g_AutoSlotSwapDisplaced := []
    g_AutoSlotSwapMoverHwnd := 0
}

AutoSlot_OnSwapF(*) {
    global g_AutoSlotSwapDisplaced, g_AutoSlotSwapMoverHwnd, g_AutoSlotHwndMon
    displaced := g_AutoSlotSwapDisplaced
    mover := g_AutoSlotSwapMoverHwnd
    AutoSlot_ClearSwapDisplaced()
    try StandardLoadingBar_CloseKeysOverlay()
    catch {
    }
    try StandardLoadingBar_Hide(0)
    catch {
    }
    claimed := Map()
    for h in displaced {
        if (!h || !DllCall("IsWindow", "ptr", h))
            continue
        monIdx := 0
        if (g_AutoSlotHwndMon.Has(h))
            monIdx := g_AutoSlotHwndMon[h]
        if (monIdx < 1)
            monIdx := AutoSlot_GetHwndMonitorIndex(h)
        try WinMinimize("ahk_id " h)
        catch {
        }
        AutoSlot_UnregisterSnapPair(h)
        AutoSlot_MarkReplaceSkip(h)
        AutoSlot_ForgetHwndMon(h)
        if (monIdx >= 1)
            claimed[monIdx] := true
    }
    for monIdx, _ in claimed
        AutoSlot_ClaimMonitor(monIdx)
    AutoSlot_ActivateHwnd(mover)
    ; [F] ends the swap: clear quiet + claim cooldown so vacated monitors fill now.
    ; (ScheduleRearrange no-ops while quiet; ClaimMonitor would also block import 1.5s.)
    global g_AutoSlotSwapQuietUntil, g_AutoSlotFillCooldown
    g_AutoSlotSwapQuietUntil := 0
    for monIdx, _ in claimed {
        if (g_AutoSlotFillCooldown.Has(monIdx))
            g_AutoSlotFillCooldown.Delete(monIdx)
    }
    ; Heal/fill what is now foreground on vacated monitors (ReplaceSkip blocks re-promote).
    try AutoSlot_ScheduleRearrange(mover)
    catch {
    }
}

AutoSlot_OnSwapTimeout(*) {
    global g_AutoSlotSwapMoverHwnd
    mover := g_AutoSlotSwapMoverHwnd
    AutoSlot_ClearSwapDisplaced()
    try StandardLoadingBar_CloseKeysOverlay()
    catch {
    }
    try StandardLoadingBar_Hide(0)
    catch {
    }
    AutoSlot_ActivateHwnd(mover)
}

AutoSlot_ShowSwapModal(srcLabel, dstLabel, moverHwnd, displacedHwnds) {
    global g_AutoSlotSwapDisplaced, g_AutoSlotSwapMoverHwnd
    g_AutoSlotSwapDisplaced := displacedHwnds
    g_AutoSlotSwapMoverHwnd := moverHwnd
    ; Keep heal/pair-max/rearrange gated for the whole Interactive Input window.
    AutoSlot_BeginSwapQuiet(AutoSlot_SWAP_MODAL_MS + 500)
    for h in displacedHwnds {
        if (h) {
            AutoSlot_PairSuppressMark(h, AutoSlot_SWAP_PAIR_SUPPRESS_MS)
            AutoSlot_ClearPairMaxPending(h)
        }
    }
    if (moverHwnd) {
        AutoSlot_PairSuppressMark(moverHwnd, AutoSlot_SWAP_PAIR_SUPPRESS_MS)
        AutoSlot_ClearPairMaxPending(moverHwnd)
    }
    try StandardLoadingBar_CloseKeysOverlay()
    catch {
    }
    try StandardLoadingBar_Hide(0)
    catch {
    }
    AutoSlot_ActivateHwnd(moverHwnd)
    keyCallbacks := Map("F", AutoSlot_OnSwapF)
    try {
        ; Interactive Input: ❓ + intermediate accent; progress countdown; overlay owns focus
        ; after MEH chord (preserveUserFocus false). Key F = replace (minimize displaced).
        StandardLoadingBar_ShowWithKeys(
            "❓ Swapped M" srcLabel " ↔ M" dstLabel " (2s)",
            keyCallbacks,
            AutoSlot_SWAP_MODAL_MS,
            0,
            AutoSlot_OnSwapTimeout,
            BANNER_ACCENT_INTERMEDIATE,
            480,
            17,
            "",
            true,
            "[F] Replace (minimize displaced)",
            true,
            true,
            false
        )
    } catch {
        AutoSlot_ClearSwapDisplaced()
        AutoSlot_Toast("ℹ️ Swapped M" srcLabel " ↔ M" dstLabel)
    }
    AutoSlot_ActivateHwnd(moverHwnd)
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
    ; Lone dest occupant (single) exchanges like a full window.
    if (dest.kind = "single")
        dest := { kind: "full", hwnd: dest.hwnd }
    if (role = "other" || dest.kind = "empty" || dest.kind = "other")
        return false

    ; half→pair allowed: incoming pair covers source; leftover companion stays behind.
    wantSwap := false
    if ((role = "full" || role = "halfAlone" || role = "half") && dest.kind = "pair")
        wantSwap := true
    else if ((role = "full" || role = "halfAlone" || role = "half") && dest.kind = "full")
        wantSwap := true
    if (!wantSwap)
        return false

    displaced := []
    if (dest.kind = "pair") {
        displaced.Push(dest.a)
        displaced.Push(dest.b)
    } else if (dest.kind = "full")
        displaced.Push(dest.hwnd)

    companion := 0
    if (role = "half") {
        try companion := Integer(mover.companion)
        catch
            companion := 0
    }

    ; Quiet heal/pair-max/rearrange; cancel queued paired-maximize; break snap links.
    AutoSlot_BeginSwapQuiet()
    parties := [hwnd]
    if (companion)
        parties.Push(companion)
    for h in displaced
        parties.Push(h)
    for h in parties {
        AutoSlot_PairSuppressMark(h, AutoSlot_SWAP_PAIR_SUPPRESS_MS)
        AutoSlot_ClearPairMaxPending(h)
        AutoSlot_UnregisterSnapPair(h)
    }

    if (dest.kind = "pair") {
        ; Displaced pair onto source FIRST; do not register snap pair until layout is stable.
        placedPair := AutoSlot_ApplyPairOnMonitor(dest.a, dest.b, dest.aPane, dest.bPane, sourceMon, false)
        if (!placedPair) {
            ; Last resort: stack both maximized on source so exchange still happens.
            AutoSlot_MaximizeOnMonitor(dest.a, sourceMon, false)
            AutoSlot_MaximizeOnMonitor(dest.b, sourceMon, false)
        }
        if (!AutoSlot_MaximizeOnMonitor(hwnd, destMon, false)) {
            global g_AutoSlotSwapQuietUntil
            g_AutoSlotSwapQuietUntil := 0
            return false
        }
        AutoSlot_RememberHwndMon(hwnd)
        if (placedPair)
            AutoSlot_FinishSwappedPair(dest.a, dest.b)
    } else if (dest.kind = "full") {
        other := dest.hwnd
        if (other = hwnd) {
            global g_AutoSlotSwapQuietUntil
            g_AutoSlotSwapQuietUntil := 0
            return false
        }
        ; Former dest window onto source first, then mover takes dest.
        ; Place other into the vacated mover pane; keep companion on its pane.
        placedOther := false
        if (companion && companion != other) {
            moverPane := AutoSlot_GetHwndPaneOnMonitor(hwnd, sourceMon)
            companionPane := AutoSlot_GetHwndPaneOnMonitor(companion, sourceMon)
            if (moverPane = "" && companionPane != "")
                moverPane := (companionPane = "start") ? "end" : "start"
            if (companionPane = "" && moverPane != "")
                companionPane := (moverPane = "start") ? "end" : "start"
            if (moverPane = "" || companionPane = "" || moverPane = companionPane) {
                moverPane := "start"
                companionPane := "end"
            }
            placedOther := AutoSlot_ApplyPairOnMonitor(other, companion, moverPane, companionPane, sourceMon, false)
            if (placedOther)
                AutoSlot_FinishSwappedPair(other, companion)
        }
        if (!placedOther)
            placedOther := AutoSlot_MaximizeOnMonitor(other, sourceMon, false)
        ; If other resisted, still continue — mover takes dest (partial swap better than none).
        if (placedOther)
            AutoSlot_RememberHwndMon(other)
        if (!AutoSlot_MaximizeOnMonitor(hwnd, destMon, false)) {
            global g_AutoSlotSwapQuietUntil
            g_AutoSlotSwapQuietUntil := 0
            return false
        }
        AutoSlot_RememberHwndMon(hwnd)
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
    AutoSlot_ShowSwapModal(srcLabel, dstLabel, hwnd, displaced)
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
    ; Clip Angel is a quick-use overlay — never 50/50 snap with it.
    if (AutoSlot_IsClipAngelHwnd(newHwnd) || AutoSlot_IsClipAngelHwnd(partnerHwnd))
        return ""

    partnerState := AutoSlot_CapturePartnerState(partnerHwnd)
    partnerWasMax := IsObject(partnerState) ? partnerState.wasMaximized : false
    if (!partnerWasMax) {
        try partnerWasMax := (WinGetMinMax("ahk_id " partnerHwnd) = 1)
        catch
            partnerWasMax := false
    }
    ; Work-area fill counts as "max" for snap (mm often 0 on Chrome after heal).
    if (!partnerWasMax && AutoSlot_CompanionAlreadyFilled(partnerHwnd, monIdx))
        partnerWasMax := true

    ; Single prepare (parity with ^!#x) — avoid double Sleep from AutoSlot_Prepare + WM_Prepare.
    if (!WM_PrepareHwndForTile(newHwnd) || !WM_PrepareHwndForTile(partnerHwnd))
        return ""
    ; Restore-from-minimized can land OS-maximized; work-area fill needs an explicit demax nudge.
    AutoSlot_EnsureRestoredForSnap(newHwnd, monIdx)
    AutoSlot_EnsureRestoredForSnap(partnerHwnd, monIdx)

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
    ; Gapless can report success while a work-area residual never shrunk (Chrome).
    if (AutoSlot_CompanionAlreadyFilled(partnerHwnd, monIdx)) {
        AutoSlot_EnsureRestoredForSnap(partnerHwnd, monIdx)
        ok := WM_SnapPairGaplessRects(monIdx, axis, newHwnd, newPane, partnerHwnd, partnerPane)
        if (!ok)
            return ""
    }
    ; Last resort: crude outer-rect place when gapless still leaves partner full.
    if (AutoSlot_CompanionAlreadyFilled(partnerHwnd, monIdx)) {
        rects := WM_ComputeSnapPairPaneRects(monIdx, axis)
        if (rects.Count > 0) {
            AutoSlot_EnsureRestoredForSnap(partnerHwnd, monIdx)
            AutoSlot_EnsureRestoredForSnap(newHwnd, monIdx)
            pRect := rects[partnerPane]
            tRect := rects[newPane]
            AutoSlot_MoveHwndToRect(partnerHwnd, pRect[1], pRect[2], pRect[3], pRect[4])
            AutoSlot_MoveHwndToRect(newHwnd, tRect[1], tRect[2], tRect[3], tRect[4])
            Sleep 80
        }
        if (AutoSlot_CompanionAlreadyFilled(partnerHwnd, monIdx))
            return ""
    }
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

; Demax OS-maximized and work-area-filled windows so 50/50 MoveWindow sticks.
AutoSlot_EnsureRestoredForSnap(hwnd, monIdx := 0) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    changed := false
    try {
        mm := WinGetMinMax("ahk_id " hwnd)
        if (mm = 1 || mm = -1) {
            WinRestore "ahk_id " hwnd
            Sleep 80
            changed := true
        }
        ; Restore-from-minimized often lands maximized — second pass.
        if (WinGetMinMax("ahk_id " hwnd) = 1) {
            WinRestore "ahk_id " hwnd
            Sleep 80
            changed := true
        }
    } catch {
    }
    if (monIdx >= 1 && AutoSlot_CompanionAlreadyFilled(hwnd, monIdx)) {
        try {
            MonitorGetWorkArea monIdx, &wl, &wt, &wr, &wb
            axis := WM_GetSnapSplitAxis(monIdx)
            ; Nudge off full work-area along the snap axis so gapless sizing sticks.
            if (axis = "h") {
                half := wl + Max(200, (wr - wl) // 2)
                okMove := AutoSlot_MoveHwndToRect(hwnd, wl, wt, half, wb)
            } else {
                half := wt + Max(200, (wb - wt) // 2)
                okMove := AutoSlot_MoveHwndToRect(hwnd, wl, wt, wr, half)
            }
            if (okMove) {
                Sleep 50
                changed := true
            }
        } catch {
        }
    }
    return changed
}

; --- Toast / undo modal ------------------------------------------------------

AutoSlot_Toast(msg) {
    global g_AutoSlotLastToastMsg, g_AutoSlotLastToastTick
    if (msg = "")
        return
    ; Re-showing the same toast resets hide timers and looks "stuck" for minutes.
    if (msg = g_AutoSlotLastToastMsg && A_TickCount - g_AutoSlotLastToastTick < AutoSlot_TOAST_DEBOUNCE_MS)
        return
    g_AutoSlotLastToastMsg := msg
    g_AutoSlotLastToastTick := A_TickCount
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

; True when hwnd still belongs to a living *visible* 50/50 pair (must not be relocated by fill).
; Minimized members are no longer active pair participants — stale registry must not block Y/fill.
AutoSlot_BackgroundCandHasLivingSnapPartner(hwnd) {
    global g_AutoSlotSnapPairs
    if (!hwnd || !g_AutoSlotSnapPairs.Has(hwnd))
        return false
    try {
        if (WinGetMinMax("ahk_id " hwnd) = -1) {
            AutoSlot_UnregisterSnapPair(hwnd)
            return false
        }
    } catch {
    }
    partner := g_AutoSlotSnapPairs[hwnd]
    if (!partner || !DllCall("IsWindow", "ptr", partner)) {
        AutoSlot_UnregisterSnapPair(hwnd)
        return false
    }
    try {
        if (WinGetMinMax("ahk_id " partner) = -1) {
            AutoSlot_UnregisterSnapPair(hwnd)
            return false
        }
    } catch {
    }
    ; Partner maximized / work-area filled (e.g. after F11 exit) is no longer a live
    ; 50/50 member — break the stale pair so this companion can fill the free half-slot.
    try {
        if (WinGetMinMax("ahk_id " partner) = 1) {
            AutoSlot_UnregisterSnapPair(hwnd)
            return false
        }
    } catch {
    }
    partnerMon := AutoSlot_GetHwndMonitorIndex(partner)
    if (partnerMon >= 1 && AutoSlot_CompanionAlreadyFilled(partner, partnerMon)) {
        AutoSlot_UnregisterSnapPair(hwnd)
        return false
    }
    return true
}

; True when a *visible* hwnd's home monitor has another window in F11 fullscreen
; (hwnd is covered there). Minimized candidates are never "covered" — they live on
; the taskbar. Normal AutoSlot maximized / work-area windows are NOT F11 (Chrome
; work-area fill was a false positive that blocked every covered background).
AutoSlot_BackgroundCandCoveredByF11(hwnd) {
    global g_AutoSlotHwndMon
    if (!hwnd)
        return false
    try {
        if (WinGetMinMax("ahk_id " hwnd) = -1)
            return false
    } catch {
    }
    try {
        if (!DllCall("IsWindowVisible", "ptr", hwnd))
            return false
    } catch {
    }
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
            ; Slotted max / work-area fill is capacity, not F11 cover.
            if (AutoSlot_CompanionAlreadyFilled(other, monIdx))
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
; forceImport: explicit user fill (Y) — bypass the [F] replace-skip like the Ctrl+Alt+Win+6 path.
AutoSlot_PickBackgroundCandidate(monIdx, occupancyRows, excludeExtra := 0, forceImport := false) {
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
        ; [F] replace: do not re-promote minimized displaced windows into freed slots.
        ; Explicit user fill (Y) overrides this, matching the Ctrl+Alt+Win+6 place path.
        if (!forceImport && AutoSlot_ReplaceSkipActive(hwnd))
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

; Split occupancy into non-filled (half-slot) vs filled (maximized / work-area).
AutoSlot_PartitionOccupancy(monIdx, excludeHwnd := 0) {
    nonFilled := []
    filled := []
    for row in AutoSlot_OccupancyOnMonitor(monIdx, excludeHwnd) {
        h := row.hwnd
        if (!h)
            continue
        if (AutoSlot_CompanionAlreadyFilled(h, monIdx))
            filled.Push(row)
        else
            nonFilled.Push(row)
    }
    return { nonFilled: nonFilled, filled: filled }
}

; Maximize the sole non-filled window on monIdx (ignores maximized-behind).
AutoSlot_HealLoneCompanion(monIdx) {
    if (monIdx < 1 || monIdx > MonitorGetCount())
        return false
    part := AutoSlot_PartitionOccupancy(monIdx)
    if (part.nonFilled.Length != 1)
        return false
    companion := part.nonFilled[1].hwnd
    if (AutoSlot_IsClipAngelHwnd(companion))
        return false
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
        ; Suppress MoveSizeEnd/paired-max rearrange so heal does not re-toast in a loop.
        AutoSlot_PairSuppressMark(companion, AutoSlot_RECENT_MS)
        AutoSlot_ClaimMonitor(monIdx)
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
; Policy uses non-filled count so maximized-behind does not block heal/fill.
; forceImport: true for explicit Ctrl+Alt+Win+Y (ignore place freeze / fill cooldown).
AutoSlot_FillMonitorFromBackground(monIdx, forceImport := false) {
    global g_AutoSlotRecent, g_AutoSlotUndo
    if (monIdx < 1 || monIdx > MonitorGetCount() || MonitorGetCount() <= 1)
        return "noop"
    part := AutoSlot_PartitionOccupancy(monIdx)
    if (part.nonFilled.Length >= 2)
        return "stale"

    order := AutoSlot_OrderForMonitorIndex(monIdx)
    label := order > 0 ? order : monIdx

    ; Place freeze / claim cooldown: companion heal only (no background import),
    ; unless the user explicitly requested fill (Y).
    blockImport := !forceImport && (AutoSlot_PlaceFreezeActive() || AutoSlot_FillCooldownActive(monIdx))

    if (part.nonFilled.Length = 1) {
        residual := part.nonFilled[1].hwnd
        if (AutoSlot_IsClipAngelHwnd(residual))
            return "noop"
        if (AutoSlot_CompanionAlreadyFilled(residual, monIdx))
            return "ok"
        ; Half + maximized/work-area on same monitor: 50/50 them (do not heal the half
        ; into a second full window — that was skipping Y after a partial snap).
        if (part.filled.Length >= 1 && !blockImport) {
            filledHwnd := part.filled[1].hwnd
            if (filledHwnd && filledHwnd != residual && !AutoSlot_IsClipAngelHwnd(filledHwnd)) {
                g_AutoSlotRecent[residual] := A_TickCount
                AutoSlot_PruneRecent()
                pane := AutoSlot_SnapPair(residual, filledHwnd, monIdx, true)
                g_AutoSlotUndo := 0
                if (pane != "") {
                    AutoSlot_RememberHwndMon(residual)
                    AutoSlot_RememberHwndMon(filledHwnd)
                    ; Suppress the snap's own move/location events + claim cooldown so this
                    ; fill does not immediately re-trigger itself (stuck "Slot filled" loop).
                    AutoSlot_PairSuppressMark(residual, AutoSlot_RECENT_MS)
                    AutoSlot_PairSuppressMark(filledHwnd, AutoSlot_RECENT_MS)
                    AutoSlot_ClaimMonitor(monIdx)
                    AutoSlot_Toast("ℹ️ Slot filled → M" label " (50/50)")
                    return "ok"
                }
            }
        }
        ; Closing one half of a pair: maximize the leftover to both slots first.
        ; Only import a background companion if heal fails (do not keep residual as a half).
        if (AutoSlot_HealLoneCompanion(monIdx))
            return "ok"
        if (blockImport)
            return "noop"
        residualRows := [{ hwnd: residual }]
        cand := AutoSlot_PickBackgroundCandidate(monIdx, residualRows, 0, forceImport)
        if (!cand)
            return "noop"
        g_AutoSlotRecent[cand] := A_TickCount
        AutoSlot_PruneRecent()
        pane := AutoSlot_SnapPair(cand, residual, monIdx, true)
        g_AutoSlotUndo := 0
        if (pane = "")
            return "noop"
        AutoSlot_RememberHwndMon(cand)
        AutoSlot_RememberHwndMon(residual)
        ; Suppress the snap's own move/location events + claim cooldown so this fill
        ; does not immediately re-trigger itself (stuck "Slot filled" loop).
        AutoSlot_PairSuppressMark(cand, AutoSlot_RECENT_MS)
        AutoSlot_PairSuppressMark(residual, AutoSlot_RECENT_MS)
        AutoSlot_ClaimMonitor(monIdx)
        AutoSlot_Toast("ℹ️ Slot filled → M" label " (50/50)")
        return "ok"
    }

    ; Lone maximized / work-area fill = half-full (1 of 2 slots) — SnapPair background in place.
    if (part.filled.Length = 1) {
        residual := part.filled[1].hwnd
        if (AutoSlot_IsClipAngelHwnd(residual))
            return "noop"
        if (blockImport)
            return "noop"
        residualRows := [{ hwnd: residual }]
        cand := AutoSlot_PickBackgroundCandidate(monIdx, residualRows, 0, forceImport)
        if (!cand)
            return "noop"
        g_AutoSlotRecent[cand] := A_TickCount
        AutoSlot_PruneRecent()
        pane := AutoSlot_SnapPair(cand, residual, monIdx, true)
        g_AutoSlotUndo := 0
        if (pane = "")
            return "noop"
        AutoSlot_RememberHwndMon(cand)
        AutoSlot_RememberHwndMon(residual)
        ; Suppress the snap's own move/location events + claim cooldown so this fill
        ; does not immediately re-trigger itself (stuck "Slot filled" loop).
        AutoSlot_PairSuppressMark(cand, AutoSlot_RECENT_MS)
        AutoSlot_PairSuppressMark(residual, AutoSlot_RECENT_MS)
        AutoSlot_ClaimMonitor(monIdx)
        AutoSlot_Toast("ℹ️ Slot filled → M" label " (50/50)")
        return "ok"
    }

    ; Multiple filled, no free half — nothing to import.
    if (part.filled.Length >= 2)
        return "ok"

    ; Empty monitor — fill both slots when two candidates exist.
    if (blockImport)
        return "noop"
    others := []
    cand1 := AutoSlot_PickBackgroundCandidate(monIdx, others, 0, forceImport)
    if (!cand1)
        return "noop"
    cand2 := AutoSlot_PickBackgroundCandidate(monIdx, others, cand1, forceImport)
    if (cand2) {
        g_AutoSlotRecent[cand1] := A_TickCount
        g_AutoSlotRecent[cand2] := A_TickCount
        AutoSlot_PruneRecent()
        pane := AutoSlot_SnapPair(cand1, cand2, monIdx, true)
        g_AutoSlotUndo := 0
        if (pane != "") {
            AutoSlot_RememberHwndMon(cand1)
            AutoSlot_RememberHwndMon(cand2)
            AutoSlot_PairSuppressMark(cand1, AutoSlot_RECENT_MS)
            AutoSlot_PairSuppressMark(cand2, AutoSlot_RECENT_MS)
            AutoSlot_ClaimMonitor(monIdx)
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
        cand2 := AutoSlot_PickBackgroundCandidate(monIdx, [{ hwnd: cand1 }], cand1, forceImport)
    if (cand2) {
        g_AutoSlotRecent[cand2] := A_TickCount
        AutoSlot_PruneRecent()
        pane := AutoSlot_SnapPair(cand2, cand1, monIdx, true)
        g_AutoSlotUndo := 0
        if (pane != "") {
            AutoSlot_RememberHwndMon(cand2)
            AutoSlot_PairSuppressMark(cand1, AutoSlot_RECENT_MS)
            AutoSlot_PairSuppressMark(cand2, AutoSlot_RECENT_MS)
            AutoSlot_ClaimMonitor(monIdx)
            AutoSlot_Toast("ℹ️ Slot filled → M" label " (50/50)")
            return "ok"
        }
    }
    AutoSlot_PairSuppressMark(cand1, AutoSlot_RECENT_MS)
    AutoSlot_ClaimMonitor(monIdx)
    AutoSlot_Toast("ℹ️ Slot filled → M" label " (maximized)")
    return "ok"
}

; --- Placement ---------------------------------------------------------------

; Place a user-picked background hwnd into free AutoSlot capacity (empty → max,
; one free half → 50/50 with residual). Returns true if placed into a slot.
; Used by Ctrl+Alt+Win+6 list open when AutoSlot is ON.
AutoSlot_TryPlaceBackgroundHwnd(hwnd) {
    global g_AutoSlotRecent, g_AutoSlotUndo
    if (!AutoSlot_IsEnabled() || !hwnd || !WinExist("ahk_id " hwnd))
        return false
    if (MonitorGetCount() <= 1)
        return false
    if (AutoSlot_IsClipAngelHwnd(hwnd) || AutoSlot_IsExcludedExeOrTitle(hwnd))
        return false
    ; Stale pair from a previous 50/50 must not block this explicit pick.
    AutoSlot_UnregisterSnapPair(hwnd)
    AutoSlot_ForgetHwndMon(hwnd)
    g_AutoSlotRecent[hwnd] := A_TickCount
    AutoSlot_PruneRecent()

    ordinalCount := Min(MonitorGetCount(), AutoSlot_MAX_ORDINAL)
    emptyMon := 0
    emptyOrder := 0
    halfMon := 0
    halfOrder := 0
    halfPartner := 0
    loop ordinalCount {
        order := A_Index
        monIdx := AutoSlot_GetMonitorIndexByOrder(order)
        if (!monIdx)
            continue
        part := AutoSlot_PartitionOccupancy(monIdx, hwnd)
        used := part.filled.Length + part.nonFilled.Length
        if (used = 0) {
            if (!emptyMon) {
                emptyMon := monIdx
                emptyOrder := order
            }
            continue
        }
        if (used = 1 && !halfMon) {
            halfMon := monIdx
            halfOrder := order
            halfPartner := part.filled.Length = 1 ? part.filled[1].hwnd : part.nonFilled[1].hwnd
        }
    }
    if (emptyMon) {
        if (!AutoSlot_MaximizeOnMonitor(hwnd, emptyMon, false))
            return false
        AutoSlot_ClaimMonitor(emptyMon)
        AutoSlot_RememberHwndMon(hwnd)
        AutoSlot_Toast("ℹ️ Opened → M" emptyOrder " (maximized)")
        return true
    }
    if (halfMon && halfPartner) {
        pane := AutoSlot_SnapPair(hwnd, halfPartner, halfMon, true)
        if (pane = "")
            return false
        g_AutoSlotUndo := 0
        AutoSlot_ClaimMonitor(halfMon)
        AutoSlot_RememberHwndMon(hwnd)
        AutoSlot_RememberHwndMon(halfPartner)
        AutoSlot_Toast("ℹ️ Opened → M" halfOrder " (50/50)")
        return true
    }
    return false
}

; Partner to 50/50 a new window with on monIdx, ignoring windows hidden behind a
; maximized one: one visible maximized/work-area window = a free half; else one lone half.
; Two filled (2 maximized) or two half-panes = genuinely full → 0.
AutoSlot_MonitorFreeHalfPartner(monIdx, excludeHwnd := 0) {
    part := AutoSlot_PartitionOccupancy(monIdx, excludeHwnd)
    if (part.filled.Length = 1)
        return part.filled[1].hwnd
    if (part.filled.Length = 0 && part.nonFilled.Length = 1)
        return part.nonFilled[1].hwnd
    return 0
}

; First ordinal monitor with a free half-slot (lone maximized or single half-pane).
AutoSlot_FindHalfFullMonitor(excludeHwnd := 0) {
    ordinalCount := Min(MonitorGetCount(), AutoSlot_MAX_ORDINAL)
    loop ordinalCount {
        order := A_Index
        monIdx := AutoSlot_GetMonitorIndexByOrder(order)
        if (!monIdx)
            continue
        partner := AutoSlot_MonitorFreeHalfPartner(monIdx, excludeHwnd)
        if (partner)
            return { order: order, monIdx: monIdx, partner: partner }
    }
    return 0
}

; Snap a new window with an explicit partner (partner demaximizes to a half via SnapPair).
AutoSlot_TrySnapNewWithPartner(hwnd, monIdx, partner, orderLabel := 0) {
    if (!partner || partner = hwnd)
        return false
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
    ; Empty = no filled and no half occupant (windows hidden behind a maximized one do not count).
    emptyOrder := 0
    emptyMon := 0

    loop ordinalCount {
        order := A_Index
        monIdx := AutoSlot_GetMonitorIndexByOrder(order)
        if (!monIdx)
            continue
        part := AutoSlot_PartitionOccupancy(monIdx, hwnd)
        if (part.filled.Length = 0 && part.nonFilled.Length = 0) {
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

    ; Prefer partitioning the origin slot when it has a free half (incl. lone maximized).
    originMon := AutoSlot_GetHwndMonitorIndex(hwnd)
    if (originMon >= 1) {
        partner := AutoSlot_MonitorFreeHalfPartner(originMon, hwnd)
        if (partner && AutoSlot_TrySnapNewWithPartner(hwnd, originMon, partner))
            return
    }

    ; Half-slots remain open on any monitor with a single visible window (maximized = 1 of 2 slots).
    half := AutoSlot_FindHalfFullMonitor(hwnd)
    if (IsObject(half) && AutoSlot_TrySnapNewWithPartner(hwnd, half.monIdx, half.partner, half.order))
        return

    if (AutoSlot_MaximizeInPlace(hwnd))
        msg := "ℹ️ Grid full — maximized"
    AutoSlot_Toast(msg)
}

; Auto-execute when #included.
AutoSlot_Init()