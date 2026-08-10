; =============================================================================
; AutoSlot — optional auto-position for newly opened windows (multi-monitor).
;
; Detection/placement policy is self-contained. 50/50 placement reuses the
; proven gapless snap APIs from WindowManagement\tile_snap.ahk (same as ^!#x).
; After a successful snap, a 2s ShowWithKeys modal offers [M] undo.
; Delete: remove the #include in WindowManagement.ahk + this folder (see README).
;
; Placement (MonitorGetCount() > 1 only); 2 slots per ordinal monitor (max 8):
;   1) First empty ordinal → maximize onto it
;   2) Else first free half (lone max / lone half-pane) → 50/50 SnapPair
;   3) Else leave as-is ("grid full" — do not maximize over existing windows)
;
; Close / minimize: heal leftover snap companion only (silent). No automatic
; background import — user fills free slots with Ctrl+Alt+Win+6 (forceImport).
;
; Maximizing one half of a 50/50 pair does NOT maximize the companion (pair breaks).
;
; Foreground monitor swap (suite MoveWinToMonitor when AutoSlot ON):
;   Full/halfAlone/half↔pair; full/halfAlone/half↔full|single. Toast + quiet only.
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
global g_AutoSlotHealPending := Map()
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
global g_AutoSlotSwapQuietUntil := 0
global g_AutoSlotJustRestored := Map()
global g_AutoSlotLastToastMsg := ""
global g_AutoSlotLastToastTick := 0
global g_AutoSlotWasF11 := Map()          ; snap-pair hwnd → true while (or after) F11 fullscreen
global g_AutoSlotF11RestorePending := Map()
; Shared background rows for one Ctrl+Alt+Win+6 pass (avoid re-enumerating per monitor).
global g_AutoSlotYBgRows := 0
global g_AutoSlotYBgActive := false
; Perf log: hwnd → origin tick for delta ms (Schedule / Place phases). Off unless diagnosing.
global g_AutoSlotPerfOrigin := Map()
; Eligibility settle: hwnd → first-miss tick (dense ~100 ms polls until budget).
global g_AutoSlotEligRetry := Map()
; Place critical section depth (WinEvent SHOW must not re-enter during Place).
global g_AutoSlotPlaceDepth := 0
; Cached BuildOccupancyByMonitor snapshot for SHOW occupied checks.
global g_AutoSlotOccSnap := 0
global g_AutoSlotOccSnapTick := 0
; SHOW coalescing: WinEvent must not run ScheduleFromShow synchronously (blocks debounce timers).
global g_AutoSlotShowPending := Map()
global g_AutoSlotShowTimerArmed := false

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
AutoSlot_SWAP_QUIET_MS := 2500
AutoSlot_SWAP_MODAL_MS := 2000
AutoSlot_SWAP_PAIR_SUPPRESS_MS := 3500
AutoSlot_RESTORE_GUARD_MS := 4000   ; MINIMIZEEND → suppress SHOW-triggered auto-place for this long
AutoSlot_TOAST_DEBOUNCE_MS := 4000  ; identical toast text — avoid hide-timer reset spam
; File IPC fallback for cross-process place (PostMessage preferred); slow poll only.
AutoSlot_PLACE_REQUEST_POLL_MS := 1000
AutoSlot_MAX_ORDINAL := 4
; Must match Utils\autoslot_place_ipc.ahk
AUTOSLOT_PLACE_MSG_NAME := "EDU_AutoSlot_PlaceHwnd"
AutoSlot_STICKY_PLACE_MS := 400
AutoSlot_STICKY_PLACE_MAX := 12
; Perf log (off by default): set env AUTOSLOT_PERF_LOG=1 → .cursor\autoslot_perf.log
AutoSlot_PERF_LOG := Trim(EnvGet("AUTOSLOT_PERF_LOG")) = "1"
; Eligibility settle after IsEligibleNewWindow miss: poll every POLL_MS until BUDGET_MS
; from first miss (empty title / HWND not ready). Do not one-shot abandon; do not use
; sparse 300/800/1500 gaps (those reintroduced multi-second Place lag).
AutoSlot_ELIG_RETRY_POLL_MS := 100
AutoSlot_ELIG_RETRY_BUDGET_MS := 2000
AutoSlot_OCC_CACHE_MS := 150
AutoSlot_SHOW_DEFER_MS := 50
AutoSlot_SHOW_BATCH_MAX := 12
global g_AutoSlotPlaceRequestFile := A_ScriptDir "\.cursor\autoslot_place_request"
global g_AutoSlotPlaceMsg := 0
global g_AutoSlotStickyHwnd := 0
global g_AutoSlotStickyAttempts := 0
AutoSlot_HSHELL_WINDOWCREATED := 1
AutoSlot_HSHELL_WINDOWDESTROYED := 2
AutoSlot_UNDO_MODAL_MS := 2000

; --- Perf log (optional; AutoSlot_PERF_LOG) ----------------------------------

AutoSlot_PerfLogPath() {
    return A_ScriptDir "\.cursor\autoslot_perf.log"
}

AutoSlot_PerfLog(hwnd, phase, detail := "") {
    if (!AutoSlot_PERF_LOG)
        return
    global g_AutoSlotPerfOrigin
    now := A_TickCount
    hwnd := Integer(hwnd)
    if (!g_AutoSlotPerfOrigin.Has(hwnd))
        g_AutoSlotPerfOrigin[hwnd] := now
    delta := now - g_AutoSlotPerfOrigin[hwnd]
    stamp := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss")
    line := stamp " +" delta "ms hwnd=" hwnd " " phase
    if (detail != "")
        line .= " " detail
    path := AutoSlot_PerfLogPath()
    try {
        DirCreate(A_ScriptDir "\.cursor")
        FileAppend(line "`n", path, "UTF-8")
    } catch as err {
        try OutputDebug("AutoSlot_PerfLog failed: " err.Message " phase=" phase)
        catch {
        }
    }
}

AutoSlot_PerfClearOrigin(hwnd) {
    global g_AutoSlotPerfOrigin
    if (hwnd && g_AutoSlotPerfOrigin.Has(hwnd))
        g_AutoSlotPerfOrigin.Delete(hwnd)
}

AutoSlot_PerfLogGlobal(phase, detail := "") {
    if (!AutoSlot_PERF_LOG)
        return
    stamp := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss")
    line := stamp " +0ms " phase
    if (detail != "")
        line .= " " detail
    path := AutoSlot_PerfLogPath()
    try {
        DirCreate(A_ScriptDir "\.cursor")
        FileAppend(line "`n", path, "UTF-8")
    } catch {
    }
}

AutoSlot_PerfSessionStart() {
    if (!AutoSlot_PERF_LOG)
        return
    path := AutoSlot_PerfLogPath()
    try {
        DirCreate(A_ScriptDir "\.cursor")
        try FileDelete(path)
        catch {
        }
        envLabel := (IsSet(IS_WORK_ENVIRONMENT) && IS_WORK_ENVIRONMENT) ? "work" : "personal"
        monCount := 0
        try monCount := MonitorGetCount()
        catch
            monCount := 0
        AutoSlot_PerfLogGlobal("session_start",
            "env=" envLabel " computer=" A_ComputerName " monitors=" monCount
            " debounce=" AutoSlot_DEBOUNCE_MS " elig_poll=" AutoSlot_ELIG_RETRY_POLL_MS
            " elig_budget=" AutoSlot_ELIG_RETRY_BUDGET_MS)
    } catch {
    }
}

; --- Place critical (block WinEvent SHOW reentrancy during Place) ------------

AutoSlot_BeginPlaceCritical() {
    global g_AutoSlotPlaceDepth
    Critical "On"
    g_AutoSlotPlaceDepth += 1
}

AutoSlot_EndPlaceCritical() {
    global g_AutoSlotPlaceDepth
    g_AutoSlotPlaceDepth := Max(0, g_AutoSlotPlaceDepth - 1)
    if (g_AutoSlotPlaceDepth = 0)
        Critical "Off"
}

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

; Act / #SingleInstance reload can deliver shell/WinEvent callbacks while maps are unset.
AutoSlot_EnsureMaps() {
    global g_AutoSlotPending, g_AutoSlotRecent, g_AutoSlotHealPending, g_AutoSlotHwndMon
    global g_AutoSlotFillCooldown, g_AutoSlotSnapPairs, g_AutoSlotMaxCompanionSuppress
    global g_AutoSlotPairMaxPending, g_AutoSlotJustRestored, g_AutoSlotWasF11
    global g_AutoSlotF11RestorePending, g_AutoSlotPerfOrigin, g_AutoSlotEligRetry
    global g_AutoSlotShowPending, g_AutoSlotShowTimerArmed, g_AutoSlotPlaceDepth
    global g_AutoSlotLastDestroyHwnd, g_AutoSlotLastDestroyTick, g_AutoSlotPlaceFreezeUntil
    global g_AutoSlotSwapQuietUntil, g_AutoSlotLastToastMsg, g_AutoSlotLastToastTick
    if (!IsSet(g_AutoSlotPending) || !(g_AutoSlotPending is Map))
        g_AutoSlotPending := Map()
    if (!IsSet(g_AutoSlotRecent) || !(g_AutoSlotRecent is Map))
        g_AutoSlotRecent := Map()
    if (!IsSet(g_AutoSlotHealPending) || !(g_AutoSlotHealPending is Map))
        g_AutoSlotHealPending := Map()
    if (!IsSet(g_AutoSlotHwndMon) || !(g_AutoSlotHwndMon is Map))
        g_AutoSlotHwndMon := Map()
    if (!IsSet(g_AutoSlotFillCooldown) || !(g_AutoSlotFillCooldown is Map))
        g_AutoSlotFillCooldown := Map()
    if (!IsSet(g_AutoSlotSnapPairs) || !(g_AutoSlotSnapPairs is Map))
        g_AutoSlotSnapPairs := Map()
    if (!IsSet(g_AutoSlotMaxCompanionSuppress) || !(g_AutoSlotMaxCompanionSuppress is Map))
        g_AutoSlotMaxCompanionSuppress := Map()
    if (!IsSet(g_AutoSlotPairMaxPending) || !(g_AutoSlotPairMaxPending is Map))
        g_AutoSlotPairMaxPending := Map()
    if (!IsSet(g_AutoSlotJustRestored) || !(g_AutoSlotJustRestored is Map))
        g_AutoSlotJustRestored := Map()
    if (!IsSet(g_AutoSlotWasF11) || !(g_AutoSlotWasF11 is Map))
        g_AutoSlotWasF11 := Map()
    if (!IsSet(g_AutoSlotF11RestorePending) || !(g_AutoSlotF11RestorePending is Map))
        g_AutoSlotF11RestorePending := Map()
    if (!IsSet(g_AutoSlotPerfOrigin) || !(g_AutoSlotPerfOrigin is Map))
        g_AutoSlotPerfOrigin := Map()
    if (!IsSet(g_AutoSlotEligRetry) || !(g_AutoSlotEligRetry is Map))
        g_AutoSlotEligRetry := Map()
    if (!IsSet(g_AutoSlotShowPending) || !(g_AutoSlotShowPending is Map))
        g_AutoSlotShowPending := Map()
    if (!IsSet(g_AutoSlotShowTimerArmed))
        g_AutoSlotShowTimerArmed := false
    if (!IsSet(g_AutoSlotPlaceDepth))
        g_AutoSlotPlaceDepth := 0
    if (!IsSet(g_AutoSlotLastDestroyHwnd))
        g_AutoSlotLastDestroyHwnd := 0
    if (!IsSet(g_AutoSlotLastDestroyTick))
        g_AutoSlotLastDestroyTick := 0
    if (!IsSet(g_AutoSlotPlaceFreezeUntil))
        g_AutoSlotPlaceFreezeUntil := 0
    if (!IsSet(g_AutoSlotSwapQuietUntil))
        g_AutoSlotSwapQuietUntil := 0
    if (!IsSet(g_AutoSlotLastToastMsg))
        g_AutoSlotLastToastMsg := ""
    if (!IsSet(g_AutoSlotLastToastTick))
        g_AutoSlotLastToastTick := 0
}

AutoSlot_Init() {
    global g_AutoSlotHook, g_AutoSlotHookCb, g_AutoSlotShellMsg, g_AutoSlotGui
    global g_AutoSlotLocHook, g_AutoSlotLocHookCb, g_AutoSlotMoveHook, g_AutoSlotMoveHookCb
    global g_AutoSlotMinHook, g_AutoSlotMinHookCb
    AutoSlot_EnsureMaps()
    if ((IsSet(g_AutoSlotHook) && g_AutoSlotHook) || (IsSet(g_AutoSlotShellMsg) && g_AutoSlotShellMsg))
        return

    AutoSlot_LoadEnabled()
    AutoSlot_PerfSessionStart()

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
    global g_AutoSlotPlaceMsg
    g_AutoSlotPlaceMsg := DllCall("RegisterWindowMessage", "str", AUTOSLOT_PLACE_MSG_NAME, "uint")
    if (g_AutoSlotPlaceMsg)
        OnMessage(g_AutoSlotPlaceMsg, AutoSlot_OnPlaceHwndMsg)
    AutoSlot_SeedHwndMonFromOccupancy()
    SetTimer(AutoSlot_CheckPlaceRequest, AutoSlot_PLACE_REQUEST_POLL_MS)
}

; PostMessage from Shift keys / Utils (lParam = hwnd to place). Preferred over file IPC.
AutoSlot_OnPlaceHwndMsg(wParam, lParam, msg, hwnd) {
    target := 0
    try target := Integer(lParam)
    catch
        target := 0
    if (target)
        AutoSlot_HandlePlaceRequest(target)
    return 0
}

; Consume cross-process place requests (e.g. Study Topic QuickLook from Shift keys).
AutoSlot_CheckPlaceRequest(*) {
    global g_AutoSlotPlaceRequestFile
    if (!FileExist(g_AutoSlotPlaceRequestFile))
        return
    raw := ""
    try raw := Trim(FileRead(g_AutoSlotPlaceRequestFile, "UTF-8"))
    catch
        raw := ""
    try FileDelete(g_AutoSlotPlaceRequestFile)
    catch {
    }
    if (raw = "")
        return
    hwnd := 0
    try hwnd := Integer(raw)
    catch
        hwnd := 0
    if (hwnd)
        AutoSlot_HandlePlaceRequest(hwnd, "file")
}

; Shared by PostMessage + file poll. QL gets sticky re-place (BeginShow.PositionWindow undoes size).
AutoSlot_HandlePlaceRequest(hwnd, path := "postmessage") {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return
    if (!AutoSlot_IsEnabled() || MonitorGetCount() <= 1)
        return
    AutoSlot_PerfLog(hwnd, "HandlePlaceRequest", "ipc path=" path)
    placed := AutoSlot_TryPlaceBackgroundHwnd(hwnd)
    if (!WinExist("ahk_id " hwnd))
        return
    if (placed && AutoSlot_IsQuickLookHwnd(hwnd)) {
        AutoSlot_ArmStickyPlace(hwnd)
        return
    }
    h := hwnd
    SetTimer(() => AutoSlot_VerifyPlaceOrRetry(h), -800)
}

AutoSlot_IsQuickLookHwnd(hwnd) {
    if (!hwnd)
        return false
    try return StrLower(WinGetProcessName("ahk_id " hwnd)) = "quicklook.exe"
    catch
        return false
}

; True when place geometry looks settled (full or still a half-pane). SnapPairs alone is not enough —
; QuickLook can keep the pair map entry while PositionWindow restores preferred size.
AutoSlot_PlaceLooksSettled(hwnd) {
    global g_AutoSlotSnapPairs
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return true
    monIdx := AutoSlot_GetHwndMonitorIndex(hwnd)
    if (monIdx >= 1 && AutoSlot_CompanionAlreadyFilled(hwnd, monIdx))
        return true
    if (!g_AutoSlotSnapPairs.Has(hwnd) || monIdx < 1)
        return false
    try {
        MonitorGetWorkArea monIdx, &wl, &wt, &wr, &wb
        workW := wr - wl
        if (workW < 1)
            return false
        rect := Buffer(16, 0)
        if !DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect)
            return false
        w := NumGet(rect, 8, "int") - NumGet(rect, 0, "int")
        ; Half pane ≈ 40–60% of work width (gapless snap).
        return (w >= Round(workW * 0.38) && w <= Round(workW * 0.62))
    } catch {
        return false
    }
}

; Second pass if QL undid maximize/move after load (skip only when geometry settled).
AutoSlot_VerifyPlaceOrRetry(hwnd) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return
    if (!AutoSlot_IsEnabled() || MonitorGetCount() <= 1)
        return
    if (AutoSlot_PlaceLooksSettled(hwnd))
        return
    AutoSlot_TryPlaceBackgroundHwnd(hwnd)
    if (AutoSlot_IsQuickLookHwnd(hwnd) && !AutoSlot_PlaceLooksSettled(hwnd))
        AutoSlot_ArmStickyPlace(hwnd)
}

; Re-apply place while QuickLook fights external resize (until maximized / half sticks).
AutoSlot_ArmStickyPlace(hwnd) {
    global g_AutoSlotStickyHwnd, g_AutoSlotStickyAttempts
    if (!hwnd)
        return
    g_AutoSlotStickyHwnd := Integer(hwnd)
    g_AutoSlotStickyAttempts := 0
    SetTimer(AutoSlot_StickyPlaceTick, 0)
    SetTimer(AutoSlot_StickyPlaceTick, -AutoSlot_STICKY_PLACE_MS)
}

AutoSlot_StickyPlaceTick(*) {
    global g_AutoSlotStickyHwnd, g_AutoSlotStickyAttempts
    hwnd := g_AutoSlotStickyHwnd
    if (!hwnd || !WinExist("ahk_id " hwnd)) {
        g_AutoSlotStickyHwnd := 0
        return
    }
    if (!AutoSlot_IsEnabled() || MonitorGetCount() <= 1) {
        g_AutoSlotStickyHwnd := 0
        return
    }
    if (AutoSlot_PlaceLooksSettled(hwnd)) {
        g_AutoSlotStickyHwnd := 0
        return
    }
    g_AutoSlotStickyAttempts += 1
    if (g_AutoSlotStickyAttempts > AutoSlot_STICKY_PLACE_MAX) {
        g_AutoSlotStickyHwnd := 0
        return
    }
    AutoSlot_TryPlaceBackgroundHwnd(hwnd)
    SetTimer(AutoSlot_StickyPlaceTick, -AutoSlot_STICKY_PLACE_MS)
}

AutoSlot_OnDisplayChange(*) {
    return 0
}

AutoSlot_OnShellHook(wParam, lParam, *) {
    AutoSlot_EnsureMaps()
    if (!lParam)
        return 0
    hwnd := Integer(lParam)
    if (wParam = AutoSlot_HSHELL_WINDOWCREATED) {
        AutoSlot_PerfLog(hwnd, "ShellCREATED")
        AutoSlot_Schedule(hwnd)
    } else if (wParam = AutoSlot_HSHELL_WINDOWDESTROYED)
        AutoSlot_OnDestroy(hwnd, true)  ; shell is primary fill trigger
    return 0
}

AutoSlot_OnWinEvent(hWinEventHook, event, hwnd, idObject, idChild, idEventThread, dwmsEventTime) {
    AutoSlot_EnsureMaps()
    if (idObject != AutoSlot_OBJID_WINDOW || !hwnd)
        return
    hwnd := Integer(hwnd)
    if (event = AutoSlot_EVENT_OBJECT_SHOW)
        AutoSlot_QueueScheduleFromShow(hwnd)
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
    AutoSlot_EnsureMaps()
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

    ; Premature WinEvent DESTROY: hwnd still a visible occupant — do not unpair / heal
    ; (would maximize the companion over a still-living 50/50 half).
    if (!fromShell && DllCall("IsWindow", "ptr", hwnd) && AutoSlot_IsOccupancyCandidate(hwnd))
        return

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
    ; Closing one half of a registered pair → maximize that partner when alone.
    ; Debounce + retry + late lone-heal for zombie HWND races (canon late HealLone).
    if (partner) {
        p := partner
        m := (partnerMon >= 1) ? partnerMon : monIdx
        SetTimer(() => AutoSlot_HealKnownCompanion(p, m), -AutoSlot_DEBOUNCE_MS)
        SetTimer(() => AutoSlot_HealKnownCompanion(p, m), -AutoSlot_FILL_RETRY_MS)
    }
    healMon := (partnerMon >= 1) ? partnerMon : monIdx
    AutoSlot_ScheduleHealOnly(healMon)
    SetTimer(() => AutoSlot_HealLoneCompanion(healMon), -(AutoSlot_FILL_RETRY_MS + 200))
    ; No automatic background import — user fills with Ctrl+Alt+Win+6.
}

; Maximize a known leftover companion after its snap-pair partner closed.
; Abort only if another *uncovered* occupant remains (living 50/50 peer / filled).
; Covered-behind noise must not block heal.
AutoSlot_HealKnownCompanion(companion, monIdx := 0) {
    global g_AutoSlotLastDestroyHwnd, g_AutoSlotLastDestroyTick
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
    ; Full-monitor uncovered set (includes companion); any other visible peer blocks.
    ; Skip the just-closed HWND while it may still linger in occupancy.
    skipHwnd := 0
    if (g_AutoSlotLastDestroyHwnd && A_TickCount - g_AutoSlotLastDestroyTick < AutoSlot_DESTROY_DEDUP_MS)
        skipHwnd := g_AutoSlotLastDestroyHwnd
    uncovered := AutoSlot_OccupancyRowsUncovered(AutoSlot_OccupancyOnMonitor(monIdx))
    for row in uncovered {
        if (!row.hwnd || row.hwnd = companion)
            continue
        if (skipHwnd && row.hwnd = skipHwnd)
            continue
        return false
    }
    healed := false
    try healed := !!WM_MaximizeHwndBackground(companion)
    catch
        healed := false
    if (!healed) {
        try AutoSlot_MaximizeHwnd(companion)
        catch {
        }
    }
    if (!healed) {
        try WinMaximize("ahk_id " companion)
        catch {
        }
    }
    if (!AutoSlot_CompanionAlreadyFilled(companion, monIdx))
        return false
    AutoSlot_RememberHwndMon(companion)
    AutoSlot_UnregisterSnapPair(companion)
    AutoSlot_PairSuppressMark(companion, AutoSlot_RECENT_MS)
    AutoSlot_ClaimMonitor(monIdx)
    return true
}

; --- 50/50 pair registry (maximize one half → unregister pair; heal on close) -

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
    AutoSlot_EnsureMaps()
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

; When hwnd is (or became) maximized, break the 50/50 pair — do NOT maximize the companion.
; User may maximize one half for focus; leftover half stays. Ctrl+Alt+Win+6 fills free capacity.
AutoSlot_OnPairedMaximize(hwnd) {
    global g_AutoSlotSnapPairs, g_AutoSlotWasF11, g_AutoSlotF11RestorePending
    if (!AutoSlot_IsEnabled() || !hwnd)
        return false
    if (AutoSlot_SwapQuietActive())
        return false
    if (AutoSlot_PairSuppressActive(hwnd))
        return false
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
    try {
        if (WinGetMinMax("ahk_id " hwnd) != 1)
            return false
    } catch {
        return false
    }
    AutoSlot_UnregisterSnapPair(hwnd)
    return true
}

AutoSlot_ClearPairMaxPending(hwnd) {
    global g_AutoSlotPairMaxPending
    if (hwnd && g_AutoSlotPairMaxPending.Has(hwnd))
        g_AutoSlotPairMaxPending.Delete(hwnd)
}

; --- Swap quiet (no post-quiet background import) ----------------------------

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

; Kept as no-ops so older call sites do not import backgrounds — removed; explicit fill only (Ctrl+Alt+Win+6).
; (ScheduleRearrange / ScheduleFill stubs deleted in efficiency pass.)

; Ctrl+Alt+Win+6 when AutoSlot ON: fill underfilled ordinal slots from background only
; (never relocates windows already occupying slots). Returns { ok, message }.
; Prefer empty monitors before splitting a lone maximized (half-slot).
; Candidates: hidden first, then visible unslotted. One collect + optional QC backfill.

; HWNDs that currently occupy AutoSlot slots on ordinal monitors (stay put during fill).
; Covered-behind a lone max are NOT slotted. Up to 2 uncovered occupants per ordinal are.
AutoSlot_BuildSlottedHwndSet() {
    slotted := Map()
    ordinalCount := Min(MonitorGetCount(), AutoSlot_MAX_ORDINAL)
    loop ordinalCount {
        monIdx := AutoSlot_GetMonitorIndexByOrder(A_Index)
        if (!monIdx)
            continue
        part := AutoSlot_PartitionUncoveredOccupancy(monIdx)
        if (part.filled.Length >= 1) {
            for row in part.filled {
                if (row.hwnd)
                    slotted[Integer(row.hwnd)] := true
            }
            continue
        }
        if (part.nonFilled.Length = 1) {
            if (part.nonFilled[1].hwnd)
                slotted[Integer(part.nonFilled[1].hwnd)] := true
            continue
        }
        if (part.nonFilled.Length >= 2) {
            n := Min(2, part.nonFilled.Length)
            loop n {
                h := part.nonFilled[A_Index].hwnd
                if (h)
                    slotted[Integer(h)] := true
            }
        }
    }
    return slotted
}

; Preview first few candidate titles for fill QC log.
AutoSlot_FillCandPreview(rows, maxN := 3) {
    parts := []
    n := 0
    if (!IsObject(rows))
        return ""
    for row in rows {
        if (n >= maxN)
            break
        title := ""
        exe := ""
        try title := row.title
        catch {
        }
        try exe := row.exe
        catch {
        }
        if (title = "" && exe = "")
            continue
        short := title != "" ? title : exe
        if (StrLen(short) > 28)
            short := SubStr(short, 1, 28) "…"
        short := StrReplace(StrReplace(StrReplace(short, "`n", " "), " ", "_"), ",", ".")
        parts.Push(short)
        n++
    }
    out := ""
    for p in parts {
        out .= (out = "" ? "" : ",") p
    }
    return out
}

AutoSlot_FillOrderLabel(monIdx) {
    order := AutoSlot_OrderForMonitorIndex(monIdx)
    return order > 0 ? order : monIdx
}

; True when any fill-candidate window center already lies on monIdx (same-monitor affinity).
AutoSlot_FillHasSameMonCandidate(monIdx, rows) {
    if (monIdx < 1 || !IsObject(rows))
        return false
    for row in rows {
        hwnd := 0
        try hwnd := Integer(row.hwnd)
        catch
            continue
        if (!hwnd || !DllCall("IsWindow", "ptr", hwnd))
            continue
        if (AutoSlot_GetHwndMonitorIndex(hwnd) = monIdx)
            return true
    }
    ; Also occupancy extras on this monitor (covered behind a max, not yet in rows).
    try {
        for occ in AutoSlot_OccupancyOnMonitor(monIdx) {
            if (!occ.hwnd)
                continue
            if (AutoSlot_CompanionAlreadyFilled(occ.hwnd, monIdx))
                continue
            return true
        }
    } catch {
    }
    return false
}

; Split halfMons: monitors with a local BG/extra first, then the rest (ordinal preserved within each).
AutoSlot_OrderHalfMonsSameMonFirst(halfMons, bgRows) {
    sameFirst := []
    rest := []
    for monIdx in halfMons {
        if (AutoSlot_FillHasSameMonCandidate(monIdx, bgRows))
            sameFirst.Push(monIdx)
        else
            rest.Push(monIdx)
    }
    out := []
    for monIdx in sameFirst
        out.Push(monIdx)
    for monIdx in rest
        out.Push(monIdx)
    return out
}

; After BG imports: expand any remaining lone half to full monitor (global expansion rule).
AutoSlot_ExpandUnderfilledAfterFill() {
    expanded := 0
    ordinalCount := Min(MonitorGetCount(), AutoSlot_MAX_ORDINAL)
    loop ordinalCount {
        monIdx := AutoSlot_GetMonitorIndexByOrder(A_Index)
        if (!monIdx)
            continue
        part := AutoSlot_PartitionUncoveredOccupancy(monIdx)
        ; Lone half (not lone max — max already fills the visual monitor).
        if (part.filled.Length != 0 || part.nonFilled.Length != 1)
            continue
        residual := part.nonFilled[1].hwnd
        if (!residual || AutoSlot_IsClipAngelHwnd(residual))
            continue
        if (AutoSlot_HealLoneCompanion(monIdx)) {
            order := AutoSlot_OrderForMonitorIndex(monIdx)
            label := order > 0 ? order : monIdx
            AutoSlot_Toast("ℹ️ Slot filled → M" label " (maximized)")
            expanded++
        }
    }
    return expanded
}

; Fill candidates: same hidden pool as Ctrl+Alt+Win+Y first, then visible unslotted floats.
; Hidden/Y-list rows are never dropped for "slotted" (list is the stay-put boundary for visible
; occupants only). Pick still rejects currently slotted hwnds as defense.
AutoSlot_CollectFillCandidates() {
    hiddenRows := []
    try hiddenRows := WM_CollectBackgroundWindows()
    catch
        hiddenRows := []
    slotted := AutoSlot_BuildSlottedHwndSet()
    seen := Map()
    rows := []
    for row in hiddenRows {
        hwnd := 0
        try hwnd := Integer(row.hwnd)
        catch
            continue
        if (!hwnd || seen.Has(hwnd))
            continue
        seen[hwnd] := true
        rows.Push(row)
    }
    hiddenCount := rows.Length
    foreHwnd := 0
    try foreHwnd := Integer(WinGetID("A"))
    catch
        foreHwnd := 0
    unslottedCount := 0
    try {
        for hwnd in WinGetList() {
            try hwnd := Integer(hwnd)
            catch
                continue
            if (!hwnd || seen.Has(hwnd))
                continue
            if (foreHwnd && hwnd = foreHwnd)
                continue
            if (slotted.Has(hwnd))
                continue
            if (!AutoSlot_IsOccupancyCandidate(hwnd))
                continue
            title := ""
            exe := ""
            try title := WinGetTitle(hwnd)
            catch
                continue
            if (title = "")
                continue
            try exe := WinGetProcessName("ahk_id " hwnd)
            catch
                exe := ""
            try {
                if (WM_BackgroundTitleIsExcluded(title, exe))
                    continue
            } catch {
            }
            rows.Push({ hwnd: hwnd, title: title, exe: exe })
            seen[hwnd] := true
            unslottedCount++
        }
    } catch {
    }
    return { rows: rows, hiddenCount: hiddenCount, unslottedCount: unslottedCount, slotted: slotted }
}

AutoSlot_ClassifyFillCapacity() {
    emptyMons := []
    halfMons := []
    loneMaxMons := []
    skippedFull := 0
    ordinalCount := Min(MonitorGetCount(), AutoSlot_MAX_ORDINAL)
    loop ordinalCount {
        monIdx := AutoSlot_GetMonitorIndexByOrder(A_Index)
        if (!monIdx)
            continue
        ; Uncovered only — covered-behind windows must not mark a monitor full or non-empty.
        part := AutoSlot_PartitionUncoveredOccupancy(monIdx)
        if (part.filled.Length = 0 && part.nonFilled.Length = 0) {
            emptyMons.Push(monIdx)
            continue
        }
        partner := AutoSlot_FreeHalfPartnerFromPart(part)
        if (partner) {
            halfMons.Push(monIdx)
            if (AutoSlot_CompanionAlreadyFilled(partner, monIdx))
                loneMaxMons.Push(monIdx)
            continue
        }
        skippedFull++
    }
    return {
        emptyMons: emptyMons,
        halfMons: halfMons,
        loneMaxMons: loneMaxMons,
        skippedFull: skippedFull,
        ordinalCount: ordinalCount
    }
}

AutoSlot_FillQualityLogPath() {
    return A_ScriptDir "\.cursor\autoslot_fill_quality.log"
}

AutoSlot_WriteFillQualityLog(stats) {
    path := AutoSlot_FillQualityLogPath()
    try {
        DirCreate(A_ScriptDir "\.cursor")
        stamp := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss")
        line := stamp
            . " pass=" stats.pass
            . " empty=" stats.empty
            . " half=" stats.half
            . " loneMax=" stats.loneMax
            . " skippedFull=" stats.skippedFull
            . " hidden=" stats.hidden
            . " unslotted=" stats.unslotted
            . " cand=" stats.cand
            . " candLeft=" stats.candLeft
            . " filled=" stats.filled
            . " healed=" stats.healed
            . " remainEmpty=" stats.remainEmpty
            . " remainLoneMax=" stats.remainLoneMax
        if (stats.HasProp("reason") && stats.reason != "")
            line .= " reason=" stats.reason
        if (stats.HasProp("mons") && stats.mons != "")
            line .= " mons=" stats.mons
        if (stats.HasProp("preview") && stats.preview != "")
            line .= " preview=" stats.preview
        line .= "`n"
        FileAppend(line, path, "UTF-8")
    } catch {
    }
}

AutoSlot_RunTileBackground() {
    global g_AutoSlotYBgRows, g_AutoSlotYBgActive
    if (!AutoSlot_IsEnabled())
        return { ok: false, message: "AutoSlot is OFF" }
    if (MonitorGetCount() <= 1)
        return { ok: false, message: "Need more than one monitor for slot fill" }

    cap := AutoSlot_ClassifyFillCapacity()
    emptyMons := cap.emptyMons
    halfMons := cap.halfMons
    loneMaxMons := cap.loneMaxMons
    skippedFull := cap.skippedFull
    ordinalCount := cap.ordinalCount

    collect := AutoSlot_CollectFillCandidates()
    bgRows := collect.rows
    hiddenCount := collect.hiddenCount
    unslottedCount := collect.unslottedCount
    candAtStart := bgRows.Length
    candPreview := AutoSlot_FillCandPreview(bgRows)

    ; Always log start so a failed press still leaves a trail (even before any place).
    AutoSlot_WriteFillQualityLog({
        pass: "start",
        empty: emptyMons.Length,
        half: halfMons.Length,
        loneMax: loneMaxMons.Length,
        skippedFull: skippedFull,
        hidden: hiddenCount,
        unslotted: unslottedCount,
        cand: candAtStart,
        candLeft: candAtStart,
        filled: 0,
        healed: 0,
        remainEmpty: emptyMons.Length,
        remainLoneMax: loneMaxMons.Length,
        mons: "",
        preview: candPreview,
        reason: ""
    })

    g_AutoSlotYBgRows := bgRows
    g_AutoSlotYBgActive := true
    filled := 0
    healed := 0
    monResults := []
    ; Prefer free-half monitors that already host a BG/extra (same-monitor pair) before
    ; ordinal steal to another monitor (Scenario A before Scenario B).
    halfMonsOrdered := AutoSlot_OrderHalfMonsSameMonFirst(halfMons, bgRows)
    try {
        for monIdx in emptyMons {
            result := AutoSlot_FillMonitorFromBackground(monIdx, true)
            monResults.Push("M" AutoSlot_FillOrderLabel(monIdx) "=" result)
            if (result = "ok")
                filled++
            else if (result = "healed")
                healed++
        }
        for monIdx in halfMonsOrdered {
            result := AutoSlot_FillMonitorFromBackground(monIdx, true)
            monResults.Push("M" AutoSlot_FillOrderLabel(monIdx) "=" result)
            if (result = "healed")
                healed++
            else if (result = "ok")
                filled++
        }

        monsStr := ""
        for r in monResults
            monsStr .= (monsStr = "" ? "" : ",") r
        capAfter := AutoSlot_ClassifyFillCapacity()
        candLeft := IsObject(g_AutoSlotYBgRows) ? g_AutoSlotYBgRows.Length : 0
        AutoSlot_WriteFillQualityLog({
            pass: "first",
            empty: emptyMons.Length,
            half: halfMons.Length,
            loneMax: loneMaxMons.Length,
            skippedFull: skippedFull,
            hidden: hiddenCount,
            unslotted: unslottedCount,
            cand: candAtStart,
            candLeft: candLeft,
            filled: filled,
            healed: healed,
            remainEmpty: capAfter.emptyMons.Length,
            remainLoneMax: capAfter.loneMaxMons.Length,
            mons: monsStr,
            preview: candPreview,
            reason: ""
        })

        ; QC backfill: remaining empty / free-half (BG snap may still apply).
        needBackfill := candLeft > 0 && (capAfter.emptyMons.Length > 0
            || capAfter.loneMaxMons.Length > 0 || capAfter.halfMons.Length > 0)
        if (needBackfill) {
            filledBefore := filled
            healedBefore := healed
            bfResults := []
            bfHalf := AutoSlot_OrderHalfMonsSameMonFirst(capAfter.halfMons, g_AutoSlotYBgRows)
            for monIdx in capAfter.emptyMons {
                result := AutoSlot_FillMonitorFromBackground(monIdx, true)
                bfResults.Push("M" AutoSlot_FillOrderLabel(monIdx) "=" result)
                if (result = "ok")
                    filled++
                else if (result = "healed")
                    healed++
            }
            for monIdx in bfHalf {
                result := AutoSlot_FillMonitorFromBackground(monIdx, true)
                bfResults.Push("M" AutoSlot_FillOrderLabel(monIdx) "=" result)
                if (result = "healed")
                    healed++
                else if (result = "ok")
                    filled++
            }
            bfMons := ""
            for r in bfResults
                bfMons .= (bfMons = "" ? "" : ",") r
            capFinal := AutoSlot_ClassifyFillCapacity()
            AutoSlot_WriteFillQualityLog({
                pass: "backfill",
                empty: capAfter.emptyMons.Length,
                half: capAfter.halfMons.Length,
                loneMax: capAfter.loneMaxMons.Length,
                skippedFull: capAfter.skippedFull,
                hidden: hiddenCount,
                unslotted: unslottedCount,
                cand: candAtStart,
                candLeft: IsObject(g_AutoSlotYBgRows) ? g_AutoSlotYBgRows.Length : 0,
                filled: filled - filledBefore,
                healed: healed - healedBefore,
                remainEmpty: capFinal.emptyMons.Length,
                remainLoneMax: capFinal.loneMaxMons.Length,
                mons: bfMons,
                preview: candPreview,
                reason: ""
            })
        }

        ; Global expansion: lone half with free second slot → maximize (after BG imports).
        expanded := AutoSlot_ExpandUnderfilledAfterFill()
        if (expanded > 0) {
            healed += expanded
            AutoSlot_WriteFillQualityLog({
                pass: "expand",
                empty: 0,
                half: 0,
                loneMax: 0,
                skippedFull: skippedFull,
                hidden: hiddenCount,
                unslotted: unslottedCount,
                cand: candAtStart,
                candLeft: IsObject(g_AutoSlotYBgRows) ? g_AutoSlotYBgRows.Length : 0,
                filled: 0,
                healed: expanded,
                remainEmpty: 0,
                remainLoneMax: 0,
                mons: "",
                preview: candPreview,
                reason: "expand_lone_half"
            })
        }
    } finally {
        g_AutoSlotYBgActive := false
        g_AutoSlotYBgRows := 0
    }

    detail := " cand=" candAtStart " empty=" emptyMons.Length " half=" halfMons.Length
    total := filled + healed
    if (total = 0) {
        reason := "place_failed"
        if (skippedFull >= ordinalCount && emptyMons.Length = 0 && halfMons.Length = 0)
            reason := "no_capacity"
        else if (candAtStart = 0)
            reason := "no_candidates"
        AutoSlot_WriteFillQualityLog({
            pass: "final",
            empty: emptyMons.Length,
            half: halfMons.Length,
            loneMax: loneMaxMons.Length,
            skippedFull: skippedFull,
            hidden: hiddenCount,
            unslotted: unslottedCount,
            cand: candAtStart,
            candLeft: 0,
            filled: 0,
            healed: 0,
            remainEmpty: emptyMons.Length,
            remainLoneMax: loneMaxMons.Length,
            mons: "",
            preview: candPreview,
            reason: reason
        })
        if (reason = "no_capacity")
            return { ok: false, message: "All ordinal slots full — nothing to fill" }
        if (reason = "no_candidates") {
            if (halfMons.Length > 0 && emptyMons.Length = 0)
                return { ok: false, message: "Could not expand or fill free half-slots" detail }
            return { ok: false, noBackground: true,
                message: "No background windows to fill free slots (visible floats may already be slotted or excluded)" detail }
        }
        return { ok: false, message: "Could not place into free slots" detail }
    }
    if (filled = 0 && healed > 0)
        msg := "Expanded " healed " half-window(s) to full monitor"
    else {
        msg := "Filled " total " slot group(s)"
        if (filled > 0)
            msg .= " from background"
        if (healed > 0)
            msg .= " (" healed " expanded)"
    }
    msg .= " — slotted windows left in place"
    return { ok: true, message: msg, filled: filled, healed: healed }
}

AutoSlot_OnMoveSizeEnd(hWinEventHook, event, hwnd, idObject, idChild, idEventThread, dwmsEventTime) {
    global g_AutoSlotHwndMon, g_AutoSlotSnapPairs, g_AutoSlotWasF11, g_AutoSlotF11RestorePending
    AutoSlot_EnsureMaps()
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
    AutoSlot_RememberHwndMon(hwnd)
}

; On minimize start: clear pair registry, heal leftover companion (parity with destroy).
AutoSlot_OnMinimize(hWinEventHook, event, hwnd, idObject, idChild, idEventThread, dwmsEventTime) {
    global g_AutoSlotHwndMon, g_AutoSlotSnapPairs
    AutoSlot_EnsureMaps()
    if (idObject != AutoSlot_OBJID_WINDOW || !hwnd)
        return
    if (!AutoSlot_IsEnabled())
        return
    hwnd := Integer(hwnd)
    ; MINIMIZEEND fires when a window is restored from the taskbar, not when it minimizes.
    ; Arm a short guard so the subsequent EVENT_OBJECT_SHOW does not auto-place the window.
    if (event = AutoSlot_EVENT_SYSTEM_MINIMIZEEND) {
        global g_AutoSlotJustRestored
        g_AutoSlotJustRestored[hwnd] := A_TickCount
        return
    }

    cached := g_AutoSlotHwndMon.Has(hwnd) ? g_AutoSlotHwndMon[hwnd] : 0
    partnerHwnd := 0
    partnerMon := 0
    if (g_AutoSlotSnapPairs.Has(hwnd)) {
        partnerHwnd := g_AutoSlotSnapPairs[hwnd]
        if (partnerHwnd && partnerHwnd != hwnd && DllCall("IsWindow", "ptr", partnerHwnd)) {
            partnerMon := AutoSlot_GetHwndMonitorIndex(partnerHwnd)
            if (partnerMon < 1 && g_AutoSlotHwndMon.Has(partnerHwnd))
                partnerMon := g_AutoSlotHwndMon[partnerHwnd]
        } else
            partnerHwnd := 0
    }
    if (!cached && !partnerHwnd && !AutoSlot_IsMinimizeRearrangeCandidate(hwnd))
        return

    monIdx := cached >= 1 ? cached : 0
    if (monIdx < 1)
        monIdx := AutoSlot_GetHwndMonitorIndex(hwnd)
    if (monIdx < 1 && partnerMon >= 1)
        monIdx := partnerMon

    ; Always break the 50/50 link when one side minimizes.
    AutoSlot_UnregisterSnapPair(hwnd)
    AutoSlot_ForgetHwndMon(hwnd)
    if (monIdx < 1 || MonitorGetCount() <= 1)
        return

    if (partnerHwnd) {
        p := partnerHwnd
        m := (partnerMon >= 1) ? partnerMon : monIdx
        SetTimer(() => AutoSlot_HealKnownCompanion(p, m), -AutoSlot_DEBOUNCE_MS)
        SetTimer(() => AutoSlot_HealKnownCompanion(p, m), -AutoSlot_FILL_RETRY_MS)
    }
    healMon := (partnerMon >= 1) ? partnerMon : monIdx
    AutoSlot_ScheduleHealOnly(healMon)
    SetTimer(() => AutoSlot_HealLoneCompanion(healMon), -(AutoSlot_FILL_RETRY_MS + 200))
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

; Place claimed this monitor — claim cooldown still used after Place / heal / Y fill.
AutoSlot_ClaimMonitor(monIdx) {
    global g_AutoSlotFillCooldown
    if (monIdx < 1)
        return
    g_AutoSlotFillCooldown[monIdx] := A_TickCount
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
    global g_AutoSlotPending, g_AutoSlotRecent, g_AutoSlotHwndMon, g_AutoSlotJustRestored,
        g_AutoSlotEligRetry
    if (!AutoSlot_IsEnabled()) {
        AutoSlot_PerfLog(hwnd, "Schedule_skip", "disabled")
        return
    }
    if (!hwnd) {
        AutoSlot_PerfLog(0, "Schedule_skip", "invalid_hwnd")
        return
    }
    if (MonitorGetCount() <= 1) {
        AutoSlot_PerfLog(hwnd, "Schedule_skip", "single_monitor")
        return
    }
    ; Already tracked — ignore SHOW spam from activation, not a new window.
    if (g_AutoSlotHwndMon.Has(hwnd)) {
        AutoSlot_PerfLog(hwnd, "Schedule_skip", "already_tracked")
        return
    }
    ; Eligibility settle owns this hwnd — do not arm a parallel debounce.
    if (g_AutoSlotEligRetry.Has(hwnd)) {
        AutoSlot_PerfLog(hwnd, "Schedule_skip", "elig_retry")
        return
    }
    ; Recently restored from taskbar — the SHOW event is from the restore animation,
    ; not a new window opening. Skip auto-placement for the guard duration.
    if (g_AutoSlotJustRestored.Has(hwnd)) {
        if (A_TickCount - g_AutoSlotJustRestored[hwnd] < AutoSlot_RESTORE_GUARD_MS) {
            AutoSlot_PerfLog(hwnd, "Schedule_skip", "restore_guard")
            return
        }
        g_AutoSlotJustRestored.Delete(hwnd)
    }
    if (g_AutoSlotRecent.Has(hwnd) && A_TickCount - g_AutoSlotRecent[hwnd] < AutoSlot_RECENT_MS) {
        AutoSlot_PerfLog(hwnd, "Schedule_skip", "recent")
        return
    }
    if (g_AutoSlotPending.Has(hwnd)) {
        AutoSlot_PerfLog(hwnd, "Schedule_skip", "pending")
        return
    }
    if (!DllCall("IsWindow", "ptr", hwnd)) {
        AutoSlot_PerfLog(hwnd, "Schedule_skip", "invalid_hwnd")
        return
    }
    ; Pending only — RememberHwndMon after eligibility OK (ProcessPending), not here.
    AutoSlot_PerfLog(hwnd, "Schedule", "debounce=" AutoSlot_DEBOUNCE_MS)
    g_AutoSlotPending[hwnd] := true
    SetTimer(() => AutoSlot_ProcessPending(hwnd), -AutoSlot_DEBOUNCE_MS)
}

; SHOW-only: activation of an existing co-occupant (e.g. 50/50 half via ^!#q/w/e/r)
; must not Place/maximize. Shell WINDOWCREATED still uses AutoSlot_Schedule directly.
; Queued from WinEvent — see AutoSlot_QueueScheduleFromShow (do not call synchronously from hook).
AutoSlot_QueueScheduleFromShow(hwnd) {
    global g_AutoSlotShowPending, g_AutoSlotShowTimerArmed
    AutoSlot_EnsureMaps()
    if (!hwnd)
        return
    g_AutoSlotShowPending[Integer(hwnd)] := true
    if (!g_AutoSlotShowTimerArmed) {
        g_AutoSlotShowTimerArmed := true
        SetTimer(AutoSlot_ProcessShowPending, -AutoSlot_SHOW_DEFER_MS)
    }
}

AutoSlot_ProcessShowPending(*) {
    global g_AutoSlotShowPending, g_AutoSlotShowTimerArmed
    AutoSlot_EnsureMaps()
    g_AutoSlotShowTimerArmed := false
    if (!g_AutoSlotShowPending.Count)
        return
    rest := Map()
    n := 0
    for hwnd, _ in g_AutoSlotShowPending {
        if (n >= AutoSlot_SHOW_BATCH_MAX) {
            rest[hwnd] := true
            continue
        }
        n += 1
        if (DllCall("IsWindow", "ptr", hwnd))
            AutoSlot_ScheduleFromShow(hwnd)
    }
    g_AutoSlotShowPending := rest
    if (g_AutoSlotShowPending.Count) {
        g_AutoSlotShowTimerArmed := true
        SetTimer(AutoSlot_ProcessShowPending, -AutoSlot_SHOW_DEFER_MS)
    }
}

AutoSlot_ScheduleFromShow(hwnd) {
    global g_AutoSlotHwndMon, g_AutoSlotEligRetry, g_AutoSlotPlaceDepth
    if (!AutoSlot_IsEnabled()) {
        AutoSlot_PerfLog(hwnd, "ScheduleFromShow_skip", "disabled")
        return
    }
    if (!hwnd) {
        AutoSlot_PerfLog(0, "ScheduleFromShow_skip", "invalid_hwnd")
        return
    }
    if (!DllCall("IsWindow", "ptr", hwnd)) {
        AutoSlot_PerfLog(hwnd, "ScheduleFromShow_skip", "invalid_hwnd")
        return
    }
    if (MonitorGetCount() <= 1) {
        AutoSlot_PerfLog(hwnd, "ScheduleFromShow_skip", "single_monitor")
        return
    }
    if (g_AutoSlotPlaceDepth > 0) {
        AutoSlot_PerfLog(hwnd, "ScheduleFromShow_skip", "place_active")
        return
    }
    if (g_AutoSlotHwndMon.Has(hwnd)) {
        AutoSlot_PerfLog(hwnd, "ScheduleFromShow_skip", "already_tracked")
        return
    }
    if (g_AutoSlotEligRetry.Has(hwnd)) {
        AutoSlot_PerfLog(hwnd, "ScheduleFromShow_skip", "elig_retry")
        return
    }
    monIdx := AutoSlot_GetHwndMonitorIndex(hwnd)
    if (monIdx >= 1) {
        cacheHit := AutoSlot_MonitorOccupiedFromCache(monIdx, hwnd)
        occupied := false
        via := ""
        if (cacheHit = "yes") {
            occupied := true
            via := "cache"
        } else if (cacheHit != "no") {
            if (AutoSlot_MonitorHasTrackedOccupant(monIdx, hwnd)) {
                occupied := true
                via := "tracked"
            } else if (AutoSlot_MonitorHasOtherOccupant(monIdx, hwnd)) {
                occupied := true
                via := "scan"
            }
        }
        if (occupied) {
            ; Already sharing this monitor — do not Place. Cache only eligible
            ; occupants (Teams/#32770 chrome must not seed HwndMon → false heal).
            if (AutoSlot_IsEligibleNewWindow(hwnd) || AutoSlot_IsOccupancyCandidate(hwnd))
                AutoSlot_RememberHwndMon(hwnd)
            detail := "occupied_mon=" monIdx
            if (via != "")
                detail .= " via=" via
            AutoSlot_PerfLog(hwnd, "ScheduleFromShow_skip", detail)
            return
        }
    }
    AutoSlot_PerfLog(hwnd, "ScheduleFromShow", "→Schedule")
    AutoSlot_Schedule(hwnd)
}

; Post-reload: track current occupants so cycle activate cannot Place existing layouts.
AutoSlot_SeedHwndMonFromOccupancy() {
    if (MonitorGetCount() <= 1)
        return
    ordinalCount := Min(MonitorGetCount(), AutoSlot_MAX_ORDINAL)
    loop ordinalCount {
        monIdx := AutoSlot_GetMonitorIndexByOrder(A_Index)
        if (!monIdx)
            continue
        part := AutoSlot_PartitionOccupancy(monIdx)
        for row in part.filled
            AutoSlot_RememberHwndMon(row.hwnd)
        for row in part.nonFilled
            AutoSlot_RememberHwndMon(row.hwnd)
    }
}

; Dense eligibility settle: ~100 ms polls from first miss until budget (~2 s).
; Required — see docs/canon/windows-rearrange.md (Eligibility settle).
AutoSlot_ScheduleEligRetry(hwnd) {
    global g_AutoSlotEligRetry
    if (!hwnd || !DllCall("IsWindow", "ptr", hwnd))
        return
    now := A_TickCount
    if (!g_AutoSlotEligRetry.Has(hwnd))
        g_AutoSlotEligRetry[hwnd] := now
    start := Integer(g_AutoSlotEligRetry[hwnd])
    elapsed := now - start
    if (elapsed >= AutoSlot_ELIG_RETRY_BUDGET_MS) {
        AutoSlot_PerfLog(hwnd, "EligRetry_giveup", "elapsed=" elapsed)
        AutoSlot_ClearEligRetry(hwnd)
        return
    }
    AutoSlot_PerfLog(hwnd, "EligRetry_arm", "poll=" AutoSlot_ELIG_RETRY_POLL_MS " elapsed=" elapsed)
    h := hwnd
    SetTimer(() => AutoSlot_ProcessEligRetry(h), -AutoSlot_ELIG_RETRY_POLL_MS)
}

AutoSlot_ProcessEligRetry(hwnd) {
    if (!hwnd || !DllCall("IsWindow", "ptr", hwnd)) {
        AutoSlot_ClearEligRetry(hwnd)
        return
    }
    AutoSlot_PerfLog(hwnd, "EligRetry_fire")
    AutoSlot_ProcessPending(hwnd)
}

AutoSlot_ClearEligRetry(hwnd) {
    global g_AutoSlotEligRetry
    if (hwnd && g_AutoSlotEligRetry.Has(hwnd))
        g_AutoSlotEligRetry.Delete(hwnd)
}

AutoSlot_ProcessPending(hwnd) {
    global g_AutoSlotPending, g_AutoSlotRecent
    if (g_AutoSlotPending.Has(hwnd))
        g_AutoSlotPending.Delete(hwnd)
    AutoSlot_PerfLog(hwnd, "ProcessPending_enter")
    if (!AutoSlot_IsEnabled()) {
        AutoSlot_PerfLog(hwnd, "ProcessPending_skip", "disabled")
        return
    }
    if (!hwnd || !DllCall("IsWindow", "ptr", hwnd)) {
        AutoSlot_PerfLog(hwnd, "ProcessPending_skip", "invalid_hwnd")
        AutoSlot_ClearEligRetry(hwnd)
        return
    }
    if (MonitorGetCount() <= 1) {
        AutoSlot_PerfLog(hwnd, "ProcessPending_skip", "single_monitor")
        return
    }
    if (g_AutoSlotRecent.Has(hwnd) && A_TickCount - g_AutoSlotRecent[hwnd] < AutoSlot_RECENT_MS) {
        AutoSlot_PerfLog(hwnd, "ProcessPending_skip", "recent")
        AutoSlot_ClearEligRetry(hwnd)
        return
    }
    if (!AutoSlot_IsEligibleNewWindow(hwnd)) {
        ; Do not Remember until eligible. Dense poll until title/HWND ready (budget).
        AutoSlot_PerfLog(hwnd, "ProcessPending_elig_fail")
        AutoSlot_ScheduleEligRetry(hwnd)
        return
    }
    AutoSlot_ClearEligRetry(hwnd)
    AutoSlot_PerfLog(hwnd, "ProcessPending_elig_ok")
    g_AutoSlotRecent[hwnd] := A_TickCount
    AutoSlot_PruneRecent()
    AutoSlot_BeginPlaceCritical()
    try {
        AutoSlot_RememberHwndMon(hwnd)
        AutoSlot_Place(hwnd)
    } finally {
        AutoSlot_EndPlaceCritical()
    }
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
; Win32 common dialogs (#32770), browser PiP, system/shell noise titles, Teams share UI,
; Win+Shift+S screen clip / Snipping Tool, and PowerPoint slide show / Presenter View
; (must not be auto-slotted or resized). ClipAngel is fully excluded — never
; auto-slotted / 50/50'd and never counted as occupancy.
AutoSlot_IsExcludedExeOrTitle(hwnd) {
    if (!hwnd)
        return false
    ; User ignore list from #!+L R (persisted autoslot_user_excludes.ini).
    try {
        if (AutoSlot_UserExcludeMatch(hwnd))
            return true
    } catch {
    }
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
    ; All Win32 common dialogs (MessageBox, Open/Save, Confirm, Print, Properties, etc.).
    if (class = "#32770")
        return true
    ; PowerPoint: keep only the main edit frame (PPTFrameClass).
    ; Slide show (screenClass) and Presenter View (PodiumParent) stay out of Place/occupancy/fill.
    if (exe = "powerpnt.exe" && class != "" && class != "pptframeclass")
        return true
    ; Teams share picker / sharing control bar / presenter toolbar — never resize.
    if (AutoSlot_IsTeamsShareUiHwnd(hwnd))
        return true
    try {
        title := WinGetTitle(hwnd)
    } catch {
        return false
    }
    t := StrLower(title)
    ; Chromium Picture-in-Picture floating player.
    if ((exe = "chrome.exe" || exe = "msedge.exe" || exe = "brave.exe" || exe = "chromium.exe")
    && (InStr(t, "picture in picture") || InStr(t, "picture-in-picture")
    || InStr(t, "imagem na imagem")))
        return true
    ; Shell / toast / widget noise (same needles as WM background scanner).
    try {
        if (WM_BackgroundIsSystemNoiseTitle(title))
            return true
    } catch {
    }
    if (InStr(t, "windowmanagement.ahk"))
        return true
    if (InStr(t, "autohotkey") && InStr(class, "ahk"))
        return true
    return false
}

; Teams chrome that must not be auto-slotted / occupancy / fill.
; Authority: WM_IsTeamsChromeHwnd (chat + full meeting eligible; share bar + compact excluded).
AutoSlot_IsTeamsShareUiHwnd(hwnd) {
    if (!hwnd)
        return false
    try {
        return WM_IsTeamsChromeHwnd(hwnd)
    } catch {
        return false
    }
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

; SHOW occupied gate: "yes" / "no" from fresh cache, "" if stale.
AutoSlot_MonitorOccupiedFromCache(monIdx, excludeHwnd := 0) {
    global g_AutoSlotOccSnap, g_AutoSlotOccSnapTick
    if (!IsObject(g_AutoSlotOccSnap) || A_TickCount - g_AutoSlotOccSnapTick > AutoSlot_OCC_CACHE_MS)
        return ""
    rows := g_AutoSlotOccSnap.Has(monIdx) ? g_AutoSlotOccSnap[monIdx] : []
    for row in rows {
        if (IsObject(row) && row.hwnd && row.hwnd != excludeHwnd)
            return "yes"
    }
    return "no"
}

AutoSlot_MonitorHasTrackedOccupant(monIdx, excludeHwnd := 0) {
    global g_AutoSlotHwndMon
    for h, m in g_AutoSlotHwndMon {
        if (h = excludeHwnd || m != monIdx)
            continue
        if (DllCall("IsWindow", "ptr", h))
            return true
    }
    return false
}

; Early-exit WinGetList: any other occupancy candidate on monIdx (no partition).
AutoSlot_MonitorHasOtherOccupant(monIdx, excludeHwnd := 0) {
    if (monIdx < 1 || monIdx > MonitorGetCount())
        return false
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
            wcx := (NumGet(rect, 0, "int") + NumGet(rect, 8, "int")) // 2
            wcy := (NumGet(rect, 4, "int") + NumGet(rect, 12, "int")) // 2
            wPoint := (wcy & 0xFFFFFFFF) << 32 | (wcx & 0xFFFFFFFF)
            hMon := DllCall("MonitorFromPoint", "int64", wPoint, "uint", 2, "ptr")
            if (Integer(hMon) = Integer(hTarget))
                return true
        } catch {
            continue
        }
    }
    return false
}

; One WinGetList for the whole desktop; bucket occupancy rows by monitor index.
; Place / TryPlace empty+half search use this once (avoid N+M full enumerations).
AutoSlot_BuildOccupancyByMonitor(excludeHwnd := 0) {
    global g_AutoSlotOccSnap, g_AutoSlotOccSnapTick
    byMon := Map()
    count := 0
    try count := MonitorGetCount()
    catch
        count := 0
    if (count < 1)
        return byMon
    monTargets := Map()
    loop count {
        monIdx := A_Index
        byMon[monIdx] := []
        try {
            MonitorGet monIdx, &ml, &mt, &mr, &mb
            cx := (ml + mr) // 2
            cy := (mt + mb) // 2
            point64 := (cy & 0xFFFFFFFF) << 32 | (cx & 0xFFFFFFFF)
            hTarget := Integer(DllCall("MonitorFromPoint", "int64", point64, "uint", 2, "ptr"))
            monTargets[hTarget] := monIdx
        } catch {
        }
    }
    hwndList := WinGetList()
    for hwnd in hwndList {
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
            wcx := (left + right) // 2
            wcy := (top + bottom) // 2
            wPoint := (wcy & 0xFFFFFFFF) << 32 | (wcx & 0xFFFFFFFF)
            hMon := Integer(DllCall("MonitorFromPoint", "int64", wPoint, "uint", 2, "ptr"))
            if (!monTargets.Has(hMon))
                continue
            monIdx := monTargets[hMon]
            byMon[monIdx].Push({
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
    g_AutoSlotOccSnap := byMon
    g_AutoSlotOccSnapTick := A_TickCount
    return byMon
}

; Keep rows whose centers are not inside a higher z-order row's rect (WinGetList order).
; Same heuristic as GetVisibleWindowsOnMonitor — buried noise must not block heal.
AutoSlot_OccupancyRowsUncovered(rows) {
    uncovered := []
    if (!IsObject(rows))
        return uncovered
    for row in rows {
        if (!IsObject(row) || !row.HasProp("hwnd") || !row.hwnd)
            continue
        cx := (row.left + row.right) // 2
        cy := (row.top + row.bottom) // 2
        covered := false
        for u in uncovered {
            if (cx >= u.left && cx <= u.right && cy >= u.top && cy <= u.bottom) {
                covered := true
                break
            }
        }
        if (!covered)
            uncovered.Push(row)
    }
    return uncovered
}

; Partition uncovered occupancy on monIdx (optional excludeHwnd).
AutoSlot_PartitionUncoveredOccupancy(monIdx, excludeHwnd := 0) {
    nonFilled := []
    filled := []
    rows := AutoSlot_OccupancyRowsUncovered(AutoSlot_OccupancyOnMonitor(monIdx, excludeHwnd))
    for row in rows {
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
    StandardLoadingBar_BusyAllMonitors_Begin()
    try {
        try {
            WinMaximize "ahk_id " hwnd
        } catch {
        }
        ; Always reinforce with SC_MAXIMIZE. WinMaximize rarely throws; WPF apps
        ; (QuickLook) need WM_SYSCOMMAND so WindowState=Maximized and PositionWindow
        ; will not undo the size after document load.
        try PostMessage 0x0112, 0xF030, , , "ahk_id " hwnd
    } finally {
        StandardLoadingBar_BusyAllMonitors_End()
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
    StandardLoadingBar_BusyAllMonitors_Begin()
    try {
        MonitorGet monIdx, &left, &top, &right, &bottom
        AutoSlot_MoveHwndToRect(hwnd, left, top, right, bottom)
        AutoSlot_MaximizeHwnd(hwnd)
        AutoSlot_ActivateHwnd(hwnd)
        ; scheduleRearrange kept for call-site compat; auto background import is disabled.
        return true
    } finally {
        StandardLoadingBar_BusyAllMonitors_End()
    }
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

AutoSlot_ShowSwapModal(srcLabel, dstLabel, moverHwnd, displacedHwnds) {
    ; Quiet heal/pair events briefly; no [F] replace — user fills with Y if needed.
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
    AutoSlot_Toast("ℹ️ Swapped M" srcLabel " ↔ M" dstLabel)
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

    StandardLoadingBar_BusyAllMonitors_Begin()
    try {
        return AutoSlot_TryForegroundSwap_Impl(hwnd, sourceMon, destMon, mover, dest, role)
    } finally {
        StandardLoadingBar_BusyAllMonitors_End()
    }
}

AutoSlot_TryForegroundSwap_Impl(hwnd, sourceMon, destMon, mover, dest, role) {
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

    StandardLoadingBar_BusyAllMonitors_Begin()
    try {
        return AutoSlot_SnapPair_Impl(newHwnd, partnerHwnd, monIdx, acceptUnvalidated)
    } finally {
        StandardLoadingBar_BusyAllMonitors_End()
    }
}

AutoSlot_SnapPair_Impl(newHwnd, partnerHwnd, monIdx, acceptUnvalidated := false) {
    global g_AutoSlotUndo
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
            Sleep 25
        }
        if (AutoSlot_CompanionAlreadyFilled(partnerHwnd, monIdx))
            return ""
    }
    ; Close-fill / Place skip strict validate wait (gapless already applied). Manual ^!#x
    ; still uses WM_WaitValidateSnapBipartitionStrict via tile_snap.
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
    AutoSlot_RememberHwndMon(newHwnd)
    AutoSlot_RememberHwndMon(partnerHwnd)
    ; Mute LOCATIONCHANGE settle (stale MinMax=1) so Place 50/50 is not immediately unpaired.
    AutoSlot_ClearPairMaxPending(newHwnd)
    AutoSlot_ClearPairMaxPending(partnerHwnd)
    AutoSlot_PairSuppressMark(newHwnd, AutoSlot_RECENT_MS)
    AutoSlot_PairSuppressMark(partnerHwnd, AutoSlot_RECENT_MS)
    AutoSlot_ActivateHwnd(newHwnd)
    return newPane
}

; Demax OS-maximized and work-area-filled windows so 50/50 MoveWindow sticks.
; Short sleeps: long 80/80/50 stacks made partner shrink visibly long before the new pane moved.
AutoSlot_EnsureRestoredForSnap(hwnd, monIdx := 0) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    changed := false
    try {
        mm := WinGetMinMax("ahk_id " hwnd)
        if (mm = 1 || mm = -1) {
            WinRestore "ahk_id " hwnd
            Sleep 25
            changed := true
        }
        ; Restore-from-minimized often lands maximized — second pass.
        if (WinGetMinMax("ahk_id " hwnd) = 1) {
            WinRestore "ahk_id " hwnd
            Sleep 25
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
                Sleep 15
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
; forceImport: explicit user fill (Ctrl+Alt+Win+6) — ignore place freeze / fill cooldown.
; When g_AutoSlotYBgActive, reuse g_AutoSlotYBgRows (one collect per fill pass); consume picked hwnds.
; Never returns an already-slotted hwnd (defense if candidate pool is stale).
; Prefer candidates whose window center is already on monIdx (same-monitor pairing).
AutoSlot_PickBackgroundCandidate(monIdx, occupancyRows, excludeExtra := 0, forceImport := false) {
    global g_AutoSlotYBgActive, g_AutoSlotYBgRows
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
    slotted := AutoSlot_BuildSlottedHwndSet()
    rows := []
    reuseY := g_AutoSlotYBgActive && IsObject(g_AutoSlotYBgRows)
    if (reuseY)
        rows := g_AutoSlotYBgRows
    else {
        try rows := WM_CollectBackgroundWindows()
        catch
            return 0
    }
    ; Two passes: same-monitor first, then any remaining candidate.
    loop 2 {
        preferSame := (A_Index = 1 && monIdx >= 1)
        i := 1
        while (i <= rows.Length) {
            row := rows[i]
            hwnd := 0
            try hwnd := Integer(row.hwnd)
            catch {
                i++
                continue
            }
            if (!hwnd || !DllCall("IsWindow", "ptr", hwnd)) {
                if (reuseY)
                    rows.RemoveAt(i)
                else
                    i++
                continue
            }
            if (occupied.Has(hwnd) || slotted.Has(hwnd)) {
                i++
                continue
            }
            if (AutoSlot_IsExcludedExeOrTitle(hwnd)) {
                i++
                continue
            }
            if (preferSame) {
                homeMon := AutoSlot_GetHwndMonitorIndex(hwnd)
                if (homeMon != monIdx) {
                    i++
                    continue
                }
            }
            if (!forceImport) {
                if (AutoSlot_BackgroundCandHasLivingSnapPartner(hwnd)) {
                    i++
                    continue
                }
                if (AutoSlot_BackgroundCandCoveredByF11(hwnd)) {
                    i++
                    continue
                }
            } else {
                AutoSlot_UnregisterSnapPair(hwnd)
            }
            if (reuseY)
                rows.RemoveAt(i)
            return hwnd
        }
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
AutoSlot_PartitionOccupancyRows(rows, monIdx) {
    nonFilled := []
    filled := []
    if (!IsObject(rows))
        rows := []
    for row in rows {
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

AutoSlot_PartitionOccupancy(monIdx, excludeHwnd := 0) {
    return AutoSlot_PartitionOccupancyRows(AutoSlot_OccupancyOnMonitor(monIdx, excludeHwnd), monIdx)
}

; Free-half partner from an already-partitioned occupancy (no re-scan).
AutoSlot_FreeHalfPartnerFromPart(part) {
    if (!IsObject(part))
        return 0
    if (part.filled.Length = 1)
        return part.filled[1].hwnd
    if (part.filled.Length = 0 && part.nonFilled.Length = 1)
        return part.nonFilled[1].hwnd
    return 0
}

; Maximize the sole uncovered non-filled window on monIdx.
; Covered-behind rows are ignored; never heal beside an uncovered filled window.
AutoSlot_HealLoneCompanion(monIdx) {
    if (monIdx < 1 || monIdx > MonitorGetCount())
        return false
    part := AutoSlot_PartitionUncoveredOccupancy(monIdx)
    if (part.filled.Length >= 1)
        return false
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
        try AutoSlot_MaximizeHwnd(companion)
        catch {
        }
    }
    ; Confirm work-area / max presentation before claiming success (WinMaximize rarely throws).
    if (!AutoSlot_CompanionAlreadyFilled(companion, monIdx))
        return false
    ; Work-area fill via WM_MaximizeHwndBackground may not set OS maximize, so
    ; OnPairedMaximize would not clear the registry — unregister explicitly.
    AutoSlot_UnregisterSnapPair(companion)
    AutoSlot_PairSuppressMark(companion, AutoSlot_RECENT_MS)
    AutoSlot_ClaimMonitor(monIdx)
    return true
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

; Promote free capacity: empty → import up to two; lone maximized → SnapPair BG;
; lone half + forceImport → SnapPair BG when candidates exist (no heal-only; no local half↔max reshuffle).
; Decisions use UNCOVERED occupancy — windows hidden behind a max do not steal the free-half path.
; forceImport: true for explicit Ctrl+Alt+Win+6 (ignore place freeze / fill cooldown).
AutoSlot_FillMonitorFromBackground(monIdx, forceImport := false) {
    global g_AutoSlotRecent, g_AutoSlotUndo
    if (monIdx < 1 || monIdx > MonitorGetCount() || MonitorGetCount() <= 1)
        return "noop"
    ; Ignore z-order-covered windows for capacity decisions (same as heal / free-half).
    part := AutoSlot_PartitionUncoveredOccupancy(monIdx)
    ; True 50/50 (two uncovered halves, no filled) is full.
    if (part.filled.Length = 0 && part.nonFilled.Length >= 2)
        return "stale"
    if (part.filled.Length >= 2)
        return "ok"

    order := AutoSlot_OrderForMonitorIndex(monIdx)
    label := order > 0 ? order : monIdx

    ; Place freeze / claim cooldown: companion heal only (no background import),
    ; unless the user explicitly requested fill (Ctrl+Alt+Win+6).
    blockImport := !forceImport && (AutoSlot_PlaceFreezeActive() || AutoSlot_FillCooldownActive(monIdx))

    if (part.nonFilled.Length = 1) {
        residual := part.nonFilled[1].hwnd
        if (AutoSlot_IsClipAngelHwnd(residual))
            return "noop"
        if (AutoSlot_CompanionAlreadyFilled(residual, monIdx))
            return "ok"
        ; Half + max on same monitor.
        if (part.filled.Length >= 1) {
            filledHwnd := part.filled[1].hwnd
            if (forceImport) {
                ; Same-monitor pair (Scenario A): snap the on-monitor extra with the max.
                ; Do not leave stale when residual is already a half — that is the desired 50/50.
                if (blockImport || AutoSlot_IsClipAngelHwnd(filledHwnd))
                    return "noop"
                if (filledHwnd && filledHwnd != residual && !AutoSlot_IsClipAngelHwnd(residual)) {
                    g_AutoSlotRecent[residual] := A_TickCount
                    AutoSlot_PruneRecent()
                    pane := AutoSlot_SnapPair(residual, filledHwnd, monIdx, true)
                    g_AutoSlotUndo := 0
                    if (pane != "") {
                        AutoSlot_RememberHwndMon(residual)
                        AutoSlot_RememberHwndMon(filledHwnd)
                        AutoSlot_PairSuppressMark(residual, AutoSlot_RECENT_MS)
                        AutoSlot_PairSuppressMark(filledHwnd, AutoSlot_RECENT_MS)
                        AutoSlot_ClaimMonitor(monIdx)
                        AutoSlot_Toast("ℹ️ Slot filled → M" label " (50/50)")
                        return "ok"
                    }
                }
                ; Residual not pairable — import a different BG beside the max.
                excludeRows := [{ hwnd: filledHwnd }, { hwnd: residual }]
                cand := AutoSlot_PickBackgroundCandidate(monIdx, excludeRows, 0, true)
                if (!cand)
                    return "noop"
                g_AutoSlotRecent[cand] := A_TickCount
                AutoSlot_PruneRecent()
                pane := AutoSlot_SnapPair(cand, filledHwnd, monIdx, true)
                g_AutoSlotUndo := 0
                if (pane = "")
                    return "noop"
                AutoSlot_RememberHwndMon(cand)
                AutoSlot_RememberHwndMon(filledHwnd)
                AutoSlot_PairSuppressMark(cand, AutoSlot_RECENT_MS)
                AutoSlot_PairSuppressMark(filledHwnd, AutoSlot_RECENT_MS)
                AutoSlot_ClaimMonitor(monIdx)
                AutoSlot_Toast("ℹ️ Slot filled → M" label " (50/50)")
                return "ok"
            }
            if (!blockImport) {
                if (filledHwnd && filledHwnd != residual && !AutoSlot_IsClipAngelHwnd(filledHwnd)) {
                    g_AutoSlotRecent[residual] := A_TickCount
                    AutoSlot_PruneRecent()
                    pane := AutoSlot_SnapPair(residual, filledHwnd, monIdx, true)
                    g_AutoSlotUndo := 0
                    if (pane != "") {
                        AutoSlot_RememberHwndMon(residual)
                        AutoSlot_RememberHwndMon(filledHwnd)
                        AutoSlot_PairSuppressMark(residual, AutoSlot_RECENT_MS)
                        AutoSlot_PairSuppressMark(filledHwnd, AutoSlot_RECENT_MS)
                        AutoSlot_ClaimMonitor(monIdx)
                        AutoSlot_Toast("ℹ️ Slot filled → M" label " (50/50)")
                        return "ok"
                    }
                }
            }
            return "noop"
        }
        ; Lone half + free slot: explicit fill imports BG only (no expand-only heal).
        residualRows := [{ hwnd: residual }]
        if (forceImport && !blockImport) {
            cand := AutoSlot_PickBackgroundCandidate(monIdx, residualRows, 0, true)
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
            AutoSlot_PairSuppressMark(cand, AutoSlot_RECENT_MS)
            AutoSlot_PairSuppressMark(residual, AutoSlot_RECENT_MS)
            AutoSlot_ClaimMonitor(monIdx)
            AutoSlot_Toast("ℹ️ Slot filled → M" label " (50/50)")
            return "ok"
        }
        if (AutoSlot_HealLoneCompanion(monIdx))
            return "healed"
        if (blockImport)
            return "noop"
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
        AutoSlot_PairSuppressMark(cand, AutoSlot_RECENT_MS)
        AutoSlot_PairSuppressMark(residual, AutoSlot_RECENT_MS)
        AutoSlot_ClaimMonitor(monIdx)
        AutoSlot_Toast("ℹ️ Slot filled → M" label " (50/50)")
        return "ok"
    }

    ; Lone maximized / work-area fill = half-full — SnapPair background (covered behind ignored).
    ; Also when filled=1 and nonFilled>=2: extras excluded from pick; still import BG.
    if (part.filled.Length = 1) {
        residual := part.filled[1].hwnd
        if (AutoSlot_IsClipAngelHwnd(residual))
            return "noop"
        if (blockImport)
            return "noop"
        residualRows := [{ hwnd: residual }]
        for row in part.nonFilled {
            if (row.hwnd)
                residualRows.Push({ hwnd: row.hwnd })
        }
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
        AutoSlot_PairSuppressMark(cand, AutoSlot_RECENT_MS)
        AutoSlot_PairSuppressMark(residual, AutoSlot_RECENT_MS)
        AutoSlot_ClaimMonitor(monIdx)
        AutoSlot_Toast("ℹ️ Slot filled → M" label " (50/50)")
        return "ok"
    }

    if (part.filled.Length >= 2)
        return "ok"

    ; Empty (uncovered) monitor — fill both slots when two candidates exist.
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
    }
    g_AutoSlotRecent[cand1] := A_TickCount
    AutoSlot_PruneRecent()
    if (!AutoSlot_MaximizeOnMonitor(cand1, monIdx))
        return "noop"
    AutoSlot_RememberHwndMon(cand1)
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
; Used by Ctrl+Alt+Win+Y list open when AutoSlot is ON.
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

    AutoSlot_BeginPlaceCritical()
    try {
        ordinalCount := Min(MonitorGetCount(), AutoSlot_MAX_ORDINAL)
        snap := AutoSlot_BuildOccupancyByMonitor(hwnd)
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
            rows := snap.Has(monIdx) ? snap[monIdx] : []
            part := AutoSlot_PartitionOccupancyRows(rows, monIdx)
            ; Same empty / free-half policy as AutoSlot_Place (empty first, then free half).
            if (part.filled.Length = 0 && part.nonFilled.Length = 0) {
                if (!emptyMon) {
                    emptyMon := monIdx
                    emptyOrder := order
                }
                continue
            }
            if (!halfMon) {
                partner := AutoSlot_FreeHalfPartnerFromPart(part)
                if (partner) {
                    halfMon := monIdx
                    halfOrder := order
                    halfPartner := partner
                }
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
    } finally {
        AutoSlot_EndPlaceCritical()
    }
}

; Partner to 50/50 a new window with on monIdx, ignoring windows hidden behind a
; maximized one: one visible maximized/work-area window with no other half-pane = free half;
; else one lone half. Two filled or two half-panes = genuinely full → 0.
AutoSlot_MonitorFreeHalfPartner(monIdx, excludeHwnd := 0) {
    return AutoSlot_FreeHalfPartnerFromPart(AutoSlot_PartitionUncoveredOccupancy(monIdx, excludeHwnd))
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
; acceptUnvalidated: skip WaitValidate poll (up to ~400 ms) — gapless already placed both panes;
; the strict wait made "partner shrinks, then long pause, then new window" feel on work PCs.
AutoSlot_TrySnapNewWithPartner(hwnd, monIdx, partner, orderLabel := 0) {
    if (!partner || partner = hwnd)
        return false
    pane := AutoSlot_SnapPair(hwnd, partner, monIdx, true)
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

    AutoSlot_BeginPlaceCritical()
    try {
        AutoSlot_PerfLog(hwnd, "Place_enter")
        t0 := A_TickCount
        g_AutoSlotUndo := 0
        msg := ""
        AutoSlot_RememberHwndMon(hwnd)
        AutoSlot_BeginPlaceFreeze()
        AutoSlot_PerfLog(hwnd, "Place_freeze_done", "ms=" (A_TickCount - t0))

        ; One desktop occupancy walk for empty + free-half search (H1).
        snap := AutoSlot_BuildOccupancyByMonitor(hwnd)
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
            rows := snap.Has(monIdx) ? snap[monIdx] : []
            part := AutoSlot_PartitionOccupancyRows(rows, monIdx)
            if (part.filled.Length = 0 && part.nonFilled.Length = 0) {
                if (!emptyMon) {
                    emptyOrder := order
                    emptyMon := monIdx
                }
                continue
            }
            if (!halfPartner) {
                partner := AutoSlot_FreeHalfPartnerFromPart(part)
                if (partner) {
                    halfOrder := order
                    halfMon := monIdx
                    halfPartner := partner
                }
            }
        }
        occDetail := "ms=" (A_TickCount - t0) " emptyMon=" emptyMon " halfPartner=" halfPartner
        AutoSlot_PerfLog(hwnd, "Place_after_occ_scan", occDetail)

        if (emptyMon) {
            tMax := A_TickCount
            tBanner := A_TickCount
            if (AutoSlot_MaximizeOnMonitor(hwnd, emptyMon)) {
                AutoSlot_ClaimMonitor(emptyMon)
                msg := "ℹ️ Auto-slotted → M" emptyOrder " (maximized)"
            }
            bannerMs := A_TickCount - tBanner
            AutoSlot_PerfLog(hwnd, "Place_maximize_done", "ms=" (A_TickCount - tMax) " banner_ms=" bannerMs)
            AutoSlot_Toast(msg)
            AutoSlot_PerfLog(hwnd, "Place_exit", "path=maximize total=" (A_TickCount - t0) " banner_ms=" bannerMs)
            AutoSlot_PerfClearOrigin(hwnd)
            return
        }

        if (halfMon && halfPartner) {
            tSnap := A_TickCount
            tBanner := A_TickCount
            if (AutoSlot_TrySnapNewWithPartner(hwnd, halfMon, halfPartner, halfOrder)) {
                bannerMs := A_TickCount - tBanner
                AutoSlot_RememberHwndMon(hwnd)
                AutoSlot_RememberHwndMon(halfPartner)
                msg := "ℹ️ Auto-slotted → M" halfOrder " (50/50)"
                AutoSlot_Toast(msg)
                AutoSlot_PerfLog(hwnd, "Place_exit", "path=snap total=" (A_TickCount - t0)
                " snapMs=" (A_TickCount - tSnap) " banner_ms=" bannerMs)
                AutoSlot_PerfClearOrigin(hwnd)
                return
            }
            AutoSlot_PerfLog(hwnd, "Place_snap_failed", "ms=" (A_TickCount - tSnap) " banner_ms=" (A_TickCount -
                tBanner))
        }

        ; No empty ordinal and no free half — leave window as the OS opened it
        ; (do not maximize over an existing pair / full monitor).
        AutoSlot_PerfLog(hwnd, "Place_exit", "path=leave total=" (A_TickCount - t0))
        AutoSlot_PerfClearOrigin(hwnd)
    } finally {
        AutoSlot_EndPlaceCritical()
    }
}

; Defer init until WindowManagement finishes the rest of auto-exec and Act's
; window churn settles — avoids unset-map races from early shell/WinEvent floods.
SetTimer(AutoSlot_Init, -500)