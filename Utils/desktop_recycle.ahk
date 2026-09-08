; =============================================================================
; Utils module: desktop_recycle.ahk
; Desktop to Recycle Bin macro
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; Move all Desktop items to Recycle Bin (recoverable)
; Trigger: Ctrl+Alt+Win+8
; Opens a temporary Desktop Explorer at 50% size (fully opaque), centered on the
; active window's monitor. Y / timeout = recycle; N / Escape = cancel. Preview hwnd
; is marked with window prop DesktopToRecycleTempExclude so AutoSlot skips it.
; =============================================================================
global g_DesktopToRecyclePath := ""
global g_DesktopToRecycleCloseHwnd := 0
global g_DesktopToRecycleWeOpenedExplorer := false
global g_DesktopToRecycleTrackTimer := ""
global g_DesktopToRecycleTrackLastMonIdx := 0
global g_DesktopToRecycleReinforceGen := 0
global DESKTOP_TO_RECYCLE_AUTOSLOT_PROP := "DesktopToRecycleTempExclude"
global DESKTOP_TO_RECYCLE_TRACK_INTERVAL := 115
global DESKTOP_TO_RECYCLE_PREVIEW_SCALE := 0.5
global g_DesktopToRecycleKeysArmTick := 0
global g_DesktopToRecycleGraceUntilTick := 0
global g_DesktopToRecycleSawSelectKeyUp := false
global g_DesktopToRecycleSessionId := 0
global g_DesktopToRecycleSessionStartTick := 0
global DESKTOP_TO_RECYCLE_KEYS_GRACE_MS := 1500  ; ignore Y/N until grace ends AND keys have been up
global DESKTOP_TO_RECYCLE_DECISION_MS := 6000
global g_DesktopToRecyclePlaceX := 0
global g_DesktopToRecyclePlaceY := 0
global g_DesktopToRecyclePlaceW := 0
global g_DesktopToRecyclePlaceH := 0
global g_DesktopToRecycleAnchorHwnd := 0

DesktopToRecycle_KeysArmed() {
    global g_DesktopToRecycleKeysArmTick, g_DesktopToRecycleSawSelectKeyUp
    return g_DesktopToRecycleKeysArmTick > 0 && A_TickCount >= g_DesktopToRecycleKeysArmTick &&
        g_DesktopToRecycleSawSelectKeyUp
}

DesktopToRecycle_StopKeysArmTimer() {
    try SetTimer(DesktopToRecycle_TryArmKeys, 0)
    catch {
    }
}

; After grace: wait until Y/N are up (defeats key-repeat), reseed poll, then arm.
DesktopToRecycle_TryArmKeys(*) {
    global g_DesktopToRecycleGraceUntilTick, g_DesktopToRecycleKeysArmTick, g_DesktopToRecycleSawSelectKeyUp
    global g_StandardLoadingBarKeysPollPrev
    if (A_TickCount < g_DesktopToRecycleGraceUntilTick)
        return
    nDown := GetKeyState("N", "P") || GetKeyState("n", "P")
    yDown := GetKeyState("Y", "P") || GetKeyState("y", "P")
    if (nDown || yDown)
        return
    g_DesktopToRecycleSawSelectKeyUp := true
    g_DesktopToRecycleKeysArmTick := A_TickCount
    ; Reseed poll so a held-then-released key cannot look like a fresh edge.
    try {
        if (IsObject(g_StandardLoadingBarKeysPollPrev)) {
            g_StandardLoadingBarKeysPollPrev["N"] := false
            g_StandardLoadingBarKeysPollPrev["n"] := false
            g_StandardLoadingBarKeysPollPrev["Y"] := false
            g_StandardLoadingBarKeysPollPrev["y"] := false
        }
    } catch {
    }
    DesktopToRecycle_StopKeysArmTimer()
}

DesktopToRecycle_StartKeysArm() {
    global g_DesktopToRecycleKeysArmTick, g_DesktopToRecycleGraceUntilTick, g_DesktopToRecycleSawSelectKeyUp
    global DESKTOP_TO_RECYCLE_KEYS_GRACE_MS
    DesktopToRecycle_StopKeysArmTimer()
    g_DesktopToRecycleSawSelectKeyUp := false
    g_DesktopToRecycleKeysArmTick := 0  ; not armed until TryArmKeys succeeds
    g_DesktopToRecycleGraceUntilTick := A_TickCount + DESKTOP_TO_RECYCLE_KEYS_GRACE_MS
    SetTimer(DesktopToRecycle_TryArmKeys, 50)
}

DesktopToRecycle_OnConfirm(*) {
    if (!DesktopToRecycle_KeysArmed())
        return
    DesktopToRecycle_StopKeysArmTimer()
    DesktopToRecycle_EndDecisionSession()
    DesktopToRecycle_StopTrack()
    DesktopToRecycle_ClosePreviewExplorer()
    PlayCleaningDesktopSound()
    DesktopToRecycle_Run()
}

DesktopToRecycle_OnCancel(*) {
    DesktopToRecycle_StopKeysArmTimer()
    DesktopToRecycle_EndDecisionSession()
    DesktopToRecycle_StopTrack()
    try StandardLoadingBar_CloseKeysOverlay()
    catch {
    }
    DesktopToRecycle_ClosePreviewExplorer()
    DesktopToRecycle_CloseMarkedTempExplorers()
    ShowCenteredOverlay_Utils("⚠ Desktop cleanup cancelled", 1500, BANNER_ACCENT_INTERMEDIATE)
}

; N and Escape are identical: always cancel (grace only protects Y / recycle).
DesktopToRecycle_OnCancelFromN(*) {
    DesktopToRecycle_OnCancel()
}

DesktopToRecycle_OnTimeout(*) {
    global g_DesktopToRecycleCloseHwnd, g_DesktopToRecycleSessionId, g_DesktopToRecycleSessionStartTick
    global DESKTOP_TO_RECYCLE_DECISION_MS
    ; Defend against a stale ShowWithKeys timer from a prior confirm (fixed in CloseKeysOverlay,
    ; but keep this guard if an old BoundFunc still fires).
    age := g_DesktopToRecycleSessionStartTick > 0 ? (A_TickCount - g_DesktopToRecycleSessionStartTick) : -1
    if (!g_DesktopToRecycleSessionId || !g_DesktopToRecycleCloseHwnd || age >= 0 && age <
        DESKTOP_TO_RECYCLE_DECISION_MS -
        400)
        return
    DesktopToRecycle_StopKeysArmTimer()
    DesktopToRecycle_EndDecisionSession()
    DesktopToRecycle_StopTrack()
    DesktopToRecycle_ClosePreviewExplorer()
    DesktopToRecycle_Run()
}

DesktopToRecycle_NormalizePath(p) {
    p := RTrim(p, "\")
    try return StrLower(p)
    return p
}

; Mark hwnd so AutoSlot (WindowManagement process) skips Place/occupancy via GetProp.
DesktopToRecycle_MarkAutoSlotExclude(hwnd) {
    global DESKTOP_TO_RECYCLE_AUTOSLOT_PROP
    if (!hwnd)
        return
    try DllCall("SetPropW", "ptr", hwnd, "wstr", DESKTOP_TO_RECYCLE_AUTOSLOT_PROP, "ptr", 1)
    catch {
    }
}

DesktopToRecycle_ClearAutoSlotExclude(hwnd) {
    global DESKTOP_TO_RECYCLE_AUTOSLOT_PROP
    if (!hwnd)
        return
    try DllCall("RemovePropW", "ptr", hwnd, "wstr", DESKTOP_TO_RECYCLE_AUTOSLOT_PROP)
    catch {
    }
}

; Cross-process suppress file (Utils + WindowManagement/AutoSlot share A_ScriptDir).
DesktopToRecycle_AutoSlotSuppressPath() {
    return A_ScriptDir "\assets\data\desktop_recycle_autoslot_suppress.ini"
}

; Call BEFORE Run explorer so AutoSlot debounce cannot Place the new window.
DesktopToRecycle_BeginAutoSlotSuppress(durationMs := 12000) {
    path := DesktopToRecycle_AutoSlotSuppressPath()
    try DirCreate(A_ScriptDir "\assets\data")
    catch {
    }
    suppressUntilTick := DllCall("GetTickCount", "UInt") + durationMs
    try {
        IniWrite(suppressUntilTick, path, "Suppress", "Until")
        IniWrite(1, path, "Suppress", "Active")
    } catch {
    }
}

DesktopToRecycle_EndAutoSlotSuppress() {
    path := DesktopToRecycle_AutoSlotSuppressPath()
    try {
        IniWrite(0, path, "Suppress", "Until")
        IniWrite(0, path, "Suppress", "Active")
    } catch {
    }
}

; Used by AutoSlot_IsExcludedExeOrTitle (same process when WM includes Utils, and via file).
DesktopToRecycle_AutoSlotSuppressActive() {
    path := DesktopToRecycle_AutoSlotSuppressPath()
    try {
        active := Integer(IniRead(path, "Suppress", "Active", 0))
        suppressUntilTick := Integer(IniRead(path, "Suppress", "Until", 0))
    } catch {
        return false
    }
    if (!active || suppressUntilTick < 1)
        return false
    return DllCall("GetTickCount", "UInt") < suppressUntilTick
}

DesktopToRecycle_IsDesktopExplorerTitle(title) {
    if (title = "")
        return false
    return InStr(title, "Desktop", false) || InStr(title, "Área de Trabalho", false)
}

; Find Explorer hwnd showing targetPath; else title Desktop / Área de Trabalho.
DesktopToRecycle_FindDesktopExplorer(targetPath) {
    if (targetPath && targetPath != "") {
        normTarget := DesktopToRecycle_NormalizePath(targetPath)
        try {
            shell := ComObject("Shell.Application")
            for window in shell.Windows {
                try {
                    if (!window || !window.hwnd)
                        continue
                    path := window.Document.Folder.Self.Path
                    if (DesktopToRecycle_NormalizePath(path) = normTarget)
                        return Integer(window.hwnd)
                } catch
                    continue
            }
        } catch {
        }
    }
    prevMode := A_TitleMatchMode
    try {
        SetTitleMatchMode 2
        hwnd := WinExist("Área de Trabalho ahk_class CabinetWClass")
        if (hwnd)
            return hwnd
        return WinExist("Desktop ahk_class CabinetWClass")
    } finally {
        SetTitleMatchMode prevMode
    }
}

; Snapshot of Desktop Explorer hwnds before we launch (to prefer a newly created one).
DesktopToRecycle_CollectDesktopExplorerHwnds(targetPath) {
    found := Map()
    if (targetPath && targetPath != "") {
        normTarget := DesktopToRecycle_NormalizePath(targetPath)
        try {
            shell := ComObject("Shell.Application")
            for window in shell.Windows {
                try {
                    if (!window || !window.hwnd)
                        continue
                    path := window.Document.Folder.Self.Path
                    if (DesktopToRecycle_NormalizePath(path) = normTarget)
                        found[Integer(window.hwnd)] := true
                } catch
                    continue
            }
        } catch {
        }
    }
    return found
}

; All Shell.Application explorer hwnds (for early claim before COM path is ready).
DesktopToRecycle_CollectAllShellHwnds() {
    found := Map()
    try {
        shell := ComObject("Shell.Application")
        for window in shell.Windows {
            try {
                if (window && window.hwnd)
                    found[Integer(window.hwnd)] := true
            } catch
                continue
        }
    } catch {
    }
    return found
}

; True if shell reports this Explorer hwnd at targetPath (COM may lag after create).
DesktopToRecycle_ExplorerPathMatches(hwnd, targetPath) {
    if (!hwnd || !targetPath)
        return false
    normTarget := DesktopToRecycle_NormalizePath(targetPath)
    try {
        shell := ComObject("Shell.Application")
        for window in shell.Windows {
            try {
                if (!window || !window.hwnd || Integer(window.hwnd) != Integer(hwnd))
                    continue
                path := window.Document.Folder.Self.Path
                return DesktopToRecycle_NormalizePath(path) = normTarget
            } catch
                continue
        }
    } catch {
    }
    return false
}

; New Shell window not in beforeShell — park immediately. Path may still be empty.
DesktopToRecycle_ClaimNewShellEarly(beforeShell, targetPath, workLeft, workTop, workRight, workBottom) {
    try {
        shell := ComObject("Shell.Application")
        for window in shell.Windows {
            try {
                if (!window || !window.hwnd)
                    continue
                h := Integer(window.hwnd)
                if (beforeShell.Has(h))
                    continue
                path := ""
                try path := window.Document.Folder.Self.Path
                catch {
                }
                title := ""
                try title := WinGetTitle("ahk_id " h)
                catch {
                }
                pathOk := (path != "" && DesktopToRecycle_NormalizePath(path) = DesktopToRecycle_NormalizePath(
                    targetPath))
                titleOk := DesktopToRecycle_IsDesktopExplorerTitle(title)
                if (!(pathOk || titleOk || path = ""))
                    continue
                DesktopToRecycle_MarkAutoSlotExclude(h)
                DesktopToRecycle_ParkPreviewOnWorkArea(h, workLeft, workTop, workRight, workBottom)
                return h
            } catch
                continue
        }
    } catch {
    }
    return 0
}

; Center hwnd at 50% of the given work area (fully opaque).
; Uses split SetWindowPos (move then size) — combined WinMove/SetWindowPos balloons ~1.5x
; on mixed-DPI.
DesktopToRecycle_ForceMoveHwnd(hwnd, x, y, w, h) {
    if (!hwnd || w < 1 || h < 1)
        return false
    ; SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE = 0x0015; SWP_NOMOVE|... = 0x0016
    okMove := DllCall("SetWindowPos", "ptr", hwnd, "ptr", 0, "int", x, "int", y, "int", 0, "int", 0, "uint",
        0x0015)
    okSize := DllCall("SetWindowPos", "ptr", hwnd, "ptr", 0, "int", 0, "int", 0, "int", w, "int", h, "uint",
        0x0016)
    if (okMove && okSize)
        return true
    try {
        WinMove(x, y, w, h, "ahk_id " hwnd)
        return true
    } catch {
        return false
    }
}

DesktopToRecycle_PlacePreviewOnWorkArea(hwnd, workLeft, workTop, workRight, workBottom) {
    global DESKTOP_TO_RECYCLE_PREVIEW_SCALE
    global g_DesktopToRecyclePlaceX, g_DesktopToRecyclePlaceY, g_DesktopToRecyclePlaceW, g_DesktopToRecyclePlaceH
    ; IsWindow — WinExist misses hidden windows (hide-before-place regression).
    if (!hwnd || !DllCall("IsWindow", "ptr", hwnd))
        return false
    monW := workRight - workLeft
    monH := workBottom - workTop
    if (monW < 1 || monH < 1)
        return false
    w := Max(200, Round(monW * DESKTOP_TO_RECYCLE_PREVIEW_SCALE))
    h := Max(150, Round(monH * DESKTOP_TO_RECYCLE_PREVIEW_SCALE))
    x := Round(workLeft + (monW - w) / 2)
    y := Round(workTop + (monH - h) / 2)
    try {
        if (WinGetMinMax("ahk_id " hwnd) = -1 || WinGetMinMax("ahk_id " hwnd) = 1)
            WinRestore("ahk_id " hwnd)
    } catch {
    }
    if (!DesktopToRecycle_ForceMoveHwnd(hwnd, x, y, w, h))
        return false
    ; If Explorer/DPI still inflated size, correct immediately (do not wait for reinforce blink).
    rx := ry := rw := rh := -1
    try WinGetPos(&rx, &ry, &rw, &rh, "ahk_id " hwnd)
    catch {
    }
    if (rw > 0 && rh > 0 && (Abs(rw - w) > 40 || Abs(rh - h) > 40 || Abs(rx - x) > 40 || Abs(ry - y) > 40))
        DesktopToRecycle_ForceMoveHwnd(hwnd, x, y, w, h)
    g_DesktopToRecyclePlaceX := x
    g_DesktopToRecyclePlaceY := y
    g_DesktopToRecyclePlaceW := w
    g_DesktopToRecyclePlaceH := h
    return true
}

DesktopToRecycle_PlacePreviewOnMonitor(hwnd, monIdx) {
    if (!hwnd || monIdx < 1)
        return false
    try MonitorGetWorkArea(monIdx, &l, &t, &r, &b)
    catch {
        return false
    }
    return DesktopToRecycle_PlacePreviewOnWorkArea(hwnd, l, t, r, b)
}

; Re-apply preview geometry if AutoSlot or Explorer raced and maximized/moved us.
DesktopToRecycle_ReinforcePlace(*) {
    global g_DesktopToRecycleCloseHwnd, g_DesktopToRecyclePlaceX, g_DesktopToRecyclePlaceY
    global g_DesktopToRecyclePlaceW, g_DesktopToRecyclePlaceH
    hwnd := g_DesktopToRecycleCloseHwnd
    if (!hwnd || !DllCall("IsWindow", "ptr", hwnd))
        return
    DesktopToRecycle_MarkAutoSlotExclude(hwnd)
    if (g_DesktopToRecyclePlaceW > 0 && g_DesktopToRecyclePlaceH > 0) {
        DesktopToRecycle_ForceMoveHwnd(hwnd, g_DesktopToRecyclePlaceX, g_DesktopToRecyclePlaceY,
            g_DesktopToRecyclePlaceW, g_DesktopToRecyclePlaceH)
    } else {
        monIdx := g_DesktopToRecycleTrackLastMonIdx
        if (monIdx < 1)
            monIdx := GetMonitorIndexForForeground_StandardBar()
        DesktopToRecycle_PlacePreviewOnMonitor(hwnd, monIdx)
    }
    ; Z-order only (no Activate) so the Y/N banner keeps focus.
    DesktopToRecycle_BringPreviewToFront(hwnd, false)
}

; Force Explorer above other apps. activate:=true on first show; false during confirm reinforce.
DesktopToRecycle_BringPreviewToFront(hwnd, activate := true) {
    if (!hwnd || !DllCall("IsWindow", "ptr", hwnd))
        return false
    try {
        mm := WinGetMinMax("ahk_id " hwnd)
        if (mm = 1 || mm = -1)
            WinRestore("ahk_id " hwnd)
    } catch {
    }
    try WinShow("ahk_id " hwnd)
    catch {
    }
    try WinSetAlwaysOnTop(true, "ahk_id " hwnd)
    catch {
    }
    ; HWND_TOPMOST = -1; SWP_NOSIZE|SWP_NOMOVE|SWP_SHOWWINDOW = 0x0043
    DllCall("SetWindowPos", "ptr", hwnd, "ptr", -1, "int", 0, "int", 0, "int", 0, "int", 0, "uint", 0x0043)
    if (activate) {
        try WinActivate("ahk_id " hwnd)
        catch {
        }
        ; AttachThreadInput fallback when WinActivate is blocked by foreground lock.
        try {
            fg := DllCall("GetForegroundWindow", "ptr")
            if (fg != hwnd) {
                curTid := DllCall("GetCurrentThreadId", "UInt")
                fgTid := DllCall("GetWindowThreadProcessId", "ptr", fg, "ptr", 0, "UInt")
                tgtTid := DllCall("GetWindowThreadProcessId", "ptr", hwnd, "ptr", 0, "UInt")
                if (fgTid && tgtTid && fgTid != curTid)
                    DllCall("AttachThreadInput", "UInt", curTid, "UInt", fgTid, "Int", 1)
                if (tgtTid && tgtTid != curTid)
                    DllCall("AttachThreadInput", "UInt", curTid, "UInt", tgtTid, "Int", 1)
                DllCall("SetForegroundWindow", "ptr", hwnd)
                DllCall("BringWindowToTop", "ptr", hwnd)
                if (fgTid && tgtTid && fgTid != curTid)
                    DllCall("AttachThreadInput", "UInt", curTid, "UInt", fgTid, "Int", 0)
                if (tgtTid && tgtTid != curTid)
                    DllCall("AttachThreadInput", "UInt", curTid, "UInt", tgtTid, "Int", 0)
            }
        } catch {
        }
    }
    return DllCall("IsWindowVisible", "ptr", hwnd) ? true : false
}

DesktopToRecycle_ReinforcePlaceIfGen(gen, *) {
    global g_DesktopToRecycleReinforceGen
    if (gen != g_DesktopToRecycleReinforceGen)
        return
    DesktopToRecycle_ReinforcePlace()
}

DesktopToRecycle_StopTrack() {
    global g_DesktopToRecycleTrackTimer, g_DesktopToRecycleTrackLastMonIdx, g_DesktopToRecycleReinforceGen
    try SetTimer(DesktopToRecycle_TrackTick, 0)
    catch {
    }
    ; Invalidate pending reinforce one-shots.
    g_DesktopToRecycleReinforceGen += 1
    g_DesktopToRecycleTrackTimer := ""
    g_DesktopToRecycleTrackLastMonIdx := 0
}

; Keep preview pinned (no focus-follow).
DesktopToRecycle_TrackTick(*) {
    global g_DesktopToRecycleCloseHwnd, g_DesktopToRecycleTrackLastMonIdx
    hwnd := g_DesktopToRecycleCloseHwnd
    if (!hwnd || !WinExist("ahk_id " hwnd)) {
        DesktopToRecycle_StopTrack()
        return
    }
    DesktopToRecycle_MarkAutoSlotExclude(hwnd)
    try WinSetAlwaysOnTop(true, "ahk_id " hwnd)
    catch {
    }
    try {
        if (WinGetMinMax("ahk_id " hwnd) = 1) {
            monIdx := g_DesktopToRecycleTrackLastMonIdx > 0 ? g_DesktopToRecycleTrackLastMonIdx : 1
            DesktopToRecycle_PlacePreviewOnMonitor(hwnd, monIdx)
        }
    } catch {
    }
}

DesktopToRecycle_StartTrack(hwnd, initialMonIdx) {
    global g_DesktopToRecycleTrackTimer, g_DesktopToRecycleTrackLastMonIdx, DESKTOP_TO_RECYCLE_TRACK_INTERVAL
    global g_DesktopToRecycleReinforceGen
    DesktopToRecycle_StopTrack()
    g_DesktopToRecycleTrackLastMonIdx := initialMonIdx
    try WinSetAlwaysOnTop(true, "ahk_id " hwnd)
    catch {
    }
    SetTimer(DesktopToRecycle_TrackTick, DESKTOP_TO_RECYCLE_TRACK_INTERVAL)
    g_DesktopToRecycleTrackTimer := DesktopToRecycle_TrackTick
    g_DesktopToRecycleReinforceGen += 1
    gen := g_DesktopToRecycleReinforceGen
    SetTimer(DesktopToRecycle_ReinforcePlaceIfGen.Bind(gen), -350)
    SetTimer(DesktopToRecycle_ReinforcePlaceIfGen.Bind(gen), -700)
    SetTimer(DesktopToRecycle_ReinforcePlaceIfGen.Bind(gen), -1200)
}

DesktopToRecycle_BeginDecisionSession() {
    global g_DesktopToRecycleSessionId, g_DesktopToRecycleSessionStartTick, DESKTOP_TO_RECYCLE_DECISION_MS
    g_DesktopToRecycleSessionStartTick := A_TickCount
    g_DesktopToRecycleSessionId := A_TickCount
    sid := g_DesktopToRecycleSessionId
    SetTimer(DesktopToRecycle_SessionExpired.Bind(sid), -DESKTOP_TO_RECYCLE_DECISION_MS)
}

DesktopToRecycle_EndDecisionSession() {
    global g_DesktopToRecycleSessionId, g_DesktopToRecycleSessionStartTick
    ; Invalidate any pending SessionExpired bind.
    g_DesktopToRecycleSessionId := 0
    g_DesktopToRecycleSessionStartTick := 0
}

DesktopToRecycle_SessionExpired(sid, *) {
    global g_DesktopToRecycleSessionId, g_DesktopToRecycleCloseHwnd
    if (sid != g_DesktopToRecycleSessionId)
        return
    if (!g_DesktopToRecycleCloseHwnd) {
        DesktopToRecycle_EndDecisionSession()
        return
    }
    DesktopToRecycle_OnTimeout()
}

; Force-close one Explorer frame (WinClose alone often leaves Desktop Explorer open).
DesktopToRecycle_ForceCloseExplorerHwnd(hwnd) {
    if (!hwnd || !DllCall("IsWindow", "ptr", hwnd))
        return
    try {
        shell := ComObject("Shell.Application")
        for window in shell.Windows {
            try {
                if (!window || !window.hwnd)
                    continue
                if (Integer(window.hwnd) = Integer(hwnd)) {
                    window.Quit()
                    break
                }
            } catch
                continue
        }
    } catch {
    }
    if (DllCall("IsWindow", "ptr", hwnd)) {
        try PostMessage(0x0010, 0, 0, , "ahk_id " hwnd)  ; WM_CLOSE
        catch {
        }
    }
    Sleep 40
    if (DllCall("IsWindow", "ptr", hwnd)) {
        try WinClose("ahk_id " hwnd)
        catch {
        }
    }
    Sleep 40
    if (DllCall("IsWindow", "ptr", hwnd)) {
        try WinKill("ahk_id " hwnd)
        catch {
        }
    }
}

; Close only the temporary preview Explorer hwnd (not every Desktop Explorer).
DesktopToRecycle_ClosePreviewExplorer() {
    global g_DesktopToRecycleCloseHwnd, g_DesktopToRecycleWeOpenedExplorer
    DesktopToRecycle_StopKeysArmTimer()
    DesktopToRecycle_EndDecisionSession()
    DesktopToRecycle_StopTrack()
    hwnd := g_DesktopToRecycleCloseHwnd
    g_DesktopToRecycleCloseHwnd := 0
    g_DesktopToRecycleWeOpenedExplorer := false
    if (hwnd) {
        try WinSetAlwaysOnTop(false, "ahk_id " hwnd)
        catch {
        }
    }
    DesktopToRecycle_EndAutoSlotSuppress()
    if (!hwnd)
        return
    DesktopToRecycle_ClearAutoSlotExclude(hwnd)
    DesktopToRecycle_ForceCloseExplorerHwnd(hwnd)
}

; Close Explorers still marked with our AutoSlot exclude prop (cancel / leak cleanup).
DesktopToRecycle_CloseMarkedTempExplorers() {
    global DESKTOP_TO_RECYCLE_AUTOSLOT_PROP
    try {
        shell := ComObject("Shell.Application")
        for window in shell.Windows {
            try {
                if (!window || !window.hwnd)
                    continue
                h := Integer(window.hwnd)
                if (!DllCall("GetPropW", "ptr", h, "wstr", DESKTOP_TO_RECYCLE_AUTOSLOT_PROP))
                    continue
                DesktopToRecycle_ClearAutoSlotExclude(h)
                DesktopToRecycle_ForceCloseExplorerHwnd(h)
            } catch
                continue
        }
    } catch {
    }
}

; Park at target work-area center as soon as hwnd is known.
DesktopToRecycle_ParkPreviewOnWorkArea(hwnd, workLeft, workTop, workRight, workBottom) {
    global DESKTOP_TO_RECYCLE_PREVIEW_SCALE
    if (!hwnd || !DllCall("IsWindow", "ptr", hwnd))
        return false
    monW := workRight - workLeft
    monH := workBottom - workTop
    if (monW < 1 || monH < 1)
        return false
    w := Max(200, Round(monW * DESKTOP_TO_RECYCLE_PREVIEW_SCALE))
    h := Max(150, Round(monH * DESKTOP_TO_RECYCLE_PREVIEW_SCALE))
    x := Round(workLeft + (monW - w) / 2)
    y := Round(workTop + (monH - h) / 2)
    try {
        if (WinGetMinMax("ahk_id " hwnd) = 1 || WinGetMinMax("ahk_id " hwnd) = -1)
            WinRestore("ahk_id " hwnd)
    } catch {
    }
    return DesktopToRecycle_ForceMoveHwnd(hwnd, x, y, w, h)
}

; Claim a newly opened Desktop Explorer within timeoutMs. Returns hwnd or 0; sets claimVia.
DesktopToRecycle_ClaimNewPreviewHwnd(before, beforeShell, targetPath, workLeft, workTop, workRight, workBottom,
    timeoutMs, &claimVia) {
    claimVia := ""
    hwnd := 0
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        early := DesktopToRecycle_ClaimNewShellEarly(beforeShell, targetPath, workLeft, workTop, workRight, workBottom)
        if (early && !before.Has(early)) {
            hwnd := early
            claimVia := "early_shell"
            break
        }
        try {
            shell := ComObject("Shell.Application")
            for window in shell.Windows {
                try {
                    if (!window || !window.hwnd)
                        continue
                    h := Integer(window.hwnd)
                    path := window.Document.Folder.Self.Path
                    if (DesktopToRecycle_NormalizePath(path) != DesktopToRecycle_NormalizePath(targetPath))
                        continue
                    if (!before.Has(h)) {
                        hwnd := h
                        claimVia := "shell_path"
                        DesktopToRecycle_MarkAutoSlotExclude(hwnd)
                        DesktopToRecycle_ParkPreviewOnWorkArea(hwnd, workLeft, workTop, workRight, workBottom)
                        break
                    }
                } catch
                    continue
            }
        } catch {
        }
        if (hwnd)
            break
        cand := DesktopToRecycle_FindDesktopExplorer(targetPath)
        if (cand && !before.Has(cand)) {
            hwnd := cand
            claimVia := "FindDesktop"
            DesktopToRecycle_MarkAutoSlotExclude(hwnd)
            DesktopToRecycle_ParkPreviewOnWorkArea(hwnd, workLeft, workTop, workRight, workBottom)
            break
        }
        Sleep 10
    }
    if (!hwnd) {
        hwnd := DesktopToRecycle_FindDesktopExplorer(targetPath)
        claimVia := "fallback"
        if (hwnd && !before.Has(hwnd)) {
            DesktopToRecycle_MarkAutoSlotExclude(hwnd)
            DesktopToRecycle_ParkPreviewOnWorkArea(hwnd, workLeft, workTop, workRight, workBottom)
        } else if (hwnd && before.Has(hwnd))
            hwnd := 0
    }
    if (hwnd && claimVia = "early_shell" && !before.Has(hwnd)) {
        confirmed := false
        vDeadline := A_TickCount + 1200
        while (A_TickCount < vDeadline) {
            t := ""
            try t := WinGetTitle("ahk_id " hwnd)
            catch {
            }
            if (DesktopToRecycle_ExplorerPathMatches(hwnd, targetPath) || DesktopToRecycle_IsDesktopExplorerTitle(t)) {
                confirmed := true
                break
            }
            Sleep 20
        }
        if (!confirmed) {
            DesktopToRecycle_ForceCloseExplorerHwnd(hwnd)
            hwnd := 0
            claimVia := "early_rejected"
        }
    }
    if (!hwnd || !DllCall("IsWindow", "ptr", hwnd) || before.Has(hwnd))
        return 0
    return hwnd
}

DesktopToRecycle_RunExplorerForPath(targetPath) {
    try Run('explorer.exe /n,"' targetPath '"')
    catch {
        try Run('explorer.exe "' targetPath '"')
        catch {
            return false
        }
    }
    return true
}

; One open attempt: snapshot → Run → claim → place → bring to front.
DesktopToRecycle_OpenPreviewExplorerOnce(targetPath, workLeft, workTop, workRight, workBottom) {
    before := DesktopToRecycle_CollectDesktopExplorerHwnds(targetPath)
    beforeShell := DesktopToRecycle_CollectAllShellHwnds()
    if (!DesktopToRecycle_RunExplorerForPath(targetPath))
        return 0
    claimVia := ""
    hwnd := DesktopToRecycle_ClaimNewPreviewHwnd(before, beforeShell, targetPath, workLeft, workTop, workRight,
        workBottom, 4000, &claimVia)
    if (!hwnd)
        return 0

    DesktopToRecycle_MarkAutoSlotExclude(hwnd)
    Sleep 40
    placed := DesktopToRecycle_PlacePreviewOnWorkArea(hwnd, workLeft, workTop, workRight, workBottom)
    if (!placed) {
        DesktopToRecycle_ClearAutoSlotExclude(hwnd)
        DesktopToRecycle_ForceCloseExplorerHwnd(hwnd)
        return 0
    }
    visible := DesktopToRecycle_BringPreviewToFront(hwnd, true)
    if (!visible) {
        DesktopToRecycle_PlacePreviewOnWorkArea(hwnd, workLeft, workTop, workRight, workBottom)
        visible := DesktopToRecycle_BringPreviewToFront(hwnd, true)
    }
    if (!visible || !DllCall("IsWindow", "ptr", hwnd)) {
        DesktopToRecycle_ClearAutoSlotExclude(hwnd)
        DesktopToRecycle_ForceCloseExplorerHwnd(hwnd)
        return 0
    }
    return hwnd
}

; Open a new Desktop Explorer, exclude from AutoSlot, place at 50% size (opaque). Returns hwnd or 0.
DesktopToRecycle_OpenPreviewExplorer(targetPath, workLeft, workTop, workRight, workBottom) {
    global g_DesktopToRecycleWeOpenedExplorer, g_DesktopToRecycleCloseHwnd
    g_DesktopToRecycleWeOpenedExplorer := false
    g_DesktopToRecycleCloseHwnd := 0
    if (!targetPath || !DirExist(targetPath))
        return 0

    initialDesktop := DesktopToRecycle_CollectDesktopExplorerHwnds(targetPath)
    ; Suppress AutoSlot BEFORE Run — SHOW/Schedule races SetProp by hundreds of ms.
    DesktopToRecycle_BeginAutoSlotSuppress(16000)
    hwnd := DesktopToRecycle_OpenPreviewExplorerOnce(targetPath, workLeft, workTop, workRight, workBottom)
    if (!hwnd) {
        ; Close stragglers from the failed first launch before retrying.
        leftovers := DesktopToRecycle_CollectDesktopExplorerHwnds(targetPath)
        for h, _ in leftovers {
            h := Integer(h)
            if (initialDesktop.Has(h))
                continue
            DesktopToRecycle_ClearAutoSlotExclude(h)
            DesktopToRecycle_ForceCloseExplorerHwnd(h)
        }
        DesktopToRecycle_CloseMarkedTempExplorers()
        Sleep 120
        hwnd := DesktopToRecycle_OpenPreviewExplorerOnce(targetPath, workLeft, workTop, workRight, workBottom)
    }
    if (!hwnd) {
        DesktopToRecycle_EndAutoSlotSuppress()
        return 0
    }
    g_DesktopToRecycleWeOpenedExplorer := true
    g_DesktopToRecycleCloseHwnd := hwnd
    ; Keep suppress active for the whole preview; End on ClosePreviewExplorer.
    return hwnd
}

DesktopToRecycle_Run() {
    global g_DesktopToRecyclePath
    path := g_DesktopToRecyclePath
    if (!path || path = "" || !DirExist(path))
        path := A_Desktop
    ui := "[Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs"
    rec := "[Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin"
    ps := "Add-Type -AssemblyName Microsoft.VisualBasic;$d='" . path .
        "';if(-not(Test-Path -LiteralPath $d)){exit 1};$files=@(Get-ChildItem -LiteralPath $d -Force|Where-Object{-not $_.PSIsContainer});$dirs=@(Get-ChildItem -LiteralPath $d -Force|Where-Object{$_.PSIsContainer});foreach($f in $files){try{[Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($f.FullName," .
        ui . "," . rec .
        ")}catch{}};foreach($dir in $dirs){try{[Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory($dir.FullName," .
        ui . "," . rec . ")}catch{}};exit 0"
    try {
        exitCode := RunWait('powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "' . ps . '"', "", "Hide")
        if (exitCode = 0)
            ShowCenteredOverlay_Utils("✅ Desktop items moved to Recycle Bin", 2000, BANNER_ACCENT_SUCCESS)
        else
            ShowCenteredOverlay_Utils("❌ Desktop path not found or error: " path, 3500, BANNER_ACCENT_ERROR)
    } catch as err {
        ShowCenteredOverlay_Utils("❌ Error moving to Recycle Bin", 2500, BANNER_ACCENT_ERROR)
    }
}

; Close prop-marked temps, then leftover Desktop-path Explorers from prior runs.
DesktopToRecycle_CloseAllTempPreviewExplorers() {
    global DESKTOP_TO_RECYCLE_AUTOSLOT_PROP, g_DesktopToRecycleCloseHwnd, g_DesktopToRecyclePath
    DesktopToRecycle_CloseMarkedTempExplorers()
    path := g_DesktopToRecyclePath
    if (!path || path = "") {
        try path := GetDesktopToRecyclePath()
        catch {
            path := ""
        }
    }
    if (path && DirExist(path)) {
        leftovers := DesktopToRecycle_CollectDesktopExplorerHwnds(path)
        for h, _ in leftovers {
            h := Integer(h)
            DesktopToRecycle_ClearAutoSlotExclude(h)
            DesktopToRecycle_ForceCloseExplorerHwnd(h)
        }
    }
    g_DesktopToRecycleCloseHwnd := 0
}

; Entry point for Desktop to Recycle macro (^!#8)
DesktopToRecycle_Trigger() {
    global g_DesktopToRecycleCloseHwnd, g_DesktopToRecyclePath, g_DesktopToRecycleWeOpenedExplorer
    global g_DesktopToRecycleAnchorHwnd

    ; D2C / standard_information_display: capture origin before close/open steals focus.
    originHwnd := 0
    try originHwnd := WinGetID("A")
    catch {
        originHwnd := 0
    }
    g_DesktopToRecycleAnchorHwnd := originHwnd
    workLeft := workTop := 0
    workRight := A_ScreenWidth
    workBottom := A_ScreenHeight
    workArea := ""
    if (originHwnd)
        workArea := GetWorkAreaForWindow_StandardBar(originHwnd)
    if (IsObject(workArea)) {
        workLeft := workArea.left
        workTop := workArea.top
        workRight := workArea.right
        workBottom := workArea.bottom
    } else
        GetActiveMonitorWorkArea_StandardBar(&workLeft, &workTop, &workRight, &workBottom)
    initialMonIdx := GetMonitorIndexForForeground_StandardBar()

    DesktopToRecycle_StopTrack()
    DesktopToRecycle_ClosePreviewExplorer()
    g_DesktopToRecyclePath := GetDesktopToRecyclePath()
    path := g_DesktopToRecyclePath
    if (!path || path = "" || !DirExist(path))
        path := A_Desktop
    g_DesktopToRecyclePath := path
    DesktopToRecycle_CloseAllTempPreviewExplorers()
    DesktopToRecycle_EndAutoSlotSuppress()
    g_DesktopToRecycleWeOpenedExplorer := false

    StandardLoadingBar_CloseKeysOverlay()
    StandardLoadingBar_Hide(0)
    Sleep 50

    ; Loading Indication until Explorer is placed and foregrounded (standard_information_display.md).
    StandardLoadingBar_Show("⏳ Opening Desktop preview...", BANNER_ACCENT_INTERMEDIATE, {
        centerOnHwnd: originHwnd,
        fontSize: 17
    })
    hwnd := 0
    try {
        hwnd := DesktopToRecycle_OpenPreviewExplorer(path, workLeft, workTop, workRight, workBottom)
    } finally {
        StandardLoadingBar_Hide(0)
    }
    g_DesktopToRecycleCloseHwnd := hwnd ? hwnd : 0
    if (!hwnd) {
        ShowCenteredOverlay_Utils("❌ Desktop preview failed to open", 2500, BANNER_ACCENT_ERROR)
        return
    }
    DesktopToRecycle_StartTrack(hwnd, initialMonIdx)

    try KeyWait("N")
    try KeyWait("Y")
    catch {
    }

    DesktopToRecycle_StartKeysArm()
    DesktopToRecycle_BeginDecisionSession()

    global DESKTOP_TO_RECYCLE_DECISION_MS
    state := "🗑️ Move all items from:`n" . g_DesktopToRecyclePath . "`nto Recycle Bin? (6s)"
    keyCallbacks := Map(
        "Y", DesktopToRecycle_OnConfirm,
        "N", DesktopToRecycle_OnCancelFromN,
        "Escape", DesktopToRecycle_OnCancelFromN)
    ; centerOnHwnd = same OriginHwnd used for preview (standard_information_display.md).
    ; skipEscapeDismiss false so Escape matches N (closes modal + Explorer + ends session).
    StandardLoadingBar_ShowWithKeys(
        state,
        keyCallbacks,
        DESKTOP_TO_RECYCLE_DECISION_MS,
        g_DesktopToRecycleAnchorHwnd,
        DesktopToRecycle_OnTimeout,
        BANNER_ACCENT_INTERMEDIATE,
        0,
        17,
        "",
        false,
        "[Y] Yes  [N]/Esc] Cancel",
        false,
        true,
        true,
        "",
        false)
}
