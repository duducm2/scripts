; =============================================================================
; Utils module: desktop_recycle.ahk
; Desktop to Recycle Bin macro
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; Move all Desktop items to Recycle Bin (recoverable)
; Trigger: Ctrl+Alt+Win+8
; Opens a temporary Desktop Explorer at 50% size / 50% opacity, centered on the
; active window's monitor; tracks focus for 4s and recenters if the user switches
; monitors. Y / timeout = recycle; N / Escape = cancel. Preview hwnd is marked
; with window prop DesktopToRecycleTempExclude so AutoSlot skips it (cross-process).
; =============================================================================
global g_DesktopToRecyclePath := ""
global g_DesktopToRecycleCloseHwnd := 0
global g_DesktopToRecycleWeOpenedExplorer := false
global g_DesktopToRecycleTrackTimer := ""
global g_DesktopToRecycleTrackLastMonIdx := 0
global g_DesktopToRecycleReinforceGen := 0
global DESKTOP_TO_RECYCLE_AUTOSLOT_PROP := "DesktopToRecycleTempExclude"
global DESKTOP_TO_RECYCLE_TRACK_INTERVAL := 115
global DESKTOP_TO_RECYCLE_PREVIEW_OPACITY := 128  ; 50% of 255
global DESKTOP_TO_RECYCLE_PREVIEW_SCALE := 0.5
global g_DesktopToRecycleKeysArmTick := 0
global g_DesktopToRecycleGraceUntilTick := 0
global g_DesktopToRecycleSawSelectKeyUp := false
global g_DesktopToRecycleSessionId := 0
global DESKTOP_TO_RECYCLE_KEYS_GRACE_MS := 1500  ; ignore Y/N until grace ends AND keys have been up
global DESKTOP_TO_RECYCLE_DECISION_MS := 6000

; #region agent log
DesktopToRecycle_DebugLog(hypothesisId, location, message, dataMap := "") {
    logPath := A_ScriptDir "\debug-65068c.log"
    dataStr := "{}"
    if (IsObject(dataMap)) {
        dataStr := "{"
        first := true
        for k, v in dataMap {
            if (!first)
                dataStr .= ","
            first := false
            vs := String(v)
            vs := StrReplace(vs, "\", "\\")
            vs := StrReplace(vs, '"', '\"')
            dataStr .= '"' k '":"' vs '"'
        }
        dataStr .= "}"
    }
    line := '{"sessionId":"65068c","hypothesisId":"' hypothesisId '","location":"' location '","message":"' message '","data":' dataStr ',"timestamp":' A_TickCount '}`n'
    try FileAppend(line, logPath, "UTF-8")
    catch {
    }
}
; #endregion

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
    ; #region agent log
    DesktopToRecycle_DebugLog("B3", "desktop_recycle.ahk:TryArmKeys", "keys_armed_clean", Map("tick", A_TickCount))
    ; #endregion
}

DesktopToRecycle_StartKeysArm() {
    global g_DesktopToRecycleKeysArmTick, g_DesktopToRecycleGraceUntilTick, g_DesktopToRecycleSawSelectKeyUp
    global DESKTOP_TO_RECYCLE_KEYS_GRACE_MS
    DesktopToRecycle_StopKeysArmTimer()
    g_DesktopToRecycleSawSelectKeyUp := false
    g_DesktopToRecycleKeysArmTick := 0  ; not armed until TryArmKeys succeeds
    g_DesktopToRecycleGraceUntilTick := A_TickCount + DESKTOP_TO_RECYCLE_KEYS_GRACE_MS
    ; #region agent log
    DesktopToRecycle_DebugLog("B3", "desktop_recycle.ahk:StartKeysArm", "keys_arm_scheduled", Map("graceMs",
        DESKTOP_TO_RECYCLE_KEYS_GRACE_MS, "graceUntil", g_DesktopToRecycleGraceUntilTick))
    ; #endregion
    SetTimer(DesktopToRecycle_TryArmKeys, 50)
}

DesktopToRecycle_OnConfirm(*) {
    global g_DesktopToRecycleCloseHwnd
    if (!DesktopToRecycle_KeysArmed()) {
        ; #region agent log
        DesktopToRecycle_DebugLog("B3", "desktop_recycle.ahk:OnConfirm", "confirm_ignored_grace", Map("hwnd",
            g_DesktopToRecycleCloseHwnd))
        ; #endregion
        return
    }
    ; #region agent log
    DesktopToRecycle_DebugLog("B", "desktop_recycle.ahk:OnConfirm", "confirm_fired", Map("hwnd",
        g_DesktopToRecycleCloseHwnd))
    ; #endregion
    DesktopToRecycle_StopKeysArmTimer()
    DesktopToRecycle_EndDecisionSession()
    DesktopToRecycle_StopTrack()
    DesktopToRecycle_ClosePreviewExplorer()
    PlayCleaningDesktopSound()
    DesktopToRecycle_Run()
}

DesktopToRecycle_OnCancel(*) {
    global g_DesktopToRecycleCloseHwnd
    ; #region agent log
    DesktopToRecycle_DebugLog("B", "desktop_recycle.ahk:OnCancel", "cancel_fired", Map(
        "hwnd", g_DesktopToRecycleCloseHwnd,
        "escP", GetKeyState("Escape", "P") ? 1 : 0,
        "nP", GetKeyState("N", "P") ? 1 : 0,
        "asyncEsc", (DllCall("user32\GetAsyncKeyState", "int", 0x1B) & 0x8000) ? 1 : 0))
    ; #endregion
    DesktopToRecycle_StopKeysArmTimer()
    DesktopToRecycle_EndDecisionSession()
    DesktopToRecycle_StopTrack()
    DesktopToRecycle_ClosePreviewExplorer()
    ShowCenteredOverlay_Utils("⚠ Desktop cleanup cancelled", 1500, BANNER_ACCENT_INTERMEDIATE)
}

DesktopToRecycle_OnCancelFromN(*) {
    global g_DesktopToRecycleCloseHwnd
    if (!DesktopToRecycle_KeysArmed()) {
        ; #region agent log
        DesktopToRecycle_DebugLog("B3", "desktop_recycle.ahk:OnCancelFromN", "cancel_ignored_grace", Map(
            "nP", GetKeyState("N", "P") ? 1 : 0,
            "hwnd", g_DesktopToRecycleCloseHwnd))
        ; #endregion
        return
    }
    ; #region agent log
    DesktopToRecycle_DebugLog("B2", "desktop_recycle.ahk:OnCancelFromN", "cancel_source_N", Map("nP", GetKeyState("N",
        "P") ? 1 : 0))
    ; #endregion
    DesktopToRecycle_OnCancel()
}

DesktopToRecycle_OnTimeout(*) {
    global g_DesktopToRecycleCloseHwnd
    ; #region agent log
    DesktopToRecycle_DebugLog("B", "desktop_recycle.ahk:OnTimeout", "timeout_fired", Map("hwnd",
        g_DesktopToRecycleCloseHwnd))
    ; #endregion
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

; Center hwnd at 50% of the given work area; apply 50% opacity (once per place).
DesktopToRecycle_PlacePreviewOnWorkArea(hwnd, workLeft, workTop, workRight, workBottom) {
    global DESKTOP_TO_RECYCLE_PREVIEW_OPACITY, DESKTOP_TO_RECYCLE_PREVIEW_SCALE
    if (!hwnd || !WinExist("ahk_id " hwnd))
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
    try WinMove(x, y, w, h, "ahk_id " hwnd)
    catch {
        return false
    }
    try WinSetTransparent(DESKTOP_TO_RECYCLE_PREVIEW_OPACITY, "ahk_id " hwnd)
    catch {
    }
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
    global g_DesktopToRecycleCloseHwnd, g_DesktopToRecycleTrackLastMonIdx
    hwnd := g_DesktopToRecycleCloseHwnd
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return
    monIdx := g_DesktopToRecycleTrackLastMonIdx
    if (monIdx < 1)
        monIdx := GetMonitorIndexForForeground_StandardBar()
    DesktopToRecycle_MarkAutoSlotExclude(hwnd)
    DesktopToRecycle_PlacePreviewOnMonitor(hwnd, monIdx)
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

; Keep preview pinned (no focus-follow). Logs: after banner, focus steal made TrackTick
; move Explorer off the user's monitor — felt like the banner killed the window.
DesktopToRecycle_TrackTick(*) {
    global g_DesktopToRecycleCloseHwnd, g_DesktopToRecycleTrackLastMonIdx
    hwnd := g_DesktopToRecycleCloseHwnd
    if (!hwnd || !WinExist("ahk_id " hwnd)) {
        ; #region agent log
        DesktopToRecycle_DebugLog("C", "desktop_recycle.ahk:TrackTick", "hwnd_gone", Map("hwnd", hwnd, "exist", 0))
        ; #endregion
        DesktopToRecycle_StopTrack()
        return
    }
    DesktopToRecycle_MarkAutoSlotExclude(hwnd)
    try WinSetAlwaysOnTop(true, "ahk_id " hwnd)
    catch {
    }
    try {
        if (WinGetMinMax("ahk_id " hwnd) = 1) {
            ; #region agent log
            DesktopToRecycle_DebugLog("C", "desktop_recycle.ahk:TrackTick", "was_maximized_restore", Map("hwnd", hwnd))
            ; #endregion
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
    global g_DesktopToRecycleSessionId, DESKTOP_TO_RECYCLE_DECISION_MS
    g_DesktopToRecycleSessionId := A_TickCount
    sid := g_DesktopToRecycleSessionId
    SetTimer(DesktopToRecycle_SessionExpired.Bind(sid), -DESKTOP_TO_RECYCLE_DECISION_MS)
    ; #region agent log
    DesktopToRecycle_DebugLog("E2", "desktop_recycle.ahk:BeginDecisionSession", "session_timer_armed", Map("sid", sid,
        "ms", DESKTOP_TO_RECYCLE_DECISION_MS))
    ; #endregion
}

DesktopToRecycle_EndDecisionSession() {
    global g_DesktopToRecycleSessionId
    ; Invalidate any pending SessionExpired bind.
    g_DesktopToRecycleSessionId := 0
}

DesktopToRecycle_SessionExpired(sid, *) {
    global g_DesktopToRecycleSessionId, g_DesktopToRecycleCloseHwnd
    if (sid != g_DesktopToRecycleSessionId)
        return
    if (!g_DesktopToRecycleCloseHwnd) {
        ; #region agent log
        DesktopToRecycle_DebugLog("E2", "desktop_recycle.ahk:SessionExpired", "session_already_done", Map("sid", sid))
        ; #endregion
        DesktopToRecycle_EndDecisionSession()
        return
    }
    ; #region agent log
    DesktopToRecycle_DebugLog("E2", "desktop_recycle.ahk:SessionExpired", "session_expired_run", Map("sid", sid,
        "hwnd", g_DesktopToRecycleCloseHwnd))
    ; #endregion
    DesktopToRecycle_OnTimeout()
}

; Close only the temporary preview Explorer hwnd (not every Desktop Explorer).
DesktopToRecycle_ClosePreviewExplorer() {
    global g_DesktopToRecycleCloseHwnd, g_DesktopToRecycleWeOpenedExplorer
    DesktopToRecycle_StopKeysArmTimer()
    DesktopToRecycle_EndDecisionSession()
    DesktopToRecycle_StopTrack()
    hwnd := g_DesktopToRecycleCloseHwnd
    ; #region agent log
    DesktopToRecycle_DebugLog("B", "desktop_recycle.ahk:ClosePreviewExplorer", "close_preview", Map("hwnd", hwnd,
        "exist", (hwnd && WinExist("ahk_id " hwnd)) ? 1 : 0))
    ; #endregion
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
    if (WinExist("ahk_id " hwnd)) {
        try WinClose("ahk_id " hwnd)
        catch {
            try WinKill("ahk_id " hwnd)
            catch {
            }
        }
    }
}

; Open a new Desktop Explorer, exclude from AutoSlot, place at 50%/50% opacity. Returns hwnd or 0.
DesktopToRecycle_OpenPreviewExplorer(targetPath, workLeft, workTop, workRight, workBottom) {
    global g_DesktopToRecycleWeOpenedExplorer
    g_DesktopToRecycleWeOpenedExplorer := false
    if (!targetPath || !DirExist(targetPath)) {
        ; #region agent log
        DesktopToRecycle_DebugLog("A", "desktop_recycle.ahk:OpenPreview", "bad_path", Map("path", String(targetPath)))
        ; #endregion
        return 0
    }

    before := DesktopToRecycle_CollectDesktopExplorerHwnds(targetPath)
    beforeCount := 0
    for , _ in before
        beforeCount += 1
    ; #region agent log
    DesktopToRecycle_DebugLog("A", "desktop_recycle.ahk:OpenPreview", "before_run", Map("beforeCount", beforeCount,
        "path", targetPath))
    ; #endregion
    ; Suppress AutoSlot BEFORE Run — SHOW/Schedule races SetProp by hundreds of ms.
    DesktopToRecycle_BeginAutoSlotSuppress(12000)
    try Run('explorer.exe /n,"' targetPath '"')
    catch {
        try Run('explorer.exe "' targetPath '"')
        catch {
            DesktopToRecycle_EndAutoSlotSuppress()
            ; #region agent log
            DesktopToRecycle_DebugLog("A", "desktop_recycle.ahk:OpenPreview", "run_failed", Map())
            ; #endregion
            return 0
        }
    }

    hwnd := 0
    deadline := A_TickCount + 2500
    while (A_TickCount < deadline) {
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
            break
        }
        Sleep 40
    }
    foundViaFallback := 0
    if (!hwnd) {
        hwnd := DesktopToRecycle_FindDesktopExplorer(targetPath)
        foundViaFallback := 1
    }
    rejectedReuse := (!hwnd || !WinExist("ahk_id " hwnd) || before.Has(hwnd)) ? 1 : 0
    ; #region agent log
    DesktopToRecycle_DebugLog("A", "desktop_recycle.ahk:OpenPreview", "after_find", Map("hwnd", hwnd, "fallback",
        foundViaFallback, "rejectedReuse", rejectedReuse, "beforeHas", (hwnd && before.Has(hwnd)) ? 1 : 0, "elapsedMs",
        A_TickCount - (deadline - 2500)))
    ; #endregion
    ; Never adopt a pre-existing Desktop Explorer (would resize/opacity/close the user's window).
    if (rejectedReuse) {
        DesktopToRecycle_EndAutoSlotSuppress()
        return 0
    }

    DesktopToRecycle_MarkAutoSlotExclude(hwnd)
    g_DesktopToRecycleWeOpenedExplorer := true
    ; Brief settle so the first paint is stable before WinMove/transparent (reduces blink).
    Sleep 120

    placed := DesktopToRecycle_PlacePreviewOnWorkArea(hwnd, workLeft, workTop, workRight, workBottom)
    ; #region agent log
    mm := -999
    try mm := WinGetMinMax("ahk_id " hwnd)
    catch {
    }
    DesktopToRecycle_DebugLog("D", "desktop_recycle.ahk:OpenPreview", "after_place", Map("hwnd", hwnd, "placed", placed ?
        1 : 0, "minmax", mm))
    ; #endregion
    if (!placed) {
        DesktopToRecycle_ClearAutoSlotExclude(hwnd)
        DesktopToRecycle_EndAutoSlotSuppress()
        try WinClose("ahk_id " hwnd)
        catch {
        }
        g_DesktopToRecycleWeOpenedExplorer := false
        return 0
    }
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

; Entry point for Desktop to Recycle macro (^!#8)
DesktopToRecycle_Trigger() {
    global g_DesktopToRecycleCloseHwnd, g_DesktopToRecyclePath, g_DesktopToRecycleWeOpenedExplorer
    ; #region agent log
    DesktopToRecycle_DebugLog("E", "desktop_recycle.ahk:Trigger", "trigger_enter", Map("tick", A_TickCount))
    ; #endregion
    DesktopToRecycle_StopTrack()
    DesktopToRecycle_ClosePreviewExplorer()
    g_DesktopToRecycleWeOpenedExplorer := false
    g_DesktopToRecyclePath := GetDesktopToRecyclePath()
    path := g_DesktopToRecyclePath
    if (!path || path = "" || !DirExist(path))
        path := A_Desktop

    ; Work area of the monitor with the *current* active window — before Explorer steals focus.
    GetActiveMonitorWorkArea_StandardBar(&workLeft, &workTop, &workRight, &workBottom)
    initialMonIdx := GetMonitorIndexForForeground_StandardBar()

    StandardLoadingBar_CloseKeysOverlay()
    StandardLoadingBar_Hide(0)
    Sleep 50

    hwnd := DesktopToRecycle_OpenPreviewExplorer(path, workLeft, workTop, workRight, workBottom)
    g_DesktopToRecycleCloseHwnd := hwnd ? hwnd : 0
    if (hwnd)
        DesktopToRecycle_StartTrack(hwnd, initialMonIdx)
    ; #region agent log
    vis := 0
    mm := -999
    if (hwnd) {
        try vis := DllCall("IsWindowVisible", "ptr", hwnd) ? 1 : 0
        try mm := WinGetMinMax("ahk_id " hwnd)
        catch {
        }
    }
    DesktopToRecycle_DebugLog("E", "desktop_recycle.ahk:Trigger", "before_banner", Map("hwnd", hwnd, "mon",
        initialMonIdx, "suppressActive", DesktopToRecycle_AutoSlotSuppressActive() ? 1 : 0, "visible", vis, "minmax",
        mm))
    ; #endregion

    try KeyWait("N")
    try KeyWait("Y")
    catch {
    }

    DesktopToRecycle_StartKeysArm()
    DesktopToRecycle_BeginDecisionSession()

    global DESKTOP_TO_RECYCLE_DECISION_MS
    state := "🗑️ Move all items from:`n" . g_DesktopToRecyclePath . "`nto Recycle Bin? (6s)"
    ; Banner does not destroy Explorer (logs: visible=1 after ShowWithKeys). Vanish was either
    ; focus-follow moving it off-monitor, or keys-overlay dismissed without OnTimeout (Explorer leaked).
    ; Fixed placement + AlwaysOnTop + independent session timer.
    keyCallbacks := Map(
        "Y", DesktopToRecycle_OnConfirm,
        "N", DesktopToRecycle_OnCancelFromN)
    StandardLoadingBar_ShowWithKeys(
        state,
        keyCallbacks,
        DESKTOP_TO_RECYCLE_DECISION_MS,
        0,
        DesktopToRecycle_OnTimeout,
        BANNER_ACCENT_INTERMEDIATE,
        0,
        17,
        "",
        false,
        "[Y] Yes  [N] Cancel",
        false,
        true,
        true,
        "",
        true)
    ; #region agent log
    vis2 := 0
    mm2 := -999
    if (g_DesktopToRecycleCloseHwnd) {
        try vis2 := DllCall("IsWindowVisible", "ptr", g_DesktopToRecycleCloseHwnd) ? 1 : 0
        try mm2 := WinGetMinMax("ahk_id " g_DesktopToRecycleCloseHwnd)
        catch {
        }
    }
    DesktopToRecycle_DebugLog("E", "desktop_recycle.ahk:Trigger", "after_showwithkeys_returns", Map("hwndStill",
        g_DesktopToRecycleCloseHwnd, "exist", (g_DesktopToRecycleCloseHwnd && WinExist("ahk_id " g_DesktopToRecycleCloseHwnd
        )) ? 1 : 0, "visible", vis2, "minmax", mm2, "runId", "post-fix2"))
    ; #endregion
}
