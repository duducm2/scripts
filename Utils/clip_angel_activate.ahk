; =============================================================================
; Utils module: clip_angel_activate.ahk
; Clip Angel activate with focus correction
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; Clip Angel: Open/Activate with focus correction (Row 0)
; =============================================================================
; Show/restore always-running Clip Angel via AHK (no native Alt+P paste hotkey). UIA: dataGridView,
; Row 0 (first clip) per clip-angel.txt. One ElementFromHandle per flow; bounded polls; layout only when not foreground/hidden.
; skipRow0: true when caller will select Row 0 after filter change (e.g. OnSubmitO leaving favorites).
; forceLayout: true always move/maximize onto targetMon (toggle-open path).
ActivateClipAngelWithFocusCorrection(silent := false, targetMon := 0, skipRow0 := false, forceLayout := false) {
    needBanner := false
    if !targetMon {
        try targetMon := GetAhkMonitorIndexFromHwnd(WinGetID("A"))
        catch
            targetMon := 0
    }
    hwnd := ClipAngel_MainHwnd()
    if !hwnd {
        if !silent
            ShowCenteredOverlay_Utils("❌ Clip Angel não está em execução.", 2000, BANNER_ACCENT_ERROR)
        return false
    }
    isActive := WinActive("ahk_id " hwnd)
    wasHidden := !ClipAngel_IsWindowShown(hwnd)
    wrongMonitor := false
    if (targetMon >= 1) {
        try wrongMonitor := (GetAhkMonitorIndexFromHwnd(hwnd) != targetMon)
        catch
            wrongMonitor := true
    }
    needsLayout := forceLayout || !isActive || wasHidden || wrongMonitor || ClipAngel_NeedsLayoutCorrection(hwnd)
    if isActive && !needsLayout {
        if !skipRow0
            ClipAngel_UiaEnsureRow0Selected(hwnd, false)
        return true
    }
    if wasHidden || !isActive || wrongMonitor || forceLayout {
        needBanner := !silent
        if needBanner
            ClipAngelBanner_Show("📂 Opening Clip Angel...", BANNER_ACCENT_INTERMEDIATE)
    }
    if !ClipAngel_ShowWindow(hwnd) {
        if needBanner
            ClipAngelBanner_Hide()
        if !silent
            ShowCenteredOverlay_Utils("❌ ClipAngel window not found.", 2000, BANNER_ACCENT_ERROR)
        return false
    }
    if needsLayout
        ClipAngel_ApplyLayoutOnMonitor(hwnd, targetMon)
    ClipAngel_EnsureWindowActive(hwnd)
    if !skipRow0
        ClipAngel_UiaEnsureRow0Selected(hwnd, true)
    if needBanner {
        ClipAngelBanner_Show("✅ Done", BANNER_ACCENT_SUCCESS)
        SetTimer(ClipAngelBanner_Hide, -500)
    }
    return true
}

; After native Alt+P / Alt+B open: one settle timer, one maximize+foreground gate (no poll / multi-pass).
; Efficiency canon: hotkey returns immediately; MoveWindowToMonitor restores — calling it repeatedly = flicker loop.
CLIPANGEL_NATIVE_OPEN_SETTLE_MS := 300
global g_ClipAngelNativeOpenTargetMon := 0

; Monitor under the mouse — where the user is looking.
ClipAngel_GetMonitorIndexFromCursor() {
    CoordMode "Mouse", "Screen"
    MouseGetPos &mx, &my
    try {
        loop MonitorGetCount() {
            MonitorGet A_Index, &l, &t, &r, &b
            if (mx >= l && mx < r && my >= t && my < b)
                return A_Index
        }
    } catch {
    }
    try return MonitorGetPrimary()
    catch
        return 1
}

; True when shown, maximized, on targetMon, and foreground — sole quality gate.
ClipAngel_IsForegroundLayoutOk(hwnd, targetMon) {
    if !hwnd || !ClipAngel_IsWindowShown(hwnd)
        return false
    if !WinActive("ahk_id " hwnd)
        return false
    try {
        if (WinGetMinMax("ahk_id " hwnd) != 1)
            return false
    } catch {
        return false
    }
    if (targetMon >= 1) {
        try {
            if (GetAhkMonitorIndexFromHwnd(hwnd) != targetMon)
                return false
        } catch {
            return false
        }
    }
    return true
}

; Capture look-at monitor; replace any pending settle timer (second Alt+P/B does not stack layouts).
ClipAngel_EnsureForegroundAfterNativeOpen() {
    global g_ClipAngelNativeOpenTargetMon
    targetMon := ClipAngel_GetMonitorIndexFromCursor()
    if (!targetMon || targetMon < 1) {
        try targetMon := MonitorGetPrimary()
        catch
            targetMon := 1
    }
    g_ClipAngelNativeOpenTargetMon := targetMon
    SetTimer(ClipAngel_ApplyForegroundMaximizeOnce, -CLIPANGEL_NATIVE_OPEN_SETTLE_MS)
}

; One layout+activate if gate fails; no inner poll loop.
ClipAngel_ApplyForegroundMaximizeOnce(*) {
    global g_ClipAngelNativeOpenTargetMon
    targetMon := g_ClipAngelNativeOpenTargetMon
    hwnd := ClipAngel_MainHwnd()
    if !hwnd
        return
    if ClipAngel_IsForegroundLayoutOk(hwnd, targetMon)
        return
    try {
        if (WinGetMinMax("ahk_id " hwnd) = -1 || DllCall("IsIconic", "ptr", hwnd))
            WinRestore("ahk_id " hwnd)
    } catch {
    }
    try WinShow("ahk_id " hwnd)
    catch {
    }
    ClipAngel_ApplyLayoutOnMonitor(hwnd, targetMon)
    ClipAngel_EnsureWindowActive(hwnd, 400)
}

; =============================================================================
; Auto-minimize when Clip Angel loses focus to another process
; =============================================================================
CLIPANGEL_AUTO_MINIMIZE_DEBOUNCE_MS := 250
global g_ClipAngelFgHook := 0
global g_ClipAngelFgHookCb := 0
global g_ClipAngelLastFgHwnd := 0
; Default false here (Utils hosts); Shift keys hotif_clipangel toggles it for the filter selector.
global g_ClipAngelFilterSelectorActive := false

; #region agent log
ClipAngel_DebugLog(hypothesisId, location, message, dataObj := "") {
    try {
        path := A_ScriptDir "\debug-be5edc.log"
        ts := A_TickCount
        dataStr := "{}"
        if (IsObject(dataObj)) {
            parts := []
            for k, v in dataObj {
                vv := v
                if (vv = true)
                    vv := "true"
                else if (vv = false)
                    vv := "false"
                else if (vv is String)
                    vv := '"' StrReplace(StrReplace(vv, "\", "\\"), '"', '\"') '"'
                parts.Push('"' k '":' vv)
            }
            dataStr := "{" . ArrayJoinComma(parts) . "}"
        }
        line := '{"sessionId":"be5edc","hypothesisId":"' hypothesisId '","location":"' location '","message":"' message '","data":' dataStr ',"timestamp":' ts ',"runId":"pre-fix"}`n'
        FileAppend(line, path, "UTF-8")
    } catch {
    }
}
ArrayJoinComma(arr) {
    out := ""
    for i, s in arr
        out .= (i > 1 ? "," : "") s
    return out
}
; #endregion

ClipAngel_InitAutoMinimizeOnDeactivate() {
    global g_ClipAngelFgHook, g_ClipAngelFgHookCb
    if (g_ClipAngelFgHook) {
        ; #region agent log
        ClipAngel_DebugLog("A", "clip_angel_activate.ahk:Init", "init_already_hooked", Map("hook", g_ClipAngelFgHook))
        ; #endregion
        return
    }
    g_ClipAngelFgHookCb := CallbackCreate(ClipAngel_ForegroundHookProc, "F", 7)
    g_ClipAngelFgHook := DllCall("user32\SetWinEventHook",
        "UInt", 0x0003,  ; EVENT_SYSTEM_FOREGROUND
        "UInt", 0x0003,
        "Ptr", 0,
        "Ptr", g_ClipAngelFgHookCb,
        "UInt", 0,
        "UInt", 0,
        "UInt", 0,
        "Ptr")
    ; #region agent log
    ClipAngel_DebugLog("A", "clip_angel_activate.ahk:Init", "init_result", Map(
        "hook", g_ClipAngelFgHook + 0,
        "cb", g_ClipAngelFgHookCb + 0,
        "script", A_ScriptName
    ))
    ; #endregion
    if (!g_ClipAngelFgHook) {
        try CallbackFree(g_ClipAngelFgHookCb)
        catch {
        }
        g_ClipAngelFgHookCb := 0
    }
}

; Hook thread: store hwnd and bounce to main thread.
ClipAngel_ForegroundHookProc(hHook, event, hwnd, idObject, idChild, idEventThread, dwmsEventTime) {
    global g_ClipAngelLastFgHwnd
    g_ClipAngelLastFgHwnd := hwnd
    ; #region agent log
    ; Throttle: only log when ClipAngel.exe exists (avoid flood).
    if (WinExist("ahk_exe ClipAngel.exe"))
        ClipAngel_DebugLog("B", "clip_angel_activate.ahk:FgHook", "fg_event", Map("hwnd", hwnd + 0))
    ; #endregion
    ; -0 is treated as period 0 (disable timer). Must use -1 for next-tick once.
    SetTimer(ClipAngel_OnForegroundChanged, -1)
}

ClipAngel_ShouldSkipAutoMinimize() {
    global g_ClipAngelAutomationBusy, g_ClipAngelFilterSelectorActive
    if (g_ClipAngelAutomationBusy)
        return true
    if (g_ClipAngelFilterSelectorActive)
        return true
    return false
}

ClipAngel_FgIsClipAngelExe(hwnd) {
    if (!hwnd)
        return false
    try return StrLower(WinGetProcessName("ahk_id " hwnd)) = "clipangel.exe"
    catch
        return false
}

; Schedule or cancel debounced minimize when foreground changes.
ClipAngel_OnForegroundChanged(*) {
    global g_ClipAngelAutomationBusy, g_ClipAngelFilterSelectorActive, g_ClipAngelLastFgHwnd
    ; #region agent log
    ClipAngel_DebugLog("C", "clip_angel_activate.ahk:OnFgChanged", "entered", Map())
    ; #endregion
    skip := ClipAngel_ShouldSkipAutoMinimize()
    ca := ClipAngel_MainHwnd()
    shown := ca ? ClipAngel_IsWindowShown(ca) : false
    fg := g_ClipAngelLastFgHwnd
    if (!fg)
        fg := WinExist("A")
    fgExe := ""
    try fgExe := StrLower(WinGetProcessName("ahk_id " fg))
    catch
        fgExe := ""
    sameCa := (fg = ca)
    isCaExe := ClipAngel_FgIsClipAngelExe(fg)
    ; #region agent log
    ClipAngel_DebugLog("C", "clip_angel_activate.ahk:OnFgChanged", "evaluate", Map(
        "skip", skip,
        "busy", !!g_ClipAngelAutomationBusy,
        "filter", !!g_ClipAngelFilterSelectorActive,
        "ca", ca + 0,
        "shown", !!shown,
        "fg", fg + 0,
        "fgExe", fgExe,
        "sameCa", !!sameCa,
        "isCaExe", !!isCaExe
    ))
    ; #endregion
    if (skip) {
        SetTimer(ClipAngel_AutoMinimizeTick, 0)
        return
    }
    if (!ca || !shown) {
        SetTimer(ClipAngel_AutoMinimizeTick, 0)
        return
    }
    if (sameCa || isCaExe) {
        SetTimer(ClipAngel_AutoMinimizeTick, 0)
        return
    }
    ; #region agent log
    ClipAngel_DebugLog("D", "clip_angel_activate.ahk:OnFgChanged", "schedule_minimize", Map("debounceMs",
        CLIPANGEL_AUTO_MINIMIZE_DEBOUNCE_MS))
    ; #endregion
    SetTimer(ClipAngel_AutoMinimizeTick, -CLIPANGEL_AUTO_MINIMIZE_DEBOUNCE_MS)
}

ClipAngel_AutoMinimizeTick(*) {
    if (ClipAngel_ShouldSkipAutoMinimize()) {
        ; #region agent log
        ClipAngel_DebugLog("E", "clip_angel_activate.ahk:Tick", "tick_skip_busy", Map())
        ; #endregion
        return
    }
    ca := ClipAngel_MainHwnd()
    if (!ca || !ClipAngel_IsWindowShown(ca)) {
        ; #region agent log
        ClipAngel_DebugLog("E", "clip_angel_activate.ahk:Tick", "tick_skip_not_shown", Map("ca", ca + 0))
        ; #endregion
        return
    }
    fg := WinExist("A")
    if (fg = ca || ClipAngel_FgIsClipAngelExe(fg)) {
        ; #region agent log
        ClipAngel_DebugLog("E", "clip_angel_activate.ahk:Tick", "tick_skip_still_ca", Map("fg", fg + 0))
        ; #endregion
        return
    }
    ok := ClipAngel_HideWindow(ca)
    ; #region agent log
    ClipAngel_DebugLog("E", "clip_angel_activate.ahk:Tick", "tick_hide", Map("ca", ca + 0, "ok", !!ok, "stillShown", !!
        ClipAngel_IsWindowShown(ca)))
    ; #endregion
}
