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
