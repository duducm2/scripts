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

; Alt+P / Alt+B: if Clip Angel already foreground and shown → minimize; else maximize on active monitor.
ClipAngel_ToggleOrShowOnActiveMonitor() {
    hwnd := ClipAngel_MainHwnd()
    if (hwnd && WinActive("ahk_id " hwnd) && ClipAngel_IsWindowShown(hwnd)) {
        ClipAngel_HideWindow(hwnd)
        return true
    }
    targetMon := 0
    try {
        activeHwnd := WinGetID("A")
        if (activeHwnd && hwnd && activeHwnd = hwnd) {
            ; Active is Clip Angel but not shown (edge) — fall through to layout on primary focus monitor.
        } else if (activeHwnd)
            targetMon := GetAhkMonitorIndexFromHwnd(activeHwnd)
    } catch {
        targetMon := 0
    }
    return ActivateClipAngelWithFocusCorrection(false, targetMon, false, true)
}
