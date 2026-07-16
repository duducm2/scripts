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

; After native Shift+P / Shift+B open: capture prior monitor, then maximize + activate (no UIA / no toggle).
; Quality gate: single early pass often races Clip Angel's own restore — poll + deferred retry.
CLIPANGEL_NATIVE_OPEN_SETTLE_MS := 200
CLIPANGEL_NATIVE_OPEN_RETRY_MS := 500
CLIPANGEL_NATIVE_OPEN_POLL_MS := 700
global g_ClipAngelNativeOpenTargetMon := 0
global g_ClipAngelNativeOpenGen := 0

; Used by WindowManagement #HotIf: assist when Clip Angel is focused but still iconic/tiny.
ClipAngel_HotIfNeedsForegroundAssist() {
    hwnd := ClipAngel_MainHwnd()
    if !hwnd
        return false
    return !ClipAngel_IsWindowShown(hwnd) || ClipAngel_NeedsLayoutCorrection(hwnd)
}

ClipAngel_EnsureForegroundAfterNativeOpen() {
    global g_ClipAngelNativeOpenTargetMon, g_ClipAngelNativeOpenGen
    targetMon := 0
    try {
        activeHwnd := WinGetID("A")
        if (activeHwnd)
            targetMon := GetAhkMonitorIndexFromHwnd(activeHwnd)
    } catch {
        targetMon := 0
    }
    g_ClipAngelNativeOpenTargetMon := targetMon
    g_ClipAngelNativeOpenGen += 1
    gen := g_ClipAngelNativeOpenGen
    SetTimer(() => ClipAngel_ApplyForegroundMaximizePass(gen), -CLIPANGEL_NATIVE_OPEN_SETTLE_MS)
    SetTimer(() => ClipAngel_ApplyForegroundMaximizePass(gen), -CLIPANGEL_NATIVE_OPEN_RETRY_MS)
}

; Force restore/show + maximize on targetMon + activate; poll until stable or timeout.
ClipAngel_ApplyForegroundMaximizePass(gen) {
    global g_ClipAngelNativeOpenTargetMon, g_ClipAngelNativeOpenGen
    if (gen != g_ClipAngelNativeOpenGen)
        return
    targetMon := g_ClipAngelNativeOpenTargetMon
    hwnd := ClipAngel_MainHwnd()
    if !hwnd
        return

    deadline := A_TickCount + CLIPANGEL_NATIVE_OPEN_POLL_MS
    loop {
        try {
            if (WinGetMinMax("ahk_id " hwnd) = -1 || DllCall("IsIconic", "ptr", hwnd))
                WinRestore("ahk_id " hwnd)
        } catch {
        }
        try WinShow("ahk_id " hwnd)
        catch {
        }
        ; Always apply layout (not only when "tiny") — visible-but-wrong-monitor was a miss.
        ClipAngel_ApplyLayoutOnMonitor(hwnd, targetMon)
        ClipAngel_EnsureWindowActive(hwnd)

        shown := ClipAngel_IsWindowShown(hwnd)
        active := !!WinActive("ahk_id " hwnd)
        onMon := true
        if (targetMon >= 1) {
            try onMon := (GetAhkMonitorIndexFromHwnd(hwnd) = targetMon)
            catch
                onMon := false
        }
        if (shown && active && onMon)
            return
        if (A_TickCount >= deadline)
            break
        Sleep 50
    }
    ; Final force attempt after poll budget.
    ClipAngel_ApplyLayoutOnMonitor(hwnd, targetMon)
    ClipAngel_EnsureWindowActive(hwnd)
}
