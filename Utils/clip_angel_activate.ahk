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
; Quality gate: poll + multi-pass retry; strong foreground (AttachThreadInput); cursor-monitor fallback.
CLIPANGEL_NATIVE_OPEN_SETTLE_MS := 200
CLIPANGEL_NATIVE_OPEN_RETRY_MS := 500
CLIPANGEL_NATIVE_OPEN_LATE_MS := 1100
CLIPANGEL_NATIVE_OPEN_POLL_MS := 800
global g_ClipAngelNativeOpenTargetMon := 0
global g_ClipAngelNativeOpenGen := 0

; Used by WindowManagement #HotIf: assist when Clip Angel is focused but still iconic/tiny.
ClipAngel_HotIfNeedsForegroundAssist() {
    hwnd := ClipAngel_MainHwnd()
    if !hwnd
        return false
    return !ClipAngel_IsWindowShown(hwnd) || ClipAngel_NeedsLayoutCorrection(hwnd)
}

; Monitor under the mouse — where the user is looking (better than ClipAngel's own hwnd when it already has focus).
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

; WinActivate often fails under foreground-lock; AttachThreadInput + SetForegroundWindow is more reliable.
ClipAngel_ForceForeground(hwnd) {
    if !hwnd || !WinExist("ahk_id " hwnd)
        return false
    if WinActive("ahk_id " hwnd)
        return true
    try DllCall("AllowSetForegroundWindow", "int", -1)  ; ASFW_ANY
    catch {
    }
    fg := 0
    try fg := DllCall("GetForegroundWindow", "ptr")
    catch {
    }
    tidCurr := DllCall("GetCurrentThreadId", "uint")
    tidFore := 0
    tidTarget := 0
    if (fg)
        tidFore := DllCall("GetWindowThreadProcessId", "ptr", fg, "ptr", 0, "uint")
    tidTarget := DllCall("GetWindowThreadProcessId", "ptr", hwnd, "ptr", 0, "uint")
    attachedFore := false
    attachedTarget := false
    try {
        if (tidFore && tidFore != tidCurr)
            attachedFore := !!DllCall("AttachThreadInput", "uint", tidCurr, "uint", tidFore, "int", 1)
        if (tidTarget && tidTarget != tidCurr && tidTarget != tidFore)
            attachedTarget := !!DllCall("AttachThreadInput", "uint", tidCurr, "uint", tidTarget, "int", 1)
        DllCall("BringWindowToTop", "ptr", hwnd)
        DllCall("SetForegroundWindow", "ptr", hwnd)
        try WinActivate("ahk_id " hwnd)
        catch {
        }
    } catch {
        try WinActivate("ahk_id " hwnd)
        catch {
        }
    }
    if (attachedFore)
        DllCall("AttachThreadInput", "uint", tidCurr, "uint", tidFore, "int", 0)
    if (attachedTarget)
        DllCall("AttachThreadInput", "uint", tidCurr, "uint", tidTarget, "int", 0)
    try WinWaitActive("ahk_id " hwnd, , 0.25)
    catch {
    }
    return !!WinActive("ahk_id " hwnd)
}

ClipAngel_EnsureForegroundAfterNativeOpen() {
    global g_ClipAngelNativeOpenTargetMon, g_ClipAngelNativeOpenGen
    targetMon := 0
    try {
        activeHwnd := WinGetID("A")
        if (activeHwnd && !WinActive("ahk_exe ClipAngel.exe"))
            targetMon := GetAhkMonitorIndexFromHwnd(activeHwnd)
        else
            targetMon := ClipAngel_GetMonitorIndexFromCursor()
    } catch {
        targetMon := ClipAngel_GetMonitorIndexFromCursor()
    }
    if (!targetMon)
        try targetMon := MonitorGetPrimary()
        catch
            targetMon := 1
    g_ClipAngelNativeOpenTargetMon := targetMon
    g_ClipAngelNativeOpenGen += 1
    gen := g_ClipAngelNativeOpenGen
    SetTimer(() => ClipAngel_ApplyForegroundMaximizePass(gen), -CLIPANGEL_NATIVE_OPEN_SETTLE_MS)
    SetTimer(() => ClipAngel_ApplyForegroundMaximizePass(gen), -CLIPANGEL_NATIVE_OPEN_RETRY_MS)
    SetTimer(() => ClipAngel_ApplyForegroundMaximizePass(gen), -CLIPANGEL_NATIVE_OPEN_LATE_MS)
}

; True when Clip Angel is shown, maximized (or large enough), on targetMon, and foreground.
ClipAngel_IsForegroundLayoutOk(hwnd, targetMon) {
    if !hwnd || !ClipAngel_IsWindowShown(hwnd)
        return false
    if !WinActive("ahk_id " hwnd)
        return false
    if (targetMon >= 1) {
        try {
            if (GetAhkMonitorIndexFromHwnd(hwnd) != targetMon)
                return false
        } catch {
            return false
        }
    }
    try {
        if (WinGetMinMax("ahk_id " hwnd) = 1)
            return true
        WinGetPos(, , &w, &h, "ahk_id " hwnd)
        return (w >= CLIPANGEL_MIN_LAYOUT_WIDTH && h >= CLIPANGEL_MIN_LAYOUT_HEIGHT)
    } catch {
        return false
    }
}

; Force restore/show + maximize on targetMon + activate; poll until stable or timeout.
ClipAngel_ApplyForegroundMaximizePass(gen) {
    global g_ClipAngelNativeOpenTargetMon, g_ClipAngelNativeOpenGen
    if (gen != g_ClipAngelNativeOpenGen)
        return
    targetMon := g_ClipAngelNativeOpenTargetMon

    deadline := A_TickCount + CLIPANGEL_NATIVE_OPEN_POLL_MS
    loop {
        hwnd := ClipAngel_MainHwnd()
        if !hwnd {
            if (A_TickCount >= deadline)
                return
            Sleep 50
            continue
        }
        try {
            if (WinGetMinMax("ahk_id " hwnd) = -1 || DllCall("IsIconic", "ptr", hwnd))
                WinRestore("ahk_id " hwnd)
        } catch {
        }
        try WinShow("ahk_id " hwnd)
        catch {
        }
        ClipAngel_ApplyLayoutOnMonitor(hwnd, targetMon)
        if !ClipAngel_ForceForeground(hwnd)
            ClipAngel_EnsureWindowActive(hwnd, 400)

        if ClipAngel_IsForegroundLayoutOk(hwnd, targetMon)
            return
        if (A_TickCount >= deadline)
            break
        Sleep 50
    }
    hwnd := ClipAngel_MainHwnd()
    if !hwnd
        return
    ClipAngel_ApplyLayoutOnMonitor(hwnd, targetMon)
    if !ClipAngel_ForceForeground(hwnd)
        ClipAngel_EnsureWindowActive(hwnd, 400)
}
