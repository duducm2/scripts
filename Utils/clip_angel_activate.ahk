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

; After native Alt+P / Alt+B open: short settle, one maximize gate, one retry if needed.
; Efficiency canon: hotkey returns immediately; same-monitor path skips MoveWindow+Sleep.
CLIPANGEL_NATIVE_OPEN_SETTLE_MS := 50
CLIPANGEL_NATIVE_OPEN_RETRY_MS := 100
global g_ClipAngelNativeOpenTargetMon := 0
global g_ClipAngelNativeOpenRetryArmed := false

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
    global g_ClipAngelNativeOpenTargetMon, g_ClipAngelNativeOpenRetryArmed
    targetMon := ClipAngel_GetMonitorIndexFromCursor()
    if (!targetMon || targetMon < 1) {
        try targetMon := MonitorGetPrimary()
        catch
            targetMon := 1
    }
    g_ClipAngelNativeOpenTargetMon := targetMon
    g_ClipAngelNativeOpenRetryArmed := false
    SetTimer(ClipAngel_ApplyForegroundMaximizeOnce, -CLIPANGEL_NATIVE_OPEN_SETTLE_MS)
}

; One layout+activate if gate fails; one short retry if still not ok.
ClipAngel_ApplyForegroundMaximizeOnce(*) {
    global g_ClipAngelNativeOpenTargetMon, g_ClipAngelNativeOpenRetryArmed
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
    if ClipAngel_IsForegroundLayoutOk(hwnd, targetMon)
        return
    if (!g_ClipAngelNativeOpenRetryArmed) {
        g_ClipAngelNativeOpenRetryArmed := true
        SetTimer(ClipAngel_ApplyForegroundMaximizeOnce, -CLIPANGEL_NATIVE_OPEN_RETRY_MS)
    }
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

ClipAngel_InitAutoMinimizeOnDeactivate() {
    global g_ClipAngelFgHook, g_ClipAngelFgHookCb
    if (g_ClipAngelFgHook)
        return
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
    if (ClipAngel_ShouldSkipAutoMinimize()) {
        SetTimer(ClipAngel_AutoMinimizeTick, 0)
        return
    }
    ca := ClipAngel_MainHwnd()
    if (!ca || !ClipAngel_IsWindowShown(ca)) {
        SetTimer(ClipAngel_AutoMinimizeTick, 0)
        return
    }
    global g_ClipAngelLastFgHwnd
    fg := g_ClipAngelLastFgHwnd
    if (!fg)
        fg := WinExist("A")
    if (fg = ca || ClipAngel_FgIsClipAngelExe(fg)) {
        SetTimer(ClipAngel_AutoMinimizeTick, 0)
        return
    }
    SetTimer(ClipAngel_AutoMinimizeTick, -CLIPANGEL_AUTO_MINIMIZE_DEBOUNCE_MS)
}

ClipAngel_AutoMinimizeTick(*) {
    if (ClipAngel_ShouldSkipAutoMinimize())
        return
    ca := ClipAngel_MainHwnd()
    if (!ca || !ClipAngel_IsWindowShown(ca))
        return
    fg := WinExist("A")
    if (fg = ca || ClipAngel_FgIsClipAngelExe(fg))
        return
    ClipAngel_HideWindow(ca)
}

; Escape while Clip Angel is focused: close filter overlay if any, then minimize.
; Cleanup lives in Shift keys\hotif_clipangel.ahk — call by name so Utils hosts do not #Warn.
ClipAngel_EscapeMinimize() {
    global g_ClipAngelFilterSelectorActive
    if (g_ClipAngelFilterSelectorActive) {
        fnName := "CleanupClipAngelFilterSelector"
        try %fnName%()
        catch {
        }
    }
    ClipAngel_CloseAndRestoreFocus(0)
}

; Resolve ClipAngel.exe for hard restart.
; Personal portable: C:\Users\eduev\Documents\ClipAngel\ClipAngel 2.13\…
; Work portable: OneDrive Bosch Documents\ClipAngel\ClipAngel.exe (also legacy
; Handy-style under fie7ca\Documents\ClipAngel 2.23 — docs/reference/clip-angel.txt).
; A_MyDocuments may be OneDrive — prefer known exact paths, then scan roots.
ClipAngel_ResolveExePath() {
    global IS_WORK_ENVIRONMENT
    try {
        hwnd := ClipAngel_MainHwnd()
        if (hwnd) {
            path := WinGetProcessPath("ahk_id " hwnd)
            if (path != "" && FileExist(path))
                return path
        }
    } catch {
    }
    if ProcessExist("ClipAngel.exe") {
        try {
            path := ProcessGetPath("ClipAngel.exe")
            if (path != "" && FileExist(path))
                return path
        } catch {
        }
    }

    ; Optional env.ahk helper when present — call by name so hosts without it do not #Warn.
    try {
        fnName := "GetClipAngelExePath"
        envPath := %fnName%()
        if (envPath != "" && FileExist(envPath))
            return envPath
    } catch {
    }

    personalExact := "C:\Users\eduev\Documents\ClipAngel\ClipAngel 2.13\ClipAngel.exe"
    workExact :=
        "C:\Users\fie7ca\OneDrive - Bosch Group\01 - Geral\16 - Others backups\Documents\ClipAngel\ClipAngel.exe"
    workLegacyExact := "C:\Users\fie7ca\Documents\ClipAngel\ClipAngel 2.23\ClipAngel.exe"
    isWork := IsSet(IS_WORK_ENVIRONMENT) && IS_WORK_ENVIRONMENT
    preferred := isWork ? workExact : personalExact
    other := isWork ? personalExact : workExact
    if FileExist(preferred)
        return preferred
    if FileExist(workLegacyExact)
        return workLegacyExact
    if FileExist(other)
        return other

    roots := []
    if (isWork) {
        roots.Push("C:\Users\fie7ca\OneDrive - Bosch Group\01 - Geral\16 - Others backups\Documents")
        roots.Push("C:\Users\fie7ca\Documents")
        userDocs := EnvGet("USERPROFILE") "\Documents"
        if (userDocs != "" && userDocs != "C:\Users\fie7ca\Documents")
            roots.Push(userDocs)
        if (A_MyDocuments != "" && A_MyDocuments != "C:\Users\fie7ca\Documents")
            roots.Push(A_MyDocuments)
    } else {
        roots.Push("C:\Users\eduev\Documents")
        userDocs := EnvGet("USERPROFILE") "\Documents"
        if (userDocs != "" && userDocs != "C:\Users\eduev\Documents")
            roots.Push(userDocs)
        if (A_MyDocuments != "" && A_MyDocuments != "C:\Users\eduev\Documents")
            roots.Push(A_MyDocuments)
    }
    for root in roots {
        for cand in [
            root "\ClipAngel\ClipAngel.exe",
            root "\ClipAngel\ClipAngel 2.23\ClipAngel.exe",
            root "\ClipAngel\ClipAngel 2.13\ClipAngel.exe"
        ] {
            if FileExist(cand)
                return cand
        }
        loop files root "\ClipAngel\*\ClipAngel.exe", "F" {
            return A_LoopFileFullPath
        }
    }
    return ""
}

; Force-kill ClipAngel.exe (not minimize). Returns true when no process remains.
ClipAngel_KillProcess(timeoutMs := 4000) {
    deadline := A_TickCount + timeoutMs
    while (ProcessExist("ClipAngel.exe") && A_TickCount < deadline) {
        try ProcessClose("ClipAngel.exe")
        catch {
        }
        Sleep 80
    }
    return !ProcessExist("ClipAngel.exe")
}

; Macros [R] — hard restart when Clip Angel is stuck (kill process → relaunch → activate).
ClipAngel_RestartHard(*) {
    global g_ClipAngelAutomationBusy
    StandardLoadingBar_Show("⏳ Restarting Clip Angel…", BANNER_ACCENT_INTERMEDIATE, { passive: false })
    try {
        exePath := ClipAngel_ResolveExePath()
        g_ClipAngelAutomationBusy := false
        if ProcessExist("ClipAngel.exe") {
            StandardLoadingBar_Update("⏳ Killing Clip Angel process…")
            if !ClipAngel_KillProcess(5000) {
                StandardLoadingBar_Hide(0)
                ShowCenteredOverlay_Utils("❌ Could not kill ClipAngel.exe", 2500, BANNER_ACCENT_ERROR)
                return
            }
            Sleep 250
        }
        if (exePath = "" || !FileExist(exePath)) {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ ClipAngel.exe not found", 2500, BANNER_ACCENT_ERROR)
            return
        }
        StandardLoadingBar_Update("⏳ Opening Clip Angel…")
        try Run('"' exePath '"')
        catch as e {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ Failed to start Clip Angel: " . e.Message, 2800, BANNER_ACCENT_ERROR)
            return
        }
        hwnd := ClipAngel_WaitForMainHwnd(10000)
        if (!hwnd) {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ Clip Angel did not start", 2500, BANNER_ACCENT_ERROR)
            return
        }
        StandardLoadingBar_Hide(0)
        ActivateClipAngelWithFocusCorrection(false, 0, false, true)
        ShowCenteredOverlay_Utils("✅ Clip Angel restarted", 1500, BANNER_ACCENT_SUCCESS)
    } catch as e {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        ShowCenteredOverlay_Utils("❌ Clip Angel restart failed: " . e.Message, 2800, BANNER_ACCENT_ERROR)
    }
}

RegisterMacro(ClipAngel_RestartHard, "🔄 Restart Clip Angel (kill process)", "r")