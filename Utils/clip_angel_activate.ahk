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

; SendMessageTimeout WM_NULL + IsHungAppWindow; false when unresponsive within timeout.
CLIPANGEL_RESPONSIVE_PROBE_MS := 300
CLIPANGEL_RESTART_HWND_WAIT_MS := 15000
CLIPANGEL_RESTART_RESPONSIVE_WAIT_MS := 20000
CLIPANGEL_RESTART_POST_ACTIVATE_RESPONSIVE_MS := 10000
CLIPANGEL_RESTART_PROCESS_WAIT_MS := 12000
CLIPANGEL_RESTART_LIST_WAIT_MS := 8000
CLIPANGEL_RESTART_QUALITY_POLL_MS := 100
CLIPANGEL_RESTART_MAX_ATTEMPTS := 2

ClipAngel_IsWindowResponsive(hwnd, timeoutMs := 0) {
    if !(hwnd is Integer) || hwnd <= 0 || !WinExist("ahk_id " hwnd)
        return false
    if (!timeoutMs)
        timeoutMs := CLIPANGEL_RESPONSIVE_PROBE_MS
    try {
        if DllCall("IsHungAppWindow", "Ptr", hwnd)
            return false
    } catch {
    }
    result := 0
    ; SMTO_ABORTIFHUNG = 0x0002
    ok := DllCall("SendMessageTimeout", "Ptr", hwnd, "UInt", 0, "Ptr", 0, "Ptr", 0, "UInt", 0x0002, "UInt",
        timeoutMs, "Ptr*", &result)
    return ok != 0
}

; Poll until Clip Angel's main window answers WM_NULL (or timeout).
ClipAngel_WaitUntilResponsive(hwnd, timeoutMs := 0) {
    if (!timeoutMs)
        timeoutMs := CLIPANGEL_RESTART_RESPONSIVE_WAIT_MS
    if !(hwnd is Integer) || hwnd <= 0
        return false
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        cur := ClipAngel_MainHwnd()
        if (cur)
            hwnd := cur
        if ClipAngel_IsWindowResponsive(hwnd)
            return true
        Sleep CLIPANGEL_RESTART_QUALITY_POLL_MS
    }
    cur := ClipAngel_MainHwnd()
    if (cur)
        hwnd := cur
    return ClipAngel_IsWindowResponsive(hwnd)
}

; True when hwnd is the ClipAngel WinForms main shell (not a random dialog).
ClipAngel_IsMainWindowIdentity(hwnd) {
    if !(hwnd is Integer) || hwnd <= 0 || !WinExist("ahk_id " hwnd)
        return false
    try {
        if (StrLower(WinGetProcessName("ahk_id " hwnd)) != "clipangel.exe")
            return false
    } catch {
        return false
    }
    try {
        title := WinGetTitle("ahk_id " hwnd)
        if !InStr(title, "ClipAngel")
            return false
    } catch {
        return false
    }
    try {
        cls := WinGetClass("ahk_id " hwnd)
        if !InStr(cls, "WindowsForms10")
            return false
    } catch {
        return false
    }
    return true
}

; New process after kill: ClipAngel.exe exists, PID != oldPid, optional path match.
ClipAngel_WaitForNewProcess(oldPid, expectedExe := "", timeoutMs := 0) {
    if (!timeoutMs)
        timeoutMs := CLIPANGEL_RESTART_PROCESS_WAIT_MS
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        pid := ProcessExist("ClipAngel.exe")
        if (pid && pid != oldPid) {
            if (expectedExe = "")
                return pid
            try path := ProcessGetPath("ClipAngel.exe")
            catch
                path := ""
            if (path = "" || StrLower(path) = StrLower(expectedExe))
                return pid
        }
        Sleep CLIPANGEL_RESTART_QUALITY_POLL_MS
    }
    return 0
}

; Wait for a main hwnd owned by newPid (when known) that passes identity checks.
ClipAngel_WaitForRestartHwnd(newPid := 0, timeoutMs := 0) {
    if (!timeoutMs)
        timeoutMs := CLIPANGEL_RESTART_HWND_WAIT_MS
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        hwnd := ClipAngel_MainHwnd()
        if (hwnd && ClipAngel_IsMainWindowIdentity(hwnd)) {
            if (!newPid)
                return hwnd
            try {
                if (WinGetPID("ahk_id " hwnd) = newPid)
                    return hwnd
            } catch {
            }
        }
        Sleep CLIPANGEL_RESTART_QUALITY_POLL_MS
    }
    hwnd := ClipAngel_MainHwnd()
    if (hwnd && ClipAngel_IsMainWindowIdentity(hwnd)) {
        if (!newPid)
            return hwnd
        try {
            if (WinGetPID("ahk_id " hwnd) = newPid)
                return hwnd
        } catch {
        }
    }
    return 0
}

; Layout usable after restart: shown, not tiny bar, preferably maximized + foreground.
ClipAngel_IsRestartLayoutOk(hwnd) {
    if !hwnd || !ClipAngel_IsWindowShown(hwnd)
        return false
    if ClipAngel_NeedsLayoutCorrection(hwnd)
        return false
    if !WinActive("ahk_id " hwnd)
        return false
    try {
        if (WinGetMinMax("ahk_id " hwnd) != 1)
            return false
    } catch {
        return false
    }
    return true
}

; UIA gates: dataGridView present, then Row 0 / list ready. Never call if hung.
ClipAngel_WaitForRestartUiReady(hwnd, timeoutMs := 0) {
    if (!timeoutMs)
        timeoutMs := CLIPANGEL_RESTART_LIST_WAIT_MS
    if !(hwnd is Integer) || hwnd <= 0
        return false
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        cur := ClipAngel_MainHwnd()
        if (cur)
            hwnd := cur
        if !ClipAngel_IsWindowResponsive(hwnd)
            return false
        try {
            if ClipAngel_UiaGetDataGrid(hwnd) {
                if ClipAngel_IsListReady(&hwnd) {
                    try ClipAngel_UiaEnsureRow0Selected(hwnd, false)
                    catch {
                    }
                    return true
                }
            }
        } catch {
        }
        Sleep CLIPANGEL_UIA_POLL_MS
    }
    if !ClipAngel_IsWindowResponsive(hwnd)
        return false
    try return ClipAngel_IsListReady()
    catch
        return false
}

; Full post-restart quality gate. Sets failReason on false.
ClipAngel_RestartQualityOk(hwnd, expectedExe, oldPid, &failReason) {
    failReason := ""
    if !(hwnd is Integer) || hwnd <= 0 || !WinExist("ahk_id " hwnd) {
        failReason := "no window"
        return false
    }
    if !ClipAngel_IsMainWindowIdentity(hwnd) {
        failReason := "window identity mismatch"
        return false
    }
    try pid := WinGetPID("ahk_id " hwnd)
    catch
        pid := 0
    if (oldPid && pid && pid = oldPid) {
        failReason := "same PID as before kill"
        return false
    }
    if (expectedExe != "") {
        try path := WinGetProcessPath("ahk_id " hwnd)
        catch
            path := ""
        if (path != "" && StrLower(path) != StrLower(expectedExe)) {
            failReason := "exe path mismatch"
            return false
        }
    }
    if !ClipAngel_IsWindowResponsive(hwnd) {
        failReason := "not responding"
        return false
    }
    if !ClipAngel_IsRestartLayoutOk(hwnd) {
        failReason := "layout/foreground not ready"
        return false
    }
    try {
        if !ClipAngel_UiaGetDataGrid(hwnd) {
            failReason := "clip list (dataGridView) missing"
            return false
        }
        if !ClipAngel_IsListReady(&hwnd) {
            failReason := "clip list Row 0 not ready"
            return false
        }
    } catch {
        failReason := "UIA list check failed"
        return false
    }
    return true
}

; One kill→launch→verify cycle. Returns true on success; failReason on false.
ClipAngel_RestartAttempt(exePath, oldPid, attempt, &failReason) {
    failReason := ""
    baselinePid := oldPid
    label := (CLIPANGEL_RESTART_MAX_ATTEMPTS > 1)
        ? " (attempt " attempt "/" CLIPANGEL_RESTART_MAX_ATTEMPTS ")"
        : ""

    if ProcessExist("ClipAngel.exe") {
        StandardLoadingBar_Update("⏳ Killing Clip Angel process…" label)
        if !ClipAngel_KillProcess(5000) {
            failReason := "could not kill ClipAngel.exe"
            return false
        }
        Sleep 250
    }

    StandardLoadingBar_Update("⏳ Opening Clip Angel…" label)
    try Run('"' exePath '"')
    catch as e {
        failReason := "failed to start: " e.Message
        return false
    }

    StandardLoadingBar_Update("⏳ Waiting for new Clip Angel process…" label)
    newPid := ClipAngel_WaitForNewProcess(baselinePid, exePath, CLIPANGEL_RESTART_PROCESS_WAIT_MS)
    if (!newPid) {
        failReason := "process did not start (or wrong exe path)"
        return false
    }

    StandardLoadingBar_Update("⏳ Waiting for Clip Angel window…" label)
    hwnd := ClipAngel_WaitForRestartHwnd(newPid, CLIPANGEL_RESTART_HWND_WAIT_MS)
    if (!hwnd) {
        failReason := "main window did not appear"
        return false
    }

    StandardLoadingBar_Update("⏳ Waiting for Clip Angel to respond…" label)
    if !ClipAngel_WaitUntilResponsive(hwnd, CLIPANGEL_RESTART_RESPONSIVE_WAIT_MS) {
        failReason := "started but is not responding"
        return false
    }

    StandardLoadingBar_Update("⏳ Activating Clip Angel…" label)
    ; silent + skipRow0: keep StandardLoadingBar as the only indicator; UIA list gate comes next.
    if !ActivateClipAngelWithFocusCorrection(true, 0, true, true) {
        failReason := "window not found after start"
        return false
    }

    StandardLoadingBar_Update("⏳ Verifying Clip Angel is responsive…" label)
    hwnd := ClipAngel_MainHwnd()
    if !hwnd || !ClipAngel_WaitUntilResponsive(hwnd, CLIPANGEL_RESTART_POST_ACTIVATE_RESPONSIVE_MS) {
        failReason := "not responding after activate"
        return false
    }

    StandardLoadingBar_Update("⏳ Waiting for Clip Angel list…" label)
    if !ClipAngel_WaitForRestartUiReady(hwnd, CLIPANGEL_RESTART_LIST_WAIT_MS) {
        failReason := "clip list not ready"
        return false
    }

    StandardLoadingBar_Update("⏳ Confirming restart quality…" label)
    hwnd := ClipAngel_MainHwnd()
    if !ClipAngel_RestartQualityOk(hwnd, exePath, baselinePid, &failReason) {
        if (failReason = "")
            failReason := "quality gate failed"
        return false
    }
    return true
}

; Macros [R] — hard restart when Clip Angel is stuck.
; Loading Indication stays up for the whole path; up to 2 attempts with multi-gate verify.
ClipAngel_RestartHard(*) {
    global g_ClipAngelAutomationBusy
    StandardLoadingBar_Show("⏳ Restarting Clip Angel…", BANNER_ACCENT_INTERMEDIATE, { passive: false })
    g_ClipAngelAutomationBusy := true
    failReason := ""
    try {
        oldPid := ProcessExist("ClipAngel.exe")
        exePath := ClipAngel_ResolveExePath()
        ; Prefer path from still-running process before kill when resolver is empty.
        if ((exePath = "" || !FileExist(exePath)) && oldPid) {
            try {
                hwndLive := ClipAngel_MainHwnd()
                if (hwndLive)
                    exePath := WinGetProcessPath("ahk_id " hwndLive)
            } catch {
            }
        }
        if (exePath = "" || !FileExist(exePath)) {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ ClipAngel.exe not found", 2500, BANNER_ACCENT_ERROR)
            return
        }

        loop CLIPANGEL_RESTART_MAX_ATTEMPTS {
            if ClipAngel_RestartAttempt(exePath, oldPid, A_Index, &failReason) {
                StandardLoadingBar_Hide(0)
                ShowCenteredOverlay_Utils("✅ Clip Angel restarted", 1500, BANNER_ACCENT_SUCCESS)
                return
            }
            if (A_Index < CLIPANGEL_RESTART_MAX_ATTEMPTS) {
                StandardLoadingBar_Update("⏳ Restart incomplete — retrying…")
                Sleep 400
                oldPid := ProcessExist("ClipAngel.exe")
            }
        }

        StandardLoadingBar_Hide(0)
        msg := "❌ Clip Angel restart failed"
        if (failReason != "")
            msg .= ": " failReason
        ShowCenteredOverlay_Utils(msg, 3200, BANNER_ACCENT_ERROR)
    } catch as e {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        ShowCenteredOverlay_Utils("❌ Clip Angel restart failed: " . e.Message, 2800, BANNER_ACCENT_ERROR)
    } finally {
        g_ClipAngelAutomationBusy := false
    }
}

RegisterMacro(ClipAngel_RestartHard, "🔄 Restart Clip Angel (kill process)", "r")