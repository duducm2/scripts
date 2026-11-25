; =============================================================================
; Utils module: outlook_teams_check.ahk
; CheckAndOpenOutlookTeams prompt helper
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; Check and Prompt to Open Outlook/Teams
; Checks if Outlook or Teams are closed and prompts user to open them if needed
; Parameters:
;   - checkOutlook: true to check Outlook, false otherwise
;   - checkTeams: true to check Teams, false otherwise
;   - waitAfterLaunchMs: ms to sleep after launching (pass 0 to continue immediately in parallel)
; Returns: true if applications are running (or were opened), false if user cancelled
; =============================================================================
CheckAndOpenOutlookTeams(checkOutlook := false, checkTeams := false, waitAfterLaunchMs := 2000) {
    outlookClosed := false
    teamsClosed := false

    ; Check Outlook status (classic OUTLOOK.EXE or Store new Outlook olk.exe)
    if (checkOutlook) {
        outlookRunning := OutlookProcessRunning()
        if (!outlookRunning) {
            outlookClosed := true
        }
    }

    ; Check Teams status
    if (checkTeams) {
        teamsRunning := ProcessExist("ms-teams.exe") || ProcessExist("Teams.exe") || ProcessExist("MSTeams.exe")
        if (!teamsRunning) {
            teamsClosed := true
        }
    }

    ; If both are open, no action needed
    if (!outlookClosed && !teamsClosed) {
        return true
    }

    ; Only prompt for Outlook; Teams opens automatically without asking.
    response := "Yes"
    if (outlookClosed) {
        message := teamsClosed
            ? "Outlook and Teams are closed. Do you want to open them?"
            : "Outlook is closed. Do you want to open it?"
        response := MsgBox(message, "Open Applications?", "YesNo Icon?")
    }

    ; If user confirms, open the applications (only open, don't toggle)
    if (response = "Yes") {
        ; Only open the closed applications, don't toggle
        try {
            ; Launch Outlook if closed
            if (outlookClosed) {
                outlookPath := ""
                if (IS_WORK_ENVIRONMENT) {
                    outlookPath := "C:\Users\fie7ca\Documents\Atalhos\Microsoft Outlook.lnk"
                    if (!FileExist(outlookPath)) {
                        outlookPath := ""
                    }
                } else {
                    outlookPath :=
                        "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Microsoft Outlook.lnk"
                    if (!FileExist(outlookPath)) {
                        outlookPath := ""
                    }
                }

                if (outlookPath != "") {
                    Run outlookPath
                } else {
                    olkPath := OutlookGetOlkExePath()
                    if (olkPath != "")
                        Run olkPath
                    else
                        Run "OUTLOOK.EXE"
                }
            }

            ; Launch Teams if closed (Min when waitAfterLaunchMs=0 for parallel background startup)
            if (teamsClosed) {
                launchOpt := (waitAfterLaunchMs = 0) ? "Min" : ""
                if (IS_WORK_ENVIRONMENT) {
                    teamsExePath :=
                        "C:\Program Files\WindowsApps\MSTeams_25332.1210.4188.1171_x64__8wekyb3d8bbwe\ms-teams.exe"
                    if (FileExist(teamsExePath)) {
                        Run teamsExePath, , launchOpt
                    } else {
                        Run "ms-teams.exe", , launchOpt
                    }
                } else {
                    teamsPath :=
                        "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Microsoft Teams.lnk"
                    if (FileExist(teamsPath)) {
                        Run teamsPath, , launchOpt
                    } else {
                        Run "ms-teams.exe", , launchOpt
                    }
                }
            }

            ; Wait a bit for applications to start (skipped when caller passes 0 for parallel startup)
            if (waitAfterLaunchMs > 0)
                Sleep waitAfterLaunchMs
            return true
        } catch Error as e {
            MsgBox "Error opening applications: " e.Message, "Error", "IconX"
            return false
        }
    }

    ; User cancelled
    return false
}

; Chime for desktop recycle Y confirmation, and after clipboard cleanup succeeds.
; Quiet: WASAPI attenuation + synchronous SoundPlay - no WMPlayer.OCX (its volume pins same-PID mixer ~10% despite later WASAPI). try/finally restores SCRIPT_MASTER_VOLUME_PERCENT deterministically.
; One-shot timer re-applies target: a new session can appear right after SoundPlay returns; first enumeration may miss it (mixer stuck ~10%).
PlayCleaningDesktopSound() {
    if (!IsSoundEnabled())
        return
    soundPath := A_ScriptDir "\assets\sounds\cleaning-desktop.wav"
    if (!FileExist(soundPath))
        return
    ScriptSoundPlay(soundPath, false)
}

; Clean Clipboard macro: session/cancel guards (prevents double-run and ignored N during automation)
global g_CleanClipboardCanceled := false
global g_CleanClipboardInProgress := false
global g_CleanClipboardSessionId := 0
global g_CleanClipboardProceedClaimed := false

CleanClipboard_BeginSession() {
    global g_CleanClipboardCanceled, g_CleanClipboardSessionId, g_CleanClipboardProceedClaimed
    g_CleanClipboardSessionId += 1
    g_CleanClipboardCanceled := false
    g_CleanClipboardProceedClaimed := false
    return g_CleanClipboardSessionId
}

CleanClipboard_EndSession() {
    global g_CleanClipboardInProgress
    g_CleanClipboardInProgress := false
}

CleanClipboard_ShouldAbort(sessionId := 0) {
    global g_CleanClipboardCanceled, g_CleanClipboardSessionId
    if (g_CleanClipboardCanceled)
        return true
    if (sessionId && sessionId != g_CleanClipboardSessionId)
        return true
    return false
}

CleanClipboard_UnwindClipAngel(hwnd := 0) {
    global g_ClipAngelFavoriteSuppressActive
    if (!hwnd)
        hwnd := ClipAngel_MainHwnd()
    ClipAngel_DismissOpenMenus(hwnd)
    if (g_ClipAngelFavoriteSuppressActive) {
        ClipAngel_EndFavoriteSuppress(hwnd)
        return
    }
    if (!hwnd)
        return
    try {
        if ClipAngel_IsWindowShown(hwnd)
            WinMinimize("ahk_id " hwnd)
    } catch {
    }
    ; Clear after minimize so layered Off does not flash an opaque maximized window.
    ClipAngel_ClearFavoriteSessionOpacity(hwnd)
}

; Shift+N / Esc while automation runs (do NOT hook bare *n so typing words with 'n' in another window never aborts)
CleanClipboard_SetAbortHotkeys(enable := true) {
    try HotIf()
    catch {
    }
    if (enable) {
        Hotkey("+n", CleanClipboard_OnCancel, "On")
        Hotkey("*Escape", CleanClipboard_OnCancel, "On")
    } else {
        try Hotkey("+n", "Off")
        catch {
        }
        try Hotkey("*Escape", "Off")
        catch {
        }
    }
}

CleanClipboard_LogError(attempt, stage, detail) {
    try {
        logDir := A_ScriptDir "\assets\data"
        if !DirExist(logDir)
            DirCreate(logDir)
        ts := FormatTime(, "yyyy-MM-dd HH:mm:ss")
        line := ts " | attempt=" attempt " | stage=" stage " | " detail "`n"
        FileAppend(line, logDir "\cleanclipboard_errors.log", "UTF-8")
    } catch {
    }
}

; Open Clip Angel suppressed (off-screen, opacity 0) for background cleanup. Returns hwnd or 0.
CleanClipboard_ActivateClipAngelForCleanup(sessionId := 0, &outRoot := 0) {
    outRoot := 0
    if (CleanClipboard_ShouldAbort(sessionId))
        return 0

    hwnd := 0
    if (!ClipAngel_OpenForAutomation("all", 0, true, &hwnd, &outRoot, true) || !hwnd)
        return 0
    if (CleanClipboard_ShouldAbort(sessionId))
        return 0
    return hwnd
}

CleanClipboard_UpdateProgress(pct, msg, barColor := BANNER_ACCENT_INTERMEDIATE) {
    try StandardLoadingBar_SetProgressValue(pct)
    catch {
    }
    try StandardLoadingBar_Update(msg, barColor)
    catch {
    }
}

CleanClipboard_FinishStatus(pct, msg, barColor, hideMs, playSuccessSound := false) {
    CleanClipboard_UpdateProgress(pct, msg, barColor)
    if (playSuccessSound) {
        try PlayCleaningDesktopSound()
        catch {
        }
    }
    try StandardLoadingBar_Hide(hideMs)
    catch {
    }
}

global g_CleanClipboardActiveRun := 0

; Kept as no-ops for backwards compatibility with any external callers.
CleanClipboard_CancelPrewarm() {
}

CleanClipboard_BeginPrewarm(sessionId := 0) {
}

CleanClipboard_CancelActiveRun() {
    global g_CleanClipboardActiveRun
    run := g_CleanClipboardActiveRun
    if (!run || run.finished)
        return
    run.cancelled := true
    if (run.osHandle && IsObject(run.osHandle)) {
        run.osHandle.cancelled := true
        if (run.osHandle.timerFn) {
            try SetTimer(run.osHandle.timerFn, 0)
            catch {
            }
        }
        run.osHandle := 0
    }
    CleanClipboard_UnwindClipAngel()
    CleanClipboard_CompleteRun(run)
}

CleanClipboard_CompleteRun(run) {
    global g_ClipAngelAutomationBusy, g_CleanClipboardActiveRun, g_CleanClipboardInProgress
    if (!run || run.finished)
        return
    run.finished := true
    if (g_CleanClipboardActiveRun = run)
        g_CleanClipboardActiveRun := 0
    g_ClipAngelAutomationBusy := run.prevBusy
    if (run.ownsSession)
        g_CleanClipboardInProgress := false
    try CleanClipboard_SetAbortHotkeys(false)
    catch {
    }
    try CleanClipboard_EndSession()
    catch {
    }
}

; Worker A: Zero-focus-steal background Clip Angel cleanup (Handy/BeginFavoriteSuppress pattern).
; Parks Clip Angel suppressed (WS_EX_NOACTIVATE + opacity 0), resets to all clips (F9),
; then sends Ctrl+A -> Delete directly to the dataGridView control via ControlSend.
; User can keep typing or switching windows without interruption.
CleanClipboard_RunClipAngelAltPDelete(run, &outStage := "", &outDetail := "") {
    global g_ClipAngelAutomationBusy
    outStage := ""
    outDetail := ""
    prevBusy := g_ClipAngelAutomationBusy
    g_ClipAngelAutomationBusy := true
    hwnd := 0
    ok := false
    prevDetect := A_DetectHiddenWindows
    DetectHiddenWindows true
    try {
        ; Dismiss any stray #32770 dialog in the background without stealing focus.
        try ClipAngel_TryConfirmDeleteDialogOnce(false)
        catch {
        }

        if (run.cancelled || CleanClipboard_ShouldAbort(run.sessionId))
            return false

        CleanClipboard_UpdateProgress(45, "⏳ " run.attemptTag "[1/2] Cleaning Clip Angel in background...")

        root := 0
        hwnd := CleanClipboard_ActivateClipAngelForCleanup(run.sessionId, &root)
        if (!hwnd) {
            if (run.cancelled || CleanClipboard_ShouldAbort(run.sessionId))
                return false
            outStage := "1/2 Open"
            outDetail := "Clip Angel is not running or window not found"
            return false
        }

        if (run.cancelled || CleanClipboard_ShouldAbort(run.sessionId))
            return false

        ; Reset filter & show all clips (normal + favorite) in background without activating.
        try ControlSend("{F9}", , "ahk_id " hwnd)
        catch {
        }
        Sleep 40

        if (run.cancelled || CleanClipboard_ShouldAbort(run.sessionId))
            return false

        ; Target dataGridView HWND directly so Ctrl+A -> Delete never hits the search/filter box.
        targetHwnd := hwnd
        try {
            dataGrid := ClipAngel_UiaGetDataGrid(hwnd, root)
            if (dataGrid) {
                gridHwnd := 0
                try gridHwnd := dataGrid.NativeWindowHandle
                catch {
                }
                if (!gridHwnd) {
                    try gridHwnd := dataGrid.GetPropertyValue(30020) ; UIA_NativeWindowHandlePropertyId
                    catch {
                    }
                }
                if (gridHwnd)
                    targetHwnd := gridHwnd
            }
        } catch {
        }

        CleanClipboard_UpdateProgress(75, "⏳ " run.attemptTag "[1/2] Selecting all & deleting non-favorites...")
        try {
            ControlSend("^a", , "ahk_id " targetHwnd)
            Sleep 40
            if (run.cancelled || CleanClipboard_ShouldAbort(run.sessionId))
                return false
            ControlSend("{Delete}", , "ahk_id " targetHwnd)
            Sleep 50
        } catch Error as e {
            outStage := "1/2 Delete"
            outDetail := "Background ControlSend failed: " e.Message
            return false
        }

        ; Defensive: dismiss any confirmation dialog in background if ConfirmationBeforeDelete is ever enabled.
        try ClipAngel_TryConfirmDeleteDialogOnce(false)
        catch {
        }

        ok := true
    } finally {
        DetectHiddenWindows prevDetect
        CleanClipboard_UnwindClipAngel(hwnd)
        ; Only restore focus if Clip Angel itself accidentally stole it — never yank the user away from another window.
        if (run.priorHwnd && WinActive("ahk_exe ClipAngel.exe"))
            ClipAngel_RestorePriorFocus(run.priorHwnd, true)
        g_ClipAngelAutomationBusy := prevBusy
    }
    return ok
}

CleanClipboard_FinalizeAttempt(run, clipAngelOk) {
    if (run.finished || run.cancelled || CleanClipboard_ShouldAbort(run.sessionId)) {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        CleanClipboard_CompleteRun(run)
        return
    }

    ; Final Guarantee Gate: re-verify OS clipboard is empty after Clip Angel delete/minimize.
    CleanClipboard_UpdateProgress(95, "⏳ " run.attemptTag "[2/2] Verifying clipboard empty...")
    osClipCleared := run.osOk
    if (!osClipCleared || !CleanClipboard_IsOsClipboardEmpty())
        osClipCleared := CleanClipboard_ClearOsClipboard(4)

    run.lastClipAngelOk := clipAngelOk
    if (clipAngelOk && osClipCleared) {
        retrySuffix := (run.attempt > 1) ? (" [attempt " run.attempt "/" run.maxAttempts "]") : ""
        CleanClipboard_FinishStatus(100, "✅ Clipboard cleaned" retrySuffix,
            BANNER_ACCENT_SUCCESS, 1200, true)
        CleanClipboard_CompleteRun(run)
        return
    }

    if (clipAngelOk && !osClipCleared) {
        run.lastStage := "2/2 OS Clipboard"
        run.lastDetail := "Windows clipboard locked by another app"
    } else {
        run.lastStage := (run.attemptStage != "") ? run.attemptStage : "1/2 Clip Angel"
        run.lastDetail := (run.attemptDetail != "") ? run.attemptDetail : "Unknown failure"
    }
    CleanClipboard_LogError(run.attempt, run.lastStage, run.lastDetail)

    if (run.attempt < run.maxAttempts && !run.cancelled && !CleanClipboard_ShouldAbort(run.sessionId)) {
        run.attempt += 1
        CleanClipboard_UpdateProgress(20,
            "⚠️ Attempt " (run.attempt - 1) "/" run.maxAttempts " failed [" run.lastStage "]: " run.lastDetail
            . " — retrying (" run.attempt "/" run.maxAttempts ")...",
            BANNER_ACCENT_INTERMEDIATE)
        SetTimer((*) => CleanClipboard_RunAttempt(run), -100)
        return
    }

    if (run.lastClipAngelOk) {
        CleanClipboard_FinishStatus(100,
            "⚠️ Clip Angel cleaned, but Windows clipboard stayed locked (" run.maxAttempts "/" run.maxAttempts ")",
            BANNER_ACCENT_INTERMEDIATE, 3000, false)
    } else {
        CleanClipboard_FinishStatus(100,
            "❌ Failed (" run.maxAttempts "/" run.maxAttempts ") at [" run.lastStage "]: " run.lastDetail,
            BANNER_ACCENT_ERROR, 3500, false)
    }
    CleanClipboard_CompleteRun(run)
}

CleanClipboard_CheckBarrier(run) {
    if (run.finished || run.cancelled || CleanClipboard_ShouldAbort(run.sessionId)) {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        CleanClipboard_CompleteRun(run)
        return
    }
    if (!run.uiDone || !run.osDone)
        return

    CleanClipboard_FinalizeAttempt(run, run.uiOk)
}

CleanClipboard_RunAttempt(run) {
    if (run.finished || run.cancelled || CleanClipboard_ShouldAbort(run.sessionId)) {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        CleanClipboard_CompleteRun(run)
        return
    }

    run.attemptTag := (run.attempt > 1) ? ("[Attempt " run.attempt "/" run.maxAttempts "] ") : ""
    run.uiDone := false
    run.osDone := false
    run.uiOk := false
    run.osOk := false
    run.attemptStage := ""
    run.attemptDetail := ""
    run.osHandle := 0

    CleanClipboard_UpdateProgress(25, "⏳ " run.attemptTag "[1/2] Cleaning clipboard & Clip Angel in parallel...")

    ; Worker B: async OS clipboard clear in parallel with Worker A.
    onOsDone(ok) {
        if (run.finished || run.cancelled)
            return
        run.osHandle := 0
        run.osDone := true
        run.osOk := ok
        CleanClipboard_CheckBarrier(run)
    }

    run.osHandle := CleanClipboard_StartClearOsClipboardAsync(6, onOsDone,
        (*) => (run.cancelled || CleanClipboard_ShouldAbort(run.sessionId)))

    ; Worker A: Clip Angel native Alt+P -> Ctrl+A -> Delete.
    uiStage := "", uiDetail := ""
    uiOk := CleanClipboard_RunClipAngelAltPDelete(run, &uiStage, &uiDetail)
    run.uiDone := true
    run.uiOk := uiOk
    if (!uiOk) {
        run.attemptStage := uiStage
        run.attemptDetail := uiDetail
    }
    CleanClipboard_CheckBarrier(run)
}

; Internal helper: Performs parallel clipboard cleanup (Clip Angel Alt+P -> Ctrl+A -> Delete
; + concurrent Windows OS clipboard clear) with retry and final guarantee barrier.
; Returns true if an async run is active (caller must NOT end session synchronously).
CleanClipboardInternal(sessionId := 0) {
    global g_ClipAngelAutomationBusy, g_CleanClipboardInProgress, g_CleanClipboardActiveRun
    if (CleanClipboard_ShouldAbort(sessionId)) {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        CleanClipboard_UnwindClipAngel()
        return false
    }

    ownsSession := (sessionId = 0)
    if (ownsSession) {
        if (g_CleanClipboardInProgress)
            return false
        g_CleanClipboardInProgress := true
    }

    priorHwnd := ClipAngel_ResolvePriorHwnd(0)
    StandardLoadingBar_Show("⏳ [1/2] Cleaning clipboard & Clip Angel in parallel...", BANNER_ACCENT_INTERMEDIATE, {
        passive: false,
        manualProgress: true,
        noActivate: true,
        centerOnHwnd: priorHwnd,
        fontSize: 17
    })
    try StandardLoadingBar_SetProgressValue(15)
    catch {
    }

    run := {
        sessionId: sessionId,
        ownsSession: ownsSession,
        priorHwnd: priorHwnd,
        prevBusy: g_ClipAngelAutomationBusy,
        attempt: 1,
        maxAttempts: 2,
        attemptTag: "",
        osHandle: 0,
        uiDone: false,
        osDone: false,
        uiOk: false,
        osOk: false,
        attemptStage: "",
        attemptDetail: "",
        lastStage: "",
        lastDetail: "",
        lastClipAngelOk: false,
        cancelled: false,
        finished: false
    }
    g_CleanClipboardActiveRun := run
    try {
        CleanClipboard_RunAttempt(run)
    } catch Error as e {
        try CleanClipboard_LogError(run.attempt, "RunAttempt", e.Message)
        catch {
        }
        try CleanClipboard_FinishStatus(100, "❌ Clipboard cleanup error: " e.Message, BANNER_ACCENT_ERROR, 3500, false)
        catch {
        }
        CleanClipboard_CompleteRun(run)
        return false
    }
    return !run.finished
}
