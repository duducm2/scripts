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
        teamsRunning := ProcessExist("ms-teams.exe")
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
global g_CleanClipboardPrewarm := 0

CleanClipboard_CancelPrewarm() {
    global g_CleanClipboardPrewarm
    pw := g_CleanClipboardPrewarm
    g_CleanClipboardPrewarm := 0
    if (!pw || pw.completed)
        return
    pw.completed := true
    pw.onCompleteCallback := 0
    if (pw.handle) {
        try ClipAngelDb_CancelAsync(pw.handle)
        catch {
        }
        pw.handle := 0
    }
}

; Starts ClipAngelDb.exe deletenonfav-staged immediately when the 3s countdown opens
; so .NET CLR + SQLite key derivation + page cache warm up while the user decides.
; No write lock is held until "GO\n" is sent at timeout; Shift+N / Esc aborts cleanly.
CleanClipboard_BeginPrewarm(sessionId) {
    global g_CleanClipboardPrewarm
    CleanClipboard_CancelPrewarm()
    try ClipAngel_TryConfirmDeleteDialogOnce(false)
    catch {
    }
    if (ClipAngelDb_ExePath() = "")
        return

    pw := {
        sessionId: sessionId,
        handle: 0,
        completed: false,
        ok: false,
        outText: "",
        exitCode: 1,
        timedOut: false,
        onCompleteCallback: 0
    }

    onPrewarmDone(ok, outText, exitCode, timedOut) {
        pw.completed := true
        pw.handle := 0
        pw.ok := ok
        pw.outText := outText
        pw.exitCode := exitCode
        pw.timedOut := timedOut
        cb := pw.onCompleteCallback
        if (IsObject(cb) && HasMethod(cb, "Call")) {
            pw.onCompleteCallback := 0
            try cb(ok, outText, exitCode, timedOut)
            catch {
            }
        }
    }

    pw.handle := ClipAngelDb_StartAsync("deletenonfav-staged", 20000, onPrewarmDone, true, true)
    if (pw.handle)
        g_CleanClipboardPrewarm := pw
}

CleanClipboard_CancelActiveRun() {
    global g_CleanClipboardActiveRun
    run := g_CleanClipboardActiveRun
    if (!run || run.finished)
        return
    run.cancelled := true
    if (run.dbHandle) {
        try ClipAngelDb_CancelAsync(run.dbHandle)
        catch {
        }
        run.dbHandle := 0
    }
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

CleanClipboard_RunUiFallback(run, &outStage := "", &outDetail := "") {
    global g_ClipAngelAutomationBusy
    outStage := ""
    outDetail := ""
    prevBusy := g_ClipAngelAutomationBusy
    g_ClipAngelAutomationBusy := true
    hwnd := 0
    confirmed := false
    try {
        CleanClipboard_UpdateProgress(60, "⏳ " run.attemptTag "[2/3] DB busy — opening Clip Angel...")
        try ClipAngel_WaitChordModifiersReleased()
        catch {
        }
        try ClipAngel_ReleaseChordModifiersForSend()
        catch {
        }

        root := 0
        hwnd := CleanClipboard_ActivateClipAngelForCleanup(run.sessionId, &root)
        if (!hwnd) {
            if (run.cancelled || CleanClipboard_ShouldAbort(run.sessionId))
                return false
            outStage := "2/3 Open"
            outDetail := "Clip Angel window not found"
            return false
        }
        if (run.cancelled || CleanClipboard_ShouldAbort(run.sessionId))
            return false

        CleanClipboard_UpdateProgress(75, "⏳ " run.attemptTag "[2/3] Sending delete non-favorites...")
        confirmed := ClipAngel_TryConfirmDeleteDialogOnce(true)
        if (!confirmed) {
            try PostMessage(0x001F, 0, 0, , "ahk_id " hwnd) ; WM_CANCELMODE
            try ClipAngel_FocusGridIfFilterCombo()
            catch {
            }
            try {
                ControlSend("^!k", , "ahk_id " hwnd)
                confirmed := ClipAngel_ConfirmDeleteNonFavoritesDialogBackground(900)
            } catch {
                confirmed := false
            }
        }
        if (!confirmed && !run.cancelled && !CleanClipboard_ShouldAbort(run.sessionId)) {
            CleanClipboard_UpdateProgress(85, "⏳ " run.attemptTag "[2/3] Trying List menu fallback...")
            if (!root) {
                try root := UIA.ElementFromHandle(hwnd)
                catch {
                    root := 0
                }
            }
            menuDialogConfirmed := false
            menuPollCb := (*) => (
                !menuDialogConfirmed && ClipAngel_TryConfirmDeleteDialogOnce(true)
                    ? (menuDialogConfirmed := true) : 0
            )
            SetTimer(menuPollCb, 50)
            menuInvoked := false
            try {
                menuInvoked := ClipAngel_InvokeDeleteNonFavoritesViaMenu(hwnd, root)
            } finally {
                SetTimer(menuPollCb, 0)
            }
            if (menuDialogConfirmed)
                confirmed := true
            else if (menuInvoked)
                confirmed := ClipAngel_ConfirmDeleteNonFavoritesDialogBackground(1500)
            else
                confirmed := ClipAngel_TryConfirmDeleteDialogOnce(true)
        }
        if (!confirmed) {
            dbPrefix := (run.attemptDetail != "") ? (run.attemptDetail " -> ") : ""
            outStage := "2/3 UI"
            outDetail := dbPrefix "Delete confirmation dialog did not appear"
        }
    } finally {
        CleanClipboard_UnwindClipAngel(hwnd)
        if (run.priorHwnd)
            ClipAngel_RestorePriorFocus(run.priorHwnd, true)
        g_ClipAngelAutomationBusy := prevBusy
    }
    return confirmed
}

CleanClipboard_FinalizeAttempt(run, clipAngelOk, usedDb) {
    if (run.finished || run.cancelled || CleanClipboard_ShouldAbort(run.sessionId)) {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        CleanClipboard_CompleteRun(run)
        return
    }

    ; Final Guarantee Gate: re-verify OS clipboard is empty after ClipAngel reload/UI actions.
    CleanClipboard_UpdateProgress(95, "⏳ " run.attemptTag "[3/3] Verifying clipboard empty...")
    osClipCleared := run.osOk
    if (!osClipCleared || !CleanClipboard_IsOsClipboardEmpty())
        osClipCleared := CleanClipboard_ClearOsClipboard(4)

    run.lastClipAngelOk := clipAngelOk
    if (clipAngelOk && osClipCleared) {
        retrySuffix := (run.attempt > 1) ? (" [attempt " run.attempt "/" run.maxAttempts "]") : ""
        if (usedDb) {
            detail := (run.deletedCount = 0) ? "already empty"
                : (run.deletedCount = 1 ? "1 clip removed" : run.deletedCount " clips removed")
            CleanClipboard_FinishStatus(100, "✅ Clipboard cleaned (" detail ")" retrySuffix,
                BANNER_ACCENT_SUCCESS, 1200, true)
        } else {
            CleanClipboard_FinishStatus(100, "✅ Clipboard cleaned" retrySuffix,
                BANNER_ACCENT_SUCCESS, 1200, true)
        }
        CleanClipboard_CompleteRun(run)
        return
    }

    if (clipAngelOk && !osClipCleared) {
        run.lastStage := "3/3 OS Clipboard"
        run.lastDetail := "Windows clipboard locked by another app"
    } else {
        run.lastStage := (run.attemptStage != "") ? run.attemptStage : "2/3 Cleanup"
        run.lastDetail := (run.attemptDetail != "") ? run.attemptDetail : "Unknown failure"
    }
    CleanClipboard_LogError(run.attempt, run.lastStage, run.lastDetail)

    if (run.attempt < run.maxAttempts && !run.cancelled && !CleanClipboard_ShouldAbort(run.sessionId)) {
        run.attempt += 1
        CleanClipboard_UpdateProgress(20,
            "⚠️ Attempt " (run.attempt - 1) "/" run.maxAttempts " failed [" run.lastStage "]: " run.lastDetail
            . " — retrying (" run.attempt "/" run.maxAttempts ")...",
            BANNER_ACCENT_INTERMEDIATE)
        SetTimer((*) => CleanClipboard_RunAttempt(run), -(200 * (run.attempt - 1)))
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
    if (!run.dbDone || !run.osDone)
        return

    if (run.dbOk) {
        ClipAngel_TriggerBackgroundReloadList()
        CleanClipboard_FinalizeAttempt(run, true, true)
        return
    }

    ; Retry fast dialog-free DB path before falling back to UI (^!k / modal confirmation).
    if (ClipAngelDb_ExePath() != "" && run.attempt < run.maxAttempts) {
        CleanClipboard_FinalizeAttempt(run, false, true)
        return
    }

    uiStage := "", uiDetail := ""
    uiOk := CleanClipboard_RunUiFallback(run, &uiStage, &uiDetail)
    if (!uiOk) {
        if (uiStage != "")
            run.attemptStage := uiStage
        if (uiDetail != "")
            run.attemptDetail := uiDetail
    }
    CleanClipboard_FinalizeAttempt(run, uiOk, false)
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
    run.dbDone := false
    run.osDone := false
    run.dbOk := false
    run.osOk := false
    run.deletedCount := 0
    run.attemptStage := ""
    run.attemptDetail := ""
    run.dbHandle := 0
    run.osHandle := 0

    CleanClipboard_UpdateProgress(25, "⏳ " run.attemptTag "[1/3] Cleaning clipboard & Clip Angel in parallel...")

    ; Worker C: dismiss any stray #32770 confirmation dialog so ClipAngel's UI thread is unblocked.
    try ClipAngel_TryConfirmDeleteDialogOnce(true)
    catch {
    }

    ; Worker A: async ClipAngelDb deletenonfav (atomic delete + verify in one process).
    onDbDone(ok, outText, exitCode, timedOut) {
        if (run.finished || run.cancelled)
            return
        run.dbHandle := 0
        run.dbDone := true
        deleted := 0, remaining := 0
        if (ok && ClipAngelDb_ParseDeleteNonFav(outText, &deleted, &remaining)) {
            run.dbOk := true
            run.deletedCount := deleted
        } else {
            run.dbOk := false
            run.attemptStage := "2/3 DB"
            run.attemptDetail := timedOut ? "ClipAngelDb timed out" : "ClipAngelDb deletenonfav failed/busy"
        }
        CleanClipboard_CheckBarrier(run)
    }

    ; Worker A: reuse pre-warmed ClipAngelDb process if alive (near-instant commit), else spawn fresh.
    usedPrewarm := false
    if (run.attempt = 1 && run.sessionId) {
        global g_CleanClipboardPrewarm
        pw := g_CleanClipboardPrewarm
        g_CleanClipboardPrewarm := 0
        if (pw && pw.sessionId = run.sessionId && !pw.completed && pw.handle) {
            pw.onCompleteCallback := onDbDone
            if (ClipAngelDb_SendStdin(pw.handle, "GO`n")) {
                run.dbHandle := pw.handle
                usedPrewarm := true
            } else {
                pw.onCompleteCallback := 0
                try ClipAngelDb_CancelAsync(pw.handle)
                catch {
                }
            }
        }
    }

    if (!usedPrewarm) {
        if (ClipAngelDb_ExePath() != "") {
            run.dbHandle := ClipAngelDb_StartAsync("deletenonfav", 15000, onDbDone, true)
            if (!run.dbHandle) {
                run.dbDone := true
                run.dbOk := false
                run.attemptStage := "2/3 DB"
                run.attemptDetail := "Failed to launch ClipAngelDb.exe"
            }
        } else {
            run.dbDone := true
            run.dbOk := false
        }
    }

    ; Worker B: async OS clipboard clear in parallel with Worker A.
    onOsDone(ok) {
        if (run.finished || run.cancelled)
            return
        run.osHandle := 0
        run.osDone := true
        run.osOk := ok
        CleanClipboard_CheckBarrier(run)
    }

    run.osHandle := CleanClipboard_StartClearOsClipboardAsync(8, onOsDone,
        (*) => (run.cancelled || CleanClipboard_ShouldAbort(run.sessionId)))

    ; If both workers resolved synchronously (e.g. OS clipboard already empty + DB exe missing), advance barrier.
    if (run.dbDone && run.osDone)
        CleanClipboard_CheckBarrier(run)
}

; Internal helper: Performs parallel clipboard cleanup with 3 retry attempts,
; atomic ClipAngelDb delete+verify, ClipAngel in-memory reload, and final guarantee barrier.
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
    StandardLoadingBar_Show("⏳ [1/3] Cleaning clipboard & Clip Angel in parallel...", BANNER_ACCENT_INTERMEDIATE, {
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
        maxAttempts: 3,
        attemptTag: "",
        dbHandle: 0,
        osHandle: 0,
        dbDone: false,
        osDone: false,
        dbOk: false,
        osOk: false,
        deletedCount: 0,
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
