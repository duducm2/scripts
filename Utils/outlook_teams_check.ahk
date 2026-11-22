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

; Esc while automation runs (do NOT hook bare *n so typing words with 'n' in another window never aborts)
CleanClipboard_SetAbortHotkeys(enable := true) {
    try HotIf()
    catch {
    }
    if (enable) {
        Hotkey("*Escape", CleanClipboard_OnCancel, "On")
    } else {
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

; Internal helper: Performs clipboard cleanup with 3 under-the-hood retry attempts,
; real step-by-step progress, ClipAngel in-memory reload, and error logging/reporting.
; sessionId: macro session from CleanClipboard_ShowCountdown; 0 = dictation legacy path
CleanClipboardInternal(sessionId := 0) {
    global g_ClipAngelAutomationBusy, g_CleanClipboardInProgress
    if (CleanClipboard_ShouldAbort(sessionId)) {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        CleanClipboard_UnwindClipAngel()
        return
    }

    ownsSession := (sessionId = 0)
    if (ownsSession) {
        if (g_CleanClipboardInProgress)
            return
        g_CleanClipboardInProgress := true
    }

    priorHwnd := ClipAngel_ResolvePriorHwnd(0)
    StandardLoadingBar_Show("⏳ [1/4] Clearing Windows clipboard...", BANNER_ACCENT_INTERMEDIATE, {
        passive: false,
        manualProgress: true,
        noActivate: true,
        centerOnHwnd: priorHwnd,
        fontSize: 17
    })
    try StandardLoadingBar_SetProgressValue(15)
    catch {
    }

    maxAttempts := 3
    lastStage := ""
    lastDetail := ""
    lastClipAngelOk := false
    prevBusy := g_ClipAngelAutomationBusy
    try {
        loop maxAttempts {
            attempt := A_Index
            attemptTag := (attempt > 1) ? ("[Attempt " attempt "/" maxAttempts "] ") : ""
            if (CleanClipboard_ShouldAbort(sessionId)) {
                StandardLoadingBar_Hide(0)
                return
            }

            CleanClipboard_UpdateProgress(15, "⏳ " attemptTag "[1/4] Clearing Windows clipboard...")
            osClipCleared := CleanClipboard_ClearOsClipboard(6)
            if (CleanClipboard_ShouldAbort(sessionId)) {
                StandardLoadingBar_Hide(0)
                return
            }

            clipAngelOk := false
            usedDb := false
            deletedCount := 0
            attemptStage := ""
            attemptDetail := ""

            ; Fast DB-direct path: delete non-favorites, verify 0 remain, and trigger ReloadList in ClipAngel.
            if (ClipAngelDb_ExePath() != "") {
                CleanClipboard_UpdateProgress(35, "⏳ " attemptTag "[2/4] Cleaning Clip Angel database...")
                dbOk := ClipAngelDb_DeleteNonFav(&deletedCount)
                if (!dbOk && !CleanClipboard_ShouldAbort(sessionId)) {
                    CleanClipboard_UpdateProgress(50, "⏳ " attemptTag "[2/4] Clip Angel DB busy — retrying...")
                    Sleep 200
                    dbOk := ClipAngelDb_DeleteNonFav(&deletedCount)
                }
                if (CleanClipboard_ShouldAbort(sessionId)) {
                    StandardLoadingBar_Hide(0)
                    return
                }
                if (dbOk) {
                    verifyDeleted := 0
                    if (ClipAngelDb_DeleteNonFav(&verifyDeleted))
                        deletedCount += verifyDeleted
                    ClipAngel_TriggerBackgroundReloadList()
                    clipAngelOk := true
                    usedDb := true
                } else {
                    attemptStage := "2/4 DB"
                    attemptDetail := "ClipAngelDb deletenonfav failed/busy"
                }
            }

            ; UI fallback when DB helper is unavailable or locked.
            if (!clipAngelOk) {
                g_ClipAngelAutomationBusy := true
                hwnd := 0
                try {
                    CleanClipboard_UpdateProgress(60, "⏳ " attemptTag "[3/4] DB unavailable — opening Clip Angel...")
                    try ClipAngel_WaitChordModifiersReleased()
                    catch {
                    }
                    try ClipAngel_ReleaseChordModifiersForSend()
                    catch {
                    }

                    root := 0
                    hwnd := CleanClipboard_ActivateClipAngelForCleanup(sessionId, &root)
                    if (!hwnd) {
                        if (CleanClipboard_ShouldAbort(sessionId)) {
                            StandardLoadingBar_Hide(0)
                            return
                        }
                        attemptStage := "3/4 Open"
                        attemptDetail := "Clip Angel window not found"
                    } else if (CleanClipboard_ShouldAbort(sessionId)) {
                        StandardLoadingBar_Hide(0)
                        return
                    } else {
                        CleanClipboard_UpdateProgress(75, "⏳ " attemptTag "[3/4] Sending delete non-favorites...")
                        ; Dismiss any leftover confirmation dialog first so ClipAngel's UI thread isn't blocked.
                        confirmed := ClipAngel_TryConfirmDeleteDialogOnce(true)
                        if (!confirmed) {
                            try PostMessage(0x001F, 0, 0, , "ahk_id " hwnd) ; WM_CANCELMODE: 0ms menu-close guard
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
                        if (!confirmed && !CleanClipboard_ShouldAbort(sessionId)) {
                            CleanClipboard_UpdateProgress(85, "⏳ " attemptTag "[3/4] Trying List menu fallback...")
                            ; Poll in background during UIA menu invoke because MessageBox.Show blocks the UI thread.
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
                        if (CleanClipboard_ShouldAbort(sessionId)) {
                            StandardLoadingBar_Hide(0)
                            return
                        }
                        if (confirmed) {
                            clipAngelOk := true
                        } else {
                            dbPrefix := (attemptDetail != "") ? (attemptDetail " -> ") : ""
                            attemptStage := "3/4 UI"
                            attemptDetail := dbPrefix "Delete confirmation dialog did not appear"
                        }
                    }
                } finally {
                    CleanClipboard_UnwindClipAngel(hwnd)
                    if (priorHwnd)
                        ClipAngel_RestorePriorFocus(priorHwnd, true)
                    g_ClipAngelAutomationBusy := prevBusy
                }
            }

            if (CleanClipboard_ShouldAbort(sessionId)) {
                StandardLoadingBar_Hide(0)
                return
            }

            ; Re-verify OS clipboard is empty after Clip Angel cleanup.
            CleanClipboard_UpdateProgress(95, "⏳ " attemptTag "[4/4] Verifying clipboard empty...")
            if (!osClipCleared || DllCall("CountClipboardFormats", "int") != 0)
                osClipCleared := CleanClipboard_ClearOsClipboard(4)

            lastClipAngelOk := clipAngelOk
            if (clipAngelOk && osClipCleared) {
                retrySuffix := (attempt > 1) ? (" [attempt " attempt "/" maxAttempts "]") : ""
                if (usedDb) {
                    detail := (deletedCount = 0) ? "already empty"
                        : (deletedCount = 1 ? "1 clip removed" : deletedCount " clips removed")
                    CleanClipboard_FinishStatus(100, "✅ Clipboard cleaned (" detail ")" retrySuffix,
                        BANNER_ACCENT_SUCCESS, 1200, true)
                } else {
                    CleanClipboard_FinishStatus(100, "✅ Clipboard cleaned" retrySuffix,
                        BANNER_ACCENT_SUCCESS, 1200, true)
                }
                return
            }

            if (clipAngelOk && !osClipCleared) {
                lastStage := "4/4 OS Clipboard"
                lastDetail := "Windows clipboard locked by another app"
            } else {
                lastStage := (attemptStage != "") ? attemptStage : "3/4 Cleanup"
                lastDetail := (attemptDetail != "") ? attemptDetail : "Unknown failure"
            }
            CleanClipboard_LogError(attempt, lastStage, lastDetail)

            if (attempt < maxAttempts && !CleanClipboard_ShouldAbort(sessionId)) {
                CleanClipboard_UpdateProgress(20,
                    "⚠️ Attempt " attempt "/" maxAttempts " failed [" lastStage "]: " lastDetail
                    . " — retrying (" (attempt + 1) "/" maxAttempts ")...",
                    BANNER_ACCENT_INTERMEDIATE)
                Sleep (250 * attempt)
            }
        }

        if (CleanClipboard_ShouldAbort(sessionId)) {
            StandardLoadingBar_Hide(0)
            return
        }
        if (lastClipAngelOk) {
            CleanClipboard_FinishStatus(100,
                "⚠️ Clip Angel cleaned, but Windows clipboard stayed locked (3/3)",
                BANNER_ACCENT_INTERMEDIATE, 3000, false)
        } else {
            CleanClipboard_FinishStatus(100,
                "❌ Failed (3/3) at [" lastStage "]: " lastDetail,
                BANNER_ACCENT_ERROR, 3500, false)
        }
    } finally {
        g_ClipAngelAutomationBusy := prevBusy
        if (ownsSession)
            g_CleanClipboardInProgress := false
    }
}
