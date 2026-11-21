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

; N/Esc while automation runs (overlay already closed; StandardLoadingBar keys are inactive)
CleanClipboard_SetAbortHotkeys(enable := true) {
    try HotIf()
    catch {
    }
    if (enable) {
        Hotkey("*n", CleanClipboard_OnCancel, "On")
        Hotkey("*Escape", CleanClipboard_OnCancel, "On")
    } else {
        try Hotkey("*n", "Off")
        catch {
        }
        try Hotkey("*Escape", "Off")
        catch {
        }
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

; Internal helper: Performs clipboard cleanup with real step-by-step progress.
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

    prevBusy := g_ClipAngelAutomationBusy
    try {
        osClipCleared := CleanClipboard_ClearOsClipboard()
        if (CleanClipboard_ShouldAbort(sessionId)) {
            StandardLoadingBar_Hide(0)
            return
        }

        ; Fast DB-direct path (with 1 brief retry on SQLITE_BUSY): deletes non-favorites without opening ClipAngel.
        if (ClipAngelDb_ExePath() != "") {
            CleanClipboard_UpdateProgress(35, "⏳ [2/4] Cleaning Clip Angel database...")
            deletedCount := 0
            dbOk := ClipAngelDb_DeleteNonFav(&deletedCount)
            if (!dbOk && !CleanClipboard_ShouldAbort(sessionId)) {
                CleanClipboard_UpdateProgress(50, "⏳ [2/4] Clip Angel DB busy — retrying...")
                Sleep 200
                dbOk := ClipAngelDb_DeleteNonFav(&deletedCount)
            }
            if (CleanClipboard_ShouldAbort(sessionId)) {
                StandardLoadingBar_Hide(0)
                return
            }
            if (dbOk) {
                if (osClipCleared) {
                    detail := (deletedCount = 0) ? "already empty"
                        : (deletedCount = 1 ? "1 clip removed" : deletedCount " clips removed")
                    CleanClipboard_FinishStatus(100, "✅ Clipboard cleaned (" detail ")", BANNER_ACCENT_SUCCESS, 1200, true)
                } else {
                    CleanClipboard_FinishStatus(100, "⚠️ Clip Angel cleaned, but Windows clipboard was locked",
                        BANNER_ACCENT_INTERMEDIATE, 2200, false)
                }
                return
            }
        }

        g_ClipAngelAutomationBusy := true
        CleanClipboard_UpdateProgress(60, "⏳ [3/4] DB unavailable — opening Clip Angel...")
        try ClipAngel_WaitChordModifiersReleased()
        catch {
        }
        try ClipAngel_ReleaseChordModifiersForSend()
        catch {
        }

        root := 0
        hwnd := CleanClipboard_ActivateClipAngelForCleanup(sessionId, &root)
        if (!hwnd || CleanClipboard_ShouldAbort(sessionId)) {
            if (!CleanClipboard_ShouldAbort(sessionId))
                CleanClipboard_FinishStatus(100, "❌ Failed at [3/4]: Clip Angel not running", BANNER_ACCENT_ERROR, 2500, false)
            else
                StandardLoadingBar_Hide(0)
            CleanClipboard_UnwindClipAngel(hwnd)
            return
        }

        CleanClipboard_UpdateProgress(75, "⏳ [3/4] Sending delete non-favorites...")
        try PostMessage(0x001F, 0, 0, , "ahk_id " hwnd) ; WM_CANCELMODE: 0ms menu-close guard
        try ClipAngel_FocusGridIfFilterCombo()
        catch {
        }
        confirmed := false
        try {
            ControlSend("^!k", , "ahk_id " hwnd)
            confirmed := ClipAngel_ConfirmDeleteNonFavoritesDialogBackground(900)
        } catch {
            confirmed := false
        }
        if (!confirmed && !CleanClipboard_ShouldAbort(sessionId)) {
            CleanClipboard_UpdateProgress(85, "⏳ [3/4] Trying Clip menu fallback...")
            if (ClipAngel_InvokeDeleteNonFavoritesViaMenu(hwnd, root))
                confirmed := ClipAngel_ConfirmDeleteNonFavoritesDialogBackground(1500)
        }
        if (CleanClipboard_ShouldAbort(sessionId)) {
            StandardLoadingBar_Hide(0)
            CleanClipboard_UnwindClipAngel(hwnd)
            return
        }

        CleanClipboard_UpdateProgress(95, "⏳ [4/4] Restoring focus & finalizing...")
        CleanClipboard_UnwindClipAngel(hwnd)
        if (priorHwnd)
            ClipAngel_RestorePriorFocus(priorHwnd, true)

        if (confirmed && osClipCleared) {
            CleanClipboard_FinishStatus(100, "✅ Clipboard cleaned", BANNER_ACCENT_SUCCESS, 1200, true)
        } else if (confirmed && !osClipCleared) {
            CleanClipboard_FinishStatus(100, "⚠️ Clip Angel cleaned, but Windows clipboard was locked",
                BANNER_ACCENT_INTERMEDIATE, 2200, false)
        } else {
            CleanClipboard_FinishStatus(100, "❌ Failed at [3/4]: Delete confirmation did not appear",
                BANNER_ACCENT_ERROR, 2500, false)
        }
    } finally {
        g_ClipAngelAutomationBusy := prevBusy
        if (ownsSession)
            g_CleanClipboardInProgress := false
    }
}
