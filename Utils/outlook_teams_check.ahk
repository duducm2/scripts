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
; Returns: true if applications are running (or were opened), false if user cancelled
; =============================================================================
CheckAndOpenOutlookTeams(checkOutlook := false, checkTeams := false) {
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

    ; Build message based on what's closed
    message := ""
    if (outlookClosed && teamsClosed) {
        message := "Outlook and Teams are closed. Do you want to open them?"
    } else if (outlookClosed) {
        message := "Outlook is closed. Do you want to open it?"
    } else if (teamsClosed) {
        message := "Teams is closed. Do you want to open it?"
    }

    ; Show message box
    response := MsgBox(message, "Open Applications?", "YesNo Icon?")

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

            ; Launch Teams if closed
            if (teamsClosed) {
                if (IS_WORK_ENVIRONMENT) {
                    teamsExePath :=
                        "C:\Program Files\WindowsApps\MSTeams_25332.1210.4188.1171_x64__8wekyb3d8bbwe\ms-teams.exe"
                    if (FileExist(teamsExePath)) {
                        Run teamsExePath
                    } else {
                        Run "ms-teams.exe"
                    }
                } else {
                    teamsPath :=
                        "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Microsoft Teams.lnk"
                    if (FileExist(teamsPath)) {
                        Run teamsPath
                    } else {
                        Run "ms-teams.exe"
                    }
                }
            }

            ; Wait a bit for applications to start
            Sleep 2000
            return true
        } catch Error as e {
            MsgBox "Error opening applications: " e.Message, "Error", "IconX"
            return false
        }
    }

    ; User cancelled
    return false
}

; Chime for "clean now" confirmations (desktop recycle Y, clean clipboard Y). Not used on auto-timeout.
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

CleanClipboard_UnwindClipAngel() {
    ; Clip Angel stays visible; no unwind needed.
}

; N/Esc while automation runs (overlay already closed; StandardLoadingBar keys are inactive)
CleanClipboard_SetAbortHotkeys(enable := true) {
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

; Open/focus Clip Angel for cleanup (native Alt+P fallback, layout correction). Returns hwnd or 0.
CleanClipboard_ActivateClipAngelForCleanup(sessionId := 0) {
    if (CleanClipboard_ShouldAbort(sessionId))
        return 0

    ClipAngel_WaitChordModifiersReleased()
    ClipAngel_ReleaseChordModifiersForSend()

    hwnd := ClipAngel_MainHwnd()
    if (!hwnd) {
        priorSendLevel := A_SendLevel
        SendLevel 0
        SendInput "{Alt up}{Shift up}{Win up}{Ctrl up}"
        Sleep CLIPANGEL_ALT_P_SETTLE_MS
        SendInput "!p"
        Sleep CLIPANGEL_ALT_P_SETTLE_MS
        SendLevel priorSendLevel
        hwnd := ClipAngel_WaitForMainHwnd()
    }

    if (CleanClipboard_ShouldAbort(sessionId))
        return 0

    if (!hwnd) {
        ShowCenteredOverlay_Utils("❌ Clip Angel not found — is it running?", 2500, BANNER_ACCENT_ERROR)
        return 0
    }

    if !ClipAngel_ShowWindow(hwnd)
        ClipAngel_EnsureVisibleAndLayout(hwnd, 0, true)

    if (!WinWaitActive("ahk_id " hwnd, , 0.8))
        ClipAngel_EnsureWindowActive(hwnd, 800)

    if (CleanClipboard_ShouldAbort(sessionId))
        return 0

    return hwnd
}

; Internal helper: Performs clipboard cleanup without showing prompt
; sessionId: macro session from CleanClipboard_ShowCountdown; 0 = dictation legacy path (no session guard)
CleanClipboardInternal(sessionId := 0) {
    Sleep 200
    if (CleanClipboard_ShouldAbort(sessionId)) {
        CleanClipboard_UnwindClipAngel()
        return
    }

    StandardLoadingBar_Show("⏳ Opening Clip Angel...", BANNER_ACCENT_INTERMEDIATE)
    hwnd := CleanClipboard_ActivateClipAngelForCleanup(sessionId)
    if (!hwnd || CleanClipboard_ShouldAbort(sessionId)) {
        StandardLoadingBar_Hide(0)
        CleanClipboard_UnwindClipAngel()
        return
    }

    StandardLoadingBar_Update("⏳ Preparing clip list...", BANNER_ACCENT_INTERMEDIATE)
    ClipAngel_LeaveFavoritesFilter(hwnd)
    Sleep 200
    if (CleanClipboard_ShouldAbort(sessionId)) {
        StandardLoadingBar_Hide(0)
        CleanClipboard_UnwindClipAngel()
        return
    }

    StandardLoadingBar_Update("⏳ Deleting non-favorites...", BANNER_ACCENT_INTERMEDIATE)
    priorSendLevel := A_SendLevel
    SendLevel 0
    ClipAngel_ReleaseChordModifiersForSend()
    SendInput "^!k"
    Sleep 600
    if (CleanClipboard_ShouldAbort(sessionId)) {
        SendLevel priorSendLevel
        StandardLoadingBar_Hide(0)
        CleanClipboard_UnwindClipAngel()
        return
    }

    SendInput "{Enter}"
    SendLevel priorSendLevel
    Sleep 300
    StandardLoadingBar_Update("✅ Clipboard cleaned", BANNER_ACCENT_SUCCESS)
    StandardLoadingBar_Hide(500)
}
