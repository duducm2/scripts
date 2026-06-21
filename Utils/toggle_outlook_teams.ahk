; =============================================================================
; Utils module: toggle_outlook_teams.ahk
; ToggleOutlookAndTeams macro
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; Toggle Outlook and Teams
; Toggles Outlook and Teams applications to manage RAM usage.
; If both are open: Closes Outlook and minimizes Teams to system tray.
; If one or both are closed: Launches both applications.
; =============================================================================
ToggleOutlookAndTeams() {
    loadingShown := false
    try {
        ; Check if both applications are running
        outlookRunning := OutlookProcessRunning()
        teamsRunning := ProcessExist("ms-teams.exe")
        isOpeningFlow := !(outlookRunning && teamsRunning)
        hadError := false
        firstError := ""

        ; Show start banner
        if (!isOpeningFlow) {
            ShowCenteredOverlay_Utils("📤 Closing Outlook and Teams...", 1500, BANNER_ACCENT_INTERMEDIATE)
        } else {
            StandardLoadingBar_Show("⏳ Opening Outlook and Teams...", BANNER_ACCENT_INTERMEDIATE, {
                passive: false,
                centerOnHwnd: 0,
                textWidth: 560,
                fontSize: 17,
                passiveBgColor: BANNER_ACCENT_INTERMEDIATE
            })
            loadingShown := true
        }

        if (outlookRunning && teamsRunning) {
            ; Both are open: Close Outlook and minimize Teams to system tray
            ; Close Outlook process(es) - classic and/or Store (olk.exe)
            try {
                if ProcessExist("OUTLOOK.EXE")
                    ProcessClose("OUTLOOK.EXE")
                if ProcessExist("olk.exe")
                    ProcessClose("olk.exe")
            } catch Error as e {
                MsgBox "Error closing Outlook: " e.Message
            }

            ; Close all Teams windows (this keeps Teams in system tray)
            try {
                ; Teams can have multiple process names, check all
                for hwnd in WinGetList("ahk_exe ms-teams.exe") {
                    WinClose(hwnd)
                }
                ; Also check for Teams.exe and MSTeams.exe variants
                for hwnd in WinGetList("ahk_exe Teams.exe") {
                    WinClose(hwnd)
                }
                for hwnd in WinGetList("ahk_exe MSTeams.exe") {
                    WinClose(hwnd)
                }
            } catch Error as e {
                MsgBox "Error closing Teams windows: " e.Message
            }
        } else {
            ; One or both are closed: Launch both applications
            ; Launch Outlook
            if (!outlookRunning) {
                StandardLoadingBar_Update("⏳ Opening Outlook...")
                try {
                    outlookPath := ""
                    if (IS_WORK_ENVIRONMENT) {
                        ; Try work environment shortcut path
                        outlookPath := "C:\Users\fie7ca\Documents\Atalhos\Microsoft Outlook.lnk"
                        if (!FileExist(outlookPath)) {
                            outlookPath := ""
                        }
                    } else {
                        ; Try personal environment shortcut path
                        outlookPath :=
                            "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Microsoft Outlook.lnk"
                        if (!FileExist(outlookPath)) {
                            outlookPath := ""
                        }
                    }

                    ; Launch using shortcut if available, otherwise olk.exe or OUTLOOK.EXE
                    if (outlookPath != "") {
                        Run outlookPath
                    } else {
                        olkPath := OutlookGetOlkExePath()
                        if (olkPath != "")
                            Run olkPath
                        else
                            Run "OUTLOOK.EXE"
                    }
                } catch Error as e {
                    hadError := true
                    if (firstError = "")
                        firstError := "Outlook: " . e.Message
                }
            }

            ; Launch/Activate Teams
            ; Simplified approach: Just run the executable. This handles both launching and bringing to front.
            StandardLoadingBar_Update("⏳ Opening Teams...")
            try {
                if (IS_WORK_ENVIRONMENT) {
                    teamsExePath :=
                        "C:\Program Files\WindowsApps\MSTeams_25332.1210.4188.1171_x64__8wekyb3d8bbwe\ms-teams.exe"
                    if (FileExist(teamsExePath)) {
                        Run teamsExePath
                    } else {
                        Run "ms-teams.exe"
                    }
                } else {
                    ; Personal environment
                    teamsPath :=
                        "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Microsoft Teams.lnk"
                    if (FileExist(teamsPath)) {
                        Run teamsPath
                    } else {
                        Run "ms-teams.exe"
                    }
                }

                ; Wait for window to appear and become active
                if (WinWaitActive("ahk_exe ms-teams.exe", , 10)) {
                } else {
                    hadError := true
                    if (firstError = "")
                        firstError := "Teams window not found"
                }
            } catch Error as e {
                hadError := true
                if (firstError = "")
                    firstError := "Teams: " . e.Message
            }

            ; Second: Activate Outlook last (so it gets final focus)
            StandardLoadingBar_Update("⏳ Activating Outlook...")
            try {
                if (OutlookProcessRunning()) {
                    ex := ProcessExist("OUTLOOK.EXE") ? "OUTLOOK.EXE" : "olk.exe"
                    WinWait("ahk_exe " ex, , 5)
                    if (!WinExist("ahk_exe " ex)) {
                        hadError := true
                        if (firstError = "")
                            firstError := "Outlook not running"
                    } else {
                        WinActivate("ahk_exe " ex)
                        WinWaitActive("ahk_exe " ex, , 2)
                    }
                } else {
                    hadError := true
                    if (firstError = "")
                        firstError := "Outlook process not detected"
                }
            } catch Error as e {
                hadError := true
                if (firstError = "")
                    firstError := "Outlook activation: " . e.Message
            }

            if (loadingShown) {
                StandardLoadingBar_Hide(0)
                loadingShown := false
            }

            if (hadError) {
                ShowCenteredOverlay_Utils("❌ Open completed with issues: " . firstError, 2500, BANNER_ACCENT_ERROR)
            } else {
                ShowCenteredOverlay_Utils("✅ Outlook and Teams opened", 1500, BANNER_ACCENT_SUCCESS)
            }

            return
        }

        ; Show finish banner
        ShowCenteredOverlay_Utils("✅ Done", 1500, BANNER_ACCENT_SUCCESS)
    } catch Error as e {
        if (loadingShown)
            StandardLoadingBar_Hide(0)
        MsgBox "Error in ToggleOutlookAndTeams macro: " e.Message
    }
}
