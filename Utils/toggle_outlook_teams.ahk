; =============================================================================
; Utils module: toggle_outlook_teams.ahk
; ToggleOutlookAndTeams macro
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; Toggle Outlook and Teams
; Toggles Outlook and Teams applications to manage RAM usage.
; If either is open: Kills both so their RAM is released. Closing wins over opening.
; That close also kills Edge, OneNote, Spotify, and WhatsApp. They are not relaunched.
; If neither is open: Launches Outlook and Teams only.
; =============================================================================
ToggleOutlookAndTeams() {
    loadingShown := false
    try {
        ; Either app running means close. Open only when both are already gone.
        outlookRunning := OutlookProcessRunning()
        teamsRunning := ProcessExist("ms-teams.exe") || ProcessExist("Teams.exe") || ProcessExist("MSTeams.exe")
        isOpeningFlow := !(outlookRunning || teamsRunning)
        hadError := false
        firstError := ""

        ; Show start banner
        if (!isOpeningFlow) {
            ShowCenteredOverlay_Utils("📤 Closing apps to free RAM...", 1500, BANNER_ACCENT_INTERMEDIATE)
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

        if (outlookRunning || teamsRunning) {
            ; Kill every matching process. ProcessClose stops one per call,
            ; and Teams keeps several ms-teams.exe processes (plus the tray host).
            try {
                KillAllProcessesByName("OUTLOOK.EXE")
                KillAllProcessesByName("olk.exe")
            } catch Error as e {
                MsgBox "Error closing Outlook: " e.Message
            }

            try {
                KillAllProcessesByName("ms-teams.exe")
                KillAllProcessesByName("Teams.exe")
                KillAllProcessesByName("MSTeams.exe")
            } catch Error as e {
                MsgBox "Error closing Teams: " e.Message
            }

            ; Extra resident apps. Not part of the open/close decision, and not started again.
            for procName in ["msedge.exe", "ONENOTE.EXE", "Spotify.exe", "WhatsApp.exe"] {
                try KillAllProcessesByName(procName)
                catch Error as e {
                    MsgBox "Error closing " procName ": " e.Message
                }
            }
            try KillWhatsAppChromeApp()
            catch Error as e {
                MsgBox "Error closing WhatsApp: " e.Message
            }
        } else {
            ; Neither is running: launch both applications
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

; ProcessClose terminates a single matching process. Loop until the name is gone.
KillAllProcessesByName(name) {
    loop 30 {
        if !ProcessExist(name)
            return
        ProcessClose(name)
    }
}

; #!+z starts native WhatsApp.exe at work, and a Chrome PWA at home
; (apps do Chrome\WhatsApp Web.lnk). Kill a Chrome process only when it hosts
; WhatsApp and no other Chrome window. Otherwise just close the WhatsApp window.
KillWhatsAppChromeApp() {
    waHwnds := []
    waPids := Map()
    otherPids := Map()
    for hwnd in WinGetList("ahk_exe chrome.exe") {
        pid := 0
        isWa := false
        try pid := WinGetPID("ahk_id " hwnd)
        try isWa := WhatsAppJump_IsWhatsAppChromeAppHwnd(hwnd)
        if isWa {
            waHwnds.Push(hwnd)
            if pid
                waPids[pid] := true
        } else if pid
            otherPids[pid] := true
    }
    if (waHwnds.Length = 0)
        return

    for pid, _ in waPids {
        if !otherPids.Has(pid)
            ProcessClose(pid)
    }
    ; No other Chrome windows: drop leftover renderer processes too.
    if (otherPids.Count = 0) {
        KillAllProcessesByName("chrome.exe")
    } else {
        for hwnd in waHwnds {
            try {
                pid := WinGetPID("ahk_id " hwnd)
                if otherPids.Has(pid)
                    WinClose(hwnd)
            }
        }
    }
    try WhatsAppJump_InvalidateHwndCache()
}
