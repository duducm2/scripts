; =============================================================================
; Utils module: teams_jump_to_chat.ahk
; Shared Jump-to-Chat logic for Teams
; Do NOT #include Lib\TeamsContext.ahk here — that file is already included by
; Microsoft Teams.ahk and Shift keys\teams_predicates.ahk; a second include
; redefines class TeamsContextCache and aborts Act reloads.
; =============================================================================

; Local process list (same as TEAMS_PROCESSES in Lib\TeamsContext.ahk).
TeamsJump_Processes() {
    return ["ms-teams.exe", "Teams.exe", "MSTeams.exe"]
}

; Returns a visible Teams HWND suitable for generic navigation (chat/search), or 0.
TeamsJump_ResolveMainHwnd() {
    for proc in TeamsJump_Processes() {
        for hwnd in WinGetList("ahk_exe " proc) {
            try {
                title := WinGetTitle(hwnd)
                if (!title)
                    continue
                if InStr(title, "Sharing control bar |")
                    continue
                if (RegExMatch(title, "i)\| Microsoft Teams"))
                    return hwnd
            } catch {
                continue
            }
        }
    }
    return 0
}

TeamsJump_ResolvePath() {
    global IS_WORK_ENVIRONMENT
    localAppData := EnvGet("LOCALAPPDATA")
    if (localAppData) {
        teamsExe := localAppData "\Microsoft\WindowsApps\ms-teams.exe"
        if (FileExist(teamsExe))
            return teamsExe
    }
    appData := EnvGet("APPDATA")
    if (appData) {
        shortcut := appData "\Microsoft\Windows\Start Menu\Programs\Microsoft Teams.lnk"
        if (FileExist(shortcut))
            return shortcut
    }
    try {
        loop reg "HKCU\SOFTWARE\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\Repository\Packages",
            "K" {
            if InStr(A_LoopRegName, "MSTeams_") && InStr(A_LoopRegName, "8wekyb3d8bbwe") {
                basePath := "C:\Program Files\WindowsApps\" A_LoopRegName "\ms-teams.exe"
                if (FileExist(basePath))
                    return basePath
            }
        }
    } catch {
    }
    return "ms-teams:"
}

TeamsJump_Run() {
    path := TeamsJump_ResolvePath()
    try {
        if (path = "ms-teams:")
            Run "ms-teams:"
        else
            Run path
    } catch {
        Run "ms-teams.exe"
    }
}

TeamsJump_ActivateWindowWithRetry(hwnd, attempts := 3, waitMs := 300) {
    if (!hwnd || hwnd <= 0 || !WinExist("ahk_id " hwnd)) {
        try ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
        catch {
        }
        return false
    }
    originalState := ""
    try {
        originalState := WinGetMinMax(hwnd)
        if !(originalState = -1 || originalState = 0 || originalState = 1)
            originalState := ""
    } catch {
        originalState := ""
    }
    loop attempts {
        try {
            if (originalState = -1) {
                WinRestore(hwnd)
                Sleep 100
            }
            WinActivate(hwnd)
            if WinWaitActive("ahk_id " hwnd, , waitMs / 1000)
                return true
        } catch {
            try ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
        }
        try {
            if (originalState = -1) {
                DllCall("ShowWindow", "Ptr", hwnd, "Int", 9)
                Sleep 100
            }
            DllCall("SetForegroundWindow", "Ptr", hwnd)
            if WinWaitActive("ahk_id " hwnd, , waitMs / 1000)
                return true
        } catch {
            try ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
        }
        try {
            DllCall("BringWindowToTop", "Ptr", hwnd)
            Sleep 100
            WinActivate(hwnd)
            if WinWaitActive("ahk_id " hwnd, , waitMs / 1000)
                return true
        } catch {
            try ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
        }
        Sleep 200
    }
    return false
}

TeamsJumpToChat(contact) {
    if (!CheckAndOpenOutlookTeams(false, true))
        return false

    oldWinDelay := A_WinDelay
    oldKeyDelay := A_KeyDelay
    oldControlDelay := A_ControlDelay

    try {
        SetWinDelay 0
        SetKeyDelay 0, 0
        SetControlDelay 0

        hwndTeams := TeamsJump_ResolveMainHwnd()
        if (hwndTeams <= 0) {
            TeamsJump_Run()
            waitStart := A_TickCount
            while ((A_TickCount - waitStart) < 15000) {
                hwndTeams := TeamsJump_ResolveMainHwnd()
                if (hwndTeams > 0)
                    break
                Sleep 150
            }
        }
        if (hwndTeams <= 0) {
            try ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
            return false
        }

        if !TeamsJump_ActivateWindowWithRetry(hwndTeams, 3, 300) {
            ShowCenteredOverlay_Utils("❌ Could not activate Teams window.", 2500, BANNER_ACCENT_ERROR)
            return false
        }

        Send "{LWin Up}{RWin Up}{LAlt Up}{RAlt Up}{LShift Up}{RShift Up}"

        Send "^g"
        Sleep 100
        loop 5 {
            A_Clipboard := ""
            A_Clipboard := contact
            if ClipWait(2) && (A_Clipboard = contact)
                break
            if A_Index = 5 {
                ShowCenteredOverlay_Utils("❌ CLIPBOARD ERROR - TRY AGAIN", 3000, BANNER_ACCENT_ERROR)
                return false
            }
            Sleep 100
        }
        Send "^v"
        Sleep 200
        Sleep 600
        Send "{Enter}"
        return true
    } finally {
        SetWinDelay oldWinDelay
        SetKeyDelay oldKeyDelay, 0
        SetControlDelay oldControlDelay
    }
}
