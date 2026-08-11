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

; Prefer a dedicated chat window; skip meeting / share-bar chrome.
TeamsJump_ResolveChatHwnd() {
    for proc in TeamsJump_Processes() {
        for hwnd in WinGetList("ahk_exe " proc) {
            try {
                title := WinGetTitle(hwnd)
                if (!title)
                    continue
                if InStr(title, "Sharing control bar |")
                    continue
                if InStr(title, "Microsoft Teams meeting") || InStr(title, "Reunião do Microsoft Teams")
                    continue
                if InStr(title, "Modo de exibição compacto da reunião")
                    continue
                if InStr(title, "Chat |") || InStr(title, "Bate-papo |")
                    return hwnd
                if (RegExMatch(title, "i)\| Microsoft Teams$") || title = "Microsoft Teams")
                    return hwnd
            } catch {
                continue
            }
        }
    }
    return TeamsJump_ResolveMainHwnd()
}

TeamsJump_ComposerNameCandidates() {
    return [
        "Type a message",
        "Type a message...",
        "Digite uma mensagem",
        "Digite uma mensagem...",
        "Start a new message",
        "Iniciar uma nova mensagem",
        "Message",
        "Mensagem"
    ]
}

; Attach UIA root for Teams (Chromium-first, then HWND root).
TeamsJump_AttachUiaRoot(hwnd) {
    if (!hwnd)
        return 0
    try {
        root := UIA.ElementFromChromium("ahk_id " hwnd, 500)
        if (root)
            return root
    } catch {
    }
    try {
        uia := UIA_Browser("ahk_id " hwnd)
        if (uia)
            return uia
    } catch {
    }
    try {
        return UIA.ElementFromHandle(hwnd)
    } catch {
        return 0
    }
}

; True when text looks like a composer placeholder name, not real typed content.
TeamsJump_IsComposerPlaceholder(text) {
    t := Trim(text)
    if (t = "")
        return true
    for name in TeamsJump_ComposerNameCandidates() {
        if (StrLower(t) = StrLower(name))
            return true
    }
    return false
}

; Find the chat compose Edit/Document element, or 0.
TeamsJump_FindComposer(hwnd) {
    if (!hwnd)
        return 0
    root := TeamsJump_AttachUiaRoot(hwnd)
    if (!root)
        return 0
    names := TeamsJump_ComposerNameCandidates()
    types := ["Edit", "Document"]
    for typeName in types {
        for name in names {
            el := 0
            try el := root.FindFirst({ Name: name, Type: typeName })
            catch {
                el := 0
            }
            if (!el) {
                try el := root.FindFirst({ Name: name, matchmode: "Substring", Type: typeName })
                catch {
                    el := 0
                }
            }
            if (el)
                return el
        }
    }
    ; Last resort: any Edit whose name looks like a message box.
    try {
        edits := root.FindAll({ Type: "Edit" })
        for el in edits {
            n := ""
            try n := el.Name
            nLower := StrLower(n)
            if (InStr(nLower, "message") || InStr(nLower, "mensagem") || InStr(nLower, "type a") || InStr(nLower,
                "digite"))
                return el
        }
    } catch {
    }
    return 0
}

; Focus the chat compose Edit/Document. Returns true if focus likely succeeded.
TeamsJump_FocusComposer(hwnd) {
    if (!hwnd)
        return false
    focused := PasteField_TryFocusMappedField(hwnd)
    el := TeamsJump_FindComposer(hwnd)
    if (el && PasteField_SetFocusWithFallback(el))
        return true
    return focused
}

; Read composer text via UIA Value/TextPattern; empty on failure / placeholder.
TeamsJump_ComposerGetTextViaUia(hwnd) {
    el := TeamsJump_FindComposer(hwnd)
    if (!el)
        return ""
    try {
        text := Trim(el.Value)
        if (text != "" && !TeamsJump_IsComposerPlaceholder(text))
            return text
    } catch {
    }
    try {
        text := Trim(el.TextPattern.DocumentRange.GetText(-1))
        if (text != "" && !TeamsJump_IsComposerPlaceholder(text))
            return text
    } catch {
    }
    return ""
}

; Focus composer and copy selection (clipboard restored). Fallback when UIA Value is empty.
TeamsJump_ComposerGetTextViaClipboard(hwnd) {
    if (!hwnd)
        return ""
    if (!TeamsJump_FocusComposer(hwnd))
        return ""
    Sleep 60
    saved := ClipboardAll()
    try {
        A_Clipboard := ""
        Send "^a"
        Sleep 40
        Send "^c"
        if !ClipWait(1, 1)
            return ""
        text := A_Clipboard
        if (Type(text) != "String")
            text := ""
        text := Trim(text)
        if (TeamsJump_IsComposerPlaceholder(text))
            return ""
        return text
    } finally {
        Sleep 40
        try A_Clipboard := saved
        catch {
        }
    }
}

TeamsJump_ComposerGetText(hwnd) {
    if (!hwnd)
        hwnd := WinExist("A")
    text := TeamsJump_ComposerGetTextViaUia(hwnd)
    if (text != "")
        return text
    return TeamsJump_ComposerGetTextViaClipboard(hwnd)
}

; Paste expectedText into focused composer; verify content landed; retry up to maxAttempts.
TeamsJump_PasteAndVerify(hwnd, expectedText, maxAttempts := 3) {
    expected := Trim(expectedText)
    if (!hwnd || expected = "") {
        ShowCenteredOverlay_Utils("❌ Teams paste failed - composer empty", 3000, BANNER_ACCENT_ERROR)
        return false
    }
    loop maxAttempts {
        if (!WinExist("ahk_id " hwnd)) {
            ShowCenteredOverlay_Utils("❌ Teams paste failed - composer empty", 3000, BANNER_ACCENT_ERROR)
            return false
        }
        if (!WinActive("ahk_id " hwnd)) {
            if (!TeamsJump_ActivateWindowWithRetry(hwnd, 2, 300)) {
                Sleep 100
                continue
            }
        }
        A_Clipboard := ""
        A_Clipboard := expected
        if (!ClipWait(2) || Trim(A_Clipboard) != expected) {
            Sleep 100
            continue
        }
        TeamsJump_FocusComposer(hwnd)
        Sleep 60
        Send "^v"
        Sleep 200
        got := TeamsJump_ComposerGetText(hwnd)
        if (got != "" && InStr(got, expected))
            return true
        Sleep 150
    }
    ShowCenteredOverlay_Utils("❌ Teams paste failed - composer empty", 3000, BANNER_ACCENT_ERROR)
    return false
}

; Activate Teams chat, focus composer via UIA, paste + verify. Never presses Enter to send.
; expectedText: fixed message to paste (default = current A_Clipboard at call time).
TeamsJump_PasteToComposer(expectedText := "") {
    if (expectedText = "")
        expectedText := A_Clipboard
    if (Type(expectedText) != "String")
        expectedText := ""

    if (!CheckAndOpenOutlookTeams(false, true))
        return false

    hwnd := TeamsJump_ResolveChatHwnd()
    if (hwnd <= 0) {
        TeamsJump_Run()
        waitStart := A_TickCount
        while ((A_TickCount - waitStart) < 15000) {
            hwnd := TeamsJump_ResolveChatHwnd()
            if (hwnd > 0)
                break
            Sleep 150
        }
    }
    if (hwnd <= 0) {
        ShowCenteredOverlay_Utils("❌ Teams chat window not found.", 2500, BANNER_ACCENT_ERROR)
        return false
    }
    if (!TeamsJump_ActivateWindowWithRetry(hwnd, 3, 300)) {
        ShowCenteredOverlay_Utils("❌ Could not activate Teams chat.", 2500, BANNER_ACCENT_ERROR)
        return false
    }

    Send "{LWin Up}{RWin Up}{LAlt Up}{RAlt Up}{LShift Up}{RShift Up}"
    Sleep 80
    return TeamsJump_PasteAndVerify(hwnd, expectedText)
}
