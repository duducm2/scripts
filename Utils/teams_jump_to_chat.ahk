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

TeamsJump_IsProcessRunning() {
    for proc in TeamsJump_Processes() {
        if (ProcessExist(proc))
            return true
    }
    return false
}

; Returns a visible Teams HWND suitable for generic navigation (chat/search), or 0.
TeamsJump_ResolveMainHwnd() {
    for proc in TeamsJump_Processes() {
        for hwnd in WinGetList("ahk_exe " proc) {
            try {
                title := Trim(WinGetTitle(hwnd))
                if (!title)
                    continue
                if InStr(title, "Sharing control bar |")
                    continue
                if (RegExMatch(title, "i)\| Microsoft Teams") || title = "Microsoft Teams")
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

; CheckAndOpenOutlookTeams shows this modal when ms-teams.exe is not running:
; "Teams is closed. Do you want to open it?" (title "Open Applications?").
TeamsJump_OpenPromptIsOpen() {
    return WinExist("Open Applications? ahk_class #32770")
}

; Stay here until that Yes/No box is gone. Do not search for a person while it is up.
TeamsJump_WaitUntilOpenPromptClosed() {
    while TeamsJump_OpenPromptIsOpen()
        Sleep 50
}

; Composer names only. An empty go-to field is not a composer.
TeamsJump_NameIsComposer(text) {
    t := Trim(text)
    if (t = "")
        return false
    for name in TeamsJump_ComposerNameCandidates() {
        if (StrLower(t) = StrLower(name))
            return true
    }
    return false
}

; True when focused element name/id looks like Teams search or Go-To bar.
TeamsJump_IsSearchOrGoToElement(fe) {
    name := ""
    aid := ""
    try name := StrLower(fe.Name)
    catch {
    }
    try aid := StrLower(fe.AutomationId)
    catch {
    }
    for knownAid in ["ms-searchux-input", "search-box-input", "searchinput"] {
        if (aid = knownAid)
            return true
    }
    for kw in ["search", "pesquisar", "go to", "ir para", "ctrl+g", "ctrl+e"] {
        if (InStr(name, kw))
            return true
    }
    return false
}

; One focused element. No tree walk. isGoTo when focus is an Edit/ComboBox and not the composer.
TeamsJump_ReadFocus() {
    info := { id: "", isGoTo: false, isNamedSearch: false }
    fe := 0
    try fe := UIA.GetFocusedElement()
    catch
        return info
    if (!fe)
        return info
    try info.id := UIA.RuntimeIdToString(fe.RuntimeId)
    catch {
    }
    type := 0
    name := ""
    try type := fe.Type
    catch
        return info
    try name := fe.Name
    catch
        name := ""
    if (type != UIA.Type.Edit && type != UIA.Type.ComboBox)
        return info
    if (TeamsJump_NameIsComposer(name))
        return info
    info.isGoTo := true
    info.isNamedSearch := TeamsJump_IsSearchOrGoToElement(fe)
    return info
}

; Ctrl+G once. True when focus moved onto the go-to field (or is already on the named search/go-to field).
TeamsJump_ProbeGoToField() {
    before := TeamsJump_ReadFocus()
    if (before.isGoTo && before.isNamedSearch)
        return true
    Send "{LWin Up}{RWin Up}{LAlt Up}{RAlt Up}{LShift Up}{RShift Up}"
    Send "^g"
    Sleep 300
    after := TeamsJump_ReadFocus()
    if (!after.isGoTo)
        return false
    if (after.isNamedSearch)
        return true
    if (before.id != "" && after.id = before.id)
        return before.isGoTo
    return true
}

; Lightweight background UIA check: true once Teams' search bar is hydrated (no WinActivate needed).
TeamsJump_IsUiHydrated(hwnd) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    try {
        root := UIA.ElementFromHandle(hwnd)
        if (!root)
            return false
        for aid in ["ms-searchux-input", "search-box-input", "searchInput"] {
            try {
                if (root.FindFirst({ AutomationId: aid }))
                    return true
            } catch {
            }
        }
    } catch {
    }
    return false
}

; Gate 2: Cold start after the open prompt. gotoAlreadyOpen: Ctrl+G already focused the field, do not send it again.
; Polls UIA in the background without stealing focus while Teams loads, then activates once to confirm Ctrl+G.
TeamsJump_WaitUntilGoToReady(&hwnd, &gotoAlreadyOpen, timeoutMs := 30000) {
    gotoAlreadyOpen := false
    TeamsJump_WaitUntilOpenPromptClosed()
    start := A_TickCount
    deadline := start + timeoutMs
    lastUpdate := 0
    shown := false
    try {
        StandardLoadingBar_Show("⏳ Waiting until Teams is ready...", BANNER_ACCENT_INTERMEDIATE)
        shown := true
    } catch {
    }
    try {
        while (A_TickCount < deadline) {
            if (TeamsJump_OpenPromptIsOpen()) {
                TeamsJump_WaitUntilOpenPromptClosed()
                continue
            }
            elapsed := Round((A_TickCount - start) / 1000)
            if ((A_TickCount - lastUpdate) >= 800) {
                try StandardLoadingBar_Update("⏳ Waiting until Teams is ready... (" elapsed "s)")
                lastUpdate := A_TickCount
            }
            fresh := TeamsJump_ResolveMainHwnd()
            if (fresh > 0)
                hwnd := fresh
            if (hwnd > 0 && WinExist("ahk_id " hwnd)) {
                ; Only activate Teams once UIA confirms the search UI is hydrated (or after 12s fallback).
                if (TeamsJump_IsUiHydrated(hwnd) || (A_TickCount - start) >= 12000) {
                    if (TeamsJump_ActivateWindowWithRetry(hwnd, 2, 300) && TeamsJump_ProbeGoToField()) {
                        gotoAlreadyOpen := true
                        return true
                    }
                }
            }
            Sleep 400
        }
        return false
    } finally {
        if (shown) {
            try StandardLoadingBar_Hide(0)
        }
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

TeamsJump_EnsureVisibleNoActivate(hwnd) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return
    try {
        if (WinGetMinMax("ahk_id " hwnd) = -1)
            DllCall("user32\ShowWindow", "ptr", hwnd, "int", 4)
    } catch {
    }
}

TeamsJump_SearchAndSelectViaUia(hwndTeams, contact) {
    if (!hwndTeams || contact = "")
        return false
    TeamsJump_EnsureVisibleNoActivate(hwndTeams)
    root := TeamsJump_AttachUiaRoot(hwndTeams)
    if (!root)
        return false
    searchEl := 0
    for aid in ["ms-searchux-input", "search-box-input", "searchInput"] {
        try searchEl := root.FindFirst({ AutomationId: aid })
        catch {
            searchEl := 0
        }
        if (searchEl)
            break
    }
    if (!searchEl) {
        for nm in ["Search", "Pesquisar", "Go to"] {
            try searchEl := root.FindFirst({ Name: nm, ControlType: "Edit", matchmode: "Substring" })
            catch {
                searchEl := 0
            }
            if (searchEl)
                break
        }
    }
    if (!searchEl)
        return false
    try {
        searchEl.Value := contact
    } catch {
        return false
    }
    tokens := StrSplit(StrLower(Trim(contact)), " ")
    deadline := A_TickCount + 3000
    while (A_TickCount < deadline) {
        for ct in ["ListItem", "MenuItem", "Option", "TreeItem"] {
            try {
                items := root.FindAll({ ControlType: ct })
                for item in items {
                    nm := ""
                    try nm := StrLower(item.Name)
                    if (nm = "")
                        continue
                    allMatch := true
                    for tok in tokens {
                        if (tok != "" && !InStr(nm, tok)) {
                            allMatch := false
                            break
                        }
                    }
                    if (allMatch) {
                        try {
                            item.Invoke()
                            return true
                        } catch {
                            try {
                                item.Select()
                                return true
                            } catch {
                            }
                        }
                    }
                }
            } catch {
            }
        }
        Sleep 80
    }
    return false
}

; Gate 3: Ensure Teams is active and focus is actually in the Go-To / Search box before pasting.
TeamsJump_EnsureGoToFocused(hwndTeams, gotoAlreadyOpen := false, attempts := 4) {
    if (!hwndTeams || !WinExist("ahk_id " hwndTeams))
        return false
    if (gotoAlreadyOpen && WinActive("ahk_id " hwndTeams) && TeamsJump_ReadFocus().isGoTo)
        return true
    loop attempts {
        if (!WinActive("ahk_id " hwndTeams)) {
            if (!TeamsJump_ActivateWindowWithRetry(hwndTeams, 2, 300)) {
                Sleep 150
                continue
            }
        }
        if (TeamsJump_ReadFocus().isGoTo)
            return true
        Send "{LWin Up}{RWin Up}{LAlt Up}{RAlt Up}{LShift Up}{RShift Up}"
        Send "^g"
        deadline := A_TickCount + 700
        while (A_TickCount < deadline) {
            if (TeamsJump_ReadFocus().isGoTo)
                return true
            Sleep 100
        }
    }
    return TeamsJump_ReadFocus().isGoTo
}

; Gate 4: Wait for search results to populate and click/invoke the matching contact item.
TeamsJump_SelectSearchResult(hwndTeams, contact, timeoutMs := 5000) {
    if (!hwndTeams || Trim(contact) = "")
        return false
    root := TeamsJump_AttachUiaRoot(hwndTeams)
    if (!root)
        return false
    tokens := StrSplit(StrLower(Trim(contact)), " ")
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        for ct in ["ListItem", "MenuItem", "Option", "TreeItem"] {
            try {
                items := root.FindAll({ ControlType: ct })
                for item in items {
                    nm := ""
                    try nm := StrLower(item.Name)
                    if (nm = "")
                        continue
                    allMatch := true
                    for tok in tokens {
                        if (tok != "" && !InStr(nm, tok)) {
                            allMatch := false
                            break
                        }
                    }
                    if (allMatch) {
                        try {
                            item.Invoke()
                            return true
                        } catch {
                            try {
                                item.Select()
                                return true
                            } catch {
                                try {
                                    item.Click()
                                    return true
                                } catch {
                                }
                            }
                        }
                    }
                }
            } catch {
            }
        }
        Sleep 100
    }
    return false
}

; Gate 5: Confirm Teams actually navigated to the contact's chat (Go-To closed, title or composer ready).
TeamsJump_WaitUntilChatOpened(hwndTeams, contact, timeoutMs := 3500) {
    firstTok := ""
    for tok in StrSplit(StrLower(Trim(contact)), " ") {
        if (tok != "") {
            firstTok := tok
            break
        }
    }
    deadline := A_TickCount + timeoutMs
    sentFallbackEnter := false
    while (A_TickCount < deadline) {
        try {
            title := StrLower(WinGetTitle("ahk_id " hwndTeams))
            if ((InStr(title, "chat |") || InStr(title, "bate-papo |")) && (firstTok = "" || InStr(title, firstTok)))
                return true
        } catch {
        }
        focusInfo := TeamsJump_ReadFocus()
        if (!focusInfo.isGoTo && TeamsJump_FindComposer(hwndTeams))
            return true
        if (!sentFallbackEnter && focusInfo.isGoTo && (deadline - A_TickCount) < (timeoutMs - 800)) {
            sentFallbackEnter := true
            try {
                if (WinActive("ahk_id " hwndTeams))
                    Send "{Enter}"
            } catch {
            }
        }
        Sleep 150
    }
    return false
}

TeamsJumpToChat(contact, backgroundOnly := false, forceColdStart := false) {
    ; Gate 1: Closed means no Teams process or no hydrated main Teams window yet.
    ; If a hydrated main window already exists, never enter the cold-start wait even if forceColdStart=true.
    hwndTeams := TeamsJump_ResolveMainHwnd()
    teamsWasClosed := !TeamsJump_IsProcessRunning() || (hwndTeams <= 0) || (forceColdStart && !TeamsJump_IsUiHydrated(hwndTeams))
    prevHwnd := 0
    try prevHwnd := WinGetID("A")
    catch {
        prevHwnd := 0
    }

    if (!CheckAndOpenOutlookTeams(false, true))
        return false
    TeamsJump_WaitUntilOpenPromptClosed()

    oldWinDelay := A_WinDelay
    oldKeyDelay := A_KeyDelay
    oldControlDelay := A_ControlDelay

    try {
        SetWinDelay 0
        SetKeyDelay 0, 0
        SetControlDelay 0

        hwndTeams := TeamsJump_ResolveMainHwnd()
        if (hwndTeams <= 0) {
            teamsWasClosed := true
            if (!TeamsJump_IsProcessRunning())
                TeamsJump_Run()
        } else if (TeamsJump_IsUiHydrated(hwndTeams)) {
            teamsWasClosed := false
        }

        ; Gate 2: Wait for main window + UIA-verified Ctrl+G readiness on cold start.
        gotoAlreadyOpen := false
        if (teamsWasClosed) {
            if (!TeamsJump_WaitUntilGoToReady(&hwndTeams, &gotoAlreadyOpen, 30000)) {
                try ShowCenteredOverlay_Utils("❌ Teams did not become ready in time.", 2500, BANNER_ACCENT_ERROR)
                return false
            }
        } else if (hwndTeams <= 0) {
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

        if (!gotoAlreadyOpen && !TeamsJump_ActivateWindowWithRetry(hwndTeams, 3, 300)) {
            ShowCenteredOverlay_Utils("❌ Could not activate Teams window.", 2500, BANNER_ACCENT_ERROR)
            return false
        }

        TeamsJump_WaitUntilOpenPromptClosed()

        ; Gate 3: Verify focus is in the Go-To / Search input before pasting the contact name.
        if (!TeamsJump_EnsureGoToFocused(hwndTeams, gotoAlreadyOpen)) {
            ShowCenteredOverlay_Utils("❌ Could not focus Teams Go-To search bar.", 2500, BANNER_ACCENT_ERROR)
            return false
        }

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
        Send "^a"
        Sleep 40
        Send "^v"

        ; Gate 4: Wait for matching contact result to appear and invoke/click it; Enter fallback if UIA misses.
        selected := TeamsJump_SelectSearchResult(hwndTeams, contact, teamsWasClosed ? 6000 : 3500)
        if (!selected) {
            if (TeamsJump_ReadFocus().isGoTo)
                Send "{Enter}"
        }

        ; Gate 5: Confirm navigation landed in the chat before returning or restoring focus.
        opened := TeamsJump_WaitUntilChatOpened(hwndTeams, contact, 3500)
        if (!opened && !selected) {
            ShowCenteredOverlay_Utils("❌ Could not select contact in Teams: " . contact, 2500, BANNER_ACCENT_ERROR)
            return false
        }

        if (backgroundOnly && prevHwnd && prevHwnd != hwndTeams && WinExist("ahk_id " prevHwnd)) {
            try WinActivate("ahk_id " prevHwnd)
            catch {
            }
        }
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
    saved := UiElements_TrySaved(root, "TeamsChat", "Composer")
    if (saved)
        return saved
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
TeamsJump_PasteAndVerify(hwnd, expectedText, maxAttempts := 3, backgroundOnly := false) {
    expected := Trim(expectedText)
    if (!hwnd || expected = "") {
        ShowCenteredOverlay_Utils("❌ Teams paste failed - composer empty", 3000, BANNER_ACCENT_ERROR)
        return false
    }
    prevHwnd := 0
    if (backgroundOnly) {
        try prevHwnd := WinGetID("A")
        catch {
            prevHwnd := 0
        }
        TeamsJump_EnsureVisibleNoActivate(hwnd)
        el := TeamsJump_FindComposer(hwnd)
        if (el) {
            try {
                el.Value := expected
                Sleep 80
                got := TeamsJump_ComposerGetTextViaUia(hwnd)
                if (got != "" && InStr(got, expected))
                    return true
            } catch {
            }
        }
    }
    try {
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
    } finally {
        if (backgroundOnly && prevHwnd && prevHwnd != hwnd && WinExist("ahk_id " prevHwnd)) {
            try WinActivate("ahk_id " prevHwnd)
            catch {
            }
        }
    }
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
