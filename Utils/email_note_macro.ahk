; =============================================================================
; Utils module: email_note_macro.ahk
; Email note: new mail in Outlook (work) or Gmail (personal), To both inboxes,
; leave focus in Subject for typing.
; =============================================================================

global EMAIL_NOTE_BOSCH := "eduardo.figueiredo@br.bosch.com"
global EMAIL_NOTE_GMAIL := "edu.evangelista.figueiredo@gmail.com"

; #region agent log
EmailNote_DbgJsonStr(s) {
    s := SubStr(s, 1, 80)
    s := StrReplace(s, "\", "\\")
    s := StrReplace(s, '"', "'")
    s := StrReplace(s, "`n", " ")
    s := StrReplace(s, "`r", "")
    return s
}
EmailNote_Dbg(hypo, loc, msg, data := "{}") {
    global g_EmailNoteDbgRun
    runId := (IsSet(g_EmailNoteDbgRun) && g_EmailNoteDbgRun != "") ? g_EmailNoteDbgRun : "pre-fix"
    try FileAppend('{"sessionId":"c18ca5","runId":"' runId '","hypothesisId":"' hypo '","location":"' loc '","message":"' msg
        . '","data":' data ',"timestamp":' A_TickCount '}`n', A_ScriptDir "\debug-c18ca5.log", "UTF-8")
    catch {
    }
}
EmailNote_DbgWin(hwnd) {
    t := "", cls := ""
    try t := WinGetTitle("ahk_id " hwnd)
    catch {
    }
    try cls := WinGetClass("ahk_id " hwnd)
    catch {
    }
    return '{"hwnd":' Integer(hwnd) ',"hasCompose":' (InStr(t, "Compose Mail") ? 1 : 0) ',"hasGmail":' (InStr(t,
        "Gmail")
        ? 1 : 0) ',"title":"' EmailNote_DbgJsonStr(t) '","class":"' EmailNote_DbgJsonStr(cls) '"}'
}
; #endregion

EmailNote_Create() {
    global IS_WORK_ENVIRONMENT
    try {
        if (IS_WORK_ENVIRONMENT) {
            if (!EmailNote_EnsureOutlookActive())
                return
            Send("^1")
            Sleep 250
            Send("^n")
            Sleep 450
            SendText(EMAIL_NOTE_BOSCH "; " EMAIL_NOTE_GMAIL)
            Sleep 80
            Send("{Tab}")
            Sleep 80
            EmailNote_FocusOutlookSubjectIfNeeded()
        } else {
            EmailNote_CreateGmail()
        }
    } catch Error as e {
        ShowCenteredOverlay_Utils("❌ Email note: " . e.Message, 2500, BANNER_ACCENT_ERROR)
    }
}

EmailNote_OutlookExe() {
    if ProcessExist("olk.exe")
        return "olk.exe"
    if ProcessExist("OUTLOOK.EXE")
        return "OUTLOOK.EXE"
    return ""
}

EmailNote_LaunchOutlook() {
    global IS_WORK_ENVIRONMENT
    outlookPath := ""
    if (IS_WORK_ENVIRONMENT) {
        outlookPath := "C:\Users\fie7ca\Documents\Atalhos\Microsoft Outlook.lnk"
        if (!FileExist(outlookPath))
            outlookPath := ""
    } else {
        outlookPath := "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Microsoft Outlook.lnk"
        if (!FileExist(outlookPath))
            outlookPath := ""
    }
    if (outlookPath != "") {
        Run outlookPath
        return
    }
    olkPath := OutlookGetOlkExePath()
    if (olkPath != "")
        Run olkPath
    else
        Run "OUTLOOK.EXE"
}

EmailNote_EnsureOutlookActive() {
    if (!OutlookProcessRunning()) {
        try {
            EmailNote_LaunchOutlook()
        } catch Error as e {
            ShowCenteredOverlay_Utils("❌ Email note: Outlook launch failed: " . e.Message, 2500, BANNER_ACCENT_ERROR)
            return false
        }
    }
    ex := ""
    deadline := A_TickCount + 10000
    while (A_TickCount < deadline) {
        ex := EmailNote_OutlookExe()
        if (ex != "" && WinExist("ahk_exe " ex))
            break
        Sleep 150
    }
    if (ex = "" || !WinExist("ahk_exe " ex)) {
        ShowCenteredOverlay_Utils("❌ Email note: Outlook window not found", 2500, BANNER_ACCENT_ERROR)
        return false
    }
    WinActivate("ahk_exe " ex)
    if (!WinWaitActive("ahk_exe " ex, , 3)) {
        ShowCenteredOverlay_Utils("❌ Email note: Could not activate Outlook", 2500, BANNER_ACCENT_ERROR)
        return false
    }
    return true
}

EmailNote_ChromeTitleHas(hwnd, needle) {
    try {
        return InStr(WinGetTitle("ahk_id " hwnd), needle)
    } catch {
        return false
    }
}

EmailNote_FindGmailHwnd(preferCompose := false) {
    prev := A_TitleMatchMode
    SetTitleMatchMode(2)
    hwnd := 0
    composeHwnd := 0
    inboxHwnd := 0
    try {
        for cand in WinGetList("ahk_exe chrome.exe") {
            try {
                title := WinGetTitle("ahk_id " cand)
            } catch {
                continue
            }
            if !InStr(title, "Gmail")
                continue
            if InStr(title, "Compose Mail") {
                if (!composeHwnd)
                    composeHwnd := cand
            } else if (!inboxHwnd) {
                inboxHwnd := cand
            }
        }
        if (preferCompose)
            hwnd := composeHwnd ? composeHwnd : inboxHwnd
        else
            hwnd := inboxHwnd ? inboxHwnd : composeHwnd
    } finally {
        SetTitleMatchMode(prev)
    }
    return hwnd
}

EmailNote_GmailUiaRoot(hwnd) {
    if !(hwnd is Integer) || hwnd <= 0
        return 0
    try {
        return UIA.ElementFromHandle(hwnd)
    } catch {
        return 0
    }
}

EmailNote_FindGmailToField(root) {
    if !root
        return 0
    for nm in ["To recipients", "Destinatários"] {
        try {
            el := root.FindFirst({ Type: 50003, Name: nm })
            if el
                return el
        } catch {
        }
    }
    return 0
}

EmailNote_FindGmailSubjectField(root) {
    if !root
        return 0
    for nm in ["Subject", "Assunto"] {
        try {
            el := root.FindFirst({ Type: 50004, Name: nm })
            if el
                return el
        } catch {
        }
    }
    return 0
}

EmailNote_GmailComposeIsReady(hwnd) {
    root := EmailNote_GmailUiaRoot(hwnd)
    if !root
        return false
    if EmailNote_FindGmailToField(root)
        return true
    if EmailNote_FindGmailSubjectField(root)
        return true
    return false
}

; Title "Gmail - Google Chrome" is a loading tab; real mailbox titles include Inbox and the address.
EmailNote_GmailInboxIsReady(hwnd) {
    if !(hwnd is Integer) || hwnd <= 0
        return false
    t := ""
    try t := WinGetTitle("ahk_id " hwnd)
    catch {
        return false
    }
    if !InStr(t, "Gmail")
        return false
    if (t = "Gmail - Google Chrome" || t = "Gmail")
        return false
    if InStr(t, "Inbox") || InStr(t, "@")
        return true
    return false
}

EmailNote_FindGmailComposeButton(hwnd) {
    root := EmailNote_GmailUiaRoot(hwnd)
    if !root
        return 0
    for nm in ["Compose", "Compose Mail", "Escrever"] {
        try {
            el := root.FindFirst({ Type: 50000, Name: nm })
            if el
                return el
        } catch {
        }
    }
    return 0
}

EmailNote_WaitUntilGmailInboxReady(initialHwnd := 0, timeoutMs := 15000, neededStreak := 2) {
    openStart := A_TickCount
    deadline := A_TickCount + timeoutMs
    readyStreak := 0
    lastUpdate := 0
    while (A_TickCount < deadline) {
        elapsed := Round((A_TickCount - openStart) / 1000)
        if ((A_TickCount - lastUpdate) >= 800) {
            StandardLoadingBar_Update("⏳ Waiting until Gmail inbox is ready... (" elapsed "s)")
            lastUpdate := A_TickCount
        }
        hwnd := EmailNote_FindGmailHwnd(false)
        if (hwnd <= 0 && initialHwnd > 0 && WinExist("ahk_id " initialHwnd))
            hwnd := initialHwnd
        if (hwnd > 0 && EmailNote_GmailInboxIsReady(hwnd)) {
            readyStreak += 1
            if (readyStreak >= neededStreak) {
                WinActivate("ahk_id " hwnd)
                WinWaitActive("ahk_id " hwnd, , 2)
                ; #region agent log
                EmailNote_Dbg("B", "email_note_macro.ahk:WaitUntilGmailInboxReady", "inboxReady", '{"hwnd":'
                    Integer(hwnd) ',"win":' EmailNote_DbgWin(hwnd) ',"hasComposeBtn":' (
                        EmailNote_FindGmailComposeButton(hwnd) ? 1 : 0) '}')
                ; #endregion
                return hwnd
            }
        } else {
            readyStreak := 0
        }
        Sleep 200
    }
    ; #region agent log
    EmailNote_Dbg("B", "email_note_macro.ahk:WaitUntilGmailInboxReady", "inboxTimeout", '{"initial":' Integer(
        initialHwnd) '}')
    ; #endregion
    return 0
}

EmailNote_TriggerGmailCompose(hwnd) {
    btn := EmailNote_FindGmailComposeButton(hwnd)
    usedBtn := 0
    if btn {
        try {
            btn.Invoke()
            usedBtn := 1
        } catch {
            try {
                btn.Click()
                usedBtn := 1
            } catch {
            }
        }
    }
    if (!usedBtn)
        Send("c")
    ; #region agent log
    EmailNote_Dbg("B", "email_note_macro.ahk:TriggerGmailCompose", "triggered", '{"usedBtn":' usedBtn ',"active":'
        EmailNote_DbgWin(WinExist("A")) '}')
    ; #endregion
}

EmailNote_FindReadyGmailComposeHwnd(inboxHwnd := 0) {
    composeHwnd := EmailNote_FindGmailHwnd(true)
    if (composeHwnd && EmailNote_ChromeTitleHas(composeHwnd, "Compose Mail") && EmailNote_GmailComposeIsReady(
        composeHwnd))
        return composeHwnd
    if (inboxHwnd && EmailNote_GmailComposeIsReady(inboxHwnd))
        return inboxHwnd
    if (composeHwnd && EmailNote_GmailComposeIsReady(composeHwnd))
        return composeHwnd
    return 0
}

; Poll until compose popout/inline UIA is present (SpotifyDictation_WaitUntilUiReady pattern).
EmailNote_WaitUntilGmailComposeReady(inboxHwnd := 0, timeoutMs := 15000, neededStreak := 2) {
    openStart := A_TickCount
    deadline := A_TickCount + timeoutMs
    readyStreak := 0
    lastUpdate := 0
    pollCount := 0
    while (A_TickCount < deadline) {
        elapsed := Round((A_TickCount - openStart) / 1000)
        if ((A_TickCount - lastUpdate) >= 800) {
            StandardLoadingBar_Update("⏳ Waiting until Gmail compose is ready... (" elapsed "s)")
            lastUpdate := A_TickCount
        }
        hwnd := EmailNote_FindReadyGmailComposeHwnd(inboxHwnd)
        pollCount += 1
        pref := EmailNote_FindGmailHwnd(true)
        if (pollCount = 1 || hwnd > 0 || Mod(pollCount, 15) = 0) {
            rootPref := EmailNote_GmailUiaRoot(pref)
            rootIn := EmailNote_GmailUiaRoot(inboxHwnd)
            ; #region agent log
            EmailNote_Dbg("C", "email_note_macro.ahk:WaitUntilGmailComposeReady", "poll", '{"poll":' pollCount
                . ',"readyHwnd":' Integer(hwnd) ',"streak":' readyStreak ',"prefTo":' (EmailNote_FindGmailToField(
                    rootPref) ? 1 : 0) ',"prefSub":' (EmailNote_FindGmailSubjectField(rootPref) ? 1 : 0) ',"inboxTo":' (
                        EmailNote_FindGmailToField(rootIn) ? 1 : 0) ',"pref":' EmailNote_DbgWin(pref) '}')
            ; #endregion
        }
        if (hwnd > 0) {
            readyStreak += 1
            if (readyStreak >= neededStreak) {
                WinActivate("ahk_id " hwnd)
                WinWaitActive("ahk_id " hwnd, , 2)
                return hwnd
            }
        } else {
            readyStreak := 0
        }
        Sleep 200
    }
    ; #region agent log
    EmailNote_Dbg("D", "email_note_macro.ahk:WaitUntilGmailComposeReady", "timeout", '{"polls":' pollCount '}')
    ; #endregion
    return 0
}

EmailNote_FillGmailCompose(hwnd) {
    global EMAIL_NOTE_BOSCH, EMAIL_NOTE_GMAIL
    root := EmailNote_GmailUiaRoot(hwnd)
    if !root
        return false
    toEl := EmailNote_FindGmailToField(root)
    if !toEl {
        ShowCenteredOverlay_Utils("❌ Email note: To field not found", 2500, BANNER_ACCENT_ERROR)
        return false
    }
    try toEl.SetFocus()
    catch {
        ShowCenteredOverlay_Utils("❌ Email note: Could not focus To", 2500, BANNER_ACCENT_ERROR)
        return false
    }
    Sleep 80
    SendText(EMAIL_NOTE_BOSCH ", " EMAIL_NOTE_GMAIL)
    Sleep 80
    root := EmailNote_GmailUiaRoot(hwnd)
    subEl := EmailNote_FindGmailSubjectField(root)
    if subEl {
        try subEl.SetFocus()
        catch {
        }
    }
    return true
}

EmailNote_CreateGmail() {
    global g_EmailNoteDbgRun
    g_EmailNoteDbgRun := "post-fix"
    barOwned := false
    try {
        StandardLoadingBar_Show("⏳ Opening compose...", BANNER_ACCENT_INTERMEDIATE, {
            fontSize: 17,
            trackActiveMonitor: true
        })
        barOwned := true

        composeHwnd := EmailNote_FindReadyGmailComposeHwnd()
        ; #region agent log
        EmailNote_Dbg("A", "email_note_macro.ahk:CreateGmail", "start", '{"alreadyReady":' (composeHwnd ? 1 : 0)
        . ',"active":' EmailNote_DbgWin(WinExist("A")) '}')
        ; #endregion
        if (!composeHwnd) {
            StandardLoadingBar_Update("⏳ Activating Gmail...")
            if (!EmailNote_EnsureGmailActive()) {
                StandardLoadingBar_Hide(0)
                barOwned := false
                return
            }
            inboxHwnd := WinExist("A")
            StandardLoadingBar_Update("⏳ Waiting until Gmail inbox is ready...")
            inboxHwnd := EmailNote_WaitUntilGmailInboxReady(inboxHwnd)
            if (!inboxHwnd) {
                StandardLoadingBar_Hide(0)
                barOwned := false
                ShowCenteredOverlay_Utils("❌ Email note: Gmail inbox did not become ready", 2500, BANNER_ACCENT_ERROR)
                return
            }
            StandardLoadingBar_Update("⏳ Opening compose...")
            ; #region agent log
            EmailNote_Dbg("A", "email_note_macro.ahk:CreateGmail", "beforeSendC", '{"inbox":' EmailNote_DbgWin(
                inboxHwnd) ',"active":' EmailNote_DbgWin(WinExist("A")) '}')
            ; #endregion
            EmailNote_TriggerGmailCompose(inboxHwnd)
            ; #region agent log
            EmailNote_Dbg("B", "email_note_macro.ahk:CreateGmail", "afterSendC", '{"active":' EmailNote_DbgWin(WinExist(
                "A")) '}')
            ; #endregion
            composeHwnd := EmailNote_WaitUntilGmailComposeReady(inboxHwnd)
        } else {
            WinActivate("ahk_id " composeHwnd)
            WinWaitActive("ahk_id " composeHwnd, , 2)
        }

        if (!composeHwnd) {
            StandardLoadingBar_Hide(0)
            barOwned := false
            ; #region agent log
            EmailNote_Dbg("E", "email_note_macro.ahk:CreateGmail", "composeNotReady", '{"active":' EmailNote_DbgWin(
                WinExist("A")) '}')
            ; #endregion
            ShowCenteredOverlay_Utils("❌ Email note: Gmail compose did not become ready", 2500, BANNER_ACCENT_ERROR)
            return
        }

        StandardLoadingBar_Update("⏳ Filling To...")
        ok := EmailNote_FillGmailCompose(composeHwnd)
        ; #region agent log
        EmailNote_Dbg("C", "email_note_macro.ahk:CreateGmail", "fillDone", '{"ok":' (ok ? 1 : 0) ',"compose":'
        EmailNote_DbgWin(composeHwnd) '}')
        ; #endregion
        StandardLoadingBar_Hide(0)
        barOwned := false
        if !ok
            return
    } catch Error as e {
        if (barOwned)
            StandardLoadingBar_Hide(0)
        ShowCenteredOverlay_Utils("❌ Email note: " . e.Message, 2500, BANNER_ACCENT_ERROR)
    }
}

EmailNote_EnsureGmailActive() {
    hwnd := EmailNote_FindGmailHwnd()
    if (!hwnd) {
        try {
            Run('chrome.exe "https://mail.google.com"')
        } catch Error as e {
            ShowCenteredOverlay_Utils("❌ Email note: Chrome launch failed: " . e.Message, 2500, BANNER_ACCENT_ERROR)
            return false
        }
        deadline := A_TickCount + 15000
        while (A_TickCount < deadline) {
            hwnd := EmailNote_FindGmailHwnd()
            if (hwnd)
                break
            Sleep 200
        }
    }
    if (!hwnd) {
        ShowCenteredOverlay_Utils("❌ Email note: Gmail window not found", 2500, BANNER_ACCENT_ERROR)
        return false
    }
    WinActivate("ahk_id " hwnd)
    if (!WinWaitActive("ahk_id " hwnd, , 3)) {
        ShowCenteredOverlay_Utils("❌ Email note: Could not activate Gmail", 2500, BANNER_ACCENT_ERROR)
        return false
    }
    return true
}

EmailNote_IsSubjectFocused() {
    try {
        focused := UIA.GetFocusedElement()
        if (!focused)
            return false
        n := ""
        try n := focused.Name
        if (n != "" && InStr(n, "Subject"))
            return true
        aid := ""
        try aid := focused.AutomationId
        if (aid = "4101")
            return true
    } catch {
    }
    return false
}

EmailNote_FocusOutlookSubjectIfNeeded() {
    if EmailNote_IsSubjectFocused()
        return
    try {
        hwnd := WinExist("A")
        if (!hwnd)
            return
        root := UIA.ElementFromHandle(hwnd)
        if (!root)
            return
        el := 0
        try el := root.FindFirst({ AutomationId: "4101" })
        if (!el)
            try el := root.FindFirst({ Name: "Subject", Type: 50004 })
        if (!el)
            try el := root.FindFirst({ Name: "Subject", ControlType: "Edit" })
        if (el)
            el.SetFocus()
    } catch {
    }
}
