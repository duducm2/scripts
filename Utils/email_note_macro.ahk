; =============================================================================
; Utils module: email_note_macro.ahk
; Email note: new mail in Outlook (work) or Gmail (personal), To both inboxes,
; leave focus in Subject for typing.
; Recipients are added one-by-one and quality-gated as chips (rounded pills)
; before Subject is focused — a single SendText of both addresses leaves the
; second (Gmail: both) as uncommitted text with the caret still in To.
; =============================================================================

global EMAIL_NOTE_BOSCH := "eduardo.figueiredo@br.bosch.com"
global EMAIL_NOTE_GMAIL := "edu.evangelista.figueiredo@gmail.com"

EmailNote_Create(subjectText := "") {
    global IS_WORK_ENVIRONMENT
    try {
        if (IS_WORK_ENVIRONMENT)
            EmailNote_CreateOutlook(subjectText)
        else
            EmailNote_CreateGmail(subjectText)
    } catch Error as e {
        ShowCenteredOverlay_Utils("❌ Email note: " . e.Message, 2500, BANNER_ACCENT_ERROR)
    }
}

EmailNote_TargetEmails() {
    global EMAIL_NOTE_BOSCH, EMAIL_NOTE_GMAIL
    return [EMAIL_NOTE_BOSCH, EMAIL_NOTE_GMAIL]
}

EmailNote_UiaRoot(hwnd) {
    if !(hwnd is Integer) || hwnd <= 0
        return 0
    try {
        return UIA.ElementFromHandle(hwnd)
    } catch {
        return 0
    }
}

EmailNote_CreateOutlook(subjectText := "") {
    barOwned := false
    try {
        StandardLoadingBar_Show("⏳ Opening compose...", BANNER_ACCENT_INTERMEDIATE, {
            fontSize: 17,
            trackActiveMonitor: true
        })
        barOwned := true
        if (!EmailNote_EnsureOutlookActive()) {
            StandardLoadingBar_Hide(0)
            barOwned := false
            return
        }
        StandardLoadingBar_Update("⏳ Opening compose...")
        ex := EmailNote_OutlookExe()
        if (ex != "") {
            WinActivate("ahk_exe " ex)
            WinWaitActive("ahk_exe " ex, , 2)
        }
        Send("^1")
        Sleep 250
        Send("^n")
        hwnd := EmailNote_WaitUntilOutlookComposeReady()
        if (!hwnd) {
            StandardLoadingBar_Hide(0)
            barOwned := false
            ShowCenteredOverlay_Utils("❌ Email note: Outlook compose did not become ready", 2500, BANNER_ACCENT_ERROR)
            return
        }
        StandardLoadingBar_Hide(0)
        barOwned := false
        WinActivate("ahk_id " hwnd)
        WinWaitActive("ahk_id " hwnd, , 2)
        EmailNote_FillCompose(hwnd, subjectText, false)
    } catch Error as e {
        if (barOwned)
            StandardLoadingBar_Hide(0)
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

EmailNote_FindOutlookToField(root) {
    if !root
        return 0
    try {
        el := root.FindFirst({ AutomationId: "4117" })
        if el
            return el
    } catch {
    }
    try {
        el := root.FindFirst({ Name: "To", Type: 50004 })
        if el
            return el
    } catch {
    }
    try {
        el := root.FindFirst({ AutomationId: "134", Type: 50026 })
        if el
            return el
    } catch {
    }
    try {
        el := root.FindFirst({ Name: "To:", matchmode: "Substring" })
        if el
            return el
    } catch {
    }
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

EmailNote_FindOutlookSubjectField(root) {
    if !root
        return 0
    try {
        el := root.FindFirst({ AutomationId: "4101" })
        if el
            return el
    } catch {
    }
    try {
        el := root.FindFirst({ Name: "Subject", Type: 50004 })
        if el
            return el
    } catch {
    }
    try {
        el := root.FindFirst({ Name: "Subject", ControlType: "Edit" })
        if el
            return el
    } catch {
    }
    try {
        el := root.FindFirst({ Name: "Assunto", Type: 50004 })
        if el
            return el
    } catch {
    }
    return 0
}

EmailNote_OutlookComposeIsReady(hwnd) {
    root := EmailNote_UiaRoot(hwnd)
    if !root
        return false
    if EmailNote_FindOutlookToField(root)
        return true
    if EmailNote_FindOutlookSubjectField(root)
        return true
    try {
        if root.FindFirst({ AutomationId: "popoutCompose" })
            return true
        if root.FindFirst({ AutomationId: "discardCompose" })
            return true
    } catch {
    }
    return false
}

; Poll until new-mail To/Subject UIA is present (Gmail compose-ready pattern).
; Do not use WinExist("A") — the loading bar can be the foreground window.
EmailNote_FindOutlookComposeHwnd() {
    ex := EmailNote_OutlookExe()
    if (ex = "")
        return 0
    try {
        for hwnd in WinGetList("ahk_exe " ex) {
            if EmailNote_OutlookComposeIsReady(hwnd)
                return hwnd
        }
    } catch {
    }
    return 0
}

EmailNote_WaitUntilOutlookComposeReady(timeoutMs := 15000, neededStreak := 2) {
    openStart := A_TickCount
    deadline := A_TickCount + timeoutMs
    readyStreak := 0
    lastUpdate := 0
    lastHwnd := 0
    while (A_TickCount < deadline) {
        elapsed := Round((A_TickCount - openStart) / 1000)
        if ((A_TickCount - lastUpdate) >= 800) {
            StandardLoadingBar_Update("⏳ Waiting until Outlook compose is ready... (" elapsed "s)")
            lastUpdate := A_TickCount
        }
        hwnd := EmailNote_FindOutlookComposeHwnd()
        if (hwnd > 0) {
            if (hwnd = lastHwnd)
                readyStreak += 1
            else {
                lastHwnd := hwnd
                readyStreak := 1
            }
            if (readyStreak >= neededStreak) {
                WinActivate("ahk_id " hwnd)
                WinWaitActive("ahk_id " hwnd, , 2)
                return hwnd
            }
        } else {
            readyStreak := 0
            lastHwnd := 0
        }
        Sleep 200
    }
    return 0
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
    return EmailNote_UiaRoot(hwnd)
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
                return hwnd
            }
        } else {
            readyStreak := 0
        }
        Sleep 200
    }
    return 0
}

EmailNote_TriggerGmailCompose(hwnd) {
    btn := EmailNote_FindGmailComposeButton(hwnd)
    if btn {
        try {
            btn.Invoke()
            return
        } catch {
            try {
                btn.Click()
                return
            } catch {
            }
        }
    }
    Send("c")
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
    while (A_TickCount < deadline) {
        elapsed := Round((A_TickCount - openStart) / 1000)
        if ((A_TickCount - lastUpdate) >= 800) {
            StandardLoadingBar_Update("⏳ Waiting until Gmail compose is ready... (" elapsed "s)")
            lastUpdate := A_TickCount
        }
        hwnd := EmailNote_FindReadyGmailComposeHwnd(inboxHwnd)
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
    return 0
}

EmailNote_NormalizeSubject(subjectText) {
    subjectText := Trim(subjectText)
    subjectText := StrReplace(subjectText, "`r`n", " ")
    subjectText := StrReplace(subjectText, "`n", " ")
    subjectText := StrReplace(subjectText, "`r", " ")
    return Trim(subjectText)
}

EmailNote_TypeSubjectText(subjectText) {
    subjectText := EmailNote_NormalizeSubject(subjectText)
    if (subjectText = "")
        return
    SendText(subjectText)
}

EmailNote_LocalPart(email) {
    p := InStr(email, "@")
    if (p > 1)
        return SubStr(email, 1, p - 1)
    return email
}

EmailNote_ArrayLength(els) {
    if !IsObject(els)
        return 0
    try {
        return Integer(els.Length)
    } catch {
        return 0
    }
}

EmailNote_FindToChipContainer(root) {
    if !root
        return 0
    el := EmailNote_FindGmailToField(root)
    if el
        return el
    try {
        el := root.FindFirst({ AutomationId: "134", Type: 50026 })
        if el
            return el
    } catch {
    }
    try {
        el := root.FindFirst({ Name: "To:", matchmode: "Substring", Type: 50026 })
        if el
            return el
    } catch {
    }
    return EmailNote_FindOutlookToField(root)
}

; Chip = Button/Group whose name includes the address (or local-part). Skip ListItems
; so an open autocomplete row cannot pass the gate. Search inside To so the
; account button in the window chrome cannot count as a recipient chip.
EmailNote_RecipientChipPresent(root, email) {
    if !root || email = ""
        return false
    scope := EmailNote_FindToChipContainer(root)
    if !scope
        scope := root
    localPart := EmailNote_LocalPart(email)
    for type in [50000, 50026] {
        try {
            el := scope.FindFirst({ Name: email, Type: type, matchmode: "Substring", cs: false })
            if el
                return true
        } catch {
        }
        if (localPart != "" && localPart != email) {
            try {
                el := scope.FindFirst({ Name: localPart, Type: type, matchmode: "Substring", cs: false })
                if el
                    return true
            } catch {
            }
        }
    }
    return false
}

EmailNote_CountOutlookRecipientEntities(root) {
    if !root
        return 0
    scope := EmailNote_FindToChipContainer(root)
    if !scope
        scope := root
    n := 0
    try {
        n := Max(n, EmailNote_ArrayLength(scope.FindAll({ ClassName: "_EType_RECIPIENT_ENTITY",
            matchmode: "Substring" })))
    } catch {
    }
    if (n > 0)
        return n
    try {
        n := Max(n, EmailNote_ArrayLength(scope.FindAll({ AutomationId: "REK", matchmode: "Substring",
            Type: 50026 })))
    } catch {
    }
    return n
}

EmailNote_CountRecipientChips(hwnd) {
    root := EmailNote_UiaRoot(hwnd)
    if !root
        return 0
    named := 0
    for email in EmailNote_TargetEmails() {
        if EmailNote_RecipientChipPresent(root, email)
            named += 1
    }
    return Max(named, EmailNote_CountOutlookRecipientEntities(root))
}

EmailNote_WaitUntilRecipientChipCount(hwnd, needed, timeoutMs := 4000, neededStreak := 2) {
    deadline := A_TickCount + timeoutMs
    readyStreak := 0
    while (A_TickCount < deadline) {
        if (hwnd > 0 && EmailNote_CountRecipientChips(hwnd) >= needed) {
            readyStreak += 1
            if (readyStreak >= neededStreak)
                return true
        } else {
            readyStreak := 0
        }
        Sleep 120
    }
    return false
}

EmailNote_FocusToField(hwnd, isGmail, allowClick := false) {
    if (hwnd > 0) {
        WinActivate("ahk_id " hwnd)
        WinWaitActive("ahk_id " hwnd, , 2)
    }
    root := EmailNote_UiaRoot(hwnd)
    el := isGmail ? EmailNote_FindGmailToField(root) : EmailNote_FindOutlookToField(root)
    if !el
        return false
    try el.SetFocus()
    catch {
        return false
    }
    Sleep 40
    ; Clicking To after a chip exists can hit the first pill instead of the input.
    if (allowClick && !isGmail) {
        try el.Click()
        catch {
        }
    }
    return true
}

EmailNote_CommitRecipientToken(isGmail) {
    ; Separator keeps caret in To so the next address can be typed.
    ; Tab after a blob often leaves (or jumps past) an uncommitted second address.
    if (isGmail)
        SendText(",")
    else
        SendText(";")
}

EmailNote_AddOneRecipient(hwnd, email, expectedCount, isGmail) {
    if !EmailNote_FocusToField(hwnd, isGmail, expectedCount = 1) {
        ShowCenteredOverlay_Utils("❌ Email note: To field not found", 2500, BANNER_ACCENT_ERROR)
        return false
    }
    needed := Max(expectedCount, EmailNote_CountRecipientChips(hwnd) + 1)
    Sleep 80
    SendText(email)
    Sleep 120
    EmailNote_CommitRecipientToken(isGmail)
    if EmailNote_WaitUntilRecipientChipCount(hwnd, needed)
        return true
    Send("{Enter}")
    if EmailNote_WaitUntilRecipientChipCount(hwnd, needed, 2500)
        return true
    Send("{Tab}")
    if EmailNote_WaitUntilRecipientChipCount(hwnd, needed, 2500)
        return true
    ShowCenteredOverlay_Utils("❌ Email note: Recipient was not added as a chip", 2500, BANNER_ACCENT_ERROR)
    return false
}

EmailNote_FocusSubjectField(hwnd, isGmail) {
    if (hwnd > 0) {
        WinActivate("ahk_id " hwnd)
        WinWaitActive("ahk_id " hwnd, , 2)
    }
    if EmailNote_IsSubjectFocused()
        return true
    root := EmailNote_UiaRoot(hwnd)
    el := isGmail ? EmailNote_FindGmailSubjectField(root) : EmailNote_FindOutlookSubjectField(root)
    if !el
        return false
    try el.SetFocus()
    catch {
        return false
    }
    Sleep 40
    try el.Click()
    catch {
    }
    return EmailNote_IsSubjectFocused()
}

; Quality gate: Subject must actually have keyboard focus before typing.
; UIA Value assignment can fill Subject while the caret stays in To.
EmailNote_WaitUntilSubjectFocused(hwnd, isGmail, timeoutMs := 5000, neededStreak := 2) {
    deadline := A_TickCount + timeoutMs
    readyStreak := 0
    lastFocusTry := 0
    while (A_TickCount < deadline) {
        if EmailNote_IsSubjectFocused() {
            readyStreak += 1
            if (readyStreak >= neededStreak)
                return true
        } else {
            readyStreak := 0
            if ((A_TickCount - lastFocusTry) >= 280) {
                EmailNote_FocusSubjectField(hwnd, isGmail)
                lastFocusTry := A_TickCount
            }
        }
        Sleep 100
    }
    return EmailNote_IsSubjectFocused()
}

EmailNote_FillCompose(hwnd, subjectText := "", isGmail := false) {
    if (hwnd > 0) {
        WinActivate("ahk_id " hwnd)
        WinWaitActive("ahk_id " hwnd, , 2)
    }
    expected := 0
    for email in EmailNote_TargetEmails() {
        expected += 1
        if !EmailNote_AddOneRecipient(hwnd, email, expected, isGmail)
            return false
    }
    if !EmailNote_WaitUntilRecipientChipCount(hwnd, 2, 2500, 2) {
        ShowCenteredOverlay_Utils("❌ Email note: Both recipient chips were not ready", 2500, BANNER_ACCENT_ERROR)
        return false
    }
    if !EmailNote_WaitUntilSubjectFocused(hwnd, isGmail) {
        ShowCenteredOverlay_Utils("❌ Email note: Could not focus Subject", 2500, BANNER_ACCENT_ERROR)
        return false
    }
    EmailNote_TypeSubjectText(subjectText)
    if !EmailNote_IsSubjectFocused()
        EmailNote_WaitUntilSubjectFocused(hwnd, isGmail, 2000, 1)
    return true
}

EmailNote_FillGmailCompose(hwnd, subjectText := "") {
    return EmailNote_FillCompose(hwnd, subjectText, true)
}

EmailNote_CreateGmail(subjectText := "") {
    barOwned := false
    try {
        StandardLoadingBar_Show("⏳ Opening compose...", BANNER_ACCENT_INTERMEDIATE, {
            fontSize: 17,
            trackActiveMonitor: true
        })
        barOwned := true

        composeHwnd := EmailNote_FindReadyGmailComposeHwnd()
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
            EmailNote_TriggerGmailCompose(inboxHwnd)
            composeHwnd := EmailNote_WaitUntilGmailComposeReady(inboxHwnd)
        } else {
            WinActivate("ahk_id " composeHwnd)
            WinWaitActive("ahk_id " composeHwnd, , 2)
        }

        if (!composeHwnd) {
            StandardLoadingBar_Hide(0)
            barOwned := false
            ShowCenteredOverlay_Utils("❌ Email note: Gmail compose did not become ready", 2500, BANNER_ACCENT_ERROR)
            return
        }

        StandardLoadingBar_Hide(0)
        barOwned := false
        WinActivate("ahk_id " composeHwnd)
        WinWaitActive("ahk_id " composeHwnd, , 2)
        ok := EmailNote_FillGmailCompose(composeHwnd, subjectText)
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
        if (n != "" && (InStr(n, "Subject") || InStr(n, "Assunto")))
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
    return EmailNote_FocusSubjectField(WinExist("A"), false)
}
