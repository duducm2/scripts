; =============================================================================
; Utils module: email_note_macro.ahk
; Email note: new mail in Outlook (work) or Gmail (personal), To both inboxes,
; leave focus in Subject for typing.
; To/Subject selectors match New Outlook (Group Name "To", MSG_*_SUBJECT) and
; Gmail compose ("Compose: New Message", To row ClassName "aoD hl"). Recipients
; are typed one-by-one with a separator; chip/text gates search only the To
; element. Outlook Cc/Bcc are open — never Tab to reach Subject.
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
    hwnd := 0
    try {
        StandardLoadingBar_Show("⏳ Opening compose...", BANNER_ACCENT_INTERMEDIATE, {
            fontSize: 17,
            trackActiveMonitor: true
        })
        if (!EmailNote_EnsureOutlookActive())
            return
        ex := EmailNote_OutlookExe()
        if (ex != "") {
            WinActivate("ahk_exe " ex)
            WinWaitActive("ahk_exe " ex, , 2)
            hwnd := WinExist("ahk_exe " ex)
        }
        if (!(hwnd is Integer) || hwnd <= 0) {
            ShowCenteredOverlay_Utils("❌ Email note: Outlook window not found", 2500, BANNER_ACCENT_ERROR)
            return
        }
        if (!EmailNote_OutlookComposeIsReady(hwnd)) {
            StandardLoadingBar_Update("⏳ Opening compose...")
            Send("^1")
            Sleep 80
            Send("^n")
            hwnd := EmailNote_WaitUntilOutlookComposeReady(hwnd)
        }
        if (!(hwnd is Integer) || hwnd <= 0) {
            ShowCenteredOverlay_Utils("❌ Email note: Outlook compose did not become ready", 2500, BANNER_ACCENT_ERROR)
            return
        }
    } catch Error as e {
        ShowCenteredOverlay_Utils("❌ Email note: " . e.Message, 2500, BANNER_ACCENT_ERROR)
        return
    } finally {
        StandardLoadingBar_Hide(0)
    }
    if (!(hwnd is Integer) || hwnd <= 0)
        return
    WinActivate("ahk_id " hwnd)
    WinWaitActive("ahk_id " hwnd, , 2)
    EmailNote_FillCompose(hwnd, subjectText, false)
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

; New Outlook: Group Name "To" (EditorClass), well MSG_*_TO — not classic 4117/134.
EmailNote_FindOutlookToField(root) {
    if !root
        return 0
    try {
        el := root.FindFirst({ Name: "To", Type: 50026 })
        if el
            return el
    } catch {
    }
    try {
        el := root.FindFirst({ AutomationId: "_TO", matchmode: "Substring" })
        if el {
            try {
                inner := el.FindFirst({ Name: "To", Type: 50026 })
                if inner
                    return inner
            } catch {
            }
            return el
        }
    } catch {
    }
    try {
        el := root.FindFirst({ AutomationId: "recipient-well-label-to" })
        if el
            return el
    } catch {
    }
    return 0
}

EmailNote_FindOutlookSubjectField(root) {
    if !root
        return 0
    try {
        el := root.FindFirst({ Name: "Subject", Type: 50004 })
        if el
            return el
    } catch {
    }
    try {
        el := root.FindFirst({ AutomationId: "_SUBJECT", matchmode: "Substring" })
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

; One OR FindFirst on the already-activated hwnd (no WinGetList, no To ladder).
EmailNote_OutlookComposeIsReady(hwnd) {
    root := EmailNote_UiaRoot(hwnd)
    if !root
        return false
    try {
        el := root.FindFirst([{ AutomationId: "discardCompose" }, { AutomationId: "popoutCompose" }, { Name: "Subject",
            Type: 50004 }])
        return !!el
    } catch {
    }
    return false
}

EmailNote_WaitUntilOutlookComposeReady(hwnd, timeoutMs := 15000) {
    if !(hwnd is Integer) || hwnd <= 0
        return 0
    openStart := A_TickCount
    deadline := A_TickCount + timeoutMs
    lastUpdate := 0
    while (A_TickCount < deadline) {
        elapsed := Round((A_TickCount - openStart) / 1000)
        if ((A_TickCount - lastUpdate) >= 800) {
            StandardLoadingBar_Update("⏳ Waiting until Outlook compose is ready... (" elapsed "s)")
            lastUpdate := A_TickCount
        }
        if EmailNote_OutlookComposeIsReady(hwnd)
            return hwnd
        Sleep 40
    }
    return 0
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

EmailNote_FindGmailComposeContainer(root) {
    if !root
        return 0
    try {
        el := root.FindFirst({ Name: "Compose: New Message", Type: 50032 })
        if el
            return el
    } catch {
    }
    try {
        el := root.FindFirst({ Name: "Compose: New Message", Type: 50026 })
        if el
            return el
    } catch {
    }
    try {
        el := root.FindFirst({ Name: "Compose: New Message" })
        if el
            return el
    } catch {
    }
    return 0
}

; ComboBox labels if present; else Gmail To row ClassName "aoD hl".
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
    try {
        el := root.FindFirst({ ClassName: "aoD hl", matchmode: "Substring" })
        if el
            return el
    } catch {
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
    try {
        el := root.FindFirst([{ Name: "Compose: New Message" }, { Name: "Subject", Type: 50004 }, { Name: "Assunto",
            Type: 50004 }])
        return !!el
    } catch {
    }
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
    if (inboxHwnd is Integer && inboxHwnd > 0 && EmailNote_GmailComposeIsReady(inboxHwnd))
        return inboxHwnd
    composeHwnd := EmailNote_FindGmailHwnd(true)
    if (composeHwnd && EmailNote_GmailComposeIsReady(composeHwnd))
        return composeHwnd
    return 0
}

EmailNote_WaitUntilGmailComposeReady(inboxHwnd := 0, timeoutMs := 15000) {
    openStart := A_TickCount
    deadline := A_TickCount + timeoutMs
    lastUpdate := 0
    while (A_TickCount < deadline) {
        elapsed := Round((A_TickCount - openStart) / 1000)
        if ((A_TickCount - lastUpdate) >= 800) {
            StandardLoadingBar_Update("⏳ Waiting until Gmail compose is ready... (" elapsed "s)")
            lastUpdate := A_TickCount
        }
        hwnd := EmailNote_FindReadyGmailComposeHwnd(inboxHwnd)
        if (hwnd > 0) {
            WinActivate("ahk_id " hwnd)
            WinWaitActive("ahk_id " hwnd, , 2)
            return hwnd
        }
        Sleep 40
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

EmailNote_FindToField(searchRoot, isGmail) {
    return isGmail ? EmailNote_FindGmailToField(searchRoot) : EmailNote_FindOutlookToField(searchRoot)
}

; One FindFirst per needle/type. Gmail uncommitted To is Text (50020), not a Button pill.
EmailNote_ScopeHasAddress(scope, email) {
    if !scope || email = ""
        return false
    localPart := EmailNote_LocalPart(email)
    needles := [email]
    if (localPart != "" && localPart != email)
        needles.Push(localPart)
    for type in [50020, 50000, 50026] {
        for needle in needles {
            try {
                el := scope.FindFirst({ Name: needle, Type: type, matchmode: "Substring", cs: false })
                if el
                    return true
            } catch {
            }
        }
    }
    try {
        n := scope.Name
        if (n != "" && (InStr(n, email, false) || (localPart != "" && InStr(n, localPart, false))))
            return true
    } catch {
    }
    ; Gmail To ComboBox (50003) puts the typed address in Value, not a named Text child.
    try {
        v := scope.Value
        if (v != "" && (InStr(v, email, false) || (localPart != "" && InStr(v, localPart, false))))
            return true
    } catch {
    }
    return false
}

; Fresh To first, then searchRoot (stale To row after the first address is committed).
EmailNote_ToHasAddress(searchRoot, isGmail, email) {
    toEl := EmailNote_FindToField(searchRoot, isGmail)
    if (toEl && EmailNote_ScopeHasAddress(toEl, email))
        return true
    if (searchRoot && EmailNote_ScopeHasAddress(searchRoot, email))
        return true
    return false
}

EmailNote_WaitUntilToHasAddress(searchRoot, isGmail, email, timeoutMs := 1500) {
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        if EmailNote_ToHasAddress(searchRoot, isGmail, email)
            return true
        Sleep 40
    }
    return false
}

; Click only on the first address — later clicks hit the Bosch pill.
EmailNote_FocusToField(searchRoot, isGmail, allowClick := false) {
    toEl := EmailNote_FindToField(searchRoot, isGmail)
    if !toEl
        return false
    try toEl.SetFocus()
    catch {
        return false
    }
    Sleep 40
    if (allowClick) {
        try toEl.Click()
        catch {
        }
    }
    return true
}

EmailNote_FindSubjectField(root, isGmail) {
    return isGmail ? EmailNote_FindGmailSubjectField(root) : EmailNote_FindOutlookSubjectField(root)
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
        if (aid != "" && InStr(aid, "_SUBJECT"))
            return true
        if (aid = "4101")
            return true
    } catch {
    }
    return false
}

EmailNote_FocusSubjectField(subEl) {
    if !subEl
        return false
    if EmailNote_IsSubjectFocused()
        return true
    try subEl.SetFocus()
    catch {
        return false
    }
    Sleep 40
    try subEl.Click()
    catch {
    }
    return EmailNote_IsSubjectFocused()
}

EmailNote_WaitUntilSubjectFocused(subEl, timeoutMs := 1500) {
    deadline := A_TickCount + timeoutMs
    lastFocusTry := 0
    while (A_TickCount < deadline) {
        if EmailNote_IsSubjectFocused()
            return true
        if ((A_TickCount - lastFocusTry) >= 200) {
            EmailNote_FocusSubjectField(subEl)
            lastFocusTry := A_TickCount
        }
        Sleep 40
    }
    return EmailNote_IsSubjectFocused()
}

EmailNote_IsToFocused() {
    try {
        focused := UIA.GetFocusedElement()
        if (!focused)
            return false
        n := ""
        try n := focused.Name
        if (n != "" && (InStr(n, "To") || InStr(n, "Destinat") || InStr(n, "recipient")))
            return true
    } catch {
    }
    return false
}

EmailNote_AddOneRecipient(searchRoot, email, isGmail, isFirst) {
    EmailNote_FocusToField(searchRoot, isGmail, isFirst)
    SendText(email)
    if (isGmail)
        SendText(",")
    else
        SendText(";")
    if EmailNote_WaitUntilToHasAddress(searchRoot, isGmail, email)
        return true
    Send("{Enter}")
    if EmailNote_WaitUntilToHasAddress(searchRoot, isGmail, email)
        return true
    ; Gmail ComboBox often never exposes a named child; caret still in To means the address was typed.
    return EmailNote_IsToFocused()
}

EmailNote_FillCompose(hwnd, subjectText := "", isGmail := false) {
    if !(hwnd is Integer) || hwnd <= 0
        return false
    WinActivate("ahk_id " hwnd)
    if (!WinWaitActive("ahk_id " hwnd, , 2))
        return false
    root := EmailNote_UiaRoot(hwnd)
    if !root
        return false
    searchRoot := root
    if (isGmail) {
        compose := EmailNote_FindGmailComposeContainer(root)
        if compose
            searchRoot := compose
    }
    toEl := EmailNote_FindToField(searchRoot, isGmail)
    if !toEl {
        ShowCenteredOverlay_Utils("❌ Email note: To field not found", 2500, BANNER_ACCENT_ERROR)
        return false
    }
    isFirst := true
    for email in EmailNote_TargetEmails() {
        if !EmailNote_AddOneRecipient(searchRoot, email, isGmail, isFirst) {
            ShowCenteredOverlay_Utils("❌ Email note: Recipient was not added as a chip", 2500, BANNER_ACCENT_ERROR)
            return false
        }
        isFirst := false
    }
    subEl := EmailNote_FindSubjectField(searchRoot, isGmail)
    if !subEl {
        ShowCenteredOverlay_Utils("❌ Email note: Could not focus Subject", 2500, BANNER_ACCENT_ERROR)
        return false
    }
    if !EmailNote_WaitUntilSubjectFocused(subEl) {
        ShowCenteredOverlay_Utils("❌ Email note: Could not focus Subject", 2500, BANNER_ACCENT_ERROR)
        return false
    }
    EmailNote_TypeSubjectText(subjectText)
    return true
}

EmailNote_FillGmailCompose(hwnd, subjectText := "") {
    return EmailNote_FillCompose(hwnd, subjectText, true)
}

EmailNote_CreateGmail(subjectText := "") {
    composeHwnd := 0
    try {
        StandardLoadingBar_Show("⏳ Opening compose...", BANNER_ACCENT_INTERMEDIATE, {
            fontSize: 17,
            trackActiveMonitor: true
        })
        composeHwnd := EmailNote_FindReadyGmailComposeHwnd()
        if (!(composeHwnd is Integer) || composeHwnd <= 0) {
            StandardLoadingBar_Update("⏳ Activating Gmail...")
            if (!EmailNote_EnsureGmailActive())
                return
            inboxHwnd := WinExist("A")
            StandardLoadingBar_Update("⏳ Waiting until Gmail inbox is ready...")
            inboxHwnd := EmailNote_WaitUntilGmailInboxReady(inboxHwnd)
            if (!(inboxHwnd is Integer) || inboxHwnd <= 0) {
                ShowCenteredOverlay_Utils("❌ Email note: Gmail inbox did not become ready", 2500, BANNER_ACCENT_ERROR)
                return
            }
            StandardLoadingBar_Update("⏳ Opening compose...")
            EmailNote_TriggerGmailCompose(inboxHwnd)
            composeHwnd := EmailNote_WaitUntilGmailComposeReady(inboxHwnd)
        }
        if (!(composeHwnd is Integer) || composeHwnd <= 0) {
            ShowCenteredOverlay_Utils("❌ Email note: Gmail compose did not become ready", 2500, BANNER_ACCENT_ERROR)
            return
        }
    } catch Error as e {
        ShowCenteredOverlay_Utils("❌ Email note: " . e.Message, 2500, BANNER_ACCENT_ERROR)
        return
    } finally {
        StandardLoadingBar_Hide(0)
    }
    if (!(composeHwnd is Integer) || composeHwnd <= 0)
        return
    WinActivate("ahk_id " composeHwnd)
    WinWaitActive("ahk_id " composeHwnd, , 2)
    EmailNote_FillGmailCompose(composeHwnd, subjectText)
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

EmailNote_FocusOutlookSubjectIfNeeded() {
    hwnd := WinExist("A")
    root := EmailNote_UiaRoot(hwnd)
    subEl := EmailNote_FindOutlookSubjectField(root)
    return EmailNote_FocusSubjectField(subEl)
}
