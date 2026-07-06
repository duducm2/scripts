; =============================================================================
; Shift keys module: outlook_helpers_01.ahk
; Outlook helper functions (part 1)
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

;-------------------------------------------------------------------
; Outlook Shortcuts
;-------------------------------------------------------------------

IsOutlookMessageActive() {
    return (WinActive("ahk_exe OUTLOOK.EXE") || WinActive("ahk_exe olk.exe"))
    && RegExMatch(WinGetTitle("A"), "i) - Message \(")
}

IsOutlookAppointmentActive() {
    ; Classic Outlook inspector windows use titles like " - Appointment/Meeting/Event".
    ; New Outlook editors often use titles like "New event - Outlook" and run under OUTLOOK.EXE or olk.exe.
    if !(WinActive("ahk_exe OUTLOOK.EXE") || WinActive("ahk_exe olk.exe"))
        return false

    t := ""
    try t := WinGetTitle("A")
    if RegExMatch(t, "i)(Appointment|Meeting|Event)")
        return true

    ; New Outlook: detect by UIA presence of the title field.
    if IsNewOutlookActive() {
        try {
            root := UIA.ElementFromHandle(WinExist("A"))
            if root.FindFirst({ Name: "Add title", ControlType: "Edit" })
                return true
            if root.FindFirst({ Name: "Add title", Type: 50004 })
                return true
            if root.FindFirst({ AutomationId: "4100" })
                return true
        } catch {
        }
    }

    return false
}

IsOutlookReminderActive() {
    return (WinActive("ahk_exe OUTLOOK.EXE") || WinActive("ahk_exe olk.exe"))
    && RegExMatch(WinGetTitle("A"), "i)Reminders?")
}

IsOutlookComposeActive() {
    ; New Outlook compose runs inside the main window and doesn't match the classic " - Message (" title.
    if !IsNewOutlookActive()
        return false
    try {
        root := UIA.ElementFromHandle(WinExist("A"))
        ; Prefer compose-only anchors seen in outlook-mail.md compose capture.
        if root.FindFirst({ AutomationId: "popoutCompose" })
            return true
        if root.FindFirst({ AutomationId: "discardCompose" })
            return true
        if root.FindFirst({ AutomationId: "splitButton-ram0__primaryActionButton" }) ; Send
            return true
        ; Fallback: presence of the compose Subject edit (MSG_*_SUBJECT) + Message body edit
        if root.FindFirst({ Name: "Subject", ControlType: "Edit" }) && root.FindFirst({ Name: "Message body",
            ControlType: "Edit" })
            return true
    } catch {
    }
    return false
}

IsOutlookMainActive() {
    if !(WinActive("ahk_exe OUTLOOK.EXE") || WinActive("ahk_exe olk.exe"))
        return false
    t := ""
    cls := ""
    try t := WinGetTitle("A")
    try cls := WinGetClass("A")
    ; Exclude inspectors and reminders
    if RegExMatch(t, "i) - Message \(")
        return false
    if RegExMatch(t, "i)(Appointment|Meeting|Event)")
        return false
    if RegExMatch(t, "i)^New event")
        return false
    if RegExMatch(t, "i)Reminder")
        return false
    ; Prefer the New Outlook shell window.
    if (cls != "" && InStr(cls, "Outlook Host"))
        return true
    if (t != "" && InStr(t, " - Outlook"))
        return true
    return true
}

Outlook_ActivateMainWindow() {
    ; Bring the main Outlook shell (Mail/Calendar) to front.
    try {
        wins := WinGetList("ahk_class Outlook Host")
        for hwnd in wins {
            t := ""
            try t := WinGetTitle("ahk_id " hwnd)
            if RegExMatch(t, "i)^(Mail|Calendar) - .* - Outlook") {
                try WinActivate("ahk_id " hwnd)
                try WinWaitActive("ahk_id " hwnd, , 1)
                return hwnd
            }
        }
    } catch {
    }
    ; Fallback: best-effort activate any Outlook process window.
    try {
        if WinExist("ahk_exe olk.exe")
            return WinActivate("ahk_exe olk.exe")
    } catch {
    }
    try {
        if WinExist("ahk_exe OUTLOOK.EXE")
            return WinActivate("ahk_exe OUTLOOK.EXE")
    } catch {
    }
    return 0
}

Outlook_FocusMainSearch() {
    Outlook_ActivateMainWindow()
    return OutlookFocusFirst([{ AutomationId: "topSearchInput", ControlType: "ComboBox" }, { AutomationId: "topSearchInput" }])
}

Outlook_SwitchToMail() {
    Outlook_ActivateMainWindow()
    ; Left-rail toggle (see outlook-mail.md): Chromium tree via OutlookMail_* â€” not OutlookClickFirst (ElementFromHandle misses WebView2).
    return OutlookMail_ClickFirst([{ AutomationId: "ddea774c-382b-47d7-aab5-adc2139a802b", ControlType: "Button" }, { Name: "Mail",
        ControlType: "Button" }, { Name: "Mail", Type: 50000 }
    ])
}

Outlook_SwitchToCalendar() {
    Outlook_ActivateMainWindow()
    return OutlookMail_ClickFirst([{ AutomationId: "8cbeb86f-83e1-43b5-aaba-cd3514322f0b", ControlType: "Button" }, { Name: "Calendar",
        ControlType: "Button" }, { Name: "Calendar", Type: 50000 }
    ])
}

; Calendar toolbar: "Go to today April 6, 2026" (dynamic date; caledar.md). WebView2 â€” use Chromium root.
; Loading bar during UIA work: docs/standard_information_display.md (Show â†’ Update â†’ Hide).
OutlookCalendar_ClickGoToToday() {
    try {
        StandardLoadingBar_Show("â³ Outlook: Go to todayâ€¦", BANNER_ACCENT_INTERMEDIATE, { passive: false,
            centerOnHwnd: 0,
            textWidth: 480, fontSize: 17 })
    } catch {
    }
    ok := false
    try {
        Outlook_ActivateMainWindow()
        try StandardLoadingBar_Update("â³ Outlook: Opening Calendarâ€¦")
        if !Outlook_SwitchToCalendar()
            return false
        Sleep 150
        try StandardLoadingBar_Update("â³ Outlook: Finding Go to todayâ€¦")
        el := OutlookMail_FindFirst([{ Name: "Go to today", matchmode: "Substring", ControlType: "Button" }, { Name: "Go to today",
            matchmode: "Substring", Type: 50000 }, { Name: "Today", ControlType: "Button" }, { Name: "Today", matchmode: "Substring",
                ControlType: "Button" }
        ])
        if !el
            return false
        try el.ScrollIntoView()
        Sleep 40
        try el.SetFocus()
        Sleep 50
        try el.Click()
        catch Error {
            try el.Invoke()
            catch Error {
                return false
            }
        }
        ok := true
    } finally {
        try StandardLoadingBar_Hide(0)
        catch {
        }
    }
    return ok
}
