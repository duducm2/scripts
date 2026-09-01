; =============================================================================
; Shift keys module: hotif_outlook_main.ahk
; Outlook main window hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

#HotIf IsOutlookMainActive()

; -------------------------------------------------------------------
; Outlook main window (New Outlook) overflow layer: Ctrl+Alt+…
; -------------------------------------------------------------------

^!f:: {  ; Focus Search
    if !Outlook_FocusMainSearch()
        ShowCenteredOverlay_Utils("❌ Outlook: Search not found", 1200, BANNER_ACCENT_ERROR)
}

^!m:: {  ; Switch to Mail
    if !Outlook_SwitchToMail()
        ShowCenteredOverlay_Utils("❌ Outlook: Mail not found", 1200, BANNER_ACCENT_ERROR)
}

^!g:: {  ; Switch to Calendar
    if !Outlook_EnsureSwitchToCalendar()
        ShowCenteredOverlay_Utils("❌ Outlook: Calendar not found", 1200, BANNER_ACCENT_ERROR)
}

^!l:: {  ; Focus message list
    if !Outlook_FocusMailMessageList()
        ShowCenteredOverlay_Utils("❌ Outlook: Message list not found", 1200, BANNER_ACCENT_ERROR)
}

^!p:: {  ; Focus reading pane
    if !Outlook_FocusMailReadingPane()
        ShowCenteredOverlay_Utils("❌ Outlook: Reading pane not found", 1200, BANNER_ACCENT_ERROR)
}

; Mail triage (Reading Pane / Ribbon)
^!r:: {  ; Reply
    if OutlookMail_ClickReadingPaneCommand("Reply")
        return
    OutlookMail_EnsureHomeTab()
    if !OutlookClickFirst([{ Name: "Reply", ControlType: "Button" }, { Name: "Reply",
        ControlType: "MenuItem" }])
        ShowCenteredOverlay_Utils("❌ Outlook: Reply not found", 1200, BANNER_ACCENT_ERROR)
}

^!a:: {  ; Reply all
    if OutlookMail_ClickReadingPaneCommand("Reply all")
        return
    OutlookMail_EnsureHomeTab()
    if !OutlookClickFirst([{ Name: "Reply all", ControlType: "Button" }, { Name: "Reply all",
        ControlType: "MenuItem" }])
        ShowCenteredOverlay_Utils("❌ Outlook: Reply all not found", 1200, BANNER_ACCENT_ERROR)
}

^!w:: {  ; Forward
    if OutlookMail_ClickReadingPaneCommand("Forward")
        return
    OutlookMail_EnsureHomeTab()
    if !OutlookClickFirst([{ Name: "Forward", ControlType: "Button" }, { Name: "Forward",
        ControlType: "MenuItem" }])
        ShowCenteredOverlay_Utils("❌ Outlook: Forward not found", 1200, BANNER_ACCENT_ERROR)
}

^!d:: {  ; Delete
    OutlookMail_EnsureHomeTab()
    if !OutlookClickFirst([{ AutomationId: "519", ControlType: "Button" }, { Name: "Delete", ControlType: "Button" }])
        ShowCenteredOverlay_Utils("❌ Outlook: Delete not found", 1200, BANNER_ACCENT_ERROR)
}

^!e:: {  ; Archive
    OutlookMail_EnsureHomeTab()
    if !OutlookClickFirst([{ AutomationId: "505", ControlType: "Button" }, { Name: "Archive", ControlType: "Button" }])
        ShowCenteredOverlay_Utils("❌ Outlook: Archive not found", 1200, BANNER_ACCENT_ERROR)
}

^!u:: {  ; Read/Unread toggle
    OutlookMail_EnsureHomeTab()
    if !OutlookClickFirst([{ AutomationId: "552", ControlType: "Button" }, { Name: "Read / Unread", ControlType: "Button" }])
        ShowCenteredOverlay_Utils("❌ Outlook: Read/Unread not found", 1200, BANNER_ACCENT_ERROR)
}

^!c:: {  ; Categorize
    OutlookMail_EnsureHomeTab()
    if !OutlookClickFirst([{ AutomationId: "509", ControlType: "Button" }, { Name: "Categorize", ControlType: "Button" }])
        ShowCenteredOverlay_Utils("❌ Outlook: Categorize not found", 1200, BANNER_ACCENT_ERROR)
}

^!v:: {  ; Move
    OutlookMail_EnsureHomeTab()
    if !OutlookClickFirst([{ AutomationId: "540", ControlType: "Button" }, { Name: "Move", ControlType: "Button" }])
        ShowCenteredOverlay_Utils("❌ Outlook: Move not found", 1200, BANNER_ACCENT_ERROR)
    else {
        Sleep 80
        Send "{Down 2}"
    }
}

^!i:: {  ; Mail Filter menu
    if !OutlookClickFirst([{ AutomationId: "mailListFilterMenu", ControlType: "Button" }, { Name: "Filter", ControlType: "Button" }])
        ShowCenteredOverlay_Utils("❌ Outlook: Filter not found", 1200, BANNER_ACCENT_ERROR)
}

^!s:: {  ; Mail Sort menu
    if !OutlookClickFirst([{ AutomationId: "mailListSortMenu", ControlType: "Button" }, { Name: "Sorted", matchmode: "Substring",
        ControlType: "Button" }])
        ShowCenteredOverlay_Utils("❌ Outlook: Sort not found", 1200, BANNER_ACCENT_ERROR)
}

; Calendar (main view)
^!n:: {  ; New item (Mail: new message, Calendar: new event)
    Outlook_ActivateMainWindow()
    ; Calendar capture exposes "New event".
    if OutlookClickFirst([{ Name: "New event", matchmode: "Substring", ControlType: "Button" }])
        return
    ; Mail: fall back to built-in new message.
    Send "^n"
}

^!t:: {  ; Today (Calendar) — toolbar "Go to today …" (Chromium; not ElementFromHandle-only)
    if !OutlookCalendar_ClickGoToToday()
        ShowCenteredOverlay_Utils("❌ Outlook: Today not found", 1200, BANNER_ACCENT_ERROR)
}

; Shift + G : Send to General - General
+G::
{
    if IsNewOutlookActive() {
        ; New Outlook: prefer the Quick Step buttons (stable IDs from outlook-mail.md).
        Outlook_ActivateMainWindow()
        OutlookMail_EnsureHomeTab()
        if OutlookClickFirst([{ AutomationId: "c46846eb-0853-7b70-b484-4d7f31f5d9db", ControlType: "RadioButton" }, ; Move to General
        { AutomationId: "c46846eb-0853-7b70-b484-4d7f31f5d9db" }, { Name: "Move to General", ControlType: "RadioButton" }, { Name: "Move to General",
            matchmode: "Substring" }, { Name: "Move to general", matchmode: "Substring" }, { Name: "Move to Gerais",
                matchmode: "Substring" }
        ]) {
            Sleep 120
            Outlook_FocusMailMessageList(true)
            return
        }
    }
    Send "!5"
    Send "O"
    Send "{Home}"
    Send "00"
    Send "{Enter}"
}

; Shift + N : Send to Newsletter - Newsletter
+N::
{
    if IsNewOutlookActive() {
        ; New Outlook: prefer the Quick Step buttons (stable IDs from outlook-mail.md).
        Outlook_ActivateMainWindow()
        OutlookMail_EnsureHomeTab()
        if OutlookClickFirst([{ AutomationId: "91476b25-0fb7-4460-f695-8905582291db", ControlType: "RadioButton" }, ; Move to Newsletter
        { AutomationId: "91476b25-0fb7-4460-f695-8905582291db" }, { Name: "Move to Newsletter", ControlType: "RadioButton" }, { Name: "Move to Newsletter",
            matchmode: "Substring" }, { Name: "Move to newsletter", matchmode: "Substring" }, { Name: "newsletter",
                matchmode: "Substring", ControlType: "RadioButton" }
        ]) {
            Sleep 120
            Outlook_FocusMailMessageList(true)
            return
        }
    }
    Send "!5"
    Send "O"
    Send "{Home}"
    Send "01"
    Send "{Enter}"
}

; Shift + I : Go to Inbox - Inbox
+I::
{
    if IsNewOutlookActive() {
        Outlook_ActivateMainWindow()
        ; New Outlook: click Inbox in the Navigation pane (unique scope).
        if OutlookMail_ClickInboxFolder()
            return
    }
    Send "{Alt}"
    Sleep 60
    Send "6"
    Sleep 80
    Send "^{Home}"
    Sleep 100
    Send "i"
    Sleep 50
    Send "n"
    Sleep 50
    Send "{Enter}"
}

; Shift + J : Jump to first mail and select it
+J::
{
    if IsNewOutlookActive() {
        Outlook_ActivateMainWindow()
        ; Keep behavior mail-centric: if user is in Calendar, switch first.
        Outlook_SwitchToMail()
        if Outlook_FocusMailMessageList(true)
            return
    }

    ; Classic / fallback: try UIA first-item selection (with non-message filter), then keyboard skip-header path.
    if Outlook_FocusMailMessageList(true)
        return
    if Outlook_FocusMailMessageList() {
        Send "{Home}"
        Sleep 40
        if !Outlook_MailList_SkipDrawerHeadersByKeyboard()
            ShowCenteredOverlay_Utils("❌ Outlook: Could not focus a message row", 1200, BANNER_ACCENT_ERROR)
    } else
        ShowCenteredOverlay_Utils("❌ Outlook: Message list not found", 1200, BANNER_ACCENT_ERROR)
}

; Shift + H : Toggle high navigation pane — ribbon Hide/Show folder pane
+H:: {
    if !IsNewOutlookActive()
        return
    Outlook_ActivateMainWindow()
    if !OutlookMail_ToggleHighNavigationPane()
        ShowCenteredOverlay_Utils("❌ Outlook: Navigation pane toggle not found", 1200, BANNER_ACCENT_ERROR)
}

; Shift + S : Subject / Title - Subject
+S:: {
    if FocusOutlookField({ AutomationId: "4101" }) ; Subject
        return
    if FocusOutlookField({ Name: "Subject", ControlType: "Edit" })
        return
}

; Shift + T : Required / To - To (meeting request in reading pane: Tentative via More options)
+T:: {
    if IsOutlookMeetingRequestReadingPaneActive() {
        if !OutlookMeeting_ClickMoreOptionsThen("Tentative")
            ShowCenteredOverlay_Utils("❌ Outlook: Tentative not found", 1200, BANNER_ACCENT_ERROR)
        return
    }
    ; New Outlook compose: “To:” row is a Group (AutomationId 134) that may be collapsed/hidden.
    if IsNewOutlookActive() && IsOutlookComposeActive() {
        ; #region agent log
        OC_STLog(msg, data := "{}", hypo := "OC_ST") {
            try {
                line := "{"
                    . '"sessionId":"b96502",'
                    . '"runId":"shiftT",'
                    . '"hypothesisId":"' hypo '",'
                    . '"timestamp":' A_TickCount ','
                    . '"location":"Shift keys.ahk:+T(compose)",'
                    . '"message":"' StrReplace(msg, '"', '\"') '",'
                    . '"data":' data
                    . "}"
                FileAppend(line "`n", "debug-b96502.log", "UTF-8")
            } catch {
            }
        }
        ; #endregion

        try {
            hwnd := WinExist("A")
            t := WinGetTitle("A")
            OC_STLog("compose_gate_passed", '{"hwnd":' hwnd ',"title":"' StrReplace(SubStr(t, 1, 120), '"', '\"') '"}',
            "OC_ST_A")
        } catch {
        }

        ; Prefer recipient focus flow (reactive UI): To row -> recipient entity group -> inner field.
        ok := OutlookCompose_FocusToRecipientsField()
        try OC_STLog("after_focus_flow", '{"ok":' (ok ? 1 : 0) '}', "OC_ST_B")
        catch {
        }
        if ok
            return

        ; Fallback experiment (logged): select Bcc then Shift+Tab once.
        try {
            bccOk := OutlookClickFirst([{ Name: "Bcc", matchmode: "Substring", ControlType: "Button" }, { Name: "Bcc",
                matchmode: "Substring" }, { Name: "Show Bcc", matchmode: "Substring", ControlType: "Button" }
            ])
            OC_STLog("bcc_click", '{"ok":' (bccOk ? 1 : 0) '}', "OC_ST_C")
            if bccOk {
                Send "{Tab}"
                Sleep 40
                fe := UIA.GetFocusedElement()
                fn := "", fa := "", ft := ""
                try fn := fe.Name
                try fa := fe.AutomationId
                try ft := fe.Type
                OC_STLog("focused_after_bcc_shift_tab", '{"name":"' StrReplace(SubStr(fn, 1, 80), '"', '\"') '","automationId":"' StrReplace(
                    SubStr(fa, 1, 80), '"', '\"') '","type":' (ft = "" ? -1 : ft) '}', "OC_ST_D")
                return
            }
        } catch {
        }

        ; Fallback: click the To row only.
        if OutlookClickFirst([{ AutomationId: "134", ControlType: "Group" }, { AutomationId: "134" }])
            return
        ; Fallback: any element whose name begins with “To:”.
        if OutlookClickFirst([{ Name: "To:", matchmode: "Substring" }, { Name: "To", matchmode: "Substring",
            ControlType: "Group" }])
            return
    }
    if FocusOutlookField({ AutomationId: "4109" }) ; Required
        return
    if FocusOutlookField({ Name: "Required", ControlType: "Edit" })
        return
    if FocusOutlookField({ AutomationId: "4117" }) ; To
        return
    if FocusOutlookField({ Name: "To", ControlType: "Edit" })
        return
}

; Shift + B : Subject -> Body - Body
+B:: {
    if FocusOutlookField({ AutomationId: "4101" }) {
        Sleep 50
        Send "{Tab}"
        return
    }
    if FocusOutlookField({ Name: "Subject", ControlType: "Edit" }) {
        Sleep 50
        Send "{Tab}"
        return
    }
}

; Shift + F : Toggle Focused / Other - Focused
+F:: {                                  ; toggle Focused / Other
    static nextOutlookButton := "Other"

    try {
        btn := OutlookFindFirst([{ Name: nextOutlookButton, ControlType: "TabItem" }, { Name: nextOutlookButton,
            ControlType: "Button" }, { Name: nextOutlookButton, Type: "Button" }
        ])

        if btn {
            btn.Click()
            nextOutlookButton := (nextOutlookButton = "Other")
                ? "Focused" : "Other"
        } else {
            MsgBox("Couldn't find '" nextOutlookButton "'.", "Button not found", "IconX")
        }

    } catch Error as err {              ; â† **only this form**
        ShowErr(err)
    }
}

; Shift+W : Calendar [W]eek view
+W:: {
    try {
        if IsNewOutlookActive() {
            if OutlookClickFirst([{ AutomationId: "2519", ControlType: "Button" }, { Name: "Week", ControlType: "Button" }])
                return
        }
        if !ClickOutlookByIdThenNameClass("WeeklyView", "Week", "NetUIRibbonButton", 50000)
            Send "^!3"
    } catch {
        Send "^!3"
    }
}

; Shift+O : Calendar m[O]nth view
+O:: {
    try {
        if IsNewOutlookActive() {
            if OutlookClickFirst([{ AutomationId: "2505", ControlType: "Button" }, { Name: "Month", ControlType: "Button" }])
                return
        }
        if !ClickOutlookByIdThenNameClass("MonthlyView", "Month", "NetUIRibbonButton", 50000)
            Send "^!4"
    } catch {
        Send "^!4"
    }
}

; Shift + P : Pop Out current item - Pop Out
+P:: {
    try {
        if !ClickOutlookByIdThenNameClass("", "Pop Out", "", 50000) {
            MsgBox("Couldn't find 'Pop Out' button.", "Outlook Pop Out", "IconX")
        }
    } catch Error as err {
        ShowErr(err)
    }
}

; -------------------------------------------------------------------
; Focus helpers â€" reuse for any field you need
; -------------------------------------------------------------------
FocusOutlookField(criteria) {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)
        ctrl := root.FindFirst(criteria)
        if ctrl {
            ctrl.SetFocus()
            return true
        }
    } catch Error {
    }
    return false
}

; -------------------------------------------------------------------
; Click helper â€" try AutomationId first, then Name+ClassName
; -------------------------------------------------------------------
ClickOutlookByIdThenNameClass(automationId, name, className, controlType := "") {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        if (automationId) {
            el := root.FindFirst({ AutomationId: automationId })
            if (el) {
                el.SetFocus()
                Sleep 50
                el.Click()
                return true
            }
        }

        crit := { Name: name }
        if (className)
            crit.ClassName := className
        if (controlType)
            crit.ControlType := controlType

        el := root.FindFirst(crit)
        if (el) {
            el.SetFocus()
            Sleep 50
            el.Click()
            return true
        }
    } catch Error as err {
        ShowErr(err)
    }
    return false
}

; -------------------------------------------------------------------
; General helper â€" visually confirm focus on the selected element
; Sends Down then Up to force a visible focus cue
; -------------------------------------------------------------------
EnsureFocus() {
    Send "{Down}"
    Send "{Up}"
}

; Helper: Select the first pinned item in Explorer sidebar (Navigation Pane)
; Global so it can be reused by both Explorer and File Dialog contexts
SelectExplorerSidebarFirstPinned() {
    try {
        explorerEl := UIA.ElementFromHandle(WinExist("A"))

        ; Look for the navigation pane (sidebar) - it's typically a Tree control
        navPane := explorerEl.FindFirst({ Type: "Tree" })

        if (navPane) {
            ; If in work environment, prefer selecting the Home tree item directly
            try {
                global IS_WORK_ENVIRONMENT
                if (IS_WORK_ENVIRONMENT) {
                    homeItem := navPane.FindFirst({ Type: "TreeItem", Name: "Home" })
                    if (homeItem) {
                        homeItem.ScrollIntoView()
                        homeItem.Select()    ; select only, no click
                        homeItem.SetFocus()
                        EnsureFocus()
                        return true
                    }
                }
            } catch Error {
                ; ignore and fallback to previous logic
            }
            ; Define the keywords to search for pinned items
            pinnedKeywords := ["fixo", "pinned", "pin", "fixado", "fixada", "fixar", "preso"]

            ; Search for the first TreeItem that contains any of the pinned keywords
            firstPinnedItem := unset
            for keyword in pinnedKeywords {
                firstPinnedItem := navPane.FindFirst({ Type: "TreeItem", Name: keyword, matchmode: "Substring" })
                if (firstPinnedItem)
                    break
            }

            ; If no pinned item found by keywords, try to find Desktop by first letter
            ; Portuguese: "Área de Trabalho" starts with "Á" or "a"
            ; English: "Desktop" starts with "D" or "d"
            if (!firstPinnedItem) {
                allTreeItems := navPane.FindAll({ Type: "TreeItem" })
                for item in allTreeItems {
                    try {
                        itemName := item.Name
                        ; Check if name starts with "a" or "Á" (Portuguese Desktop) or "d" or "D" (English Desktop)
                        ; Case-insensitive check
                        firstChar := SubStr(itemName, 1, 1)
                        if (firstChar = "a" || firstChar = "A" || firstChar = "Á" || firstChar = "á" ||
                            firstChar = "d" || firstChar = "D") {
                            ; Additional check: must be Desktop-related (not just any item starting with a/d)
                            if (InStr(itemName, "Desktop", false) || InStr(itemName, "Área de Trabalho", false) ||
                            InStr(itemName, "Trabalho", false)) {
                                firstPinnedItem := item
                                break
                            }
                        }
                    } catch {
                        ; Skip items without names
                    }
                }
            }

            if (firstPinnedItem) {
                firstPinnedItem.ScrollIntoView()
                firstPinnedItem.Select()
                firstPinnedItem.SetFocus()
                EnsureFocus()
                return true
            }

            ; If we didn't find a pinned item, at least focus the tree and press Home
            navPane.SetFocus()
            Sleep 100
            Send "{Home}"
            EnsureFocus()
            return false
        }
    } catch Error {
        ; swallow and continue to fallback
    }

    ; Robust fallback â€" cycle through panes up to 6 times to reach navigation, then Home
    loop 6 {
        Send "{F6}"
        Sleep 120
        try {
            explorerEl := UIA.ElementFromHandle(WinExist("A"))
            navPane := explorerEl.FindFirst({ Type: "Tree" })
            if (navPane && navPane.HasKeyboardFocus) {
                Send "{Home}"
                EnsureFocus()
                return false
            }
        } catch Error {
        }
    }
    ; Last resort â€" send Home anyway
    Send "{Home}"
    EnsureFocus()
    return false
}

; Shift + K : Send Shift+F6
+K:: Send "+{F6}"

; Shift + L : Send F6
+L:: Send "{F6}"

; Shift + M : Toggle Mail / Calendar - Mail/Calendar
+M:: {
    try {
        if IsNewOutlookActive() {
            if Outlook_ToggleMailCalendarRail()
                return
            ; Fallback if rail toggles not found / no toggle pattern
            t := WinGetTitle("A")
            if RegExMatch(t, "i)Calendar") {
                if Outlook_SwitchToMail()
                    return
            } else {
                if Outlook_SwitchToCalendar()
                    return
            }
        }
        root := UIA.ElementFromHandle(WinExist("A"))

        ; Find Mail and Calendar list items
        mailItem := root.FindFirst({ Name: "Mail", Type: "50007" })
        if !mailItem {
            mailItem := root.FindFirst({ Name: "Mail", ClassName: "NetUIListViewItem" })
        }

        calendarItem := root.FindFirst({ Name: "Calendar", Type: "50007" })
        if !calendarItem {
            calendarItem := root.FindFirst({ Name: "Calendar", ClassName: "NetUIListViewItem" })
        }

        ; Check which is selected and toggle
        if (mailItem && calendarItem) {
            try {
                isMailSelected := mailItem.IsSelected
                isCalendarSelected := calendarItem.IsSelected

                if (isMailSelected) {
                    calendarItem.SetFocus()
                    Sleep 50
                    calendarItem.Click()
                } else {
                    mailItem.SetFocus()
                    Sleep 50
                    mailItem.Click()
                }
            } catch Error as err {
                ; Fallback: if pattern check fails, try clicking Calendar
                calendarItem.SetFocus()
                Sleep 50
                calendarItem.Click()
            }
        } else {
            MsgBox "Could not find Mail or Calendar items.", "Outlook Toggle", "IconX"
        }
    } catch Error as err {
        MsgBox "Error toggling Mail/Calendar:`n" err.Message, "Outlook Toggle", "IconX"
    }
}

; Meeting request - main Mail window reading pane only (not popped-out invitation window).
#HotIf IsOutlookMainActive() && IsOutlookMeetingRequestReadingPaneActive()

; Accept / Follow (header row). Alt+F = Follow (Shift+F stays Focused / Other in the base hotkey above).
+A:: {
    if !OutlookMeeting_ClickAccept()
        ShowCenteredOverlay_Utils("❌ Outlook: Accept the meeting not found", 1200, BANNER_ACCENT_ERROR)
}

!f:: {
    if !OutlookMeeting_ClickFollow()
        ShowCenteredOverlay_Utils("❌ Outlook: Follow not found", 1200, BANNER_ACCENT_ERROR)
}

+D:: {
    if !OutlookMeeting_ConfirmDecline()
        return
    if !OutlookMeeting_DeclineSilently()
        ShowCenteredOverlay_Utils("❌ Outlook: Decline silently not found", 1200, BANNER_ACCENT_ERROR)
}

; Message inspector-specific hotkeys (Subject/To/Body)
#HotIf IsOutlookMessageActive()

; Shift + S : Subject / Title - Subject
+S:: {
    if FocusOutlookField({ AutomationId: "4101" }) ; Subject
        return
    if FocusOutlookField({ Name: "Subject", ControlType: "Edit" })
        return
}

; Shift + T : Required / To - To
+T:: {
    if FocusOutlookField({ AutomationId: "4109" }) ; Required
        return
    if FocusOutlookField({ Name: "Required", ControlType: "Edit" })
        return
    if FocusOutlookField({ AutomationId: "4117" }) ; To
        return
    if FocusOutlookField({ Name: "To", ControlType: "Edit" })
        return
}

; Shift + B : Body (Subject -> Body) - Body
+B:: {
    if FocusOutlookField({ AutomationId: "4101" }) {
        Sleep 50
        Send "{Tab}"
        return
    }
    if FocusOutlookField({ Name: "Subject", ControlType: "Edit" }) {
        Sleep 50
        Send "{Tab}"
        return
    }
}

; Popped-out meeting invitation (same Accept/Follow/Tentative as reading pane; overrides generic message Shift+T / Alt+F when both match).
#HotIf IsOutlookMeetingInvitationPopOutActive()

+A:: {
    if !OutlookMeeting_ClickAccept()
        ShowCenteredOverlay_Utils("❌ Outlook: Accept the meeting not found", 1200, BANNER_ACCENT_ERROR)
}

!f:: {
    if !OutlookMeeting_ClickFollow()
        ShowCenteredOverlay_Utils("❌ Outlook: Follow not found", 1200, BANNER_ACCENT_ERROR)
}

+T:: {
    if !OutlookMeeting_ClickMoreOptionsThen("Tentative")
        ShowCenteredOverlay_Utils("❌ Outlook: Tentative not found", 1200, BANNER_ACCENT_ERROR)
}

+D:: {
    if !OutlookMeeting_ConfirmDecline()
        return
    if !OutlookMeeting_DeclineSilently()
        ShowCenteredOverlay_Utils("❌ Outlook: Decline silently not found", 1200, BANNER_ACCENT_ERROR)
}

; Canceled meeting — Button "Remove event" (Shift+E); no confirmation (see outlook-remove-evet.md).
#HotIf IsOutlookMainActive() && IsOutlookRemoveEventReadingPaneActive()

+E:: {
    if !OutlookMeeting_ClickRemoveEvent()
        ShowCenteredOverlay_Utils("❌ Outlook: Remove event not found", 1200, BANNER_ACCENT_ERROR)
}

#HotIf IsOutlookRemoveEventPopOutActive()

+E:: {
    if !OutlookMeeting_ClickRemoveEvent()
        ShowCenteredOverlay_Utils("❌ Outlook: Remove event not found", 1200, BANNER_ACCENT_ERROR)
}

#HotIf