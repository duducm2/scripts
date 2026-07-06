; =============================================================================
; Shift keys module: outlook_helpers_02.ahk
; Outlook helper functions (part 2)
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

; True if toggle appears "on" (left-rail app bar toggle buttons).
Outlook_RailToggleIsPressed(el) {
    if !el
        return false
    try {
        if el.IsTogglePatternAvailable
            return (el.TogglePattern.CurrentToggleState = 1)
    } catch {
    }
    try {
        return (el.GetCurrentPropertyValue(UIA.Property.ToggleToggleState) = 1)
    } catch {
    }
    return false
}

; Shift+M: switch between Mail and Calendar using rail toggle state (not window title). See outlook-mail.md left-rail-appbar.
Outlook_ToggleMailCalendarRail() {
    Outlook_ActivateMainWindow()
    root := OutlookMail_RootElement()
    if !root
        return false
    searchRoot := root
    try {
        r := root.FindFirst({ Name: "left-rail-appbar", matchmode: "Substring", ControlType: "Group" })
        if !r
            r := root.FindFirst({ Name: "left-rail-appbar", matchmode: "Substring" })
        if r
            searchRoot := r
    } catch {
    }
    mailBtn := ""
    calBtn := ""
    try mailBtn := searchRoot.FindFirst({ AutomationId: "ddea774c-382b-47d7-aab5-adc2139a802b", ControlType: "Button" })
    if !mailBtn
        try mailBtn := searchRoot.FindFirst({ Name: "Mail", ControlType: "Button" })
    try calBtn := searchRoot.FindFirst({ AutomationId: "8cbeb86f-83e1-43b5-aaba-cd3514322f0b", ControlType: "Button" })
    if !calBtn
        try calBtn := searchRoot.FindFirst({ Name: "Calendar", ControlType: "Button" })
    ; If rail group scope missed (build differences), search full Chromium root
    if !mailBtn
        try mailBtn := root.FindFirst({ AutomationId: "ddea774c-382b-47d7-aab5-adc2139a802b", ControlType: "Button" })
    if !calBtn
        try calBtn := root.FindFirst({ AutomationId: "8cbeb86f-83e1-43b5-aaba-cd3514322f0b", ControlType: "Button" })
    if !mailBtn || !calBtn
        return false

    mailOn := Outlook_RailToggleIsPressed(mailBtn)
    calOn := Outlook_RailToggleIsPressed(calBtn)
    target := ""

    if (mailOn && !calOn)
        target := calBtn
    else if (calOn && !mailOn)
        target := mailBtn
    else {
        ; Both off, both on, or no toggle pattern: infer from title
        t := WinGetTitle("A")
        if RegExMatch(t, "i)Calendar")
            target := mailBtn
        else
            target := calBtn
    }

    if !target
        return false
    try target.SetFocus()
    Sleep 50
    try target.Click()
    catch {
        try target.Invoke()
    }
    return true
}

Outlook_MailList_GetFirstListItem(root) {
    if !root
        return 0
    listEl := 0
    try {
        g := root.FindFirst({ AutomationId: "MailList" })
        if g
            listEl := g.FindFirst({ ControlType: "List" })
    } catch {
    }
    if !listEl {
        try listEl := root.FindFirst({ Name: "Message list", ControlType: "List" })
    }
    if !listEl {
        try listEl := root.FindFirst({ Name: "Message list", Type: 50008 })
    }
    if !listEl
        return 0

    ; New Outlook may expose date/header bars near the top of the list. Pick the first real
    ; message row instead of UI chrome/header artifacts.
    try {
        items := listEl.FindAll({ ControlType: "ListItem" })
        if IsObject(items) {
            for item in items {
                if !Outlook_MailList_IsNonMessageItem(item)
                    return item
            }
        }
    } catch {
    }

    try {
        items := listEl.FindAll({ Type: 50007 })
        if IsObject(items) {
            for item in items {
                if !Outlook_MailList_IsNonMessageItem(item)
                    return item
            }
        }
    } catch {
    }

    item := 0
    try item := listEl.FindFirst({ ControlType: "ListItem" })
    if !item {
        try item := listEl.FindFirst({ Type: 50007 })
    }
    if item && !Outlook_MailList_IsNonMessageItem(item)
        return item
    return 0
}

Outlook_MailList_IsNonMessageItem(item) {
    if !item
        return true
    aid := ""
    name := ""
    hasSelectionPattern := false
    try aid := item.AutomationId
    try name := Trim(item.Name)
    try hasSelectionPattern := !!item.GetCurrentPropertyValue(UIA.Property.IsSelectionItemPatternAvailable)

    ; Date group headers / list chrome can appear before the first actual mail row.
    if (aid != "" && RegExMatch(aid, "i)^groupHeader"))
        return true
    if (name != "" && RegExMatch(name, "i)^((Today|Yesterday)\b|Header action menu)$"))
        return true

    ; Tiny unnamed bars without selection pattern are not message rows.
    if (!hasSelectionPattern && StrLen(name) < 4)
        return true

    return false
}

Outlook_MailList_FocusedIsDrawerHeader() {
    try {
        fe := UIA.GetFocusedElement()
        if !fe
            return false
        name := ""
        aid := ""
        try name := Trim(fe.Name)
        try aid := fe.AutomationId
        if (aid != "" && RegExMatch(aid, "i)^groupHeader"))
            return true
        if (name != "" && RegExMatch(name, "i)^(Today|Yesterday)\b"))
            return true
    } catch {
    }
    return false
}

Outlook_MailList_FocusedIsLikelyMessageRow() {
    try {
        fe := UIA.GetFocusedElement()
        if !fe
            return false
        if Outlook_MailList_IsNonMessageItem(fe)
            return false
        ctlType := 0
        hasSelectionPattern := false
        try ctlType := fe.ControlType
        try hasSelectionPattern := !!fe.GetCurrentPropertyValue(UIA.Property.IsSelectionItemPatternAvailable)
        if (ctlType = UIA.ControlType.ListItem || ctlType = 50007)
            return true
        return hasSelectionPattern
    } catch {
    }
    return false
}

Outlook_MailList_SkipDrawerHeadersByKeyboard(maxSteps := 5) {
    loop maxSteps {
        if Outlook_MailList_FocusedIsLikelyMessageRow()
            return true
        Send "{Down}"
        Sleep 30
    }
    return Outlook_MailList_FocusedIsLikelyMessageRow()
}

Outlook_MailList_TrySelectFirstItem(root) {
    item := Outlook_MailList_GetFirstListItem(root)
    if !item
        return false
    try item.ScrollIntoView()
    catch {
    }
    Sleep 40
    try item.Select()
    catch {
        try item.Click()
        catch {
            try item.Invoke()
            catch {
                return false
            }
        }
    }
    try item.SetFocus()
    return true
}

; selectFirst: when true, select the first message row in the list (if any) for faster triage after moves.
Outlook_FocusMailMessageList(selectFirst := false) {
    Outlook_ActivateMainWindow()
    hwnd := WinExist("A")
    if !hwnd
        return false
    if selectFirst {
        try {
            root := OutlookMail_RootElementForHwnd(hwnd)
            if !root
                root := UIA.ElementFromHandle(hwnd)
            if Outlook_MailList_TrySelectFirstItem(root)
                return true
        } catch {
        }
    }
    if !OutlookFocusFirst([{ AutomationId: "Skip to message list-region" }, { Name: "Message list", matchmode: "Substring" }])
        return false

    if !selectFirst {
        try EnsureFocus()
        return true
    }

    ; Retry selection after focusing list container.
    try {
        root := OutlookMail_RootElementForHwnd(WinExist("A"))
        if !root
            root := UIA.ElementFromHandle(WinExist("A"))
        if Outlook_MailList_TrySelectFirstItem(root)
            return true
    } catch {
    }

    ; Last fallback: Home may land on Today/Yesterday header; step down until a message row is focused.
    Send "{Home}"
    Sleep 40
    return Outlook_MailList_SkipDrawerHeadersByKeyboard()
}

Outlook_FocusMailReadingPane() {
    Outlook_ActivateMainWindow()
    if OutlookFocusFirst([{ AutomationId: "Skip to message-region" }, { Name: "Reading Pane", matchmode: "Substring" }]) {
        try EnsureFocus()
        return true
    }
    return false
}

OutlookMail_ClickReadingPaneCommand(cmdName) {
    Outlook_ActivateMainWindow()
    try {
        root := OutlookMail_RootElement()
        if !root
            root := UIA.ElementFromHandle(WinExist("A"))
        pane := root.FindFirst({ AutomationId: "Skip to message-region" })
        if !pane
            return false
        el := 0
        try el := pane.FindFirst({ Name: cmdName, ControlType: "MenuItem" })
        if !el
            try el := pane.FindFirst({ Name: cmdName, ControlType: "Button" })
        if el {
            try el.SetFocus()
            Sleep 40
            try el.Click()
            catch {
                try el.Invoke()
            }
            return true
        }
    } catch {
    }
    return false
}

; Reading pane group (New Outlook WebView2 UI). Uses the active window only — no Outlook_ActivateMainWindow (see IsOutlookMeetingRequestReadingPaneActive vs pop-out invitation).
OutlookMail_GetReadingPaneElement() {
    try {
        hwnd := WinExist("A")
        if !hwnd
            return ""
        root := OutlookMail_RootElementForHwnd(hwnd)
        if root {
            pane := root.FindFirst({ AutomationId: "Skip to message-region" })
            if pane
                return pane
        }
    } catch {
    }
    try {
        root := UIA.ElementFromHandle(WinExist("A"))
        if root {
            pane := root.FindFirst({ AutomationId: "Skip to message-region" })
            if pane
                return pane
        }
    } catch {
    }
    return ""
}

; True when the main Mail/Calendar shell shows a meeting request in the reading pane (not a popped-out invitation window).
IsOutlookMeetingRequestReadingPaneActive() {
    if !IsOutlookMainActive()
        return false
    try {
        pane := OutlookMail_GetReadingPaneElement()
        if !pane
            return false
        return pane.FindFirst({ Name: "Accept the meeting", ControlType: "MenuItem" }) ? true : false
    } catch {
    }
    return false
}

; Popped-out meeting invitation: " - Message (" inspector with Accept/Decline row (same UI as reading pane, different host window).
IsOutlookMeetingInvitationPopOutActive() {
    if !(WinActive("ahk_exe OUTLOOK.EXE") || WinActive("ahk_exe olk.exe"))
        return false
    if !RegExMatch(WinGetTitle("A"), "i) - Message \(")
        return false
    try {
        root := OutlookMail_RootElementForHwnd(WinExist("A"))
        if !root
            return false
        return root.FindFirst({ Name: "Accept the meeting", ControlType: "MenuItem" }) ? true : false
    } catch {
    }
    return false
}

; Organizer canceled: primary action is Button "Remove event" (not Accept/Decline).
IsOutlookRemoveEventReadingPaneActive() {
    if !IsOutlookMainActive()
        return false
    try {
        root := OutlookMail_RootElementForHwnd(WinExist("A"))
        if !root
            return false
        return root.FindFirst({ Name: "Remove event", ControlType: "Button" }) ? true : false
    } catch {
    }
    return false
}

IsOutlookRemoveEventPopOutActive() {
    if !(WinActive("ahk_exe OUTLOOK.EXE") || WinActive("ahk_exe olk.exe"))
        return false
    if !RegExMatch(WinGetTitle("A"), "i) - Message \(")
        return false
    try {
        root := OutlookMail_RootElementForHwnd(WinExist("A"))
        if !root
            return false
        return root.FindFirst({ Name: "Remove event", ControlType: "Button" }) ? true : false
    } catch {
    }
    return false
}

OutlookMeeting_ClickMenuItemInActiveWindow(criteriaList) {
    hwnd := WinExist("A")
    root := OutlookMail_RootElementForHwnd(hwnd)
    if !root
        return false
    for criteria in criteriaList {
        try {
            el := root.FindFirst(criteria)
            if el {
                try el.SetFocus()
                Sleep 50
                try el.Click()
                catch {
                    try el.Invoke()
                }
                return true
            }
        } catch {
        }
    }
    return false
}

OutlookMeeting_ClickAccept() {
    return OutlookMeeting_ClickMenuItemInActiveWindow([{ Name: "Accept the meeting", ControlType: "MenuItem" }])
}

OutlookMeeting_ClickFollow() {
    return OutlookMeeting_ClickMenuItemInActiveWindow([{ Name: "Follow;", matchmode: "Substring", ControlType: "MenuItem" }])
}

OutlookMeeting_ClickDecline() {
    return OutlookMeeting_ClickMenuItemInActiveWindow([{ Name: "Decline the meeting", ControlType: "MenuItem" }])
}

; True if user confirms (Yes). Default button is Yes (Enter confirms).
OutlookMeeting_ConfirmDecline() {
    return MsgBox("Decline this meeting invitation?", "Confirm decline", "Icon? YesNo Default1") = "Yes"
}

; Canceled meeting: Button "Remove event" (see outlook-remove-evet.md). No confirmation in script.
OutlookMeeting_ClickRemoveEvent() {
    return OutlookMeeting_ClickMenuItemInActiveWindow([{ Name: "Remove event", ControlType: "Button" }])
}

; Opens "More options" (…) then clicks a MenuItem in the overflow menu (see mark-appointment-request.md).
OutlookMeeting_ClickMoreOptionsThen(menuItemName) {
    try {
        hwnd := WinExist("A")
        root := OutlookMail_RootElementForHwnd(hwnd)
        if !root
            return false
        pane := root.FindFirst({ AutomationId: "Skip to message-region" })
        searchRoot := pane ? pane : root
        moreBtn := ""
        try moreBtn := searchRoot.FindFirst({ AutomationId: "menur7c4", ControlType: "Button" })
        if !moreBtn
            try moreBtn := searchRoot.FindFirst({ Name: "More options", ControlType: "Button", matchmode: "Substring" })
        if !moreBtn
            try moreBtn := root.FindFirst({ AutomationId: "menur7c4", ControlType: "Button" })
        if !moreBtn
            return false
        try moreBtn.SetFocus()
        Sleep 50
        try moreBtn.Click()
        catch {
            try moreBtn.Invoke()
        }
        Sleep 120
        el := ""
        try el := root.FindFirst({ Name: menuItemName, ControlType: "MenuItem" })
        if !el
            try el := UIA.ElementFromHandle(WinExist("A")).FindFirst({ Name: menuItemName, ControlType: "MenuItem" })
        if !el
            try el := root.FindFirst({ Name: menuItemName, matchmode: "Substring", ControlType: "MenuItem" })
        if !el
            return false
        try el.SetFocus()
        Sleep 40
        try el.Click()
        catch {
            try el.Invoke()
        }
        return true
    } catch {
    }
    return false
}

OutlookMail_EnsureNavigationPaneVisible() {
    Outlook_ActivateMainWindow()
    try {
        root := OutlookMail_RootElement()
        if !root
            return false
        ; If the navigation pane exists (or at least the folder tree), we’re good.
        try {
            if root.FindFirst({ Name: "Navigation pane", matchmode: "Substring" })
                return true
            if root.FindFirst({ ControlType: "Tree" })
                return true
        } catch {
        }

        ; Otherwise toggle the nav pane (label may be "Show…" or "Hide…", depending on state).
        navToggleCriteria := [{ Name: "navigation pane", matchmode: "Substring", ControlType: "Button" }, { Name: "Navigation pane",
            matchmode: "Substring", ControlType: "Button" }]
        for x in OutlookMail_CriteriaToggleNavigationPaneRibbon()
            navToggleCriteria.Push(x)
        if OutlookMail_ClickFirst(navToggleCriteria) {
            Sleep 120
            try {
                if root.FindFirst({ ControlType: "Tree" })
                    return true
            } catch {
            }
        }
    } catch {
    }
    return false
}

; Ribbon folder-pane control: Type 50000 (Button), Fluent className prefix fui-Button …
; Name is "Show navigation pane" when the folder pane is collapsed, "Hide navigation pane" when expanded.
OutlookMail_CriteriaShowNavigationPaneRibbon() {
    return [{ Name: "Show navigation pane", Type: 50000, matchmode: "Substring" }, { Name: "Show navigation pane",
        ControlType: "Button", matchmode: "Substring" }, { Name: "Show navigation pane", ClassName: "fui-Button",
            matchmode: "Substring" }
    ]
}

OutlookMail_CriteriaHideNavigationPaneRibbon() {
    return [{ Name: "Hide navigation pane", Type: 50000, matchmode: "Substring" }, { Name: "Hide navigation pane",
        ControlType: "Button", matchmode: "Substring" }, { Name: "Hide navigation pane", ClassName: "fui-Button",
            matchmode: "Substring" }
    ]
}

OutlookMail_CriteriaToggleNavigationPaneRibbon() {
    c := []
    for x in OutlookMail_CriteriaHideNavigationPaneRibbon()
        c.Push(x)
    for x in OutlookMail_CriteriaShowNavigationPaneRibbon()
        c.Push(x)
    return c
}

; New Outlook mail surface is WebView2 (Chromium). Ribbon + lists live under Chrome_RenderWidgetHostHWND1;
; UIA.ElementFromHandle(top-level hwnd) may not include that subtree (see UIA.ElementFromChromium).
OutlookMail_RootElement() {
    Outlook_ActivateMainWindow()
    hwnd := WinExist("A")
    if !hwnd
        return ""
    return OutlookMail_RootElementForHwnd(hwnd)
}

; Chromium root for a specific top-level hwnd — does not activate another window (reading pane vs pop-out invitation).
OutlookMail_RootElementForHwnd(hwnd) {
    if !hwnd
        return ""
    try {
        return UIA.ElementFromChromium("ahk_id " hwnd, 500)
    } catch {
    }
    try {
        return UIA.ElementFromHandle(hwnd)
    } catch {
    }
    return ""
}

OutlookMail_FindFirst(criteriaList) {
    root := OutlookMail_RootElement()
    if !root
        return ""
    for criteria in criteriaList {
        try {
            el := root.FindFirst(criteria)
            if el
                return el
        } catch {
        }
    }
    return ""
}

OutlookMail_ClickFirst(criteriaList) {
    el := OutlookMail_FindFirst(criteriaList)
    if !el
        return false
    try el.SetFocus()
    Sleep 50
    try el.Click()
    catch Error {
        try el.Invoke()
        catch Error {
            return false
        }
    }
    return true
}

; New Outlook: ribbon command buttons (Delete, Move, …) live on the Home tab panel only (outlook.md: TabItem Home, AutomationId "1").
; Call before UIA clicks that target AutomationIds on that panel when View/Help might be active.
OutlookMail_EnsureHomeTab() {
    Outlook_ActivateMainWindow()
    homeTab := OutlookMail_FindFirst([{ Name: "Home", ControlType: "TabItem", AutomationId: "1" }, { AutomationId: "1",
        ControlType: "TabItem" }, { Name: "Home", ControlType: "TabItem" }
    ])
    if !homeTab {
        try {
            Send "!h"
            Sleep 100
        } catch {
        }
        return true
    }
    try {
        if homeTab.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable) && homeTab.SelectionItemPattern.IsSelected {
            Sleep 30
            return true
        }
    } catch {
    }
    try homeTab.SetFocus()
    Sleep 30
    try homeTab.Click()
    catch Error {
        try homeTab.Invoke()
        catch Error {
            try {
                Send "!h"
                Sleep 100
            } catch {
            }
            return true
        }
    }
    Sleep 100
    return true
}

; Left folder list collapsed: ribbon shows "Show navigation pane" (outlook-mail.md: Ribbon … 8,1).
OutlookMail_IsLeftSidePanelHidden() {
    Outlook_ActivateMainWindow()
    return OutlookMail_FindFirst(OutlookMail_CriteriaShowNavigationPaneRibbon()) != ""
}

; Ribbon "high navigation" toggle: show the left folder pane (same control as "Hide navigation pane" when open).
OutlookMail_ClickHighNavigationShowPane() {
    Outlook_ActivateMainWindow()
    if OutlookMail_ClickFirst(OutlookMail_CriteriaShowNavigationPaneRibbon())
        return true
    return OutlookMail_EnsureNavigationPaneVisible()
}

; Nav pane already visible: open Inbox via folder tree only (no ribbon toggle).
OutlookMail_GoToInboxShortcut() {
    Outlook_ActivateMainWindow()
    try {
        root := OutlookMail_RootElement()
        if !root
            return false

        ; Scope search to the Navigation pane subtree to avoid colliding with other “Inbox” elements.
        nav := 0
        try nav := root.FindFirst({ Name: "Navigation pane", matchmode: "Substring" })
        if !nav
            nav := root

        inbox := 0
        ; Prefer the uniquely-named selected variant.
        try inbox := nav.FindFirst({ Name: "Inbox selected", ControlType: "TreeItem" })
        if !inbox
            try inbox := nav.FindFirst({ Name: "Inbox", ControlType: "TreeItem" })
        if !inbox
            try inbox := nav.FindFirst({ Name: "Inbox", matchmode: "Substring", ControlType: "TreeItem" })

        if inbox {
            try inbox.ScrollIntoView()
            try inbox.SetFocus()
            Sleep 40
            try inbox.Click()
            catch {
                try inbox.Invoke()
            }
            return true
        }
    } catch {
    }
    return false
}

OutlookMail_ClickInboxFolder() {
    Outlook_ActivateMainWindow()
    if OutlookMail_IsLeftSidePanelHidden() {
        if !OutlookMail_ClickHighNavigationShowPane()
            return false
        Sleep 120
        return OutlookMail_GoToInboxShortcut()
    }
    return OutlookMail_GoToInboxShortcut()
}

; Ribbon: Hide navigation pane / Show navigation pane (outlook-mail.md: Ribbon … 8,1).
OutlookMail_ToggleHighNavigationPane() {
    Outlook_ActivateMainWindow()
    return OutlookMail_ClickFirst(OutlookMail_CriteriaToggleNavigationPaneRibbon())
}

OutlookCompose_FocusToRecipientsField() {
    Outlook_ActivateMainWindow()
    try {
        root := UIA.ElementFromHandle(WinExist("A"))

        ; #region agent log
        OC_ToLog(msg, data := "{}", hypo := "OC_To") {
            try {
                line := "{"
                    . '"sessionId":"b96502",'
                    . '"runId":"shiftT",'
                    . '"hypothesisId":"' hypo '",'
                    . '"timestamp":' A_TickCount ','
                    . '"location":"Shift keys.ahk:OutlookCompose_FocusToRecipientsField",'
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
            c := WinGetClass("A")
            p := WinGetProcessName("A")
            OC_ToLog("entry", '{"hwnd":' hwnd ',"proc":"' StrReplace(p, '"', '\"') '","class":"' StrReplace(c, '"',
                '\"') '","title":"' StrReplace(SubStr(t, 1, 120), '"', '\"') '"}', "OC_To_A")
        } catch {
        }

        ; Step 1: click the To: row (reactive UI may expand recipients editor)
        okTo := OutlookClickFirst([{ AutomationId: "134", ControlType: "Group" }, { AutomationId: "134" }, { Name: "To:",
            matchmode: "Substring" }])
        try OC_ToLog("after_to_click", '{"ok":' (okTo ? 1 : 0) '}', "OC_To_B")
        catch {
        }
        if !okTo
            return false
        Sleep 80

        ; Step 2: click the recipient entity group (AutomationId looks like REK000070; can change)
        recipGroup := 0
        try recipGroup := root.FindFirst({ AutomationId: "REK", matchmode: "Substring", ControlType: "Group" })
        if !recipGroup
            try recipGroup := root.FindFirst({ AutomationId: "REK", matchmode: "Substring" })
        if !recipGroup {
            ; Broad fallback: find any group that looks like a recipient entity (class contains _EType_RECIPIENT_ENTITY)
            try recipGroup := root.FindFirst({ ClassName: "_EType_RECIPIENT_ENTITY", matchmode: "Substring",
                ControlType: "Group" })
        }
        if recipGroup {
            n := "", aid := "", cn := "", ct := "", off := "", en := ""
            try n := recipGroup.Name
            try aid := recipGroup.AutomationId
            try cn := recipGroup.ClassName
            try ct := recipGroup.ControlType
            try off := recipGroup.IsOffscreen
            try en := recipGroup.IsEnabled
            try OC_ToLog("recip_group_found", '{"name":"' StrReplace(SubStr(n, 1, 80), '"', '\"') '","automationId":"' StrReplace(
                SubStr(aid, 1, 80), '"', '\"') '","className":"' StrReplace(SubStr(cn, 1, 80), '"', '\"') '","controlType":"' StrReplace(
                    ct, '"', '\"') '","isOffscreen":' (off ? 1 : 0) ',"isEnabled":' (en ? 1 : 0) '}', "OC_To_C")
            catch {
            }
        } else {
            try OC_ToLog("recip_group_not_found", "{}", "OC_To_C")
            catch {
            }
        }
        if recipGroup {
            try recipGroup.ScrollIntoView()
            try recipGroup.SetFocus()
            Sleep 30
            try recipGroup.Click()
            catch {
                try recipGroup.Invoke()
            }
            Sleep 60

            ; Step 3: focus the editable field inside the recipient group (if exposed)
            edit := 0
            try edit := recipGroup.FindFirst({ ControlType: "Edit" })
            if !edit
                try edit := recipGroup.FindFirst({ Type: 50004 })
            if edit {
                try OC_ToLog("inner_edit_found", "{}", "OC_To_D")
                catch {
                }
                try edit.SetFocus()
                Sleep 20
                try edit.Click()
                try {
                    fe := UIA.GetFocusedElement()
                    fn := "", fa := "", ft := ""
                    try fn := fe.Name
                    try fa := fe.AutomationId
                    try ft := fe.Type
                    OC_ToLog("focused_after_edit", '{"name":"' StrReplace(SubStr(fn, 1, 80), '"', '\"') '","automationId":"' StrReplace(
                        SubStr(fa, 1, 80), '"', '\"') '","type":' (ft = "" ? -1 : ft) '}', "OC_To_E")
                } catch {
                }
                return true
            }
            ; Fallback: click the hover target wrapper (often the direct text/caret host)
            wrap := 0
            try wrap := recipGroup.FindFirst({ ClassName: "lpcWrapper", matchmode: "Substring" })
            if wrap {
                try OC_ToLog("wrapper_found", "{}", "OC_To_D")
                catch {
                }
                try wrap.SetFocus()
                Sleep 20
                try wrap.Click()
                try {
                    fe := UIA.GetFocusedElement()
                    fn := "", fa := "", ft := ""
                    try fn := fe.Name
                    try fa := fe.AutomationId
                    try ft := fe.Type
                    OC_ToLog("focused_after_wrapper", '{"name":"' StrReplace(SubStr(fn, 1, 80), '"', '\"') '","automationId":"' StrReplace(
                        SubStr(fa, 1, 80), '"', '\"') '","type":' (ft = "" ? -1 : ft) '}', "OC_To_E")
                } catch {
                }
                return true
            }
            try OC_ToLog("no_inner_target", "{}", "OC_To_D")
            catch {
            }
            return true
        }
    } catch {
    }
    return false
}

IsNewOutlookActive() {
    if !(WinActive("ahk_exe OUTLOOK.EXE") || WinActive("ahk_exe olk.exe"))
        return false
    cls := ""
    title := ""
    exe := ""
    try cls := WinGetClass("A")
    try title := WinGetTitle("A")
    try exe := WinGetProcessName("A")
    return InStr(cls, "Outlook Host")
    || InStr(title, " - Outlook")
    || RegExMatch(title, "i)^(New event|Reminders?)")
}

OutlookFindFirst(criteriaList) {
    try {
        root := UIA.ElementFromHandle(WinExist("A"))
        for criteria in criteriaList {
            el := root.FindFirst(criteria)
            if el
                return el
        }
    } catch Error {
    }
    return ""
}

OutlookFocusFirst(criteriaList) {
    el := OutlookFindFirst(criteriaList)
    if !el
        return false
    try el.SetFocus()
    return true
}

OutlookClickFirst(criteriaList) {
    el := OutlookFindFirst(criteriaList)
    if !el
        return false
    try el.SetFocus()
    Sleep 50
    try el.Click()
    catch Error {
        try el.Invoke()
        catch Error {
            return false
        }
    }
    return true
}
