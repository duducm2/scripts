; =============================================================================
; Shift keys module: outlook_appointment_hotif_02.ahk
; Outlook appointment date/time helpers and hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

; ----- Outlook Appointment: Date/Time helpers -----
Outlook_ClickStartDate() {
    ClickOutlookByIdThenNameClass("4098", "Start date, combo", "RichEdit20WPT", "Edit")
}

Outlook_ClickStartDatePicker() {
    ; Robust open: focus the Date Picker and press Enter
    if FocusOutlookField({ AutomationId: "4352" }) {
        Sleep 80
        Send "{Enter}"
        return
    }
    if FocusOutlookField({ Name: "Date Picker", ControlType: "Button" }) {
        Sleep 80
        Send "{Enter}"
        return
    }
}

Outlook_ClickStartTime() {
    ClickOutlookByIdThenNameClass("4096", "Start time, combo", "RichEdit20WPT", "Edit")
}

Outlook_ClickStartTime_1100AM() {
    ; Clicks the button showing 11:00 AM (start)
    ClickOutlookByIdThenNameClass("4354", "11:00 AM", "AfxWndW", "Button")
}

Outlook_ClickEndDate() {
    ClickOutlookByIdThenNameClass("4099", "End date, combo", "RichEdit20WPT", "Edit")
}

Outlook_ClickEndDatePicker() {
    ; Date Picker next to End date
    ClickOutlookByIdThenNameClass("4353", "Date Picker", "AfxWndW", "Button")
}

Outlook_ClickEndTime() {
    ClickOutlookByIdThenNameClass("4097", "End time, combo", "RichEdit20WPT", "Edit")
}

Outlook_ClickEndTime_1200PM() {
    ; Clicks the button showing 12:00 PM (end)
    ClickOutlookByIdThenNameClass("4355", "12:00 PM", "AfxWndW", "Button")
}

; Shift + S : prefer Send (meeting), else Save.
; A bare Click() is not used: on a split button it expands the chevron menu and
; the old shortcut treated that as success. Each candidate must match a current
; Send/Save label, and the click counts only when the form leaves or Outlook
; opens a dialog. Save is skipped once a Send control was found.
+S:: {
    result := ""
    try result := Appt_RunWithLoading("Send/Save", Appt_ClickSendOrSave)
    if (result = true)
        return
    if (result = "send-miss")
        ShowCenteredOverlay_Utils("❌ Appointment: Send did not complete", 1600, BANNER_ACCENT_ERROR)
    else if (result = "save-miss")
        ShowCenteredOverlay_Utils("❌ Appointment: Save did not complete", 1600, BANNER_ACCENT_ERROR)
    else
        ShowCenteredOverlay_Utils("❌ Appointment: Send/Save not found", 1400, BANNER_ACCENT_ERROR)
}

Appt_ClickSendOrSave() {
    root := Appt_GetRootActive()
    if !root
        return "missing"
    hwnd := WinExist("A")
    if !hwnd
        return "missing"
    candidates := Appt_SendSave_Collect(root)
    if (candidates.Length = 0)
        return "missing"
    sawSend := false
    for c in candidates {
        if (c.action != "send")
            continue
        sawSend := true
        if Appt_SendSave_Try(hwnd, c.el)
            return true
    }
    if sawSend
        return "send-miss"
    for c in candidates {
        if (c.action != "save")
            continue
        if Appt_SendSave_Try(hwnd, c.el)
            return true
    }
    return "save-miss"
}

Appt_SendSave_Try(hwnd, el) {
    return Appt_SendSave_Commit(hwnd, el, Appt_SendSave_IsEnabled(el), !Appt_SendSave_IsOffscreen(el))
}

Appt_SendSave_Collect(root) {
    list := []
    seen := Map()
    for elementId in ["Send", "Save"] {
        el := 0
        try el := UiElements_TrySaved(root, "OutlookAppointment", elementId)
        if IsObject(el)
            Appt_SendSave_Consider(list, seen, el, "saved", 0, 0)
    }
    scopes := []
    try {
        for tb in root.FindAll({ Type: UIA.Type.ToolBar })
            scopes.Push(tb)
    } catch {
    }
    for barName in ["Event form commands", "Meeting form commands", "Appointment commands", "Command bar"] {
        for typeName in ["ToolBar", "Group", "Pane"] {
            bar := 0
            try bar := root.FindFirst({ Name: barName, ControlType: typeName })
            if IsObject(bar)
                scopes.Push(bar)
        }
    }
    for scope in scopes
        Appt_SendSave_Scan(list, seen, scope, "toolbar", 0, 0)
    top := 0
    limit := 0
    try {
        loc := root.Location
        top := loc.y
        band := loc.h * 0.28
        if (band < 140)
            band := 140
        if (band > 260)
            band := 260
        limit := loc.y + band
    } catch {
        top := 0
        limit := 0
    }
    Appt_SendSave_Scan(list, seen, root, "window", top, limit)
    return Appt_SendSave_Sort(list)
}

Appt_SendSave_Scan(list, seen, scope, band, top, limit) {
    if !IsObject(scope)
        return
    for typeId in [UIA.Type.Button, UIA.Type.SplitButton] {
        try {
            for el in scope.FindAll({ Type: typeId })
                Appt_SendSave_Consider(list, seen, el, band, top, limit)
        } catch {
        }
    }
    if (band != "toolbar")
        return
    try {
        for el in scope.FindAll({ Type: UIA.Type.Text })
            Appt_SendSave_Consider(list, seen, el, band, top, limit)
    } catch {
    }
}

Appt_SendSave_Consider(list, seen, el, band, top, limit) {
    if !IsObject(el)
        return
    action := Appt_SendSave_Classify(Appt_SendSave_Label(el))
    if (action = "") {
        id := ""
        try id := el.AutomationId
        action := Appt_SendSave_ClassifyId(id)
    }
    if (action = "")
        return
    btn := Appt_SendSave_AsButton(el)
    if !IsObject(btn)
        return
    if (band = "window" && !Appt_SendSave_InTopBand(btn, top, limit))
        return
    key := Appt_SendSave_Key(btn)
    if (key = "" || seen.Has(key))
        return
    seen[key] := true
    list.Push({
        el: btn,
        action: action,
        score: Appt_SendSave_Score(btn, action, band)
    })
}

Appt_SendSave_Commit(hwnd, el, wasEnabled, wasOnScreen) {
    hadMenu := Appt_SendSave_HasPopupMenu(hwnd)
    wasSplitOpen := Appt_SendSave_SplitOpened(el)
    if Appt_SendSave_Fire(el, "invoke") && Appt_SendSave_Wait(hwnd, el, wasEnabled, wasOnScreen, hadMenu, wasSplitOpen)
        return true
    if Appt_SendSave_Fire(el, "mouse") && Appt_SendSave_Wait(hwnd, el, wasEnabled, wasOnScreen, hadMenu, wasSplitOpen)
        return true
    try el.SetFocus()
    catch {
    }
    Sleep 40
    if !Appt_SendSave_FocusIsOn(el) || Appt_SendSave_SplitOpened(el)
        return false
    try Send "{Enter}"
    return Appt_SendSave_Wait(hwnd, el, wasEnabled, wasOnScreen, hadMenu, wasSplitOpen)
}

Appt_SendSave_Wait(hwnd, el, wasEnabled, wasOnScreen, hadMenu, wasSplitOpen) {
    Sleep 120
    deadline := A_TickCount + 900
    checkedMenu := false
    loop {
        if !WinExist("ahk_id " hwnd)
            return true
        menuOpened := !wasSplitOpen && Appt_SendSave_SplitOpened(el)
        if (!menuOpened && !hadMenu && !checkedMenu) {
            checkedMenu := true
            menuOpened := Appt_SendSave_HasPopupMenu(hwnd)
        }
        if menuOpened {
            try Send "{Esc}"
            Sleep 60
            return false
        }
        fg := 0
        try fg := WinExist("A")
        if (fg && fg != hwnd) {
            if Appt_SendSave_ForegroundIsMenu(fg) {
                try Send "{Esc}"
                Sleep 60
                return false
            }
            exe := ""
            try exe := WinGetProcessName("ahk_id " fg)
            if (exe = "OUTLOOK.EXE" || exe = "olk.exe")
                return true
        }
        if Appt_SendSave_Retired(el, wasEnabled, wasOnScreen)
            return true
        if (A_TickCount >= deadline)
            break
        Sleep 80
    }
    return false
}

Appt_SendSave_Fire(el, how) {
    if (how = "invoke") {
        try {
            if el.GetPropertyValue(UIA.Property.IsInvokePatternAvailable) {
                el.InvokePattern.Invoke()
                return true
            }
        } catch {
        }
        return false
    }
    loc := ""
    try loc := el.Location
    if !IsObject(loc) || loc.w < 4 || loc.h < 4
        return false
    span := loc.w
    if (loc.w >= 56)
        span := loc.w - 28
    x := Round(loc.x + span / 2)
    y := Round(loc.y + loc.h / 2)
    prevMode := A_CoordModeMouse
    CoordMode("Mouse", "Screen")
    ok := false
    try {
        Click(x " " y)
        ok := true
    } catch {
    }
    CoordMode("Mouse", prevMode)
    return ok
}

Appt_SendSave_Classify(raw) {
    n := Appt_SendSave_NormalizeLabel(raw)
    if (n = "")
        return ""
    if RegExMatch(n,
        "i)cancel|discard|don.?t|do not|n[aã]o enviar|response|availability|reminder|delete|remove|follow|tentative|decline|\\baccept\\b|forward|reply|\\boptions?\\b|chevron|dropdown"
    )
        return ""
    if (n = "send" || n = "enviar" || InStr(n, "send update") = 1 || InStr(n, "send invitation") = 1
    || InStr(n, "send invite") = 1 || InStr(n, "enviar atualiza") = 1 || InStr(n, "enviar convite") = 1)
        return "send"
    if (n = "save" || n = "salvar" || InStr(n, "save and close") = 1 || InStr(n, "save & close") = 1
    || InStr(n, "save event") = 1 || InStr(n, "save meeting") = 1 || InStr(n, "salvar e fechar") = 1
    || InStr(n, "salvar evento") = 1 || InStr(n, "salvar reuni") = 1)
        return "save"
    return ""
}

Appt_SendSave_ClassifyId(raw) {
    id := StrLower(Trim(raw))
    if (id = "")
        return ""
    if RegExMatch(id, "i)cancel|discard|delete|remove|option|chevron|expand")
        return ""
    if InStr(id, "send") || InStr(id, "enviar")
        return "send"
    if InStr(id, "save") || InStr(id, "salvar")
        return "save"
    return ""
}

Appt_SendSave_NormalizeLabel(raw) {
    s := Trim(raw)
    s := RegExReplace(s, "i)\s*\([^)]*\)\s*$", "")
    s := RegExReplace(s, "\s+", " ")
    return StrLower(Trim(s))
}

Appt_SendSave_Label(el) {
    name := ""
    try name := Trim(el.Name)
    if (name != "")
        return name
    try {
        for t in el.FindAll({ Type: UIA.Type.Text }) {
            tn := ""
            try tn := Trim(t.Name)
            if (tn != "")
                return tn
        }
    } catch {
    }
    help := ""
    try help := Trim(el.HelpText)
    return help
}

Appt_SendSave_AsButton(el) {
    cur := el
    loop 8 {
        if !IsObject(cur)
            return 0
        t := 0
        try t := cur.Type
        if (t = UIA.Type.Button || t = UIA.Type.SplitButton)
            return cur
        parent := 0
        try parent := cur.Parent
        catch
            return 0
        if !IsObject(parent)
            return 0
        cur := parent
    }
    return 0
}

Appt_SendSave_Score(el, action, band) {
    score := (action = "send") ? 0 : 100
    n := Appt_SendSave_NormalizeLabel(Appt_SendSave_Label(el))
    if (n != "send" && n != "enviar" && n != "save" && n != "salvar")
        score += 10
    try {
        if (el.Type = UIA.Type.SplitButton)
            score += 5
    } catch {
    }
    if (band != "toolbar" && band != "saved")
        score += 20
    if !Appt_SendSave_IsEnabled(el)
        score += 40
    if Appt_SendSave_IsOffscreen(el)
        score += 30
    try score += el.Location.x / 10000
    return score
}

Appt_SendSave_Sort(list) {
    n := list.Length
    loop n {
        i := A_Index
        loop (n - i) {
            j := A_Index
            if (list[j].score > list[j + 1].score) {
                tmp := list[j]
                list[j] := list[j + 1]
                list[j + 1] := tmp
            }
        }
    }
    return list
}

Appt_SendSave_Key(el) {
    try {
        loc := el.Location
        name := ""
        try name := el.Name
        return el.Type "|" Round(loc.x) "|" Round(loc.y) "|" name
    } catch {
    }
    return ""
}

Appt_SendSave_InTopBand(el, top, limit) {
    if (limit <= top)
        return false
    try {
        loc := el.Location
        if (loc.w < 8 || loc.h < 8)
            return false
        mid := loc.y + loc.h / 2
        return mid >= top - 4 && mid <= limit
    } catch {
    }
    return false
}

Appt_SendSave_IsEnabled(el) {
    try return !!el.GetPropertyValue(UIA.Property.IsEnabled)
    return true
}

Appt_SendSave_IsOffscreen(el) {
    try return !!el.GetPropertyValue(UIA.Property.IsOffscreen)
    return false
}

Appt_SendSave_Retired(el, wasEnabled, wasOnScreen) {
    try {
        if (wasOnScreen && el.GetPropertyValue(UIA.Property.IsOffscreen))
            return true
        if (wasEnabled && !el.GetPropertyValue(UIA.Property.IsEnabled))
            return true
    } catch {
    }
    return false
}

Appt_SendSave_SplitOpened(el) {
    cur := el
    loop 4 {
        if !IsObject(cur)
            return false
        opened := false
        try {
            if cur.GetPropertyValue(UIA.Property.IsExpandCollapsePatternAvailable)
                opened := cur.ExpandCollapsePattern.ExpandCollapseState = UIA.ExpandCollapseState.Expanded
        } catch {
        }
        if opened
            return true
        parent := 0
        try parent := cur.Parent
        catch
            return false
        if !IsObject(parent) || Appt_SendSave_Same(parent, cur)
            return false
        cur := parent
    }
    return false
}

Appt_SendSave_HasPopupMenu(hwnd) {
    root := 0
    try root := UIA.ElementFromHandle(hwnd)
    if !IsObject(root)
        return false
    el := 0
    try el := root.FindFirst({ ControlType: "Menu" })
    return IsObject(el)
}

Appt_SendSave_ForegroundIsMenu(hwnd) {
    root := 0
    try root := UIA.ElementFromHandle(hwnd)
    if !IsObject(root)
        return false
    t := 0
    try t := root.Type
    if (t = UIA.Type.Menu)
        return true
    el := 0
    try el := root.FindFirst({ ControlType: "Menu" })
    if IsObject(el)
        return true
    try el := root.FindFirst({ ControlType: "MenuItem" })
    return IsObject(el)
}

Appt_SendSave_FocusIsOn(el) {
    fe := 0
    try fe := UIA.GetFocusedElement()
    cur := fe
    loop 8 {
        if !IsObject(cur)
            return false
        if Appt_SendSave_Same(cur, el)
            return true
        parent := 0
        try parent := cur.Parent
        catch
            return false
        if !IsObject(parent) || Appt_SendSave_Same(parent, cur)
            return false
        cur := parent
    }
    return false
}

Appt_SendSave_Same(a, b) {
    if !IsObject(a) || !IsObject(b)
        return false
    try {
        idA := a.RuntimeId
        idB := b.RuntimeId
        if (idA != "" && idA = idB)
            return true
    } catch {
    }
    try {
        la := a.Location
        lb := b.Location
        return Round(la.x) = Round(lb.x) && Round(la.y) = Round(lb.y) && Round(la.w) = Round(lb.w)
    } catch {
    }
    return false
}

; Shift + D : Start date (combo) - Start Date
+D:: {
    isNew := false
    try isNew := IsNewOutlookActive()

    Appt_RunWithLoading("Start date", (*) => (
        isNew
            ? (Appt_ClickSaved("StartDate") || Appt_PopoverInvokeFirst([{ Name: "Start date", ControlType: "ComboBox" }, { Name: "Start date",
                Type: 50003 }, { Name: "Start date",
                    ControlType: "Button" }, { Name: "Start date", Type: 50000 }, { AutomationId: "DatePicker",
                        matchmode: "Substring" }
            ]) || (ShowCenteredOverlay_Utils("❌ Appointment: Start date not found", 1400, BANNER_ACCENT_ERROR), false))
            : (Appt_ClickSaved("StartDate") || (Outlook_ClickStartDate(), true))
    ))
}

; Shift + P : Start date picker - Picker
+P:: {
    ; New Outlook: repurpose Shift+P to Private toggle modal (date picker concept removed).
    if IsNewOutlookActive() {
        Appt_RunWithLoading("Private", Appt_RunPrivate)
        return
    }
    Outlook_ClickStartDatePicker()
}

Appt_RunPrivate() {
    choice := Appt_SelectFromModal("Appointment privacy", [{ k: "1", label: "Private" }, { k: "2", label: "Not private" }],
    "[1-2] Select  [Esc] Cancel")
    if !choice
        return false
    if (choice = "1")
        return Appt_ClickSaved("Private") || Appt_OpenMenuAndPick([{ Name: "Private", ControlType: "Button" }, { Name: "Not private",
            ControlType: "Button" }, { Name: "Private", matchmode: "Substring", ControlType: "Button" }, { Name: "Not private",
                matchmode: "Substring", ControlType: "Button" }
        ], "Private")
    return Appt_OpenMenuAndPick([{ Name: "Private", ControlType: "Button" }, { Name: "Not private",
        ControlType: "Button" }, { Name: "Private", matchmode: "Substring", ControlType: "Button" }, { Name: "Not private",
            matchmode: "Substring", ControlType: "Button" }
    ], "Not private")
}

; Shift + T : Start time (combo) - Time
+T:: {
    Appt_RunWithLoading("Start time", (*) => (
        IsNewOutlookActive()
            ? (Appt_ClickSaved("StartTime") || Appt_PopoverFocusFirst([{ Name: "Start time", ControlType: "ComboBox" }, { Name: "Start time",
                Type: 50003 }, { AutomationId: "ComboBox",
                    matchmode: "Substring" }]) || (ShowCenteredOverlay_Utils(
                        "❌ Appointment: Start time not found", 1400,
                        BANNER_ACCENT_ERROR), false))
            : (Appt_ClickSaved("StartTime") || (Outlook_ClickStartTime(), true))
    ))
}

; Shift + E : End date (combo) - End Date
+E:: {
    Appt_RunWithLoading("End time", (*) => (
        IsNewOutlookActive()
            ? (Appt_ClickSaved("EndTime") || Appt_PopoverFocusFirst([{ Name: "End time", ControlType: "ComboBox" }, { Name: "End time",
                Type: 50003 }]) ||
            (ShowCenteredOverlay_Utils("❌ Appointment: End time not found", 1400, BANNER_ACCENT_ERROR), false))
            : (Appt_ClickSaved("EndTime") || (Outlook_ClickEndDate(), true))
    ))
}

; Shift + H : Scheduler / Scheduling assistant (New Outlook) or End time (classic)
+H:: {
    if IsNewOutlookActive() {
        Appt_RunWithLoading("Scheduler", (*) => (
            Appt_ClickSaved("Scheduler")
            || Appt_ClickAny([{ Name: "Scheduler", ControlType: "Button" }, { Name: "Scheduling assistant",
                matchmode: "Substring",
                ControlType: "Button" }, { Name: "Scheduling", matchmode: "Substring", ControlType: "Button" }
            ]) || (ShowCenteredOverlay_Utils("❌ Appointment: Scheduler not found", 1400, BANNER_ACCENT_ERROR),
            false)
        ))
        return
    }
    Outlook_ClickEndTime()
}

; Shift + A : All day checkbox - All Day
+A:: {
    if IsNewOutlookActive() {
        Appt_RunWithLoading("All day", (*) => (
            Appt_ClickSaved("AllDay") || Appt_PopoverToggleFirst([{ Name: "All day", ControlType: "CheckBox" }, { Name: "All day",
                Type: 50002 },
            ; New Outlook exposes this as a switch (button) with a stable AutomationId (e.g. Toggle9777).
            { AutomationId: "Toggle", matchmode: "Substring", ControlType: "Button" }, { AutomationId: "Toggle",
                matchmode: "Substring", Type: 50000 }, { Name: "All day", ControlType: "Button" }, { Name: "All day",
                    Type: 50000 }
            ]) || (ShowCenteredOverlay_Utils("❌ Appointment: All day not found", 1400, BANNER_ACCENT_ERROR),
            false)
        ))
        return
    }
    ; Classic fallback (existing behavior)
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)
        if Appt_ClickSaved("AllDay")
            return
        checkbox := root.FindFirst({ AutomationId: "4226", ControlType: "CheckBox" })
        if !checkbox
            checkbox := root.FindFirst({ Name: "All day", ControlType: "CheckBox" })
        if checkbox
            checkbox.Invoke()
    } catch {
    }
}

; Shift + I : Title field - Title
+I:: {
    if Appt_ClickSaved("Title")
        return
    if FocusOutlookField({ Name: "Add title", ControlType: "Edit" })
        return
    if FocusOutlookField({ AutomationId: "4100" }) ; Title
        return
    if FocusOutlookField({ Name: "Title", ControlType: "Edit" })
        return
}

; Shift + R : Required / To field - Required
+R:: {
    if Appt_ClickSaved("Required")
        return
    if FocusOutlookField({ Name: "Invite required attendees", ControlType: "Group" })
        return
    if FocusOutlookField({ Name: "Invite required attendees", ControlType: "Text" })
        return
    if FocusOutlookField({ AutomationId: "4109" }) ; Required
        return
    if FocusOutlookField({ Name: "Required", ControlType: "Edit" })
        return
}

; Shift + O : Location -> Body - lOcation (moved off Shift+L)
+O:: {
    if Appt_ClickSaved("Location") {
        Sleep 100
        Send "{Tab}"
        return
    }
    if FocusOutlookField({ AutomationId: "location-suggestions-picker-input", ControlType: "ComboBox" }) {
        Sleep 100
        Send "{Tab}"
        return
    }
    if FocusOutlookField({ Name: "Add a room or location", ControlType: "ComboBox" }) {
        Sleep 100
        Send "{Tab}"
        return
    }
    if FocusOutlookField({ AutomationId: "4111" }) { ; Location
        Sleep 100
        Send "{Tab}"
        return
    }
    if FocusOutlookField({ Name: "Location", ControlType: "Edit" }) {
        Sleep 100
        Send "{Tab}"
        return
    }
}

Appt_IsSchedulerView() {
    root := Appt_GetRootActive()
    if !root
        return false
    try {
        if root.FindFirst({ Name: "Scheduling grid", matchmode: "Substring" })
            return true
        if root.FindFirst({ AutomationId: "Jump to Scheduling grid-region", matchmode: "Substring" })
            return true
    } catch {
    }
    return false
}

Appt_ClickSchedulerSuggestionNav(isNext) {
    if Appt_ClickSaved(isNext ? "NextSuggestion" : "PrevSuggestion")
        return true
    root := Appt_GetRootActive()
    if !root
        return false
    needle := isNext ? "Selects the next time suggestion" : "Selects the previous time suggestion"
    try {
        btn := root.FindFirst({ Name: needle, matchmode: "Substring", ControlType: "Button" })
        if btn {
            try btn.Click()
            catch {
                try btn.Invoke()
            }
            return true
        }
    } catch {
    }
    return false
}

Appt_ClickDayNav(isNext) {
    ; Best-effort: look for previous/next day arrow buttons in the schedule header.
    if Appt_ClickSaved(isNext ? "NextDay" : "PrevDay")
        return true
    root := Appt_GetRootActive()
    if !root
        return false
    ; Primary targeting (New Outlook): "Go to previous day <date>" / "Go to next day <date>"
    candidates := isNext
        ? [{ Name: "Go to next", matchmode: "Substring", ControlType: "Button" }, { Name: "go to next",
            matchmode: "Substring",
            ControlType: "Button" }, { Name: "Next", matchmode: "Substring", ControlType: "Button" }, { Name: "Forward",
                matchmode: "Substring", ControlType: "Button" }, { Name: "Next day", matchmode: "Substring",
                    ControlType: "Button" }
        ]
            : [{ Name: "Go to previous", matchmode: "Substring", ControlType: "Button" }, { Name: "go to previous",
                matchmode: "Substring", ControlType: "Button" }, { Name: "Previous", matchmode: "Substring",
                    ControlType: "Button" }, { Name: "Back", matchmode: "Substring", ControlType: "Button" }, { Name: "Previous day",
                        matchmode: "Substring", ControlType: "Button" }
            ]
    for crit in candidates {
        try {
            btn := root.FindFirst(crit)
            if btn {
                try btn.Click()
                catch {
                    try btn.Invoke()
                }
                return true
            }
        } catch {
        }
    }
    return false
}

Appt_SchedulerClickBack() {
    return Appt_ClickSaved("Back")
    || Appt_ClickInCommandBar([{ Name: "Back", ControlType: "Button" }])
    || Appt_ClickAny([{ Name: "Back", ControlType: "Button" }])
}

Appt_SchedulerClickOptions() {
    return Appt_ClickSaved("Options")
    || Appt_ClickInCommandBar([{ Name: "Options", ControlType: "Button" }, { Name: "Options", matchmode: "Substring",
        ControlType: "Button" }])
    || Appt_ClickAny([{ Name: "Options", ControlType: "Button" }, { Name: "Options", matchmode: "Substring",
        ControlType: "Button" }])
}

Appt_SchedulerClickAddAttendee(isOptional) {
    elementId := isOptional ? "AddOptional" : "AddRequired"
    name := isOptional ? "Add optional attendee" : "Add required attendee"
    return Appt_ClickSaved(elementId) || Appt_ClickAny([{ Name: name, ControlType: "Button" }, { Name: name, matchmode: "Substring",
        ControlType: "Button" }])
}

Appt_SchedulerFocusDateTimeControl(kind) {
    ; Focus core controls in scheduler/editor view by visible names.
    if (kind = "start_date")
        return Appt_ClickSaved("StartDate") || Appt_PopoverFocusFirst([{ Name: "Start date", ControlType: "ComboBox" }, { Name: "Start date",
            Type: 50003 }])
    if (kind = "start_time")
        return Appt_ClickSaved("StartTime") || Appt_PopoverFocusFirst([{ Name: "Start time", ControlType: "ComboBox" }, { Name: "Start time",
            Type: 50003 }])
    if (kind = "end_time")
        return Appt_ClickSaved("EndTime") || Appt_PopoverFocusFirst([{ Name: "End time", ControlType: "ComboBox" }, { Name: "End time",
            Type: 50003 }])
    if (kind = "all_day")
        return Appt_ClickSaved("AllDay") || Appt_PopoverInvokeFirst([{ Name: "All day", ControlType: "CheckBox" }, { Name: "All day",
            Type: 50002 }, { Name: "All day",
                ControlType: "Button" }])
    if (kind = "time_zone")
        return Appt_ClickSaved("TimeZone") || Appt_ToggleOrClickAny([{ Name: "Show event time zones", matchmode: "Substring",
            ControlType: "Button" }, { Name: "Time zone",
                matchmode: "Substring", ControlType: "Button" }])
    return false
}

Appt_ToggleOrClickAny(criteriaList) {
    root := Appt_GetRootActive()
    if !root
        return false
    for crit in criteriaList {
        try {
            el := root.FindFirst(crit)
            if !el
                continue
            try {
                if el.IsTogglePatternAvailable {
                    el.TogglePattern.Toggle()
                    return true
                }
            } catch {
            }
            try el.Click()
            catch {
                try el.Invoke()
            }
            return true
        } catch {
        }
    }
    return false
}

; Shift + L / Shift + K : Previous/Next navigation (day in editor, suggestions in scheduler view)
+K:: {
    if !IsNewOutlookActive()
        return
    Appt_RunWithLoading("Prev", (*) => (
        (Appt_IsSchedulerView() ? (Appt_ClickSchedulerSuggestionNav(false) || Appt_ClickDayNav(false)) :
            Appt_ClickDayNav(false))
        || (ShowCenteredOverlay_Utils("❌ Appointment: Previous not found", 1400, BANNER_ACCENT_ERROR), false)
    ))
}

+L:: {
    if !IsNewOutlookActive()
        return
    Appt_RunWithLoading("Next", (*) => (
        (Appt_IsSchedulerView() ? (Appt_ClickSchedulerSuggestionNav(true) || Appt_ClickDayNav(true)) :
            Appt_ClickDayNav(
                true))
        || (ShowCenteredOverlay_Utils("❌ Appointment: Next not found", 1400, BANNER_ACCENT_ERROR), false)
    ))
}

; Shift + Y : Today
+Y:: {
    if !IsNewOutlookActive()
        return
    Appt_RunWithLoading("Today", (*) => (
        Appt_ClickSaved("Today")
        || Appt_ClickAny([{ Name: "Today", ControlType: "Button" }, { Name: "Today", matchmode: "Substring",
            ControlType: "Button" }])
        || (ShowCenteredOverlay_Utils("❌ Appointment: Today not found", 1400, BANNER_ACCENT_ERROR), false)
    ))
}

; Shift + F : Current date header button
+F:: {
    if !IsNewOutlookActive()
        return
    Appt_RunWithLoading("Date", (*) => (
        Appt_ClickSaved("DateHeader")
        || Appt_ClickAny([{ Name: "Thu", matchmode: "Substring", ControlType: "Button" }, { Name: "Apr", matchmode: "Substring",
            ControlType: "Button" }, { Name: "Week", matchmode: "Substring", ControlType: "Button" }
        ]) || (ShowCenteredOverlay_Utils("❌ Appointment: Date header not found", 1400, BANNER_ACCENT_ERROR),
        false)
    ))
}

; Scheduling Assistant view controls (work in scheduler view; safe no-ops otherwise)
; Shift + Backspace : Back (avoid collision with Shift+B = Body)
+Backspace:: {
    if !IsNewOutlookActive()
        return
    Appt_RunWithLoading("Back", (*) => (
        Appt_SchedulerClickBack() || (ShowCenteredOverlay_Utils("❌ Appointment: Back not found", 1400,
            BANNER_ACCENT_ERROR), false)
    ))
}

; Shift + N : Options
+N:: {
    if !IsNewOutlookActive()
        return
    Appt_RunWithLoading("Options", (*) => (
        Appt_SchedulerClickOptions() || (ShowCenteredOverlay_Utils("❌ Appointment: Options not found", 1400,
            BANNER_ACCENT_ERROR), false)
    ))
}

; Shift + T : Start date (scheduler quick focus) (doesn't override existing popover binding, since it's same key)
; (No extra binding needed.)

; Shift + Z : Time zone
+Z:: {
    if !IsNewOutlookActive()
        return
    Appt_RunWithLoading("Time zone", (*) => (
        Appt_SchedulerFocusDateTimeControl("time_zone") || (ShowCenteredOverlay_Utils(
            "❌ Appointment: Time zone not found", 1400, BANNER_ACCENT_ERROR), false)
    ))
}

; Shift + J : Add required attendee
+J:: {
    if !IsNewOutlookActive()
        return
    Appt_RunWithLoading("Add required", (*) => (
        Appt_SchedulerClickAddAttendee(false) || (ShowCenteredOverlay_Utils(
            "❌ Appointment: Add required not found",
            1400, BANNER_ACCENT_ERROR), false)
    ))
}

; Alt + O : Add optional attendee
!o:: {
    if !IsNewOutlookActive()
        return
    Appt_RunWithLoading("Add optional", (*) => (
        Appt_SchedulerClickAddAttendee(true) || (ShowCenteredOverlay_Utils(
            "❌ Appointment: Add optional not found",
            1400, BANNER_ACCENT_ERROR), false)
    ))
}

; Shift + B : Body (from Location) - Body
+B:: {
    Appt_RunWithLoading("Body", (*) => (
        IsNewOutlookActive()
            ? (Appt_FocusBodyField_NewOutlook() || (ShowCenteredOverlay_Utils("❌ Appointment: Body not found",
                1400,
                BANNER_ACCENT_ERROR), false))
            : (true)
    ))
    if IsNewOutlookActive()
        return
    ; Classic fallback: tab from Location into body
    if FocusOutlookField({ AutomationId: "location-suggestions-picker-input", ControlType: "ComboBox" }) {
        Sleep 100
        Send "{Tab}"
        return
    }
    if FocusOutlookField({ Name: "Add a room or location", ControlType: "ComboBox" }) {
        Sleep 100
        Send "{Tab}"
        return
    }
    if FocusOutlookField({ AutomationId: "4111" }) { ; Location
        Sleep 100
        Send "{Tab}"
        return
    }
    if FocusOutlookField({ Name: "Location", ControlType: "Edit" }) {
        Sleep 100
        Send "{Tab}"
        return
    }
}

; Shift + C : Make Recurring - Recurring
+C:: {
    Appt_RunWithLoading("Recurring", (*) => (
        IsNewOutlookActive()
            ? (Appt_ClickSaved("Recurring") || Appt_PopoverInvokeFirst([{ Name: "Make recurring", ControlType: "Button" }, { Name: "recurring",
                matchmode: "Substring", ControlType: "Button" }]) || (ShowCenteredOverlay_Utils(
                    "❌ Appointment: Recurring not found", 1400, BANNER_ACCENT_ERROR), false))
            : (false)
    ))
    if IsNewOutlookActive()
        return
    if Appt_ClickSaved("Recurring")
        return
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)
        btn := root.FindFirst({ AutomationId: "4364", ControlType: "Button" })
        if !btn
            btn := root.FindFirst({ Name: "Make Recurring", ControlType: "Button" })
        if btn
            btn.Invoke()
    } catch {
    }
}

; -------------------------------------------------------------------
; New Outlook Appointment command bar shortcuts + modals
; -------------------------------------------------------------------

; Shift + M : Teams meeting - Meeting
+M:: {
    if !IsNewOutlookActive()
        return
    Appt_RunWithLoading("Teams meeting", (*) => (
        Appt_ClickSaved("Teams")
        || Appt_ClickInCommandBar([{ Name: "Teams meeting", matchmode: "Substring", ControlType: "Button" }, { Name: "Teams",
            matchmode: "Substring", ControlType: "Button" }]) || Appt_ClickAny([{ Name: "Teams meeting",
                matchmode: "Substring",
                ControlType: "Button" }, { Name: "Teams", matchmode: "Substring", ControlType: "Button" }
            ]) || (ShowCenteredOverlay_Utils("❌ Appointment: Teams meeting not found", 1400,
                BANNER_ACCENT_ERROR),
            false)
    ))
}

; Shift + U : Series (recurring) - sUper series
+U:: {
    if !IsNewOutlookActive()
        return
    Appt_RunWithLoading("Series", (*) => (
        Appt_ClickSaved("Series")
        || Appt_ClickInCommandBar([{ Name: "Series", ControlType: "Button" }, { Name: "Series", ControlType: "TabItem" }]) ||
        Appt_ClickAny([{ Name: "Series", ControlType: "Button" }, { Name: "Series", ControlType: "TabItem" }]) ||
        Appt_PopoverInvokeFirst([{ Name: "Make recurring", ControlType: "Button" }, { Name: "recurring",
            matchmode: "Substring",
            ControlType: "Button" }]) || (ShowCenteredOverlay_Utils("❌ Appointment: Series/Recurring not found",
                1400,
                BANNER_ACCENT_ERROR), false)
    ))
}

; Shift + V : Status/Busy selection modal - aVailability
+V:: {
    if !IsNewOutlookActive()
        return
    Appt_RunWithLoading("Status", (*) => (
        (choice := Appt_SelectFromModal("Appointment status", [{ k: "1", label: "Free" }, { k: "2", label: "Working elsewhere" }, { k: "3",
            label: "Tentative" }, { k: "4", label: "Busy" }, { k: "5", label: "Out of office" }
        ], "[1-5] Select  [Esc] Cancel"))
            ? (
                (target := (choice = "1") ? "Free"
                    : (choice = "2") ? "Working elsewhere"
                        : (choice = "3") ? "Tentative"
                            : (choice = "4") ? "Busy"
                                : "Out of office"),
                (savedId := (choice = "1") ? "StatusFree"
                    : (choice = "2") ? "StatusElsewhere"
                        : (choice = "3") ? "StatusTentative"
                            : (choice = "4") ? "StatusBusy"
                                : "StatusOoo"),
                Appt_ClickSaved(savedId) || Appt_OpenMenuAndPick([{ Name: "Free", ControlType: "Button" }, { Name: "Busy",
                    ControlType: "Button" }, { Name: "Tentative",
                        ControlType: "Button" }, { Name: "Working elsewhere", ControlType: "Button" }, { Name: "Out of office",
                            ControlType: "Button" }, { Name: "Free", matchmode: "Substring", ControlType: "Button" }, { Name: "Busy",
                                matchmode: "Substring", ControlType: "Button" }
                ], target)
            )
            : false
    ))
}

; Shift + Q : Reminder selection modal - Q for reminder freQuency
+Q:: {
    if !IsNewOutlookActive()
        return
    Appt_RunWithLoading("Reminder", (*) => (
        RemQ_Run()
    ))
}

RemQ_Run() {
    choice := Appt_SelectFromModal("Appointment reminder", [{ k: "1", label: "Don't remind me" }, { k: "2",
        label: "15 minutes before" }, { k: "3",
            label: "1 hour before" }, { k: "4", label: "12 hours before" }, { k: "5", label: "1 day before" }, { k: "6",
                label: "1 week before" }
    ], "[1-6] Select  [Esc] Cancel")
    if !choice
        return false

    target := (choice = "1") ? "Don't remind me"
        : (choice = "2") ? "15 minutes before"
            : (choice = "3") ? "1 hour before"
                : (choice = "4") ? "12 hours before"
                    : (choice = "5") ? "1 day before"
                        : "1 week before"
    savedId := (choice = "1") ? "RemNone"
        : (choice = "2") ? "Rem15m"
            : (choice = "3") ? "Rem1h"
                : (choice = "4") ? "Rem12h"
                    : (choice = "5") ? "Rem1d"
                        : "Rem1w"

    RemQ_VisualizeSelection("Reminder", target)
    if Appt_ClickSaved(savedId)
        return true
    return Appt_OpenMenuAndPick([{ Name: "Don't remind me", ControlType: "Button" }, { Name: "15 minutes before",
        ControlType: "Button" }, { Name: "1 week before", ControlType: "Button" }, { Name: "15 minutes",
            matchmode: "Substring",
            ControlType: "Button" }, { Name: "1 hour", matchmode: "Substring", ControlType: "Button" }, { Name: "12 hours",
                matchmode: "Substring", ControlType: "Button" }, { Name: "1 day", matchmode: "Substring",
                    ControlType: "Button" }, { Name: "Reminder",
                        matchmode: "Substring", ControlType: "Button" }
    ], target, 300)
}

RemQ_VisualizeSelection(label, target) {
    try StandardLoadingBar_Update("👁️ Appointment: selecting " label " → " target, BANNER_ACCENT_INTERMEDIATE)
    catch {
    }
    try ShowCenteredOverlay_Utils("👁️ Selecting " label ": " target, 900, BANNER_ACCENT_INTERMEDIATE)
    catch {
    }
    return true
}

; Shift + G : Category selection modal - cateGory
+G:: {
    if !IsNewOutlookActive()
        return
    Appt_RunWithLoading("Category", (*) => (
        (choice := Appt_SelectFromModal("Appointment category", [{ k: "1", label: "Aniversário" }, { k: "2",
            label: "Importante" }, { k: "3",
                label: "Pessoal" }], "[1-3] Select  [Esc] Cancel"))
            ? (
                (target := (choice = "1") ? "Aniversário"
                    : (choice = "2") ? "Importante"
                        : "Pessoal"),
                (savedId := (choice = "1") ? "CatBirthday"
                    : (choice = "2") ? "CatImportant"
                        : "CatPersonal"),
                Appt_ClickSaved(savedId) || Appt_OpenMenuAndPick([{ Name: "Aniversário", ControlType: "Button" }, { Name: "Importante",
                    ControlType: "Button" }, { Name: "Pessoal",
                        ControlType: "Button" }, { Name: "Category", matchmode: "Substring", ControlType: "Button" }, { Name: "Categories",
                            matchmode: "Substring", ControlType: "Button" }
                ], target)
            )
            : false
    ))
}

; Shift + 1 / Shift + 2 : Select time suggestions (New Outlook popover)
+1:: {
    Appt_RunWithLoading("Time suggestion 1", (*) => (
        IsNewOutlookActive()
            ? (Appt_ClickSaved("Suggestion1") || Appt_PopoverSelectTimeSuggestion(1) || (ShowCenteredOverlay_Utils(
                "❌ Appointment: Suggestion 1 not found", 1400, BANNER_ACCENT_ERROR), false))
            : (false)
    ))
}

+2:: {
    Appt_RunWithLoading("Time suggestion 2", (*) => (
        IsNewOutlookActive()
            ? (Appt_ClickSaved("Suggestion2") || Appt_PopoverSelectTimeSuggestion(2) || (ShowCenteredOverlay_Utils(
                "❌ Appointment: Suggestion 2 not found", 1400, BANNER_ACCENT_ERROR), false))
            : (false)
    ))
}
