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

; Shift + S : prefer Send (meeting), else Save (Event form commands)
+S:: {
    Appt_RunWithLoading("Send/Save", (*) => (
        Appt_ClickInCommandBar([{ Name: "Send", ControlType: "Button" }])
        || Appt_ClickAny([{ Name: "Send", ControlType: "Button" }])
        || Appt_ClickInCommandBar([{ Name: "Save", ControlType: "Button" }])
        || Appt_ClickAny([{ Name: "Save", ControlType: "Button" }])
        || (ShowCenteredOverlay_Utils("❌ Appointment: Send/Save not found", 1400, BANNER_ACCENT_ERROR), false)
    ))
}

; Shift + D : Start date (combo) - Start Date
+D:: {
    isNew := false
    try isNew := IsNewOutlookActive()

    Appt_RunWithLoading("Start date", (*) => (
        isNew
            ? (Appt_PopoverInvokeFirst([{ Name: "Start date", ControlType: "ComboBox" }, { Name: "Start date", Type: 50003 }, { Name: "Start date",
                ControlType: "Button" }, { Name: "Start date", Type: 50000 }, { AutomationId: "DatePicker", matchmode: "Substring" }
            ]) || (ShowCenteredOverlay_Utils("❌ Appointment: Start date not found", 1400, BANNER_ACCENT_ERROR), false))
            : (Outlook_ClickStartDate(), true)
    ))
}

; Shift + P : Start date picker - Picker
+P:: {
    ; New Outlook: repurpose Shift+P to Private toggle modal (date picker concept removed).
    if IsNewOutlookActive() {
        Appt_RunWithLoading("Private", (*) => (
            (choice := Appt_SelectFromModal("Appointment privacy", [{ k: "1", label: "Private" }, { k: "2", label: "Not private" }],
            "[1-2] Select  [Esc] Cancel"))
                ? (
                    (choice = "1")
                        ? Appt_OpenMenuAndPick([{ Name: "Private", ControlType: "Button" }, { Name: "Not private",
                            ControlType: "Button" }, { Name: "Private", matchmode: "Substring", ControlType: "Button" }, { Name: "Not private",
                                matchmode: "Substring", ControlType: "Button" }
                        ], "Private")
                        : Appt_OpenMenuAndPick([{ Name: "Private", ControlType: "Button" }, { Name: "Not private",
                            ControlType: "Button" }, { Name: "Private", matchmode: "Substring", ControlType: "Button" }, { Name: "Not private",
                                matchmode: "Substring", ControlType: "Button" }
                        ], "Not private")
                )
                : false
        ))
        return
    }
    Outlook_ClickStartDatePicker()
}

; Shift + T : Start time (combo) - Time
+T:: {
    Appt_RunWithLoading("Start time", (*) => (
        IsNewOutlookActive()
            ? (Appt_PopoverFocusFirst([{ Name: "Start time", ControlType: "ComboBox" }, { Name: "Start time", Type: 50003 }, { AutomationId: "ComboBox",
                matchmode: "Substring" }]) || (ShowCenteredOverlay_Utils("❌ Appointment: Start time not found", 1400,
                    BANNER_ACCENT_ERROR), false))
            : (Outlook_ClickStartTime(), true)
    ))
}

; Shift + E : End date (combo) - End Date
+E:: {
    Appt_RunWithLoading("End time", (*) => (
        IsNewOutlookActive()
            ? (Appt_PopoverFocusFirst([{ Name: "End time", ControlType: "ComboBox" }, { Name: "End time", Type: 50003 }]) ||
            (ShowCenteredOverlay_Utils("❌ Appointment: End time not found", 1400, BANNER_ACCENT_ERROR), false))
            : (Outlook_ClickEndDate(), true)
    ))
}

; Shift + H : Scheduler / Scheduling assistant (New Outlook) or End time (classic)
+H:: {
    if IsNewOutlookActive() {
        Appt_RunWithLoading("Scheduler", (*) => (
            Appt_ClickAny([{ Name: "Scheduler", ControlType: "Button" }, { Name: "Scheduling assistant", matchmode: "Substring",
                ControlType: "Button" }, { Name: "Scheduling", matchmode: "Substring", ControlType: "Button" }
            ]) || (ShowCenteredOverlay_Utils("❌ Appointment: Scheduler not found", 1400, BANNER_ACCENT_ERROR), false)
        ))
        return
    }
    Outlook_ClickEndTime()
}

; Shift + A : All day checkbox - All Day
+A:: {
    if IsNewOutlookActive() {
        Appt_RunWithLoading("All day", (*) => (
            Appt_PopoverToggleFirst([{ Name: "All day", ControlType: "CheckBox" }, { Name: "All day", Type: 50002 },
            ; New Outlook exposes this as a switch (button) with a stable AutomationId (e.g. Toggle9777).
            { AutomationId: "Toggle", matchmode: "Substring", ControlType: "Button" }, { AutomationId: "Toggle",
                matchmode: "Substring", Type: 50000 }, { Name: "All day", ControlType: "Button" }, { Name: "All day",
                    Type: 50000 }
            ]) || (ShowCenteredOverlay_Utils("❌ Appointment: All day not found", 1400, BANNER_ACCENT_ERROR), false)
        ))
        return
    }
    ; Classic fallback (existing behavior)
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)
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
    if FocusOutlookField({ Name: "Add title", ControlType: "Edit" })
        return
    if FocusOutlookField({ AutomationId: "4100" }) ; Title
        return
    if FocusOutlookField({ Name: "Title", ControlType: "Edit" })
        return
}

; Shift + R : Required / To field - Required
+R:: {
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
    root := Appt_GetRootActive()
    if !root
        return false
    ; Primary targeting (New Outlook): "Go to previous day <date>" / "Go to next day <date>"
    candidates := isNext
        ? [{ Name: "Go to next", matchmode: "Substring", ControlType: "Button" }, { Name: "go to next", matchmode: "Substring",
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
    return Appt_ClickInCommandBar([{ Name: "Back", ControlType: "Button" }])
    || Appt_ClickAny([{ Name: "Back", ControlType: "Button" }])
}

Appt_SchedulerClickOptions() {
    return Appt_ClickInCommandBar([{ Name: "Options", ControlType: "Button" }, { Name: "Options", matchmode: "Substring",
        ControlType: "Button" }])
    || Appt_ClickAny([{ Name: "Options", ControlType: "Button" }, { Name: "Options", matchmode: "Substring",
        ControlType: "Button" }])
}

Appt_SchedulerClickAddAttendee(isOptional) {
    name := isOptional ? "Add optional attendee" : "Add required attendee"
    return Appt_ClickAny([{ Name: name, ControlType: "Button" }, { Name: name, matchmode: "Substring", ControlType: "Button" }])
}

Appt_SchedulerFocusDateTimeControl(kind) {
    ; Focus core controls in scheduler/editor view by visible names.
    if (kind = "start_date")
        return Appt_PopoverFocusFirst([{ Name: "Start date", ControlType: "ComboBox" }, { Name: "Start date", Type: 50003 }])
    if (kind = "start_time")
        return Appt_PopoverFocusFirst([{ Name: "Start time", ControlType: "ComboBox" }, { Name: "Start time", Type: 50003 }])
    if (kind = "end_time")
        return Appt_PopoverFocusFirst([{ Name: "End time", ControlType: "ComboBox" }, { Name: "End time", Type: 50003 }])
    if (kind = "all_day")
        return Appt_PopoverInvokeFirst([{ Name: "All day", ControlType: "CheckBox" }, { Name: "All day", Type: 50002 }, { Name: "All day",
            ControlType: "Button" }])
    if (kind = "time_zone")
        return Appt_ToggleOrClickAny([{ Name: "Show event time zones", matchmode: "Substring", ControlType: "Button" }, { Name: "Time zone",
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
        (Appt_IsSchedulerView() ? (Appt_ClickSchedulerSuggestionNav(true) || Appt_ClickDayNav(true)) : Appt_ClickDayNav(
            true))
        || (ShowCenteredOverlay_Utils("❌ Appointment: Next not found", 1400, BANNER_ACCENT_ERROR), false)
    ))
}

; Shift + Y : Today
+Y:: {
    if !IsNewOutlookActive()
        return
    Appt_RunWithLoading("Today", (*) => (
        Appt_ClickAny([{ Name: "Today", ControlType: "Button" }, { Name: "Today", matchmode: "Substring", ControlType: "Button" }])
        || (ShowCenteredOverlay_Utils("❌ Appointment: Today not found", 1400, BANNER_ACCENT_ERROR), false)
    ))
}

; Shift + F : Current date header button
+F:: {
    if !IsNewOutlookActive()
        return
    Appt_RunWithLoading("Date", (*) => (
        Appt_ClickAny([{ Name: "Thu", matchmode: "Substring", ControlType: "Button" }, { Name: "Apr", matchmode: "Substring",
            ControlType: "Button" }, { Name: "Week", matchmode: "Substring", ControlType: "Button" }
        ]) || (ShowCenteredOverlay_Utils("❌ Appointment: Date header not found", 1400, BANNER_ACCENT_ERROR), false)
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
        Appt_SchedulerClickAddAttendee(false) || (ShowCenteredOverlay_Utils("❌ Appointment: Add required not found",
            1400, BANNER_ACCENT_ERROR), false)
    ))
}

; Alt + O : Add optional attendee
!o:: {
    if !IsNewOutlookActive()
        return
    Appt_RunWithLoading("Add optional", (*) => (
        Appt_SchedulerClickAddAttendee(true) || (ShowCenteredOverlay_Utils("❌ Appointment: Add optional not found",
            1400, BANNER_ACCENT_ERROR), false)
    ))
}

; Shift + B : Body (from Location) - Body
+B:: {
    Appt_RunWithLoading("Body", (*) => (
        IsNewOutlookActive()
            ? (Appt_FocusBodyField_NewOutlook() || (ShowCenteredOverlay_Utils("❌ Appointment: Body not found", 1400,
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
            ? (Appt_PopoverInvokeFirst([{ Name: "Make recurring", ControlType: "Button" }, { Name: "recurring",
                matchmode: "Substring", ControlType: "Button" }]) || (ShowCenteredOverlay_Utils(
                    "❌ Appointment: Recurring not found", 1400, BANNER_ACCENT_ERROR), false))
            : (false)
    ))
    if IsNewOutlookActive()
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
        Appt_ClickInCommandBar([{ Name: "Teams meeting", matchmode: "Substring", ControlType: "Button" }, { Name: "Teams",
            matchmode: "Substring", ControlType: "Button" }]) || Appt_ClickAny([{ Name: "Teams meeting", matchmode: "Substring",
                ControlType: "Button" }, { Name: "Teams", matchmode: "Substring", ControlType: "Button" }
            ]) || (ShowCenteredOverlay_Utils("❌ Appointment: Teams meeting not found", 1400, BANNER_ACCENT_ERROR),
            false)
    ))
}

; Shift + U : Series (recurring) - sUper series
+U:: {
    if !IsNewOutlookActive()
        return
    Appt_RunWithLoading("Series", (*) => (
        Appt_ClickInCommandBar([{ Name: "Series", ControlType: "Button" }, { Name: "Series", ControlType: "TabItem" }]) ||
        Appt_ClickAny([{ Name: "Series", ControlType: "Button" }, { Name: "Series", ControlType: "TabItem" }]) ||
        Appt_PopoverInvokeFirst([{ Name: "Make recurring", ControlType: "Button" }, { Name: "recurring", matchmode: "Substring",
            ControlType: "Button" }]) || (ShowCenteredOverlay_Utils("❌ Appointment: Series/Recurring not found", 1400,
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
                Appt_OpenMenuAndPick([{ Name: "Free", ControlType: "Button" }, { Name: "Busy", ControlType: "Button" }, { Name: "Tentative",
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
    choice := Appt_SelectFromModal("Appointment reminder", [{ k: "1", label: "Don't remind me" }, { k: "2", label: "15 minutes before" }, { k: "3",
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

    RemQ_VisualizeSelection("Reminder", target)
    return Appt_OpenMenuAndPick([{ Name: "Don't remind me", ControlType: "Button" }, { Name: "15 minutes before",
        ControlType: "Button" }, { Name: "1 week before", ControlType: "Button" }, { Name: "15 minutes", matchmode: "Substring",
            ControlType: "Button" }, { Name: "1 hour", matchmode: "Substring", ControlType: "Button" }, { Name: "12 hours",
                matchmode: "Substring", ControlType: "Button" }, { Name: "1 day", matchmode: "Substring", ControlType: "Button" }, { Name: "Reminder",
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
        (choice := Appt_SelectFromModal("Appointment category", [{ k: "1", label: "Aniversário" }, { k: "2", label: "Importante" }, { k: "3",
            label: "Pessoal" }], "[1-3] Select  [Esc] Cancel"))
            ? (
                (target := (choice = "1") ? "Aniversário"
                    : (choice = "2") ? "Importante"
                        : "Pessoal"),
                Appt_OpenMenuAndPick([{ Name: "Aniversário", ControlType: "Button" }, { Name: "Importante", ControlType: "Button" }, { Name: "Pessoal",
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
            ? (Appt_PopoverSelectTimeSuggestion(1) || (ShowCenteredOverlay_Utils(
                "❌ Appointment: Suggestion 1 not found", 1400, BANNER_ACCENT_ERROR), false))
            : (false)
    ))
}

+2:: {
    Appt_RunWithLoading("Time suggestion 2", (*) => (
        IsNewOutlookActive()
            ? (Appt_PopoverSelectTimeSuggestion(2) || (ShowCenteredOverlay_Utils(
                "❌ Appointment: Suggestion 2 not found", 1400, BANNER_ACCENT_ERROR), false))
            : (false)
    ))
}
