; =============================================================================
; Shift keys module: outlook_appointment_uia.ahk
; Outlook appointment UIA state checking
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

; =============================================================================
; Outlook Appointment Control State Checking Functions (UIA-based)
; =============================================================================

Outlook_CheckPrivacyState() {
    try {
        win := WinExist("A")
        if (!win) {
            return ""
        }
        root := UIA.ElementFromHandle(win)
        Sleep 100  ; Allow UI to update

        ; Look for Private checkbox - typically has AutomationId or specific Name
        checkbox := root.FindFirst({ AutomationId: "4227", ControlType: "CheckBox" })
        if (!checkbox) {
            checkbox := root.FindFirst({ Name: "Private", ControlType: "CheckBox" })
        }

        if (checkbox) {
            ; Check if checkbox is checked
            isChecked := checkbox.GetCurrentPropertyValue(UIA.Property.ToggleToggleState)
            ; ToggleState: 0 = Off, 1 = On
            return (isChecked = 1) ? "On" : "Off"
        }
    } catch Error {
        ; Silently fail - return empty string
    }
    return ""
}

Outlook_CheckAllDayState() {
    try {
        win := WinExist("A")
        if (!win) {
            return ""
        }
        root := UIA.ElementFromHandle(win)
        Sleep 100  ; Allow UI to update

        checkbox := root.FindFirst({ AutomationId: "4226", ControlType: "CheckBox" })
        if (!checkbox) {
            checkbox := root.FindFirst({ Name: "All day", ControlType: "CheckBox" })
        }

        if (checkbox) {
            isChecked := checkbox.GetCurrentPropertyValue(UIA.Property.ToggleToggleState)
            return (isChecked = 1) ? "Yes" : "No"
        }
    } catch Error {
        ; Silently fail
    }
    return ""
}

Outlook_CheckStatusState() {
    try {
        win := WinExist("A")
        if (!win) {
            return ""
        }
        root := UIA.ElementFromHandle(win)
        Sleep 100  ; Allow UI to update

        ; Look for Status dropdown/button - may need to find by AutomationId or Name
        statusControl := root.FindFirst({ AutomationId: "4356", ControlType: "Button" })
        if (!statusControl) {
            statusControl := root.FindFirst({ Name: "Busy", ControlType: "Button" })
        }
        if (!statusControl) {
            ; Try to find any control with Status-related names
            statusControl := root.FindFirst({ Name: "Free", ControlType: "Button" })
        }

        if (statusControl) {
            ; Get the text/value of the status control
            statusText := statusControl.GetCurrentPropertyValue(UIA.Property.Name)
            if (InStr(statusText, "Free", false)) {
                return "Free"
            } else if (InStr(statusText, "Busy", false)) {
                return "Busy"
            } else if (InStr(statusText, "Out of office", false) || InStr(statusText, "Out of Office", false)) {
                return "Out of office"
            }
        }
    } catch Error {
        ; Silently fail
    }
    return ""
}

Outlook_CheckCategoryState() {
    try {
        win := WinExist("A")
        if (!win) {
            return ""
        }
        root := UIA.ElementFromHandle(win)
        Sleep 100  ; Allow UI to update

        ; Look for Category control - may be a button or dropdown
        categoryControl := root.FindFirst({ AutomationId: "4357", ControlType: "Button" })
        if (!categoryControl) {
            categoryControl := root.FindFirst({ Name: "Categorize", ControlType: "Button" })
        }

        if (categoryControl) {
            ; Try to get category text/value
            categoryText := categoryControl.GetCurrentPropertyValue(UIA.Property.Name)
            if (InStr(categoryText, "Important", false)) {
                return "Important"
            } else if (InStr(categoryText, "Personal", false)) {
                return "Personal"
            }
        }
    } catch Error {
        ; Silently fail
    }
    return ""
}

Outlook_CheckReminderState() {
    try {
        win := WinExist("A")
        if (!win) {
            return ""
        }
        root := UIA.ElementFromHandle(win)
        Sleep 100  ; Allow UI to update

        ; Look for Reminder dropdown/field
        reminderControl := root.FindFirst({ AutomationId: "4358", ControlType: "ComboBox" })
        if (!reminderControl) {
            reminderControl := root.FindFirst({ Name: "Reminder", ControlType: "ComboBox" })
        }
        if (!reminderControl) {
            reminderControl := root.FindFirst({ Name: "Reminder", ControlType: "Edit" })
        }

        if (reminderControl) {
            reminderText := reminderControl.GetCurrentPropertyValue(UIA.Property.Value)
            if (reminderText) {
                return reminderText
            }
        }
    } catch Error {
        ; Silently fail
    }
    return ""
}

RunOutlookAppointmentWizard() {
    ; New Outlook only. Requires an active New Outlook appointment window.
    if !IsNewOutlookActive() || !IsOutlookAppointmentActive() {
        ShowCenteredOverlay_Utils("❌ Appointment Wizard: open a New event window first", 1700, BANNER_ACCENT_ERROR)
        return
    }
    global g_ApptWizardMainHwnd
    g_ApptWizardMainHwnd := WinExist("A")

    ; STEP 1/5 – Status
    c1 := Appt_SelectFromModal("Wizard 1/5: status", [{ k: "1", label: "🟢 Free" }, { k: "2", label: "🟡 Tentative" }, { k: "3",
        label: "🔴 Busy" }, { k: "4", label: "🔴 Out of office" }
    ], "[1-4] Select  [Esc] Cancel")
    if (c1 = "")
        return
    status := (c1 = "1") ? "Free" : (c1 = "2") ? "Tentative" : (c1 = "3") ? "Busy" : "Out of office"

    ; STEP 2/5 – Privacy
    c2 := Appt_SelectFromModal("Wizard 2/5: privacy", [{ k: "1", label: "🔓 Not private" }, { k: "2", label: "🔒 Private" }],
    "[1-2] Select  [Esc] Cancel")
    if (c2 = "")
        return
    privacy := (c2 = "2") ? "Private" : "Not private"

    ; STEP 3/5 – Category
    c4 := Appt_SelectFromModal("Wizard 3/5: category", [{ k: "1", label: "🚫 None" }, { k: "2", label: "⭐ Important" }, { k: "3",
        label: "👤 Personal" }], "[1-3] Select  [Esc] Cancel")
    if (c4 = "")
        return
    ; New Outlook UI is localized (PT-BR) for categories in this setup.
    category := (c4 = "1") ? "" : (c4 = "2") ? "Importante" : "Pessoal"

    ; STEP 4/5 – Reminder (align with new Appointment menu labels we already use)
    c5 := Appt_SelectFromModal("Wizard 4/5: reminder", [{ k: "1", label: "🔕 Don't remind me" }, { k: "2", label: "⏰ 15 minutes before" }, { k: "3",
        label: "⏰ 1 hour before" }, { k: "4", label: "⏰ 12 hours before" }, { k: "5", label: "🗓️ 1 day before" }, { k: "6",
            label: "📅 1 week before" }
    ], "[1-6] Select  [Esc] Cancel")
    if (c5 = "")
        return
    reminder := (c5 = "1") ? "Don't remind me"
        : (c5 = "2") ? "15 minutes before"
            : (c5 = "3") ? "1 hour before"
                : (c5 = "4") ? "12 hours before"
                    : (c5 = "5") ? "1 day before"
                        : "1 week before"

    ; STEP 5/5 – All-day (final)
    c3 := Appt_SelectFromModal("Wizard 5/5: all-day", [{ k: "1", label: "⏰ Timed (All-day OFF)" }, { k: "2", label: "📅 All-day ON" }],
    "[1-2] Select  [Esc] Cancel")
    if (c3 = "")
        return
    allDayOn := (c3 = "2")

    ApptWizard_ApplySelection(status, privacy, allDayOn, category, reminder)
}

; Shift + w → Cascaded text wizard for Outlook Appointment
+w:: {
    if !IsOutlookAppointmentActive() || !IsNewOutlookActive() {
        ShowCenteredOverlay_Utils("❌ Appointment Wizard: open a New event window first", 1700, BANNER_ACCENT_ERROR)
        return
    }
    RunOutlookAppointmentWizard()
}

ApptWizard_ApplySelection(status, privacy, allDayOn, category, reminder) {
    global APPT_WIZARD_STEP_DELAY_MS, APPT_WIZARD_REMINDER_PRECLICK_DELAY_MS
    if !IsSet(APPT_WIZARD_STEP_DELAY_MS)
        APPT_WIZARD_STEP_DELAY_MS := 2200
    if !IsSet(APPT_WIZARD_REMINDER_PRECLICK_DELAY_MS)
        APPT_WIZARD_REMINDER_PRECLICK_DELAY_MS := 820

    try StandardLoadingBar_Show("⏳ Wizard: applying…", BANNER_ACCENT_INTERMEDIATE, { passive: false, centerOnHwnd: 0,
        textWidth: 640, fontSize: 17 })
    catch {
    }
    try {
        try StandardLoadingBar_Update("🔄 Wizard: Status → " status, BANNER_ACCENT_INTERMEDIATE)
        Appt_OpenMenuAndPick([{ Name: "Free", ControlType: "Button" }, { Name: "Busy", ControlType: "Button" }, { Name: "Tentative",
            ControlType: "Button" }, { Name: "Working elsewhere", ControlType: "Button" }, { Name: "Out of office",
                ControlType: "Button" }, { Name: "Free", matchmode: "Substring", ControlType: "Button" }, { Name: "Busy",
                    matchmode: "Substring", ControlType: "Button" }
        ], status)
        Sleep APPT_WIZARD_STEP_DELAY_MS

        try StandardLoadingBar_Update("🔄 Wizard: Privacy → " privacy, BANNER_ACCENT_INTERMEDIATE)
        Appt_OpenMenuAndPick([{ Name: "Private", ControlType: "Button" }, { Name: "Not private", ControlType: "Button" }, { Name: "Private",
            matchmode: "Substring", ControlType: "Button" }, { Name: "Not private", matchmode: "Substring", ControlType: "Button" }
        ], privacy)
        Sleep APPT_WIZARD_STEP_DELAY_MS

        if (category != "") {
            try StandardLoadingBar_Update("🔄 Wizard: Category → " category, BANNER_ACCENT_INTERMEDIATE)
            Appt_OpenMenuAndPick([{ Name: "Aniversário", ControlType: "Button" }, { Name: "Importante", ControlType: "Button" }, { Name: "Pessoal",
                ControlType: "Button" }, { Name: "Important", ControlType: "Button" }, { Name: "Personal", ControlType: "Button" }, { Name: "Category",
                    matchmode: "Substring", ControlType: "Button" }, { Name: "Categories", matchmode: "Substring",
                        ControlType: "Button" }
            ], category)
            Sleep APPT_WIZARD_STEP_DELAY_MS
        }

        try StandardLoadingBar_Update("🔄 Wizard: Reminder → " reminder, BANNER_ACCENT_INTERMEDIATE)
        Appt_OpenMenuAndPick([{ Name: "Don't remind me", ControlType: "Button" }, { Name: "15 minutes before",
            ControlType: "Button" }, { Name: "1 week before", ControlType: "Button" }, { Name: "15 minutes", matchmode: "Substring",
                ControlType: "Button" }, { Name: "1 hour", matchmode: "Substring", ControlType: "Button" }, { Name: "12 hours",
                    matchmode: "Substring", ControlType: "Button" }, { Name: "1 day", matchmode: "Substring",
                        ControlType: "Button" }, { Name: "Reminder", matchmode: "Substring", ControlType: "Button" }
        ], reminder, APPT_WIZARD_REMINDER_PRECLICK_DELAY_MS)
        Sleep APPT_WIZARD_STEP_DELAY_MS

        try StandardLoadingBar_Update("🔄 Wizard: All-day → " (allDayOn ? "On" : "Off"), BANNER_ACCENT_INTERMEDIATE)
        ApptWizard_SetAllDay(allDayOn)
        Sleep APPT_WIZARD_STEP_DELAY_MS

        try StandardLoadingBar_Update("✅ Wizard: applied", BANNER_ACCENT_SUCCESS)
        try StandardLoadingBar_Hide(700)
    } catch {
        try StandardLoadingBar_Update("❌ Wizard: failed", BANNER_ACCENT_ERROR)
        try StandardLoadingBar_Hide(1200)
    }
}

ApptWizard_FocusTitleField() {
    try {
        ; Clear any overlay that might keep focus.
        StandardLoadingBar_Hide(0)
    } catch {
    }
    ; Close any open context menu/popover that may be holding focus.
    try Send "{Esc}"
    try Send "{Esc}"
    Sleep 260
    global g_ApptWizardMainHwnd
    if IsSet(g_ApptWizardMainHwnd) && g_ApptWizardMainHwnd
        try WinActivate("ahk_id " g_ApptWizardMainHwnd)
    ok := false
    try ok := ApptWizard_TryFocusTitleByCriteria(g_ApptWizardMainHwnd, { Name: "Add title", ControlType: "Edit" })
    catch {
    }
    if !ok {
        try ok := ApptWizard_TryFocusTitleByCriteria(g_ApptWizardMainHwnd, { Name: "Add title", Type: 50004 })
        catch {
        }
    }
    if !ok {
        ; Try alternate label (some builds expose Title vs Add title).
        try ok := ApptWizard_TryFocusTitleByCriteria(g_ApptWizardMainHwnd, { Name: "Title", ControlType: "Edit" })
        catch {
        }
    }
    if !ok {
        try ok := ApptWizard_TryFocusTitleByCriteria(g_ApptWizardMainHwnd, { AutomationId: "4100" })
        catch {
        }
    }

    ; Last-resort: keyboard focus traversal (some builds expose no Edit controls).
    if !ok {
        ok := ApptWizard_FocusTitleField_ByTabbing(28)
    }

    if !ok {
        try ShowCenteredOverlay_Utils("⚠️ Wizard: Title field not found", 1200, BANNER_ACCENT_INTERMEDIATE)
        catch {
        }
    }
    return ok
}

ApptWizard_IsTitleName(name) {
    if (name = "")
        return false
    n := StrLower(name)
    return InStr(n, "add title") || InStr(n, "title") || InStr(n, "adicionar titulo") || InStr(n, "adicionar título")
}

ApptWizard_IsBodyEditorName(name) {
    if (name = "")
        return false
    n := StrLower(name)
    return InStr(n, "type {0} to insert files and more") || InStr(n, "insert files and more")
}

ApptWizard_IsTitleFocused() {
    try {
        fe := UIA.GetFocusedElement()
        if !fe
            return false
        nm := ""
        try nm := fe.Name
        return ApptWizard_IsTitleName(nm) && !ApptWizard_IsBodyEditorName(nm)
    } catch {
    }
    return false
}

ApptWizard_TryFocusTitleByCriteria(hwnd, criteria) {
    try {
        if !hwnd
            hwnd := WinExist("A")
        root := UIA.ElementFromHandle(hwnd)
        if !root
            return false
        ctrl := root.FindFirst(criteria)
        if !ctrl
            return false
        try ctrl.SetFocus()
        Sleep 220
        if ApptWizard_IsTitleFocused()
            return true
        try ctrl.Click()
        Sleep 220
        return ApptWizard_IsTitleFocused()
    } catch {
    }
    return false
}

FocusOutlookFieldOnHwnd(hwnd, criteria) {
    try {
        if !hwnd
            hwnd := WinExist("A")
        root := UIA.ElementFromHandle(hwnd)
        ctrl := root.FindFirst(criteria)
        if ctrl {
            ctrl.SetFocus()
            return true
        }
    } catch {
    }
    return false
}

ApptWizard_FocusTitleField_ByTabbing(maxSteps := 24) {
    try {
        ; Anchor: focus Save in command bar (stable) then tab forward.
        tb := Appt_FindCommandBar()
        if tb {
            btn := 0
            try btn := tb.FindFirst({ Name: "Save", ControlType: "Button" })
            if btn {
                try btn.SetFocus()
                Sleep 220
            }
        }
    } catch {
    }

    loop maxSteps {
        fe := 0, name := "", aid := "", ty := ""
        try fe := UIA.GetFocusedElement()
        if fe {
            try name := fe.Name
            try aid := fe.AutomationId
            try ty := fe.Type
        }

        ; Match both EN/PT variants.
        if (name != "") {
            if InStr(name, "Add title", false) || InStr(name, "Title", false) || InStr(name, "Adicionar título", false) ||
            InStr(name, "Adicionar titulo", false) {
                ; Ensure caret by clicking focused element if possible.
                try fe.Click()
                catch {
                }
                return true
            }
            if ApptWizard_IsBodyEditorName(name) {
                ; If we landed in the body editor, backtrack once to avoid typing spaces there.
                Send "+{Tab}"
                Sleep 240
                if ApptWizard_IsTitleFocused()
                    return true
                continue
            }
        }
        Send "{Tab}"
        Sleep 240
    }
    return false
}

ApptWizard_SetAllDay(desiredOn) {
    pop := Appt_OpenPopoverIfNeeded()
    if !pop
        return false
    el := 0
    try el := pop.FindFirst({ Name: "All day", Type: 50002 })
    catch {
    }
    if !el {
        try el := pop.FindFirst({ Name: "All day", Type: 50000 })
        catch {
        }
    }
    if !el {
        try el := pop.FindFirst({ Name: "All day", matchmode: "Substring" })
        catch {
        }
    }
    if !el
        return false

    state := ""
    try {
        if el.IsTogglePatternAvailable
            state := el.TogglePattern.CurrentToggleState
    } catch {
    }
    ; ToggleState: 0=Off, 1=On (typical)
    if (state != "" && ((state = 1) = desiredOn))
        return true

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
}
