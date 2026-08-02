; =============================================================================
; Shift keys module: hotif_gmail.ahk
; Gmail hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

Gmail_ToggleReadStatus(uia) {
    readPattern := "i)^(Mark as read|Marcar como lida|Marcar como lido)$"
    unreadPattern := "i)^(Mark as unread|Marcar como n[oÃ³] lida|Marcar como n[oÃ³] lido)$"
    if (btn := WaitForButton(uia, readPattern, 1000)) {
        btn.Invoke()
        return true
    }
    if (btn := WaitForButton(uia, unreadPattern, 1000)) {
        btn.Invoke()
        return true
    }
    return false
}

Gmail_CheckboxToggleState(el) {
    try {
        if el.GetPropertyValue(UIA.Property.IsTogglePatternAvailable)
            return el.TogglePattern.CurrentToggleState
    } catch {
    }
    return -1
}

; Set toolbar master Select checkbox: wantChecked 1=select all, 0=deselect.
; Screen Click("left") does not toggle this control; TogglePattern does.
Gmail_SetMasterSelect(uia, wantChecked) {
    selectBtn := 0
    try selectBtn := uia.FindElement({ Name: "Select", Type: "Button" })
    if (!selectBtn)
        return false

    checkbox := 0
    try checkbox := selectBtn.FindElement({ Type: "CheckBox" })

    if (checkbox) {
        if (Gmail_CheckboxToggleState(checkbox) = wantChecked)
            return true

        try {
            if checkbox.GetPropertyValue(UIA.Property.IsTogglePatternAvailable) {
                checkbox.TogglePattern.Toggle()
                Sleep 150
                if (Gmail_CheckboxToggleState(checkbox) = wantChecked)
                    return true
            }
        } catch {
        }

        try {
            checkbox.ControlClick("left")
            Sleep 150
            if (Gmail_CheckboxToggleState(checkbox) = wantChecked)
                return true
        } catch {
        }
    }

    ; Fallback: Gmail *a select-all / *n deselect-all
    try {
        Send(wantChecked ? "{*}a" : "{*}n")
        Sleep 200
        return true
    } catch {
    }
    return false
}

#HotIf WinActive("Gmail")

; Shift + I : Go to main inbox - Inbox
+i:: Send("gi")

; Shift + U : Go to updates - Updates
+u::
{
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 300 ; Give UIA time to attach

        ; Find the "Updates" tab (Name may start with "Updates" or include message counts)
        updatesButton := uia.FindElement({ Name: "Updates", Type: "TabItem", matchmode: "Substring" })

        if (updatesButton) {
            updatesButton.Click()
        }
        else {
            MsgBox "Could not find the 'Updates' button."
        }
    }
    catch Error as e {
        MsgBox "An error occurred: " e.Message
    }
}

; Shift + F : Go to forums - Forums
+f::
{
    try
    {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 300

        ; Try English and Portuguese names for the Forums tab
        forumsButton := uia.FindElement({ Name: "Forums", Type: "TabItem", matchmode: "Substring" })
        if (!forumsButton)
            forumsButton := uia.FindElement({ Name: "FÃ³runs", Type: "TabItem", matchmode: "Substring" })

        if (forumsButton) {
            forumsButton.Click()
        }
        else {
            MsgBox "Could not find the 'Forums' button."
        }
    }
    catch Error as e {
        MsgBox "An error occurred: " e.Message
    }
}

; Shift + R : Toggle read / unread - Read
+r::
{
    try
    {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 300
        if (!Gmail_ToggleReadStatus(uia))
            MsgBox "Could not find a 'Mark as read' or 'Mark as unread' button."
    }
    catch Error as e {
        MsgBox "An error occurred: " e.Message
    }
}

; Shift + E : Select all visible, mark read/unread, then deselect
+e::
{
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 300

        if (!Gmail_SetMasterSelect(uia, 1)) {
            MsgBox "Could not find the 'Select' checkbox."
            return
        }

        Sleep 400
        if (!Gmail_ToggleReadStatus(uia)) {
            MsgBox "Could not find a 'Mark as read' or 'Mark as unread' button."
            return
        }

        Sleep 300
        Gmail_SetMasterSelect(uia, 0)
    }
    catch Error as e {
        MsgBox "An error occurred: " e.Message
    }
}

; Shift + P : Previous conversation - Previous
+p:: Send("p")

; Shift + N : Next conversation - Next
+n:: Send("n")

; Shift + A : Archive conversation - Archive
+a:: Send("e")

; Shift + S : Select conversation - Select
+s:: Send("x")

; Shift + R : Reply - Reply (Note: conflicts with Read/unread, but Reply is more common)
; Actually, let me use Y for Reply to avoid conflict
+y:: Send("r")

; Shift + A : Reply all - All (conflicts with Archive!)
; Let me use G for Reply all (G for Group/all)
+g:: Send("a")

; Shift + W : Forward - Forward
+w:: Send("f")

; Shift + S : Star/unstar conversation - Star (conflicts with Select!)
; Let me use T for Star (T for sTar)
+t:: Send("s")

; Shift + D : Delete - Delete
+d:: Send("#")

; Shift + X : Report as spam - Spam
+x:: Send("!")

; Shift + C : Compose new email - Compose
+c:: Send("c")

; Shift + M : Move to folder - Move
+m:: Send("v")

; Shift + H : Show keyboard shortcuts help - Help
+h:: Send("?")

; Shift + B : Click inbox button - Button
+b::
{
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 300 ; Give UIA time to attach

        ; Try to find inbox link by name (may include unread count)
        inboxLink := uia.FindElement({ Name: "Inbox", Type: "50005", matchmode: "Substring" })

        ; Fallback: try by ClassName
        if (!inboxLink) {
            inboxLink := uia.FindElement({ ClassName: "J-Ke n0", Type: "50005" })
        }

        ; Fallback: try by Value (URL)
        if (!inboxLink) {
            inboxLink := uia.FindElement({ Value: "#inbox", Type: "50005", matchmode: "Substring" })
        }

        if (inboxLink) {
            inboxLink.Click()
        }
        else {
            MsgBox "Could not find the 'Inbox' button."
        }
    }
    catch Error as e {
        MsgBox "An error occurred: " e.Message
    }
}
