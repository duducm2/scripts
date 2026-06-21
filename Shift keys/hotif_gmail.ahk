; =============================================================================
; Shift keys module: hotif_gmail.ahk
; Gmail hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

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

        ; Regex patterns for the buttons (English & Portuguese)
        readPattern := "i)^(Mark as read|Marcar como lida|Marcar como lido)$"
        unreadPattern := "i)^(Mark as unread|Marcar como n[oÃ³] lida|Marcar como n[oÃ³] lido)$"

        ; Prefer clicking "Mark as read" if present; otherwise "Mark as unread"
        if (btn := WaitForButton(uia, readPattern, 1000)) {
            btn.Invoke()
        }
        else if (btn := WaitForButton(uia, unreadPattern, 1000)) {
            btn.Invoke()
        }
        else {
            MsgBox "Could not find a 'Mark as read' or 'Mark as unread' button."
        }
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

