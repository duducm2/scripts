/********************************************************************
 *   Win+Alt+Shift symbol layer shortcuts (AHK v2)
 *   • Provides system-wide symbol shortcuts
 ********************************************************************/

#Requires AutoHotkey v2.0+

#SingleInstance Force

SetTitleMatchMode 2

#include %A_ScriptDir%\env.ahk
#include UIA-v2\Lib\UIA.ahk
#include UIA-v2\Lib\UIA_Browser.ahk

; Function to send symbol characters
SendSymbol(sym) {
    SendText(sym)
}

;-------------------------------------------------------------------
; Cheat-sheet overlay (Win + Alt + Shift + A) – shows remapped shortcuts
;-------------------------------------------------------------------

; Map that stores the pop-up text for each application.  Extend freely.
cheatSheets := Map()

;---------------------------------------- Shift + keys ----------------------------------------------
; ----- You can have repeated keys, depending on the software.
; ----- Prefered Keys sequences (most important first): Y U I O P H J K L N M , . 6 7 8 9 0 W E R T D F G C V B

; --- WhatsApp desktop -------------------------------------------------------
cheatSheets["WhatsApp"] := "
(
Shift+Y  →  Toggle voice message
Shift+U  →  Search chats
Shift+I  →  Reply
Shift+O  →  Sticker panel
Shift+P  →  Toggle Unread filter
Shift+H  →  Focus current chat
)"  ; end WhatsApp

; --- Outlook main window ----------------------------------------------------
cheatSheets["OUTLOOK.EXE"] := "
(
Shift+Y  →  Send to General
Shift+U  →  Send to Newsletter
Shift+H  →  Subject / Title
Shift+J  →  Required / To
Shift+K  →  Date Picker
Shift+L  →  Subject → Body
Shift+N  →  Focused / Other
Shift+M  →  Make Recurring → Tab
)"  ; end Outlook

; --- Microsoft Teams – meeting window --------------------------------------
cheatSheets["TeamsMeeting"] := "
(
Shift+Y  →  Open Chat pane
Shift+U  →  Maximize meeting window
)"  ; end TeamsMeeting

; --- Microsoft Teams – chat window -----------------------------------------
cheatSheets["TeamsChat"] := "
(
Shift+K  →  Pin chat
Shift+J  →  Mark unread
Shift+Y  →  Like
Shift+U  →  Heart
Shift+I  →  Laugh
Shift+L  →  Remove pin
Shift+R  →  Reply
Shift+E  →  Edit message
Shift+O  →  Home panel
)"  ; end TeamsChat

; --- Spotify ---------------------------------------------------------------
cheatSheets["Spotify.exe"] := "
(
Shift+Y  →  Toggle Connect panel
Shift+U  →  Toggle Full screen
)"  ; end Spotify

; --- OneNote ---------------------------------------------------------------
cheatSheets["ONENOTE.EXE"] := "
(
Shift+Y  →  Select line and children
Shift+U  →  Expand
Shift+I  →  Collapse
Shift+J  →  Expand all
Shift+K  →  Collapse all
Shift+D  →  Delete line and children
)"  ; end OneNote

; --- Chrome general shortcuts ----------------------------------------------
cheatSheets["chrome.exe"] := "
(
Shift+G  →  Pop current tab to new window
)"  ; end Chrome

; --- Cursor / VS Code ------------------------------------------------------
cheatSheets["Cursor.exe"] := "
(
Shift+Y  →  Unfold
Shift+U  →  Fold
Shift+I  →  Unfold all
Shift+O  →  Fold all
Shift+P  →  Go to terminal
Shift+H  →  New terminal
Shift+J  →  Go to file explorer
Shift+K  →  Format code
Shift+L  →  Command palette
Shift+M  →  Change project
Shift+,  →  Show chat history
Shift+.  →  Extensions
Shift+6  →  Switch brackets
Shift+7  →  Search
Shift+8  →  Save all documents
)"  ; end Cursor

cheatSheets["Code.exe"] := "
(
Shift+Y  →  Unfold
Shift+U  →  Fold
Shift+I  →  Unfold all
Shift+O  →  Fold all
Shift+P  →  Go to terminal
Shift+H  →  New terminal
Shift+J  →  Go to file explorer
Shift+K  →  Format code
Shift+L  →  Command palette
Shift+M  →  Change project
Shift+,  →  Show chat history
Shift+.  →  Extensions
Shift+6  →  Switch brackets
Shift+7  →  Search
Shift+8  →  Save all documents
)"  ; end VS Code

; --- Windows Explorer ------------------------------------------------------
cheatSheets["explorer.exe"] := "
(
Shift+Y  →  Select first file
Shift+U  →  New folder
Shift+I  →  New shortcut
)"  ; end Explorer

; --- ClipAngel -------------------------------------------------------------
cheatSheets["ClipAngel.exe"] := "
(
Shift+Y  →  Select filtered content and copy
)"  ; end ClipAngel

; --- Figma -----------------------------------------------------------------
cheatSheets["Figma.exe"] := "
(
Shift+Y  →  Show/Hide UI
Shift+U  →  Component search
Shift+I  →  Select parent
Shift+O  →  Create component
Shift+P  →  Detach instance
Shift+H  →  Add auto layout
Shift+J  →  Remove auto layout
Shift+K  →  Suggest auto layout
Shift+L  →  Export
Shift+N  →  Copy as PNG
Shift+M  →  Actions...
Shift+,  →  Align left
Shift+.  →  Align right
Shift+6  →  Align top
Shift+7  →  Align bottom
Shift+8  →  Align center horizontal
Shift+9  →  Align center vertical
Shift+0  →  Distribute horizontal spacing
Shift+W  →  Distribute vertical spacing
Shift+E  →  Tidy up
)"  ; end Figma

; ========== Helper to decide which sheet applies ===========================
GetCheatSheetText() {
    global cheatSheets

    exe := WinGetProcessName("A") ; active process name (e.g. chrome.exe)
    title := WinGetTitle("A")

    ; Special handling for Chrome-based apps that share chrome.exe
    if (exe = "chrome.exe") {
        chromeShortcuts := cheatSheets.Has("chrome.exe") ? cheatSheets["chrome.exe"] : ""
        appShortcuts := ""

        if InStr(title, "WhatsApp")
            appShortcuts := cheatSheets.Has("WhatsApp") ? cheatSheets["WhatsApp"] : ""
        if InStr(title, "Gmail")
            appShortcuts :=
                "(Gmail)`r`nShift+Y → Inbox`r`nShift+U → Updates`r`nShift+I → Mark read`r`nShift+O → Mark unread"
        if InStr(title, "ChatGPT") || InStr(title, "chatgpt")
            appShortcuts :=
                "(ChatGPT)`r`nShift+Y → Cut all`r`nShift+U → Model selector`r`nShift+I → Toggle sidebar`r`nShift+O → Type " "chatgpt" "`r`nShift+P → New chat`r`nShift+H → Copy code block"

        ; Combine Chrome general + app-specific shortcuts
        if (appShortcuts != "" && chromeShortcuts != "")
            return "(Chrome)`r`n" chromeShortcuts "`r`n`r`n" appShortcuts
        else if (appShortcuts != "")
            return appShortcuts
        else if (chromeShortcuts != "")
            return "(Chrome)`r`n" chromeShortcuts
        else
            return ""
    }

    ; Microsoft Teams – differentiate meeting vs chat via helper predicates
    if IsTeamsMeetingActive()
        return cheatSheets.Has("TeamsMeeting") ? cheatSheets["TeamsMeeting"] : ""
    if IsTeamsChatActive()
        return cheatSheets.Has("TeamsChat") ? cheatSheets["TeamsChat"] : ""

    ; Direct match by process name
    if cheatSheets.Has(exe)
        return cheatSheets[exe]

    ; Try case-insensitive match for process name
    for key, value in cheatSheets {
        if (StrLower(key) = StrLower(exe))
            return value
    }

    ; Nothing found → blank → caller will show fallback message
    return ""
}

; ========== GUI creation & showing ========================================
ToggleShortcutHelp() {
    static helpGui := 0

    static shown := false

    ; Toggle off if currently shown
    if (IsObject(helpGui) && shown) {
        helpGui.Hide()
        shown := false
        return
    }

    ; Ensure text for current context
    text := GetCheatSheetText()
    if (text = "") {
        exe := WinGetProcessName("A")
        text := "No cheat-sheet registered for:`n" exe
    }

    static cheatCtrl

    if !IsObject(helpGui) {
        helpGui := Gui(
            "+AlwaysOnTop -Caption +ToolWindow +Border +Owner +LastFound"
        )
        helpGui.BackColor := "000000"
        helpGui.SetFont("s12 cFFFF00", "Consolas")
        cheatCtrl := helpGui.Add("Edit", "ReadOnly +Multi -E0x200 -VScroll -HScroll -Border Background000000 w600 r1")

        ; Esc also hides
        Hotkey "Esc", (*) => (helpGui.Hide(), shown := false), "Off"
    }

    ; Update cheat-sheet text and resize height to fit
    cheatCtrl.Value := text
    lineCnt := StrLen(text) ? StrSplit(text, "`n").Length : 1

    ; Calculate height based on line count (font size 12 ≈ 20px per line + margins)
    controlHeight := lineCnt * 20 + 10
    guiHeight := controlHeight + 20  ; Add padding for GUI borders

    ; Resize the control and GUI explicitly
    cheatCtrl.Move(, , 600, controlHeight)
    helpGui.Show("NoActivate Center w620 h" guiHeight)
    shown := true
    Hotkey "Esc", "On"
}

; Hotkey: Win + Alt + Shift + A  (#!+a) — toggles the cheat-sheet
#!+a:: ToggleShortcutHelp()

;-------------------------------------------------------------------
; Environment paths (unchanged)
;-------------------------------------------------------------------
global WORK_SCRIPTS_PATH := "C:\Users\fie7ca\Documents\01 - Scripts"
global PERSONAL_SCRIPTS_PATH := "G:\Meu Drive\12 - Scripts"
; global IS_WORK_ENVIRONMENT   := true    ; set to false on personal rig // This will now be loaded from env.ahk

; ---------------------------------------------------------------------------
; ShowErr(msgOrErr)  – uniform MsgBox for any thrown value
; ---------------------------------------------------------------------------
ShowErr(err) {
    text := (Type(err) = "Error") ? err.Message : err
    MsgBox("Error:`n" text, "Error", "IconX")
}

;-------------------------------------------------------------------
; OneNote Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe onenote.exe")

; Shift + y : Onenote: select line and children
+y:: Send("^+-") ; Remaps to Ctrl + Shift + -~

; Shift + U : Onenote: select line and children
+d:: {
    Send("^+-") ; Remaps to Ctrl + Shift + -
    Send "{Del}"
}

; Shift + U : Onenote: expand all
+u:: Send("!+-")     ; Remaps to Alt + Shift + 0

; Shift + I : Onenote: collapse all
+i:: Send("!+{+}")     ; Remaps to Alt + Shift + 1

; Shift + J : Onenote: expand all
+j:: Send("!+1")     ; Remaps to Alt + Shift + 0

; Shift + K : Onenote: collapse all
+k:: Send("!+0")     ; Remaps to Alt + Shift + 1

#HotIf

;-------------------------------------------------------------------
; ClipAngel Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ClipAngel")

; Shift + Y : Select filtered content and copy
+y:: {
    Send "{F10}"
    Sleep "300"
    Send "{F10}"
}

#HotIf

;-------------------------------------------------------------------
; WhatsApp Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("WhatsApp")

; Remap Shift+J to Ctrl+Alt+Shift+U
+j:: Send "^!+u"

; Remap Shift+K to Ctrl+Alt+Shift+P
+k:: Send "^!+p"

; Shift + U: Extended search (Alt+K)
+u:: Send("!k")

; Shift + I: Reply (Alt+R)
+i:: Send("!r")

; Shift + O: Sticker panel (Ctrl+Alt+S)
+o:: Send("^!s")

; Shift + P: Edit last message (Win+Up)
; The user updated this to be Toggle Unread
+p::
{
    try
    {
        uia := UIA_Browser()
        Sleep 300 ; Give UIA time to attach

        ; Find the "Unread" and "All" filter buttons
        unreadButton := uia.FindElement({ Name: "Unread", AutomationId: "unread-filter", Type: "TabItem" })
        allButton := uia.FindElement({ Name: "All", AutomationId: "all-filter", Type: "TabItem" })

        if (unreadButton && allButton) {
            ; Check if the "Unread" button is currently selected.
            ; The .IsSelected property is part of the SelectionItemPattern.
            if (unreadButton.IsSelected) {
                allButton.Click() ; If Unread is selected, click All
            }
            else {
                unreadButton.Click() ; Otherwise, click Unread
            }
        }
        else if (unreadButton) {
            ; Fallback if only the Unread button is found
            unreadButton.Click()
        }
        else {
            MsgBox "Could not find the 'Unread' filter button."
        }
    }
    catch Error as e {
        MsgBox "An error occurred: " e.Message
    }
}

global isRecording := false          ; persists between hotkey presses

; ---------- HOTKEY ----------------------------------------------------------
+y:: ToggleVoiceMessage()             ; Shift + Y

; ---------------------------------------------------------------------------
ToggleVoiceMessage() {
    global isRecording

    try {
        chrome := UIA_Browser()      ; top-level Chrome UIA element
        if !IsObject(chrome) {
            MsgBox "Can't attach to Chrome."
            return
        }

        Sleep 400                    ; let Chrome finish drawing

        ; Exact-name regexes (case-insensitive, anchored ^ $)
        voicePattern := "i)^(Voice message|Record voice message)$"
        sendPattern := "i)^(Send|Stop recording)$"

        ; Helper to grab a button by pattern
        FindBtn(p) => WaitForButton(chrome, p)

        if (isRecording) {           ; ► we’re supposed to stop & send
            if (btn := FindBtn(sendPattern)) {
                btn.Invoke()
                isRecording := false
            } else {
                ; Assume you clicked Send manually → reset & start new rec
                isRecording := false
                if (btn := FindBtn(voicePattern)) {
                    btn.Invoke()
                    isRecording := true
                } else
                    MsgBox "Couldn't restart recording (Voice-message button missing)."
            }
        } else {                     ; ► start recording
            if (btn := FindBtn(voicePattern)) {
                btn.Invoke()
                isRecording := true
            } else
                MsgBox "Couldn't find the Voice-message button."
        }
    } catch Error as err {
        MsgBox "Error:`n" err.Message
    }
}

; ---------------------------------------------------------------------------
; WaitForButton(root, pattern, timeout := 5000)
;   • Searches all descendant buttons of `root` until Name matches `pattern`
;   • Returns the UIA element or 0 if none matched within `timeout` ms
; ---------------------------------------------------------------------------
WaitForButton(root, pattern, timeout := 5000) {
    if !IsObject(root)
        return 0
    deadline := A_TickCount + timeout
    while (A_TickCount < deadline) {
        for btn in root.FindAll({ Type: "Button" }) {
            if RegExMatch(btn.Name, pattern)
                return btn
        }
        Sleep 150
    }
    return 0
}

; ---------------------------------------------------------------------------
; WaitForList(root, pattern := "", timeout := 5000)
;   • Searches descendant List controls; Name must match `pattern` if provided
;   • Returns the UIA element or 0 after `timeout` ms
; ---------------------------------------------------------------------------
WaitForList(root, pattern := "", timeout := 5000) {
    if !IsObject(root)
        return 0
    deadline := A_TickCount + timeout
    while (A_TickCount < deadline) {
        for lst in root.FindAll({ Type: "List" }) {
            if (!pattern || RegExMatch(lst.Name, pattern))
                return lst
        }
        Sleep 150
    }
    return 0
}

; Shift + h: Focus the current conversation
+h::
{
    try
    {
        ; WhatsApp desktop is Chromium-based, so we can use UIA_Browser.
        ; It should attach to the active window, which is WhatsApp thanks to #HotIf.
        uia := UIA_Browser()
        Sleep 300 ; Give UIA time to attach to the browser. A similar delay is in the reference script.

        ; Find the "Archived" button to use as an anchor.
        ; The user provided: Name:"Archived "
        archivedButton := uia.FindElement({ Name: "Archived ", Type: "Button" })

        if (archivedButton) {
            ; Focus the button without clicking it.
            archivedButton.SetFocus()
            ; Send Tab to move to the main conversation list.
            ; From there, the focus should be on the selected chat.
            SendInput "{Tab}"
        }
        else {
            MsgBox "Could not find the 'Archived' button."
        }
    }
    catch Error as e {
        MsgBox "An error occurred while trying to focus WhatsApp conversation: " e.Message
    }
}

#HotIf

;-------------------------------------------------------------------
; Outlook Reminder Window Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe OUTLOOK.EXE") && RegExMatch(WinGetTitle("A"), "i)\d+\s+Reminder\(s\)")

; Shift + Y : Select first reminder list item
+Y::
{
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Find the first list item in the reminder window
        ; Looking for ListItem type (50007) with AutomationId pattern "ListViewItem-0"
        firstItem := root.FindFirst({ AutomationId: "ListViewItem-0", ControlType: "ListItem" })

        ; Fallback: if specific AutomationId doesn't work, find first ListItem
        if !firstItem {
            firstItem := root.FindFirst({ ControlType: "ListItem" })
        }

        ; Fallback: try finding by type number if ControlType doesn't work
        if !firstItem {
            firstItem := root.FindFirst({ Type: "50007" })
        }

        if firstItem {
            firstItem.Select()
            ; Alternative method: click the item to ensure selection
            ; firstItem.Click()
        }
        else {
            MsgBox("Could not find the first reminder item.", "Reminder Selection", "IconX")
        }
    }
    catch Error as e {
        MsgBox("UIA error:`n" e.Message, "Outlook Reminder Error", "IconX")
    }
}

; ativa a janela de lembretes do Outlook
ActivateReminder() {
    WinActivate("ahk_exe OUTLOOK.EXE")
    WinWaitActive("ahk_exe OUTLOOK.EXE", , 1)
}

; digita o tempo e aperta Alt+S
QuickSnooze(t) {
    ActivateReminder()
    Send("{Tab}")              ; chega ao combo
    Send("{Tab}")              ; chega ao combo
    Sleep 100
    Send("^a{Delete}" . t)     ; substitui o texto
    Sleep 120
    Send("!s")                 ; Alt+S = Snooze
    Sleep 200
    Send("{Tab}")
    Send("{Tab}")
    Send("{Tab}")
}

; caixa de confirmação antes de executar
Confirm(t) {
    if MsgBox("Snooze for " t "?", "Confirm Snooze", "YesNo Icon?") = "Yes"
        QuickSnooze(t)
}

; HOTKEYS de teste
; Shift+U  → 1 minuto
; Shift+I  → 2 minutos
; Shift+O  → 3 minutos
+U:: Confirm("1 hour")
+I:: Confirm("4 hours")
+O:: Confirm("1 day")
+P:: ConfirmDismissAll()  ; Shift + P : Dismiss all reminders with confirmation

#HotIf

;-------------------------------------------------------------------
; Microsoft Teams Helper functions
;-------------------------------------------------------------------
IsTeamsMeetingTitle(title) {
    if InStr(title, "Chat |") || InStr(title, "Sharing control bar |")
        return false
    if InStr(title, "Microsoft Teams meeting")
        return true
    return RegExMatch(title, "i)^.*\| Microsoft Teams.*$")
}

IsTeamsChatTitle(title) {
    if InStr(title, "Sharing control bar |") || InStr(title, "Microsoft Teams meeting")
        return false
    return InStr(title, "Chat |") && RegExMatch(title, "i)\| Microsoft Teams$")
}

; -------------------------------------------------------------------
; Helper predicates to detect which Teams window is active
; -------------------------------------------------------------------
IsTeamsMeetingActive() {
    return IsTeamsMeetingTitle(WinGetTitle("A"))
}
IsTeamsChatActive() {
    return IsTeamsChatTitle(WinGetTitle("A"))
}

; -------------------------------------------------------------------
; Microsoft Teams Shortcuts – MEETING WINDOW
; -------------------------------------------------------------------
#HotIf IsTeamsMeetingActive()

; Shift + Y : Open Chat pane
+Y:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        btn := root.FindFirst({ AutomationId: "chat-button" })
        if !btn
            btn := root.FindFirst({ Name: "Chat", ControlType: "Button" })

        if btn
            btn.Click()
        else
            MsgBox("Couldn't find the Chat button.", "Control not found", "IconX")
    }
    catch as e {
        MsgBox("UIA error:`n" e.Message, "Error", "IconX")
    }
}

; Shift + U : Maximize meeting window
+U:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        btn := root.FindFirst({ Name: "Maximize meeting window", ControlType: "Button" })
        if btn
            btn.Click()
        else
            MsgBox("Couldn't find the Maximize button.", "Control not found", "IconX")
    }
    catch as e {
        MsgBox("UIA error:`n" e.Message, "Error", "IconX")
    }
}

#HotIf

;-------------------------------------------------------------------
; Microsoft Teams Shortcuts (chat)
;-------------------------------------------------------------------
#HotIf IsTeamsChatActive()

; Shift + K : Pin
+K::
{
    Sleep "500"
    Send "^1"
    Sleep "300"
    Send("{AppsKey}")
    Sleep "300"
    Send "{Down}"
    Send "{Down}"
    Send "{Right}"
    Send "{Enter}"
}

; Shift + Y : Like
+y::
{
    Send "{Enter}"
    Send "{Enter}"
    Send "{Esc}"
}

; Shift + U : Heart
+u::
{
    Send "{Enter}"
    Send "{Down}"
    Send "{Enter}"
    Send "{Esc}"
}

; Shift + I : Laugh
+i::
{
    Send "{Enter}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
    Send "{Esc}"
}

; Shift + Ç : Remove pin
+l::
{
    Sleep "500"
    Send "^1"
    Sleep "300"
    Send("{AppsKey}")
    Sleep "300"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Shift + R : Reply Message
+r::
{
    Send "{Enter}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Shift + E : Edit message
+e::
{
    Send "{Enter}"
    Sleep 100
    Send "{Left}"
    Sleep 100
    Send "{Left}"
    Sleep 100
    Send "{Enter}"
}

; Shift + J : Mark as unread
+j::
{
    Send "^1"
    Sleep "400"
    Send("{AppsKey}")
    Sleep "400"
    Send "{Down}"
    Send "{Enter}"
}

; Shift + O : Ctrl+1 then Ctrl+Shift+Home
+o::
{
    Send "^1"
    Sleep "200"          ; 200 ms
    Send "^+{Home}"
}

#HotIf

;-------------------------------------------------------------------
; Outlook Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe OUTLOOK.EXE")

; Shift + U : Send to general
+Y::
{
    Send "!5"
    Send "O"
    Send "{Home}"
    Send "00"
    Send "{Enter}"
}

; Shift + I : Send to newsletter
+U::
{
    Send "!5"
    Send "O"
    Send "{Home}"
    Send "01"
    Send "{Enter}"
}

; Shift + H: Copy last code block
+I:: Send("?")

; Shift + L  →  focus Subject (e-mail) or Location (appointment) and Tab to body
+L::
{
    hwnd := WinActive("ahk_class rctrl_renwnd32 ahk_exe OUTLOOK.EXE")
    if !hwnd
        return                                    ; no inspector on top

    ui := UIA.ElementFromHandle(hwnd)

    ; helper that returns the element or blank (instead of throwing)
    safeFind(condObj) {
        try return ui.FindFirst(condObj)
        catch
            return ""
    }

    ; 1) e-mail Subject field  (AutomationId 4101)
    target := safeFind({ AutomationId: "4101" })

    ; 2) appointment Location field  (AutomationId 4111)
    if (!target)
        target := safeFind({ AutomationId: "4111" })

    ; 3) fallback - any writable edit nearest the top (first Edit under Ribbon)
    if (!target)
        target := safeFind({ ControlType: "Edit", IsReadOnly: 0 })

    if (target) {
        target.SetFocus()
        Sleep 50           ; small pause for reliability
        Send "{Tab}"       ; jump into the body
    } else {
        Send "{F6}{F6}"    ; last-resort Outlook cycle
    }
}

+N:: {                                  ; toggle Focused / Other
    static nextOutlookButton := "Other"

    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        btn := root.FindFirst({
            Name: nextOutlookButton, Type: "Button"
        })

        if btn {
            btn.Click()
            nextOutlookButton := (nextOutlookButton = "Other")
                ? "Focused" : "Other"
        } else {
            MsgBox("Couldn't find ‘" nextOutlookButton "’.",
                "Button not found", "IconX")
        }

    } catch Error as err {              ; ← **only this form**
        ShowErr(err)
    }
}

; -------------------------------------------------------------------
; Focus helpers – reuse for any field you need
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
; New shortcuts
;   Shift + J  → Title field
;   Shift + K  → Required (To…) field
;   Shift + L  → Date Picker button
; -------------------------------------------------------------------

+H:: {                                    ; go to Title or Subject
    if FocusOutlookField({ AutomationId: "4100" })
        return
    if FocusOutlookField({ Name: "Title", ControlType: "Edit" })
        return
    ; fallback to Subject
    if FocusOutlookField({ AutomationId: "4101" })
        return
    if FocusOutlookField({ Name: "Subject", ControlType: "Edit" })
        return
    MsgBox "Couldn't find the Title or Subject field.", "Control not found", "IconX"
}

+J:: {                                    ; go to Required or To
    if FocusOutlookField({ AutomationId: "4109" })
        return
    if FocusOutlookField({ Name: "Required", ControlType: "Edit" })
        return
    ; fallback to To
    if FocusOutlookField({ AutomationId: "4117" })
        return
    if FocusOutlookField({ Name: "To", ControlType: "Edit" })
        return
    MsgBox "Couldn't find the Required or To field.", "Control not found", "IconX"
}

+K:: {                                    ; go to Date Picker
    if FocusOutlookField({ AutomationId: "4352" })
        return
    if FocusOutlookField({ Name: "Date Picker", ControlType: "Button" })
        return
    MsgBox "Couldn't find the Date Picker.", "Control not found", "IconX"
}

+M:: {  ; Shift + M → focus the “Make Recurring” button, then press Tab once
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        btn := root.FindFirst({ AutomationId: "4364", ControlType: "Button" })
        if !btn
            btn := root.FindFirst({ Name: "Make Recurring", ControlType: "Button" })

        if btn {
            btn.SetFocus()
            Sleep 100
            Send "{Tab}"
        } else {
            MsgBox "Couldn't find the Make Recurring button.", "Control not found", "IconX"
        }
    } catch Error as err {
        ShowErr(err)
    }
}

#HotIf

;-------------------------------------------------------------------
; Google Chrome Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe chrome.exe")

; Shift + G : Pop up the current tab
+g::
{
    Send "{F6}"                        ; Focus address bar (omnibox)
    Sleep 100
    Send "{F6}"                        ; Focus the tab strip (current tab)
    Sleep 100
    Send "{AppsKey}"                   ; Open the tab's context menu (AppsKey or Shift+F10)
    Sleep 100                          ; Wait a moment for menu to open
    Send "m"                           ; Select "Move tab to new window" (press 'm')
    Sleep 100
    Send "{Enter}"                     ; Confirm the action (detach tab)
    Sleep 100
    Send "{Enter}"                     ; Confirm the action (detach tab)
}

#HotIf

;-------------------------------------------------------------------
; ChatGPT Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("chatgpt")

; Shift + Y : Select all and cut
+y::
{
    Send "^a"
    Sleep 50
    Send "^x"
}

; Shift + U → click ChatGPT's model selector (any language)
+u:: {
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 300                 ; brief settle-time

        ; 1️⃣  The button/menu item always carries an AutomationId that starts with "radix-".
        modelCtl := uia.FindElement({ AutomationId: "radix-", matchmode: "StartsWith" })

        ; 2️⃣  If, for some reason, that fails, fall back to the label text (no type filter)
        if (!modelCtl) {
            for name in ["Model selector", "Seletor de modelo"]
                if (modelCtl := uia.FindElement({ Name: name, matchmode: "Substring" }))
                    break
        }

        ; 3️⃣  Click or complain
        if (modelCtl)
            modelCtl.Click()
        else
            MsgBox "Couldn't find the model-selector (ID or label)."
    }
    catch Error as e {
        MsgBox "UIA error: " e.Message
    }
}

; Shift + I: Toggle sidebar
+i:: Send("^+s")

; Shift + O: Write chatgpt
+o:: Send("chatgpt")

; Shift + P: New chat
+p:: Send("^+o")

; Shift + H: Copy last code block
+h:: Send("^+;")

#HotIf

;-------------------------------------------------------------------
; Windows Explorer Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe explorer.exe")

; Shift + Y : Select first item in list
+y::
{
    Send "{ESC}"
    try
    {
        explorerEl := UIA.ElementFromHandle(WinExist("A"))
        firstItem := explorerEl.FindFirst({ Type: "ListItem" }) ; grab first ListItem regardless of AutomationId
        firstItem.Select()
        firstItem.SetFocus()
    }
    catch Error {
        ; Fallback: just press Home to select the first item
        Send "{Home}"
    }
}

; Shift + U : New folder
+u:: Send("^+n")

; Shift + I : New Shortcut
+i::
{
    try
    {
        explorerEl := UIA.ElementFromHandle(WinExist("A"))
        newButton := explorerEl.FindFirst({ Name: "Novo", Type: "Button" })
        newButton.Click()
        Sleep 150
        Send "{Down}"
        Send "{Enter}"
    }
    catch Error {
        ; Fallback: open context menu and choose Shortcut
        Send "+{F10}"
        Sleep 150
        Send "w"  ; New submenu (assumes English—adjust if needed)
        Sleep 100
        Send "s"  ; Shortcut option (assumes English—adjust if needed)
    }
}

#HotIf

;-------------------------------------------------------------------
; Gmail Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("Gmail")

; Shift + Y: Go to main inbox
+y:: Send("gi")

; Shift + U: Go to updates
+u::
{
    try
    {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 300 ; Give UIA time to attach

        ; Find the "Updates" tab. The name can change (e.g., "Updates, 1 new message,"),
        ; so we search for a TabItem element that starts with "Updates".
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

; Shift + I: Mark as read
+i:: Send("+i")

; Shift + O: Mark as unread
+o:: Send("+u")

#HotIf

;-----------------------------------------
;  Detect which editor is active
;-----------------------------------------
IsEditorActive() {
    global IS_WORK_ENVIRONMENT
    return IS_WORK_ENVIRONMENT
        ? WinActive("ahk_exe Code.exe")       ; Visual Studio Code
            : WinActive("ahk_exe Cursor.exe")     ; Cursor (VS Code-based)
}

;-------------------------------------------------------------------
; Cursor / VS Code Shortcuts
;-------------------------------------------------------------------
#HotIf IsEditorActive()

; Shift + Y : Unfold
+y::
{
    Send "^+8"
}

; Shift + U : Fold
+u::
{
    Send "^+9"
}

; Shift + I : Unfold all
+i::
{
    Send "^m"
    Send "^j"
}

; Shift + O : Fold all
+o::
{
    Send "^m"
    Send "^0"
}

; Shift + P : Go to terminal
+p:: Send "^'"

; Shift + H : New terminal
+h:: Send '^+"'

; Shift + J : Go to file explorer
+j:: Send "^+e"

; Shift + K : Format code
+k:: Send "!+f"

; Shift + L : command palette
+l:: Send "^+p"

; Shift + M : Change project
+m:: Send "^r"

; Shift + , : Show chat history
+,::
{
    Send "^+p" ; Open command palette
    Sleep 200
    Send "show history"
    Sleep 200
    Send "{Enter}"
}

; Shift + . : Extensions
+.:: Send "^+x"

; Shift + 6 : Switch the brackets open/close
+6:: Send "^+]"

; Shift + 7 : Search
+7:: Send "^+f"

; Shift + 8 : Save all documents
+8::
{
    Send "^k"
    Send "s"
}

#HotIf

;-------------------------------------------------------------------
; Spotify Shortcuts
;-------------------------------------------------------------------ww
#HotIf WinActive("ahk_exe Spotify.exe")

; Shift + Y : Toggle Connect panel and select device
+y::
{
    try {
        spot := UIA_Browser("ahk_exe Spotify.exe")
        Sleep 300

        ; Check if panel is open by looking for "This computer" button
        thisComputerPattern := "i)^This computer$"
        thisComputerBtn := WaitForButton(spot, thisComputerPattern, 1000)

        if (thisComputerBtn) {
            ; Panel is open, close it by clicking Connect button
            connectPattern := "i)^(Connect to a device|Conectar a um dispositivo|Connect)$"
            if (connectBtn := WaitForButton(spot, connectPattern))
                connectBtn.Invoke()
            else
                MsgBox "Couldn't find the Connect button to close panel."
        } else {
            ; Panel is closed, open it and select "This computer"
            connectPattern := "i)^(Connect to a device|Conectar a um dispositivo|Connect)$"
            if (connectBtn := WaitForButton(spot, connectPattern)) {
                connectBtn.Invoke()
                Sleep 500  ; Wait for panel to open

                ; Now select "This computer"
                if (thisComputerBtn := WaitForButton(spot, thisComputerPattern))
                    thisComputerBtn.SetFocus()
                else
                    MsgBox "Panel opened but couldn't find 'This computer' button."
            } else {
                MsgBox "Couldn't find the Connect-to-device button."
            }
        }
    } catch Error as e {
        MsgBox "Error: " e.Message
    }
}

; Shift + U : Toggle full screen
+u::
{
    retryCount := 0
    maxRetries := 1

    while (retryCount <= maxRetries) {
        try {
            spot := UIA_Browser("ahk_exe Spotify.exe")
            Sleep 300

            ; Check if any RadioButton elements are visible (indicates full screen mode)
            inFullScreen := false
            try {
                radioButtons := spot.FindAll({ Type: "RadioButton" })
                if (radioButtons.Length > 0) {
                    inFullScreen := true
                }
            } catch {
                ; No radio buttons found, not in full screen
            }

            if (inFullScreen) {
                ; We're in full screen, exit with ESC twice
                Send "{Esc}"
                Sleep 100
                Send "{Esc}"
            } else {
                ; Not in full screen, look for Enter Full screen button
                enterFsPattern := "i)^Enter Full screen$"
                if (btn := WaitForButton(spot, enterFsPattern, 1000)) {
                    btn.Invoke()
                } else {
                    MsgBox "Couldn't find the Enter Full screen button."
                }
            }
            return ; Success, exit the retry loop
        } catch Error as e {
            retryCount++
            if (retryCount > maxRetries) {
                MsgBox "Error after retry: " e.Message
            } else {
                Sleep 500 ; Wait before retry
            }
        }
    }
}

#HotIf

;-------------------------------------------------------------------
; Figma Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe Figma.exe")

; Shift + Y : Show/Hide UI (Ctrl + \)
+y:: Send("^]")

; Shift + U : Component search (Shift + I)
+u:: Send("+i")

; Shift + I : Select parent (\)
+i:: Send("]")

; Shift + O : Create component (Ctrl + Alt + K)
+o:: Send("^!k")

; Shift + P : Detach instance (Ctrl + Alt + B)
+p:: Send("^!b")

; Shift + H : Add auto layout (Shift + A)
+h:: Send("+a")

; Shift + J : Remove auto layout (Alt + Shift + A)
+j:: Send("!+a")

; Shift + K : Suggest auto layout (Ctrl + Alt + Shift + A)
+k:: Send("^!+a")

; Shift + L : Export (Ctrl + Shift + E)
+l:: Send("^+e")

; Shift + N : Copy as PNG (Ctrl + Shift + C)
+n:: Send("^+c")

; Shift + M : Actions... (Ctrl + K)
+m:: Send("^k")

; Shift + , : Align left (Alt + A)
+,:: Send("!a")

; Shift + . : Align right (Alt + D)
+.:: Send("!d")

; Shift + 6 : Align top (Alt + W)
+6:: Send("!w")

; Shift + 7 : Align bottom (Alt + S)
+7:: Send("!s")

; Shift + 8 : Align center horizontal (Alt + H)
+8:: Send("!h")

; Shift + 9 : Align center vertical (Alt + V)
+9:: Send("!v")

; Shift + 0 : Distribute horizontal spacing (Alt + Shift + H)
+0:: Send("!+h")

; Shift + W : Distribute vertical spacing (Alt + Shift + V)
+w:: Send("!+v")

; Shift + E : Tidy up (Ctrl + Alt + Shift + T)
+e:: Send("^!+t")

#HotIf

;-------------------------------------------------------------------
; Mobills Finance App Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("Mobills")

; Shift + Y : Dashboard
+y:: {
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 300

        btn := uia.FindElement({ AutomationId: "menu-dashboard-item", Type: "Button" })
        if (btn) {
            btn.Click()
        } else {
            MsgBox "Could not find the Dashboard button.", "Mobills Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error navigating to Dashboard: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + U : Contas
+u:: {
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 300

        btn := uia.FindElement({ AutomationId: "menu-accounts-item", Type: "Button" })
        if (btn) {
            btn.Click()
        } else {
            MsgBox "Could not find the Contas button.", "Mobills Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error navigating to Contas: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + I : Transações
+i:: {
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 300

        btn := uia.FindElement({ AutomationId: "menu-transactions-item", Type: "Button" })
        if (btn) {
            btn.Click()
        } else {
            MsgBox "Could not find the Transações button.", "Mobills Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error navigating to Transações: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + O : Cartões de crédito
+o:: {
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 300

        btn := uia.FindElement({ AutomationId: "menu-creditCards-item", Type: "Button" })
        if (btn) {
            btn.Click()
        } else {
            MsgBox "Could not find the Cartões de crédito button.", "Mobills Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error navigating to Cartões de crédito: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + P : Planejamento
+p:: {
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 300

        btn := uia.FindElement({ AutomationId: "menu-budgets-item", Type: "Button" })
        if (btn) {
            btn.Click()
        } else {
            MsgBox "Could not find the Planejamento button.", "Mobills Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error navigating to Planejamento: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + H : Relatórios
+h:: {
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 300

        btn := uia.FindElement({ AutomationId: "menu-reports-item", Type: "Button" })
        if (btn) {
            btn.Click()
        } else {
            MsgBox "Could not find the Relatórios button.", "Mobills Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error navigating to Relatórios: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + J : Mais opções
+j:: {
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 300

        btn := uia.FindElement({ AutomationId: "menu-moreOptions-item", Type: "Button" })
        if (btn) {
            btn.Click()
        } else {
            MsgBox "Could not find the Mais opções button.", "Mobills Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error navigating to Mais opções: " e.Message, "Mobills Error", "IconX"
    }
}

#HotIf

ConfirmDismissAll() {
    if MsgBox("Dismiss all reminders?", "Confirm Dismiss", "YesNo Icon?") = "Yes"
        DismissAllReminders()
}

DismissAllReminders() {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)
        ; Try by AutomationId first
        btn := root.FindFirst({ AutomationId: "8345", ControlType: "Button" })
        ; Fallback: search by name
        if !btn
            btn := root.FindFirst({ Name: "Dismiss All", ControlType: "Button" })
        if btn {
            btn.Click()
        } else {
            MsgBox("Could not find the 'Dismiss All' button.", "Dismiss All", "IconX")
        }
    } catch Error as e {
        MsgBox("UIA error:`n" e.Message, "Dismiss All Error", "IconX")
    }
}
