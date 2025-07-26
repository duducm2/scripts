/********************************************************************
 *   Win+Alt+Shift symbol layer shortcuts (AHK v2)
 *   • Provides system-wide symbol shortcuts
 ********************************************************************/

/********************************************************************
 *   AVAILABLE WIN+ALT+SHIFT COMBINATIONS
 *   The following combinations are not currently in use:
 *   • C - Available
 *   • K - Available
 *   • 1 - Available
 *   • , - Available
 *   • . - Available
 ********************************************************************/

#Requires AutoHotkey v2.0+

#SingleInstance Force

SetTitleMatchMode 2

#include %A_ScriptDir%\env.ahk
#include UIA-v2\Lib\UIA.ahk
#include UIA-v2\Lib\UIA_Browser.ahk
#include %A_ScriptDir%\ChatGPT_Loading.ahk

; --- Global Variables ---
global smallLoadingGuis_ChatGPT := []

; --- Config ---------------------------------------------------------------
PROMPT_FILE := A_ScriptDir "\\ChatGPT_Prompt.txt"

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
; ----- Prefered Keys sequences (most important first): Y U I O P H J K L N M , . W E R T D F G C V B

; --- WhatsApp desktop -------------------------------------------------------
cheatSheets["WhatsApp"] := "
(
WhatsApp
Shift+Y  →  Toggle voice message
Shift+U  →  Search chats
Shift+I  →  Reply
Shift+O  →  Sticker panel
Shift+P  →  Toggle Unread filter
Shift+H  →  Focus current chat
Shift+J  →  Mark as read or unread
Shift+K  →  Pin chat or unpin chat
)"  ; end WhatsApp

; --- Outlook main window ----------------------------------------------------
cheatSheets["OUTLOOK.EXE"] := "
(
Outlook
Shift+Y  →  Send to General
Shift+U  →  Send to Newsletter
Shift+H  →  Subject / Title
Shift+J  →  Required / To
Shift+K  →  Date Picker
Shift+L  →  Subject → Body
Shift+N  →  Focused / Other
Shift+M  →  Make Recurring → Tab
)"  ; end Outlook

; --- Outlook Reminder window -------------------------------------------------
cheatSheets["OutlookReminder"] := "
(
Outlook – Reminders
Shift+Y  →  Select first reminder
Shift+U  →  Snooze 1 hour
Shift+I  →  Snooze 4 hours
Shift+O  →  Snooze 1 day
Shift+P  →  Dismiss all reminders
)"  ; end Outlook Reminder

; --- Outlook Appointment window ---------------------------------------------
cheatSheets["OutlookAppointment"] := "
(
Outlook – Appointment
Shift+H  →  Title field
Shift+J  →  Required / To field
Shift+K  →  Date Picker
Shift+L  →  Location → Body
Shift+M  →  Make Recurring → Tab
)"  ; end Outlook Appointment

; --- Microsoft Teams – meeting window --------------------------------------
cheatSheets["TeamsMeeting"] := "
(
Teams
Shift+Y  →  Open Chat pane
Shift+U  →  Maximize meeting window
)"  ; end TeamsMeeting

; --- Microsoft Teams – chat window -----------------------------------------
cheatSheets["TeamsChat"] := "
(
Teams
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
Spotify
Shift+Y  →  Toggle Connect panel
Shift+U  →  Toggle Full screen
Shift+I  →  Open Search
Shift+O  →  Go to Playlists
Shift+P  →  Go to Artists
Shift+H  →  Go to Albums
Shift+J  →  Go to Search
Shift+K  →  Go to Home
Shift+L  →  Go to Now Playing
Shift+N  →  Go to Made For You
Shift+M  →  Go to New Releases
Shift+,  →  Go to Charts
Shift+.  →  Toggle Now Playing View
Shift+W  →  Toggle Library Sidebar
Shift+E  →  Toggle Fullscreen Library
Shift+R  →  Toggle lyrics
Shift+T  →  Toggle play/pause
)"  ; end Spotify

; --- OneNote ---------------------------------------------------------------
cheatSheets["ONENOTE.EXE"] := "
(
OneNote
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
Chrome
Shift+G  →  Pop current tab to new window
)"  ; end Chrome

; --- Cursor / VS Code ------------------------------------------------------
cheatSheets["Cursor.exe"] := "
(
Cursor
Shift+Y  →  Fold
Shift+U  →  Unfold
Shift+I  →  Fold all
Shift+O  →  Unfold all
Shift+P  →  Go to terminal
Shift+H  →  New terminal
Shift+J  →  Go to file explorer
Shift+K  →  Format code
Shift+L  →  Command palette
Shift+M  →  Change project
Shift+,  →  Show chat history
Shift+.  →  Extensions
Shift+W  →  Switch brackets
Shift+E  →  Search
Shift+R  →  Save all documents
Shift+T  →  Change ML model
Shift+D  →  Git section
Shift+F  →  Generate commit message with AI
Shift+G  →  Commit
Shift+C  →  Commit and push
Shift+V  →  Git pull
Shift+B  →  Git push
)"  ; end Cursor

cheatSheets["Code.exe"] := "
(
VS Code
Shift+Y  →  Fold
Shift+U  →  Unfold
Shift+I  →  Fold all
Shift+O  →  Unfold all
Shift+P  →  Go to terminal
Shift+H  →  New terminal
Shift+J  →  Go to file explorer
Shift+K  →  Format code
Shift+L  →  Command palette
Shift+M  →  Change project
Shift+,  →  Show chat history
Shift+.  →  Extensions
Shift+W  →  Switch brackets
Shift+E  →  Search
Shift+R  →  Save all documents
Shift+T  →  Change ML model
Shift+D  →  Git section
Shift+F  →  Generate commit message with AI
Shift+G  →  Commit
Shift+C  →  Commit and push
Shift+V  →  Git pull
Shift+B  →  Git push
)"  ; end VS Code

; --- Windows Explorer ------------------------------------------------------
cheatSheets["explorer.exe"] := "
(
Explorer
Shift+Y  →  Select first file
Shift+U  →  Focus search bar
Shift+I  →  Focus address bar
Shift+O  →  New folder
)"  ; end Explorer

; --- ClipAngel -------------------------------------------------------------
cheatSheets["ClipAngel.exe"] := "
(
ClipAngel
Shift+Y  →  Select filtered content and copy
Shift+U  →  Switch focus list/text
Shift+I  →  Delete all non-favorite
Shift+O  →  Clear filters
Shift+P  →  Mark as favorite
Shift+H  →  Unmark as favorite
Shift+J  →  Edit text
Shift+K  →  Save as file
)"  ; end ClipAngel

; --- Figma -----------------------------------------------------------------
cheatSheets["Figma.exe"] := "
(
Figma
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
Shift+R  →  Align top
Shift+T  →  Align bottom
Shift+D  →  Align center horizontal
Shift+F  →  Align center vertical
Shift+G  →  Distribute horizontal spacing
Shift+W  →  Distribute vertical spacing
Shift+E  →  Tidy up
)"  ; end Figma

; --- Gmail ---------------------------------------------------------------
cheatSheets["Gmail"] := "
(
Gmail
Shift+Y  →  Go to main inbox
Shift+U  →  Go to updates
Shift+I  →  Go to forums
Shift+O  →  Toggle read/unread
Shift+P  →  Previous conversation
Shift+H  →  Next conversation
Shift+J  →  Archive conversation
Shift+K  →  Select conversation
Shift+L  →  Reply
Shift+N  →  Reply all
Shift+M  →  Forward
Shift+,  →  Star/unstar conversation
Shift+.  →  Delete
Shift+W  →  Report as spam
Shift+E  →  Compose new email
Shift+R  →  Search mail
Shift+T  →  Move to folder
Shift+D  →  Show keyboard shortcuts help
)"  ; end Gmail

; --- Google Keep ---------------------------------------------------------------
cheatSheets["Google Keep"] := "
(
Google Keep
Shift+Y  →  Search and select note
Shift+U  →  Toggle main menu
)"  ; end Google Keep

; --- File Dialog ---------------------------------------------------------------
cheatSheets["FileDialog"] :=
"(File Dialog)`r`nShift+Y → Select first item`r`nShift+U → Focus file name field`r`nShift+I → Click Insert/Open/Save button`r`nShift+O → Click Cancel button"

; --- UIA Tree Inspector -------------------------------------------------
cheatSheets["UIATreeInspector"] := "(UIA Tree Inspector)`r`nShift+Y → Refresh list"

; --- SettleUp Shortcuts -----------------------------------------------------
cheatSheets["Settle Up"] := "
(
Settle Up
Shift+Y  →  Add transaction
Shift+U  →  Focus expense value field
Shift+I  →  Focus expense name field
)"

; --- Miro Shortcuts -----------------------------------------------------
cheatSheets["Miro"] := "
(
Miro
Shift+Y  →  Command palette
Shift+U  →  Frame list
Shift+I  →  Group
Shift+O  →  Ungroup
Shift+P  →  Lock/Unlock
Shift+H  →  Add/Edit link
)"

; --- Wikipedia ---------------------------------------------------------------
cheatSheets["Wikipedia"] := "
(
Wikipedia
Shift+Y  →  Click search button
)"

; ========== Helper to decide which sheet applies ===========================
GetCheatSheetText() {
    global cheatSheets

    exe := WinGetProcessName("A") ; active process name (e.g. chrome.exe)
    title := WinGetTitle("A")
    hwnd := WinExist("A")

    ; Check for file dialog first (works in any app)
    if WinGetClass("ahk_id " hwnd) = "#32770" {
        txt := WinGetText("ahk_id " hwnd)
        if InStr(txt, "Namespace Tree Control") || InStr(txt, "Controle da Árvore de Namespace")
            return cheatSheets["FileDialog"]
    }

    ; Special handling for Chrome-based apps that share chrome.exe
    if (exe = "chrome.exe") {
        chromeShortcuts := cheatSheets.Has("chrome.exe") ? cheatSheets["chrome.exe"] : ""
        appShortcuts := ""

        if InStr(title, "WhatsApp")
            appShortcuts := cheatSheets.Has("WhatsApp") ? cheatSheets["WhatsApp"] : ""
        if InStr(title, "Gmail")
            appShortcuts := cheatSheets.Has("Gmail") ? cheatSheets["Gmail"] : ""
        if InStr(title, "ChatGPT") || InStr(title, "chatgpt")
            appShortcuts :=
                "(ChatGPT)`r`nShift+Y → Cut all`r`nShift+U → Model selector`r`nShift+I → Toggle sidebar`r`nShift+O → Re-send rules`r`nShift+P → New chat`r`nShift+H → Copy code block`r`nShift+L → Send and show AI banner"
        if InStr(title, "Mobills")
            appShortcuts :=
                "(Mobills)`r`nShift+Y → Dashboard`r`nShift+U → Contas`r`nShift+I → Transações`r`nShift+O → Cartões de crédito`r`nShift+P → Planejamento`r`nShift+H → Relatórios`r`nShift+J → Mais opções`r`nShift+K → Previous month`r`nShift+L → Next month`r`nShift+N → Ignore transaction`r`nShift+M → Name field"
        if InStr(title, "Google Keep") || InStr(title, "keep.google.com")
            appShortcuts := cheatSheets.Has("Google Keep") ? cheatSheets["Google Keep"] : ""
        if InStr(title, "YouTube")
            appShortcuts := "(YouTube)`r`nShift+Y → Focus search box"
        if InStr(title, "UIATreeInspector")
            appShortcuts := cheatSheets["UIATreeInspector"]
        if InStr(title, "Settle Up")
            appShortcuts := cheatSheets.Has("Settle Up") ? cheatSheets["Settle Up"] : ""
        if InStr(title, "Miro")
            appShortcuts := cheatSheets.Has("Miro") ? cheatSheets["Miro"] : ""
        if InStr(title, "Wikipedia")
            appShortcuts := cheatSheets.Has("Wikipedia") ? cheatSheets["Wikipedia"] : ""

        ; Combine Chrome general + app-specific shortcuts
        if (appShortcuts != "" && chromeShortcuts != "")
            return chromeShortcuts "`r`n`r`n" appShortcuts
        else if (appShortcuts != "")
            return appShortcuts
        else if (chromeShortcuts != "")
            return chromeShortcuts
        else
            return ""
    }

    ; UIA Tree Inspector - check both process and window title
    if (exe = "AutoHotkey64.exe" && InStr(title, "UIATreeInspector"))
        return cheatSheets["UIATreeInspector"]

    ; Microsoft Teams – differentiate meeting vs chat via helper predicates
    if IsTeamsMeetingActive()
        return cheatSheets.Has("TeamsMeeting") ? cheatSheets["TeamsMeeting"] : ""
    if IsTeamsChatActive()
        return cheatSheets.Has("TeamsChat") ? cheatSheets["TeamsChat"] : ""
    if IsFileDialogActive()
        return cheatSheets["FileDialog"]

    ; Special handling for Outlook-based apps
    if (exe = "OUTLOOK.EXE") {
        ; Detect Reminders window – e.g. "1 Reminder" / "2 Reminders" or generic "Reminders"
        if RegExMatch(title, "i)(^|\s)\d*\s*Reminder(s)?") {
            return cheatSheets.Has("OutlookReminder") ? cheatSheets["OutlookReminder"] : cheatSheets["OUTLOOK.EXE"]
        }
        ; Detect Appointment or Meeting inspector windows
        if RegExMatch(title, "i)(Appointment|Meeting)") {
            return cheatSheets.Has("OutlookAppointment") ? cheatSheets["OutlookAppointment"] : cheatSheets[
                "OUTLOOK.EXE"]
        }
        ; Fallback to generic Outlook sheet
        if cheatSheets.Has("OUTLOOK.EXE")
            return cheatSheets["OUTLOOK.EXE"]
    }

    ; Direct match by process name (generic fallback)
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

; ========== Shared variables for cheat sheet state ========================
global g_helpGui := 0
global g_helpShown := false
global g_globalGui := 0
global g_globalShown := false

; ========== GUI creation & showing ========================================
ToggleShortcutHelp() {
    global g_helpGui, g_helpShown

    ; Toggle off if currently shown
    if (IsObject(g_helpGui) && g_helpShown) {
        g_helpGui.Hide()
        g_helpShown := false
        ; Hotkey "Esc", "Off"  ; (disabled)
        return
    }

    ; Ensure text for current context
    text := GetCheatSheetText()
    if (text = "") {
        exe := WinGetProcessName("A")
        text := "No cheat-sheet registered for:`n" exe
    }

    static cheatCtrl

    if !IsObject(g_helpGui) {
        g_helpGui := Gui(
            "+AlwaysOnTop -Caption +ToolWindow +Border +Owner +LastFound"
        )
        g_helpGui.BackColor := "000000"
        g_helpGui.SetFont("s12 cFFFF00", "Consolas")
        cheatCtrl := g_helpGui.Add("Edit", "ReadOnly +Multi -E0x200 -VScroll -HScroll -Border Background000000 w600 r1"
        )

        ; Esc also hides  ; (disabled – use Win+Alt+Shift+A to hide)
        ; Hotkey "Esc", (*) => (g_helpGui.Hide(), g_helpShown := false), "Off"
    }

    ; Update cheat-sheet text and resize height to fit
    cheatCtrl.Value := text
    lineCnt := StrLen(text) ? StrSplit(text, "`n").Length : 1

    ; Calculate height based on line count (font size 12 ≈ 20px per line + margins)
    controlHeight := lineCnt * 20 + 10
    guiHeight := controlHeight + 20  ; Add padding for GUI borders

    ; Resize the control and GUI explicitly
    cheatCtrl.Move(, , 600, controlHeight)
    ; Show → measure → centre
    g_helpGui.Show("AutoSize Hide")
    CenterGuiOnActiveMonitor(g_helpGui)
    g_helpGui.Show("NoActivate")  ; ensure visible after centring
    g_helpShown := true
    ; Hotkey "Esc", "On"  ; (disabled)
}

; ========== Global shortcuts cheat sheet (Win+Alt+Shift+key) ===============
ShowGlobalShortcutsHelp() {
    global g_globalGui, g_globalShown

    ; Toggle off if currently shown
    if (IsObject(g_globalGui) && g_globalShown) {
        g_globalGui.Hide()
        g_globalShown := false
        ; Hotkey "Esc", "Off"  ; (disabled)
        return
    }

    ; Create the global shortcuts text with categories
    globalText := "
(
=== ONENOTE ===
Win+Alt+Shift+N  →  Opens or activates OneNote

=== SPOTIFY ===
Win+Alt+Shift+S  →  Opens or activates Spotify
Win+Alt+Shift+D  →  Go to library

=== CHATGPT ===
Win+Alt+Shift+8  →  Get word pronunciation, definition, and Portuguese translation
Win+Alt+Shift+9  →  Clicks on the last microphone icon
Win+Alt+Shift+0  →  Speak with ChatGPT
Win+Alt+Shift+7  →  Speak with ChatGPT (send message automatically)
Win+Alt+Shift+P  →  Copy last ChatGPT prompt
Win+Alt+Shift+J  →  Copy last ChatGPT message (in the editing box)
Win+Alt+Shift+U  →  Activate ChatGPT and copy last code block
Win+Alt+Shift+O  →  Check grammar and improve text in both English and Portuguese
Win+Alt+Shift+I  →  Opens ChatGPT
Win+Alt+Shift+L  →  Talk with ChatGPT through voice
Win+Alt+Shift+Y  →  Copy last message and read it aloud

=== YOUTUBE ===
Win+Alt+Shift+H  →  Activates Youtube

=== GOOGLE ===
Win+Alt+Shift+F  →  Opens Google

=== GMAIL ===
Win+Alt+Shift+W  →  Opens Gmail

=== OUTLOOK ===
Win+Alt+Shift+B  →  Open mail
Win+Alt+Shift+V  →  Open Reminder
Win+Alt+Shift+G  →  Open calendar

=== MICROSOFT TEAMS ===
Win+Alt+Shift+R  →  New conversation
Win+Alt+Shift+5  →  Toggle Mute (meeting)
Win+Alt+Shift+4  →  Toggle camera (meeting)
Win+Alt+Shift+T  →  Screen share (meeting)
Win+Alt+Shift+2  →  Exit meeting
Win+Alt+Shift+E  →  Select the chats window
Win+Alt+Shift+3  →  Select the meeting window

=== WHATSAPP ===
Win+Alt+Shift+Z  →  Opens WhatsApp

=== WINDOWS ===
Win+Alt+Shift+6  →  Minimizes windows
Win+Alt+Shift+M  →  Maximizes the current window

=== GENERAL ===
Win+Alt+Shift+Q  →  Jump mouse on the middle
Win+Alt+Shift+X  →  Activate hunt and Peck

=== SHORTCUTS ===
Win+Alt+Shift+A  →  Show app-specific shortcuts (quick press)
Win+Alt+Shift+A  →  Show global shortcuts (hold 400ms+)

=== WIKIPEDIA ===
Win+Alt+Shift+K  →  Opens or activates Wikipedia
)"

    static globalCtrl

    if !IsObject(g_globalGui) {
        g_globalGui := Gui(
            "+AlwaysOnTop -Caption +ToolWindow +Border +Owner +LastFound"
        )
        g_globalGui.BackColor := "000000"
        g_globalGui.SetFont("s10 cFFFF00", "Consolas")  ; Smaller font for more content
        globalCtrl := g_globalGui.Add("Edit", "ReadOnly +Multi +VScroll -HScroll -Border Background000000 w760 h540")

        ; Esc also hides  ; (disabled – use Win+Alt+Shift+A to hide)
        ; Hotkey "Esc", (*) => (g_globalGui.Hide(), g_globalShown := false), "Off"
    }

    ; Update text and show
    globalCtrl.Value := globalText
    g_globalGui.Show("AutoSize Hide")
    CenterGuiOnActiveMonitor(g_globalGui)
    g_globalGui.Show("NoActivate")
    g_globalShown := true
    ; Hotkey "Esc", "On"  ; (disabled)
}

; ========== Hotkey with hold detection ====================================
; Win + Alt + Shift + A with hold detection
#!+a::
{
    global g_helpGui, g_helpShown, g_globalGui, g_globalShown

    ; First check if any cheat sheet is currently open - if so, close it
    if (IsObject(g_helpGui) && g_helpShown) {
        g_helpGui.Hide()
        g_helpShown := false
        ; Hotkey "Esc", "Off"  ; (disabled)
        return
    }

    if (IsObject(g_globalGui) && g_globalShown) {
        g_globalGui.Hide()
        g_globalShown := false
        ; Hotkey "Esc", "Off"  ; (disabled)
        return
    }

    ; No cheat sheet is open, determine which one to show based on hold time
    static pressTime := 0
    pressTime := A_TickCount

    ; Wait for key release or timeout
    KeyWait "a", "T0.4"  ; Wait max 400ms for key release

    holdTime := A_TickCount - pressTime

    if (holdTime >= 400) {
        ; Long hold - show global shortcuts
        ShowGlobalShortcutsHelp()
    } else {
        ; Quick press - show app-specific shortcuts
        ToggleShortcutHelp()
    }
}

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

; ---------------------------------------------------------------------------
; Helper: centre a GUI over the active window (fallback: primary monitor)
; ---------------------------------------------------------------------------
CenterGuiOnActiveMonitor(guiObj) {
    ; Ensure GUI has its final size
    guiObj.GetPos(, , &guiW, &guiH)

    activeWin := WinGetID("A")
    ; Default to primary monitor work area (monitor #1)
    MonitorGetWorkArea(1, &lPrim, &tPrim, &rPrim, &bPrim)
    wx := lPrim, wy := tPrim, ww := rPrim - lPrim, wh := bPrim - tPrim

    ; If we know the active window, find which monitor contains its centre
    if (activeWin) {
        rect := Buffer(16, 0)
        if (DllCall("GetWindowRect", "ptr", activeWin, "ptr", rect)) {
            cx := NumGet(rect, 0, "int") + (NumGet(rect, 8, "int") - NumGet(rect, 0, "int")) // 2
            cy := NumGet(rect, 4, "int") + (NumGet(rect, 12, "int") - NumGet(rect, 4, "int")) // 2

            count := MonitorGetCount()
            loop count {
                idx := A_Index
                MonitorGetWorkArea(idx, &l, &t, &r, &b)
                if (cx >= l && cx <= r && cy >= t && cy <= b) {
                    wx := l, wy := t, ww := r - l, wh := b - t
                    break
                }
            }
        }
    }

    guiX := wx + (ww - guiW) / 2
    guiY := wy + (wh - guiH) / 2
    guiObj.Show("NoActivate x" Round(guiX) " y" Round(guiY))
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
    Send "{Tab}"
    Sleep 100
    Send "{Tab}"
    Sleep 300
    Send "^a"  ; Select all
    Sleep 100
    Send "^c"  ; Copy
    Sleep 100
    Send "{F10}"
}

; Shift + U : Switch focus between list and text (F10)
+u:: Send "{F10}"

; Shift + I : Delete all non-favorite (Ctrl+Alt+K)
+i:: Send "^!k"

; Shift + O : Clear filters (F7)
+o:: Send "{F7}"

; Shift + P : Mark as favorite (Alt+Q)
+p:: Send "!q"

; Shift + H : Unmark as favorite (Alt+W)
+h:: Send "!w"

; Shift + J : Edit text (F4)
+j:: Send "{F4}"

; Shift + K : Save as file (Ctrl+S)
+k:: Send "^s"

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

        if (isRecording) {           ; ► we're supposed to stop & send
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
; Wikipedia Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe chrome.exe") && InStr(WinGetTitle("A"), "Wikipedia")

; Shift + Y: Toggle the search button (Type 50005 Link, Name "Search")
+y::
{
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 300 ; Give UIA time to attach

        ; Find the search link by Type and Name
        searchLink := uia.FindElement({ Type: 50005, Name: "Search" })
        if (!searchLink) {
            ; Fallback: try by Name only, in case Type is not exposed
            searchLink := uia.FindElement({ Name: "Search" })
        }

        if (searchLink) {
            searchLink.Click()
        } else {
            MsgBox "Could not find the Wikipedia search link (Type 50005, Name 'Search')."
        }
    }
    catch Error as e {
        MsgBox "An error occurred: " e.Message
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
            MsgBox("Couldn't find '" nextOutlookButton "'.", "Button not found", "IconX")
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

+M:: {  ; Shift + M → focus the "Make Recurring" button, then press Tab once
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

        ; Find the model selector button containing "model selector" text
        modelCtl := uia.FindElement({
            Name: "Model selector",
            Type: "Button",
            matchmode: "Substring"
        })

        ; Click or complain
        if (modelCtl)
            modelCtl.Click()
        else
            MsgBox "Couldn't find the model-selector button."
    }
    catch Error as e {
        MsgBox "UIA error: " e.Message
    }
}

; Shift + I: Toggle sidebar
+i:: Send("^+s")

; Shift + O : Re-send rules & ask ChatGPT to correct mistake
+o::
{
    ; Ensure composer is focused
    Send "{Esc}"
    Sleep 150

    promptText := ""
    try promptText := FileRead(PROMPT_FILE, "UTF-8")
    if (StrLen(promptText) = 0)
        promptText := "[Prompt file missing]"

    msg :=
        "It seems you violated one of the conversation rules (e.g., incorrect name spelling). Read the rules below, identify your mistake, and reply ONLY with the corrected content." .
        "`n`n" . promptText

    oldClip := A_Clipboard
    A_Clipboard := ""
    A_Clipboard := msg
    ClipWait 1
    Send "^v"
    Sleep 100
    Send "{Enter}"
    Sleep 100
    A_Clipboard := oldClip

    ; After sending, show loading for Stop streaming
    buttonNames := ["Stop streaming", "Interromper transmissão"]
    WaitForChatGPTButtonAndShowLoading(buttonNames, "Waiting for response...")
}

; Shift + P: New chat
+p:: Send("^+o")

; Shift + H: Copy last code block
+h:: Send("^+;")

; Shift + L: Send and show AI banner
+l:: {
    ; --- Button Names (EN/PT) ---
    pt_stopStreamingName := "Interromper transmissão"
    en_stopStreamingName := "Stop streaming"
    currentStopStreamingName := IS_WORK_ENVIRONMENT ? pt_stopStreamingName : en_stopStreamingName

    ; Step 1: Send Escape to ensure composer is focused
    Send "{Esc}"
    Sleep 100
    ; Step 2: Send Enter to submit the prompt
    Send "{Enter}"
    Sleep 100
    ; Step 3: Alt+Tab to previous window
    Send "!{Tab}"
    Sleep 300
    ; Step 4: Wait for the Stop streaming button to appear, show indicator, then wait for it to disappear
    WaitForButtonAndShowSmallLoading_ChatGPT([currentStopStreamingName], "AI is responding…")
}

#HotIf

;-------------------------------------------------------------------
; Windows Explorer Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe explorer.exe")

; Shift + Y : Select first item in list (force-focus ItemsView)
+y::
{
    ; Send a right-click to shift focus into the main pane
    Click "Right"
    Sleep 300
    ; Clear any in-place edits or text focus first
    Send "{ESC}"

    EnsureItemsViewFocus()

    try {
        explorerEl := UIA.ElementFromHandle(WinExist("A"))

        itemsView := explorerEl.FindFirst({ AutomationId: "ItemsView", Type: "List" })
            ? explorerEl.FindFirst({ AutomationId: "ItemsView", Type: "List" })
            : explorerEl.FindFirst({ ClassName: "UIItemsView", Type: "List" })
                ? explorerEl.FindFirst({ ClassName: "UIItemsView", Type: "List" })
                : explorerEl.FindFirst({ Name: "Items View", Type: "List", matchmode: "Substring" })

        ; Fallback to entire window if we still did not find a dedicated list
        listRoot := itemsView ? itemsView : explorerEl

        ; Pick the very first ListItem inside that list root
        firstItem := listRoot.FindFirst({ Type: "ListItem" })

        if (firstItem) {
            firstItem.ScrollIntoView()
            firstItem.Select()
            firstItem.SetFocus()
            return
        }
    } catch Error {
        ; swallow and fallback below
    }

    ; Last-chance fallback – press Home which works if focus is already inside the list
    Send "{Home}"
}

; Helper to force focus to the ItemsView pane (file list)
EnsureItemsViewFocus() {
    try {
        explorerHwnd := WinExist("A")
        root := UIA.ElementFromHandle(explorerHwnd)

        ; quick check – if ItemsView already has keyboard focus, we're done
        iv := root.FindFirst({ AutomationId: "ItemsView", Type: "List" })
        if iv && iv.HasKeyboardFocus
            return

        ; Send up to 6 F6 cycles to reach the pane
        loop 6 {
            Send "{F6}"
            Sleep 120
            iv := root.FindFirst({ AutomationId: "ItemsView", Type: "List" })
            if iv && iv.HasKeyboardFocus
                break
        }
    } catch Error {
    }
}

; Shift + U : Focus search bar (Ctrl+E/F)
+u:: Send "^e"

; Shift + I : Focus address bar (Alt+D)
+i:: Send "!d"

; Shift + O : New folder
+o:: Send("^+n")

#HotIf

;-------------------------------------------------------------------
; Gmail Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("Gmail")

; Shift + Y: Go to main inbox (already implemented)
+y:: Send("gi")

; Shift + U: Go to updates
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

; Shift + O: Toggle read / unread on the selected message
+o::
{
    try
    {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 300

        ; Regex patterns for the buttons (English & Portuguese)
        readPattern := "i)^(Mark as read|Marcar como lida|Marcar como lido)$"
        unreadPattern := "i)^(Mark as unread|Marcar como n[oó] lida|Marcar como n[oó] lido)$"

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

; Shift + I: Go to forums
+i::
{
    try
    {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 300

        ; Try English and Portuguese names for the Forums tab
        forumsButton := uia.FindElement({ Name: "Forums", Type: "TabItem", matchmode: "Substring" })
        if (!forumsButton)
            forumsButton := uia.FindElement({ Name: "Fóruns", Type: "TabItem", matchmode: "Substring" })

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

; Shift + P: Previous conversation
+p:: Send("p")

; Shift + H: Next conversation
+h:: Send("n")

; Shift + J: Archive conversation
+j:: Send("e")

; Shift + K: Select conversation
+k:: Send("x")

; Shift + L: Reply
+l:: Send("r")

; Shift + N: Reply all
+n:: Send("a")

; Shift + M: Forward
+m:: Send("f")

; Shift + ,: Star/unstar conversation
+,:: Send("s")

; Shift + .: Delete
+.:: Send("#")

; Shift + W: Report as spam
+w:: Send("!")

; Shift + E: Compose new email
+e:: Send("c")

; Shift + R: Search mail
+r:: Send("/")

; Shift + T: Move to folder
+t:: Send("v")

; Shift + D: Show keyboard shortcuts help
+d:: Send("?")

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

; Shift + Y : Fold
+y::
{
    Send "^+8"
}

; Shift + U : Unfold
+u::
{
    Send "^+9"
}

; Shift + I : Fold all
+i::
{
    Send "^m"
    Send "^0"
}

; Shift + O : Unfold all
+o::
{
    Send "^m"
    Send "^j"
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

; Shift + W : Switch the brackets open/close
+w:: Send "^+]"

; Shift + E : Search
+e:: Send "^+f"

; Shift + R : Save all documents
+r::
{
    Send "^k"
    Send "s"
}

; Shift + T : Trigger Ctrl+;
+t:: Send "^;"

; --- Git Repository Selection Functions ---
;
; This section provides multi-repository support for VS Code Git operations.
;
; Features:
; • Automatically detects all Git repositories in the Source Control panel
; • Shows repository selection popup when multiple repositories are found
; • Auto-selects when only one repository is available
; • Provides fallback mechanisms for robust operation
; • Includes comprehensive error handling and user feedback
;
; Usage: All Git hotkeys (+f, +g, +c, +v, +b) now support repository selection
;

; Function to detect all Git repositories in VS Code Source Control panel
GetGitRepositories() {
    try {
        win := WinExist("A")
        if (!win) {
            return []
        }

        root := UIA.ElementFromHandle(win)
        if (!root) {
            return []
        }

        ; Find the Source Control Management tree - try multiple approaches
        scmTree := root.FindFirst({ ControlType: "Tree", Name: "Source Control Management" })
        if (!scmTree) {
            ; Alternative: try finding by AutomationId or other properties
            scmTree := root.FindFirst({ ControlType: "Tree", AutomationId: "scm-viewlet" })
        }
        if (!scmTree) {
            ; Try finding any tree in the source control area
            scmTrees := root.FindAll({ ControlType: "Tree" })
            for tree in scmTrees {
                try {
                    if (InStr(tree.Name, "Source Control") || InStr(tree.AutomationId, "scm")) {
                        scmTree := tree
                        break
                    }
                } catch {
                    continue
                }
            }
        }

        if (!scmTree) {
            return []
        }

        repositories := []
        ; Find all repository blocks - they appear as groups with "Git" labels
        allElements := scmTree.FindAll({ ControlType: "TreeItem" })

        for element in allElements {
            try {
                elementName := element.Name
                if (!elementName) {
                    continue
                }

                ; Look for repository names by finding labels that end with "Git"
                if (InStr(elementName, "Git")) {
                    ; Extract repository name by removing " Git" suffix
                    repoName := StrReplace(elementName, " Git", "")
                    repoName := Trim(repoName)

                    ; Additional validation: ensure it's not just "Git" and has actual content
                    if (repoName && repoName != "Git" && repoName != "") {
                        ; Check if we already have this repository to avoid duplicates
                        isDuplicate := false
                        for existingRepo in repositories {
                            if (existingRepo.name == repoName) {
                                isDuplicate := true
                                break
                            }
                        }

                        if (!isDuplicate) {
                            repositories.Push({ name: repoName, element: element })
                        }
                    }
                }
            } catch {
                continue
            }
        }

        return repositories
    } catch Error as e {
        ; Log error for debugging (optional)
        ; FileAppend("Git repo detection error: " . e.Message . "`n", "debug.log")
        return []
    }
}

; Function to show repository selection popup and return selected repository element
SelectRepository() {
    repositories := GetGitRepositories()

    if (repositories.Length == 0) {
        MsgBox(
            "No Git repositories found in Source Control panel.`n`nPlease ensure:`n• VS Code is open`n• Source Control panel is visible`n• At least one Git repository is loaded",
            "Repository Selection", "IconX T10000")
        return ""
    }

    if (repositories.Length == 1) {
        ; Only one repository, use it automatically
        return repositories[1].element
    }

    ; Multiple repositories - show selection dialog
    repoNames := ""
    for i, repo in repositories {
        repoNames .= i . ". " . repo.name . "`n"
    }

    try {
        result := InputBox("Select repository:`n`n" . repoNames . "`nEnter number (1-" . repositories.Length .
            ") or press Cancel to abort:", "Repository Selection", "w400 h250")

        if (result.Result == "Cancel") {
            return ""
        }

        selectedIndex := result.Value
        if (!selectedIndex || selectedIndex == "") {
            return ""
        }

        index := Integer(selectedIndex)
        if (index >= 1 && index <= repositories.Length) {
            return repositories[index].element
        } else {
            MsgBox("Invalid selection (" . selectedIndex . "). Please enter a number between 1 and " . repositories.Length .
                ".", "Repository Selection", "IconX T5000")
            return ""
        }
    } catch Error as e {
        if (InStr(e.Message, "cancelled") || InStr(e.Message, "Cancel")) {
            return ""
        }
        MsgBox("Invalid input. Please enter a valid number.", "Repository Selection", "IconX T5000")
        return ""
    }
}

; Function to find elements within a specific repository scope
FindInRepository(repoElement, searchCriteria) {
    try {
        ; First try to find in the repository element itself
        element := repoElement.FindFirst(searchCriteria)
        if (element) {
            return element
        }

        ; If not found, try to find in the parent tree structure
        ; This handles cases where buttons might be in sibling elements
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)
        scmTree := root.FindFirst({ ControlType: "Tree", Name: "Source Control Management" })

        if (scmTree) {
            ; Find all elements and filter by proximity to our repository
            allElements := scmTree.FindAll(searchCriteria)

            for testElement in allElements {
                try {
                    ; Check if this element is related to our repository
                    ; by walking up the tree to see if we find our repo element
                    current := testElement
                    while (current) {
                        if (current == repoElement) {
                            return testElement
                        }
                        try {
                            current := current.TreeWalker.GetParentElement(current)
                        } catch {
                            break
                        }
                    }
                } catch {
                    continue
                }
            }
        }

        return ""
    } catch {
        return ""
    }
}

; Function to test repository detection (for debugging/validation)
; Usage: Call TestRepositoryDetection() from AutoHotkey console or add temporary hotkey
TestRepositoryDetection() {
    repositories := GetGitRepositories()

    if (repositories.Length == 0) {
        MsgBox("No repositories detected.`n`nPlease ensure VS Code is open with Source Control panel visible.",
            "Repository Detection Test", "IconI")
        return
    }

    repoList := "Found " . repositories.Length . " repository(ies):`n`n"
    for i, repo in repositories {
        repoList .= i . ". " . repo.name . "`n"
    }

    MsgBox(repoList, "Repository Detection Test", "IconI")
}

; --- Git Shortcuts ---
; Shift + D : Git section
+d:: Send "^+g"

; Shift + F : Generate commit message with AI
+f::
{
    try {
        ; Select repository first to ensure we're working with the right context
        repoElement := SelectRepository()
        if (!repoElement) {
            return ; User cancelled or no repositories found
        }

        ; The Ctrl+M shortcut should work in the context of the selected repository
        ; First, try to focus on the repository's commit message input field
        messageInput := FindInRepository(repoElement, { ControlType: "Edit", Name: "Message", matchmode: "Substring" })
        if (!messageInput) {
            ; Fallback: try to find globally
            win := WinExist("A")
            root := UIA.ElementFromHandle(win)
            messageInput := root.FindFirst({ ControlType: "Edit", Name: "Message", matchmode: "Substring" })
        }

        if (messageInput) {
            messageInput.SetFocus()
            Sleep 100
        }

        ; Send the Ctrl+M command to generate commit message
        Send "^m"
    } catch Error as e {
        ; If there's an error with UIA, fall back to the simple command
        Send "^m"
    }
}

; Shift + G : Commit
+g::
{
    try {
        ; Select repository first
        repoElement := SelectRepository()
        if (!repoElement) {
            return ; User cancelled or no repositories found
        }

        ; Find commit button within the selected repository scope
        commitBtn := FindInRepository(repoElement, { ControlType: "Button", Name: "Commit Changes on", matchmode: "Substring" })
        if (!commitBtn) {
            ; Fallback: try to find globally and check if it's the right one
            win := WinExist("A")
            root := UIA.ElementFromHandle(win)
            commitBtn := root.FindFirst({ ControlType: "Button", Name: "Commit Changes on", matchmode: "Substring" })
        }

        if (commitBtn) {
            commitBtn.Click()
        } else {
            MsgBox("Could not find the 'Commit' button for the selected repository.", "VS Code Git", "IconX")
        }
    } catch Error as e {
        MsgBox("UIA Error: " e.Message, "VS Code Git Error", "IconX")
    }
}

; Shift + C : Commit and push
+c::
{
    try {
        ; Select repository first
        repoElement := SelectRepository()
        if (!repoElement) {
            return ; User cancelled or no repositories found
        }

        ; Find More Actions button within the selected repository scope
        moreActionsBtn := FindInRepository(repoElement, { ControlType: "MenuItem", Name: "More Actions..." })
        if (!moreActionsBtn) {
            moreActionsBtn := FindInRepository(repoElement, { ControlType: "Button", Name: "More Actions..." })
        }
        if (!moreActionsBtn) {
            ; Fallback: try to find globally
            win := WinExist("A")
            root := UIA.ElementFromHandle(win)
            moreActionsBtn := root.FindFirst({ ControlType: "MenuItem", Name: "More Actions..." })
            if (!moreActionsBtn) {
                moreActionsBtn := root.FindFirst({ ControlType: "Button", Name: "More Actions..." })
            }
        }

        if (moreActionsBtn) {
            moreActionsBtn.Click()
            Sleep 200
            Send "{Down 3}"
            Sleep 200
            Send "{Enter}"
        } else {
            MsgBox("Could not find the 'More Actions...' button for the selected repository.", "VS Code Git", "IconX")
        }
    } catch Error as e {
        MsgBox("UIA Error: " e.Message, "VS Code Git Error", "IconX")
    }
}

; Shift + V : Git pull
+v::
{
    try {
        ; Select repository first
        repoElement := SelectRepository()
        if (!repoElement) {
            return ; User cancelled or no repositories found
        }

        ; Find Pull button within the selected repository scope
        pullBtn := FindInRepository(repoElement, { ControlType: "Button", Name: "Pull" })
        if (!pullBtn) {
            ; Fallback: try to find globally
            win := WinExist("A")
            root := UIA.ElementFromHandle(win)
            pullBtn := root.FindFirst({ ControlType: "Button", Name: "Pull" })
        }

        if (pullBtn) {
            pullBtn.Click()
        } else {
            MsgBox("Could not find the 'Pull' button for the selected repository.", "VS Code Git", "IconX")
        }
    } catch Error as e {
        MsgBox("UIA Error: " e.Message, "VS Code Git Error", "IconX")
    }
}

; Shift + B : Git push
+b::
{
    try {
        ; Select repository first
        repoElement := SelectRepository()
        if (!repoElement) {
            return ; User cancelled or no repositories found
        }

        ; Find Push button within the selected repository scope
        pushBtn := FindInRepository(repoElement, { ControlType: "Button", Name: "Push" })
        if (!pushBtn) {
            ; Fallback: try to find globally
            win := WinExist("A")
            root := UIA.ElementFromHandle(win)
            pushBtn := root.FindFirst({ ControlType: "Button", Name: "Push" })
        }

        if (pushBtn) {
            pushBtn.Click()
        } else {
            MsgBox("Could not find the 'Push' button for the selected repository.", "VS Code Git", "IconX")
        }
    } catch Error as e {
        MsgBox("UIA Error: " e.Message, "VS Code Git Error", "IconX")
    }
}

#HotIf

;-------------------------------------------------------------------
; Spotify Shortcuts
;-------------------------------------------------------------------ww
#HotIf WinActive("ahk_exe Spotify.exe")

; Shift + Y : Toggle Connect panel
+y::
{
    try {
        spot := UIA_Browser("ahk_exe Spotify.exe")
        Sleep 300

        ; Find and click the Connect button
        connectPattern := "i)^(Connect to a device|Conectar a um dispositivo|Connect)$"
        if (connectBtn := WaitForButton(spot, connectPattern)) {
            connectBtn.Invoke()

            ; After connecting, wait a moment for the device list to load
            Sleep 500

            ; Search for Office button (e.g., "Office Google Cast")
            officePattern := "i)Office"
            if (officeBtn := WaitForButton(spot, officePattern, 3000)) {
                officeBtn.Invoke()
            }
            ; If Office button not found, continue without error (as requested)
        }
        else
            MsgBox "Couldn't find the Connect-to-device button."
    } catch Error as e {
        MsgBox "Error: " e.Message
    }
}

; Shift + U : Toggle full screen
+u::
{
    try {
        spot := UIA_Browser("ahk_exe Spotify.exe")
        Sleep 300

        ; Look for either Enter or Exit full screen button with case-insensitive pattern
        enterFsPattern := "i)^Enter Full[- ]?screen$"
        exitFsPattern := "i)^Exit Full[- ]?screen$"

        ; First attempt - immediate check
        enterFsBtn := WaitForButton(spot, enterFsPattern, 500)
        if (!enterFsBtn) {
            exitFsBtn := WaitForButton(spot, exitFsPattern, 500)
            if (!exitFsBtn) {
                ; Wait 1 second and try again
                Sleep 1000
                exitFsBtn := WaitForButton(spot, exitFsPattern, 500)
                if (exitFsBtn)
                    exitFsBtn.Invoke()
            } else {
                exitFsBtn.Invoke()
            }
        } else {
            enterFsBtn.Invoke()
        }
    }
}

; Shift + I : Open Search (Ctrl+K)
+i:: Send "^k"

; Shift + O : Go to Playlists (Alt+Shift+1)
+o:: Send "!+1"

; Shift + P : Go to Artists (Alt+Shift+3)
+p:: Send "!+3"

; Shift + H : Go to Albums (Alt+Shift+4)
+h:: Send "!+4"

; Shift + J : Go to Search (Ctrl+L)
+j:: Send "^l"

; Shift + K : Go to Home (Alt+Shift+H)
+k:: Send "!+h"

; Shift + L : Go to Now Playing (Alt+Shift+J)
+l:: Send "!+j"

; Shift + N : Go to Made For You (Alt+Shift+M)
+n:: Send "!+m"

; Shift + M : Go to New Releases (Alt+Shift+N)
+m:: Send "!+n"

; Shift + , : Go to Charts (Alt+Shift+C)
+,:: Send "!+c"

; Shift + . : Toggle Now Playing View Sidebar (Alt+Shift+R)
+.:: Send "!+r"

; Shift + W : Toggle Your Library Sidebar (Alt+Shift+L)
+w:: Send "!+l"

; Shift + E : Toggle Fullscreen Library
+e::
{
    try {
        spot := UIA_Browser("ahk_exe Spotify.exe")
        Sleep 300
        fullscreenLibBtn := spot.FindElement({ Name: "fullscreen library", Type: "Button" })
        if (fullscreenLibBtn) {
            fullscreenLibBtn.Click()
        } else {
            MsgBox "Could not find the fullscreen library button.", "Spotify Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error toggling fullscreen library: " e.Message, "Spotify Error", "IconX"
    }
}

; Shift + R : Toggle lyrics
+r::
{
    try {
        spot := UIA_Browser("ahk_exe Spotify.exe")
        Sleep 300

        ; Find and click the Lyrics button
        lyricsBtn := spot.FindElement({ Name: "Lyrics", Type: "Button" })
        if (lyricsBtn) {
            lyricsBtn.Click()
        } else {
            MsgBox "Could not find the Lyrics button.", "Spotify Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error toggling lyrics: " e.Message, "Spotify Error", "IconX"
    }
}

; Shift + T : Toggle play/pause
+t::
{
    ; Simple approach: try to find Play button, if not found just send media key
    try {
        spot := UIA_Browser("ahk_exe Spotify.exe")
        Sleep 300
        playBtn := spot.FindElement({ Name: "Play", Type: "Button" })
        if (playBtn) {
            playBtn.Click()
            return
        }
    }

    ; If we get here, either Play button wasn't found or there was an error
    ; In either case, just send media key to toggle play/pause
    Send "{Media_Play_Pause}"
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

; Shift + R : Align top (Alt + W)
+r:: Send("!w")

; Shift + T : Align bottom (Alt + S)
+t:: Send("!s")

; Shift + D : Align center horizontal (Alt + H)
+d:: Send("!h")

; Shift + F : Align center vertical (Alt + V)
+f:: Send("!v")

; Shift + G : Distribute horizontal spacing (Alt + Shift + H)
+g:: Send("!+h")

; Shift + W : Distribute vertical spacing (Alt + Shift + V)
+w:: Send("!+v")

; Shift + E : Tidy up (Ctrl + Alt + Shift + T)
+e:: Send("^!+t")

#HotIf

;-------------------------------------------------------------------
; Mobills Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("Mobills")

; Shift + Y : Dashboard
+y:: {
    try {
        btn := GetMobillsButton("menu-dashboard-item", "Dashboard")
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
        btn := GetMobillsButton("menu-accounts-item", "Accounts")
        if (btn) {
            btn.Click()
        } else {
            MsgBox "Could not find the Contas/Accounts button.", "Mobills Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error navigating to Contas/Accounts: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + I : Transações
+i:: {
    try {
        btn := GetMobillsButton("menu-transactions-item", "Transactions")
        if (btn) {
            btn.Click()
        } else {
            MsgBox "Could not find the Transações/Transactions button.", "Mobills Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error navigating to Transações/Transactions: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + O : Cartões de crédito
+o:: {
    try {
        btn := GetMobillsButton("menu-creditCards-item", "Credit cards")
        if (btn) {
            btn.Click()
        } else {
            MsgBox "Could not find the Cartões de crédito/Credit cards button.", "Mobills Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error navigating to Cartões de crédito/Credit cards: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + P : Planejamento
+p:: {
    try {
        btn := GetMobillsButton("menu-budgets-item", "Budgets")
        if (btn) {
            btn.Click()
        } else {
            MsgBox "Could not find the Planejamento/Budgets button.", "Mobills Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error navigating to Planejamento/Budgets: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + H : Relatórios
+h:: {
    try {
        btn := GetMobillsButton("menu-reports-item", "Reports")
        if (btn) {
            btn.Click()
        } else {
            MsgBox "Could not find the Relatórios/Reports button.", "Mobills Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error navigating to Relatórios/Reports: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + J : Mais opções
+j:: {
    try {
        btn := GetMobillsButton("menu-moreOptions-item", "More options")
        if (btn) {
            btn.Click()
        } else {
            MsgBox "Could not find the Mais opções/More options button.", "Mobills Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error navigating to Mais opções/More options: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + K : Previous month
+k:: PrevMobillsMonth()

; Shift + L : Next month
+l:: NextMobillsMonth()

PrevMobillsMonth() {
    try {
        uia := TryAttachBrowser()
        if !uia {
            MsgBox "Could not attach to the browser window.", "Mobills Navigation", "IconX"
            return
        }
        grp := FindMonthGroup(uia)
        if !grp {
            MsgBox "Could not locate the month container.", "Mobills Navigation", "IconX"
            return
        }
        ; First try: immediate previous sibling
        btn := grp.WalkTree("-1", { Type: "Button" })
        if !btn {
            ; Fallback: search all buttons inside parent and pick the one to the LEFT of the group
            parent := UIA.TreeWalkerTrue.GetParentElement(grp)
            if (parent) {
                grpPos := grp.Location
                for , el in parent.FindAll({ Type: "Button" }) {
                    pos := el.Location
                    if (pos.y >= grpPos.y - 10 && pos.y <= grpPos.y + grpPos.h + 10 && pos.x < grpPos.x) {
                        btn := el                          ; closest left candidate
                    }
                }
            }
        }
        if btn {
            btn.Click()
        } else {
            MsgBox "Could not find the previous-month button.", "Mobills Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error navigating to previous month:`n" e.Message, "Mobills Error", "IconX"
    }
}

NextMobillsMonth() {
    try {
        uia := TryAttachBrowser()
        if !uia {
            MsgBox "Could not attach to the browser window.", "Mobills Navigation", "IconX"
            return
        }
        grp := FindMonthGroup(uia)
        if !grp {
            MsgBox "Could not locate the month container.", "Mobills Navigation", "IconX"
            return
        }
        ; First try: immediate next sibling
        btn := grp.WalkTree("+1", { Type: "Button" })
        if !btn {
            ; Fallback: search all buttons inside parent and pick the one to the RIGHT of the group
            parent := UIA.TreeWalkerTrue.GetParentElement(grp)
            if (parent) {
                grpPos := grp.Location
                for , el in parent.FindAll({ Type: "Button" }) {
                    pos := el.Location
                    if (pos.y >= grpPos.y - 10 && pos.y <= grpPos.y + grpPos.h + 10 && pos.x > grpPos.x + grpPos.w) {
                        btn := el                          ; closest right candidate
                    }
                }
            }
        }
        if btn {
            btn.Click()
        } else {
            MsgBox "Could not find the next-month button.", "Mobills Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error navigating to next month:`n" e.Message, "Mobills Error", "IconX"
    }
}

; ---- New helper to jump from "Open" button ----
FocusViaOpenButton(tabs, pressSpace := false) {
    try {
        uia := TryAttachBrowser()
        if !uia
            return false
        ; Anchor = Button named "Open"
        openBtn := uia.FindElement({ Name: "Open", Type: "Button" })
        if !openBtn {
            ; fallback by class substring
            openBtn := uia.FindElement({ ClassName: "MuiAutocomplete-popupIndicator", Type: "Button", matchmode: "Substring" })
        }
        if !openBtn
            return false
        openBtn.SetFocus()
        Sleep 200
        ; Tab forward specified times
        loop tabs {
            Send "+{Tab}"
            Sleep 80
        }
        if pressSpace {
            Sleep 80
            Send "{Space}"
        }
        return true
    } catch Error {
        return false
    }
}

; Shift + N : Toggle "Ignore transaction"
+n:: FocusViaOpenButton(3, true)

; Shift + M : Focus expense name field
+m:: FocusViaOpenButton(2, false)

#HotIf

;-------------------------------------------------------------------
; Google Keep Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe chrome.exe") && (WinActive("Google Keep") || WinActive("keep.google.com") || InStr(
    WinGetTitle("A"), "Google Keep"))

; Shift + Y : Search and select note
+y::
{
    ; Store the current active window handle
    currentWindow := WinExist("A")

    ; Show message box to get search text from user
    searchText := InputBox("Enter text to search for in your notes:", "Google Keep Search", "w300 h100")

    if (searchText.Result = "OK" && searchText.Value != "") {
        ; Explicitly activate the Google Keep window to ensure we're working with the right window
        WinActivate("ahk_id " currentWindow)
        WinWaitActive("ahk_id " currentWindow, , 2)

        ; Store the search text in clipboard
        oldClip := A_Clipboard
        A_Clipboard := searchText.Value

        ; Wait a moment for clipboard to be ready
        Sleep 200

        ; Send Escape to clear any current selection/focus
        Send "{Esc}"
        Sleep 300

        ; Open search with Ctrl+F
        Send "^f"
        Sleep 200

        ; Paste the search text
        Send "^v"
        Sleep 900

        ; Press Escape to close search
        Send "{Esc}"
        Sleep 300

        ; Press Enter to confirm selection
        Send "{Enter}"
        Sleep 300

        ; Restore original clipboard
        A_Clipboard := oldClip
    }
}

; Shift + U : Toggle main menu
+u::
{
    try {
        ; Store the current active window handle
        currentWindow := WinExist("A")

        ; Explicitly activate the Google Keep window to ensure we're working with the right window
        WinActivate("ahk_id " currentWindow)
        WinWaitActive("ahk_id " currentWindow, , 2)

        ; Use UIA to find and click the main menu button
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 300 ; Give UIA time to attach

        ; Find the main menu button by its properties
        mainMenuBtn := uia.FindElement({
            Name: "Main menu",
            Type: "Button",
            ClassName: "gb_Lc"
        })

        if (mainMenuBtn) {
            mainMenuBtn.Click()
        } else {
            ; Fallback: try to find by name only
            mainMenuBtn := uia.FindElement({ Name: "Main menu", Type: "Button" })
            if (mainMenuBtn) {
                mainMenuBtn.Click()
            } else {
                MsgBox "Could not find the Main menu button.", "Google Keep", "IconX"
            }
        }
    } catch Error as e {
        MsgBox "Error toggling main menu: " e.Message, "Google Keep Error", "IconX"
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

; ---------------------------------------------------------------------------
; Helper for Mobills buttons – language-neutral search
; ---------------------------------------------------------------------------
GetMobillsButton(autoId, btnName) {
    try {
        root := UIA.ElementFromHandle(WinExist("A"))
        btn := root.FindFirst({ AutomationId: autoId, ControlType: "Button" })
        if !btn
            btn := root.FindFirst({ Name: btnName, ControlType: "Button" })
        return btn
    } catch Error {
        return ""
    }
}

;-------------------------------------------
; Helper functions
;-------------------------------------------
TryAttachBrowser() {
    ; Try Chrome first, then Edge
    try return UIA_Browser("ahk_exe chrome.exe")
    catch {
        try return UIA_Browser("ahk_exe msedge.exe")
        catch {
            return 0
        }
    }
}

FindMonthGroup(uia) {
    ; Strategy 1 – look for known class name on the container
    try {
        grp := uia.FindElement({ Type: "Group", ClassName: "sc-kAyceB", matchmode: "Substring" })
        if grp
            return grp
    }
    catch {
    }
    ; Strategy 2 – locate by month text (any language)
    months := ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
        "November", "December",
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro",
        "Dezembro"]
    for , m in months {
        try {
            el := uia.FindElement({ Name: m, Type: "Text", mm: 1, cs: false })
            if el {
                grp := el.WalkTree("p", { Type: "Group" })
                if grp
                    return grp
            }
        }
        catch {
        }
    }
    return 0
}

;-------------------------------------------------------------------
; YouTube Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe chrome.exe") && InStr(WinGetTitle("A"), "YouTube")

; Shift + Y : Focus search box
+y:: {
    try {
        uia := UIA_Browser()
        Sleep 300

        ; Try multiple search strategies
        searchBox := uia.FindFirst({ Type: "ComboBox", Name: "Search" })
        if !searchBox
            searchBox := uia.FindFirst({ Type: "Edit", Name: "Search" })
        if !searchBox
            searchBox := uia.FindFirst({ ClassName: "ytSearchboxComponentInput" })
        if !searchBox
            searchBox := uia.FindFirst({ Type: "SearchBox" })
        if !searchBox
            searchBox := uia.FindFirst({ AutomationId: "search" })

        if (searchBox) {
            searchBox.SetFocus()
            ; Additional fallback - if SetFocus doesn't work, try sending keyboard shortcut
            Sleep 100
            if !searchBox.HasKeyboardFocus {
                Send "/"  ; YouTube's built-in shortcut to focus search
            }
        } else {
            ; Last resort - just use YouTube's built-in keyboard shortcut
            Send "/"
        }
    } catch Error as e {
        ; If all else fails, use the keyboard shortcut
        Send "/"
    }
}

#HotIf

;-------------------------------------------------------------------
; File Dialog (Namespace Tree Control) Shortcuts
;-------------------------------------------------------------------
#HotIf IsFileDialogActive()

; Shift + Y : select first file (via header focus + Shift+Tab)
+y:: {
    try {
        root := UIA.ElementFromHandle(WinExist("A"))

        ; First attempt: find and focus Items View list directly
        itemsList := root.FindFirst({ Type: "List", ClassName: "UIItemsView" })
        if !itemsList
            itemsList := root.FindFirst({ Type: "List", Name: "Items View" })
        if !itemsList
            itemsList := root.FindFirst({ Type: "List", AutomationId: "ItemsView" })

        if itemsList {
            itemsList.SetFocus()
            Sleep 120
            Send "{Home}"  ; Go to first item
            return
        }

        ; Second attempt: find header (Header control)
        hdr := root.FindFirst({ Type: "Header" })
        if !hdr {
            hdr := root.FindFirst({ Name: "Header", Type: "Header" })
            if !hdr
                hdr := root.FindFirst({ Name: "Cabeçalho", Type: "Header" })
        }
        if hdr {
            hdr.SetFocus()
            Sleep 120
            Send "+{Tab}"
            Send "{Home}"
            return
        }

        ; Third attempt: find file name ComboBox by AutomationId and Type
        ; This should work regardless of the name (File name: or Nome:)
        fileNameCombo := root.FindFirst({ Type: "ComboBox", AutomationId: "1148" })

        ; If not found, try by various possible names
        if !fileNameCombo {
            possibleNames := ["File name:", "Nome:", "Filename:", "File Name:"]
            for name in possibleNames {
                fileNameCombo := root.FindFirst({ Type: "ComboBox", Name: name })
                if fileNameCombo
                    break
            }
        }

        ; If ComboBox found, use it
        if fileNameCombo {
            fileNameCombo.SetFocus()
            Sleep 120
            Send "+{Tab}"
            Send "{Home}"
            return
        }
    } catch Error {
    }
    ; Last resort fallback: simple Shift+Tab then Home
    Send "+{Tab}"
    Sleep 120
    Send "{Home}"
}

#HotIf

IsFileDialogActive() {
    hwnd := WinActive("A")
    if !hwnd {
        ToolTip("No active window")
        return false
    }

    winClass := WinGetClass("ahk_id " hwnd)
    winTitle := WinGetTitle("ahk_id " hwnd)
    winProcess := WinGetProcessName("ahk_id " hwnd)

    ToolTip("Window: " winTitle "`nClass: " winClass "`nProcess: " winProcess)
    SetTimer () => ToolTip(), -3000  ; clear after 3s

    if winClass != "#32770" {
        return false
    }

    try {
        root := UIA.ElementFromHandle(hwnd)
        ; Try to find ANY useful identifiers
        for type in ["List", "Tree", "Pane", "Window"] {
            elements := root.FindAll({ Type: type })
            if elements.Length {
                info := "Found " elements.Length " " type "s:`n"
                for el in elements {
                    info .= "- Name: " el.Name " AutoId: " el.AutomationId " Class: " el.ClassName "`n"
                }
                ToolTip(info)
                SetTimer () => ToolTip(), -5000
                break
            }
        }
        return true
    } catch Error as e {
        ToolTip("UIA Error: " e.Message)
        SetTimer () => ToolTip(), -3000
        return false
    }
}

;-------------------------------------------------------------------
; UIA Tree Inspector Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("UIATreeInspector") || WinActive("ahk_exe UIATreeInspectorAutoHotkey64.exe")

; Shift + Y : Refresh list
+y:: {
    try {
        root := UIA.ElementFromHandle(WinExist("A"))
        Sleep 200
        btn := root.FindFirst({ Name: "Refresh list", Type: "Button" })
        if !btn
            btn := root.FindFirst({ AutomationId: "5", Type: "Button" })
        if btn {
            btn.Invoke()
        } else {
            MsgBox "Could not find the Refresh list button.", "UIA Tree Inspector", "IconX"
        }
    } catch Error as e {
        MsgBox "Error refreshing list:`n" e.Message, "UIA Tree Inspector", "IconX"
    }
}

#HotIf

; Shift + U : Focus file name edit field
+u:: {
    try {
        root := UIA.ElementFromHandle(WinExist("A"))

        ; First attempt: Find by AutomationId and Type (most reliable)
        fileNameEdit := root.FindFirst({ Type: "Edit", AutomationId: "1148" })

        ; Second attempt: Try various possible names
        if !fileNameEdit {
            possibleNames := [
                "File name:",      ; English standard
                "Nome:",          ; Portuguese standard
                "Filename:",      ; Alternative English
                "File Name:",     ; Alternative capitalization
                "Name:",          ; Generic English
                "Nome do arquivo:", ; Full Portuguese
                "Save As:",       ; Save dialog English
                "Salvar como:"    ; Save dialog Portuguese
            ]
            for name in possibleNames {
                fileNameEdit := root.FindFirst({ Type: "Edit", Name: name })
                if fileNameEdit
                    break
            }
        }

        ; Third attempt: Try to find through parent ComboBox
        if !fileNameEdit {
            fileNameCombo := root.FindFirst({ Type: "ComboBox", AutomationId: "1148" })
            if fileNameCombo {
                fileNameEdit := fileNameCombo.FindFirst({ Type: "Edit" })
            }
        }

        if fileNameEdit {
            fileNameEdit.SetFocus()
            Sleep 50
            Send "^a"  ; Select all existing text
            return
        }
    } catch Error {
    }
    ; Fallback: Try to focus using keyboard navigation
    Send "!n"  ; Alt+N is a common shortcut for file name field
}

; Shift + I : Click Insert/Open/Save button
+i:: {
    try {
        root := UIA.ElementFromHandle(WinExist("A"))

        ; First attempt: Find by AutomationId and Type (most reliable)
        actionBtn := root.FindFirst({ Type: "Button", AutomationId: "1" })

        ; Second attempt: Try various possible names
        if !actionBtn {
            possibleNames := [
                ; English variations
                "Insert",
                "Open",
                "Save",
                "Save As",
                "OK",
                ; Portuguese variations
                "Abrir",
                "Salvar",
                "Salvar como",
                "Inserir",
                ; Spanish variations (common in some systems)
                "Insertar",
                "Guardar",
                "Guardar como",
                ; French variations (common in some systems)
                "Insérer",
                "Ouvrir",
                "Enregistrer",
                "Enregistrer sous"
            ]
            for name in possibleNames {
                actionBtn := root.FindFirst({ Type: "Button", Name: name })
                if actionBtn
                    break
            }
        }

        ; Third attempt: Try SplitButton type (some dialogs use this instead)
        if !actionBtn {
            actionBtn := root.FindFirst({ Type: "SplitButton", AutomationId: "1" })
            if !actionBtn {
                for name in possibleNames {
                    actionBtn := root.FindFirst({ Type: "SplitButton", Name: name })
                    if actionBtn
                        break
                }
            }
        }

        if actionBtn {
            actionBtn.Click()
            return
        }
    } catch Error {
    }
    ; Fallback: Try common keyboard shortcuts
    Send "!s"  ; Alt+S (Save)
    Sleep 50
    Send "!o"  ; Alt+O (Open)
}

; Shift + O : Click Cancel button
+o:: {
    try {
        root := UIA.ElementFromHandle(WinExist("A"))

        ; First attempt: Find by AutomationId and Type (most reliable)
        cancelBtn := root.FindFirst({ Type: "Button", AutomationId: "2" })

        ; Second attempt: Try various possible names
        if !cancelBtn {
            possibleNames := [
                ; English variations
                "Cancel",
                "Close",
                "Exit",
                "Dismiss",
                ; Portuguese variations
                "Cancelar",
                "Fechar",
                "Sair",
                ; Spanish variations
                "Cancelar",
                "Cerrar",
                ; French variations
                "Annuler",
                "Fermer",
                ; German variations
                "Abbrechen",
                "Schließen",
                ; Italian variations
                "Annulla",
                "Chiudi",
                ; Generic
                "No",
                "Não",
                "×",  ; Sometimes used as close symbol
                "✕"   ; Alternative close symbol
            ]
            for name in possibleNames {
                cancelBtn := root.FindFirst({ Type: "Button", Name: name })
                if cancelBtn
                    break
            }
        }

        if cancelBtn {
            cancelBtn.Click()
            return
        }
    } catch Error {
    }
    ; Fallback: Try common keyboard shortcuts
    Send "{Esc}"  ; Escape key is universal for cancel
}

;-------------------------------------------------------------------
; SettleUp Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("Settle Up")

; Shift + Y : Click Add Transaction button
+y:: {
    try {
        uia := UIA_Browser()
        Sleep 300

        ; Try multiple ways to find the Add Transaction button
        addBtn := uia.FindFirst({
            Name: "Adicionar transação"
        })

        ; If not found by full match, try partial match
        if !addBtn {
            possibleNames := [
                "Adicionar transação",    ; Portuguese
                "Add transaction",        ; English
                "Nueva transacción",      ; Spanish
                "Ajouter une transaction" ; French
            ]
            for name in possibleNames {
                addBtn := uia.FindFirst({ Name: name })
                if addBtn
                    break
            }
        }

        if addBtn {
            addBtn.Click()
            return
        }
    } catch Error as e {
        ; Silently handle errors
    }
}

; Shift + U : Focus expense value field (via name field + tab)
+u:: {
    try {
        uia := UIA_Browser()
        Sleep 300

        ; Find the "who paid" combo box
        paidByCombo := uia.FindFirst({
            Type: "ComboBox",
            Name: "Eduardo Figueiredo pagou"
        })

        ; If not found by exact match, try partial matches
        if !paidByCombo {
            possibleNames := [
                " pagou",           ; Portuguese suffix
                " paid",            ; English suffix
                " pagó",            ; Spanish suffix
                " a payé"           ; French suffix
            ]
            for suffix in possibleNames {
                paidByCombo := uia.FindFirst({
                    Type: "ComboBox",
                    Name: A_UserName . suffix,
                    matchmode: "Substring"
                })
                if paidByCombo
                    break
            }
        }

        ; Try by AutomationId if name matching failed
        if !paidByCombo {
            paidByCombo := uia.FindFirst({
                Type: "ComboBox",
                AutomationId: "mat-select-54"
            })
        }

        if paidByCombo {
            paidByCombo.Click()
            Sleep 100
            Send "{Tab}"  ; Move to expense value field
            return
        }
    } catch Error as e {
        ; Silently handle errors
    }
}

; Shift + I : Focus expense name field (via receipt button + shift-tab)
+i:: {
    try {
        uia := UIA_Browser()
        Sleep 300

        ; Find the "Add receipt" button
        receiptBtn := uia.FindFirst({
            Type: "Button",
            Name: "Adicionar comprovante",
            ClassName: "mdc-button mat-mdc-button mat-primary mat-mdc-button-base ng-star-inserted"
        })

        ; If not found by exact match, try other languages
        if !receiptBtn {
            possibleNames := [
                "Adicionar comprovante",    ; Portuguese
                "Add receipt",              ; English
                "Añadir recibo",           ; Spanish
                "Ajouter un reçu",         ; French
                "Beleg hinzufügen"         ; German
            ]
            for name in possibleNames {
                receiptBtn := uia.FindFirst({ Type: "Button", Name: name })
                if receiptBtn
                    break
            }
        }

        if receiptBtn {
            receiptBtn.SetFocus()  ; Just focus, don't click
            Sleep 100
            Send "+{Tab}"  ; Move backwards to expense name field
            return
        }
    } catch Error as e {
        ; Silently handle errors
    }
}

#HotIf

;-------------------------------------------------------------------
; Miro Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe chrome.exe") && InStr(WinGetTitle("A"), "Miro")

; Shift + Y : Command palette (Ctrl+K)
+y:: Send "^k"

; Shift + U : Frame list (Ctrl+Shift+F)
+u:: Send "^+f"

; Shift + I : Group (Ctrl+G)
+i:: Send "^g"

; Shift + O : Ungroup (Ctrl+Shift+G)
+o:: Send "^+g"

; Shift + P : Lock/Unlock (Ctrl+Shift+L)
+p:: Send "^+l"

; Shift + H : Add/Edit link (Alt+Ctrl+K)
+h:: Send "!^k"

#HotIf

; --- Duplicated small loading indicator functions from ChatGPT.ahk ---
global smallLoadingGuis_ChatGPT := []
ShowSmallLoadingIndicator_ChatGPT(state := "Loading…", bgColor := "00FF00") {
    global smallLoadingGuis_ChatGPT
    if (smallLoadingGuis_ChatGPT.Length > 0) {
        try {
            if (smallLoadingGuis_ChatGPT[1].Controls.Length > 0)
                smallLoadingGuis_ChatGPT[1].Controls[1].Text := state
        } catch {
            ; In case the GUI or control is invalid, proceed to recreate
        }
        return
    }

    ; --- Configuration for the simplified dual-border indicator ---
    colors := ["000000", "FFFFFF"] ; Black and White borders
    borderThickness := 2 ; pixels for each border
    alpha := 90 ; Reduced opacity for better visibility

    ; --- Central Text GUI ---
    textGui := Gui()
    textGui.Opt("+AlwaysOnTop -Caption +ToolWindow")
    textGui.BackColor := bgColor
    textGui.SetFont("s10 c000000 Bold", "Segoe UI")
    textGui.Add("Text", "w250 Center", state)
    smallLoadingGuis_ChatGPT.Push(textGui)

    ; --- Calculate Position ---
    activeWin := WinGetID("A")
    if (activeWin) {
        WinGetPos(&wx, &wy, &ww, &wh, activeWin)
    } else {
        work := SysGet.MonitorWorkArea(SysGet.MonitorPrimary)
        wx := work.Left, wy := work.Top, ww := work.Right - work.Left, wh := work.Bottom - work.Top
    }

    ; --- Show Text GUI to get its dimensions ---
    textGui.Show("AutoSize Hide")
    textGui.GetPos(, , &gw, &gh)
    cx := wx + (ww - gw) / 2
    cy := wy + (wh - gh) / 2

    ; --- Create Border GUIs ---
    currentW := gw, currentH := gh
    for color in colors {
        currentW += borderThickness * 2
        currentH += borderThickness * 2

        borderGui := Gui("+AlwaysOnTop -Caption +ToolWindow")
        borderGui.BackColor := color
        smallLoadingGuis_ChatGPT.Push(borderGui)

        xPos := cx - (currentW - gw) / 2
        yPos := cy - (currentH - gh) / 2

        borderGui.Show("x" Round(xPos) " y" Round(yPos) " w" Round(currentW) " h" Round(currentH) " NA")
        WinSetTransparent(alpha, borderGui.Hwnd)
    }

    ; --- Show the main text GUI on top ---
    textGui.Show("x" Round(cx) " y" Round(cy) " NA")
    WinSetTransparent(alpha, textGui.Hwnd)
}

HideSmallLoadingIndicator_ChatGPT() {
    global smallLoadingGuis_ChatGPT
    if (smallLoadingGuis_ChatGPT.Length > 0) {
        for gui in smallLoadingGuis_ChatGPT {
            try gui.Destroy()
        }
        smallLoadingGuis_ChatGPT := []
    }
}

WaitForButtonAndShowSmallLoading_ChatGPT(buttonNames, stateText := "Loading…", timeout := 15000) {
    ; Store ChatGPT's window handle before Alt+Tab
    chatGPTHwnd := WinExist("chatgpt")
    if !chatGPTHwnd {
        return ; ChatGPT window not found
    }

    ; Obtain UIA context for ChatGPT window specifically
    try cUIA := UIA_Browser("ahk_id " chatGPTHwnd)
    catch {
        return ; Failed to get UIA context
    }

    start := A_TickCount
    btn := ""

    ; Wait for the target button to appear and monitor it until it disappears
    while ((A_TickCount - start) < timeout) {
        btn := ""
        for n in buttonNames {
            try btn := cUIA.FindElement({ Name: n, Type: "Button" })
            catch {
                btn := ""
            }
            if btn
                break
        }
        if btn {
            ShowSmallLoadingIndicator_ChatGPT(stateText)
            while btn && ((A_TickCount - start) < timeout) {
                Sleep 250
                btn := ""
                for n in buttonNames {
                    try btn := cUIA.FindElement({ Name: n, Type: "Button" })
                    catch {
                        btn := ""
                    }
                    if btn
                        break
                }
            }
            break
        }
        Sleep 250
    }

    ; Always hide the indicator at the end
    HideSmallLoadingIndicator_ChatGPT()
}
