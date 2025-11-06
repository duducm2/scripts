#Requires AutoHotkey v2.0+
#SingleInstance Force

; -----------------------------------------------------------------------------
; This script consolidates all Outlook related hotkeys and functions.
; -----------------------------------------------------------------------------

; --- Includes ----------------------------------------------------------------
#include UIA-v2\Lib\UIA.ahk

; --- Hotkeys & Functions -----------------------------------------------------

; =============================================================================
; Open Outlook Mail
; Hotkey: Win+Alt+Shift+B
; Original File: Outlook - Open mail.ahk
; =============================================================================
#!+b::
{
    email := "Eduardo.Figueiredo@br.bosch.com"
    exclusion := "Calendar"
    for hwnd in WinGetList("ahk_exe OUTLOOK.EXE") {
        title := WinGetTitle(hwnd)
        if InStr(title, email) && !InStr(title, exclusion) {
            WinActivate(hwnd)
            return
        }
    }
}

; =============================================================================
; Open Outlook Calendar
; Hotkey: Win+Alt+Shift+G
; Original File: Outlook - Open calendar.ahk
; =============================================================================
#!+g::
{
    SetTitleMatchMode 1
    WinActivate "Calendar - Eduardo"
}

; =============================================================================
; Open Outlook Reminders
; Hotkey: Win+Alt+Shift+V
; Original File: Outlook - Open Reminder.ahk
; =============================================================================
#!+v::
{
    SetTitleMatchMode 2
    WinActivate "Reminder"
}

; =============================================================================
; Voice Aloud Email
; Hotkey: Win+Alt+Shift+D
; =============================================================================
#!+d::
{
    try {
        ; Remember current target window before showing GUI
        gVoiceAloudTargetWin := WinExist("A")
        ; Create GUI for voice aloud selection with auto-submit
        voiceAloudGui := Gui("+AlwaysOnTop +ToolWindow", "Voice Aloud Email")
        voiceAloudGui.SetFont("s10", "Segoe UI")

        ; Add instruction text
        voiceAloudGui.AddText("w350 Center",
            "Select voice aloud option:`n`n1) Voice aloud the email (from cursor)`n2) Voice aloud the email from the beginning`n`nType 1 or 2:"
        )

        ; Add input field with auto-submit functionality
        voiceAloudGui.AddEdit("w50 Center vVoiceAloudInput Limit1 Number")

        ; Add OK and Cancel buttons (as backup)
        voiceAloudGui.AddButton("w80 xp-40 y+10", "OK").OnEvent("Click", SubmitVoiceAloud)
        voiceAloudGui.AddButton("w80 xp+90", "Cancel").OnEvent("Click", CancelVoiceAloud)

        ; Set up auto-submit on text change
        voiceAloudGui["VoiceAloudInput"].OnEvent("Change", AutoSubmitVoiceAloud)

        ; Show GUI and focus input
        voiceAloudGui.Show("w350 h200")
        voiceAloudGui["VoiceAloudInput"].Focus()

    } catch Error as e {
        MsgBox "Error in voice aloud selector: " e.Message, "Voice Aloud Error", "IconX"
    }
}

; Global variable for voice aloud target window
global gVoiceAloudTargetWin := 0

; Function to get voice aloud option by number
GetVoiceAloudOptionByNumber(numberText) {
    try number := Integer(numberText)
    catch {
        return ""
    }
    optionMap := Map()
    optionMap[1] := "from_cursor"
    optionMap[2] := "from_beginning"
    return (optionMap.Has(number)) ? optionMap[number] : ""
}

; Function to execute voice aloud option
ExecuteVoiceAloudOption(option) {
    if (option = "")
        return

    ; First: Pause/stop music (do not risk resuming if already paused)
    Send "{Media_Stop}"
    Sleep 200

    if (option = "from_cursor") {
        ; Option 1: Voice aloud from cursor position
        Send "#!+b"  ; Go to Outlook email
        Sleep 300
        Send "{Alt down}1{Alt up}"  ; Alt+1 to start voice aloud
        Sleep 200
        Send "{Escape}"  ; Stop voice aloud
    }
    else if (option = "from_beginning") {
        ; Option 2: Voice aloud from beginning
        Send "#!+b"  ; Go to Outlook email
        Sleep 100
        Send "{Right}"
        Sleep 300
        Send "^{Home}"  ; Go to beginning of email
        Sleep 200
        Send "{Alt down}1{Alt up}"  ; Alt+1 to start voice aloud
        Sleep 200
        Send "{Escape}"  ; Stop voice aloud
    }
}

; Auto-submit function for voice aloud
AutoSubmitVoiceAloud(ctrl, *) {
    currentValue := ctrl.Text
    if (currentValue != "" && IsInteger(currentValue)) {
        choice := Integer(currentValue)
        if (choice >= 1 && choice <= 2) {
            ctrl.Gui.Destroy()
            ExecuteVoiceAloudOption(GetVoiceAloudOptionByNumber(currentValue))
        }
    }
}

; Manual submit function for voice aloud (backup)
SubmitVoiceAloud(ctrl, *) {
    currentValue := ctrl.Gui["VoiceAloudInput"].Text
    if (currentValue != "" && IsInteger(currentValue)) {
        choice := Integer(currentValue)
        if (choice >= 1 && choice <= 2) {
            ctrl.Gui.Destroy()
            ExecuteVoiceAloudOption(GetVoiceAloudOptionByNumber(currentValue))
        } else {
            MsgBox "Invalid selection. Please choose 1-2.", "Voice Aloud Selection", "IconX"
        }
    }
}

; Cancel function for voice aloud
CancelVoiceAloud(ctrl, *) {
    ctrl.Gui.Destroy()
}