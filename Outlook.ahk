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
