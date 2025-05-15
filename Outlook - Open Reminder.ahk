#Requires AutoHotkey v2.0+

#include UIA-v2\Lib\UIA.ahk

; Win+Alt+Shift+R to open Outlook reminders
#!+r::
{
    SetTitleMatchMode 2
    WinActivate "Reminder"
}
