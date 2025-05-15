#Requires AutoHotkey v2.0+

#include UIA-v2\Lib\UIA.ahk

; Win+Alt+Shift+A to open Outlook inbox
#!+a::
{
    ; Activate the inbox window
    SetTitleMatchMode 1
    WinActivate "Inbox - Eduardo"
}
