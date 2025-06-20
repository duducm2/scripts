#Requires AutoHotkey v2.0+

#include UIA-v2\Lib\UIA.ahk

; Win+Alt+Shift+X to open Outlook calendar
#!+g::
{
    ; Activate the "Calendar - Eduardo" window
    SetTitleMatchMode 1
    WinActivate "Calendar - Eduardo"
}
