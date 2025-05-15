#Requires AutoHotkey v2.0+

#include UIA-v2\Lib\UIA.ahk

; Win+Alt+Shift+9 to send to newsletter in Outlook
#!+9::
{
    Send "^+v"
    Send "{Home}"
    Send "01"
    Send "{Enter}"
}
