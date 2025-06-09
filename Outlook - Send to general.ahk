#Requires AutoHotkey v2.0+

#include UIA-v2\Lib\UIA.ahk

; Win+Alt+Shift+8 to send to general in Outlook
+^8::
{
    Send "!5"
    Send "O"
    Send "{Home}"
    Send "00"
    Send "{Enter}"
}
