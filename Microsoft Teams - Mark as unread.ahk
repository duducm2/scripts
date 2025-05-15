#Requires AutoHotkey v2.0+

; Win+Alt+Shift+U to mark Teams message as unread
#!+y::
{
    Send "^1"
    Sleep "200"
    Send("{AppsKey}")
    Sleep "200"
    Send "{Down}"
    Send "{Enter}"
}
