#Requires AutoHotkey v2.0+

; Win+Alt+Shift+J to pin message in Teams
#!+j::
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
