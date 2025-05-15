#Requires AutoHotkey v2.0+

; Win+Alt+Shift+J to pin message in Teams
#!+j::
{
    Send "^1"
    Sleep "200"
    Send("{AppsKey}")
    Sleep "200"
    Send "{Down}"
    Send "{Down}"
    Send "{Right}"
    Send "{Enter}"
}
