#Requires AutoHotkey v2.0+

; Win+Alt+Shift+K to remove pin from Teams message
#!+k::
{
    Send "^1"
    Sleep "200"
    Send("{AppsKey}")
    Sleep "200"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}
