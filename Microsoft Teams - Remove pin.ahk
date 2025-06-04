#Requires AutoHotkey v2.0+

; Win+Alt+Shift+K to remove pin from Teams message
#!+k::
{
    Sleep "500"
    Send "^1"
    Sleep "300"
    Send("{AppsKey}")
    Sleep "300"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}
