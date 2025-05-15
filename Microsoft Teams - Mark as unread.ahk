#Requires AutoHotkey v2

F12 & y::
{
    Send "^1"
    Sleep "200"
    Send("{AppsKey}")
    Sleep "200"
    Send "{Down}"
    Send "{Enter}"
}
