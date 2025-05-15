#Requires AutoHotkey v2

F12 & k::
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
