#Requires AutoHotkey v2


CapsLock & y::
{
	Send "^1"
    Sleep "200"
    Send("{AppsKey}")
    Sleep "200"
	Send "{Down}"
	Send "{Enter}"
    Send("{CapsLock}")
}