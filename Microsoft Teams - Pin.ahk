#Requires AutoHotkey v2


CapsLock & j::
{
	Send "^1"
    Sleep "200"
    Send("{AppsKey}")
    Sleep "200"
	Send "{Down}"
	Send "{Down}"
	Send "{Right}"
	Send "{Enter}"
	Send("{CapsLock}")
}