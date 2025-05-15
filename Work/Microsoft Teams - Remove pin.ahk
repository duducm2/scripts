#Requires AutoHotkey v2


CapsLock & k::
{
	Send "^1"
    Sleep "200"
    Send("{AppsKey}")
    Sleep "200"
	Send "{Down}"
	Send "{Down}"
	Send "{Down}"
	Send "{Enter}"
	Send("{CapsLock}")
}