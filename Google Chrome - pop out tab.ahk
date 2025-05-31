#Requires AutoHotkey v2.0+

#!+l::
{
    ; Press F6 twice
    Send "{F6}"
    Sleep 100
    Send "{F6}"
    Sleep 100
    ; Open context menu
    Send "{AppsKey}"
    Sleep 100
    ; Press Down arrow three times
    Send "{Down}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Enter}"
}
