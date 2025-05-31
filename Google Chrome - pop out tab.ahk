#Requires AutoHotkey v2.0+

#!+l::
{
    ; Press F6 twice
    Send "{F6}"
    Sleep 30
    Send "{F6}"
    Sleep 30
    ; Open context menu
    Send "{AppsKey}"
    Sleep 100
    ; Press Down arrow three times
    Send "{Down}"
    Sleep 30
    Send "{Down}"
    Sleep 30
    Send "{Down}"
    Sleep 30
    Send "{Enter}"
}
