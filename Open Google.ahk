#Requires AutoHotkey v2

#include C:\Users\eduev\Documents\UIA-v2\Lib\UIA.ahk
#include C:\Users\eduev\Documents\UIA-v2\Lib\UIA_Browser.ahk

!+g::
{

    ; Run in Incognito mode to avoid any extensions interfering.
    Run "chrome.exe"
    WinWaitActive "ahk_exe chrome.exe"
    ; Sleep 500

    ; Initialize UIA_Browser, use Last Found Window
    ; cUIA := UIA_Browser()
}
