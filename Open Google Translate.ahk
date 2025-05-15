#Requires AutoHotkey v2

#include C:\Users\eduev\Documents\UIA-v2\Lib\UIA.ahk
#include C:\Users\eduev\Documents\UIA-v2\Lib\UIA_Browser.ahk

!+t::
{

    Run "chrome.exe https://translate.google.com.br/?hl=en&en=ru&sl=en&tl=pt&op=translate -incognito"
    WinWaitActive "ahk_exe chrome.exe"
    Sleep 3000 ; Give enough time to load the page

}
