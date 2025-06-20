#Requires AutoHotkey v2.0+

; Win+Alt+Shift+G to open Google Chrome
#!+f::
{
    Run "chrome.exe"
    WinWaitActive "ahk_exe chrome.exe"
}
