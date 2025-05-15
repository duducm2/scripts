#Requires AutoHotkey v2.0+

; Win+Alt+Shift+G to open Google Chrome
#!+g::
{
    Run "chrome.exe"
    WinWaitActive "ahk_exe chrome.exe"
}
