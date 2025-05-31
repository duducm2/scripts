#Requires AutoHotkey v2.0+

; Win+Alt+Shift+C to open ChatGPT
#!+c::
{
    SetTitleMatchMode(2)
    if WinExist("chatgpt") {
        WinActivate("chatgpt")
        Send("{Esc}")
    } else {
        Run "chrome.exe --new-window https://chatgpt.com/"
    }
}
