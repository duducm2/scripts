#Requires AutoHotkey v2

!+4:: {  ; Win + A hotkey
    ; Set title match mode to find the word anywhere in the title
    SetTitleMatchMode 2  ; Mode 2 means the WinTitle can contain the title anywhere

    ; Try to activate a window containing "YouTube"
    if WinExist("YouTube") {
        WinActivate
    } else {
        ; If no YouTube window exists, open YouTube in Chrome
        Run "chrome.exe https://www.youtube.com"
        WinWaitActive "ahk_exe chrome.exe", , 10  ; Wait up to 10 seconds for Chrome
    }
}
