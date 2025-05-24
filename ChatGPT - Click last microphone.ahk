#Requires AutoHotkey v2.0+
#include UIA-v2\Lib\UIA.ahk
#include UIA-v2\Lib\UIA_Browser.ahk

; Win+Alt+Shift+T to toggle ChatGPT's read aloud (clicks "Read aloud" or "Stop" as appropriate)
#!+t::
{
    SetTitleMatchMode 2
    winTitle := "chatgpt - transcription"
    WinActivate winTitle
    WinWaitActive "ahk_exe chrome.exe"

    cUIA := UIA_Browser()
    Sleep 300

    ; First, try to find all "Stop" buttons (if ChatGPT is currently reading aloud)
    stopBtns := cUIA.FindAll({ Name: "Stop", Type: "Button" })
    if stopBtns.Length {
        stopBtns[stopBtns.Length].Click()
        return
    }

    ; Otherwise, find all "Read aloud" buttons and click the last one
    readBtns := cUIA.FindAll({ Name: "Read aloud", Type: "Button" })
    if readBtns.Length {
        readBtns[readBtns.Length].Click()
    } else {
        MsgBox "No 'Read aloud' or 'Stop' button found!"
    }
}
