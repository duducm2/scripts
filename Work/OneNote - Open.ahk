#Requires AutoHotkey v2
#include C:\Users\fie7ca\Documents\UIA-v2\Lib\UIA.ahk

CapsLock & n:: {
    ; Try to find the existing OneNote window
    oneNoteWin := WinExist("ahk_exe ONENOTE.EXE")

    if (oneNoteWin) {
        ; If found, activate the window
        WinActivate(oneNoteWin)
    }
    Send("{CapsLock}")
}
