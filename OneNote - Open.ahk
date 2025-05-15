#Requires AutoHotkey v2
#include C:\Users\eduev\Documents\UIA-v2\Lib\UIA.ahk

!+n:: {
    ; Try to find the existing OneNote window
    oneNoteWin := WinExist("ahk_exe ONENOTE.EXE")

    if (oneNoteWin) {
        ; If found, activate the window
        WinActivate(oneNoteWin)
    }
}
