#Requires AutoHotkey v2

F12 & n:: {
    ; Try to find the existing OneNote window
    oneNoteWin := WinExist("ahk_exe ONENOTE.EXE")

    if (oneNoteWin) {
        ; If found, activate the window
        WinActivate(oneNoteWin)
    }
}
