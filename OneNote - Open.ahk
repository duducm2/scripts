#Requires AutoHotkey v2.0+
#Include env.ahk

; Win+Alt+Shift+N to activate OneNote
#!+n::
{
    if oneNoteWin := WinExist("ahk_exe ONENOTE.EXE") {
        WinActivate(oneNoteWin)
    } else {
        if (IS_WORK_ENVIRONMENT) {
            Run "c:\ProgramData\Microsoft\Windows\Start Menu\Programs\OneNote.lnk"
        } else {
            Run "c:\ProgramData\Microsoft\Windows\Start Menu\Programs\OneNote.lnk"
        }
    }
}
