#Requires AutoHotkey v2.0+

; Win+Alt+Shift+N to activate OneNote
#!+n::
{
    if oneNoteWin := WinExist("ahk_exe ONENOTE.EXE") {
        WinActivate(oneNoteWin)
    } else {
        Run "c:\ProgramData\Microsoft\Windows\Start Menu\Programs\OneNote.lnk"
    }
}
