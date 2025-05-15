#Requires AutoHotkey v2.0+

; Win+Alt+Shift+W to clear ChatGPT input
#!+w::
{
    SetTitleMatchMode(2)
    WinActivate("chatgpt - transcription")
    Send("{Esc}")
}
