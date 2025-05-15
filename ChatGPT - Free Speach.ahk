#Requires AutoHotkey v2

!+w::
{
    ; Activate the window containing "chatgpt - transcription"
    SetTitleMatchMode 2
    WinActivate "chatgpt - transcription"
    ; WinWaitActive "ahk_exe chrome.exe" ; Ensure the correct window is active

    Send "{Esc}" ; Focus main area
}
