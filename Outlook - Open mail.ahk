#Requires AutoHotkey v2

#include C:\Users\fie7ca\Documents\UIA-v2\Lib\UIA.ahk

CapsLock & a::
{
    ; Activate the "Google Tradutor" window
    SetTitleMatchMode 1
    WinActivate "Inbox - Eduardo"
    Send("{CapsLock}")
}