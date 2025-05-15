#Requires AutoHotkey v2

#include C:\Users\fie7ca\Documents\UIA-v2\Lib\UIA.ahk

CapsLock & x::
{
    ; Activate the "Google Tradutor" window
    SetTitleMatchMode 1
    WinActivate "Calendar - Eduardo"
    Send("{CapsLock}")
}