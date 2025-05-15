#Requires AutoHotkey v2

#include C:\Users\fie7ca\Documents\UIA-v2\Lib\UIA.ahk

CapsLock & r::
{
    SetTitleMatchMode 2
    WinActivate "Reminder"
    Send("{CapsLock}")
}