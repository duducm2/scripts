#Requires AutoHotkey v2

#include C:\Users\fie7ca\Documents\UIA-v2\Lib\UIA.ahk

CapsLock & 9::
{
    Send "^+v"
    Send "{Home}"
    Send "01"
    Send "{Enter}"
    Send("{CapsLock}")
}