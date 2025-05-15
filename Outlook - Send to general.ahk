#Requires AutoHotkey v2

#include C:\Users\fie7ca\Documents\UIA-v2\Lib\UIA.ahk

CapsLock & 8::
{
    Send "^+v"
    Send "{Home}"
    Send "00"
    Send "{Enter}"
    Send("{CapsLock}")
}