#Requires AutoHotkey v2

#include C:\Users\fie7ca\Documents\UIA-v2\Lib\UIA.ahk

CapsLock & 2::
{
    Send("{AppsKey}")
    Sleep 1500
    ; Send("{c}")
    ; Send("{o}")
    Send("{CapsLock}")
}