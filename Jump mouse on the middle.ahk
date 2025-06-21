#Requires AutoHotkey v2.0+
#SingleInstance Force

#+!3::  ; Win + Alt + Shift + H
{
    hwnd := WinExist("A")  ; Get the handle of the active window
    if !hwnd {
        MsgBox "Still no active window!"
        return
    }

    ; Get window's screen rectangle
    rect := Buffer(16, 0)
    if !DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect) {
        MsgBox "GetWindowRect failed"
        return
    }

    left := NumGet(rect, 0, "int")
    top := NumGet(rect, 4, "int")
    right := NumGet(rect, 8, "int")
    bottom := NumGet(rect, 12, "int")

    ; Calculate center
    centerX := left + (right - left) // 2
    centerY := top + (bottom - top) // 2

    ; Move the mouse cursor instantly
    DllCall("SetCursorPos", "int", centerX, "int", centerY)
}
