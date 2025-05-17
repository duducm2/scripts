#Requires AutoHotkey v2.0+

; CapsLock+Win+[Arrow keys] for fast movement
CapsLock & Up::
{
    if GetKeyState("LWin", "P") || GetKeyState("RWin", "P") {
        Send "{Up}"
        Send "{Up}"
        Send "{Up}"
        Send "{Up}"
        Send "{Up}"
        Send "{CapsLock}"
    }
    else
        Send "{Blind}{Up}"
}

CapsLock & Down::
{
    if GetKeyState("LWin", "P") || GetKeyState("RWin", "P") {
        Send "{Down}"
        Send "{Down}"
        Send "{Down}"
        Send "{Down}"
        Send "{Down}"
        Send "{CapsLock}"
    }
    else
        Send "{Blind}{Down}"
}

CapsLock & Right::
{
    if GetKeyState("LWin", "P") || GetKeyState("RWin", "P") {
        Send "{Right}"
        Send "{Right}"
        Send "{Right}"
        Send "{Right}"
        Send "{Right}"
        Send "{CapsLock}"
    }
    else
        Send "{Blind}{Right}"
}

CapsLock & Left::
{
    if GetKeyState("LWin", "P") || GetKeyState("RWin", "P") {
        Send "{Left}"
        Send "{Left}"
        Send "{Left}"
        Send "{Left}"
        Send "{Left}"
        Send "{CapsLock}"
    }
    else
        Send "{Blind}{Left}"
}
