; AutoHotkey Version: 2
; Script to trigger three tab presses when pressing Control + Shift + Right Arrow

CapsLock & Up:: ; This is the hotkey definition for Control + Shift + Right Arrow
{
    Send "{Up}"
    Send "{Up}"
    Send "{Up}"
    Send "{Up}"
    Send "{Up}"
    Send("{CapsLock}")
}

CapsLock & Down:: ; This is the hotkey definition for Control + Shift + Right Arrow
{
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send("{CapsLock}")
}

CapsLock & Right:: ; This is the hotkey definition for Control + Shift + Right Arrow
{
    Send "{Right}"
    Send "{Right}"
    Send "{Right}"
    Send "{Right}"
    Send "{Right}"
    Send("{CapsLock}")
}

CapsLock & Left:: ; This is the hotkey definition for Control + Shift + Right Arrow
{
    Send "{Left}"
    Send "{Left}"
    Send "{Left}"
    Send "{Left}"
    Send "{Left}"
    Send("{CapsLock}")
}