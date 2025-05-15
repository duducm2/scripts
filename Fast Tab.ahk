#Requires AutoHotkey v2.0+
#InputLevel 2
#UseHook
#SingleInstance Ignore

; #UseHook                                   ; grab the key before Windows/apps do
; #SingleInstance Force

#!Right:: ; This is the hotkey definition for Control + Shift + Right Arrow
{
    Send "{Tab}" ; This sends the Tab key three times
    Sleep "10"
    Send "{Tab}" ; This sends the Tab key three times
    Sleep "10"
    Send "{Tab}" ; This sends the Tab key three times
    Sleep "10"
    Send "{Tab}" ; This sends the Tab key three times
}

#!Left:: ; This is the hotkey definition for Control + Shift + Left Arrow
{
    Send "+{Tab}" ; This sends the+ Tab key three times
    Sleep "10"
    Send "+{Tab}" ; This sends the+ Tab key three times
    Sleep "10"
    Send "+{Tab}" ; This sends the Tab key three times
    Sleep "10"
    Send "+{Tab}" ; This sends the Tab key three times
}

#!Up:: ; This is the hotkey definition for Control + Shift + Left Arrow
{
    Send "+{Tab}" ; This sends the+ Tab key three times
    Sleep "10"
    Send "+{Tab}" ; This sends the+ Tab key three times
    Sleep "10"
    Send "+{Tab}" ; This sends the Tab key three times
    Sleep "10"
    Send "+{Tab}" ; This sends the Tab key three times
    Sleep "10"
    Send "+{Tab}" ; This sends the Tab key three times
    Sleep "10"
    Send "+{Tab}" ; This sends the Tab key three times
    Sleep "10"
    Send "+{Tab}" ; This sends the Tab key three times
    ; Sleep "10"
    ; Send "+{Tab}" ; This sends the Tab key three times
    ; Sleep "10"
    ; Send "+{Tab}" ; This sends the Tab key three times
}

#!Down:: ; This is the hotkey definition for Control + Shift + Left Arrow
{
    Send "{Tab}" ; This sends the+ Tab key three times
    Sleep "10"
    Send "{Tab}" ; This sends the+ Tab key three times
    Sleep "10"
    Send "{Tab}" ; This sends the Tab key three times
    Sleep "10"
    Send "{Tab}" ; This sends the Tab key three times
    Sleep "10"
    Send "{Tab}" ; This sends the Tab key three times
    Sleep "10"
    Send "{Tab}" ; This sends the Tab key three times
    Sleep "10"
    Send "{Tab}" ; This sends the Tab key three times
    ; Sleep "10"
    ; Send "{Tab}" ; This sends the Tab key three times
    ; Sleep "10"
    ; Send "{Tab}" ; This sends the Tab key three times
}
