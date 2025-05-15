; AutoHotkey Version: 2
; Script to trigger three tab presses when pressing Control + Shift + Right Arrow

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