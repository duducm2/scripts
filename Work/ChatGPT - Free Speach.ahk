; ---------- CAPSLOCK + W  ----------
CapsLock & w:: {                ; mantém CapsLock pressionado e bate W
    SetTitleMatchMode(2)
    WinActivate("chatgpt - transcription")
    Send("{Esc}")
    Send("{CapsLock}")
}