#Requires AutoHotkey v2.0

#UseHook  ; <-- This is important

; Change the default shortcuts (Let's say you want ctrl + , instead of ctrl + ;)
CapsLock & 0:: {
    Send "!ç"
    Send("{CapsLock}")
}
