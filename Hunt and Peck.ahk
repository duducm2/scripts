#Requires AutoHotkey v2.0

Run "C:\Users\eduev\OneDrive\Documentos\HuntAndPeck\hap.exe"

#UseHook  ; <-- This is important

; Change the default shortcuts (Let's say you want `ctrl + ,` instead of `ctrl + ;`)
!+0:: {
    ; Send "{Escape}"
    ; Send "{Escape}"
    Send "!ç"
}
