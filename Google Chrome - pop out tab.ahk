#Requires AutoHotkey v2.0
#HotIf WinActive("ahk_exe chrome.exe")  ; Only active when Chrome is the active window
#!+l:: {  ; Ctrl+Alt+N triggers tab detach
    Send "{F6}"                        ; Focus address bar (omnibox)
    Sleep 100
    Send "{F6}"                        ; Focus the tab strip (current tab):contentReference[oaicite:6]{index=6}:contentReference[oaicite:7]{index=7}
    Sleep 100
    Send "{AppsKey}"                   ; Open the tab's context menu (AppsKey or Shift+F10)
    Sleep 100                          ; Wait a moment for menu to open
    Send "m"                           ; Select "Move tab to new window" (press 'm'):contentReference[oaicite:8]{index=8}
    Sleep 100
    Send "{Enter}"                     ; Confirm the action (detach tab)
    Sleep 100
    Send "{Enter}"                     ; Confirm the action (detach tab)
}
#HotIf