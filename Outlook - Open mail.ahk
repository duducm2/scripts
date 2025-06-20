#Requires AutoHotkey v2.0+
#include UIA-v2\Lib\UIA.ahk   ; keep if you need UIA elsewhere

; Win+Alt+Shift+A → activate any Outlook window with your e-mail
; (except the Calendar window)
#!+b::
{
    email := "Eduardo.Figueiredo@br.bosch.com"
    exclusion := "Calendar"

    ; Grab every Outlook window
    for hwnd in WinGetList("ahk_exe OUTLOOK.EXE") {
        title := WinGetTitle(hwnd)
        if InStr(title, email) && !InStr(title, exclusion) {
            WinActivate(hwnd)
            return                       ; stop after the first suitable window
        }
    }
}
