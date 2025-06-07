#Requires AutoHotkey v2.0+

#Include env.ahk ; Ensures IS_WORK_ENVIRONMENT is available

; Win+Alt+Shift+W to open WhatsApp
#!+z::
{
    SetTitleMatchMode(2)
    if WinExist("WhatsApp") {
        WinActivate("WhatsApp")
    } else {
        if (IS_WORK_ENVIRONMENT) {
            Run "C:\Users\fie7ca\Documents\Atalhos\WhatsApp.lnk"
        } else { ; Personal Environment
            Run "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\apps do Chrome\WhatsApp Web.lnk"
        }
    }
}
