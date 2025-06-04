#Requires AutoHotkey v2.0+
#SingleInstance Force
#include UIA-v2\Lib\UIA.ahk

#Include %A_ScriptDir%\env.ahk  ; IS_WORK_ENVIRONMENT (true/false)

; Using Win+Alt+E instead of Win+E to avoid conflicts
+#e:: OpenDesktopUIA()

OpenDesktopUIA() {
    SetTitleMatchMode 2
    ; 1. Try to find an Explorer window already on Desktop
    if WinExist("Área de Trabalho ahk_class CabinetWClass") || WinExist("Desktop ahk_class CabinetWClass") {
        WinActivate
        WinWaitActive
    } else {
        ; 2. Open Explorer directly to Desktop (work or personal)
        target := IS_WORK_ENVIRONMENT ? "C:\Users\fie7ca\Desktop" : "C:\Users\eduev\OneDrive\Desktop"
        Run 'explorer.exe "' target '"'
        WinWaitActive("ahk_class CabinetWClass", , 5)
    }
}
