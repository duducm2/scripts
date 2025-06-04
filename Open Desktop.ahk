#Requires AutoHotkey v2.0+
#InputLevel 2
#SingleInstance Ignore

#Include %A_ScriptDir%\env.ahk  ; IS_WORK_ENVIRONMENT (true/false)

; Win + E  →  abrir ou ativar janela do Desktop
#e:: OpenDesktop()

OpenDesktop() {
    SetTitleMatchMode 2
    desktopTitles := ["Desktop", "Área de Trabalho"]

    ; 1. Tenta ativar qualquer Explorer já aberto que tenha “Desktop” no título
    for title in desktopTitles {
        if WinExist(title " ahk_class CabinetWClass") {
            WinRestore
            WinActivate
            return
        }
    }

    ; 2. Se não existe, abre Desktop (work ou pessoal)
    personal := A_Desktop
    work     := "C:\Users\fie7ca\Desktop"
    target   := IS_WORK_ENVIRONMENT ? work : personal
    Run 'explorer.exe "' target '"'

    ; 3. Espera a janela aparecer e ativa
    Loop 50 {                      ; até 10 s (50 × 200 ms)
        Sleep 200
        for title in desktopTitles {
            if WinExist(title " ahk_class CabinetWClass") {
                WinRestore
                WinActivate
                return
            }
        }
    }
}
