#Requires AutoHotkey v2.0+
#SingleInstance Force

#include UIA-v2\Lib\UIA.ahk
#include UIA-v2\Lib\UIA_Browser.ahk
#include %A_ScriptDir%\env.ahk

#!+b:: CopyLastPrompt()

CopyLastPrompt() {
    global IS_WORK_ENVIRONMENT

    static pt_copyName := "Copiar"
    static en_copyName := "Copy"
    copyName := IS_WORK_ENVIRONMENT ? pt_copyName : en_copyName

    SetTitleMatchMode 2
    WinActivate "chatgpt"
    WinWaitActive "ahk_exe chrome.exe"
    cUIA := UIA_Browser()
    Sleep 300

    try {
        ; Find all elements named "Copy" (or "Copiar")
        buttons := cUIA.FindAll({ Name: copyName, Type: "Button" })

        if buttons.Length {
            lastBtn := buttons[buttons.Length]  ; get the LAST match
            lastBtn.Click()
        } else {
            MsgBox(IS_WORK_ENVIRONMENT ? "Nenhum botão 'Copiar' encontrado." : "No 'Copy' button found.")
        }
    } catch as e {
        MsgBox(IS_WORK_ENVIRONMENT ? "Erro ao copiar:`n" e.Message : "Error copying:`n" e.Message)
    }
}
