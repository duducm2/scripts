; Requires AutoHotkey v2.0+
#include UIA-v2\Lib\UIA.ahk
#include UIA-v2\Lib\UIA_Browser.ahk

; Win+Alt+Shift+T  →  alterna o “Ler em voz alta / Read aloud”
#!+t::
{
    SetTitleMatchMode 2
    winTitle := "chatgpt"
    WinActivate winTitle
    WinWaitActive "ahk_exe chrome.exe"

    cUIA := UIA_Browser()
    Sleep 300

    ; --- Nomes dos botões em cada idioma ------------------------------------
    readNames := ["Read aloud", "Ler em voz alta"]
    stopNames := ["Stop", "Parar"]               ; “Parar” por segurança
    ; ------------------------------------------------------------------------

    ; 1) Tenta encontrar qualquer botão de “Stop / Parar”
    stopBtns := []
    for name in stopNames
        for btn in cUIA.FindAll({ Name: name, Type: "Button" })
            stopBtns.Push(btn)

    if stopBtns.Length
    {
        stopBtns[stopBtns.Length].Click()        ; clica no último encontrado
        return
    }

    ; 2) Caso não esteja lendo, procura “Read aloud / Ler em voz alta”
    readBtns := []
    for name in readNames
        for btn in cUIA.FindAll({ Name: name, Type: "Button" })
            readBtns.Push(btn)

    if readBtns.Length
        readBtns[readBtns.Length].Click()        ; clica no último encontrado
    else
        MsgBox "Nenhum botão 'Read aloud/Ler em voz alta' ou 'Stop/Parar' encontrado!"
}
