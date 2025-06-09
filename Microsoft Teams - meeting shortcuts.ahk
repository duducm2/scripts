#Requires AutoHotkey v2.0+

#include UIA-v2\Lib\UIA.ahk      ; ajuste o caminho se necessário

; -----------------------------------------------------------------
; Função utilitária – ativa a janela da reunião do Microsoft Teams
; (ignora janelas de chat e barra de compartilhamento)
; Compatível com o Teams "clássico" (Teams.exe) e o novo Teams
; (ms-teams.exe / MSTeams.exe)
; -----------------------------------------------------------------
ActivateTeamsMeetingWindow() {
    static processes := ["ms-teams.exe", "Teams.exe", "MSTeams.exe"]

    ; Procura entre todos os processos candidatos
    for proc in processes {
        for hwnd in WinGetList("ahk_exe " proc) {
            title := WinGetTitle(hwnd)
            if (
                (InStr(title, "Microsoft Teams meeting") 
                || RegExMatch(title, "^[^|]+\| Microsoft Teams$")
                || RegExMatch(title, "^.*\(.*\) \| Microsoft Teams$")
                )
                && !InStr(title, "Chat |")
                && !InStr(title, "Sharing control bar |")
            ) {
                WinActivate(hwnd)
                return true
            }
        }
    }

    ; Fallback: busca genérica por título caso o exe mude
    if hwnd := WinExist("Microsoft Teams meeting | Microsoft Teams") {
        WinActivate(hwnd)
        return true
    }

    MsgBox "Não foi possível encontrar uma janela de reunião ativa.", "Microsoft Teams", "Iconi"
    return false
}

; -----------------------------------------------------------------
; HOTKEYS (Win  + Alt + Shift + …)
; -----------------------------------------------------------------

; 1) Microfone  –  Win + Alt + Shift + Q
#!+q::                       ; Q
{
    if ActivateTeamsMeetingWindow()
        Send "^+m"           ; Ctrl+Shift+M
}

; 2) Câmera  –  Win + Alt + Shift + F
#!+f::                       ; F
{
    if ActivateTeamsMeetingWindow()
        Send "^+o"           ; Ctrl+Shift+O
}

; 3) Compartilhar tela  –  Win + Alt + Shift + E
#!+e::                       ; E
{
    if ActivateTeamsMeetingWindow() {
        ; garante que o foco não fique preso em menus anteriores
        Send "{Esc}"
        Sleep 200

        ; abre o menu de compartilhamento
        Send "^+e"           ; Ctrl+Shift+E
        Sleep 800            ; aguarda o menu carregar

        ; navega lentamente pelas opções (4 TABs)
        Loop 4 {
            Send "{Tab}"
            Sleep 200        ; deixa tempo para o Teams responder
        }
    }
}

; 4) Sair da reunião  –  Win + Alt + Shift + U
#!+u::                       ; U
{
    if !ActivateTeamsMeetingWindow()
        return
    response := MsgBox("Tem certeza de que deseja sair da reunião?", "Sair da reunião?", "YesNo Icon!")
    if response = "Yes"
        Send "^+h"           ; Ctrl+Shift+H
}
