#Requires AutoHotkey v2.0+
#SingleInstance Force
#include %A_ScriptDir%\env.ahk  ; define IS_WORK_ENVIRONMENT := true/false

; Win+Alt+Shift+S → abrir/ativar Spotify
#!+s:: OpenSpotify()

OpenSpotify() {
    SetTitleMatchMode(2)

    ; 1) Já está rodando? Só ativa.
    if WinExist("ahk_exe Spotify.exe") || WinExist("Spotify") {
        WinActivate
        return
    }

    ; 2) Decide como abrir de acordo com o ambiente
    global IS_WORK_ENVIRONMENT
    if IS_WORK_ENVIRONMENT  ; PC do trabalho
    {
        ; A) Usa o atalho que você mencionou
        link := "C:\Users\fie7ca\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Spotify.lnk"
        if FileExist(link) {
            Run(link)
            return
        }

        ; B) Se o .lnk não existir, tenta chamar o app da Microsoft Store
        storeApp := "explorer.exe shell:AppsFolder\SpotifyAB.SpotifyMusic_zpdnekdrzrea0!Spotify"
        Run(storeApp)
    }
    else  ; PC pessoal
    {
        ; Normalmente o alias "spotify" já resolve
        Run("spotify")
    }
}
