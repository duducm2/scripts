#Requires AutoHotkey v2.0+

; Win+Alt+Shift+S to activate Spotify
#!+s::
{
    SetTitleMatchMode(2)
    if WinExist("Spotify") {
        WinActivate("Spotify")
    } else {
        Run "spotify"
    }
}
