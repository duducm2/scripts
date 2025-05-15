#Requires AutoHotkey v2

F12 & s:: {
    ; Try to find the existing Spotify window
    spotifyWin := WinExist("ahk_exe Spotify.exe")

    if (spotifyWin) {
        ; If found, activate the window
        WinActivate(spotifyWin)
    }
}
