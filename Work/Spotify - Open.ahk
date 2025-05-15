#Requires AutoHotkey v2
#include C:\Users\fie7ca\Documents\UIA-v2\Lib\UIA.ahk

CapsLock & s:: {
    ; Try to find the existing Spotify window
    spotifyWin := WinExist("ahk_exe Spotify.exe")

    if (spotifyWin) {
        ; If found, activate the window
        WinActivate(spotifyWin)
    }
    Send("{CapsLock}")
}
