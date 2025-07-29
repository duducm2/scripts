#Requires AutoHotkey v2.0+
#SingleInstance Force
#UseHook  ; Ensure Volume hotkeys are captured before the OS processes them

; -----------------------------------------------------------------------------
; This script consolidates all Spotify related hotkeys and functions.
; -----------------------------------------------------------------------------

; --- Includes ----------------------------------------------------------------
#include UIA-v2\Lib\UIA.ahk
#include UIA-v2\Lib\UIA_Browser.ahk
#include %A_ScriptDir%\env.ahk

; --- Hotkeys & Functions -----------------------------------------------------

; =============================================================================
; Open or Activate Spotify
; Hotkey: Win+Alt+Shift+S
; Original File: Spotify - Open.ahk
; =============================================================================
#!+s:: OpenSpotify()

OpenSpotify() {
    SetTitleMatchMode(2)

    ; 1) If Spotify is already running, just activate it.
    if WinExist("ahk_exe Spotify.exe") || WinExist("Spotify") {
        WinActivate()
        if WinWaitActive("ahk_exe Spotify.exe", , 2)
            CenterMouse()
        return
    }

    ; 2) Decide how to open it based on the environment.
    global IS_WORK_ENVIRONMENT
    if IS_WORK_ENVIRONMENT {
        ; Work PC: Try the shortcut first.
        link := "C:\Users\fie7ca\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Spotify.lnk"
        if FileExist(link) {
            Run(link)
            if WinWaitActive("ahk_exe Spotify.exe", , 5)
                CenterMouse()
            return
        }

        ; Fallback for Work PC: Use the Store App command.
        Run("explorer.exe shell:AppsFolder\SpotifyAB.SpotifyMusic_zpdnekdrzrea0!Spotify")
        if WinWaitActive("ahk_exe Spotify.exe", , 5)
            CenterMouse()
    }
    else {
        ; Personal PC: Directly run the Microsoft Store App command.
        Run("explorer.exe shell:AppsFolder\SpotifyAB.SpotifyMusic_zpdnekdrzrea0!Spotify")
        if WinWaitActive("ahk_exe Spotify.exe", , 5)
            CenterMouse()
    }
}

*Volume_Down:: {
    if GetKeyState("Ctrl", "P") {
        ; Ctrl held: adjust Spotify volume
        spotify := ActivateSpotify()
        if WinWaitActive("ahk_exe Spotify.exe", , 2) {
            Send("^{Down}")
        }
        ; Handle window state after action
        if IsObject(spotify) {
            if spotify.wasMinimized {
                ; Re-minimize if the window was originally minimized
                hwnd := spotify.hwnd
                SetTimer((id := hwnd) => WinMinimize("ahk_id " id), -3500)
            } else {
                ; Use Alt+Tab to go back to previous window if it wasn't minimized
                ScheduleAltTab()
            }
        }
    } else if GetKeyState("Alt", "P") {
        ; Alt held: adjust YouTube volume
        yt := FocusYouTube()
        if IsObject(yt) {
            Send("{Down}")
            if yt.wasMinimized {
                hwnd := yt.hwnd
                SetTimer((id := hwnd) => WinMinimize("ahk_id " id), -3500)
            } else {
                ; Use Alt+Tab to go back to previous window if it wasn't minimized
                ScheduleAltTab()
            }
        } else {
            Send("{Volume_Down}")
        }
    } else {
        ; Default: master volume down
        Send("{Volume_Down}")
    }
}

*Volume_Up:: {
    if GetKeyState("Ctrl", "P") {
        spotify := ActivateSpotify()
        if WinWaitActive("ahk_exe Spotify.exe", , 2) {
            Send("^{Up}")
        }
        ; Handle window state after action
        if IsObject(spotify) {
            if spotify.wasMinimized {
                ; Re-minimize if the window was originally minimized
                hwnd := spotify.hwnd
                SetTimer((id := hwnd) => WinMinimize("ahk_id " id), -3500)
            } else {
                ; Use Alt+Tab to go back to previous window if it wasn't minimized
                ScheduleAltTab()
            }
        }
    } else if GetKeyState("Alt", "P") {
        yt := FocusYouTube()
        if IsObject(yt) {
            Send("{Up}")
            if yt.wasMinimized {
                hwnd := yt.hwnd
                SetTimer((id := hwnd) => WinMinimize("ahk_id " id), -3500)
            } else {
                ; Use Alt+Tab to go back to previous window if it wasn't minimized
                ScheduleAltTab()
            }
        } else {
            Send("{Volume_Up}")
        }
    } else {
        Send("{Volume_Up}")
    }
}

ActivateSpotify() {
    ; Activate Spotify and return information about its previous state
    winTitle := "ahk_exe Spotify.exe"
    hwnd := WinExist(winTitle)
    wasMinimized := hwnd && (WinGetMinMax(hwnd) == -1) ; -1 = minimized

    ; Bring the window to the foreground (restores if it was minimized)
    WinActivate(winTitle)

    ; Return a small object so callers can inspect state if desired
    return { hwnd: hwnd, wasMinimized: wasMinimized }
}

; =============================================================================
; Updated FocusYouTube to also return hwnd & minimized flag
; =============================================================================
FocusYouTube() {
    winList := WinGetList("ahk_exe chrome.exe")
    for win in winList {
        title := WinGetTitle(win)
        if InStr(title, "YouTube") {
            wasMinimized := (WinGetMinMax(win) == -1) ; check state before activation
            WinActivate(win)
            WinWaitActive(win, , 2)
            return { hwnd: win, wasMinimized: wasMinimized }
        }
    }
    return false ; not found
}

; =============================================================================
; Helper function to center mouse on the active window
; =============================================================================
CenterMouse() {
    Sleep(200)
    Send("#!+q")
}

; =============================================================================
; Helper function to schedule a single Alt+Tab after 3.5 seconds of inactivity
; =============================================================================
ScheduleAltTab() {
    ; Cancel any existing scheduled Alt+Tab so we only trigger it once
    SetTimer(DoAltTab, 0)
    ; Schedule a new one-shot Alt+Tab for 3.5 seconds from now
    SetTimer(DoAltTab, -3500)
}

DoAltTab() {
    Send("!{Tab}")
}
