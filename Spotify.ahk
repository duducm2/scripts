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

; =============================================================================
; Go to Spotify Library
; Hotkey: Win+Alt+Shift+D
; Original File: Spotify - Go to library.ahk
; =============================================================================
#!+d:: GoToSpotifyLibrary()

GoToSpotifyLibrary() {
    loadingGui := "" ; Defer GUI creation

    try {
        SetTitleMatchMode(2)
        spotifyWin := "ahk_exe Spotify.exe"

        if !WinExist(spotifyWin) {
            ; If Spotify isn't running, create a GUI centered on the primary screen.
            loadingGui := Gui()
            loadingGui.Opt("+AlwaysOnTop -Caption +ToolWindow")
            loadingGui.BackColor := "333333"
            loadingGui.SetFont("s12 cFFFFFF", "Segoe UI")
            loadingGui.Add("Text", "w300 Center", "Spotify is opening...")
            loadingGui.Show("AutoSize Center NA")
            WinSetTransparent(220, loadingGui)
            OpenSpotify()
            Sleep(4000)
            return ; Exit after handling the opening process
        }

        ; --- Spotify window exists, so center the GUI on it ---
        WinActivate(spotifyWin)
        if WinWaitActive(spotifyWin, , 2) {
            CenterMouse()
        }

        ; Get Spotify window's position and size
        spotifyX := 0, spotifyY := 0, spotifyW := 0, spotifyH := 0
        WinGetPos(&spotifyX, &spotifyY, &spotifyW, &spotifyH, spotifyWin)

        ; Create and measure the loading GUI
        loadingGui := Gui()
        loadingGui.Opt("+AlwaysOnTop -Caption +ToolWindow")
        loadingGui.BackColor := "333333"
        loadingGui.SetFont("s12 cFFFFFF", "Segoe UI")
        statusText := loadingGui.Add("Text", "w300 Center", "Navigating to your library...")
        loadingGui.Show("AutoSize Hide") ; Show hidden to get its size
        guiW := 0, guiH := 0
        loadingGui.GetPos(, , &guiW, &guiH)

        ; Calculate the centered position and show the GUI
        guiX := spotifyX + (spotifyW - guiW) / 2
        guiY := spotifyY + (spotifyH - guiH) / 2
        loadingGui.Show("x" . Round(guiX) . " y" . Round(guiY) . " NA")
        WinSetTransparent(220, loadingGui)

        Sleep(700)
        cUIA := UIA_Browser(spotifyWin)

        filterFieldName := "Type to filter your library. The list of content below will update as you type."
        filterFieldType := "Edit"
        fullscreenButtonName := "fullscreen library"
        fullscreenButtonType := "Button"
        goBackButtonName := "Go back"
        goBackButtonType := "Button"

        filterField := ""
        try filterField := cUIA.FindElement({ Name: filterFieldName, Type: filterFieldType })

        if !filterField {
            if fullscreenButton := cUIA.WaitElement({ Name: fullscreenButtonName, Type: fullscreenButtonType }, 3000) {
                fullscreenButton.Click()
                Sleep(1000)
            } else {
                statusText.Value := "Error: 'fullscreen' button not found."
                MsgBox("Could not find 'fullscreen library' button.")
                return
            }
        }

        if initialFilterField := cUIA.WaitElement({ Name: filterFieldName, Type: filterFieldType }, 1000) {
            initialFilterField.SetFocus()
            Sleep(600)
            SendInput("{Shift Down}{Tab 5}{Shift Up}")
            Sleep(600)

            loop (10) {
                focusedEl := ""
                try focusedEl := cUIA.GetFocusedElement()

                if IsObject(focusedEl) && (focusedEl.Name == goBackButtonName) {
                    SendInput("{Enter}")
                    Sleep(600)
                } else {
                    break
                }
            }
        }

        if filterField := cUIA.WaitElement({ Name: filterFieldName, Type: filterFieldType }, 2000) {
            filterField.Click()
            statusText.Value := "Done! Library is ready."
        } else {
            statusText.Value := "Error: Could not find filter field."
            MsgBox("Could not find the library filter text field.")
        }

    } catch as e {
        statusText.Value := "An error occurred."
        MsgBox("An error occurred during Spotify automation: `n" e.Message)
    } finally {
        Sleep(2000)
        if IsObject(loadingGui) && loadingGui.Hwnd {
            loadingGui.Destroy()
        }
    }
}

*Volume_Down:: {
    if GetKeyState("Ctrl", "P") {
        ; Ctrl held: adjust Spotify volume
        spotify := ActivateSpotify()
        if WinWaitActive("ahk_exe Spotify.exe", , 2) {
            Send("^{Down}")
        }
        ; Re-minimize if the window was originally minimized
        if IsObject(spotify) && spotify.wasMinimized {
            hwnd := spotify.hwnd
            ; One-shot timer after 3.5 seconds; capture hwnd by value
            SetTimer((id := hwnd) => WinMinimize("ahk_id " id), -3500)
        }
    } else if GetKeyState("Alt", "P") {
        ; Alt held: adjust YouTube volume
        yt := FocusYouTube()
        if IsObject(yt) {
            Send("{Down}")
            if yt.wasMinimized {
                hwnd := yt.hwnd
                SetTimer((id := hwnd) => WinMinimize("ahk_id " id), -3500)
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
        if IsObject(spotify) && spotify.wasMinimized {
            hwnd := spotify.hwnd
            SetTimer((id := hwnd) => WinMinimize("ahk_id " id), -3500)
        }
    } else if GetKeyState("Alt", "P") {
        yt := FocusYouTube()
        if IsObject(yt) {
            Send("{Up}")
            if yt.wasMinimized {
                hwnd := yt.hwnd
                SetTimer((id := hwnd) => WinMinimize("ahk_id " id), -3500)
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
