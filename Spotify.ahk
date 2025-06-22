#Requires AutoHotkey v2.0+
#SingleInstance Force

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
    if WinExist("ahk_exe Spotify.exe") || WinExist("Spotify") {
        WinActivate
        return
    }

    global IS_WORK_ENVIRONMENT
    if IS_WORK_ENVIRONMENT {
        link := "C:\Users\fie7ca\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Spotify.lnk"
        if FileExist(link) {
            Run(link)
            return
        }
        storeApp := "explorer.exe shell:AppsFolder\SpotifyAB.SpotifyMusic_zpdnekdrzrea0!Spotify"
        Run(storeApp)
    } else {
        Run("spotify")
    }
}

; =============================================================================
; Go to Spotify Library
; Hotkey: Win+Alt+Shift+D
; Original File: Spotify - Go to library.ahk
; =============================================================================
#!+d:: GoToSpotifyLibrary()

GoToSpotifyLibrary() {
    loadingGui := Gui()
    loadingGui.Opt("+AlwaysOnTop -Caption +ToolWindow")
    loadingGui.BackColor := "333333"
    loadingGui.SetFont("s12 cFFFFFF", "Segoe UI")
    statusText := loadingGui.Add("Text", "w300 Center", "Navigating to your library...")
    loadingGui.Show("AutoSize Center NA")
    WinSetTransparent(220, loadingGui)

    try {
        SetTitleMatchMode(2)
        spotifyWin := "ahk_exe Spotify.exe"

        if !WinExist(spotifyWin) {
            OpenSpotify()
            statusText.Value := "Spotify is opening..."
            Sleep(4000)
            return
        }

        WinActivate(spotifyWin)
        WinWaitActive(spotifyWin, , 2)
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
