#Requires AutoHotkey v2.0+
#SingleInstance Force

; Assuming UIA-v2 folder is now a subfolder in the same directory as this script.
#include UIA-v2\Lib\UIA.ahk
#include UIA-v2\Lib\UIA_Browser.ahk
#include %A_ScriptDir%\env.ahk ; Include environment configuration

; Win+Alt+Shift+D -> Go to Spotify Library
#!+d:: GoToSpotifyLibrary()

GoToSpotifyLibrary() {
    SetTitleMatchMode(2)
    spotifyWin := "ahk_exe Spotify.exe"

    ; If Spotify is not open, open it and finish the script.
    if !WinExist(spotifyWin) {
        OpenSpotify()
        return
    }

    ; If Spotify is already open, activate its window.
    WinActivate(spotifyWin)
    WinWaitActive(spotifyWin, , 2)
    Sleep(300)

    try {
        cUIA := UIA_Browser(spotifyWin)

        ; Define element properties
        filterFieldName := "Type to filter your library. The list of content below will update as you type."
        filterFieldType := "Edit"
        fullscreenButtonName := "fullscreen library"
        fullscreenButtonType := "Button"
        goBackButtonName := "Go back"
        goBackButtonType := "Button"

        ; Check if the library filter text field exists.
        filterField := ""
        try filterField := cUIA.FindElement({ Name: filterFieldName, Type: filterFieldType })

        ; If the filter field is not found, we are not in library view.
        if !filterField {
            ; Click on the "fullscreen library" button to enter library view.
            if fullscreenButton := cUIA.WaitElement({ Name: fullscreenButtonName, Type: fullscreenButtonType }, 3000) {
                fullscreenButton.Click()
                Sleep(500) ; Wait for UI to update
            } else {
                MsgBox("Could not find 'fullscreen library' button.")
                return
            }
        }

        ; Now we should be in the library view.
        ; Click the "Go back" button repeatedly using a more efficient keyboard navigation.

        ; First, navigate to the "Go back" button location once.
        if initialFilterField := cUIA.WaitElement({ Name: filterFieldName, Type: filterFieldType }, 1000) {
            initialFilterField.SetFocus()
            Sleep(200)
            SendInput("{Shift Down}{Tab 5}{Shift Up}")
            Sleep(200)

            ; Now, repeatedly press Enter as long as the focus remains on a "Go back" button.
            loop (10) {
                focusedEl := ""
                try focusedEl := cUIA.GetFocusedElement()

                ; If the focused element is our button, press Enter.
                if IsObject(focusedEl) && (focusedEl.Name == goBackButtonName) {
                    SendInput("{Enter}")
                    Sleep(500) ; Wait for UI to update after the click.
                } else {
                    ; The focus has moved away, so we are done.
                    break
                }
            }
        }

        ; Finally, select the text field to filter the library.
        if filterField := cUIA.WaitElement({ Name: filterFieldName, Type: filterFieldType }, 2000) {
            filterField.Click()
        } else {
            MsgBox("Could not find the library filter text field.")
        }

    } catch as e {
        MsgBox("An error occurred during Spotify automation: `n" e.Message)
    }
}

OpenSpotify() {
    ; This function opens Spotify based on the environment (work/personal).
    ; It's adapted from your "Spotify - Open.ahk" script.
    global IS_WORK_ENVIRONMENT
    if (IS_WORK_ENVIRONMENT) { ; Work PC
        ; Try to use the Start Menu shortcut first
        link := "C:\Users\fie7ca\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Spotify.lnk"
        if FileExist(link) {
            Run(link)
            return
        }
        ; If shortcut doesn't exist, try the Microsoft Store app path
        storeApp := "explorer.exe shell:AppsFolder\SpotifyAB.SpotifyMusic_zpdnekdrzrea0!Spotify"
        Run(storeApp)
    } else { ; Personal PC
        ; On personal PC, "spotify" alias usually works
        Run("spotify")
    }
}
