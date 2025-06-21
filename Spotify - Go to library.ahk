#Requires AutoHotkey v2.0+
#SingleInstance Force

; Assuming UIA-v2 folder is now a subfolder in the same directory as this script.
#include UIA-v2\Lib\UIA.ahk
#include UIA-v2\Lib\UIA_Browser.ahk
#include %A_ScriptDir%\env.ahk ; Include environment configuration

; Win+Alt+Shift+D -> Go to Spotify Library
#!+d:: GoToSpotifyLibrary()

GoToSpotifyLibrary() {
    ; Create and show a custom GUI for loading status
    loadingGui := Gui()
    loadingGui.Opt("+AlwaysOnTop -Caption +ToolWindow") ; Frameless, always on top
    loadingGui.BackColor := "333333" ; Dark background
    loadingGui.SetFont("s12 cFFFFFF", "Segoe UI") ; White text
    statusText := loadingGui.Add("Text", "w300 Center", "Navigating to your library...")
    loadingGui.Show("AutoSize Center NA")
    WinSetTransparent(220, loadingGui) ; Make it semi-transparent

    try {
        SetTitleMatchMode(2)
        spotifyWin := "ahk_exe Spotify.exe"

        ; If Spotify is not open, open it and finish the script.
        if !WinExist(spotifyWin) {
            OpenSpotify()
            statusText.Value := "Spotify is opening..." ; Update text
            Sleep(4000) ; Give Spotify time to launch
            return
        }

        ; If Spotify is already open, activate its window.
        WinActivate(spotifyWin)
        WinWaitActive(spotifyWin, , 2)
        Sleep(700)

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
                Sleep(1000) ; Wait for UI to update
            } else {
                statusText.Value := "Error: 'fullscreen' button not found."
                MsgBox("Could not find 'fullscreen library' button.")
                return
            }
        }

        ; Now we should be in the library view.
        ; Click the "Go back" button repeatedly using a more efficient keyboard navigation.

        ; First, navigate to the "Go back" button location once.
        if initialFilterField := cUIA.WaitElement({ Name: filterFieldName, Type: filterFieldType }, 1000) {
            initialFilterField.SetFocus()
            Sleep(600)
            SendInput("{Shift Down}{Tab 5}{Shift Up}")
            Sleep(600)

            ; Now, repeatedly press Enter as long as the focus remains on a "Go back" button.
            loop (10) {
                focusedEl := ""
                try focusedEl := cUIA.GetFocusedElement()

                ; If the focused element is our button, press Enter.
                if IsObject(focusedEl) && (focusedEl.Name == goBackButtonName) {
                    SendInput("{Enter}")
                    Sleep(600) ; Wait for UI to update after the click.
                } else {
                    ; The focus has moved away, so we are done.
                    break
                }
            }
        }

        ; Finally, select the text field to filter the library.
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
        Sleep(2000) ; Keep the final message visible for a bit
        if IsObject(loadingGui) && loadingGui.Hwnd {
            loadingGui.Destroy()
        }
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
