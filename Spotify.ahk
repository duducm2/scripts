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
        WinActivate
        return
    }

    ; 2) Decide how to open it based on the environment.
    global IS_WORK_ENVIRONMENT
    if IS_WORK_ENVIRONMENT {
        ; Work PC: Try the shortcut first.
        link := "C:\Users\fie7ca\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Spotify.lnk"
        if FileExist(link) {
            Run(link)
            return
        }

        ; Fallback for Work PC: Use the Store App command.
        Run("explorer.exe shell:AppsFolder\SpotifyAB.SpotifyMusic_zpdnekdrzrea0!Spotify")
    }
    else {
        ; Personal PC: Directly run the Microsoft Store App command.
        Run("explorer.exe shell:AppsFolder\SpotifyAB.SpotifyMusic_zpdnekdrzrea0!Spotify")
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
        WinWaitActive(spotifyWin, , 2)

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

; --- NEW Hotkeys -------------------------------------------------------------
; Volume wheel – handle differently when Ctrl is held.
*Volume_Down:: {
    if GetKeyState("Ctrl", "P") {
        ; Ctrl held: adjust Spotify only
        OpenSpotify()
        WinWaitActive("ahk_exe Spotify.exe", , 2)
        Send("^{Down}")
    } else {
        ; No Ctrl: let OS handle master volume
        Send "{Volume_Down}"
    }
}

*Volume_Up:: {
    if GetKeyState("Ctrl", "P") {
        OpenSpotify()
        WinWaitActive("ahk_exe Spotify.exe", , 2)
        Send("^{Up}")
    } else {
        Send "{Volume_Up}"
    }
}

; --- NEW Hotkeys -------------------------------------------------------------
; Alt + Volume wheel → control YouTube volume (when a YouTube tab is active)
!Volume_Down:: {
    winTitle := WinGetTitle("A")
    if InStr(winTitle, "YouTube") {
        ; Release Alt so the arrow key is not modified
        SendInput "{Alt Up}{Down}{Alt Down}"
    } else {
        ; Fallback – pass the key through to the system master-volume
        Send "{Volume_Down}"
    }
}

!Volume_Up:: {
    winTitle := WinGetTitle("A")
    if InStr(winTitle, "YouTube") {
        SendInput "{Alt Up}{Up}{Alt Down}"
    } else {
        Send "{Volume_Up}"
    }
}
