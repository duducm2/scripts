#e:: {
    SetTitleMatchMode 2  ; Match anywhere in the title

    personalDesktopPath := "C:\Users\fie7ca\Documents\Atalhos\Desktop"
    workDesktopPath := "C:\Users\fie7ca\Desktop"

    chosenDesktopPath := ""
    windowTitle := "Área de Trabalho" ; Assuming this title is consistent for both paths.
    ; If your "Work" environment uses a different title (e.g., "Desktop"),
    ; you may need to adjust this or make it dynamic.

    ; Determine the correct desktop path based on the DESKTOP_MODE environment variable
    envDesktopMode := EnvGet("DESKTOP_MODE")

    if (envDesktopMode = "Work") {
        chosenDesktopPath := workDesktopPath
    } else { ; Default to Personal if DESKTOP_MODE is not "Work" (e.g., "Personal", empty, or any other value)
        chosenDesktopPath := personalDesktopPath
    }

    ; Try to activate existing Desktop window matching the title
    if WinExist(windowTitle) {
        WinActivate(windowTitle)
        ; Wait to confirm activation (up to 2 seconds)
        if !WinWaitActive(windowTitle, , 2) {
            ; If activation failed, try to open a new window using the chosen path
            Run chosenDesktopPath
            Sleep 500  ; Give Windows time to start the process
            WinWaitActive(windowTitle, , 5)  ; Wait up to 5 seconds for the window with the specific title
        }
    } else {
        ; If no window exists, open a new one using the chosen path
        Run chosenDesktopPath
        Sleep 500  ; Give Windows time to start the process
        WinWaitActive(windowTitle, , 5)  ; Wait up to 5 seconds for the window with the specific title
    }

    ; Final check - if a window with the title is still not active, try one last time to activate it
    if !WinActive(windowTitle) {
        Sleep 1000  ; Wait a full second
        WinActivate(windowTitle)
    }
}
