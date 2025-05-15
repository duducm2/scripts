#Requires AutoHotkey v2.0+
#InputLevel 2
#SingleInstance Ignore

#Include %A_ScriptDir%\env.ahk

; Win+E hotkey to open Desktop folder (original)
#e:: {
    SetTitleMatchMode 2  ; Match anywhere in the title

    personalDesktopPath := "C:\Users\eduev\OneDrive\Desktop"
    workDesktopPath := "C:\Users\fie7ca\Desktop"

    chosenDesktopPath := ""
    windowTitle := "Área de Trabalho" ; Assuming this title is consistent for both paths.
    ; If your "Work" environment uses a different title (e.g., "Desktop"),
    ; you may need to adjust this or make it dynamic.

    ; Determine the correct desktop path based on the IS_WORK_ENVIRONMENT variable from env.ahk
    if (IS_WORK_ENVIRONMENT) {
        chosenDesktopPath := workDesktopPath
    } else { ; Default to Personal if IS_WORK_ENVIRONMENT is false
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
}
