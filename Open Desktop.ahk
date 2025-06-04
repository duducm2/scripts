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
    possibleTitles := ["Área de Trabalho", "Desktop", "Volume"]

    ; Determine the correct desktop path based on the IS_WORK_ENVIRONMENT variable from env.ahk
    if (IS_WORK_ENVIRONMENT) {
        chosenDesktopPath := workDesktopPath
    } else { ; Default to Personal if IS_WORK_ENVIRONMENT is false
        chosenDesktopPath := personalDesktopPath
    }

    ; Try to activate existing Desktop window matching any possible title
    found := false
    for title in possibleTitles {
        if WinExist(title) {
            WinActivate(title)
            if WinWaitActive(title, , 2) {
                found := true
                break
            }
        }
    }
    if !found {
        ; If not found or not activated, open the folder and activate when ready
        Run chosenDesktopPath
        Sleep 500
        ActivateDesktopWindow(possibleTitles)
    }
}

ActivateDesktopWindow(possibleTitles) {
    for title in possibleTitles {
        if WinWait(title, , 5) { ; Wait up to 5 seconds for any title
            WinActivate(title)
            WinWaitActive(title, , 2)
            return true
        }
    }
    return false
}
