#e:: {
    SetTitleMatchMode 2  ; Match anywhere in the title

    ; Try to activate existing Desktop window
    if WinExist("Área de Trabalho") {
        WinActivate("Área de Trabalho")
        ; Wait to confirm activation (up to 2 seconds)
        if !WinWaitActive("Área de Trabalho", , 2) {
            ; If activation failed, try to open a new window
            Run "C:\Users\eduev\OneDrive\Desktop"
            Sleep 500  ; Give Windows time to start the process
            WinWaitActive("Área de Trabalho", , 5)  ; Wait up to 5 seconds
        }
    } else {
        ; If no window exists, open new one
        Run "C:\Users\eduev\OneDrive\Desktop"
        Sleep 500  ; Give Windows time to start the process
        WinWaitActive("Área de Trabalho", , 5)  ; Wait up to 5 seconds
    }

    ; Final check - if window still not active, try one last time
    if !WinActive("Área de Trabalho") {
        Sleep 1000  ; Wait a full second
        WinActivate("Área de Trabalho")
    }
}
