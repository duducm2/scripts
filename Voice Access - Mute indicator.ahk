#Requires AutoHotkey v2.0
; No UIA includes needed
#SingleInstance Force ; Good practice to add

global isDictating := false ; State variable

; Define the hotkey Alt + Shift + B
!+b:: {
    global isDictating ; Access the global variable

    targetTitle := "ChatGPT"

    ; Common steps: Activate window and set CapsLock state
    SetTitleMatchMode 2
    if !WinExist(targetTitle) {
        MsgBox Format("Window '{}' not found!", targetTitle)
        return
    }
    WinActivate(targetTitle)
    if !WinWaitActive(targetTitle, , 2) {
        MsgBox Format("Failed to activate window '{}'!", targetTitle)
        return
    }
    Sleep 300 ; Slightly increased sleep after activation
    SetCapsLockState "AlwaysOff"

    if (!isDictating) {
        ; --- First Press: Start Dictation ---
        ; MsgBox "Starting Dictation..." ; Optional: Uncomment for debug

        numberOfTabsStart := 7

        Send "+{Esc}" ; Focus main area
        Sleep 200

        loop numberOfTabsStart {
            Send "{Tab}"
            Sleep 50 ; Adjusted sleep
        }
        Sleep 100

        Send "{Enter}" ; Activate Dictate button

        ; MsgBox Format("Sent Shift+Esc, {} Tabs, and Enter to start dictation.", numberOfTabsStart) ; Optional: Uncomment for debug
        isDictating := true ; Update state

    } else {
        ; --- Second Press: Stop Dictation & Copy ---
        ; MsgBox "Stopping Dictation and Copying..." ; Optional: Uncomment for debug

        numberOfTabsStop := 5
        copyDelay := 300 ; Milliseconds to wait before copying

        loop numberOfTabsStop {
            Send "{Tab}"
            Sleep 50 ; Adjusted sleep
        }
        Sleep 100

        Send "{Enter}" ; Stop dictation / Finalize input
        Sleep copyDelay ; Wait for processing

        Send "^+c" ; Copy transcript (Ctrl+Shift+C)
        Sleep 100

        ; MsgBox Format("Sent {} Tabs, Enter, waited {}ms, and sent Ctrl+Shift+C.", numberOfTabsStop, copyDelay) ; Optional: Uncomment for debug
        isDictating := false ; Update state
    }
}
