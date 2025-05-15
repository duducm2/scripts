#Requires AutoHotkey v2.0
#SingleInstance Force ; Good practice to add

#include UIA-v2\Lib\UIA.ahk
#include UIA-v2\Lib\UIA_Browser.ahk

CapsLock & 7::
{

    static isDictating := false ; State variable
    static submitFailCount := 0 ; Counter for consecutive submit failures

    ; Activate the window containing "chatgpt - transcription"
    SetTitleMatchMode 2
    WinActivate "chatgpt - transcription"
    WinWaitActive "ahk_exe chrome.exe" ; Ensure the correct window is active

    ; Initialize UIA Browser for the active window
    cUIA := UIA_Browser()
    Sleep 300 ; Wait a bit for the UI

    if !isDictating { ; If not currently dictating, prepare field, paste prompt and start it
        try {
            ; Find and click the Botão de ditado
            dictateBtn := cUIA.FindElement({ Name: "Botão de ditado", Type: "Button" })
            if dictateBtn {
                dictateBtn.Click()
                isDictating := true ; Update state
                submitFailCount := 0 ; Reset fail counter on successful start
            } else {
                MsgBox "Botão de ditado not found."
            }
        } catch Error as e {
            MsgBox "Error during pre-dictation or starting dictation: " e.Message
            isDictating := false ; Reset state on error
        }
    } else { ; If already dictating, stop it (click Submit dictation)
        try {
            ; Find the Submit dictation button using its Name and Type
            submitBtn := cUIA.FindElement({ Name: "Enviar ditado", Type: "Button" })
            if submitBtn {
                submitBtn.Click()
                isDictating := false ; Update state
                submitFailCount := 0 ; Reset fail counter on successful submit

                ; >>> NEW: Wait for Send prompt button and press Enter <<<
                try {
                    Sleep 200 ; Small pause after clicking submit
                    ; Wait for the button to reappear/enable (Timeout 10 seconds)
                    finalSendBtn := cUIA.WaitElement({ Name: "Enviar prompt", AutomationId: "composer-submit-button" },
                    10000)
                    if finalSendBtn {
                        SendInput "{Enter}"

                    } else {
                        MsgBox "Timeout: 'Send prompt' button did not reappear after submitting dictation."
                    }
                } catch Error as e_wait {
                    MsgBox "Error waiting for/clicking final Send Prompt button: " e_wait.Message
                }
                ; >>> End NEW <<<

            } else {
                MsgBox "Submit dictation button not found."
                submitFailCount++ ; Increment fail counter
            }
        } catch Error as e {
            MsgBox "Error finding or clicking Submit dictation button: " e.Message
            submitFailCount++ ; Increment fail counter
            ; Decide if state should be reset here? Maybe not, user might want to try again.
        }

        ; Check if submit failed too many times
        if submitFailCount >= 1 {
            MsgBox "Failed to submit dictation 1 time. Assuming dictation stopped. Press hotkey again to start."
            isDictating := false ; Reset state
            submitFailCount := 0 ; Reset counter
        }
    }

    Send("{CapsLock}")
}
